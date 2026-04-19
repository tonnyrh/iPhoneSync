<#


This application is a Windows desktop tool for syncing media files from an iPhone’s Internal Storage to a folder on a PC.

It is designed to handle iPhone folders more reliably by processing one top-level folder at a time, repeatedly scanning and “warming up” the folder contents before and after copying. This helps detect files that appear gradually through the Windows Shell interface and ensures the destination stays in sync with the source.

The app includes:

automatic iPhone detection
target folder selection
progress, status, and log output
cancel support during sync
basic persistence of the last used source display and target folder in a local config file

In short, it is a more robust iPhone-to-PC media sync utility built for situations where normal file enumeration from iPhone storage can be incomplete or delayed.

Tonny Roger Holm - 2026

#>


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================
# Config
# =========================
$script:Config = @{
    TopFolderWarmupMaxPasses            = 8
    TopFolderWarmupStablePassesRequired = 2
    TopFolderWarmupPauseSeconds         = 3
    TopFolderProcessMaxRounds           = 20
    MasterDiscoveryMaxRounds            = 20
    MasterEmptyDiscoveryTolerance       = 4
    EnumerationRetryCount               = 3
    EnumerationRetryDelayMs             = 700
    FileReadyTimeoutSeconds             = 45
    SkipExtensions                      = @()
    VerboseFolderLogging                = $false
}

$script:IPhoneSourceFolder = $null
$script:CancelRequested    = $false
$script:IsRunning          = $false
$script:AppConfig          = $null
$script:AppConfigPath      = $null

# =========================
# App config
# =========================
function Get-AppConfigRootPath {
    $appData = [Environment]::GetFolderPath('ApplicationData')

    if (-not [string]::IsNullOrWhiteSpace($appData) -and (Test-Path -LiteralPath $appData)) {
        return (Join-Path -Path $appData -ChildPath 'iPhoneMediaSync')
    }

    $tempPath = [System.IO.Path]::GetTempPath()
    return (Join-Path -Path $tempPath -ChildPath 'iPhoneMediaSync')
}

function Get-AppConfigFilePath {
    $root = Get-AppConfigRootPath
    return (Join-Path -Path $root -ChildPath 'config.json')
}

function New-DefaultAppConfig {
    return [ordered]@{
        LastSourceDisplay = ''
        LastTargetFolder  = ''
        PendingSyncState  = $null
    }
}

function Initialize-AppConfig {
    $script:AppConfigPath = Get-AppConfigFilePath
    $script:AppConfig = New-DefaultAppConfig

    $configFolder = Split-Path -Path $script:AppConfigPath -Parent
    if (-not (Test-Path -LiteralPath $configFolder)) {
        New-Item -ItemType Directory -Path $configFolder -Force | Out-Null
    }

    if (Test-Path -LiteralPath $script:AppConfigPath) {
        try {
            $raw = Get-Content -LiteralPath $script:AppConfigPath -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $loaded = $raw | ConvertFrom-Json -ErrorAction Stop

                if ($null -ne $loaded.PSObject.Properties['LastSourceDisplay']) {
                    $script:AppConfig.LastSourceDisplay = [string]$loaded.LastSourceDisplay
                }

                if ($null -ne $loaded.PSObject.Properties['LastTargetFolder']) {
                    $script:AppConfig.LastTargetFolder = [string]$loaded.LastTargetFolder
                }

                if ($null -ne $loaded.PSObject.Properties['PendingSyncState']) {
                    $script:AppConfig.PendingSyncState = $loaded.PendingSyncState
                }
            }
        }
        catch {
            $script:AppConfig = New-DefaultAppConfig
        }
    }
}

function Save-AppConfig {
    try {
        if ($null -eq $script:AppConfig) {
            $script:AppConfig = New-DefaultAppConfig
        }

        if ([string]::IsNullOrWhiteSpace($script:AppConfigPath)) {
            $script:AppConfigPath = Get-AppConfigFilePath
        }

        $configFolder = Split-Path -Path $script:AppConfigPath -Parent
        if (-not (Test-Path -LiteralPath $configFolder)) {
            New-Item -ItemType Directory -Path $configFolder -Force | Out-Null
        }

        $json = $script:AppConfig | ConvertTo-Json -Depth 5
        [System.IO.File]::WriteAllText($script:AppConfigPath, $json, [System.Text.Encoding]::UTF8)
    }
    catch {
    }
}

function Update-AppConfigFromUi {
    param(
        [string]$SourceDisplay,
        [string]$TargetFolder
    )

    if ($null -eq $script:AppConfig) {
        $script:AppConfig = New-DefaultAppConfig
    }

    if ($PSBoundParameters.ContainsKey('SourceDisplay')) {
        $script:AppConfig.LastSourceDisplay = [string]$SourceDisplay
    }

    if ($PSBoundParameters.ContainsKey('TargetFolder')) {
        $script:AppConfig.LastTargetFolder = [string]$TargetFolder
    }

    Save-AppConfig
}

function New-PendingSyncState {
    param(
        [string]$SourceDisplay,
        [string]$TargetFolder
    )

    return [ordered]@{
        IsPending                  = $true
        SourceDisplay              = [string]$SourceDisplay
        TargetFolder               = [string]$TargetFolder
        CompletedTopFolders        = @()
        LastCompletedTopFolderName = ''
        LastCompletedTopFolderSafeName = ''
        UpdatedAt                  = (Get-Date).ToString('o')
    }
}

function Get-PendingSyncStateForContext {
    param(
        [string]$SourceDisplay,
        [string]$TargetFolder
    )

    if ($null -eq $script:AppConfig) {
        return $null
    }

    $state = $script:AppConfig.PendingSyncState
    if ($null -eq $state) {
        return $null
    }

    $stateSource = [string]$state.SourceDisplay
    $stateTarget = [string]$state.TargetFolder

    if ($stateSource -ne [string]$SourceDisplay) {
        return $null
    }

    if ($stateTarget -ne [string]$TargetFolder) {
        return $null
    }

    if ($state.PSObject.Properties.Name -notcontains 'IsPending') {
        return $null
    }

    if (-not [bool]$state.IsPending) {
        return $null
    }

    return $state
}

function Start-PendingSyncTracking {
    param(
        [string]$SourceDisplay,
        [string]$TargetFolder
    )

    if ($null -eq $script:AppConfig) {
        $script:AppConfig = New-DefaultAppConfig
    }

    $existingState = Get-PendingSyncStateForContext -SourceDisplay $SourceDisplay -TargetFolder $TargetFolder

    if ($null -ne $existingState) {
        $script:AppConfig.PendingSyncState = $existingState
        $script:AppConfig.PendingSyncState.UpdatedAt = (Get-Date).ToString('o')
    }
    else {
        $script:AppConfig.PendingSyncState = New-PendingSyncState -SourceDisplay $SourceDisplay -TargetFolder $TargetFolder
    }

    Save-AppConfig
}

function Update-PendingSyncProgress {
    param(
        [string]$TopFolderName,
        [string]$TopFolderSafeName
    )

    if ($null -eq $script:AppConfig) {
        $script:AppConfig = New-DefaultAppConfig
    }

    if ($null -eq $script:AppConfig.PendingSyncState) {
        $script:AppConfig.PendingSyncState = New-PendingSyncState -SourceDisplay $script:AppConfig.LastSourceDisplay -TargetFolder $script:AppConfig.LastTargetFolder
    }

    $completed = New-Object System.Collections.Generic.List[string]

    if ($null -ne $script:AppConfig.PendingSyncState.CompletedTopFolders) {
        foreach ($name in @($script:AppConfig.PendingSyncState.CompletedTopFolders)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$name) -and -not $completed.Contains([string]$name)) {
                $completed.Add([string]$name)
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($TopFolderSafeName) -and -not $completed.Contains($TopFolderSafeName)) {
        $completed.Add($TopFolderSafeName)
    }

    $script:AppConfig.PendingSyncState.CompletedTopFolders = @($completed.ToArray())
    $script:AppConfig.PendingSyncState.LastCompletedTopFolderName = [string]$TopFolderName
    $script:AppConfig.PendingSyncState.LastCompletedTopFolderSafeName = [string]$TopFolderSafeName
    $script:AppConfig.PendingSyncState.UpdatedAt = (Get-Date).ToString('o')
    Save-AppConfig
}

function Clear-PendingSyncState {
    if ($null -eq $script:AppConfig) {
        $script:AppConfig = New-DefaultAppConfig
    }

    $script:AppConfig.PendingSyncState = $null
    Save-AppConfig
}

# =========================
# Logging / UI
# =========================
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox
    )

    $timestamp = (Get-Date).ToString('G', [System.Globalization.CultureInfo]::CurrentCulture)
    $line = "[{0}] {1}" -f $timestamp, $Message

    if ($LogBox.InvokeRequired) {
        $null = $LogBox.Invoke([Action]{
            $script:txtLog.AppendText($line + [Environment]::NewLine)
            $script:txtLog.SelectionStart = $script:txtLog.Text.Length
            $script:txtLog.ScrollToCaret()
        })
    }
    else {
        $LogBox.AppendText($line + [Environment]::NewLine)
        $LogBox.SelectionStart = $LogBox.Text.Length
        $LogBox.ScrollToCaret()
    }
}

function Set-Status {
    param(
        [string]$Text
    )

    if ($script:lblStatus.InvokeRequired) {
        $null = $script:lblStatus.Invoke([Action]{
            $script:lblStatus.Text = $Text
        })
    }
    else {
        $script:lblStatus.Text = $Text
    }
}

function Set-Stats {
    param(
        [int]$TopFoldersSeen,
        [int]$TopFoldersDone,
        [int]$FilesSeen,
        [int]$Pending,
        [int]$Copied,
        [int]$Skipped,
        [int]$Errors
    )

    $globalText = "Global folders - Seen: {0} | Completed: {1} || Global files - Copied: {2} | Skipped: {3} | Errors: {4}" -f $TopFoldersSeen, $TopFoldersDone, $Copied, $Skipped, $Errors
    $currentFolderText = "Current folder files - Seen: {0} | Pending/Diff: {1}" -f $FilesSeen, $Pending

    if ($script:lblGlobalStats.InvokeRequired) {
        $null = $script:lblGlobalStats.Invoke([Action]{
            $script:lblGlobalStats.Text = $globalText
            $script:lblFolderStats.Text = $currentFolderText
        })
    }
    else {
        $script:lblGlobalStats.Text = $globalText
        $script:lblFolderStats.Text = $currentFolderText
    }
}

# =========================
# Helper functions
# =========================
function Get-SafeName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return ($Name -replace '[\\/:*?"<>|]', '_')
}

function Join-RelativePath {
    param(
        [AllowEmptyString()]
        [string]$Base = '',

        [Parameter(Mandatory = $true)]
        [string]$Child
    )

    if ([string]::IsNullOrWhiteSpace($Base)) {
        return $Child
    }

    return ("{0}\{1}" -f $Base, $Child)
}

function Split-RelativePath {
    param(
        [AllowEmptyString()]
        [string]$RelativePath = ''
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return @()
    }

    return ($RelativePath -split '\\')
}

function Ensure-DirectoryExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-TargetFileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    try {
        if (Test-Path -LiteralPath $TargetPath) {
            return $true
        }
    }
    catch {
    }

    try {
        $item = Get-Item -LiteralPath $TargetPath -ErrorAction Stop
        if ($null -ne $item) {
            return $true
        }
    }
    catch {
    }

    try {
        $parentPath = Split-Path -Path $TargetPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($parentPath) -and (Test-Path -LiteralPath $parentPath)) {
            $existingInFolder = Get-ChildItem -LiteralPath $parentPath -Force -File -ErrorAction Stop |
                Where-Object { $_.Name -ieq $FileName } |
                Select-Object -First 1

            if ($null -ne $existingInFolder) {
                return $true
            }
        }
    }
    catch {
    }

    return $false
}

function Get-ShellItemsSafe {
    param(
        [Parameter(Mandatory = $true)]
        $ShellFolder
    )

    $lastError = $null

    for ($i = 1; $i -le $script:Config.EnumerationRetryCount; $i++) {
        try {
            $items = @($ShellFolder.Items())
            return $items
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds $script:Config.EnumerationRetryDelayMs
        }
    }

    throw $lastError
}

function Get-ItemSizeSafe {
    param(
        [Parameter(Mandatory = $true)]
        $ShellItem
    )

    try {
        $value = $ShellItem.ExtendedProperty('Size')
        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
            return $null
        }

        return [int64]$value
    }
    catch {
        return $null
    }
}

function Get-ShellItemFileNameSafe {
    param(
        [Parameter(Mandatory = $true)]
        $ShellItem
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    try {
        $value = [string]$ShellItem.ExtendedProperty('System.FileName')
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $candidates.Add($value)
        }
    }
    catch {
    }

    try {
        $value = [string]$ShellItem.ExtendedProperty('System.ParsingName')
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $candidates.Add($value)
        }
    }
    catch {
    }

    try {
        $value = [string]$ShellItem.Path
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $candidates.Add($value)
        }
    }
    catch {
    }

    try {
        $value = [string]$ShellItem.Name
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $candidates.Add($value)
        }
    }
    catch {
    }

    foreach ($candidate in $candidates) {
        $leaf = try { Split-Path -Path $candidate -Leaf } catch { $candidate }
        if (-not [string]::IsNullOrWhiteSpace($leaf)) {
            return $leaf
        }
    }

    return $null
}

function Wait-ForFileReady {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$TimeoutSeconds = 45
    )

    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()

    while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            if (Test-Path -LiteralPath $Path) {
                $item1 = Get-Item -LiteralPath $Path -ErrorAction Stop
                $size1 = $item1.Length

                Start-Sleep -Milliseconds 750

                $item2 = Get-Item -LiteralPath $Path -ErrorAction Stop
                $size2 = $item2.Length

                if ($size1 -eq $size2) {
                    return $true
                }
            }
        }
        catch {
        }

        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
    }

    return $false
}

# =========================
# iPhone discovery
# =========================
function Get-IPhoneInternalStorageFolder {
    $shell = New-Object -ComObject Shell.Application
    $thisPC = $shell.Namespace(17)

    if ($null -eq $thisPC) {
        return $null
    }

    foreach ($device in $thisPC.Items()) {
        if ($device.Name -eq 'Apple iPhone') {
            try {
                $deviceFolder = $device.GetFolder()
                if ($null -eq $deviceFolder) {
                    continue
                }

                foreach ($child in $deviceFolder.Items()) {
                    if ($child.Name -eq 'Internal Storage') {
                        return $child.GetFolder()
                    }
                }
            }
            catch {
                continue
            }
        }
    }

    return $null
}

function Get-TopFoldersFromInternalStorage {
    param(
        [Parameter(Mandatory = $true)]
        $RootFolder
    )

    $list = New-Object System.Collections.Generic.List[object]
    $items = Get-ShellItemsSafe -ShellFolder $RootFolder

    foreach ($item in $items) {
        if ($item.IsFolder) {
            $safeName = Get-SafeName -Name $item.Name

            $list.Add([pscustomobject]@{
                Name     = $item.Name
                SafeName = $safeName
            })
        }
    }

    return @($list | Sort-Object SafeName -Unique)
}

function Resolve-TopFolderByName {
    param(
        [Parameter(Mandatory = $true)]
        $RootFolder,

        [Parameter(Mandatory = $true)]
        [string]$TopFolderName
    )

    $items = Get-ShellItemsSafe -ShellFolder $RootFolder

    foreach ($item in $items) {
        if ($item.IsFolder -and $item.Name -eq $TopFolderName) {
            try {
                return $item.GetFolder()
            }
            catch {
                return $null
            }
        }
    }

    return $null
}

# =========================
# Mapping / index
# =========================
function Resolve-ShellFolderByRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        $RootFolder,

        [AllowEmptyString()]
        [string]$RelativeFolderPath = ''
    )

    if ([string]::IsNullOrWhiteSpace($RelativeFolderPath)) {
        return $RootFolder
    }

    $current = $RootFolder
    $segments = Split-RelativePath -RelativePath $RelativeFolderPath

    foreach ($segment in $segments) {
        $found = $null
        $items = Get-ShellItemsSafe -ShellFolder $current

        foreach ($item in $items) {
            if ($item.IsFolder -and (Get-SafeName -Name $item.Name) -eq $segment) {
                $found = $item
                break
            }
        }

        if ($null -eq $found) {
            return $null
        }

        try {
            $current = $found.GetFolder()
        }
        catch {
            return $null
        }

        if ($null -eq $current) {
            return $null
        }
    }

    return $current
}

function Get-ShellItemInFolderByName {
    param(
        [Parameter(Mandatory = $true)]
        $ShellFolder,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $items = Get-ShellItemsSafe -ShellFolder $ShellFolder

    foreach ($item in $items) {
        if ($item.IsFolder) {
            continue
        }

        $fileName = Get-ShellItemFileNameSafe -ShellItem $item

        if ($item.Name -eq $Name -or $fileName -eq $Name) {
            return $item
        }
    }

    return $null
}

function New-SourceIndex {
    return [ordered]@{
        Files          = [System.Collections.Generic.Dictionary[string, object]]::new()
        FolderCount    = 0
        FileCount      = 0
        EnumerationErr = 0
    }
}

function Add-FileRecord {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Index,

        [AllowEmptyString()]
        [string]$RelativeFolderPath = '',

        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [string]$DisplayName = '',

        [Nullable[int64]]$Size
    )

    $safeFileName = Get-SafeName -Name $FileName
    $relativeFilePath = Join-RelativePath -Base $RelativeFolderPath -Child $safeFileName

    $Index.Files[$relativeFilePath] = [pscustomobject]@{
        RelativeFilePath   = $relativeFilePath
        RelativeFolderPath = $RelativeFolderPath
        FileName           = $FileName
        DisplayName        = $DisplayName
        SafeFileName       = $safeFileName
        Size               = $Size
    }

    $Index.FileCount = $Index.Files.Count
}

function Build-SourceIndexRecursive {
    param(
        [Parameter(Mandatory = $true)]
        $ShellFolder,

        [AllowEmptyString()]
        [string]$RelativeFolderPath = '',

        [Parameter(Mandatory = $true)]
        [hashtable]$Index,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ($script:CancelRequested) {
        return
    }

    try {
        $items = Get-ShellItemsSafe -ShellFolder $ShellFolder
    }
    catch {
        $Index.EnumerationErr++
        Write-Log -Message ("Failed to list folder '{0}': {1}" -f $RelativeFolderPath, $_.Exception.Message) -LogBox $LogBox
        return
    }

    foreach ($item in $items) {
        if ($script:CancelRequested) {
            return
        }

        [System.Windows.Forms.Application]::DoEvents()

        $name = $item.Name

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($item.IsFolder) {
            $safeFolderName = Get-SafeName -Name $name
            $childRelativeFolderPath = Join-RelativePath -Base $RelativeFolderPath -Child $safeFolderName
            $Index.FolderCount++

            if ($script:Config.VerboseFolderLogging) {
                Write-Log -Message ("Scanning folder: {0}" -f $childRelativeFolderPath) -LogBox $LogBox
            }

            try {
                $subFolder = $item.GetFolder()
                if ($null -ne $subFolder) {
                    Build-SourceIndexRecursive -ShellFolder $subFolder -RelativeFolderPath $childRelativeFolderPath -Index $Index -LogBox $LogBox
                }
            }
            catch {
                $Index.EnumerationErr++
                Write-Log -Message ("Error opening folder '{0}': {1}" -f $childRelativeFolderPath, $_.Exception.Message) -LogBox $LogBox
            }

            continue
        }

        $fileName = Get-ShellItemFileNameSafe -ShellItem $item
        if ([string]::IsNullOrWhiteSpace($fileName)) {
            $fileName = $name
        }

        $extension = [System.IO.Path]::GetExtension($fileName)
        if ($script:Config.SkipExtensions.Count -gt 0 -and $script:Config.SkipExtensions -contains $extension) {
            continue
        }

        $size = Get-ItemSizeSafe -ShellItem $item
        Add-FileRecord -Index $Index -RelativeFolderPath $RelativeFolderPath -FileName $fileName -DisplayName $name -Size $size
    }
}

function Get-SourceSignature {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Index
    )

    $keys = @($Index.Files.Keys | Sort-Object)
    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($key in $keys) {
        $file = $Index.Files[$key]
        $sizeText = if ($null -ne $file.Size) { [string]$file.Size } else { 'null' }
        $parts.Add(('{0}|{1}' -f $key, $sizeText))
    }

    return ($parts -join "`n")
}

function Warm-AndIndex-OneTopFolder {
    param(
        [Parameter(Mandatory = $true)]
        $TopFolderShell,

        [Parameter(Mandatory = $true)]
        [string]$TopFolderSafeName,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox
    )

    $previousSignature = $null
    $stablePasses = 0
    $lastIndex = $null

    for ($pass = 1; $pass -le $script:Config.TopFolderWarmupMaxPasses; $pass++) {
        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected during warm-up.' -f $TopFolderSafeName) -LogBox $LogBox
            break
        }

        Set-Status -Text ('Warming up folder {0} - pass {1}/{2}' -f $TopFolderSafeName, $pass, $script:Config.TopFolderWarmupMaxPasses)
        Write-Log -Message ('[{0}] Warm-up/scan pass {1} starting...' -f $TopFolderSafeName, $pass) -LogBox $LogBox

        $index = New-SourceIndex
        Build-SourceIndexRecursive -ShellFolder $TopFolderShell -RelativeFolderPath '' -Index $index -LogBox $LogBox

        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected while building index.' -f $TopFolderSafeName) -LogBox $LogBox
            return $lastIndex
        }

        $signature = Get-SourceSignature -Index $index

        if ($signature -eq $previousSignature) {
            $stablePasses++
            Write-Log -Message ('[{0}] Pass {1}: no change. Stability {2}/{3}' -f $TopFolderSafeName, $pass, $stablePasses, $script:Config.TopFolderWarmupStablePassesRequired) -LogBox $LogBox
        }
        else {
            $stablePasses = 0
            Write-Log -Message ('[{0}] Pass {1}: new files/folders discovered.' -f $TopFolderSafeName, $pass) -LogBox $LogBox
        }

        Write-Log -Message ('[{0}] Pass {1}: folders={2}, files={3}, enumeration-errors={4}' -f $TopFolderSafeName, $pass, $index.FolderCount, $index.FileCount, $index.EnumerationErr) -LogBox $LogBox

        $previousSignature = $signature
        $lastIndex = $index

        if ($stablePasses -ge $script:Config.TopFolderWarmupStablePassesRequired) {
            Write-Log -Message ('[{0}] Folder is considered stable enough for this round.' -f $TopFolderSafeName) -LogBox $LogBox
            break
        }

        if ($pass -lt $script:Config.TopFolderWarmupMaxPasses) {
            for ($s = 1; $s -le $script:Config.TopFolderWarmupPauseSeconds; $s++) {
                if ($script:CancelRequested) {
                    Write-Log -Message ('[{0}] Cancel detected during warm-up pause.' -f $TopFolderSafeName) -LogBox $LogBox
                    break
                }

                Start-Sleep -Seconds 1
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }

    return $lastIndex
}

function Get-PendingFiles {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SourceIndex,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot
    )

    $pending = New-Object System.Collections.Generic.List[object]

    foreach ($relativeFilePath in ($SourceIndex.Files.Keys | Sort-Object)) {
        $sourceFile = $SourceIndex.Files[$relativeFilePath]
        $targetPath = Join-Path -Path $TargetRoot -ChildPath $relativeFilePath

        if (-not (Test-TargetFileExists -TargetPath $targetPath -FileName $sourceFile.SafeFileName)) {
            $pending.Add($sourceFile)
            continue
        }
    }

    return $pending
}

function Copy-OnePendingFile {
    param(
        [Parameter(Mandatory = $true)]
        $TopFolderShell,

        [Parameter(Mandatory = $true)]
        $PendingFile,

        [Parameter(Mandatory = $true)]
        [string]$TargetTopFolderRoot,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox
    )

    $destinationFolderPath = if ([string]::IsNullOrWhiteSpace($PendingFile.RelativeFolderPath)) {
        $TargetTopFolderRoot
    }
    else {
        Join-Path -Path $TargetTopFolderRoot -ChildPath $PendingFile.RelativeFolderPath
    }

    Ensure-DirectoryExists -Path $destinationFolderPath

    $sourceFolder = Resolve-ShellFolderByRelativePath -RootFolder $TopFolderShell -RelativeFolderPath $PendingFile.RelativeFolderPath
    if ($null -eq $sourceFolder) {
        throw ('Source folder not found: {0}' -f $PendingFile.RelativeFolderPath)
    }

    $sourceItem = Get-ShellItemInFolderByName -ShellFolder $sourceFolder -Name $PendingFile.FileName
    if ($null -eq $sourceItem) {
        throw ('Source file not found: {0}' -f $PendingFile.RelativeFilePath)
    }

    $targetPath = Join-Path -Path $TargetTopFolderRoot -ChildPath $PendingFile.RelativeFilePath

    if (Test-TargetFileExists -TargetPath $targetPath -FileName $PendingFile.SafeFileName) {
        Write-Log -Message ('Skipping existing file: {0}' -f $PendingFile.RelativeFilePath) -LogBox $LogBox
        return $false
    }

    $shell = New-Object -ComObject Shell.Application
    $destinationShell = $shell.Namespace($destinationFolderPath)

    if ($null -eq $destinationShell) {
        throw ('Failed to open destination folder in Shell: {0}' -f $destinationFolderPath)
    }

    Write-Log -Message ('Copying: {0}' -f $PendingFile.RelativeFilePath) -LogBox $LogBox

    $copyFlags = 16 + 4 + 1024
    $destinationShell.CopyHere($sourceItem, $copyFlags)

    $ready = Wait-ForFileReady -Path $targetPath -TimeoutSeconds $script:Config.FileReadyTimeoutSeconds
    if (-not $ready) {
        throw ('Timeout while verifying: {0}' -f $PendingFile.RelativeFilePath)
    }

    return $true
}

function Process-OneTopFolderUntilStable {
    param(
        [Parameter(Mandatory = $true)]
        $InternalStorageRootFolder,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$TopFolderInfo,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox,

        [Parameter(Mandatory = $true)]
        [ref]$TotalCopied,

        [Parameter(Mandatory = $true)]
        [ref]$TotalSkipped,

        [Parameter(Mandatory = $true)]
        [ref]$TotalErrors,

        [Parameter(Mandatory = $true)]
        [int]$TopFoldersSeen,

        [Parameter(Mandatory = $true)]
        [int]$TopFoldersDone
    )

    $topName = $TopFolderInfo.Name
    $topSafe = $TopFolderInfo.SafeName
    $targetTopFolderRoot = Join-Path -Path $TargetRoot -ChildPath $topSafe
    Ensure-DirectoryExists -Path $targetTopFolderRoot

    Write-Log -Message ('Starting processing of top folder: {0}' -f $topName) -LogBox $LogBox

    $stableZeroDiffRounds = 0
    $lastPendingCount = -1
    $countedSkippedFiles = [System.Collections.Generic.HashSet[string]]::new()

    for ($round = 1; $round -le $script:Config.TopFolderProcessMaxRounds; $round++) {
        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected before round start.' -f $topSafe) -LogBox $LogBox
            break
        }

        $topFolderShell = Resolve-TopFolderByName -RootFolder $InternalStorageRootFolder -TopFolderName $topName
        if ($null -eq $topFolderShell) {
            Write-Log -Message ('[{0}] Top folder is no longer available on the iPhone.' -f $topSafe) -LogBox $LogBox
            break
        }

        Write-Log -Message ('[{0}] -------- Round {1} --------' -f $topSafe, $round) -LogBox $LogBox

        $sourceIndex = Warm-AndIndex-OneTopFolder -TopFolderShell $topFolderShell -TopFolderSafeName $topSafe -LogBox $LogBox

        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected after warm-up/index.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        if ($null -eq $sourceIndex) {
            Write-Log -Message ('[{0}] Index was not created.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        $pending = Get-PendingFiles -SourceIndex $sourceIndex -TargetRoot $targetTopFolderRoot
        $pendingCount = $pending.Count
        $skipThisRound = $sourceIndex.FileCount - $pendingCount
        if ($skipThisRound -lt 0) {
            $skipThisRound = 0
        }

        $pendingPathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($pendingFile in $pending) {
            if (-not [string]::IsNullOrWhiteSpace([string]$pendingFile.RelativeFilePath)) {
                $null = $pendingPathSet.Add([string]$pendingFile.RelativeFilePath)
            }
        }

        foreach ($sourceRelativeFilePath in $sourceIndex.Files.Keys) {
            if (-not $pendingPathSet.Contains([string]$sourceRelativeFilePath) -and -not $countedSkippedFiles.Contains([string]$sourceRelativeFilePath)) {
                $null = $countedSkippedFiles.Add([string]$sourceRelativeFilePath)
                $TotalSkipped.Value++
            }
        }

        Write-Log -Message ('[{0}] Round {1}: folders={2}, files={3}, diff={4}' -f $topSafe, $round, $sourceIndex.FolderCount, $sourceIndex.FileCount, $pendingCount) -LogBox $LogBox
        Set-Stats -TopFoldersSeen $TopFoldersSeen -TopFoldersDone $TopFoldersDone -FilesSeen $sourceIndex.FileCount -Pending $pendingCount -Copied $TotalCopied.Value -Skipped $TotalSkipped.Value -Errors $TotalErrors.Value

        if ($pendingCount -eq 0) {
            $stableZeroDiffRounds++
            Write-Log -Message ('[{0}] Zero diff. Stability {1}/2' -f $topSafe, $stableZeroDiffRounds) -LogBox $LogBox

            if ($stableZeroDiffRounds -ge 2) {
                Write-Log -Message ('[{0}] Folder is complete and stable.' -f $topSafe) -LogBox $LogBox
                return $true
            }

            continue
        }
        else {
            $stableZeroDiffRounds = 0
        }

        $copiedThisRound = 0
        $errorsThisRound = 0

        foreach ($file in $pending) {
            if ($script:CancelRequested) {
                Write-Log -Message ('[{0}] Cancel detected during file copy loop.' -f $topSafe) -LogBox $LogBox
                return $false
            }

            try {
                Set-Status -Text ('Folder {0}: copying {1}/{2}' -f $topSafe, ($copiedThisRound + $errorsThisRound + 1), $pendingCount)
                $copied = Copy-OnePendingFile -TopFolderShell $topFolderShell -PendingFile $file -TargetTopFolderRoot $targetTopFolderRoot -LogBox $LogBox

                if ($copied) {
                    $copiedThisRound++
                    $TotalCopied.Value++
                }
                else {
                    if (-not $countedSkippedFiles.Contains([string]$file.RelativeFilePath)) {
                        $null = $countedSkippedFiles.Add([string]$file.RelativeFilePath)
                        $TotalSkipped.Value++
                    }
                }
            }
            catch {
                $errorsThisRound++
                $TotalErrors.Value++
                Write-Log -Message ('[{0}] Error copying ''{1}'': {2}' -f $topSafe, $file.RelativeFilePath, $_.Exception.Message) -LogBox $LogBox
            }

            [System.Windows.Forms.Application]::DoEvents()
            Set-Stats -TopFoldersSeen $TopFoldersSeen -TopFoldersDone $TopFoldersDone -FilesSeen $sourceIndex.FileCount -Pending ($pendingCount - $copiedThisRound - $errorsThisRound) -Copied $TotalCopied.Value -Skipped $TotalSkipped.Value -Errors $TotalErrors.Value
        }

        Write-Log -Message ('[{0}] Round {1}: copied={2}, errors={3}' -f $topSafe, $round, $copiedThisRound, $errorsThisRound) -LogBox $LogBox

        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected after copy round.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        $topFolderShellAfter = Resolve-TopFolderByName -RootFolder $InternalStorageRootFolder -TopFolderName $topName
        if ($null -eq $topFolderShellAfter) {
            Write-Log -Message ('[{0}] Top folder disappeared after copy.' -f $topSafe) -LogBox $LogBox
            break
        }

        $postIndex = Warm-AndIndex-OneTopFolder -TopFolderShell $topFolderShellAfter -TopFolderSafeName $topSafe -LogBox $LogBox

        if ($script:CancelRequested) {
            Write-Log -Message ('[{0}] Cancel detected after post-warmup.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        if ($null -eq $postIndex) {
            Write-Log -Message ('[{0}] Post index was not created.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        $postPending = Get-PendingFiles -SourceIndex $postIndex -TargetRoot $targetTopFolderRoot
        $postPendingCount = $postPending.Count

        Write-Log -Message ('[{0}] Diff after copy/post-warmup = {1}' -f $topSafe, $postPendingCount) -LogBox $LogBox

        if ($postPendingCount -eq 0) {
            $stableZeroDiffRounds++
            Write-Log -Message ('[{0}] Zero diff after post-warmup. Stability {1}/2' -f $topSafe, $stableZeroDiffRounds) -LogBox $LogBox

            if ($stableZeroDiffRounds -ge 2) {
                Write-Log -Message ('[{0}] Folder is complete and stable.' -f $topSafe) -LogBox $LogBox
                return $true
            }
        }
        else {
            $stableZeroDiffRounds = 0
        }

        if ($lastPendingCount -eq $postPendingCount -and $copiedThisRound -eq 0) {
            Write-Log -Message ('[{0}] No progress detected. Stopping processing of this folder in this run.' -f $topSafe) -LogBox $LogBox
            return $false
        }

        $lastPendingCount = $postPendingCount
    }

    return $false
}

function Run-PerTopFolderSync {
    param(
        [Parameter(Mandatory = $true)]
        $InternalStorageRootFolder,

        [Parameter(Mandatory = $true)]
        [string]$TargetRoot,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogBox,

        [string[]]$ResumeCompletedTopFolders = @(),

        [string]$ResumeLastCompletedTopFolderName = ''
    )

    $totalCopied = 0
    $totalSkipped = 0
    $totalErrors = 0
    $knownTopFolders = [System.Collections.Generic.HashSet[string]]::new()
    $completedTopFolders = [System.Collections.Generic.HashSet[string]]::new()
    $masterStableRounds = 0
    $emptyDiscoveryRounds = 0

    foreach ($completedTopFolder in @($ResumeCompletedTopFolders)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$completedTopFolder)) {
            $null = $completedTopFolders.Add([string]$completedTopFolder)
        }
    }

    Ensure-DirectoryExists -Path $TargetRoot

    if ($completedTopFolders.Count -gt 0) {
        if ([string]::IsNullOrWhiteSpace($ResumeLastCompletedTopFolderName)) {
            Write-Log -Message ('Resume mode: skipping {0} previously completed top folder(s).' -f $completedTopFolders.Count) -LogBox $LogBox
        }
        else {
            Write-Log -Message ('Resume mode: continuing after top folder: {0}' -f $ResumeLastCompletedTopFolderName) -LogBox $LogBox
        }
    }

    for ($masterRound = 1; $masterRound -le $script:Config.MasterDiscoveryMaxRounds; $masterRound++) {
        if ($script:CancelRequested) {
            Write-Log -Message 'Cancelled by user.' -LogBox $LogBox
            break
        }

        Write-Log -Message ('================ Master round {0} =================' -f $masterRound) -LogBox $LogBox
        Set-Status -Text ('Master round {0}: discovering top folders...' -f $masterRound)

        $freshInternalStorageRootFolder = Get-IPhoneInternalStorageFolder
        if ($null -ne $freshInternalStorageRootFolder) {
            $InternalStorageRootFolder = $freshInternalStorageRootFolder
        }
        else {
            Write-Log -Message 'Could not refresh iPhone Internal Storage before discovery. Using existing shell reference.' -LogBox $LogBox
        }

        $topFolders = Get-TopFoldersFromInternalStorage -RootFolder $InternalStorageRootFolder

        if ($topFolders.Count -eq 0 -and $knownTopFolders.Count -eq 0) {
            $emptyDiscoveryRounds++
            Write-Log -Message ('Master round {0}: iPhone returned 0 top folders. Empty discovery {1}/{2}' -f $masterRound, $emptyDiscoveryRounds, $script:Config.MasterEmptyDiscoveryTolerance) -LogBox $LogBox
            Write-Log -Message 'If the iPhone asks whether to trust or allow this PC, unlock it and approve the USB connection.' -LogBox $LogBox
            Set-Status -Text 'Waiting for iPhone approval/unlock...'

            if ($emptyDiscoveryRounds -lt $script:Config.MasterEmptyDiscoveryTolerance) {
                Start-Sleep -Seconds 2
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }
        }
        else {
            $emptyDiscoveryRounds = 0
        }

        $newTopFoldersThisRound = 0
        foreach ($top in $topFolders) {
            if (-not $knownTopFolders.Contains($top.SafeName)) {
                $null = $knownTopFolders.Add($top.SafeName)
                $newTopFoldersThisRound++
                Write-Log -Message ('New top folder discovered: {0}' -f $top.Name) -LogBox $LogBox
            }
        }

        $pendingTopFolders = @($topFolders | Where-Object { -not $completedTopFolders.Contains($_.SafeName) })
        Write-Log -Message ('Master round {0}: top folders seen={1}, new={2}, remaining={3}' -f $masterRound, $topFolders.Count, $newTopFoldersThisRound, $pendingTopFolders.Count) -LogBox $LogBox

        if ($pendingTopFolders.Count -eq 0 -and $newTopFoldersThisRound -eq 0) {
            $masterStableRounds++
            Write-Log -Message ('No new or unfinished top folders. Master stability {0}/2' -f $masterStableRounds) -LogBox $LogBox

            if ($masterStableRounds -ge 2) {
                Write-Log -Message 'No new top folders and all folders are stable. Sync complete.' -LogBox $LogBox
                break
            }

            continue
        }
        else {
            $masterStableRounds = 0
        }

        foreach ($top in $pendingTopFolders) {
            if ($script:CancelRequested) {
                break
            }

            $finished = Process-OneTopFolderUntilStable `
                -InternalStorageRootFolder $InternalStorageRootFolder `
                -TopFolderInfo $top `
                -TargetRoot $TargetRoot `
                -LogBox $LogBox `
                -TotalCopied ([ref]$totalCopied) `
                -TotalSkipped ([ref]$totalSkipped) `
                -TotalErrors ([ref]$totalErrors) `
                -TopFoldersSeen $knownTopFolders.Count `
                -TopFoldersDone $completedTopFolders.Count

            if ($finished) {
                if (-not $completedTopFolders.Contains($top.SafeName)) {
                    $null = $completedTopFolders.Add($top.SafeName)
                }

                Update-PendingSyncProgress -TopFolderName $top.Name -TopFolderSafeName $top.SafeName
            }

            Set-Stats -TopFoldersSeen $knownTopFolders.Count -TopFoldersDone $completedTopFolders.Count -FilesSeen 0 -Pending 0 -Copied $totalCopied -Skipped $totalSkipped -Errors $totalErrors
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    Set-Status -Text 'Ready'

    if (-not $script:CancelRequested) {
        Clear-PendingSyncState
    }
}

# =========================
# GUI
# =========================
Initialize-AppConfig

$form = New-Object System.Windows.Forms.Form
$form.Text = 'iPhone Media Sync'
$form.Size = New-Object System.Drawing.Size(900, 720)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $false

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Location = New-Object System.Drawing.Point(20, 20)
$lblSource.Size = New-Object System.Drawing.Size(260, 20)
$lblSource.Text = 'Source on iPhone (Internal Storage):'
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(20, 45)
$txtSource.Size = New-Object System.Drawing.Size(690, 25)
$txtSource.ReadOnly = $true
$form.Controls.Add($txtSource)

$btnDetect = New-Object System.Windows.Forms.Button
$btnDetect.Location = New-Object System.Drawing.Point(730, 43)
$btnDetect.Size = New-Object System.Drawing.Size(130, 28)
$btnDetect.Text = 'Find iPhone'
$form.Controls.Add($btnDetect)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(20, 90)
$lblTarget.Size = New-Object System.Drawing.Size(220, 20)
$lblTarget.Text = 'Target folder on disk:'
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(20, 115)
$txtTarget.Size = New-Object System.Drawing.Size(690, 25)
$form.Controls.Add($txtTarget)

$btnBrowseTarget = New-Object System.Windows.Forms.Button
$btnBrowseTarget.Location = New-Object System.Drawing.Point(730, 113)
$btnBrowseTarget.Size = New-Object System.Drawing.Size(130, 28)
$btnBrowseTarget.Text = 'Select target'
$form.Controls.Add($btnBrowseTarget)

$btnSync = New-Object System.Windows.Forms.Button
$btnSync.Location = New-Object System.Drawing.Point(20, 160)
$btnSync.Size = New-Object System.Drawing.Size(120, 35)
$btnSync.Text = 'Start sync'
$form.Controls.Add($btnSync)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(155, 160)
$btnCancel.Size = New-Object System.Drawing.Size(120, 35)
$btnCancel.Text = 'Cancel'
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(290, 160)
$btnClose.Size = New-Object System.Drawing.Size(120, 35)
$btnClose.Text = 'Close'
$form.Controls.Add($btnClose)

$lblStatusTitle = New-Object System.Windows.Forms.Label
$lblStatusTitle.Location = New-Object System.Drawing.Point(20, 235)
$lblStatusTitle.Size = New-Object System.Drawing.Size(60, 20)
$lblStatusTitle.Text = 'Status:'
$form.Controls.Add($lblStatusTitle)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Location = New-Object System.Drawing.Point(85, 235)
$script:lblStatus.Size = New-Object System.Drawing.Size(775, 20)
$script:lblStatus.Text = 'Ready'
$form.Controls.Add($script:lblStatus)

$script:lblGlobalStats = New-Object System.Windows.Forms.Label
$script:lblGlobalStats.Location = New-Object System.Drawing.Point(20, 260)
$script:lblGlobalStats.Size = New-Object System.Drawing.Size(840, 20)
$script:lblGlobalStats.Text = 'Global folders - Seen: 0 | Completed: 0 || Global files - Copied: 0 | Skipped: 0 | Errors: 0'
$form.Controls.Add($script:lblGlobalStats)

$script:lblFolderStats = New-Object System.Windows.Forms.Label
$script:lblFolderStats.Location = New-Object System.Drawing.Point(20, 285)
$script:lblFolderStats.Size = New-Object System.Drawing.Size(840, 20)
$script:lblFolderStats.Text = 'Current folder files - Seen: 0 | Pending/Diff: 0'
$form.Controls.Add($script:lblFolderStats)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 315)
$txtLog.Size = New-Object System.Drawing.Size(840, 335)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)
$script:txtLog = $txtLog

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = 'Select target folder'

function Set-UiRunningState {
    param(
        [bool]$IsRunning
    )

    $btnSync.Enabled = -not $IsRunning
    $btnDetect.Enabled = -not $IsRunning
    $btnBrowseTarget.Enabled = -not $IsRunning
    $btnCancel.Enabled = $IsRunning
}

if ($null -ne $script:AppConfig) {
    if (-not [string]::IsNullOrWhiteSpace($script:AppConfig.LastSourceDisplay)) {
        $txtSource.Text = $script:AppConfig.LastSourceDisplay
    }

    if (-not [string]::IsNullOrWhiteSpace($script:AppConfig.LastTargetFolder)) {
        $txtTarget.Text = $script:AppConfig.LastTargetFolder
    }
}

$btnDetect.Add_Click({
    $txtLog.Clear()
    Write-Log -Message 'Looking for Apple iPhone...' -LogBox $txtLog
    Set-Status -Text 'Looking for iPhone...'

    try {
        $sourceFolder = Get-IPhoneInternalStorageFolder

        if ($null -eq $sourceFolder) {
            [System.Windows.Forms.MessageBox]::Show(
                'Could not find Apple iPhone\Internal Storage. Make sure the iPhone is connected, unlocked, and trusted on this PC.',
                'iPhone not found',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            Write-Log -Message 'Could not find Internal Storage on the iPhone.' -LogBox $txtLog
            $txtSource.Text = ''
            $script:IPhoneSourceFolder = $null
            Update-AppConfigFromUi -SourceDisplay '' -TargetFolder $txtTarget.Text
            Set-Status -Text 'iPhone not found'
            return
        }

        $script:IPhoneSourceFolder = $sourceFolder
        $txtSource.Text = 'Apple iPhone\Internal Storage'
        Update-AppConfigFromUi -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text
        Write-Log -Message 'Found iPhone Internal Storage.' -LogBox $txtLog
        Set-Status -Text 'iPhone found'
    }
    catch {
        Write-Log -Message ('Error while searching for iPhone: {0}' -f $_.Exception.Message) -LogBox $txtLog
        Set-Status -Text 'Error while searching for iPhone'
    }
})

$btnBrowseTarget.Add_Click({
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtTarget.Text = $folderBrowser.SelectedPath
        Update-AppConfigFromUi -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text
    }
})

$btnCancel.Add_Click({
    $script:CancelRequested = $true
    Write-Log -Message 'Cancel requested by user...' -LogBox $txtLog
    Set-Status -Text 'Cancelling...'
})

$btnSync.Add_Click({
    if ($null -eq $script:IPhoneSourceFolder) {
        [System.Windows.Forms.MessageBox]::Show(
            'You must click "Find iPhone" first.',
            'Missing source',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($txtTarget.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            'You must select a target folder.',
            'Missing target folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $targetRoot = $txtTarget.Text
    $resumeState = Get-PendingSyncStateForContext -SourceDisplay $txtSource.Text -TargetFolder $targetRoot
    $resumeCompletedTopFolders = @()
    $resumeLastCompletedTopFolderName = ''

    if ($null -ne $resumeState) {
        $lastCompletedTopFolderName = [string]$resumeState.LastCompletedTopFolderName
        $updatedAtText = [string]$resumeState.UpdatedAt
        $resumeMessage = if ([string]::IsNullOrWhiteSpace($lastCompletedTopFolderName)) {
            "A previous sync did not finish.`r`n`r`nDo you want to resume the unfinished sync?"
        }
        else {
            "A previous sync did not finish.`r`nLast completed top folder: $lastCompletedTopFolderName`r`nSaved: $updatedAtText`r`n`r`nDo you want to resume from there?"
        }

        $resumeChoice = [System.Windows.Forms.MessageBox]::Show(
            $resumeMessage,
            'Resume previous sync?',
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($resumeChoice -eq [System.Windows.Forms.DialogResult]::Cancel) {
            return
        }

        if ($resumeChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
            $resumeCompletedTopFolders = @($resumeState.CompletedTopFolders)
            $resumeLastCompletedTopFolderName = [string]$resumeState.LastCompletedTopFolderName
        }
        else {
            Clear-PendingSyncState
        }
    }

    try {
        $script:IsRunning = $true
        $script:CancelRequested = $false
        Set-UiRunningState -IsRunning $true
        $txtLog.Clear()

        Ensure-DirectoryExists -Path $targetRoot
        Update-AppConfigFromUi -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text
        Start-PendingSyncTracking -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text

        Write-Log -Message ('Starting sync to: {0}' -f $targetRoot) -LogBox $txtLog
        Write-Log -Message ('Using config file: {0}' -f $script:AppConfigPath) -LogBox $txtLog
        Write-Log -Message 'Strategy: one top folder at a time -> warm-up -> diff -> copy -> post-warmup -> stable.' -LogBox $txtLog

        Run-PerTopFolderSync `
            -InternalStorageRootFolder $script:IPhoneSourceFolder `
            -TargetRoot $targetRoot `
            -LogBox $txtLog `
            -ResumeCompletedTopFolders $resumeCompletedTopFolders `
            -ResumeLastCompletedTopFolderName $resumeLastCompletedTopFolderName

        if ($script:CancelRequested) {
            [System.Windows.Forms.MessageBox]::Show(
                'Sync was cancelled.',
                'Cancelled',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                'Sync completed.',
                'Completed',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
    }
    catch {
        Write-Log -Message ('Unexpected error: {0}' -f $_.Exception.Message) -LogBox $txtLog

        [System.Windows.Forms.MessageBox]::Show(
            ('Error: {0}' -f $_.Exception.Message),
            'Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $script:IsRunning = $false
        Set-UiRunningState -IsRunning $false
        Set-Status -Text 'Ready'
    }
})

$btnClose.Add_Click({
    Update-AppConfigFromUi -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text

    if ($script:IsRunning) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            'Sync is still running. Do you want to cancel and close?',
            'Confirm',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:CancelRequested = $true
            $form.Close()
        }

        return
    }

    $form.Close()
})

$form.Add_FormClosing({
    Update-AppConfigFromUi -SourceDisplay $txtSource.Text -TargetFolder $txtTarget.Text
})

[void]$form.ShowDialog()
