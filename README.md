# iPhone Media Sync

A Windows desktop utility for copying media files from an iPhone's **Internal Storage** to a folder on your PC in a more reliable way than ordinary drag-and-drop.

This tool is designed for scenarios where iPhone files and folders appear gradually through the Windows Shell interface, causing normal copy operations to miss files or stop too early.


## Disclaimer

This tool interacts with iPhone media through the Windows Shell interface, which can behave differently across systems and devices. 
Always test with a safe destination folder first and verify results and formats BEFORE deleting anything from the iPhone.

## Features

- Detects **Apple iPhone > Internal Storage** through Windows Shell
- Lets you select a local target folder on disk
- Processes **one top-level iPhone folder at a time**
- Performs repeated **warm-up scans** before and after copy operations
- Builds a file index and compares source vs destination
- Copies only missing or changed files
- Shows live **status**, **statistics**, and **log output**
- Supports **Cancel** during scanning and copying
- Saves the last used source display and target folder in a local config file

---

## Why this tool exists

When browsing an iPhone through Windows Explorer, the content exposed through MTP / Shell can be incomplete at first. Files and folders may "show up" gradually, especially when there is a large photo or video library.

A normal copy job may therefore:

- start too early
- miss files that appear later
- stop with an incomplete result
- behave inconsistently between runs

This application tries to reduce that problem by:

1. scanning one top-level folder at a time
2. repeating scans until the folder looks stable
3. comparing the detected source files with the destination
4. copying only the files still missing or different
5. scanning again after copy to catch late-appearing files

---

## How it works

The sync flow is roughly:

1. Detect iPhone Internal Storage
2. Enumerate top-level folders
3. Pick one top-level folder
4. Warm up and scan it multiple times
5. Build a source index
6. Compare with the destination folder
7. Copy pending files
8. Warm up and scan again
9. Repeat until the folder is considered stable
10. Continue with the next top-level folder

This approach is slower than a simple bulk copy, but more robust when Windows does not expose all iPhone files immediately.

---

## Requirements

- Windows
- PowerShell
- iPhone connected by USB
- iPhone unlocked
- The PC must be trusted on the iPhone
- Access to **Apple iPhone > Internal Storage** in Windows Explorer

---

## Usage

1. Connect the iPhone to the PC with a cable
2. Unlock the iPhone
3. Confirm **Trust This Computer** if prompted
4. Run the script
5. Click **Find iPhone**
6. Select the target folder on disk
7. Click **Start sync**

During sync, the application will:

- log discovered folders
- show current copy progress
- show file and folder statistics
- allow cancellation

---

## Configuration

The application stores a small JSON config file with the last used values.

Stored values:

- last source display text
- last target folder

Config file location:

- `%APPDATA%\iPhoneMediaSync\config.json`
- fallback: `%TEMP%\iPhoneMediaSync\config.json`

Note that the iPhone source itself is a live Shell object and cannot be fully restored from config, so you still need to click **Find iPhone** after restarting the application.

---

## Current sync strategy

The current implementation uses this strategy:

- one top-level folder at a time
- repeated warm-up passes
- stability detection based on repeated scans
- diff against target
- post-copy warm-up and re-check
- graceful cancel handling

This is intended to improve reliability when iPhone enumeration is delayed or inconsistent.

---

## Limitations

- Windows-only
- Depends on Windows Shell / MTP behavior
- iPhone must remain connected and unlocked during operation
- Enumeration speed and consistency may vary depending on the device, cable, Windows state, and library size
- The source is not mounted as a normal filesystem, so some operations are inherently slower and less predictable than standard file copy

---

## Error handling

The tool includes basic protections for common issues:

- retry logic for folder enumeration
- verification that copied files are stable on disk
- graceful cancel handling during warm-up and copy phases
- logging of enumeration and copy errors

---

## UI overview

The application provides:

- source field for detected iPhone storage
- target folder field
- **Find iPhone**
- **Select target**
- **Start sync**
- **Cancel**
- **Close**
- status line
- statistics line
- log output window

---

## Example use case

This tool is useful when you have:

- a large iPhone photo/video library
- incomplete copy results from Explorer
- folders that appear slowly
- a need to repeat sync until the destination is fully caught up

---

## Running the script

Run the `.ps1` file in PowerShell.

If needed, start PowerShell with an execution policy that allows local scripts, for example:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\iPhoneSync.ps1

