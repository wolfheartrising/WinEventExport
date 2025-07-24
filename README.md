#Windows Event Log Export & Maintenance Script

![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-Custom-lightgrey)

## Overview

This PowerShell script exports selected Windows Event Logs to CSV files and maintains them by removing entries older than a user-defined retention period. Designed for easy automation via Windows Task Scheduler.

## Features

- Export logs: `Application`, `Security`, `Setup`, `System`, `ForwardedEvents`
- Appends new events to CSV files without overwriting
- Automatically trims old entries from each log file
- GUI folder picker and custom retention prompt on first run
- Logs export status and skipped logs in a summary file
- Persistent configuration stored in `config.json`

## Setup Instructions

### Initial Run
1. Launch the script in PowerShell with GUI support.
2. Select the output directory using the folder picker.
3. Enter the number of days of logs to retain (e.g., 30).
4. Settings will be saved to `config.json`.

### Future Runs
- The script reads from `config.json` and runs automatically.
- To reset configuration, delete `config.json`.

### Automating with Task Scheduler
1. Open **Task Scheduler** and create a new task.
2. Set to run with highest privileges and specify PowerShell as the action.
3. Use command:  
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ExportWinEventLog.ps1"# WinEventExport
This script exports selected Windows Event Logs to CSV format and maintains them by automatically removing entries older than a user-defined retention period.
