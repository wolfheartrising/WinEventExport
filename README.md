# Windows Event Log Export & Maintenance Script

![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-Custom-lightgrey)
[![Signed Execution Optional](https://img.shields.io/badge/Execution-Signing%20Optional-green)](https://github.com/wolfheartrising/PSScriptSigner/blob/main/PSScriptSigner.ps1)

## Overview

This PowerShell script exports selected Windows Event Logs to CSV files and automatically maintains them by removing entries older than a user-defined retention period. It is designed to run manually or be scheduled with Windows Task Scheduler for hands-free operation.

## Features

- Export logs: `Application`, `Security`, `Setup`, `System`, and `ForwardedEvents`
- Appends new events to CSV files without overwriting
- Automatically trims old entries from each log file
- GUI folder picker to choose output directory on first run
- Configurable retention period (e.g., 7, 30, 90 days)
- Summary log tracks skipped categories and successful exports
- Retention-aware cleanup of both event data and summary log entries
- Persistent configuration saved to `config.json`

## Setup Instructions

### Initial Run
1. Launch the script in PowerShell with GUI support.
2. Select the output directory using the folder picker window.
3. Enter the number of days of logs to retain (e.g., 30).
4. Settings will be saved to `config.json` for future runs.

### Future Runs
- The script reads from `config.json` and runs automatically.
- To reconfigure, delete `config.json` and rerun the script.

## Scheduling the Script

This script is designed to be run on a recurring basis using **Windows Task Scheduler**:

1. Open **Task Scheduler** and create a new task.

2. Set the task to run with **highest privileges** and specify **PowerShell** as the action.

3. Use the following action command:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ExportWinEventLog.ps1"
Alternate Step 3 (for enhanced security):
If you'd prefer to maintain stricter security and avoid policy bypasses, consider signing the script using PSScriptSigner — a PowerShell script signing tool created by Foresta.
This lets you run the script under a more secure execution policy (e.g., AllSigned) while ensuring integrity and trust for scheduled tasks and shared environments.
