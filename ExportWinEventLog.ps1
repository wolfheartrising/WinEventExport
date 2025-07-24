<#
===============================================================================
    Windows Event Log Export & Maintenance Script
===============================================================================
    Author:          Foresta
    Created:         2025-07-23
    Last Modified:   2025-07-23
    Version:         1.2.0

    Description:
        This script exports selected Windows Event Logs to CSV format and 
        maintains them by automatically removing entries older than a user-
        defined retention period.

        - Appends new event data to persistent CSV logs per category
        - Cleans out entries older than the retention window
        - Allows user to choose output folder and retention period via prompt
        - Logs skipped categories and successful exports in summary log
        - All settings are saved to config.json for future runs

    Usage Notes:
        - On first run, the user will be prompted to select a save location 
          and define the number of days of logs to retain
        - All logs are saved in CSV format in the specified output directory
        - Summary log shows results of each execution and is trimmed by age
        - This script is designed to be scheduled via Windows Task Scheduler 
          for fully automated execution on a recurring basis

    Requirements:
        - Windows PowerShell 5.1 or later
        - Access to Windows Event Logs
        - GUI environment for folder selection dialog
        - Write access to configured output directory

    Version History:
        1.0.0 - Initial release with static 30-day retention
        1.1.0 - Configurable retention period and persistent settings
        1.2.0 - Folder picker, export stats, and summary log with cleanup

    LICENSE:
        This script was developed for internal organizational use.
        You may modify and redistribute it with appropriate credit.

===============================================================================
#>

# Track script start time
$ScriptStartTime = Get-Date

# Path to config file
$ConfigPath = "$PSScriptRoot\config.json"

# Prompt for setup if config is missing
if (-not (Test-Path $ConfigPath)) {
    Write-Host "Initial setup required..."

    # Open folder picker
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select the folder where Windows Event Logs will be saved"
    $folderDialog.ShowNewFolderButton = $true

    $dialogResult = $folderDialog.ShowDialog()
    if ($dialogResult -eq "OK" -and $folderDialog.SelectedPath) {
        $logDir = $folderDialog.SelectedPath
    } else {
        Write-Error "No folder selected. Configuration aborted."
        exit
    }

    # Ask for retention period
    $daysInput = Read-Host "Enter number of days of logs to keep (e.g., 30)"
    if (-not ($daysInput -as [int] -and [int]$daysInput -gt 0)) {
        Write-Error "Invalid input for days. Please enter a positive number."
        exit
    }

    # Save configuration
    $Settings = @{
        LogOutputDirectory = $logDir
        RetentionDays = [int]$daysInput
    } | ConvertTo-Json -Depth 3
    $Settings | Out-File $ConfigPath -Encoding UTF8
    Write-Host "Configuration saved to: $ConfigPath"
}

# Load settings
$Config = Get-Content $ConfigPath | ConvertFrom-Json
$LogOutputDirectory = $Config.LogOutputDirectory
$RetentionDays = $Config.RetentionDays
$CutoffDate = (Get-Date).AddDays(-$RetentionDays)
$SummaryPath = "$LogOutputDirectory\ExportSummary.txt"

# Trim old summary entries
if (Test-Path $SummaryPath) {
    $AllSummaryEntries = Get-Content -Path $SummaryPath
    $FilteredSummary = $AllSummaryEntries | Where-Object {
        $timestamp = $_.Substring(0, 19) -as [datetime]
        $timestamp -gt $CutoffDate
    }
    $FilteredSummary | Set-Content -Path $SummaryPath
}

# Event logs to export
$EventTypesToExport = @('Application', 'Security', 'Setup', 'System', 'ForwardedEvents')

foreach ($EventType in $EventTypesToExport)
{
    $LogOutputTopic = "Windows Event Log - $EventType"
    $LogOutputCSVFilePath = "$LogOutputDirectory\$LogOutputTopic.csv"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    Write-Output "Processing logs for: $EventType"
    Write-Output "Target file: $LogOutputCSVFilePath"

    # Retrieve events newer than cutoff
    $RecentEvents = Get-WinEvent -FilterHashtable @{
        LogName = $EventType
        StartTime = $CutoffDate
    }

    # Log skipped categories
    if ($RecentEvents.Count -eq 0) {
        $logMessage = "$timestamp - No recent events found for log type: $EventType"
        Write-Output $logMessage
        Add-Content -Path $SummaryPath -Value $logMessage
        continue
    }

    # Export to CSV
    if (-Not (Test-Path $LogOutputCSVFilePath)) {
        $RecentEvents | Export-Csv -Path $LogOutputCSVFilePath -NoTypeInformation
        Write-Output "Created new CSV file with recent entries."
    } else {
        $RecentEvents | Export-Csv -Path $LogOutputCSVFilePath -Append -NoTypeInformation
        Write-Output "Appended recent entries to existing file."
    }

    # Clean old entries
    $AllEvents = Import-Csv -Path $LogOutputCSVFilePath
    $FilteredEvents = $AllEvents | Where-Object {
        ($_.'TimeCreated' -as [datetime]) -gt $CutoffDate
    }
    $FilteredEvents | Export-Csv -Path $LogOutputCSVFilePath -NoTypeInformation

    # Log export results
    $exportedCount = $FilteredEvents.Count
    $successLog = "$timestamp - Exported and cleaned $exportedCount event(s) for log type: $EventType"
    Write-Output $successLog
    Add-Content -Path $SummaryPath -Value $successLog
}

# Report execution duration
$ScriptEndTime = Get-Date
$ScriptDuration = New-TimeSpan -Start $ScriptStartTime -End $ScriptEndTime
Write-Output "Log export and cleanup completed in: $ScriptDuration"
