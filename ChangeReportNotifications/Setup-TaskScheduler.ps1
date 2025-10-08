#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Setup script for Windows Task Scheduler integration with Change Report Email Notifications

.DESCRIPTION
    This script creates and configures a Windows scheduled task to run the Change Report
    Email Notifications system daily. It includes validation and testing functionality.

.PARAMETER InstallPath
    The path where the Change Report system is installed. Default: C:\ChangeReportNotifications

.PARAMETER TaskName
    The name for the scheduled task. Default: "Change Report Email Notifications"

.PARAMETER ExecutionTime
    The daily execution time in HH:MM format. If not specified, reads from config file.

.PARAMETER UserAccount
    The user account to run the task under. Default: current user

.PARAMETER TestOnly
    Only test the existing scheduled task without creating or modifying it

.PARAMETER Remove
    Remove the existing scheduled task

.EXAMPLE
    .\Setup-TaskScheduler.ps1
    
.EXAMPLE
    .\Setup-TaskScheduler.ps1 -InstallPath "D:\Scripts\ChangeReports" -ExecutionTime "09:30"
    
.EXAMPLE
    .\Setup-TaskScheduler.ps1 -TestOnly
    
.EXAMPLE
    .\Setup-TaskScheduler.ps1 -Remove
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath = "C:\ChangeReportNotifications",
    
    [Parameter(Mandatory = $false)]
    [string]$TaskName = "Change Report Email Notifications",
    
    [Parameter(Mandatory = $false)]
    [string]$ExecutionTime,
    
    [Parameter(Mandatory = $false)]
    [string]$UserAccount = $env:USERNAME,
    
    [Parameter(Mandatory = $false)]
    [switch]$TestOnly,
    
    [Parameter(Mandatory = $false)]
    [switch]$Remove
)

function Write-TaskLog {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Error" { Write-Host $logMessage -ForegroundColor Red }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        default { Write-Host $logMessage -ForegroundColor White }
    }
}

function Test-Prerequisites {
    Write-TaskLog "Checking prerequisites for Task Scheduler setup..."
    
    # Check if running as Administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-TaskLog "This script must be run as Administrator to manage scheduled tasks" -Level "Error"
        return $false
    }
    
    # Check if installation path exists
    if (-not (Test-Path $InstallPath)) {
        Write-TaskLog "Installation path not found: $InstallPath" -Level "Error"
        Write-TaskLog "Please run the installation script first or specify correct path" -Level "Error"
        return $false
    }
    
    # Check if main script exists
    $mainScript = "$InstallPath\Send-ChangeReport.ps1"
    if (-not (Test-Path $mainScript)) {
        Write-TaskLog "Main script not found: $mainScript" -Level "Error"
        return $false
    }
    
    # Check if configuration file exists
    $configFile = "$InstallPath\config\config.json"
    if (-not (Test-Path $configFile)) {
        Write-TaskLog "Configuration file not found: $configFile" -Level "Error"
        Write-TaskLog "Please run the installation script to create configuration" -Level "Error"
        return $false
    }
    
    Write-TaskLog "Prerequisites check completed successfully" -Level "Success"
    return $true
}

function Get-ConfigurationTime {
    param([string]$ConfigPath)
    
    try {
        $configFile = "$ConfigPath\config\config.json"
        $config = Get-Content $configFile -Raw | ConvertFrom-Json
        
        if ($config.Schedule -and $config.Schedule.ExecutionTime) {
            return $config.Schedule.ExecutionTime
        }
        else {
            Write-TaskLog "No execution time found in configuration file" -Level "Warning"
            return $null
        }
    }
    catch {
        Write-TaskLog "Failed to read execution time from configuration: $($_.Exception.Message)" -Level "Warning"
        return $null
    }
}

function Get-ExecutionTime {
    param([string]$ConfigPath)
    
    # Use parameter if provided
    if (-not [string]::IsNullOrEmpty($ExecutionTime)) {
        if ($ExecutionTime -match "^([01]?[0-9]|2[0-3]):[0-5][0-9]$") {
            return $ExecutionTime
        }
        else {
            Write-TaskLog "Invalid time format: $ExecutionTime. Use HH:MM format (e.g., 08:30)" -Level "Error"
            return $null
        }
    }
    
    # Try to get from configuration file
    $configTime = Get-ConfigurationTime -ConfigPath $ConfigPath
    if ($configTime) {
        if ($configTime -match "^([01]?[0-9]|2[0-3]):[0-5][0-9]$") {
            return $configTime
        }
    }
    
    # Prompt user for time
    do {
        $userTime = Read-Host "Enter daily execution time (HH:MM format, e.g., 08:30)"
        if ($userTime -match "^([01]?[0-9]|2[0-3]):[0-5][0-9]$") {
            return $userTime
        }
        else {
            Write-TaskLog "Invalid time format. Please use HH:MM format (e.g., 08:30)" -Level "Warning"
        }
    } while ($true)
}

function Test-ExistingTask {
    param([string]$TaskName)
    
    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        return $task -ne $null
    }
    catch {
        return $false
    }
}

function Remove-ExistingTask {
    param([string]$TaskName)
    
    Write-TaskLog "Removing existing scheduled task: $TaskName"
    
    try {
        if (Test-ExistingTask -TaskName $TaskName) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-TaskLog "Successfully removed scheduled task: $TaskName" -Level "Success"
            return $true
        }
        else {
            Write-TaskLog "Scheduled task not found: $TaskName" -Level "Warning"
            return $true
        }
    }
    catch {
        Write-TaskLog "Failed to remove scheduled task: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function New-ScheduledTaskDefinition {
    param(
        [string]$TaskName,
        [string]$InstallPath,
        [string]$ExecutionTime,
        [string]$UserAccount
    )
    
    Write-TaskLog "Creating scheduled task definition..."
    
    try {
        # Parse execution time
        $timeParts = $ExecutionTime -split ":"
        $hour = [int]$timeParts[0]
        $minute = [int]$timeParts[1]
        
        # Create task action
        $scriptPath = "$InstallPath\Send-ChangeReport.ps1"
        $batchPath = "$InstallPath\Run-ChangeReport.bat"
        
        # Use batch file if it exists, otherwise use PowerShell directly
        if (Test-Path $batchPath) {
            $action = New-ScheduledTaskAction -Execute $batchPath -WorkingDirectory $InstallPath
        }
        else {
            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WorkingDirectory $InstallPath
        }
        
        # Create task trigger (daily at specified time)
        $trigger = New-ScheduledTaskTrigger -Daily -At (Get-Date).Date.AddHours($hour).AddMinutes($minute)
        
        # Create task settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable
        
        # Create task principal (user context)
        $principal = New-ScheduledTaskPrincipal -UserId $UserAccount -LogonType Interactive -RunLevel Highest
        
        # Register the scheduled task
        $task = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Daily email notifications for critical and high priority change records"
        
        Write-TaskLog "Successfully created scheduled task: $TaskName" -Level "Success"
        Write-TaskLog "Task will run daily at $ExecutionTime as user: $UserAccount" -Level "Success"
        
        return $task
    }
    catch {
        Write-TaskLog "Failed to create scheduled task: $($_.Exception.Message)" -Level "Error"
        return $null
    }
}

function Test-ScheduledTask {
    param([string]$TaskName)
    
    Write-TaskLog "Testing scheduled task: $TaskName"
    
    try {
        # Get the task
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        
        # Display task information
        Write-TaskLog "Task Name: $($task.TaskName)" -Level "Success"
        Write-TaskLog "Task State: $($task.State)" -Level "Success"
        Write-TaskLog "Last Run Time: $($task.LastRunTime)" -Level "Success"
        Write-TaskLog "Next Run Time: $($task.NextRunTime)" -Level "Success"
        
        # Get task info
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-TaskLog "Last Task Result: $($taskInfo.LastTaskResult)" -Level "Success"
        
        # Test run the task
        Write-TaskLog "Starting test run of the scheduled task..."
        Start-ScheduledTask -TaskName $TaskName
        
        # Wait a moment and check status
        Start-Sleep -Seconds 5
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-TaskLog "Task execution initiated. Check logs for results." -Level "Success"
        
        return $true
    }
    catch {
        Write-TaskLog "Failed to test scheduled task: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Show-TaskInformation {
    param([string]$TaskName)
    
    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
        
        Write-Host "`n=== Scheduled Task Information ===" -ForegroundColor Cyan
        Write-Host "Task Name: $($task.TaskName)" -ForegroundColor White
        Write-Host "Description: $($task.Description)" -ForegroundColor White
        Write-Host "State: $($task.State)" -ForegroundColor White
        Write-Host "Author: $($task.Author)" -ForegroundColor White
        Write-Host "User ID: $($task.Principal.UserId)" -ForegroundColor White
        Write-Host "Run Level: $($task.Principal.RunLevel)" -ForegroundColor White
        Write-Host "Last Run Time: $($taskInfo.LastRunTime)" -ForegroundColor White
        Write-Host "Next Run Time: $($taskInfo.NextRunTime)" -ForegroundColor White
        Write-Host "Last Task Result: $($taskInfo.LastTaskResult)" -ForegroundColor White
        Write-Host "Number of Missed Runs: $($taskInfo.NumberOfMissedRuns)" -ForegroundColor White
        
        # Show triggers
        Write-Host "`n--- Triggers ---" -ForegroundColor Green
        foreach ($trigger in $task.Triggers) {
            Write-Host "Type: $($trigger.CimClass.CimClassName)" -ForegroundColor White
            if ($trigger.StartBoundary) {
                Write-Host "Start Time: $($trigger.StartBoundary)" -ForegroundColor White
            }
            if ($trigger.DaysInterval) {
                Write-Host "Repeat: Every $($trigger.DaysInterval) day(s)" -ForegroundColor White
            }
        }
        
        # Show actions
        Write-Host "`n--- Actions ---" -ForegroundColor Green
        foreach ($action in $task.Actions) {
            Write-Host "Execute: $($action.Execute)" -ForegroundColor White
            if ($action.Arguments) {
                Write-Host "Arguments: $($action.Arguments)" -ForegroundColor White
            }
            if ($action.WorkingDirectory) {
                Write-Host "Working Directory: $($action.WorkingDirectory)" -ForegroundColor White
            }
        }
        
        return $true
    }
    catch {
        Write-TaskLog "Failed to retrieve task information: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Validate-TaskConfiguration {
    param([string]$TaskName, [string]$InstallPath)
    
    Write-TaskLog "Validating task configuration..."
    
    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        
        # Check if task is enabled
        if ($task.State -ne "Ready") {
            Write-TaskLog "Task is not in Ready state. Current state: $($task.State)" -Level "Warning"
        }
        
        # Check if the script file exists
        $scriptExists = $false
        foreach ($action in $task.Actions) {
            if ($action.Execute -like "*PowerShell*" -or $action.Execute -like "*Send-ChangeReport*" -or $action.Execute -like "*Run-ChangeReport*") {
                $scriptExists = $true
                break
            }
        }
        
        if (-not $scriptExists) {
            Write-TaskLog "Task action does not appear to reference the correct script" -Level "Warning"
        }
        
        # Check working directory
        $correctWorkingDir = $false
        foreach ($action in $task.Actions) {
            if ($action.WorkingDirectory -eq $InstallPath) {
                $correctWorkingDir = $true
                break
            }
        }
        
        if (-not $correctWorkingDir) {
            Write-TaskLog "Task working directory may not be set correctly" -Level "Warning"
        }
        
        Write-TaskLog "Task configuration validation completed" -Level "Success"
        return $true
    }
    catch {
        Write-TaskLog "Failed to validate task configuration: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

# Main execution
try {
    Write-Host "=== Change Report Email Notifications - Task Scheduler Setup ===" -ForegroundColor Cyan
    Write-Host "This script will configure Windows Task Scheduler for automated execution.`n" -ForegroundColor Yellow
    
    # Handle remove operation
    if ($Remove) {
        if (Remove-ExistingTask -TaskName $TaskName) {
            Write-Host "`nScheduled task removed successfully." -ForegroundColor Green
        }
        else {
            Write-Host "`nFailed to remove scheduled task." -ForegroundColor Red
            exit 1
        }
        exit 0
    }
    
    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        Write-TaskLog "Prerequisites check failed. Setup aborted." -Level "Error"
        exit 1
    }
    
    # Handle test-only operation
    if ($TestOnly) {
        if (Test-ExistingTask -TaskName $TaskName) {
            Show-TaskInformation -TaskName $TaskName
            Validate-TaskConfiguration -TaskName $TaskName -InstallPath $InstallPath
            Test-ScheduledTask -TaskName $TaskName
        }
        else {
            Write-TaskLog "Scheduled task not found: $TaskName" -Level "Error"
            Write-TaskLog "Please run the setup script without -TestOnly to create the task" -Level "Info"
            exit 1
        }
        exit 0
    }
    
    # Get execution time
    $execTime = Get-ExecutionTime -ConfigPath $InstallPath
    if (-not $execTime) {
        Write-TaskLog "Failed to determine execution time. Setup aborted." -Level "Error"
        exit 1
    }
    
    # Check if task already exists
    if (Test-ExistingTask -TaskName $TaskName) {
        Write-TaskLog "Scheduled task already exists: $TaskName" -Level "Warning"
        $response = Read-Host "Do you want to replace the existing task? (Y/N)"
        if ($response -eq "Y" -or $response -eq "y") {
            if (-not (Remove-ExistingTask -TaskName $TaskName)) {
                Write-TaskLog "Failed to remove existing task. Setup aborted." -Level "Error"
                exit 1
            }
        }
        else {
            Write-TaskLog "Setup cancelled by user." -Level "Info"
            exit 0
        }
    }
    
    # Create the scheduled task
    $task = New-ScheduledTaskDefinition -TaskName $TaskName -InstallPath $InstallPath -ExecutionTime $execTime -UserAccount $UserAccount
    
    if ($task) {
        Write-Host "`n=== Task Scheduler Setup Completed Successfully ===" -ForegroundColor Green
        
        # Show task information
        Show-TaskInformation -TaskName $TaskName
        
        # Validate configuration
        Validate-TaskConfiguration -TaskName $TaskName -InstallPath $InstallPath
        
        Write-Host "`nNext Steps:" -ForegroundColor Yellow
        Write-Host "1. Test the scheduled task: .\Setup-TaskScheduler.ps1 -TestOnly" -ForegroundColor White
        Write-Host "2. Monitor the first few executions in Task Scheduler" -ForegroundColor White
        Write-Host "3. Check log files in: $InstallPath\logs" -ForegroundColor White
        Write-Host "4. Verify email delivery to configured recipients" -ForegroundColor White
        
        Write-Host "`nTask Scheduler Management:" -ForegroundColor Yellow
        Write-Host "- View task: Get-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
        Write-Host "- Run manually: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
        Write-Host "- Disable task: Disable-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
        Write-Host "- Remove task: .\Setup-TaskScheduler.ps1 -Remove" -ForegroundColor White
    }
    else {
        Write-TaskLog "Failed to create scheduled task. Setup failed." -Level "Error"
        exit 1
    }
}
catch {
    Write-TaskLog "Task Scheduler setup failed with error: $($_.Exception.Message)" -Level "Error"
    Write-TaskLog "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    exit 1
}