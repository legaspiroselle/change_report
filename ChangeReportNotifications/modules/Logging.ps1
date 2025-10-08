# Logging Module for Change Report Notifications
# Provides comprehensive logging functionality with different log levels

# Define log levels
enum LogLevel {
    Debug = 0
    Info = 1
    Warning = 2
    Error = 3
}

<#
.SYNOPSIS
    Writes a log entry with timestamp and level support
.DESCRIPTION
    Creates a timestamped log entry with the specified level and message
.PARAMETER Message
    The message to log
.PARAMETER Level
    The log level (Debug, Info, Warning, Error)
.PARAMETER LogPath
    The path to the log file directory
.EXAMPLE
    Write-Log -Message "Database connection successful" -Level Info -LogPath "C:\Logs"
#>
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [LogLevel]$Level = [LogLevel]::Info,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    try {
        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        
        # Get current log file name
        $logFileName = Get-LogFileName -Date (Get-Date)
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName
        
        # Create timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        # Format log entry
        $logEntry = "[$timestamp] [$($Level.ToString().ToUpper())] $Message"
        
        # Write to log file
        Add-Content -Path $fullLogPath -Value $logEntry -Encoding UTF8
        
        # Also write to console for immediate feedback
        switch ($Level) {
            ([LogLevel]::Debug) { Write-Verbose $logEntry }
            ([LogLevel]::Info) { Write-Host $logEntry -ForegroundColor Green }
            ([LogLevel]::Warning) { Write-Warning $logEntry }
            ([LogLevel]::Error) { Write-Error $logEntry }
        }
    }
    catch {
        # Fallback to console if file logging fails
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        Write-Host "[$timestamp] [$($Level.ToString().ToUpper())] $Message" -ForegroundColor Yellow
    }
}

<#
.SYNOPSIS
    Initializes a daily log session
.DESCRIPTION
    Creates a new log session for the day and writes session start information
.PARAMETER LogPath
    The path to the log file directory
.EXAMPLE
    Start-LogSession -LogPath "C:\Logs"
#>
function Start-LogSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    try {
        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        
        # Write session start marker
        $sessionId = [System.Guid]::NewGuid().ToString("N").Substring(0, 8)
        $computerName = $env:COMPUTERNAME
        $userName = $env:USERNAME
        
        Write-Log -Message "=== LOG SESSION STARTED ===" -Level Info -LogPath $LogPath
        Write-Log -Message "Session ID: $sessionId" -Level Info -LogPath $LogPath
        Write-Log -Message "Computer: $computerName" -Level Info -LogPath $LogPath
        Write-Log -Message "User: $userName" -Level Info -LogPath $LogPath
        Write-Log -Message "PowerShell Version: $($PSVersionTable.PSVersion)" -Level Info -LogPath $LogPath
        
        return $sessionId
    }
    catch {
        Write-Warning "Failed to start log session: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Generates a date-based log file name
.DESCRIPTION
    Creates a standardized log file name based on the provided date
.PARAMETER Date
    The date to use for the log file name
.EXAMPLE
    Get-LogFileName -Date (Get-Date)
#>
function Get-LogFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [DateTime]$Date
    )
    
    # Format: ChangeReport_YYYY-MM-DD.log
    $dateString = $Date.ToString("yyyy-MM-dd")
    return "ChangeReport_$dateString.log"
}

<#
.SYNOPSIS
    Ends a log session with cleanup information
.DESCRIPTION
    Writes session end information and performs any necessary cleanup
.PARAMETER LogPath
    The path to the log file directory
.PARAMETER SessionId
    The session ID from Start-LogSession
.EXAMPLE
    Stop-LogSession -LogPath "C:\Logs" -SessionId "abc12345"
#>
function Stop-LogSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$SessionId
    )
    
    try {
        if ($SessionId) {
            Write-Log -Message "Session ID: $SessionId" -Level Info -LogPath $LogPath
        }
        Write-Log -Message "=== LOG SESSION ENDED ===" -Level Info -LogPath $LogPath
        # Add empty line for separation
        $logFileName = Get-LogFileName -Date (Get-Date)
        $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName
        Add-Content -Path $fullLogPath -Value "" -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to end log session: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Cleans up old log files
.DESCRIPTION
    Removes log files older than the specified number of days
.PARAMETER LogPath
    The path to the log file directory
.PARAMETER DaysToKeep
    Number of days of log files to retain (default: 30)
.EXAMPLE
    Remove-OldLogFiles -LogPath "C:\Logs" -DaysToKeep 7
#>
function Remove-OldLogFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [int]$DaysToKeep = 30
    )
    
    try {
        if (-not (Test-Path -Path $LogPath)) {
            Write-Log -Message "Log path does not exist: $LogPath" -Level Warning -LogPath $LogPath
            return
        }
        
        $cutoffDate = (Get-Date).AddDays(-$DaysToKeep)
        $logFiles = Get-ChildItem -Path $LogPath -Filter "ChangeReport_*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($logFiles.Count -gt 0) {
            Write-Log -Message "Cleaning up $($logFiles.Count) old log files (older than $DaysToKeep days)" -Level Info -LogPath $LogPath
            
            foreach ($file in $logFiles) {
                Remove-Item -Path $file.FullName -Force
                Write-Log -Message "Removed old log file: $($file.Name)" -Level Debug -LogPath $LogPath
            }
        }
        else {
            Write-Log -Message "No old log files to clean up" -Level Debug -LogPath $LogPath
        }
    }
    catch {
        Write-Log -Message "Failed to clean up old log files: $($_.Exception.Message)" -Level Error -LogPath $LogPath
    }
}

# Functions are available when dot-sourced
# Note: Export-ModuleMember is not needed when dot-sourcing

<#
.SYNOPSIS
    Sends error notification emails for system failures
.DESCRIPTION
    Sends formatted error notification emails when critical system failures occur
.PARAMETER Config
    Configuration object containing email settings
.PARAMETER ErrorMessage
    The error message to include in the notification
.PARAMETER ErrorType
    The type of error (Database, Email, Configuration, etc.)
.PARAMETER LogPath
    Path to log directory for fallback logging
.EXAMPLE
    Send-ErrorNotification -Config $config -ErrorMessage "Database connection failed" -ErrorType "Database" -LogPath "C:\Logs"
#>
function Send-ErrorNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,
        
        [Parameter(Mandatory = $true)]
        [string]$ErrorType,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    try {
        # Log the error first
        Write-Log -Message "SYSTEM ERROR [$ErrorType]: $ErrorMessage" -Level Error -LogPath $LogPath
        
        # Check if we have email configuration
        if (-not $Config.Email) {
            Write-Log -Message "No email configuration available for error notification" -Level Warning -LogPath $LogPath
            return
        }
        
        # Generate error notification email
        $subject = "ALERT: Change Report System Error - $ErrorType"
        $body = Get-ErrorEmailTemplate -ErrorType $ErrorType -ErrorMessage $ErrorMessage -Timestamp (Get-Date)
        
        # Attempt to send error notification
        $emailParams = @{
            SmtpServer = $Config.Email.SMTPServer
            Port = $Config.Email.Port
            UseSsl = $Config.Email.EnableSSL
            From = $Config.Email.From
            To = $Config.Email.To
            Subject = $subject
            Body = $body
            BodyAsHtml = $true
        }
        
        # Add authentication if configured
        if ($Config.Email.Username -and $Config.Email.Password) {
            $credential = New-Object System.Management.Automation.PSCredential($Config.Email.Username, $Config.Email.Password)
            $emailParams.Credential = $credential
        }
        
        # Send the error notification
        Send-MailMessage @emailParams
        Write-Log -Message "Error notification sent successfully" -Level Info -LogPath $LogPath
    }
    catch {
        # If email fails, log to file as fallback
        Write-Log -Message "CRITICAL: Failed to send error notification email: $($_.Exception.Message)" -Level Error -LogPath $LogPath
        Write-Log -Message "Original error was: [$ErrorType] $ErrorMessage" -Level Error -LogPath $LogPath
        
        # Write error details to a separate critical error file
        $criticalErrorFile = Join-Path -Path $LogPath -ChildPath "CRITICAL_ERRORS.log"
        $criticalEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - UNDELIVERED ERROR NOTIFICATION`n"
        $criticalEntry += "Error Type: $ErrorType`n"
        $criticalEntry += "Error Message: $ErrorMessage`n"
        $criticalEntry += "Email Failure: $($_.Exception.Message)`n"
        $criticalEntry += "----------------------------------------`n"
        
        Add-Content -Path $criticalErrorFile -Value $criticalEntry -Encoding UTF8
    }
}

<#
.SYNOPSIS
    Generates HTML email template for error notifications
.DESCRIPTION
    Creates formatted HTML email content for different error scenarios
.PARAMETER ErrorType
    The type of error (Database, Email, Configuration, etc.)
.PARAMETER ErrorMessage
    The detailed error message
.PARAMETER Timestamp
    When the error occurred
.EXAMPLE
    Get-ErrorEmailTemplate -ErrorType "Database" -ErrorMessage "Connection timeout" -Timestamp (Get-Date)
#>
function Get-ErrorEmailTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ErrorType,
        
        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,
        
        [Parameter(Mandatory = $true)]
        [DateTime]$Timestamp
    )
    
    $computerName = $env:COMPUTERNAME
    $formattedTimestamp = $Timestamp.ToString("yyyy-MM-dd HH:mm:ss")
    
    # Get error-specific details and recommendations
    $errorDetails = Get-ErrorTypeDetails -ErrorType $ErrorType
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #d32f2f; color: white; padding: 15px; border-radius: 5px; }
        .content { background-color: #f5f5f5; padding: 20px; border-radius: 5px; margin: 10px 0; }
        .error-details { background-color: #ffebee; border-left: 4px solid #d32f2f; padding: 15px; margin: 10px 0; }
        .recommendations { background-color: #e3f2fd; border-left: 4px solid #1976d2; padding: 15px; margin: 10px 0; }
        .footer { font-size: 12px; color: #666; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <div class="header">
        <h2>🚨 Change Report System Error Alert</h2>
    </div>
    
    <div class="content">
        <h3>Error Summary</h3>
        <table>
            <tr><th>Error Type</th><td>$ErrorType</td></tr>
            <tr><th>Timestamp</th><td>$formattedTimestamp</td></tr>
            <tr><th>Server</th><td>$computerName</td></tr>
            <tr><th>Severity</th><td>$($errorDetails.Severity)</td></tr>
        </table>
    </div>
    
    <div class="error-details">
        <h3>Error Details</h3>
        <p><strong>Message:</strong> $ErrorMessage</p>
        <p><strong>Description:</strong> $($errorDetails.Description)</p>
    </div>
    
    <div class="recommendations">
        <h3>Recommended Actions</h3>
        <ul>
"@
    
    foreach ($action in $errorDetails.Actions) {
        $html += "<li>$action</li>"
    }
    
    $html += @"
        </ul>
    </div>
    
    <div class="footer">
        <p>This is an automated error notification from the Change Report Email System.</p>
        <p>Generated on $computerName at $formattedTimestamp</p>
    </div>
</body>
</html>
"@
    
    return $html
}

<#
.SYNOPSIS
    Gets error-specific details and recommendations
.DESCRIPTION
    Returns structured information about different error types including severity and recommended actions
.PARAMETER ErrorType
    The type of error to get details for
.EXAMPLE
    Get-ErrorTypeDetails -ErrorType "Database"
#>
function Get-ErrorTypeDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ErrorType
    )
    
    switch ($ErrorType.ToLower()) {
        "database" {
            return @{
                Severity = "High"
                Description = "The system was unable to connect to or query the SQL Server database. This prevents the change report from being generated."
                Actions = @(
                    "Verify SQL Server is running and accessible",
                    "Check database connection string and credentials",
                    "Ensure network connectivity to database server",
                    "Verify database permissions for the service account",
                    "Check SQL Server logs for additional error details"
                )
            }
        }
        "email" {
            return @{
                Severity = "Medium"
                Description = "The system was unable to send email notifications. Change data may have been retrieved successfully but notifications were not delivered."
                Actions = @(
                    "Verify SMTP server settings and connectivity",
                    "Check email authentication credentials",
                    "Ensure firewall allows SMTP traffic",
                    "Verify recipient email addresses are valid",
                    "Check SMTP server logs for delivery issues"
                )
            }
        }
        "configuration" {
            return @{
                Severity = "High"
                Description = "The system configuration is invalid or missing required parameters. This prevents the system from operating correctly."
                Actions = @(
                    "Verify configuration file exists and is readable",
                    "Check all required configuration parameters are present",
                    "Validate configuration file JSON syntax",
                    "Ensure file permissions allow read access",
                    "Review configuration documentation for required fields"
                )
            }
        }
        "authentication" {
            return @{
                Severity = "High"
                Description = "Authentication failed for database or email services. This indicates credential or permission issues."
                Actions = @(
                    "Verify service account credentials are correct",
                    "Check if passwords have expired or been changed",
                    "Ensure service account has necessary permissions",
                    "Test credentials manually if possible",
                    "Contact system administrators for credential verification"
                )
            }
        }
        default {
            return @{
                Severity = "Medium"
                Description = "An unexpected error occurred in the Change Report system."
                Actions = @(
                    "Review system logs for additional error details",
                    "Check system resources (disk space, memory)",
                    "Verify PowerShell execution policy allows script execution",
                    "Contact system administrator for further investigation",
                    "Consider restarting the scheduled task"
                )
            }
        }
    }
}

# Functions are available when dot-sourced
# Note: Export-ModuleMember is not needed when dot-sourcing