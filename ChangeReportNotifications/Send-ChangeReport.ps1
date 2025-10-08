#Requires -Version 5.1

<#
.SYNOPSIS
    Main orchestrator script for Change Report Email Notifications
.DESCRIPTION
    Coordinates all modules to query database for critical/high priority changes
    and send formatted email notifications with comprehensive error handling
.PARAMETER ConfigPath
    Path to the JSON configuration file (default: .\config\config.json)
.PARAMETER TestMode
    Run in test mode without sending emails (logs actions only)
.PARAMETER ReportDate
    Specific date to generate report for (default: current date)
.PARAMETER Verbose
    Enable verbose logging output
.EXAMPLE
    .\Send-ChangeReport.ps1
.EXAMPLE
    .\Send-ChangeReport.ps1 -ConfigPath "C:\Config\myconfig.json" -TestMode
.EXAMPLE
    .\Send-ChangeReport.ps1 -ReportDate "2024-01-15" -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Path to the JSON configuration file")]
    [string]$ConfigPath = ".\config\config.json",
    
    [Parameter(Mandatory = $false, HelpMessage = "Run in test mode without sending emails")]
    [switch]$TestMode,
    
    [Parameter(Mandatory = $false, HelpMessage = "Specific date to generate report for (YYYY-MM-DD format)")]
    [datetime]$ReportDate = (Get-Date),
    
    [Parameter(Mandatory = $false, HelpMessage = "Force execution even if recent execution detected")]
    [switch]$Force,
    
    [Parameter(Mandatory = $false, HelpMessage = "Display help information")]
    [switch]$Help
)

# Set error action preference for consistent error handling
$ErrorActionPreference = "Stop"

# Initialize variables
$script:Config = $null
$script:DatabaseConnection = $null
$script:SessionId = $null
$script:ExitCode = 0

# Import required modules
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "modules"

try {
    . (Join-Path -Path $ModulePath -ChildPath "SecureConfig.ps1")
    . (Join-Path -Path $ModulePath -ChildPath "Config.ps1")
    . (Join-Path -Path $ModulePath -ChildPath "Logging.ps1")
    . (Join-Path -Path $ModulePath -ChildPath "Database.ps1")
    . (Join-Path -Path $ModulePath -ChildPath "Email.ps1")
}
catch {
    Write-Error "Failed to import required modules: $($_.Exception.Message)"
    exit 1
}

<#
.SYNOPSIS
    Main execution function that orchestrates the entire process
#>
function Invoke-ChangeReportProcess {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host "Starting Change Report Email Notification Process..." -ForegroundColor Cyan
        
        # Step 1: Load and validate configuration
        Write-Host "Step 1: Loading configuration..." -ForegroundColor Yellow
        $script:Config = Initialize-Configuration -ConfigPath $ConfigPath
        
        # Step 2: Initialize logging
        Write-Host "Step 2: Initializing logging..." -ForegroundColor Yellow
        $script:SessionId = Initialize-Logging
        
        # Step 3: Test database connectivity
        Write-Host "Step 3: Testing database connectivity..." -ForegroundColor Yellow
        Test-DatabaseConnectivity
        
        # Step 4: Connect to database and retrieve changes
        Write-Host "Step 4: Retrieving change records..." -ForegroundColor Yellow
        $changes = Get-ChangeRecords
        
        # Step 5: Format and send email notification (always send, even if no changes)
        Write-Host "Step 5: Generating and sending email notification..." -ForegroundColor Yellow
        Send-EmailNotification -Changes $changes
        
        # Step 6: Cleanup and finalize
        Write-Host "Step 6: Finalizing process..." -ForegroundColor Yellow
        Complete-Process -Success $true
        
        Write-Host "Change Report process completed successfully!" -ForegroundColor Green
        
    }
    catch {
        $errorMessage = $_.Exception.Message
        $errorType = Get-ErrorType -Exception $_.Exception
        
        Write-Host "CRITICAL ERROR: $errorMessage" -ForegroundColor Red
        
        # Log the error
        if ($script:Config -and $script:Config.Logging) {
            Write-Log -Message "CRITICAL ERROR [$errorType]: $errorMessage" -Level Error -LogPath $script:Config.Logging.LogPath
            Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level Error -LogPath $script:Config.Logging.LogPath
        }
        
        # Attempt to send error notification
        try {
            if ($script:Config -and $script:Config.Email) {
                Send-ErrorNotification -Config $script:Config -ErrorMessage $errorMessage -ErrorType $errorType -LogPath $script:Config.Logging.LogPath
            }
        }
        catch {
            Write-Warning "Failed to send error notification: $($_.Exception.Message)"
        }
        
        # Cleanup and set error exit code
        Complete-Process -Success $false
        $script:ExitCode = 1
        
        # Re-throw for proper error handling
        throw
    }
}

<#
.SYNOPSIS
    Initializes and validates configuration
#>
function Initialize-Configuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        # Resolve full path
        $fullConfigPath = Resolve-Path -Path $ConfigPath -ErrorAction Stop
        Write-Verbose "Loading configuration from: $fullConfigPath"
        
        # Load secure configuration
        $config = Get-SecureConfiguration -ConfigPath $fullConfigPath
        
        # Validate configuration
        $isValid = Test-Configuration -Config $config
        if (-not $isValid) {
            throw "Configuration validation failed. Please check the configuration file and try again."
        }
        
        Write-Verbose "Configuration loaded and validated successfully"
        return $config
        
    }
    catch {
        throw "Configuration initialization failed: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Initializes logging session
#>
function Initialize-Logging {
    [CmdletBinding()]
    param()
    
    try {
        # Ensure log directory exists
        if (-not (Test-Path -Path $script:Config.Logging.LogPath)) {
            New-Item -ItemType Directory -Path $script:Config.Logging.LogPath -Force | Out-Null
        }
        
        # Start logging session
        $sessionId = Start-LogSession -LogPath $script:Config.Logging.LogPath
        
        # Log execution parameters
        Write-Log -Message "Execution Parameters:" -Level Info -LogPath $script:Config.Logging.LogPath
        Write-Log -Message "  Config Path: $ConfigPath" -Level Info -LogPath $script:Config.Logging.LogPath
        Write-Log -Message "  Report Date: $($ReportDate.ToString('yyyy-MM-dd'))" -Level Info -LogPath $script:Config.Logging.LogPath
        Write-Log -Message "  Test Mode: $TestMode" -Level Info -LogPath $script:Config.Logging.LogPath
        Write-Log -Message "  Verbose: $($VerbosePreference -ne 'SilentlyContinue')" -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Clean up old log files (keep 30 days by default)
        Remove-OldLogFiles -LogPath $script:Config.Logging.LogPath -DaysToKeep 30
        
        return $sessionId
        
    }
    catch {
        throw "Logging initialization failed: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Tests database connectivity before proceeding
#>
function Test-DatabaseConnectivity {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Message "Testing database connectivity..." -Level Info -LogPath $script:Config.Logging.LogPath
        
        $isConnected = Test-DatabaseConnection -Config $script:Config
        
        if (-not $isConnected) {
            throw "Database connectivity test failed. Please check database configuration and network connectivity."
        }
        
        Write-Log -Message "Database connectivity test successful" -Level Info -LogPath $script:Config.Logging.LogPath
        
    }
    catch {
        throw "Database connectivity test failed: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Retrieves change records from database
#>
function Get-ChangeRecords {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log -Message "Connecting to database..." -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Connect to database
        $script:DatabaseConnection = Connect-Database -Config $script:Config
        
        if (-not $script:DatabaseConnection) {
            throw "Failed to establish database connection"
        }
        
        Write-Log -Message "Database connection established successfully" -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Query for critical and high priority changes
        Write-Log -Message "Querying for critical and high priority changes for date: $($ReportDate.ToString('yyyy-MM-dd'))" -Level Info -LogPath $script:Config.Logging.LogPath
        
        $changes = Get-CriticalChanges -Connection $script:DatabaseConnection -ReportDate $ReportDate
        
        Write-Log -Message "Retrieved $($changes.Count) change records" -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Log summary by priority
        $criticalCount = ($changes | Where-Object { $_.Priority -eq "Critical" }).Count
        $highCount = ($changes | Where-Object { $_.Priority -eq "High" }).Count
        
        Write-Log -Message "Change summary: $criticalCount Critical, $highCount High priority changes" -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Log individual changes for audit trail
        if ($changes.Count -gt 0) {
            Write-Log -Message "Change details:" -Level Info -LogPath $script:Config.Logging.LogPath
            foreach ($change in $changes) {
                Write-Log -Message "  - $($change.ChangeID): $($change.Priority) - $($change.ShortDescription)" -Level Info -LogPath $script:Config.Logging.LogPath
            }
        }
        else {
            Write-Log -Message "No critical or high priority changes found for the specified date" -Level Info -LogPath $script:Config.Logging.LogPath
        }
        
        return $changes
        
    }
    catch {
        throw "Failed to retrieve change records: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Formats and sends email notification
#>
function Send-EmailNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Changes
    )
    
    try {
        Write-Log -Message "Formatting email notification..." -Level Info -LogPath $script:Config.Logging.LogPath
        
        # Handle no-changes scenario with specific template
        if ($Changes.Count -eq 0) {
            Write-Log -Message "No changes found - using no-changes email template" -Level Info -LogPath $script:Config.Logging.LogPath
            $emailBody = Get-NoChangesEmailTemplate -Date $ReportDate
        }
        else {
            # Format email content for changes
            $emailBody = Format-ChangeReport -Changes $Changes -Date $ReportDate
        }
        
        $emailSubject = Get-EmailSubject -Date $ReportDate -ChangeCount $Changes.Count
        
        Write-Log -Message "Email subject: $emailSubject" -Level Info -LogPath $script:Config.Logging.LogPath
        Write-Log -Message "Email recipients: $($script:Config.Email.To -join ', ')" -Level Info -LogPath $script:Config.Logging.LogPath
        
        if ($TestMode) {
            Write-Log -Message "TEST MODE: Email would be sent but TestMode is enabled" -Level Warning -LogPath $script:Config.Logging.LogPath
            Write-Host "TEST MODE: Email content generated successfully but not sent" -ForegroundColor Yellow
            
            # Optionally save email content to file for review
            $testEmailPath = Join-Path -Path $script:Config.Logging.LogPath -ChildPath "test_email_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
            $emailBody | Out-File -FilePath $testEmailPath -Encoding UTF8
            Write-Log -Message "Test email content saved to: $testEmailPath" -Level Info -LogPath $script:Config.Logging.LogPath
            
            # Show what type of email would be sent
            if ($Changes.Count -eq 0) {
                Write-Host "TEST MODE: Would send NO-CHANGES notification email" -ForegroundColor Cyan
            }
            else {
                Write-Host "TEST MODE: Would send change report with $($Changes.Count) changes" -ForegroundColor Cyan
            }
        }
        else {
            # Always send email notification (requirement 5.4 - daily email regardless of changes)
            Write-Log -Message "Sending email notification..." -Level Info -LogPath $script:Config.Logging.LogPath
            
            if ($Changes.Count -eq 0) {
                Write-Log -Message "Sending no-changes notification email" -Level Info -LogPath $script:Config.Logging.LogPath
            }
            else {
                Write-Log -Message "Sending change report email with $($Changes.Count) changes" -Level Info -LogPath $script:Config.Logging.LogPath
            }
            
            $emailSent = Send-ChangeNotification -Config $script:Config -Body $emailBody -Subject $emailSubject
            
            if ($emailSent) {
                Write-Log -Message "Email notification sent successfully" -Level Info -LogPath $script:Config.Logging.LogPath
                
                if ($Changes.Count -eq 0) {
                    Write-Host "No-changes notification email sent successfully" -ForegroundColor Green
                }
                else {
                    Write-Host "Change report email sent successfully ($($Changes.Count) changes)" -ForegroundColor Green
                }
            }
            else {
                throw "Email delivery failed"
            }
        }
        
    }
    catch {
        throw "Failed to send email notification: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Completes the process with cleanup
#>
function Complete-Process {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Success
    )
    
    try {
        # Close database connection if open
        if ($script:DatabaseConnection) {
            Write-Log -Message "Closing database connection..." -Level Info -LogPath $script:Config.Logging.LogPath
            Close-Database -Connection $script:DatabaseConnection
            $script:DatabaseConnection = $null
        }
        
        # Log completion status
        if ($Success) {
            Write-Log -Message "Change Report process completed successfully" -Level Info -LogPath $script:Config.Logging.LogPath
            $script:ExitCode = 0
        }
        else {
            Write-Log -Message "Change Report process completed with errors" -Level Error -LogPath $script:Config.Logging.LogPath
            $script:ExitCode = 1
        }
        
        # End logging session
        if ($script:SessionId -and $script:Config) {
            Stop-LogSession -LogPath $script:Config.Logging.LogPath -SessionId $script:SessionId
        }
        
    }
    catch {
        Write-Warning "Error during cleanup: $($_.Exception.Message)"
        $script:ExitCode = 1
    }
}

<#
.SYNOPSIS
    Determines error type from exception for categorization
#>
function Get-ErrorType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )
    
    $errorMessage = $Exception.Message.ToLower()
    
    if ($errorMessage -match "database|sql|connection|timeout") {
        $script:ExitCode = 3
        return "Database"
    }
    elseif ($errorMessage -match "email|smtp|mail|send") {
        $script:ExitCode = 4
        return "Email"
    }
    elseif ($errorMessage -match "config|json|parameter|validation") {
        $script:ExitCode = 2
        return "Configuration"
    }
    elseif ($errorMessage -match "auth|credential|login|permission") {
        $script:ExitCode = 3
        return "Authentication"
    }
    else {
        $script:ExitCode = 1
        return "General"
    }
}

<#
.SYNOPSIS
    Displays help information for the script
#>
function Show-Help {
    Write-Host @"

Change Report Email Notification System
======================================

DESCRIPTION:
    Queries SQL Server database for critical and high priority changes and sends
    formatted email notifications to configured recipients.

USAGE:
    .\Send-ChangeReport.ps1 [parameters]

PARAMETERS:
    -ConfigPath <string>     Path to JSON configuration file
                            Default: .\config\config.json
    
    -TestMode               Run in test mode (no emails sent, content saved to file)
    
    -ReportDate <datetime>  Specific date for report (YYYY-MM-DD format)
                            Default: Current date
    
    -Force                  Force execution even if recent execution detected
    
    -Verbose               Enable detailed logging output
    
    -Help                  Display this help information

EXAMPLES:
    .\Send-ChangeReport.ps1
        Run with default settings
    
    .\Send-ChangeReport.ps1 -TestMode -Verbose
        Run in test mode with verbose output
    
    .\Send-ChangeReport.ps1 -ConfigPath "C:\Config\prod.json" -ReportDate "2024-01-15"
        Run with custom config and specific date
    
    .\Send-ChangeReport.ps1 -Force
        Force execution ignoring recent execution check

EXIT CODES:
    0 = Success
    1 = General error
    2 = Configuration error
    3 = Database/Authentication error
    4 = Email error
    5 = Recent execution detected (use -Force to override)

REQUIREMENTS:
    - PowerShell 5.1 or higher
    - Valid configuration file
    - Network access to SQL Server and SMTP server
    - Appropriate permissions for database and email operations

"@ -ForegroundColor Cyan
}

<#
.SYNOPSIS
    Checks for recent execution to prevent duplicate runs
#>
function Test-RecentExecution {
    [CmdletBinding()]
    param()
    
    try {
        # Create a lock file to prevent concurrent executions
        $lockFile = Join-Path -Path $env:TEMP -ChildPath "ChangeReport_$(Get-Date -Format 'yyyy-MM-dd').lock"
        
        if (Test-Path -Path $lockFile) {
            $lockContent = Get-Content -Path $lockFile -ErrorAction SilentlyContinue
            if ($lockContent) {
                $lastRun = [datetime]::ParseExact($lockContent, "yyyy-MM-dd HH:mm:ss", $null)
                $timeSinceLastRun = (Get-Date) - $lastRun
                
                # Prevent multiple executions within 1 hour
                if ($timeSinceLastRun.TotalHours -lt 1) {
                    Write-Warning "Recent execution detected at $($lastRun.ToString('yyyy-MM-dd HH:mm:ss'))"
                    Write-Warning "Use -Force parameter to override this check"
                    $script:ExitCode = 5
                    throw "Recent execution detected. Use -Force to override."
                }
            }
        }
        
        # Create/update lock file
        (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") | Out-File -FilePath $lockFile -Encoding UTF8
        
    }
    catch {
        if ($script:ExitCode -eq 5) {
            throw
        }
        # If lock file operations fail, continue anyway
        Write-Verbose "Could not check for recent execution: $($_.Exception.Message)"
    }
}

# Display help if requested
if ($Help) {
    Show-Help
    exit 0
}

# Main execution block
try {
    # Validate PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5.1 or higher is required. Current version: $($PSVersionTable.PSVersion)"
    }
    
    # Check for recent execution if not forced
    if (-not $Force) {
        Test-RecentExecution
    }
    
    # Execute main process
    Invoke-ChangeReportProcess
    
    Write-Host "Process completed with exit code: $script:ExitCode" -ForegroundColor $(if ($script:ExitCode -eq 0) { "Green" } else { "Red" })
}
catch {
    Write-Host "FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    $script:ExitCode = 1
}
finally {
    # Ensure cleanup happens even if main process fails
    if ($script:DatabaseConnection) {
        try {
            Close-Database -Connection $script:DatabaseConnection
        }
        catch {
            Write-Warning "Failed to close database connection during cleanup: $($_.Exception.Message)"
        }
    }
}

# Exit with appropriate code for monitoring systems
# Exit codes:
# 0 = Success
# 1 = General error
# 2 = Configuration error
# 3 = Database error
# 4 = Email error
# 5 = Recent execution detected (when not forced)
exit $script:ExitCode