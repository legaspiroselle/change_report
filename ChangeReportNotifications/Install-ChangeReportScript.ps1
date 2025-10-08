#Requires -Version 5.1
<#
.SYNOPSIS
    Installation script for Change Report Email Notifications system

.DESCRIPTION
    This script automates the setup of the Change Report Email Notifications system including:
    - Directory structure creation
    - PowerShell module dependency checking
    - Initial configuration file generation
    - Permission setup

.PARAMETER InstallPath
    The path where the Change Report system will be installed. Default: C:\ChangeReportNotifications

.PARAMETER ConfigOnly
    Only generate configuration file without full installation

.EXAMPLE
    .\Install-ChangeReportScript.ps1
    
.EXAMPLE
    .\Install-ChangeReportScript.ps1 -InstallPath "D:\Scripts\ChangeReports"
    
.EXAMPLE
    .\Install-ChangeReportScript.ps1 -ConfigOnly
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath = "C:\ChangeReportNotifications",
    
    [Parameter(Mandatory = $false)]
    [switch]$ConfigOnly
)

# Import required modules
Import-Module Microsoft.PowerShell.Security -Force

function Write-InstallLog {
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
    Write-InstallLog "Checking prerequisites..."
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-InstallLog "PowerShell 5.1 or higher is required. Current version: $($PSVersionTable.PSVersion)" -Level "Error"
        return $false
    }
    
    # Check if running as Administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-InstallLog "This script must be run as Administrator for proper installation" -Level "Warning"
    }
    
    Write-InstallLog "Prerequisites check completed" -Level "Success"
    return $true
}

function Install-RequiredModules {
    Write-InstallLog "Checking PowerShell module dependencies..."
    
    $requiredModules = @(
        @{ Name = "SqlServer"; MinVersion = "21.0.0"; Description = "SQL Server PowerShell module for database connectivity" }
    )
    
    foreach ($module in $requiredModules) {
        Write-InstallLog "Checking module: $($module.Name)"
        
        $installedModule = Get-Module -Name $module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $installedModule) {
            Write-InstallLog "Module $($module.Name) not found. Installing..." -Level "Warning"
            try {
                Install-Module -Name $module.Name -MinimumVersion $module.MinVersion -Force -AllowClobber -Scope CurrentUser
                Write-InstallLog "Successfully installed $($module.Name)" -Level "Success"
            }
            catch {
                Write-InstallLog "Failed to install $($module.Name): $($_.Exception.Message)" -Level "Error"
                Write-InstallLog "You can manually install using: Install-Module -Name $($module.Name)" -Level "Info"
            }
        }
        elseif ($installedModule.Version -lt [Version]$module.MinVersion) {
            Write-InstallLog "Module $($module.Name) version $($installedModule.Version) is below minimum required $($module.MinVersion). Updating..." -Level "Warning"
            try {
                Update-Module -Name $module.Name -Force
                Write-InstallLog "Successfully updated $($module.Name)" -Level "Success"
            }
            catch {
                Write-InstallLog "Failed to update $($module.Name): $($_.Exception.Message)" -Level "Error"
            }
        }
        else {
            Write-InstallLog "Module $($module.Name) version $($installedModule.Version) is already installed" -Level "Success"
        }
    }
}

function New-DirectoryStructure {
    param([string]$BasePath)
    
    Write-InstallLog "Creating directory structure at: $BasePath"
    
    $directories = @(
        $BasePath,
        "$BasePath\modules",
        "$BasePath\config",
        "$BasePath\logs",
        "$BasePath\TestArtifacts"
    )
    
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            try {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
                Write-InstallLog "Created directory: $dir" -Level "Success"
            }
            catch {
                Write-InstallLog "Failed to create directory $dir: $($_.Exception.Message)" -Level "Error"
                return $false
            }
        }
        else {
            Write-InstallLog "Directory already exists: $dir"
        }
    }
    
    return $true
}

function Set-DirectoryPermissions {
    param([string]$BasePath)
    
    Write-InstallLog "Setting directory permissions..."
    
    try {
        # Set permissions for logs directory (allow write access)
        $logsPath = "$BasePath\logs"
        if (Test-Path $logsPath) {
            $acl = Get-Acl $logsPath
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $env:USERNAME, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
            )
            $acl.SetAccessRule($accessRule)
            Set-Acl -Path $logsPath -AclObject $acl
            Write-InstallLog "Set permissions for logs directory" -Level "Success"
        }
        
        return $true
    }
    catch {
        Write-InstallLog "Failed to set permissions: $($_.Exception.Message)" -Level "Warning"
        return $false
    }
}

function Copy-ScriptFiles {
    param([string]$DestinationPath)
    
    Write-InstallLog "Copying script files..."
    
    $currentPath = Split-Path -Parent $MyInvocation.ScriptName
    $filesToCopy = @(
        "Send-ChangeReport.ps1",
        "Run-ChangeReport.bat",
        "modules\Config.ps1",
        "modules\Database.ps1",
        "modules\Email.ps1",
        "modules\Logging.ps1"
    )
    
    foreach ($file in $filesToCopy) {
        $sourcePath = Join-Path $currentPath $file
        $destPath = Join-Path $DestinationPath $file
        
        if (Test-Path $sourcePath) {
            try {
                $destDir = Split-Path -Parent $destPath
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                }
                
                Copy-Item -Path $sourcePath -Destination $destPath -Force
                Write-InstallLog "Copied: $file" -Level "Success"
            }
            catch {
                Write-InstallLog "Failed to copy $file: $($_.Exception.Message)" -Level "Error"
            }
        }
        else {
            Write-InstallLog "Source file not found: $sourcePath" -Level "Warning"
        }
    }
}

function New-ConfigurationFile {
    param([string]$ConfigPath)
    
    Write-InstallLog "Creating secure configuration file..."
    
    # Import secure configuration module
    . "$ConfigPath\modules\SecureConfig.ps1"
    
    # Prompt for configuration values
    Write-Host "`n=== Secure Configuration Setup ===" -ForegroundColor Cyan
    Write-Host "Please provide the following configuration details:" -ForegroundColor Yellow
    Write-Host "Note: Passwords will be encrypted using Windows DPAPI for security." -ForegroundColor Green
    
    # Database configuration
    Write-Host "`n--- Database Configuration ---" -ForegroundColor Green
    $dbServer = Read-Host "Database Server (e.g., SERVER\INSTANCE or SERVER,PORT)"
    $dbName = Read-Host "Database Name"
    
    $authTypes = @("Windows", "SQL")
    do {
        $dbAuthType = Read-Host "Authentication Type (Windows/SQL)"
    } while ($dbAuthType -notin $authTypes)
    
    $dbUsername = ""
    $dbEncryptedPassword = ""
    if ($dbAuthType -eq "SQL") {
        Write-Host "Database credentials will be encrypted for security." -ForegroundColor Yellow
        $dbCreds = Get-SecureCredentialInput -Title "Database Authentication" -Message "Enter database credentials:"
        $dbUsername = $dbCreds.Username
        $dbEncryptedPassword = $dbCreds.EncryptedPassword
    }
    
    # Email configuration
    Write-Host "`n--- Email Configuration ---" -ForegroundColor Green
    $smtpServer = Read-Host "SMTP Server (e.g., smtp.company.com)"
    $smtpPort = Read-Host "SMTP Port (default: 587)"
    if ([string]::IsNullOrEmpty($smtpPort)) { $smtpPort = 587 }
    
    $enableSSL = Read-Host "Enable SSL/TLS? (Y/N, default: Y)"
    $enableSSL = ($enableSSL -ne "N" -and $enableSSL -ne "n")
    
    $fromEmail = Read-Host "From Email Address"
    $toEmails = Read-Host "To Email Addresses (comma-separated)"
    $toEmailArray = $toEmails -split "," | ForEach-Object { $_.Trim() }
    
    $smtpUsername = Read-Host "SMTP Username (leave blank if not required)"
    $smtpEncryptedPassword = ""
    if (-not [string]::IsNullOrEmpty($smtpUsername)) {
        Write-Host "SMTP credentials will be encrypted for security." -ForegroundColor Yellow
        $smtpCreds = Get-SecureCredentialInput -Title "SMTP Authentication" -Message "Enter SMTP credentials:"
        $smtpUsername = $smtpCreds.Username
        $smtpEncryptedPassword = $smtpCreds.EncryptedPassword
    }
    
    # Execution time
    Write-Host "`n--- Scheduling Configuration ---" -ForegroundColor Green
    $executionTime = Read-Host "Daily execution time (HH:MM format, e.g., 08:00)"
    
    # Create secure configuration object
    $config = @{
        Database = @{
            Server = $dbServer
            Database = $dbName
            AuthType = $dbAuthType
            Username = $dbUsername
            EncryptedPassword = $dbEncryptedPassword
        }
        Email = @{
            SMTPServer = $smtpServer
            Port = [int]$smtpPort
            EnableSSL = $enableSSL
            From = $fromEmail
            To = $toEmailArray
            Username = $smtpUsername
            EncryptedPassword = $smtpEncryptedPassword
        }
        Logging = @{
            LogPath = "$ConfigPath\logs"
            LogLevel = "Info"
        }
        Schedule = @{
            ExecutionTime = $executionTime
        }
    }
    
    # Save secure configuration to JSON file
    $configFilePath = "$ConfigPath\config\config.json"
    try {
        $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $configFilePath -Encoding UTF8
        Write-InstallLog "Secure configuration file created: $configFilePath" -Level "Success"
        Write-InstallLog "Passwords have been encrypted using Windows DPAPI" -Level "Success"
        
        # Create a sample configuration file for reference
        $sampleConfigPath = "$ConfigPath\config\config.sample.json"
        $sampleConfig = @{
            Database = @{
                Server = "YOUR_SERVER\INSTANCE"
                Database = "YOUR_DATABASE"
                AuthType = "Windows"
                Username = ""
                EncryptedPassword = ""
            }
            Email = @{
                SMTPServer = "smtp.company.com"
                Port = 587
                EnableSSL = $true
                From = "noreply@company.com"
                To = @("admin@company.com", "manager@company.com")
                Username = ""
                EncryptedPassword = ""
            }
            Logging = @{
                LogPath = "$ConfigPath\logs"
                LogLevel = "Info"
            }
            Schedule = @{
                ExecutionTime = "08:00"
            }
        }
        
        $sampleConfig | ConvertTo-Json -Depth 3 | Out-File -FilePath $sampleConfigPath -Encoding UTF8
        Write-InstallLog "Sample configuration file created: $sampleConfigPath" -Level "Success"
        
        return $true
    }
    catch {
        Write-InstallLog "Failed to create configuration file: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Test-Installation {
    param([string]$InstallPath)
    
    Write-InstallLog "Testing installation..."
    
    $requiredFiles = @(
        "$InstallPath\Send-ChangeReport.ps1",
        "$InstallPath\config\config.json",
        "$InstallPath\modules\Config.ps1",
        "$InstallPath\modules\Database.ps1",
        "$InstallPath\modules\Email.ps1",
        "$InstallPath\modules\Logging.ps1"
    )
    
    $allFilesExist = $true
    foreach ($file in $requiredFiles) {
        if (-not (Test-Path $file)) {
            Write-InstallLog "Missing required file: $file" -Level "Error"
            $allFilesExist = $false
        }
    }
    
    if ($allFilesExist) {
        Write-InstallLog "Installation validation completed successfully" -Level "Success"
        return $true
    }
    else {
        Write-InstallLog "Installation validation failed" -Level "Error"
        return $false
    }
}

# Main installation process
try {
    Write-Host "=== Change Report Email Notifications - Installation Script ===" -ForegroundColor Cyan
    Write-Host "This script will install and configure the Change Report system.`n" -ForegroundColor Yellow
    
    # Check prerequisites
    if (-not (Test-Prerequisites)) {
        Write-InstallLog "Prerequisites check failed. Installation aborted." -Level "Error"
        exit 1
    }
    
    if (-not $ConfigOnly) {
        # Install required modules
        Install-RequiredModules
        
        # Create directory structure
        if (-not (New-DirectoryStructure -BasePath $InstallPath)) {
            Write-InstallLog "Failed to create directory structure. Installation aborted." -Level "Error"
            exit 1
        }
        
        # Set permissions
        Set-DirectoryPermissions -BasePath $InstallPath
        
        # Copy script files
        Copy-ScriptFiles -DestinationPath $InstallPath
    }
    
    # Create configuration file
    if (-not (New-ConfigurationFile -ConfigPath $InstallPath)) {
        Write-InstallLog "Failed to create configuration file. Installation aborted." -Level "Error"
        exit 1
    }
    
    if (-not $ConfigOnly) {
        # Test installation
        if (-not (Test-Installation -InstallPath $InstallPath)) {
            Write-InstallLog "Installation validation failed." -Level "Error"
            exit 1
        }
    }
    
    Write-Host "`n=== Installation Completed Successfully ===" -ForegroundColor Green
    Write-InstallLog "Change Report Email Notifications system has been installed to: $InstallPath" -Level "Success"
    
    if (-not $ConfigOnly) {
        Write-Host "`nNext Steps:" -ForegroundColor Yellow
        Write-Host "1. Review and test the configuration file: $InstallPath\config\config.json" -ForegroundColor White
        Write-Host "2. Run the setup script for Task Scheduler: .\Setup-TaskScheduler.ps1" -ForegroundColor White
        Write-Host "3. Test the system manually: .$InstallPath\Send-ChangeReport.ps1 -TestMode" -ForegroundColor White
        Write-Host "4. Check the logs directory for any issues: $InstallPath\logs" -ForegroundColor White
    }
    else {
        Write-Host "`nConfiguration file has been created. Please review and modify as needed." -ForegroundColor Yellow
    }
}
catch {
    Write-InstallLog "Installation failed with error: $($_.Exception.Message)" -Level "Error"
    Write-InstallLog "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    exit 1
}