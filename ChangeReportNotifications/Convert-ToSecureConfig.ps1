#Requires -Version 5.1
<#
.SYNOPSIS
    Converts existing plain text configuration to secure encrypted format

.DESCRIPTION
    This script migrates existing configuration files that contain plain text passwords
    to the new secure format using Windows DPAPI encryption. Creates a backup of the
    original configuration file before conversion.

.PARAMETER ConfigPath
    Path to the existing configuration file to convert

.PARAMETER BackupPath
    Path where to save the backup of the original file (optional)

.PARAMETER Force
    Force conversion even if the configuration appears to already be encrypted

.EXAMPLE
    .\Convert-ToSecureConfig.ps1 -ConfigPath ".\config\config.json"
    
.EXAMPLE
    .\Convert-ToSecureConfig.ps1 -ConfigPath ".\config\config.json" -BackupPath ".\config\config.backup.json"
    
.EXAMPLE
    .\Convert-ToSecureConfig.ps1 -ConfigPath ".\config\config.json" -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath,
    
    [Parameter(Mandatory = $false)]
    [string]$BackupPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Import secure configuration module
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "modules"
. (Join-Path -Path $ModulePath -ChildPath "SecureConfig.ps1")

function Write-ConversionLog {
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

try {
    Write-Host "=== Configuration Security Migration Tool ===" -ForegroundColor Cyan
    Write-Host "This tool converts plain text passwords to encrypted format using Windows DPAPI`n" -ForegroundColor Yellow
    
    # Validate input file
    if (-not (Test-Path $ConfigPath)) {
        throw "Configuration file not found: $ConfigPath"
    }
    
    $fullConfigPath = Resolve-Path $ConfigPath
    Write-ConversionLog "Processing configuration file: $fullConfigPath"
    
    # Set default backup path if not provided
    if (-not $BackupPath) {
        $configDir = Split-Path -Parent $fullConfigPath
        $configName = [System.IO.Path]::GetFileNameWithoutExtension($fullConfigPath)
        $configExt = [System.IO.Path]::GetExtension($fullConfigPath)
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $BackupPath = Join-Path $configDir "$configName.backup.$timestamp$configExt"
    }
    
    # Load and analyze current configuration
    Write-ConversionLog "Analyzing current configuration format..."
    
    $configContent = Get-Content -Path $fullConfigPath -Raw
    $config = $configContent | ConvertFrom-Json
    
    $hasPlainTextPasswords = $false
    $conversionNeeded = $false
    
    # Check database password
    if ($config.Database.AuthType -eq "SQL") {
        if ($config.Database.Password -and -not (Test-EncryptedString -String $config.Database.Password)) {
            Write-ConversionLog "Found plain text database password" -Level "Warning"
            $hasPlainTextPasswords = $true
            $conversionNeeded = $true
        }
        elseif ($config.Database.EncryptedPassword) {
            Write-ConversionLog "Database password already encrypted"
        }
    }
    
    # Check email password
    if ($config.Email.Password -and -not (Test-EncryptedString -String $config.Email.Password)) {
        Write-ConversionLog "Found plain text email password" -Level "Warning"
        $hasPlainTextPasswords = $true
        $conversionNeeded = $true
    }
    elseif ($config.Email.EncryptedPassword) {
        Write-ConversionLog "Email password already encrypted"
    }
    
    # Check if conversion is needed
    if (-not $conversionNeeded -and -not $Force) {
        Write-ConversionLog "Configuration appears to already be in secure format" -Level "Success"
        Write-ConversionLog "Use -Force parameter to convert anyway"
        exit 0
    }
    
    if ($hasPlainTextPasswords) {
        Write-Host "`nSECURITY WARNING:" -ForegroundColor Red
        Write-Host "Plain text passwords detected in configuration file!" -ForegroundColor Red
        Write-Host "This is a security risk and should be converted immediately.`n" -ForegroundColor Red
    }
    
    # Confirm conversion
    if (-not $Force) {
        $response = Read-Host "Do you want to proceed with the conversion? (Y/N)"
        if ($response -ne "Y" -and $response -ne "y") {
            Write-ConversionLog "Conversion cancelled by user"
            exit 0
        }
    }
    
    # Create backup
    Write-ConversionLog "Creating backup: $BackupPath"
    Copy-Item -Path $fullConfigPath -Destination $BackupPath -Force
    Write-ConversionLog "Backup created successfully" -Level "Success"
    
    # Perform conversion
    Write-ConversionLog "Converting configuration to secure format..."
    
    $updated = $false
    
    # Convert database password
    if ($config.Database.AuthType -eq "SQL" -and $config.Database.Password) {
        if (-not (Test-EncryptedString -String $config.Database.Password) -or $Force) {
            Write-ConversionLog "Encrypting database password..."
            
            $encryptedDbPassword = Protect-ConfigString -PlainText $config.Database.Password
            
            # Update configuration object
            $config.Database | Add-Member -NotePropertyName "EncryptedPassword" -NotePropertyValue $encryptedDbPassword -Force
            
            # Remove plain text password
            if ($config.Database.PSObject.Properties.Name -contains "Password") {
                $config.Database.PSObject.Properties.Remove("Password")
            }
            
            $updated = $true
            Write-ConversionLog "Database password encrypted successfully" -Level "Success"
        }
    }
    
    # Convert email password
    if ($config.Email.Password) {
        if (-not (Test-EncryptedString -String $config.Email.Password) -or $Force) {
            Write-ConversionLog "Encrypting email password..."
            
            $encryptedEmailPassword = Protect-ConfigString -PlainText $config.Email.Password
            
            # Update configuration object
            $config.Email | Add-Member -NotePropertyName "EncryptedPassword" -NotePropertyValue $encryptedEmailPassword -Force
            
            # Remove plain text password
            if ($config.Email.PSObject.Properties.Name -contains "Password") {
                $config.Email.PSObject.Properties.Remove("Password")
            }
            
            $updated = $true
            Write-ConversionLog "Email password encrypted successfully" -Level "Success"
        }
    }
    
    # Save updated configuration
    if ($updated -or $Force) {
        Write-ConversionLog "Saving updated configuration..."
        
        $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $fullConfigPath -Encoding UTF8
        
        Write-ConversionLog "Configuration saved successfully" -Level "Success"
    }
    
    # Verify the conversion
    Write-ConversionLog "Verifying converted configuration..."
    
    try {
        $verifyConfig = Get-SecureConfiguration -ConfigPath $fullConfigPath
        Write-ConversionLog "Configuration verification successful" -Level "Success"
    }
    catch {
        Write-ConversionLog "Configuration verification failed: $($_.Exception.Message)" -Level "Error"
        
        # Restore backup
        Write-ConversionLog "Restoring backup configuration..."
        Copy-Item -Path $BackupPath -Destination $fullConfigPath -Force
        throw "Conversion failed and backup has been restored"
    }
    
    Write-Host "`n=== Conversion Completed Successfully ===" -ForegroundColor Green
    Write-ConversionLog "Configuration has been converted to secure format" -Level "Success"
    Write-ConversionLog "Passwords are now encrypted using Windows DPAPI" -Level "Success"
    Write-ConversionLog "Original configuration backed up to: $BackupPath" -Level "Success"
    
    Write-Host "`nSecurity Benefits:" -ForegroundColor Yellow
    Write-Host "✓ Passwords are encrypted and can only be decrypted by the current user" -ForegroundColor Green
    Write-Host "✓ Configuration file no longer contains plain text credentials" -ForegroundColor Green
    Write-Host "✓ Encrypted passwords are tied to the current Windows user account" -ForegroundColor Green
    Write-Host "✓ Backup of original configuration is available for rollback if needed" -ForegroundColor Green
    
    Write-Host "`nNext Steps:" -ForegroundColor Yellow
    Write-Host "1. Test the configuration with: .\Send-ChangeReport.ps1 -TestMode" -ForegroundColor White
    Write-Host "2. Verify that the system can decrypt passwords and connect successfully" -ForegroundColor White
    Write-Host "3. Once verified, securely delete the backup file: $BackupPath" -ForegroundColor White
    Write-Host "4. Update any documentation to reflect the new secure configuration format" -ForegroundColor White
    
}
catch {
    Write-ConversionLog "Conversion failed: $($_.Exception.Message)" -Level "Error"
    Write-ConversionLog "Stack trace: $($_.ScriptStackTrace)" -Level "Error"
    
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have read/write access to the configuration file" -ForegroundColor White
    Write-Host "2. Verify the configuration file contains valid JSON" -ForegroundColor White
    Write-Host "3. Check that you're running as the same user who will execute the main script" -ForegroundColor White
    Write-Host "4. Ensure Windows DPAPI is available (Windows Vista/Server 2008 or later)" -ForegroundColor White
    
    exit 1
}