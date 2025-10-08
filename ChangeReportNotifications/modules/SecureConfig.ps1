# Secure Configuration Module for Change Report Email Notifications
# Handles secure credential storage and retrieval using Windows DPAPI

<#
.SYNOPSIS
    Secure configuration management module with encrypted credential storage
.DESCRIPTION
    This module provides secure storage and retrieval of sensitive configuration data
    using Windows Data Protection API (DPAPI) for encryption
#>

Add-Type -AssemblyName System.Security

<#
.SYNOPSIS
    Encrypts a plain text string using Windows DPAPI
.DESCRIPTION
    Uses Windows Data Protection API to encrypt sensitive data for the current user
.PARAMETER PlainText
    The plain text string to encrypt
.RETURNS
    Base64 encoded encrypted string
.EXAMPLE
    $encrypted = Protect-ConfigString -PlainText "mypassword"
#>
function Protect-ConfigString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlainText
    )
    
    try {
        if ([string]::IsNullOrEmpty($PlainText)) {
            return ""
        }
        
        # Convert string to bytes
        $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        
        # Encrypt using DPAPI for current user
        $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
            $plainBytes,
            $null,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        
        # Convert to Base64 for storage
        return [System.Convert]::ToBase64String($encryptedBytes)
    }
    catch {
        Write-Error "Failed to encrypt string: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Decrypts an encrypted string using Windows DPAPI
.DESCRIPTION
    Uses Windows Data Protection API to decrypt sensitive data for the current user
.PARAMETER EncryptedString
    The Base64 encoded encrypted string to decrypt
.RETURNS
    Plain text string
.EXAMPLE
    $plainText = Unprotect-ConfigString -EncryptedString $encrypted
#>
function Unprotect-ConfigString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EncryptedString
    )
    
    try {
        if ([string]::IsNullOrEmpty($EncryptedString)) {
            return ""
        }
        
        # Convert from Base64
        $encryptedBytes = [System.Convert]::FromBase64String($EncryptedString)
        
        # Decrypt using DPAPI
        $plainBytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $encryptedBytes,
            $null,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        
        # Convert back to string
        return [System.Text.Encoding]::UTF8.GetString($plainBytes)
    }
    catch {
        Write-Error "Failed to decrypt string: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Creates a secure credential object from username and password
.DESCRIPTION
    Converts username and encrypted password to PSCredential object
.PARAMETER Username
    The username
.PARAMETER EncryptedPassword
    The encrypted password string
.RETURNS
    PSCredential object
.EXAMPLE
    $cred = New-SecureCredential -Username "user" -EncryptedPassword $encryptedPass
#>
function New-SecureCredential {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
        
        [Parameter(Mandatory = $true)]
        [string]$EncryptedPassword
    )
    
    try {
        if ([string]::IsNullOrEmpty($Username) -or [string]::IsNullOrEmpty($EncryptedPassword)) {
            return $null
        }
        
        # Decrypt password
        $plainPassword = Unprotect-ConfigString -EncryptedString $EncryptedPassword
        
        # Create secure string
        $securePassword = ConvertTo-SecureString -String $plainPassword -AsPlainText -Force
        
        # Clear plain text password from memory
        $plainPassword = $null
        [System.GC]::Collect()
        
        # Create credential object
        return New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    }
    catch {
        Write-Error "Failed to create secure credential: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Loads configuration with secure credential handling
.DESCRIPTION
    Reads configuration file and decrypts sensitive data as needed
.PARAMETER ConfigPath
    Path to the JSON configuration file
.RETURNS
    Configuration hashtable with decrypted credentials
.EXAMPLE
    $config = Get-SecureConfiguration -ConfigPath ".\config\config.json"
#>
function Get-SecureConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        Write-Verbose "Loading secure configuration from: $ConfigPath"
        
        $configContent = Get-Content -Path $ConfigPath -Raw
        $config = $configContent | ConvertFrom-Json
        
        # Convert to hashtable
        $configHash = @{}
        
        # Database configuration with secure credentials
        $configHash.Database = @{
            Server = $config.Database.Server
            Database = $config.Database.Database
            AuthType = $config.Database.AuthType
            Username = $config.Database.Username
            EncryptedPassword = $config.Database.EncryptedPassword
        }
        
        # Add decrypted credential if SQL authentication
        if ($configHash.Database.AuthType -eq "SQL" -and 
            -not [string]::IsNullOrEmpty($configHash.Database.Username) -and
            -not [string]::IsNullOrEmpty($configHash.Database.EncryptedPassword)) {
            
            $configHash.Database.Credential = New-SecureCredential -Username $configHash.Database.Username -EncryptedPassword $configHash.Database.EncryptedPassword
        }
        
        # Email configuration with secure credentials
        $configHash.Email = @{
            SMTPServer = $config.Email.SMTPServer
            Port = $config.Email.Port
            EnableSSL = $config.Email.EnableSSL
            From = $config.Email.From
            To = $config.Email.To
            Username = $config.Email.Username
            EncryptedPassword = $config.Email.EncryptedPassword
        }
        
        # Add decrypted credential if SMTP authentication is configured
        if (-not [string]::IsNullOrEmpty($configHash.Email.Username)) {
            if (-not [string]::IsNullOrEmpty($configHash.Email.EncryptedPassword)) {
                # Full credential with username and password
                $configHash.Email.Credential = New-SecureCredential -Username $configHash.Email.Username -EncryptedPassword $configHash.Email.EncryptedPassword
            }
            else {
                # Username only (some SMTP servers accept username without password)
                Write-Verbose "Email username provided without password - will attempt username-only authentication"
            }
        }
        else {
            # No credentials - will use anonymous authentication
            Write-Verbose "No email credentials configured - will use anonymous SMTP authentication"
        }
        
        # Logging configuration (no sensitive data)
        $configHash.Logging = @{
            LogPath = $config.Logging.LogPath
            LogLevel = $config.Logging.LogLevel
        }
        
        # Schedule configuration (no sensitive data)
        if ($config.Schedule) {
            $configHash.Schedule = @{
                ExecutionTime = $config.Schedule.ExecutionTime
            }
        }
        
        Write-Verbose "Configuration loaded successfully with secure credential handling"
        return $configHash
    }
    catch {
        Write-Error "Failed to load secure configuration: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Saves configuration with encrypted passwords
.DESCRIPTION
    Encrypts sensitive data before saving to configuration file
.PARAMETER Config
    Configuration hashtable to save
.PARAMETER ConfigPath
    Path where to save the configuration file
.EXAMPLE
    Set-SecureConfiguration -Config $config -ConfigPath ".\config\config.json"
#>
function Set-SecureConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        Write-Verbose "Saving secure configuration to: $ConfigPath"
        
        # Create configuration object for JSON serialization
        $configObject = @{
            Database = @{
                Server = $Config.Database.Server
                Database = $Config.Database.Database
                AuthType = $Config.Database.AuthType
                Username = $Config.Database.Username
                EncryptedPassword = ""
            }
            Email = @{
                SMTPServer = $Config.Email.SMTPServer
                Port = $Config.Email.Port
                EnableSSL = $Config.Email.EnableSSL
                From = $Config.Email.From
                To = $Config.Email.To
                Username = $Config.Email.Username
                EncryptedPassword = ""
            }
            Logging = @{
                LogPath = $Config.Logging.LogPath
                LogLevel = $Config.Logging.LogLevel
            }
        }
        
        # Add schedule if present
        if ($Config.Schedule) {
            $configObject.Schedule = @{
                ExecutionTime = $Config.Schedule.ExecutionTime
            }
        }
        
        # Encrypt database password if provided
        if ($Config.Database.AuthType -eq "SQL" -and -not [string]::IsNullOrEmpty($Config.Database.Password)) {
            $configObject.Database.EncryptedPassword = Protect-ConfigString -PlainText $Config.Database.Password
        }
        
        # Encrypt email password if provided
        if (-not [string]::IsNullOrEmpty($Config.Email.Password)) {
            $configObject.Email.EncryptedPassword = Protect-ConfigString -PlainText $Config.Email.Password
        }
        
        # Ensure directory exists
        $configDir = Split-Path -Parent $ConfigPath
        if (-not (Test-Path $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        }
        
        # Save to JSON file
        $configObject | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8
        
        Write-Verbose "Secure configuration saved successfully"
        
        # Clear sensitive data from memory
        if ($Config.Database.Password) {
            $Config.Database.Password = $null
        }
        if ($Config.Email.Password) {
            $Config.Email.Password = $null
        }
        [System.GC]::Collect()
        
        return $true
    }
    catch {
        Write-Error "Failed to save secure configuration: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Prompts user for credentials and encrypts them
.DESCRIPTION
    Interactive credential collection with immediate encryption
.PARAMETER Title
    Title for the credential prompt
.PARAMETER Message
    Message to display to user
.RETURNS
    Hashtable with Username and EncryptedPassword
.EXAMPLE
    $creds = Get-SecureCredentialInput -Title "Database" -Message "Enter database credentials"
#>
function Get-SecureCredentialInput {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    try {
        Write-Host $Message -ForegroundColor Yellow
        
        # Get username
        $username = Read-Host "Username"
        
        # Get password securely
        $securePassword = Read-Host "Password" -AsSecureString
        
        # Convert secure string to plain text temporarily for encryption
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
        
        # Encrypt the password
        $encryptedPassword = Protect-ConfigString -PlainText $plainPassword
        
        # Clear sensitive data from memory
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
        $plainPassword = $null
        $securePassword.Dispose()
        [System.GC]::Collect()
        
        return @{
            Username = $username
            EncryptedPassword = $encryptedPassword
        }
    }
    catch {
        Write-Error "Failed to get secure credential input: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Tests if a string is encrypted (Base64 format check)
.DESCRIPTION
    Simple check to determine if a string appears to be encrypted
.PARAMETER String
    String to test
.RETURNS
    Boolean indicating if string appears encrypted
.EXAMPLE
    $isEncrypted = Test-EncryptedString -String $passwordString
#>
function Test-EncryptedString {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$String
    )
    
    if ([string]::IsNullOrEmpty($String)) {
        return $false
    }
    
    try {
        # Try to decode as Base64 - if successful and reasonable length, likely encrypted
        $bytes = [System.Convert]::FromBase64String($String)
        return $bytes.Length -gt 16  # Encrypted data should be at least 16 bytes
    }
    catch {
        return $false
    }
}

<#
.SYNOPSIS
    Migrates existing plain text configuration to encrypted format
.DESCRIPTION
    Converts existing configuration file with plain text passwords to encrypted format
.PARAMETER ConfigPath
    Path to existing configuration file
.PARAMETER BackupPath
    Path to save backup of original file
.EXAMPLE
    Convert-ToSecureConfiguration -ConfigPath ".\config\config.json" -BackupPath ".\config\config.backup.json"
#>
function Convert-ToSecureConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $false)]
        [string]$BackupPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        Write-Host "Converting configuration to secure format..." -ForegroundColor Yellow
        
        # Create backup if path specified
        if ($BackupPath) {
            Copy-Item -Path $ConfigPath -Destination $BackupPath -Force
            Write-Host "Backup created: $BackupPath" -ForegroundColor Green
        }
        
        # Load existing configuration
        $configContent = Get-Content -Path $ConfigPath -Raw
        $config = $configContent | ConvertFrom-Json
        
        $needsUpdate = $false
        
        # Check and convert database password
        if ($config.Database.AuthType -eq "SQL" -and 
            -not [string]::IsNullOrEmpty($config.Database.Password) -and
            -not (Test-EncryptedString -String $config.Database.Password)) {
            
            Write-Host "Encrypting database password..." -ForegroundColor Yellow
            $encryptedDbPassword = Protect-ConfigString -PlainText $config.Database.Password
            
            # Update config object
            $config.Database | Add-Member -NotePropertyName "EncryptedPassword" -NotePropertyValue $encryptedDbPassword -Force
            $config.Database.Password = $null
            $config.Database.PSObject.Properties.Remove("Password")
            
            $needsUpdate = $true
        }
        
        # Check and convert email password
        if (-not [string]::IsNullOrEmpty($config.Email.Password) -and
            -not (Test-EncryptedString -String $config.Email.Password)) {
            
            Write-Host "Encrypting email password..." -ForegroundColor Yellow
            $encryptedEmailPassword = Protect-ConfigString -PlainText $config.Email.Password
            
            # Update config object
            $config.Email | Add-Member -NotePropertyName "EncryptedPassword" -NotePropertyValue $encryptedEmailPassword -Force
            $config.Email.Password = $null
            $config.Email.PSObject.Properties.Remove("Password")
            
            $needsUpdate = $true
        }
        
        # Save updated configuration if changes were made
        if ($needsUpdate) {
            $config | ConvertTo-Json -Depth 3 | Out-File -FilePath $ConfigPath -Encoding UTF8
            Write-Host "Configuration converted to secure format successfully!" -ForegroundColor Green
            Write-Host "Passwords are now encrypted using Windows DPAPI" -ForegroundColor Green
        }
        else {
            Write-Host "Configuration is already in secure format" -ForegroundColor Green
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to convert configuration to secure format: $($_.Exception.Message)"
        throw
    }
}

# Functions are available when dot-sourced
# Note: Export-ModuleMember is not needed when dot-sourcing