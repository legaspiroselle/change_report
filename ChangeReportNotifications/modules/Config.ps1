# Configuration Module for Change Report Email Notifications
# Handles loading, validation, and secure string conversion for configuration parameters

<#
.SYNOPSIS
    Loads configuration from JSON file and validates parameters
.DESCRIPTION
    Reads the configuration file and returns a hashtable with all settings
.PARAMETER ConfigPath
    Path to the JSON configuration file
.EXAMPLE
    $config = Get-Configuration -ConfigPath ".\config\config.json"
#>
function Get-Configuration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }
        
        $configContent = Get-Content -Path $ConfigPath -Raw
        $config = $configContent | ConvertFrom-Json
        
        # Convert PSCustomObject to hashtable for easier manipulation
        $configHash = @{}
        
        # Database configuration
        $configHash.Database = @{
            Server = $config.Database.Server
            Database = $config.Database.Database
            AuthType = $config.Database.AuthType
            Username = $config.Database.Username
            Password = $config.Database.Password
        }
        
        # Email configuration
        $configHash.Email = @{
            SMTPServer = $config.Email.SMTPServer
            Port = $config.Email.Port
            EnableSSL = $config.Email.EnableSSL
            From = $config.Email.From
            To = $config.Email.To
            Username = $config.Email.Username
            Password = $config.Email.Password
        }
        
        # Logging configuration
        $configHash.Logging = @{
            LogPath = $config.Logging.LogPath
            LogLevel = $config.Logging.LogLevel
        }
        
        return $configHash
    }
    catch {
        Write-Error "Failed to load configuration: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Validates configuration parameters for completeness and correctness
.DESCRIPTION
    Checks all required configuration parameters and validates their values
.PARAMETER Config
    Configuration hashtable to validate
.EXAMPLE
    $isValid = Test-Configuration -Config $config
#>
function Test-Configuration {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    $isValid = $true
    $errors = @()
    
    # Validate Database configuration
    if (-not $Config.Database) {
        $errors += "Database configuration section is missing"
        $isValid = $false
    } else {
        if ([string]::IsNullOrWhiteSpace($Config.Database.Server)) {
            $errors += "Database server is required"
            $isValid = $false
        }
        
        if ([string]::IsNullOrWhiteSpace($Config.Database.Database)) {
            $errors += "Database name is required"
            $isValid = $false
        }
        
        if ($Config.Database.AuthType -notin @("Windows", "SQL")) {
            $errors += "Database AuthType must be 'Windows' or 'SQL'"
            $isValid = $false
        }
        
        if ($Config.Database.AuthType -eq "SQL") {
            if ([string]::IsNullOrWhiteSpace($Config.Database.Username)) {
                $errors += "Database username is required for SQL authentication"
                $isValid = $false
            }
            if ([string]::IsNullOrWhiteSpace($Config.Database.Password)) {
                $errors += "Database password is required for SQL authentication"
                $isValid = $false
            }
        }
    }
    
    # Validate Email configuration
    if (-not $Config.Email) {
        $errors += "Email configuration section is missing"
        $isValid = $false
    } else {
        if ([string]::IsNullOrWhiteSpace($Config.Email.SMTPServer)) {
            $errors += "SMTP server is required"
            $isValid = $false
        }
        
        if ($Config.Email.Port -lt 1 -or $Config.Email.Port -gt 65535) {
            $errors += "SMTP port must be between 1 and 65535"
            $isValid = $false
        }
        
        if ([string]::IsNullOrWhiteSpace($Config.Email.From)) {
            $errors += "From email address is required"
            $isValid = $false
        }
        
        if (-not $Config.Email.To -or $Config.Email.To.Count -eq 0) {
            $errors += "At least one recipient email address is required"
            $isValid = $false
        }
        
        # Validate email format for From address
        if ($Config.Email.From -and $Config.Email.From -notmatch "^[^@]+@[^@]+\.[^@]+$") {
            $errors += "From email address format is invalid"
            $isValid = $false
        }
        
        # Validate email format for To addresses
        foreach ($email in $Config.Email.To) {
            if ($email -notmatch "^[^@]+@[^@]+\.[^@]+$") {
                $errors += "Recipient email address format is invalid: $email"
                $isValid = $false
            }
        }
    }
    
    # Validate Logging configuration
    if (-not $Config.Logging) {
        $errors += "Logging configuration section is missing"
        $isValid = $false
    } else {
        if ([string]::IsNullOrWhiteSpace($Config.Logging.LogPath)) {
            $errors += "Log path is required"
            $isValid = $false
        }
        
        if ($Config.Logging.LogLevel -notin @("Debug", "Info", "Warning", "Error")) {
            $errors += "Log level must be one of: Debug, Info, Warning, Error"
            $isValid = $false
        }
    }
    
    # Output validation errors if any
    if ($errors.Count -gt 0) {
        Write-Warning "Configuration validation failed:"
        foreach ($error in $errors) {
            Write-Warning "  - $error"
        }
    }
    
    return $isValid
}

<#
.SYNOPSIS
    Converts plain text to secure string for password protection
.DESCRIPTION
    Takes a plain text string and converts it to a SecureString object
.PARAMETER PlainText
    The plain text string to convert
.EXAMPLE
    $securePassword = ConvertTo-SecureString -PlainText "mypassword"
#>
function ConvertTo-ConfigSecureString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PlainText
    )
    
    try {
        return $PlainText | ConvertTo-SecureString -AsPlainText -Force
    }
    catch {
        Write-Error "Failed to convert string to secure string: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Converts secure string back to plain text (use with caution)
.DESCRIPTION
    Converts a SecureString back to plain text for use in connections
.PARAMETER SecureString
    The SecureString to convert
.EXAMPLE
    $plainText = ConvertFrom-SecureString -SecureString $securePassword
#>
function ConvertFrom-ConfigSecureString {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]$SecureString
    )
    
    try {
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    catch {
        Write-Error "Failed to convert secure string to plain text: $($_.Exception.Message)"
        throw
    }
    finally {
        if ($ptr) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
        }
    }
}

# Export functions for module use (only when running as a module)
if ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript') {
    # Running as script, don't export
} else {
    # Functions are available when dot-sourced
    # Note: Export-ModuleMember is not needed when dot-sourcing
}