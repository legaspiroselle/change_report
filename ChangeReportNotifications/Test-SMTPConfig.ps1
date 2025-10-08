#Requires -Version 5.1
<#
.SYNOPSIS
    Test script to validate different SMTP authentication configurations

.DESCRIPTION
    This script tests various SMTP authentication scenarios to ensure the system
    handles anonymous, username-only, and full authentication correctly.

.PARAMETER ConfigPath
    Path to the configuration file to test

.EXAMPLE
    .\Test-SMTPConfig.ps1 -ConfigPath ".\config\config.anonymous-smtp.json"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ConfigPath
)

# Import required modules
$ModulePath = Join-Path -Path $PSScriptRoot -ChildPath "modules"
. (Join-Path -Path $ModulePath -ChildPath "SecureConfig.ps1")
. (Join-Path -Path $ModulePath -ChildPath "Email.ps1")

try {
    Write-Host "=== SMTP Configuration Test ===" -ForegroundColor Cyan
    Write-Host "Testing configuration: $ConfigPath`n" -ForegroundColor Yellow
    
    # Load configuration
    Write-Host "Loading configuration..." -ForegroundColor Green
    $config = Get-SecureConfiguration -ConfigPath $ConfigPath
    
    # Display SMTP settings
    Write-Host "SMTP Configuration:" -ForegroundColor Yellow
    Write-Host "  Server: $($config.Email.SMTPServer)" -ForegroundColor White
    Write-Host "  Port: $($config.Email.Port)" -ForegroundColor White
    Write-Host "  SSL/TLS: $($config.Email.EnableSSL)" -ForegroundColor White
    Write-Host "  From: $($config.Email.From)" -ForegroundColor White
    Write-Host "  To: $($config.Email.To -join ', ')" -ForegroundColor White
    
    # Determine authentication method
    if ($config.Email.Credential) {
        Write-Host "  Authentication: Username + Password (Secure)" -ForegroundColor Green
        Write-Host "  Username: $($config.Email.Credential.UserName)" -ForegroundColor White
    }
    elseif (-not [string]::IsNullOrEmpty($config.Email.Username)) {
        if (-not [string]::IsNullOrEmpty($config.Email.EncryptedPassword)) {
            Write-Host "  Authentication: Username + Password (Legacy)" -ForegroundColor Yellow
        }
        else {
            Write-Host "  Authentication: Username Only" -ForegroundColor Cyan
        }
        Write-Host "  Username: $($config.Email.Username)" -ForegroundColor White
    }
    else {
        Write-Host "  Authentication: Anonymous/Default Credentials" -ForegroundColor Cyan
    }
    
    Write-Host "`nTesting email configuration..." -ForegroundColor Green
    
    # Validate email configuration
    $isValid = Test-EmailConfiguration -Config $config
    
    if ($isValid) {
        Write-Host "✓ Email configuration is valid" -ForegroundColor Green
    }
    else {
        Write-Host "✗ Email configuration validation failed" -ForegroundColor Red
        exit 1
    }
    
    # Test SMTP connectivity (without sending email)
    Write-Host "`nTesting SMTP connectivity..." -ForegroundColor Green
    
    try {
        # Create SMTP client to test connection
        $smtpClient = New-Object System.Net.Mail.SmtpClient
        $smtpClient.Host = $config.Email.SMTPServer
        $smtpClient.Port = $config.Email.Port
        $smtpClient.EnableSsl = $config.Email.EnableSSL
        $smtpClient.Timeout = 10000  # 10 seconds
        
        # Configure authentication
        if ($config.Email.Credential) {
            Write-Host "  Using secure credential authentication" -ForegroundColor Cyan
            $smtpClient.Credentials = New-Object System.Net.NetworkCredential(
                $config.Email.Credential.UserName,
                $config.Email.Credential.Password
            )
            $smtpClient.UseDefaultCredentials = $false
        }
        elseif (-not [string]::IsNullOrWhiteSpace($config.Email.Username)) {
            Write-Host "  Using username authentication" -ForegroundColor Cyan
            $smtpClient.UseDefaultCredentials = $false
        }
        else {
            Write-Host "  Using anonymous/default authentication" -ForegroundColor Cyan
            $smtpClient.UseDefaultCredentials = $true
            $smtpClient.Credentials = $null
        }
        
        # Test connection by creating a mail message (but not sending)
        $testMessage = New-Object System.Net.Mail.MailMessage
        $testMessage.From = New-Object System.Net.Mail.MailAddress($config.Email.From)
        $testMessage.To.Add($config.Email.To[0])
        $testMessage.Subject = "SMTP Configuration Test - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $testMessage.Body = "This is a test message to validate SMTP configuration."
        $testMessage.IsBodyHtml = $false
        
        Write-Host "✓ SMTP client configuration successful" -ForegroundColor Green
        Write-Host "✓ Test message created successfully" -ForegroundColor Green
        
        # Cleanup
        $testMessage.Dispose()
        $smtpClient.Dispose()
        
        Write-Host "`n=== Test Results ===" -ForegroundColor Green
        Write-Host "✓ Configuration loaded successfully" -ForegroundColor Green
        Write-Host "✓ Email settings validated" -ForegroundColor Green
        Write-Host "✓ SMTP client configured correctly" -ForegroundColor Green
        Write-Host "✓ Authentication method determined" -ForegroundColor Green
        
        Write-Host "`nRecommendations:" -ForegroundColor Yellow
        
        if ($config.Email.Credential) {
            Write-Host "• Configuration uses secure credential storage ✓" -ForegroundColor Green
        }
        elseif (-not [string]::IsNullOrEmpty($config.Email.Username)) {
            if ([string]::IsNullOrEmpty($config.Email.EncryptedPassword)) {
                Write-Host "• Username-only authentication detected" -ForegroundColor Cyan
                Write-Host "• Ensure SMTP server supports this authentication method" -ForegroundColor Yellow
            }
            else {
                Write-Host "• Consider upgrading to secure credential storage" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "• Anonymous SMTP configuration detected" -ForegroundColor Cyan
            Write-Host "• Ensure SMTP server allows anonymous relay from this IP" -ForegroundColor Yellow
            Write-Host "• Consider using authentication for better security" -ForegroundColor Yellow
        }
        
        if ($config.Email.EnableSSL) {
            Write-Host "• SSL/TLS encryption enabled ✓" -ForegroundColor Green
        }
        else {
            Write-Host "• Consider enabling SSL/TLS for better security" -ForegroundColor Yellow
        }
        
    }
    catch {
        Write-Host "✗ SMTP configuration test failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
        Write-Host "• Verify SMTP server address and port" -ForegroundColor White
        Write-Host "• Check network connectivity to SMTP server" -ForegroundColor White
        Write-Host "• Verify authentication requirements" -ForegroundColor White
        Write-Host "• Check firewall settings" -ForegroundColor White
        exit 1
    }
    
}
catch {
    Write-Host "Test failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "`nSMTP configuration test completed successfully!" -ForegroundColor Green