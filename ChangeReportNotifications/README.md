# Change Report Email Notifications

A PowerShell script system that automatically sends daily email notifications for critical and high priority change records from a SQL database.

## Features

- **Automated Daily Notifications**: Sends email reports at scheduled times
- **Priority Filtering**: Focuses on Critical and High priority changes only
- **Professional Email Format**: HTML-formatted emails with tables and styling
- **Secure Database Access**: Supports Windows and SQL Server authentication
- **Comprehensive Logging**: Detailed logs with error tracking and audit trails
- **Task Scheduler Integration**: Automated execution via Windows Task Scheduler
- **Flexible Configuration**: JSON-based configuration with validation
- **Error Handling**: Robust error handling with email notifications for failures

## System Requirements

### Prerequisites
- **Operating System**: Windows Server 2012 R2 or later, Windows 10/11
- **PowerShell**: Version 5.1 or higher
- **Database Access**: SQL Server with appropriate permissions
- **Email Access**: SMTP server access for email delivery
- **Permissions**: Administrator rights for Task Scheduler setup

### PowerShell Modules
- **SqlServer**: Version 21.0.0 or higher (automatically installed during setup)

## Installation

### Automated Installation (Recommended)

1. **Download and Extract**: Extract all files to a temporary directory
2. **Run Installation Script**: Execute as Administrator
   ```powershell
   .\Install-ChangeReportScript.ps1
   ```
3. **Follow Prompts**: Provide database and email configuration details
4. **Setup Task Scheduler**: Run the scheduler setup script
   ```powershell
   .\Setup-TaskScheduler.ps1
   ```

### Manual Installation

1. **Create Directory Structure**:
   ```
   C:\ChangeReportNotifications\
   ├── config\
   ├── logs\
   ├── modules\
   └── TestArtifacts\
   ```

2. **Copy Files**: Place all script files in appropriate directories

3. **Install PowerShell Modules**:
   ```powershell
   Install-Module -Name SqlServer -MinimumVersion 21.0.0 -Force
   ```

4. **Create Configuration**: Copy and modify `config.sample.json` to `config.json`

## Configuration

### Database Configuration
```json
{
  "Database": {
    "Server": "SERVER\\INSTANCE",
    "Database": "YourDatabase",
    "AuthType": "Windows",
    "Username": "",
    "Password": ""
  }
}
```

**Authentication Types**:
- **Windows**: Uses current user's Windows credentials
- **SQL**: Requires username and password

### Email Configuration

#### Authenticated SMTP (Username + Password)
```json
{
  "Email": {
    "SMTPServer": "smtp.company.com",
    "Port": 587,
    "EnableSSL": true,
    "From": "noreply@company.com",
    "To": ["admin@company.com", "manager@company.com"],
    "Username": "smtp-user",
    "EncryptedPassword": "AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAA..."
  }
}
```

#### Anonymous SMTP (No Authentication)
```json
{
  "Email": {
    "SMTPServer": "mail.company.com",
    "Port": 25,
    "EnableSSL": false,
    "From": "noreply@company.com",
    "To": ["admin@company.com", "manager@company.com"],
    "Username": "",
    "EncryptedPassword": ""
  }
}
```

#### Username-Only SMTP (Username without Password)
```json
{
  "Email": {
    "SMTPServer": "relay.company.com",
    "Port": 587,
    "EnableSSL": true,
    "From": "noreply@company.com",
    "To": ["admin@company.com", "manager@company.com"],
    "Username": "service-account@company.com",
    "EncryptedPassword": ""
  }
}
```

### Logging Configuration
```json
{
  "Logging": {
    "LogPath": "C:\\ChangeReportNotifications\\logs",
    "LogLevel": "Info"
  }
}
```

**Log Levels**: Debug, Info, Warning, Error

### Schedule Configuration
```json
{
  "Schedule": {
    "ExecutionTime": "08:00"
  }
}
```

## SMTP Authentication Scenarios

The system supports multiple SMTP authentication methods to accommodate different server configurations:

### 1. **Anonymous SMTP** (No Authentication Required)
- **Use Case**: Internal corporate mail servers, development environments
- **Configuration**: Leave `Username` and `EncryptedPassword` empty
- **Ports**: Typically port 25 (unencrypted) or 587 (with STARTTLS)
- **Security**: Relies on network-level security (IP whitelisting, VPN, etc.)

### 2. **Username + Password Authentication**
- **Use Case**: External SMTP services (Gmail, Office 365, SendGrid, etc.)
- **Configuration**: Provide both `Username` and `EncryptedPassword`
- **Ports**: Typically port 587 (STARTTLS) or 465 (SSL/TLS)
- **Security**: Credentials are encrypted using Windows DPAPI

### 3. **Username-Only Authentication**
- **Use Case**: Some corporate SMTP relays that identify users by username only
- **Configuration**: Provide `Username` but leave `EncryptedPassword` empty
- **Ports**: Varies by server configuration
- **Security**: Username-based identification without password verification

### 4. **Default Credentials (Windows Authentication)**
- **Use Case**: SMTP servers that accept current Windows user credentials
- **Configuration**: Leave both `Username` and `EncryptedPassword` empty
- **Behavior**: Uses current user's Windows credentials automatically
- **Security**: Leverages existing Windows authentication

## Usage

### Manual Execution
```powershell
# Standard execution
.\Send-ChangeReport.ps1

# Test mode (doesn't send emails)
.\Send-ChangeReport.ps1 -TestMode

# Verbose output
.\Send-ChangeReport.ps1 -Verbose
```

### Batch File Execution
```cmd
# Simple execution
Run-ChangeReport.bat

# With logging
Run-ChangeReport.bat > execution.log 2>&1
```

### Task Scheduler Management
```powershell
# Test existing task
.\Setup-TaskScheduler.ps1 -TestOnly

# Remove scheduled task
.\Setup-TaskScheduler.ps1 -Remove

# Recreate with different settings
.\Setup-TaskScheduler.ps1 -ExecutionTime "09:30"
```

## Email Output

### Sample Email Content
- **Subject**: "Daily Change Report - [Date] - [X] Critical/High Priority Changes"
- **Summary Section**: Count of changes by priority level
- **Detailed Table**: All change information including:
  - Change ID
  - Priority Level
  - Change Type
  - Configuration Item
  - Short Description
  - Assignment Group
  - Assigned To
  - Actual Start/End Dates

### No Changes Scenario
When no critical or high priority changes are found, a confirmation email is sent stating "No critical or high priority changes for [date]".

## Logging and Monitoring

### Log Files
- **Location**: `logs\ChangeReport_YYYY-MM-DD.log`
- **Rotation**: Daily log files with automatic cleanup
- **Content**: Execution details, errors, email delivery status

### Log Levels
- **Debug**: Detailed execution information
- **Info**: General execution flow
- **Warning**: Non-critical issues
- **Error**: Failures requiring attention

### Monitoring Points
1. **Daily Execution**: Verify task runs at scheduled time
2. **Email Delivery**: Confirm recipients receive emails
3. **Database Connectivity**: Monitor connection success
4. **Log File Growth**: Check for excessive error logging

## Testing

### Manual Testing
```powershell
# Test configuration and email delivery (no actual emails sent)
.\Send-ChangeReport.ps1 -TestMode -Verbose

# Test with verbose output for debugging
.\Send-ChangeReport.ps1 -Verbose

# Test using batch file
.\Run-ChangeReport.bat
```

### Validation Steps
1. **Configuration Test**: Verify config.json is valid JSON
2. **Database Connection**: Ensure database connectivity and query execution
3. **Email Formatting**: Check HTML email template rendering
4. **Task Scheduler**: Verify scheduled task creation and execution

## Security Considerations

### Credential Protection
- **Encrypted Storage**: All passwords are encrypted using Windows DPAPI (Data Protection API)
- **User-Specific Encryption**: Encrypted passwords can only be decrypted by the user who encrypted them
- **No Plain Text**: Configuration files never contain plain text passwords
- **Secure Memory Handling**: Passwords are cleared from memory after use
- **Database Connections**: Use SqlCredential objects instead of connection string passwords

### Database Security
- **Encrypted Connections**: Database connections use encryption by default
- **Least Privilege**: Use dedicated service accounts with minimal required permissions
- **Parameterized Queries**: All SQL queries use parameters to prevent injection attacks
- **Connection Timeouts**: Configurable timeouts prevent hanging connections

### Email Security
- **SMTP Authentication**: Secure credential handling for SMTP authentication
- **SSL/TLS Encryption**: Email transmission uses encrypted connections
- **Credential Objects**: Use .NET NetworkCredential objects for secure authentication
- **Retry Logic**: Secure retry mechanisms for transient failures

### File System Security
- **Configuration Protection**: Restrict access to configuration files
- **Log File Security**: Secure log file permissions and rotation
- **Backup Security**: Encrypted backups of configuration files
- **Temporary Files**: Secure handling and cleanup of temporary files

### Network Security
- **Encrypted Protocols**: Use SSL/TLS for all network communications
- **Certificate Validation**: Proper certificate validation for secure connections
- **Firewall Configuration**: Document required ports and protocols
- **Network Isolation**: Consider network segmentation for database access

### Audit and Monitoring
- **Execution Logging**: Comprehensive logging of all operations
- **Security Events**: Log authentication attempts and failures
- **Configuration Changes**: Monitor and log configuration modifications
- **Access Patterns**: Track database and email access patterns

### Migration from Plain Text
If you have existing configurations with plain text passwords:

```powershell
# Convert existing configuration to secure format
.\Convert-ToSecureConfig.ps1 -ConfigPath ".\config\config.json"

# Verify the conversion
.\Send-ChangeReport.ps1 -TestMode -Verbose
```

## Performance Optimization

### Database Queries
- **Indexing**: Ensure Priority and ActualStartDate columns are indexed
- **Query Timeout**: Default 30 seconds with retry logic
- **Result Limits**: Consider pagination for large result sets

### Email Delivery
- **Size Limits**: Monitor email size for large change lists
- **Delivery Time**: Track SMTP response times
- **Retry Logic**: Automatic retry for transient failures

### System Resources
- **Memory Usage**: Monitor PowerShell process memory
- **Execution Time**: Target completion under 5 minutes
- **Concurrent Execution**: Prevent overlapping executions

## Troubleshooting

### Common Issues and Solutions

#### Installation Problems
- **Permission Errors**: Run PowerShell as Administrator
- **Execution Policy**: Set execution policy with `Set-ExecutionPolicy RemoteSigned`
- **Module Installation**: Manually install SqlServer module if automatic installation fails

#### Database Connection Issues
- **Authentication Failures**: Verify server name, database name, and credentials
- **Network Connectivity**: Test connection with `Test-NetConnection -ComputerName "server" -Port 1433`
- **Permissions**: Ensure database user has SELECT permissions on change table

#### Email Delivery Problems
- **SMTP Connection**: Verify SMTP server, port, and SSL settings
- **Authentication**: Check SMTP username and password if required
- **Firewall**: Ensure SMTP port (usually 587 or 25) is not blocked

#### Task Scheduler Issues
- **Task Not Running**: Verify task is enabled and user has proper permissions
- **Execution Failures**: Check working directory and file paths are absolute
- **User Context**: Ensure task user has "Log on as a batch job" rights

For detailed troubleshooting steps, check the log files in the `logs` directory.

## Support and Maintenance

### Regular Maintenance
- **Log Cleanup**: Implement log rotation and cleanup
- **Configuration Review**: Periodic review of settings
- **Module Updates**: Keep PowerShell modules current
- **Performance Monitoring**: Track execution times and resource usage

### Backup Considerations
- **Configuration Files**: Include in backup procedures
- **Log Files**: Consider retention requirements
- **Script Files**: Version control recommended

### Change Management
- **Testing**: Test all changes in non-production environment
- **Documentation**: Update documentation for configuration changes
- **Rollback Plan**: Maintain previous working configurations

## Directory Structure

```
ChangeReportNotifications/
├── config/
│   ├── config.json              # Main configuration file (created during setup)
│   └── config.sample.json       # Sample configuration template
├── logs/                        # Daily log files (created during execution)
├── modules/
│   ├── Config.ps1              # Configuration management
│   ├── Database.ps1            # Database connectivity
│   ├── Email.ps1               # Email formatting and delivery
│   └── Logging.ps1             # Logging functionality
├── Install-ChangeReportScript.ps1  # Automated installation
├── Setup-TaskScheduler.ps1     # Task Scheduler configuration
├── Send-ChangeReport.ps1       # Main orchestrator script
├── Run-ChangeReport.bat        # Batch file wrapper
└── README.md                   # This documentation
```

**Note**: Test files, additional documentation, and generated artifacts are excluded from version control but may be created during installation and testing.

## Version History

- **v1.0**: Initial release with core functionality
- **v1.1**: Added comprehensive error handling and logging
- **v1.2**: Enhanced email formatting and Task Scheduler integration
- **v1.3**: Added installation and setup automation scripts

## License

This project is licensed under the MIT License - see the LICENSE file for details.