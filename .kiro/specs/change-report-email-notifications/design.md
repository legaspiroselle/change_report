# Design Document

## Overview

The Change Report Email Notification system is a PowerShell script that connects to a SQL Server database, queries for critical and high priority change records, and sends formatted email notifications daily. The solution uses native PowerShell modules for database connectivity and email functionality, ensuring compatibility with Windows environments and easy integration with Task Scheduler.

## Architecture

The system follows a modular PowerShell script architecture with the following components:

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Task Scheduler │───▶│  Main Script     │───▶│  Email Module   │
└─────────────────┘    │  (Orchestrator)  │    └─────────────────┘
                       └──────────────────┘
                                │
                                ▼
                       ┌──────────────────┐    ┌─────────────────┐
                       │  Database Module │───▶│  Config Module  │
                       └──────────────────┘    └─────────────────┘
```

### Core Components:
1. **Main Orchestrator Script** - Entry point that coordinates all operations
2. **Configuration Module** - Handles loading and validation of settings
3. **Database Module** - Manages SQL Server connections and queries
4. **Email Module** - Formats and sends email notifications
5. **Logging Module** - Provides error handling and audit trail

## Components and Interfaces

### 1. Configuration Module (`Config.ps1`)

**Purpose:** Centralized configuration management with validation

**Interface:**
```powershell
# Configuration structure
$Config = @{
    Database = @{
        Server = "server-name"
        Database = "database-name"
        AuthType = "Windows" # or "SQL"
        Username = "" # for SQL auth
        Password = "" # for SQL auth (encrypted)
    }
    Email = @{
        SMTPServer = "smtp.company.com"
        Port = 587
        EnableSSL = $true
        From = "noreply@company.com"
        To = @("admin@company.com", "manager@company.com")
        Username = "" # SMTP auth
        Password = "" # SMTP auth (encrypted)
    }
    Logging = @{
        LogPath = "C:\Logs\ChangeReports"
        LogLevel = "Info" # Debug, Info, Warning, Error
    }
}
```

**Functions:**
- `Get-Configuration()` - Loads config from JSON file
- `Test-Configuration($Config)` - Validates all config parameters
- `ConvertTo-SecureString($PlainText)` - Encrypts sensitive data

### 2. Database Module (`Database.ps1`)

**Purpose:** SQL Server connectivity and change record retrieval

**Interface:**
```powershell
# Change record structure
$ChangeRecord = @{
    ChangeID = ""
    Priority = ""
    Type = ""
    ConfigurationItem = ""
    ShortDescription = ""
    AssignmentsGroup = ""
    AssignedTo = ""
    ActualStartDate = ""
    ActualEndDate = ""
}
```

**Functions:**
- `Connect-Database($Config)` - Establishes secure database connection
- `Get-CriticalChanges($Connection, $Date)` - Retrieves filtered change records
- `Close-Database($Connection)` - Properly closes database connections
- `Test-DatabaseConnection($Config)` - Validates connectivity

**SQL Query Design:**
```sql
SELECT 
    ChangeID,
    Priority,
    Type,
    ConfigurationItem,
    ShortDescription,
    AssignmentsGroup,
    AssignedTo,
    ActualStartDate,
    ActualEndDate
FROM ChangeTable 
WHERE Priority IN ('Critical', 'High')
    AND CAST(ActualStartDate AS DATE) = CAST(@ReportDate AS DATE)
ORDER BY 
    CASE Priority 
        WHEN 'Critical' THEN 1 
        WHEN 'High' THEN 2 
    END,
    ActualStartDate
```

### 3. Email Module (`Email.ps1`)

**Purpose:** Email formatting and delivery

**Functions:**
- `Format-ChangeReport($Changes, $Date)` - Creates HTML email body
- `Send-ChangeNotification($Config, $Body, $Subject)` - Sends email via SMTP
- `Format-ChangeTable($Changes)` - Converts change data to HTML table
- `Get-EmailSubject($Date, $ChangeCount)` - Generates dynamic subject line

**Email Template Design:**
- HTML format with CSS styling for professional appearance
- Summary section with change counts by priority
- Detailed table with all change information
- Footer with generation timestamp and system information

### 4. Logging Module (`Logging.ps1`)

**Purpose:** Comprehensive logging and error handling

**Functions:**
- `Write-Log($Message, $Level, $LogPath)` - Writes timestamped log entries
- `Start-LogSession($LogPath)` - Initializes daily log file
- `Send-ErrorNotification($Config, $Error)` - Sends failure alerts
- `Get-LogFileName($Date)` - Generates date-based log file names

### 5. Main Script (`Send-ChangeReport.ps1`)

**Purpose:** Orchestrates the entire process

**Workflow:**
1. Load and validate configuration
2. Initialize logging session
3. Connect to database
4. Query for critical/high priority changes
5. Format email content
6. Send email notification
7. Log results and cleanup
8. Handle errors gracefully

## Data Models

### Change Record Model
```powershell
class ChangeRecord {
    [string]$ChangeID
    [string]$Priority
    [string]$Type
    [string]$ConfigurationItem
    [string]$ShortDescription
    [string]$AssignmentsGroup
    [string]$AssignedTo
    [datetime]$ActualStartDate
    [datetime]$ActualEndDate
    
    [string] ToString() {
        return "$($this.ChangeID) - $($this.Priority) - $($this.ShortDescription)"
    }
}
```

### Configuration Model
```powershell
class EmailConfig {
    [string]$SMTPServer
    [int]$Port
    [bool]$EnableSSL
    [string]$From
    [string[]]$To
    [string]$Username
    [securestring]$Password
}

class DatabaseConfig {
    [string]$Server
    [string]$Database
    [string]$AuthType
    [string]$Username
    [securestring]$Password
}
```

## Error Handling

### Database Connection Errors
- **Retry Logic:** 3 attempts with exponential backoff
- **Fallback:** Send notification about database unavailability
- **Logging:** Detailed connection error information

### Email Delivery Errors
- **Retry Logic:** 2 attempts for transient failures
- **Fallback:** Log email content to file for manual review
- **Notification:** Alert to backup email address if configured

### Configuration Errors
- **Validation:** Pre-flight checks for all required parameters
- **Graceful Degradation:** Use default values where appropriate
- **User Feedback:** Clear error messages for missing/invalid config

### SQL Query Errors
- **Parameter Validation:** Ensure date parameters are valid
- **Query Timeout:** 30-second timeout with retry
- **Result Validation:** Verify expected columns are returned

## Testing Strategy

### Unit Testing Approach
- **PowerShell Pester Framework:** For individual function testing
- **Mock Objects:** Database connections and SMTP servers
- **Test Data:** Sample change records with various scenarios
- **Configuration Testing:** Valid and invalid config combinations

### Integration Testing
- **Database Connectivity:** Test with actual SQL Server instance
- **Email Delivery:** Test with development SMTP server
- **End-to-End:** Full workflow with test data
- **Error Scenarios:** Network failures, invalid credentials

### Test Scenarios
1. **Normal Operation:** Critical and high priority changes exist
2. **No Changes:** Empty result set handling
3. **Database Failure:** Connection timeout and retry logic
4. **Email Failure:** SMTP server unavailable
5. **Configuration Issues:** Missing or invalid parameters
6. **Large Result Sets:** Performance with many change records

### Performance Considerations
- **Query Optimization:** Indexed columns for Priority and ActualStartDate
- **Memory Management:** Stream large result sets
- **Email Size Limits:** Pagination for large change lists
- **Execution Time:** Target completion under 5 minutes

## Security Considerations

### Credential Management
- **Encrypted Storage:** Use PowerShell SecureString for passwords
- **Windows Credential Manager:** Store sensitive data securely
- **Least Privilege:** Database user with read-only access
- **SMTP Authentication:** Secure email server credentials

### Data Protection
- **SQL Injection Prevention:** Parameterized queries only
- **Email Content:** No sensitive data in email headers
- **Log Security:** Restrict access to log files
- **Network Security:** Use SSL/TLS for all connections

### Audit Trail
- **Execution Logging:** Record all script executions
- **Access Logging:** Database connection attempts
- **Email Tracking:** Delivery confirmation and failures
- **Configuration Changes:** Log config file modifications