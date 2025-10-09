# Email Module for Change Report Notifications
# Handles HTML email formatting and delivery functionality

<#
.SYNOPSIS
    Email module for formatting and sending change report notifications
.DESCRIPTION
    This module provides functions for creating HTML-formatted emails with change data
    and sending them via SMTP with proper authentication and error handling
#>

function Format-ChangeReport {
    <#
    .SYNOPSIS
        Formats change records into HTML email content
    .PARAMETER Changes
        Array of change record objects to format
    .PARAMETER Date
        Report date for the email header
    .RETURNS
        HTML formatted email body as string
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Changes,
        
        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )
    
    try {
        $formattedDate = $Date.ToString("MMMM dd, yyyy")
        $criticalCount = ($Changes | Where-Object { $_.Priority -eq "Critical" }).Count
        $highCount = ($Changes | Where-Object { $_.Priority -eq "High" }).Count
        $totalCount = $Changes.Count
        
        # Get HTML template and replace placeholders
        $htmlTemplate = Get-EmailTemplate
        $changeTable = Format-ChangeTable -Changes $Changes
        
        $htmlBody = $htmlTemplate -replace "{{REPORT_DATE}}", $formattedDate
        $htmlBody = $htmlBody -replace "{{TOTAL_COUNT}}", $totalCount
        $htmlBody = $htmlBody -replace "{{CRITICAL_COUNT}}", $criticalCount
        $htmlBody = $htmlBody -replace "{{HIGH_COUNT}}", $highCount
        $htmlBody = $htmlBody -replace "{{CHANGE_TABLE}}", $changeTable
        $htmlBody = $htmlBody -replace "{{GENERATION_TIME}}", (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        
        return $htmlBody
    }
    catch {
        Write-Error "Failed to format change report: $($_.Exception.Message)"
        throw
    }
}

function Format-ChangeTable {
    <#
    .SYNOPSIS
        Converts change data to HTML table format
    .PARAMETER Changes
        Array of change record objects
    .RETURNS
        HTML table string with formatted change data
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Changes
    )
    
    try {
        if ($Changes.Count -eq 0) {
            return "<tr><td colspan='9' style='text-align: center; font-style: italic; color: #666;'>No critical or high priority changes found for this date.</td></tr>"
        }
        
        $tableRows = ""
        
        # Sort changes by priority (Critical first, then High) and then by start date
        $sortedChanges = $Changes | Sort-Object @{
            Expression = {
                switch ($_.Priority) {
                    "Critical" { 1 }
                    "High" { 2 }
                    default { 3 }
                }
            }
        }, ActualStartDate
        
        foreach ($change in $sortedChanges) {
            $priorityClass = switch ($change.Priority) {
                "Critical" { "priority-critical" }
                "High" { "priority-high" }
                default { "priority-normal" }
            }
            
            $startDate = if ($change.ActualStartDate) { 
                ([datetime]$change.ActualStartDate).ToString("MM/dd/yyyy HH:mm") 
            } else { 
                "Not Set" 
            }
            
            $endDate = if ($change.ActualEndDate) { 
                ([datetime]$change.ActualEndDate).ToString("MM/dd/yyyy HH:mm") 
            } else { 
                "Not Set" 
            }
            
            $tableRows += @"
                <tr>
                    <td>$($change.ChangeID)</td>
                    <td><span class="$priorityClass">$($change.Priority)</span></td>
                    <td>$($change.Type)</td>
                    <td>$($change.ConfigurationItem)</td>
                    <td>$($change.ShortDescription)</td>
                    <td>$($change.AssignmentsGroup)</td>
                    <td>$($change.AssignedTo)</td>
                    <td>$startDate</td>
                    <td>$endDate</td>
                </tr>
"@
        }
        
        return $tableRows
    }
    catch {
        Write-Error "Failed to format change table: $($_.Exception.Message)"
        throw
    }
}

function Get-EmailSubject {
    <#
    .SYNOPSIS
        Generates dynamic email subject line
    .PARAMETER Date
        Report date
    .PARAMETER ChangeCount
        Number of changes in the report
    .RETURNS
        Formatted email subject string
    #>
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$Date,
        
        [Parameter(Mandatory = $true)]
        [int]$ChangeCount
    )
    
    try {
        $formattedDate = $Date.ToString("yyyy-MM-dd")
        
        if ($ChangeCount -eq 0) {
            return "Daily Change Report - No Critical/High Priority Changes - $formattedDate"
        }
        elseif ($ChangeCount -eq 1) {
            return "Daily Change Report - 1 Critical/High Priority Change - $formattedDate"
        }
        else {
            return "Daily Change Report - $ChangeCount Critical/High Priority Changes - $formattedDate"
        }
    }
    catch {
        Write-Error "Failed to generate email subject: $($_.Exception.Message)"
        throw
    }
}

function Get-NoChangesEmailTemplate {
    <#
    .SYNOPSIS
        Returns HTML email template for no-changes scenario
    .PARAMETER Date
        Report date
    .RETURNS
        HTML formatted email body for no changes
    #>
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )
    
    $formattedDate = $Date.ToString("MMMM dd, yyyy")
    
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Change Report - No Changes</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 300;
        }
        .header .date {
            font-size: 16px;
            opacity: 0.9;
            margin-top: 5px;
        }
        .content {
            padding: 40px 30px;
            text-align: center;
        }
        .no-changes-icon {
            font-size: 64px;
            color: #28a745;
            margin-bottom: 20px;
        }
        .message {
            font-size: 18px;
            color: #495057;
            margin-bottom: 20px;
        }
        .sub-message {
            font-size: 14px;
            color: #6c757d;
            margin-bottom: 30px;
        }
        .info-box {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            border-radius: 6px;
            padding: 20px;
            margin: 20px 0;
        }
        .info-box h3 {
            color: #0c5460;
            margin-top: 0;
        }
        .info-box p {
            color: #0c5460;
            margin-bottom: 0;
        }
        .footer {
            background-color: #f8f9fa;
            padding: 20px 30px;
            text-align: center;
            color: #6c757d;
            font-size: 12px;
            border-top: 1px solid #e9ecef;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Daily Change Report</h1>
            <div class="date">$formattedDate</div>
        </div>
        
        <div class="content">
            <div class="no-changes-icon">✅</div>
            <div class="message">
                <strong>No Critical or High Priority Changes</strong>
            </div>
            <div class="sub-message">
                No changes with Critical or High priority were found for $formattedDate
            </div>
            
            <div class="info-box">
                <h3>What This Means</h3>
                <p>
                    This is a positive indicator that no urgent changes requiring immediate attention 
                    were scheduled or executed on this date. The system continues to monitor for 
                    critical and high priority changes daily.
                </p>
            </div>
            
            <div class="info-box">
                <h3>Next Steps</h3>
                <p>
                    No action is required. You will receive another report tomorrow, or immediately 
                    if any critical or high priority changes are detected.
                </p>
            </div>
        </div>
        
        <div class="footer">
            <p>This report was automatically generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p>Change Report Notification System</p>
        </div>
    </div>
</body>
</html>
"@
}

function Get-EmailTemplate {
    <#
    .SYNOPSIS
        Returns the HTML email template
    .RETURNS
        HTML template string with placeholders
    #>
    
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Daily Change Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 300;
        }
        .header .date {
            font-size: 16px;
            opacity: 0.9;
            margin-top: 5px;
        }
        .summary {
            padding: 25px 30px;
            background-color: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        .summary h2 {
            margin: 0 0 15px 0;
            color: #495057;
            font-size: 20px;
        }
        .summary-stats {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        .stat-item {
            background: white;
            padding: 15px 20px;
            border-radius: 6px;
            border-left: 4px solid #007bff;
            flex: 1;
            min-width: 150px;
        }
        .stat-item.critical {
            border-left-color: #dc3545;
        }
        .stat-item.high {
            border-left-color: #fd7e14;
        }
        .stat-number {
            font-size: 24px;
            font-weight: bold;
            color: #495057;
        }
        .stat-label {
            font-size: 14px;
            color: #6c757d;
            margin-top: 5px;
        }
        .content {
            padding: 30px;
        }
        .content h2 {
            color: #495057;
            margin-bottom: 20px;
            font-size: 20px;
        }
        .table-container {
            overflow-x: auto;
            margin-bottom: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        th {
            background: #495057;
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: 600;
            font-size: 14px;
        }
        td {
            padding: 12px 8px;
            border-bottom: 1px solid #e9ecef;
            font-size: 13px;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .priority-critical {
            background-color: #dc3545;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
        }
        .priority-high {
            background-color: #fd7e14;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
        }
        .footer {
            background-color: #f8f9fa;
            padding: 20px 30px;
            text-align: center;
            color: #6c757d;
            font-size: 12px;
            border-top: 1px solid #e9ecef;
        }
        @media (max-width: 768px) {
            body {
                padding: 10px;
            }
            .header {
                padding: 20px;
            }
            .header h1 {
                font-size: 24px;
            }
            .summary, .content {
                padding: 20px;
            }
            .summary-stats {
                flex-direction: column;
            }
            th, td {
                padding: 8px 4px;
                font-size: 12px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Daily Change Report</h1>
            <div class="date">{{REPORT_DATE}}</div>
        </div>
        
        <div class="summary">
            <h2>Summary</h2>
            <div class="summary-stats">
                <div class="stat-item">
                    <div class="stat-number">{{TOTAL_COUNT}}</div>
                    <div class="stat-label">Total Changes</div>
                </div>
                <div class="stat-item critical">
                    <div class="stat-number">{{CRITICAL_COUNT}}</div>
                    <div class="stat-label">Critical Priority</div>
                </div>
                <div class="stat-item high">
                    <div class="stat-number">{{HIGH_COUNT}}</div>
                    <div class="stat-label">High Priority</div>
                </div>
            </div>
        </div>
        
        <div class="content">
            <h2>Change Details</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Change ID</th>
                            <th>Priority</th>
                            <th>Type</th>
                            <th>Configuration Item</th>
                            <th>Description</th>
                            <th>Assignment Group</th>
                            <th>Assigned To</th>
                            <th>Start Date</th>
                            <th>End Date</th>
                        </tr>
                    </thead>
                    <tbody>
                        {{CHANGE_TABLE}}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>This report was automatically generated on {{GENERATION_TIME}}</p>
            <p>Change Report Notification System</p>
        </div>
    </div>
</body>
</html>
"@
}

function Send-ChangeNotification {
    <#
    .SYNOPSIS
        Sends change notification email via SMTP
    .PARAMETER Config
        Configuration object containing email settings
    .PARAMETER Body
        HTML email body content
    .PARAMETER Subject
        Email subject line
    .RETURNS
        Boolean indicating success/failure
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Mandatory = $true)]
        [string]$Subject
    )
    
    $maxRetries = 2
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -le $maxRetries) {
        try {
            Write-Verbose "Attempting to send email (attempt $($retryCount + 1) of $($maxRetries + 1))"
            
            # Create mail message
            $mailMessage = New-Object System.Net.Mail.MailMessage
            $mailMessage.From = New-Object System.Net.Mail.MailAddress($Config.Email.From)
            $mailMessage.Subject = $Subject
            $mailMessage.Body = $Body
            $mailMessage.IsBodyHtml = $true
            $mailMessage.Priority = [System.Net.Mail.MailPriority]::Normal
            
            # Add recipients
            foreach ($recipient in $Config.Email.To) {
                if (-not [string]::IsNullOrWhiteSpace($recipient)) {
                    $mailMessage.To.Add($recipient)
                }
            }
            
            if ($mailMessage.To.Count -eq 0) {
                throw "No valid recipients specified in configuration"
            }
            
            # Create SMTP client
            $smtpClient = New-Object System.Net.Mail.SmtpClient
            $smtpClient.Host = $Config.Email.SMTPServer
            $smtpClient.Port = $Config.Email.Port
            $smtpClient.EnableSsl = $Config.Email.EnableSSL
            
            # Configure authentication based on configuration
            if ($Config.Email.Credential) {
                # Use secure credential object (preferred method)
                Write-Verbose "Using secure credential authentication for SMTP"
                $smtpClient.Credentials = New-Object System.Net.NetworkCredential(
                    $Config.Email.Credential.UserName,
                    $Config.Email.Credential.Password
                )
                $smtpClient.UseDefaultCredentials = $false
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Config.Email.Username)) {
                # Fallback for legacy configuration or plain text credentials
                Write-Verbose "Using username/password authentication for SMTP"
                if (-not [string]::IsNullOrWhiteSpace($Config.Email.Password)) {
                    Write-Warning "Using legacy email authentication method. Consider updating to secure configuration."
                    $credential = New-Object System.Net.NetworkCredential(
                        $Config.Email.Username,
                        $Config.Email.Password
                    )
                    $smtpClient.Credentials = $credential
                    $smtpClient.UseDefaultCredentials = $false
                }
                else {
                    Write-Warning "Username provided but no password found. Using default credentials."
                    $smtpClient.UseDefaultCredentials = $true
                }
            }
            else {
                # No credentials provided - use anonymous or default credentials
                Write-Verbose "No SMTP credentials provided - using anonymous/default authentication"
                $smtpClient.UseDefaultCredentials = $true
                $smtpClient.Credentials = $null
            }
            
            # Set timeout (30 seconds)
            $smtpClient.Timeout = 30000
            
            # Send the email
            $smtpClient.Send($mailMessage)
            
            Write-Verbose "Email sent successfully to $($mailMessage.To.Count) recipients"
            $success = $true
            
            # Cleanup
            $mailMessage.Dispose()
            $smtpClient.Dispose()
            
            return $true
        }
        catch [System.Net.Mail.SmtpException] {
            $retryCount++
            $errorMessage = "SMTP Error: $($_.Exception.Message)"
            
            if ($retryCount -le $maxRetries) {
                Write-Warning "$errorMessage - Retrying in 5 seconds... (Attempt $retryCount of $maxRetries)"
                Start-Sleep -Seconds 5
            }
            else {
                Write-Error "$errorMessage - All retry attempts failed"
                
                # Try to save email content to file as fallback
                try {
                    $fallbackPath = Join-Path $Config.Logging.LogPath "failed_email_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
                    $Body | Out-File -FilePath $fallbackPath -Encoding UTF8
                    Write-Warning "Email content saved to fallback file: $fallbackPath"
                }
                catch {
                    Write-Error "Failed to save email content to fallback file: $($_.Exception.Message)"
                }
                
                throw $_.Exception
            }
        }
        catch [System.Exception] {
            $retryCount++
            $errorMessage = "Email delivery error: $($_.Exception.Message)"
            
            if ($retryCount -le $maxRetries -and $_.Exception.Message -match "timeout|network|connection") {
                Write-Warning "$errorMessage - Retrying in 5 seconds... (Attempt $retryCount of $maxRetries)"
                Start-Sleep -Seconds 5
            }
            else {
                Write-Error "$errorMessage - Cannot retry this error type or max retries exceeded"
                throw $_.Exception
            }
        }
        finally {
            # Ensure cleanup even if exceptions occur
            if ($mailMessage) {
                try { $mailMessage.Dispose() } catch { }
            }
            if ($smtpClient) {
                try { $smtpClient.Dispose() } catch { }
            }
        }
    }
    
    return $false
}

function Test-EmailConfiguration {
    <#
    .SYNOPSIS
        Validates email configuration parameters
    .PARAMETER Config
        Configuration object to validate
    .RETURNS
        Boolean indicating if configuration is valid
    #>
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        $errors = @()
        
        # Check required email configuration
        if ([string]::IsNullOrWhiteSpace($Config.Email.SMTPServer)) {
            $errors += "SMTP Server is required"
        }
        
        if ($Config.Email.Port -le 0 -or $Config.Email.Port -gt 65535) {
            $errors += "SMTP Port must be between 1 and 65535"
        }
        
        if ([string]::IsNullOrWhiteSpace($Config.Email.From)) {
            $errors += "From email address is required"
        }
        elseif ($Config.Email.From -notmatch "^[^@]+@[^@]+\.[^@]+$") {
            $errors += "From email address format is invalid"
        }
        
        if (-not $Config.Email.To -or $Config.Email.To.Count -eq 0) {
            $errors += "At least one recipient email address is required"
        }
        else {
            foreach ($recipient in $Config.Email.To) {
                if ($recipient -notmatch "^[^@]+@[^@]+\.[^@]+$") {
                    $errors += "Recipient email address '$recipient' format is invalid"
                }
            }
        }
        
        # Validate authentication settings (optional)
        if (-not [string]::IsNullOrWhiteSpace($Config.Email.Username)) {
            # If username is provided, password should also be provided (but not required for anonymous SMTP)
            if ([string]::IsNullOrWhiteSpace($Config.Email.Password) -and 
                [string]::IsNullOrWhiteSpace($Config.Email.EncryptedPassword)) {
                Write-Warning "Username provided without password - will attempt anonymous authentication"
            }
        }
        
        if ($errors.Count -gt 0) {
            Write-Error "Email configuration validation failed:`n$($errors -join "`n")"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Error "Failed to validate email configuration: $($_.Exception.Message)"
        return $false
    }
}

# Send-ErrorNotification function moved to Logging.ps1 to avoid conflicts

# Functions are available when dot-sourced for testing
# Note: Export-ModuleMember is not needed when dot-sourcing