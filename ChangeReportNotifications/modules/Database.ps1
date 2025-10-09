# Database Module for Change Report Notifications
# Handles SQL Server connectivity and database operations

# Import required modules
Add-Type -AssemblyName System.Data

<#
.SYNOPSIS
    Establishes a connection to SQL Server database
.DESCRIPTION
    Creates a secure database connection using either Windows Authentication or SQL Server Authentication
.PARAMETER Config
    Configuration object containing database connection parameters
.RETURNS
    SqlConnection object if successful, $null if failed
.EXAMPLE
    $connection = Connect-Database -Config $config
#>
function Connect-Database {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        Write-Verbose "Attempting to connect to database: $($Config.Database.Server)"
        
        # Create connection object
        $connection = New-Object System.Data.SqlClient.SqlConnection
        
        # Build connection string based on authentication type
        if ($Config.Database.AuthType -eq "Windows") {
            $connectionString = "Server=$($Config.Database.Server);Database=$($Config.Database.Database);Integrated Security=True;Connection Timeout=30;Encrypt=True;TrustServerCertificate=False;"
            $connection.ConnectionString = $connectionString
        }
        elseif ($Config.Database.AuthType -eq "SQL") {
            if (-not $Config.Database.Credential) {
                throw "Database credential object is required for SQL Server Authentication"
            }
            
            # Use secure connection string without embedded password
            $connectionString = "Server=$($Config.Database.Server);Database=$($Config.Database.Database);Connection Timeout=30;Encrypt=True;TrustServerCertificate=False;"
            $connection.ConnectionString = $connectionString
            
            # Set credentials securely using SqlCredential
            # Make SecureString read-only as required by SqlCredential
            $securePassword = $Config.Database.Credential.Password.Copy()
            $securePassword.MakeReadOnly()
            $connection.Credential = New-Object System.Data.SqlClient.SqlCredential($Config.Database.Credential.UserName, $securePassword)
        }
        else {
            throw "Invalid AuthType. Must be 'Windows' or 'SQL'"
        }
        
        # Open connection
        $connection.Open()
        
        Write-Verbose "Successfully connected to database"
        return $connection
    }
    catch {
        Write-Error "Failed to connect to database: $($_.Exception.Message)"
        if ($connection) {
            try { $connection.Dispose() } catch { }
        }
        return $null
    }
}

<#
.SYNOPSIS
    Tests database connectivity without maintaining the connection
.DESCRIPTION
    Validates that a database connection can be established with the provided configuration
.PARAMETER Config
    Configuration object containing database connection parameters
.RETURNS
    $true if connection successful, $false otherwise
.EXAMPLE
    $isConnected = Test-DatabaseConnection -Config $config
#>
function Test-DatabaseConnection {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )
    
    try {
        Write-Verbose "Testing database connection to: $($Config.Database.Server)"
        
        $connection = Connect-Database -Config $Config
        
        if ($connection -and $connection.State -eq 'Open') {
            Close-Database -Connection $connection
            Write-Verbose "Database connection test successful"
            return $true
        }
        else {
            Write-Warning "Database connection test failed"
            return $false
        }
    }
    catch {
        Write-Error "Database connection test failed: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
    Properly closes a database connection
.DESCRIPTION
    Safely closes and disposes of a SQL Server connection object
.PARAMETER Connection
    SqlConnection object to close
.EXAMPLE
    Close-Database -Connection $connection
#>
function Close-Database {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection
    )
    
    try {
        if ($Connection -and $Connection.State -ne 'Closed') {
            Write-Verbose "Closing database connection"
            $Connection.Close()
            $Connection.Dispose()
            Write-Verbose "Database connection closed successfully"
        }
    }
    catch {
        Write-Error "Error closing database connection: $($_.Exception.Message)"
    }
}

# Define ChangeRecord class for structured data handling
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
    
    # Constructor
    ChangeRecord() {}
    
    # Constructor with parameters
    ChangeRecord([hashtable]$Properties) {
        foreach ($key in $Properties.Keys) {
            if ($this.PSObject.Properties.Name -contains $key) {
                $this.$key = $Properties[$key]
            }
        }
    }
    
    # String representation
    [string] ToString() {
        return "$($this.ChangeID) - $($this.Priority) - $($this.ShortDescription)"
    }
}

<#
.SYNOPSIS
    Retrieves critical and high priority change records from the database
.DESCRIPTION
    Executes a parameterized query to fetch change records with Critical or High priority
    for a specified date, with proper sorting by priority and start date
.PARAMETER Connection
    Active SqlConnection object
.PARAMETER ReportDate
    Date to filter changes by ActualStartDate (defaults to current date)
.RETURNS
    Array of ChangeRecord objects
.EXAMPLE
    $changes = Get-CriticalChanges -Connection $connection -ReportDate (Get-Date)
#>
function Get-CriticalChanges {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SqlClient.SqlConnection]$Connection,
        
        [Parameter(Mandatory = $false)]
        [datetime]$ReportDate = (Get-Date)
    )
    
    $changes = @()
    
    try {
        Write-Verbose "Querying for critical and high priority changes for date: $($ReportDate.ToString('yyyy-MM-dd'))"
        
        # Parameterized SQL query to prevent SQL injection
        $sqlQuery = @"
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
"@
        
        # Create SQL command with parameter
        $command = New-Object System.Data.SqlClient.SqlCommand($sqlQuery, $Connection)
        $command.CommandTimeout = 30
        
        # Add parameter to prevent SQL injection
        $dateParam = $command.Parameters.Add("@ReportDate", [System.Data.SqlDbType]::Date)
        $dateParam.Value = $ReportDate.Date
        
        # Execute query
        $reader = $command.ExecuteReader()
        
        # Process results
        while ($reader.Read()) {
            $changeRecord = [ChangeRecord]::new()
            
            # Safely read values with null checking
            $changeRecord.ChangeID = if ($reader["ChangeID"] -ne [DBNull]::Value) { $reader["ChangeID"].ToString() } else { "" }
            $changeRecord.Priority = if ($reader["Priority"] -ne [DBNull]::Value) { $reader["Priority"].ToString() } else { "" }
            $changeRecord.Type = if ($reader["Type"] -ne [DBNull]::Value) { $reader["Type"].ToString() } else { "" }
            $changeRecord.ConfigurationItem = if ($reader["ConfigurationItem"] -ne [DBNull]::Value) { $reader["ConfigurationItem"].ToString() } else { "" }
            $changeRecord.ShortDescription = if ($reader["ShortDescription"] -ne [DBNull]::Value) { $reader["ShortDescription"].ToString() } else { "" }
            $changeRecord.AssignmentsGroup = if ($reader["AssignmentsGroup"] -ne [DBNull]::Value) { $reader["AssignmentsGroup"].ToString() } else { "" }
            $changeRecord.AssignedTo = if ($reader["AssignedTo"] -ne [DBNull]::Value) { $reader["AssignedTo"].ToString() } else { "" }
            
            # Handle datetime fields with null checking
            if ($reader["ActualStartDate"] -ne [DBNull]::Value) {
                $changeRecord.ActualStartDate = [datetime]$reader["ActualStartDate"]
            } else {
                $changeRecord.ActualStartDate = [datetime]::MinValue
            }
            
            if ($reader["ActualEndDate"] -ne [DBNull]::Value) {
                $changeRecord.ActualEndDate = [datetime]$reader["ActualEndDate"]
            } else {
                $changeRecord.ActualEndDate = [datetime]::MinValue
            }
            
            $changes += $changeRecord
        }
        
        $reader.Close()
        
        Write-Verbose "Retrieved $($changes.Count) change records"
        return $changes
    }
    catch {
        Write-Error "Failed to retrieve change records: $($_.Exception.Message)"
        if ($reader -and -not $reader.IsClosed) {
            $reader.Close()
        }
        return @()
    }
}

# Functions are available when dot-sourced
# Note: Export-ModuleMember is not needed when dot-sourcing