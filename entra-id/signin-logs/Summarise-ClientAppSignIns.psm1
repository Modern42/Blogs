<#
.SYNOPSIS
    Summarizes sign-in counts per user per application for specific client applications.

.DESCRIPTION
    This function retrieves sign-in logs from Microsoft Entra ID (Azure AD) for specific client
    applications and summarizes the count of sign-ins per user per app. Results are exported
    to a CSV file for analysis.

.PARAMETER AppId
    One or more application/client IDs (GUIDs) to filter sign-in logs by.
    Can be a single AppId or an array of AppIds.

.PARAMETER DaysBack
    Number of days to look back for sign-in logs. Defaults to 7 days.
    Maximum value is 30 days (Entra ID log retention limit).

.PARAMETER StartDate
    Optional start date for the sign-in log query (DateTime format).
    If specified, DaysBack parameter is ignored.

.PARAMETER EndDate
    Optional end date for the sign-in log query (DateTime format).
    Defaults to current date/time if StartDate is specified.

.PARAMETER OutputPath
    Optional path for the output CSV file. Defaults to script directory.

.PARAMETER LogPath
    Optional path for the log file. Defaults to script directory.

.PARAMETER EnableLogging
    Switch parameter to enable logging to file. Logging is disabled by default.

.PARAMETER IncludeNonInteractive
    Switch parameter to include non-interactive sign-in logs in addition to interactive sign-ins.

.PARAMETER SuccessfulOnly
    Switch parameter to filter for successful sign-ins only (errorCode eq 0).
    When enabled, only sign-ins with successful authentication will be included.

.EXAMPLE
    Import-Module .\Summarise-ClientAppSignIns.psm1
    Get-ClientAppSignInSummary -AppId "12345678-1234-1234-1234-123456789abc"
    Retrieves and summarizes interactive sign-in logs for the specified client application from the last 7 days

.EXAMPLE
    Get-ClientAppSignInSummary -AppId "12345678-1234-1234-1234-123456789abc" -DaysBack 30 -EnableLogging
    Retrieves and summarizes sign-in logs for the last 30 days with logging enabled

.EXAMPLE
    Get-ClientAppSignInSummary -AppId @("12345678-1234-1234-1234-123456789abc", "87654321-4321-4321-4321-cba987654321") -StartDate "2025-01-01" -EndDate "2025-01-31"
    Retrieves and summarizes sign-in logs for multiple client applications for a specific date range

.EXAMPLE
    Get-ClientAppSignInSummary -AppId "12345678-1234-1234-1234-123456789abc" -IncludeNonInteractive
    Retrieves and summarizes both interactive and non-interactive sign-in logs

.EXAMPLE
    $results = Get-ClientAppSignInSummary -AppId "12345678-1234-1234-1234-123456789abc" -DaysBack 14
    $results | Where-Object { $_.SignInCount -gt 10 }
    Store results in a variable and filter for users with more than 10 sign-ins

.NOTES
    Requires Microsoft.Graph PowerShell module and appropriate permissions.
    Required Graph API permissions: AuditLog.Read.All or Directory.Read.All
    Requires Entra ID P1 or P2 license
#>
function Get-ClientAppSignInSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({
            foreach ($id in $_) {
                if ($id -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                    throw "Invalid AppId format: $id. Must be a valid GUID."
                }
            }
            return $true
        })]
        [string[]]$AppId,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1, 30)]
        [int]$DaysBack = 7,

        [Parameter(Mandatory=$false)]
        [DateTime]$StartDate,

        [Parameter(Mandatory=$false)]
        [DateTime]$EndDate,

        [Parameter(Mandatory=$false)]
        [string]$OutputPath = "$PSScriptRoot\ClientAppSignInSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

        [Parameter(Mandatory=$false)]
        [string]$LogPath = "$PSScriptRoot\ClientAppSignInSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",

        [Parameter(Mandatory=$false)]
        [switch]$EnableLogging,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeNonInteractive,

        [Parameter(Mandatory=$false)]
        [switch]$SuccessfulOnly
    )

    # Function to write to log file and console
    function Write-Log {
        param([string]$Message, [string]$Level = "INFO")

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"

        # Write to console with color
        switch ($Level) {
            "ERROR" { Write-Host $logMessage -ForegroundColor Red }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
            "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
            default { Write-Host $logMessage }
        }

        # Write to log file only if logging is enabled
        if ($EnableLogging) {
            Add-Content -Path $LogPath -Value $logMessage
        }
    }

    Write-Log "Starting sign-in log summarization for client application(s): $($AppId -join ', ')"

    # Always use beta endpoint
    $apiVersion = "beta"
    Write-Log "Using beta endpoint"

    # Calculate date range
    if ($StartDate) {
        if (-not $EndDate) {
            $EndDate = Get-Date
        }
        Write-Log "Using custom date range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
    }
    else {
        $EndDate = Get-Date
        $StartDate = $EndDate.AddDays(-$DaysBack)
        Write-Log "Using date range: Last $DaysBack days ($($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd')))"
    }

    # Validate date range
    if ($StartDate -gt $EndDate) {
        Write-Log "Start date cannot be after end date" "ERROR"
        throw "Start date cannot be after end date"
    }

    # Format dates for Graph API (ISO 8601)
    $startDateFormatted = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $endDateFormatted = $EndDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    # Check if Microsoft Graph module is available
    try {
        Import-Module Microsoft.Graph.Reports -ErrorAction Stop
        Write-Log "Microsoft.Graph.Reports module loaded successfully"
    }
    catch {
        Write-Log "Failed to load Microsoft.Graph.Reports module. Please install it using: Install-Module Microsoft.Graph.Reports" "ERROR"
        throw
    }

    # Connect to Microsoft Graph
    try {
        Write-Log "Connecting to Microsoft Graph..."
        $context = Get-MgContext

        if (-not $context) {
            Connect-MgGraph -Scopes "AuditLog.Read.All" -ErrorAction Stop
            Write-Log "Successfully connected to Microsoft Graph"
        }
        else {
            Write-Log "Already connected to Microsoft Graph as $($context.Account)"

            # Verify required permissions
            $requiredScopes = @("AuditLog.Read.All", "Directory.Read.All")
            $currentScopes = $context.Scopes

            $hasRequiredPermission = $false
            foreach ($scope in $requiredScopes) {
                if ($scope -in $currentScopes) {
                    $hasRequiredPermission = $true
                    break
                }
            }

            if (-not $hasRequiredPermission) {
                Write-Log "Missing required permissions. Reconnecting..." "WARNING"
                Connect-MgGraph -Scopes "AuditLog.Read.All" -ErrorAction Stop
            }
        }
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" "ERROR"
        throw
    }

    # Build filter query
    # For multiple AppIds, create an OR condition
    $appIdFilter = if ($AppId.Count -eq 1) {
        "appId eq '$($AppId[0])'"
    }
    else {
        $appIdConditions = $AppId | ForEach-Object { "appId eq '$_'" }
        "(" + ($appIdConditions -join ' or ') + ")"
    }

    # Build the base filter with date range and app ID
    $filter = "(createdDateTime ge $startDateFormatted and createdDateTime le $endDateFormatted) and $appIdFilter"

    # Add signInEventTypes filter based on IncludeNonInteractive parameter
    if ($IncludeNonInteractive) {
        $filter += " and (signInEventTypes/any(t: t eq 'interactiveUser') or signInEventTypes/any(t: t eq 'nonInteractiveUser'))"
        Write-Log "Added signInEventTypes filter for both interactive and non-interactive sign-ins"
    } else {
        $filter += " and signInEventTypes/any(t: t eq 'interactiveUser')"
        Write-Log "Added signInEventTypes filter for interactive sign-ins only"
    }

    # Add filter for successful sign-ins only if specified
    if ($SuccessfulOnly) {
        $filter += " and status/errorCode eq 0"
        Write-Log "Added filter for successful sign-ins only (errorCode eq 0)"
    }

    Write-Log "Filter query: $filter"

    # Retrieve sign-in logs
    try {
        Write-Log "Retrieving sign-in logs..."

        # Use Get-MgAuditLogSignIn with filter
        $signInLogs = @()
        $pageCount = 0

        # Get all sign-in logs using pagination
        $uri = "https://graph.microsoft.com/$apiVersion/auditLogs/signIns?`$filter=$filter&`$top=999"

        do {
            $pageCount++
            Write-Log "Fetching page $pageCount..."

            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

            if ($response.value) {
                $signInLogs += $response.value
                Write-Log "Retrieved $($response.value.Count) records (Total so far: $($signInLogs.Count))"
            }

            # Get next page URL if available
            $uri = $response.'@odata.nextLink'

        } while ($uri)

        Write-Log "Successfully retrieved $($signInLogs.Count) sign-in log entries" "SUCCESS"
    }
    catch {
        Write-Log "Failed to retrieve sign-in logs: $_" "ERROR"
        Disconnect-MgGraph
        throw
    }

    # Check if any logs were found
    if ($signInLogs.Count -eq 0) {
        Write-Log "No sign-in logs found for the specified client application(s) and date range" "WARNING"
        Write-Log "This could mean:" "WARNING"
        Write-Log "  - No sign-ins occurred for this client application during the specified period" "WARNING"
        Write-Log "  - The application ID(s) may be incorrect" "WARNING"
        Write-Log "  - Your account may not have access to these logs" "WARNING"
        Disconnect-MgGraph
        return
    }

    # Summarize sign-ins per user per app
    Write-Log "Summarizing sign-in counts per user per application..."

    $summary = $signInLogs | Group-Object -Property @{Expression={$_.userPrincipalName}}, @{Expression={$_.appId}} | ForEach-Object {
        $firstLog = $_.Group[0]
        $group = $_.Group

        # Collect unique 'set' values for output
        $authReqSet = [System.Collections.Generic.HashSet[string]]::new()
        $authProtocolSet = [System.Collections.Generic.HashSet[string]]::new()
        $operatingSystemSet = [System.Collections.Generic.HashSet[string]]::new()
        $browserSet = [System.Collections.Generic.HashSet[string]]::new()

        foreach ($log in $group) {
            if ($log.authenticationRequirement) { [void]$authReqSet.Add($log.authenticationRequirement) }
            if ($log.authenticationProtocol) { [void]$authProtocolSet.Add($log.authenticationProtocol) }
            if ($log.deviceDetail.operatingSystem) { [void]$operatingSystemSet.Add($log.deviceDetail.operatingSystem) }
            if ($log.deviceDetail.browser) { [void]$browserSet.Add($log.deviceDetail.browser) }
        }

        [PSCustomObject]@{
            UserPrincipalName          = $firstLog.userPrincipalName
            UserDisplayName            = $firstLog.userDisplayName
            UserId                     = $firstLog.userId
            AppDisplayName             = $firstLog.appDisplayName
            AppId                      = $firstLog.appId
            SignInCount                = $_.Count
            AuthenticationRequirements = ($authReqSet | Sort-Object) -join ', '
            AuthenticationProtocols    = ($authProtocolSet | Sort-Object) -join ', '
            OperatingSystems           = ($operatingSystemSet | Sort-Object) -join ', '
            Browsers                   = ($browserSet | Sort-Object) -join ', '
        }
    } | Sort-Object AppDisplayName, UserPrincipalName

    # Export to CSV
    try {
        $summary | Export-Csv -Path $OutputPath -NoTypeInformation -ErrorAction Stop
        Write-Log "Successfully exported summary to: $OutputPath" "SUCCESS"
    }
    catch {
        Write-Log "Failed to export CSV: $_" "ERROR"
        Disconnect-MgGraph
        throw
    }

    # Display summary statistics
    Write-Log "=== SUMMARY ===" "INFO"
    Write-Log "Client Application ID(s): $($AppId -join ', ')" "INFO"
    Write-Log "Date Range: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))" "INFO"
    Write-Log "Total Sign-In Logs Retrieved: $($signInLogs.Count)" "INFO"
    Write-Log "Unique User/App Combinations: $($summary.Count)" "INFO"
    Write-Log "Unique Users: $(($summary | Select-Object -Unique UserPrincipalName).Count)" "INFO"
    Write-Log "Unique Apps: $(($summary | Select-Object -Unique AppId).Count)" "INFO"
    Write-Log "Output file: $OutputPath" "INFO"
    if ($EnableLogging) {
        Write-Log "Log file: $LogPath" "INFO"
    }

    # Display summary table in console
    Write-Host "`n=== SIGN-IN SUMMARY BY USER AND APP ===" -ForegroundColor Cyan
    $summary | Format-Table -AutoSize

    # Disconnect from Microsoft Graph
    Disconnect-MgGraph
    Write-Log "Disconnected from Microsoft Graph"

    # Return the summary object
    return $summary
}

# Export the function
Export-ModuleMember -Function Get-ClientAppSignInSummary
