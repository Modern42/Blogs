<#
.SYNOPSIS
    Retrieves and summarizes sign-in logs for specific Office 365 client applications.

.DESCRIPTION
    This script imports the Summarise-ClientAppSignIns module and retrieves sign-in logs
    for the following Azure client applications:
    - 04b07795-8ddb-461a-bbee-02f9e1bf7b46 (Azure command-line interface (Azure CLI))
    - 1950a258-227b-4e31-a9cf-717495945fc2 (Azure PowerShell)
    - 0c1307d4-29d6-4389-a11c-5cbe7f65d7fa (Azure mobile app)

.PARAMETER IncludeNonInteractive
    Switch parameter to include non-interactive sign-in logs. Disabled by default.
    When enabled, uses the Microsoft Graph beta endpoint.

.PARAMETER DaysBack
    Number of days to look back for sign-in logs. Defaults to 7 days.
    Maximum value is 30 days (Entra ID log retention limit).

.PARAMETER EnableLogging
    Switch parameter to enable logging to file. Disabled by default.

.PARAMETER OutputPath
    Optional path for the output CSV file. If not specified, defaults to script directory
    with a timestamped filename.

.EXAMPLE
    .\Get-Phase2ClientAppSignIns.ps1
    Retrieves interactive sign-in logs for the specified Azure applications from the last 7 days

.EXAMPLE
    .\Get-Phase2ClientAppSignIns.ps1 -IncludeNonInteractive
    Retrieves both interactive and non-interactive sign-in logs using the beta endpoint

.EXAMPLE
    .\Get-Phase2ClientAppSignIns.ps1 -DaysBack 30 -EnableLogging
    Retrieves sign-in logs for the last 30 days with logging enabled

.EXAMPLE
    .\Get-Phase2ClientAppSignIns.ps1 -OutputPath "C:\Reports\Phase2_SignIns.csv"
    Retrieves sign-in logs and saves them to the specified output path

.NOTES
    Requires Microsoft.Graph PowerShell module and appropriate permissions.
    Required Graph API permissions: AuditLog.Read.All or Directory.Read.All
    Requires Entra ID P1 or P2 license
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [switch]$IncludeNonInteractive,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 30)]
    [int]$DaysBack = 7,

    [Parameter(Mandatory=$false)]
    [switch]$EnableLogging,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath
)

# Import the Summarise-ClientAppSignIns module
$modulePath = Join-Path $PSScriptRoot "Summarise-ClientAppSignIns.psm1"

if (-not (Test-Path $modulePath)) {
    Write-Error "Could not find Summarise-ClientAppSignIns.psm1 in $PSScriptRoot"
    exit 1
}

Write-Host "Importing module from $modulePath" -ForegroundColor Cyan
Import-Module $modulePath -Force

# Define the Azure client application IDs
$appIds = @(
    "04b07795-8ddb-461a-bbee-02f9e1bf7b46",  # Azure command-line interface (Azure CLI)
    "1950a258-227b-4e31-a9cf-717495945fc2",  # Azure PowerShell
    "0c1307d4-29d6-4389-a11c-5cbe7f65d7fa"   # Azure mobile app
)

Write-Host "`nRetrieving sign-in logs for the following Azure applications:" -ForegroundColor Cyan
Write-Host "  - 04b07795-8ddb-461a-bbee-02f9e1bf7b46 (Azure command-line interface (Azure CLI))" -ForegroundColor Gray
Write-Host "  - 1950a258-227b-4e31-a9cf-717495945fc2 (Azure PowerShell)" -ForegroundColor Gray
Write-Host "  - 0c1307d4-29d6-4389-a11c-5cbe7f65d7fa (Azure mobile app)" -ForegroundColor Gray
Write-Host ""

# Build parameters for the Get-ClientAppSignInSummary cmdlet
$params = @{
    AppId = $appIds
    DaysBack = $DaysBack
    SuccessfulOnly = $true
}

# Add optional parameters if specified
if ($IncludeNonInteractive) {
    $params.IncludeNonInteractive = $true
}

if ($EnableLogging) {
    $params.EnableLogging = $true
}

if ($OutputPath) {
    $params.OutputPath = $OutputPath
}

# Call the cmdlet with splatting
$results = Get-ClientAppSignInSummary @params

# Return the results
return $results
