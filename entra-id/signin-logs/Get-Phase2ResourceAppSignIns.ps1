<#
.SYNOPSIS
    Retrieves and summarizes sign-in logs for Azure Resource Manager resource application.

.DESCRIPTION
    This script imports the Summarise-ResourceAppSignIns module and retrieves sign-in logs
    for the Azure Resource Manager resource application:
    - 797f4846-ba00-4fd7-ba43-dac1f8f63013 (Azure Resource Manager)

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
    .\Get-AzureResourceManagerSignIns.ps1
    Retrieves interactive sign-in logs for Azure Resource Manager from the last 7 days

.EXAMPLE
    .\Get-AzureResourceManagerSignIns.ps1 -IncludeNonInteractive
    Retrieves both interactive and non-interactive sign-in logs using the beta endpoint

.EXAMPLE
    .\Get-AzureResourceManagerSignIns.ps1 -DaysBack 30 -EnableLogging
    Retrieves sign-in logs for the last 30 days with logging enabled

.EXAMPLE
    .\Get-AzureResourceManagerSignIns.ps1 -OutputPath "C:\Reports\AzureRM_SignIns.csv"
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

# Import the Summarise-ResourceAppSignIns module
$modulePath = Join-Path $PSScriptRoot "Summarise-ResourceAppSignIns.psm1"

if (-not (Test-Path $modulePath)) {
    Write-Error "Could not find Summarise-ResourceAppSignIns.psm1 in $PSScriptRoot"
    exit 1
}

Write-Host "Importing module from $modulePath" -ForegroundColor Cyan
Import-Module $modulePath -Force

# Define the Azure Resource Manager resource application ID
$appIds = @(
    "797f4846-ba00-4fd7-ba43-dac1f8f63013"   # Azure Resource Manager
)

Write-Host "`nRetrieving sign-in logs for the following resource application:" -ForegroundColor Cyan
Write-Host "  - 797f4846-ba00-4fd7-ba43-dac1f8f63013 (Azure Resource Manager)" -ForegroundColor Gray
Write-Host ""

# Build parameters for the Get-ResourceAppSignInSummary cmdlet
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
$results = Get-ResourceAppSignInSummary @params

# Return the results
return $results
