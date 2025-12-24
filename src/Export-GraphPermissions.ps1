<#
        .SYNOPSIS
        Creates a list of Microsoft Graph permission roles.

        .DESCRIPTION

        .EXAMPLE
        ./src/Export-GraphPermissions.ps1

        Creates a list of all Graph permissions with output written to .\_info

        .EXAMPLE
        Export-GraphPermissions.ps1 $OutputPath ".\myOutputFolder"

        Creates a list using custom folders for the output
#>

param (
    [Parameter(Mandatory = $false, HelpMessage = "Path to output the csv and json files")]
    [string]$OutputPath = ".\_info"
)

function GetPermissionsFromMicrosoftGraph() {
    Write-Host "Retrieving permissions from Microsoft Graph"
    $graphApps = Import-Csv ".\_info\MicrosoftApps.csv" | where-object { $_.source -eq "Graph" -or $_.source -eq "GitHub" }
    # $graphAppId = @("00000003-0000-0000-c000-000000000000", "c5393580-f805-4401-95e8-94b7a6ef2fc2", "00000002-0000-0000-c000-000000000000")

    $i = 0
    $spFinal = @()
    foreach ($app in $graphApps) {
        try {
            $i++
            $appId = $app.appId
            $appName = $app.displayName
            Write-Host "Processing App: $($i) of $($graphApps.Count)"
            $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -All
            if ($sp) {
                $spFinal += $sp
            }
        }
        catch {
            Write-Host "Error retrieving service principal for App: $appName. Error: $_"
        }
    }
    # $sp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -All

    return $spFinal
}

$sp = GetPermissionsFromMicrosoftGraph

Write-Host "Exporting to csv and json"
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$outputFilePathCsv = Join-Path $OutputPath "GraphAppRoles.csv"
$outputFilePathJson = Join-Path $OutputPath "GraphAppRoles.json"

$sp.AppRoles | Select-Object Id, Value, DisplayName, Description | Export-Csv $outputFilePathCsv
$sp.AppRoles | ConvertTo-Json | Out-File $outputFilePathJson

$outputFilePathCsv = Join-Path $OutputPath "GraphDelegateRoles.csv"
$outputFilePathJson = Join-Path $OutputPath "GraphDelegateRoles.json"

$sp.Oauth2PermissionScopes | Select-Object Id, Value, AdminConsentDisplayName, AdminConsentDescription | Export-Csv $outputFilePathCsv
$sp.Oauth2PermissionScopes | ConvertTo-Json | Out-File $outputFilePathJson

# if (!$sp.Oauth2PermissionScopes) {
#     $sp.publishedPermissionScopes | Select-Object Id, Value, AdminConsentDisplayName, AdminConsentDescription | Export-Csv $outputFilePathCsv
#     $sp.publishedPermissionScopes | ConvertTo-Json | Out-File $outputFilePathJson
# }
# else {
#     $sp.Oauth2PermissionScopes | Select-Object Id, Value, AdminConsentDisplayName, AdminConsentDescription | Export-Csv $outputFilePathCsv
#     $sp.Oauth2PermissionScopes | ConvertTo-Json | Out-File $outputFilePathJson
# }