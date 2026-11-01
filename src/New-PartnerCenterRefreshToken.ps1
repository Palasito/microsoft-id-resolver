#Requires -Version 7.0
<#
.SYNOPSIS
    Obtains a Partner Center refresh token for the license price export workflow.

.DESCRIPTION
    Runs the OAuth 2.0 authorization code flow (with PKCE) once in the browser for the Entra app that the
    workflow uses. Sign in with a user that has the Admin Agent or Sales Agent role in Partner Center.

    The refresh token is never printed. It is copied to the clipboard, or written straight to the
    PARTNER_CENTER_REFRESH_TOKEN repository secret when -GitHubRepository is used (requires the GitHub CLI).

    The app registration needs a "Web" platform redirect URI that matches -RedirectUri exactly.

.PARAMETER TenantId
    Tenant ID of the CSP partner tenant (same value as the PARTNER_CENTER_TENANT_ID variable).

.PARAMETER ClientId
    Application (client) ID of the app registration (same value as the PARTNER_CENTER_CLIENT_ID variable).

.PARAMETER ClientSecret
    Client secret of the app registration. Prompted for when not supplied.

.PARAMETER RedirectUri
    Loopback redirect URI registered on the app. Defaults to http://localhost:8400/.

.PARAMETER GitHubRepository
    Optional owner/repo. When set, the token is stored with "gh secret set PARTNER_CENTER_REFRESH_TOKEN".

.EXAMPLE
    ./src/New-PartnerCenterRefreshToken.ps1 -TenantId <tenant-id> -ClientId <client-id>

    Prompts for the client secret, opens the browser and copies the refresh token to the clipboard.

.EXAMPLE
    ./src/New-PartnerCenterRefreshToken.ps1 -TenantId <tenant-id> -ClientId <client-id> -GitHubRepository Palasito/microsoft-id-resolver

    Stores the refresh token directly as the PARTNER_CENTER_REFRESH_TOKEN repository secret.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [securestring]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$RedirectUri = "http://localhost:8400/",

    [Parameter(Mandatory = $false)]
    [string]$GitHubRepository
)

$ErrorActionPreference = "Stop"

$scope = "https://api.partner.microsoft.com/user_impersonation offline_access"
$authority = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0"

function ConvertTo-Base64Url {
    param([byte[]]$Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-JwtClaims {
    param([string]$Token)
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    $payload = $payload.PadRight($payload.Length + (4 - $payload.Length % 4) % 4, '=')
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
}

if (-not $RedirectUri.EndsWith('/')) { $RedirectUri += '/' }
if (-not $ClientSecret) { $ClientSecret = Read-Host -Prompt "Client secret" -AsSecureString }
if ($GitHubRepository -and -not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "The GitHub CLI (gh) is required when -GitHubRepository is used."
}

$listener = [System.Net.HttpListener]::new()
try {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Partner Center Refresh Token - Starting" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan

    $verifier = ConvertTo-Base64Url ([Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
    $challenge = ConvertTo-Base64Url ([Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($verifier)))
    $state = [guid]::NewGuid().ToString('N')

    $query = @{
        client_id             = $ClientId
        response_type         = "code"
        redirect_uri          = $RedirectUri
        response_mode         = "query"
        scope                 = $scope
        state                 = $state
        code_challenge        = $challenge
        code_challenge_method = "S256"
        prompt                = "select_account"
    }
    $queryString = ($query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString($_.Value))" }) -join '&'
    $authorizeUrl = "$authority/authorize?$queryString"

    $listener.Prefixes.Add($RedirectUri)
    $listener.Start()

    Write-Host "`nStep 1: Sign in as an Admin Agent or Sales Agent user" -ForegroundColor Yellow
    Write-Host "If the browser does not open, browse to:`n$authorizeUrl"
    Start-Process $authorizeUrl

    $context = $listener.GetContext()
    $callback = [System.Web.HttpUtility]::ParseQueryString($context.Request.Url.Query)

    $message = "Sign-in complete. You can close this window and return to PowerShell."
    if ($callback["error"]) { $message = "Sign-in failed: $($callback["error"]). Check PowerShell for details." }
    $html = [Text.Encoding]::UTF8.GetBytes("<html><body style='font-family:sans-serif'><p>$([System.Net.WebUtility]::HtmlEncode($message))</p></body></html>")
    $context.Response.ContentType = "text/html; charset=utf-8"
    $context.Response.OutputStream.Write($html, 0, $html.Length)
    $context.Response.Close()

    if ($callback["error"]) {
        throw "Authorization failed: $($callback["error"]) - $($callback["error_description"])"
    }
    if ($callback["state"] -ne $state) {
        throw "Authorization response state does not match the request."
    }
    Write-Host "[OK] Received authorization code" -ForegroundColor Green

    Write-Host "`nStep 2: Redeeming authorization code" -ForegroundColor Yellow
    $tokens = Invoke-RestMethod -Method Post -Uri "$authority/token" -ContentType "application/x-www-form-urlencoded" -Body @{
        client_id     = $ClientId
        client_secret = (ConvertFrom-SecureString -SecureString $ClientSecret -AsPlainText)
        grant_type    = "authorization_code"
        code          = $callback["code"]
        redirect_uri  = $RedirectUri
        scope         = $scope
        code_verifier = $verifier
    }
    if (-not $tokens.refresh_token) {
        throw "No refresh token was returned. Make sure offline_access is allowed for the app."
    }

    $claims = Get-JwtClaims -Token $tokens.access_token
    Write-Host "[OK] Signed in as $($claims.upn) for $($claims.aud)" -ForegroundColor Green

    Write-Host "`nStep 3: Storing refresh token" -ForegroundColor Yellow
    if ($GitHubRepository) {
        $tokens.refresh_token | gh secret set PARTNER_CENTER_REFRESH_TOKEN --repo $GitHubRepository
        if ($LASTEXITCODE -ne 0) { throw "gh secret set failed with exit code $LASTEXITCODE." }
        Write-Host "[OK] Updated PARTNER_CENTER_REFRESH_TOKEN in $GitHubRepository" -ForegroundColor Green
    }
    else {
        Set-Clipboard -Value $tokens.refresh_token
        Write-Host "[OK] Refresh token copied to the clipboard. Paste it into the PARTNER_CENTER_REFRESH_TOKEN secret, then clear the clipboard." -ForegroundColor Green
    }

    Write-Host "`n==================================================" -ForegroundColor Cyan
    Write-Host "Partner Center Refresh Token - Complete" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
}
catch {
    Write-Error "Failed to obtain Partner Center refresh token: $_"
    exit 1
}
finally {
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
}
