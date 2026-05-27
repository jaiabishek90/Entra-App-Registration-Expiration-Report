<#
.SYNOPSIS
    Fetches all Entra ID (Azure AD) app registrations and reports on expiring or expired
    client secrets and certificates. Uses app-only authentication with a certificate thumbprint.

.DESCRIPTION
    Connects to Microsoft Graph using a registered app + certificate (no interactive login),
    enumerates every application registration in the tenant, inspects all client secrets
    (passwordCredentials) and certificates (keyCredentials), calculates days remaining until
    expiry, and outputs a consolidated report. Optionally exports to CSV.

.PARAMETER DaysUntilExpiration
    Only return credentials expiring within this many days (or already expired). Default 30.

.PARAMETER ExportPath
    Optional. Full path to a CSV file to export results to.

.PARAMETER IncludeExpired
    Include already-expired credentials in the output. Default $true.

.EXAMPLE
    .\Check-EntraAppExpiringCredentials.ps1

.EXAMPLE
    .\Check-EntraAppExpiringCredentials.ps1 -DaysUntilExpiration 90 -ExportPath "C:\Temp\expiring.csv"

.NOTES
    Requires : Microsoft.Graph.Applications PowerShell module
    Auth     : App-only (client credentials) using a certificate installed in CurrentUser\My
               or LocalMachine\My with the thumbprint defined below.
    API perm : Application.Read.All (Application permission, admin-consented)
#>

[CmdletBinding()]
param(
    [int]$DaysUntilExpiration = 30,
    [string]$ExportPath,
    [bool]$IncludeExpired = $true
)

# ===========================================================================
# >>> HARDCODED APP-ONLY AUTHENTICATION SETTINGS - EDIT THESE <<<
# ===========================================================================
$TenantId              = '00000000-0000-0000-0000-000000000000'   # Your Entra tenant ID (GUID) or domain (contoso.onmicrosoft.com)
$ClientId              = '11111111-1111-1111-1111-111111111111'   # Application (client) ID of the registered app used to sign in
$CertificateThumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F90A1B2C3D4' # Thumbprint of the cert (no spaces, uppercase hex)
# ===========================================================================

# ---------------------------------------------------------------------------
# Ensure Microsoft.Graph.Applications module is available
# ---------------------------------------------------------------------------
$requiredModule = 'Microsoft.Graph.Applications'
if (-not (Get-Module -ListAvailable -Name $requiredModule)) {
    Write-Host "Module '$requiredModule' is not installed. Installing for current user..." -ForegroundColor Yellow
    try {
        Install-Module -Name $requiredModule -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to install module '$requiredModule'. Install manually: Install-Module Microsoft.Graph -Scope CurrentUser"
        return
    }
}
Import-Module $requiredModule -ErrorAction Stop

# ---------------------------------------------------------------------------
# Verify the certificate exists locally before attempting connect
# ---------------------------------------------------------------------------
$cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
if (-not $cert) {
    $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
}
if (-not $cert) {
    Write-Error "Certificate with thumbprint '$CertificateThumbprint' was not found in CurrentUser\My or LocalMachine\My."
    return
}
Write-Host "Found certificate: $($cert.Subject) (expires $($cert.NotAfter))" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph (app-only, certificate-based)
# ---------------------------------------------------------------------------
try {
    # Disconnect any prior session to avoid context conflicts
    if (Get-MgContext) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }

    Write-Host "Connecting to Microsoft Graph (app-only)..." -ForegroundColor Cyan
    Connect-MgGraph -TenantId $TenantId `
                    -ClientId $ClientId `
                    -CertificateThumbprint $CertificateThumbprint `
                    -NoWelcome -ErrorAction Stop

    $context = Get-MgContext
    Write-Host "Connected. Tenant: $($context.TenantId)  AuthType: $($context.AuthType)  App: $($context.ClientId)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    return
}

# ---------------------------------------------------------------------------
# Fetch all app registrations
# ---------------------------------------------------------------------------
Write-Host "Fetching all application registrations (this may take a while in large tenants)..." -ForegroundColor Cyan
try {
    $apps = Get-MgApplication -All -Property 'Id,DisplayName,AppId,PasswordCredentials,KeyCredentials,Notes' -ErrorAction Stop
}
catch {
    Write-Error "Failed to fetch applications: $_"
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

Write-Host "Found $($apps.Count) application registrations. Analyzing credentials..." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Build report
# ---------------------------------------------------------------------------
$now = Get-Date
$results = New-Object System.Collections.Generic.List[object]

foreach ($app in $apps) {

    # ----- Client secrets -----
    foreach ($secret in $app.PasswordCredentials) {
        if (-not $secret.EndDateTime) { continue }

        $endDate  = [datetime]$secret.EndDateTime
        $daysLeft = [int]([math]::Floor(($endDate - $now).TotalDays))

        $status = if ($daysLeft -lt 0) { 'Expired' }
                  elseif ($daysLeft -le $DaysUntilExpiration) { 'Expiring Soon' }
                  else { 'Valid' }

        if ($status -eq 'Valid') { continue }
        if ($status -eq 'Expired' -and -not $IncludeExpired) { continue }

        $results.Add([pscustomobject]@{
            AppDisplayName  = $app.DisplayName
            AppId           = $app.AppId
            ObjectId        = $app.Id
            CredentialType  = 'ClientSecret'
            CredentialName  = $secret.DisplayName
            KeyId           = $secret.KeyId
            StartDate       = $secret.StartDateTime
            EndDate         = $endDate
            DaysUntilExpiry = $daysLeft
            Status          = $status
            Notes           = $app.Notes
        })
    }

    # ----- Certificates (keyCredentials) -----
    foreach ($keyCred in $app.KeyCredentials) {
        if (-not $keyCred.EndDateTime) { continue }

        $endDate  = [datetime]$keyCred.EndDateTime
        $daysLeft = [int]([math]::Floor(($endDate - $now).TotalDays))

        $status = if ($daysLeft -lt 0) { 'Expired' }
                  elseif ($daysLeft -le $DaysUntilExpiration) { 'Expiring Soon' }
                  else { 'Valid' }

        if ($status -eq 'Valid') { continue }
        if ($status -eq 'Expired' -and -not $IncludeExpired) { continue }

        $results.Add([pscustomobject]@{
            AppDisplayName  = $app.DisplayName
            AppId           = $app.AppId
            ObjectId        = $app.Id
            CredentialType  = "Certificate ($($keyCred.Type))"
            CredentialName  = $keyCred.DisplayName
            KeyId           = $keyCred.KeyId
            StartDate       = $keyCred.StartDateTime
            EndDate         = $endDate
            DaysUntilExpiry = $daysLeft
            Status          = $status
            Notes           = $app.Notes
        })
    }
}

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
if ($results.Count -eq 0) {
    Write-Host "`nNo credentials found that are expired or expiring within $DaysUntilExpiration day(s)." -ForegroundColor Green
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    return
}

$sorted = $results | Sort-Object DaysUntilExpiry

Write-Host "`n===== Summary =====" -ForegroundColor Cyan
Write-Host ("Expired       : {0}" -f ($sorted | Where-Object Status -eq 'Expired').Count) -ForegroundColor Red
Write-Host ("Expiring Soon : {0}" -f ($sorted | Where-Object Status -eq 'Expiring Soon').Count) -ForegroundColor Yellow
Write-Host ("Total flagged : {0}`n" -f $sorted.Count) -ForegroundColor White

$sorted | Format-Table AppDisplayName, CredentialType, CredentialName, EndDate, DaysUntilExpiry, Status -AutoSize

if ($ExportPath) {
    try {
        $sorted | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $ExportPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export CSV: $_"
    }
}

# Clean up session
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

# Return objects so they can be piped further
return $sorted
