# Check-EntraAppExpiringCredentials

A PowerShell script that audits **all Entra ID (Azure AD) app registrations** in a tenant and reports any **client secrets** or **certificates** that are expired or about to expire. Authenticates to Microsoft Graph **app-only** using a certificate thumbprint — no interactive sign-in required, making it suitable for scheduled runs.

---

## Features

- Enumerates every app registration in the tenant (handles paging automatically)
- Inspects both client secrets (`passwordCredentials`) and certificates (`keyCredentials`)
- Calculates days remaining and flags each item as `Expired`, `Expiring Soon`, or `Valid`
- Configurable expiry threshold (default 30 days)
- Console summary plus a sorted, formatted table
- Optional CSV export
- Pre-flight checks for the required module and local certificate
- Clean session teardown after each run

---

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or later (PowerShell 7+ recommended) |
| Module | `Microsoft.Graph.Applications` (auto-installed if missing) |
| Entra app registration | Used for authentication (see setup below) |
| API permission | `Application.Read.All` (Microsoft Graph, **Application** permission, admin-consented) |
| Certificate | Uploaded to the app registration **and** installed locally in `CurrentUser\My` or `LocalMachine\My` with its private key |

---

## One-time setup

### 1. Create the app registration used for authentication

1. In the Entra admin center go to **Identity → Applications → App registrations → New registration**.
2. Give it a name (e.g. `ps-credential-auditor`), leave the redirect URI blank, and register.
3. Note the **Application (client) ID** and **Directory (tenant) ID** — you'll need both.

### 2. Grant the Graph permission

1. Open the new app → **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Add **`Application.Read.All`**.
3. Click **Grant admin consent for \<your tenant\>**.

### 3. Generate / obtain a certificate

Generate a self-signed certificate on the machine that will run the script:

```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=ps-credential-auditor" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export the public key (.cer) — this is what you upload to Entra
Export-Certificate -Cert $cert -FilePath ".\ps-credential-auditor.cer" | Out-Null

$cert.Thumbprint   # copy this value into the script
```

### 4. Upload the public key to the app registration

1. App registration → **Certificates & secrets → Certificates → Upload certificate**.
2. Upload the `.cer` file produced above.
3. Confirm the thumbprint shown in Entra matches `$cert.Thumbprint` from the previous step.

### 5. Fill in the hardcoded values in the script

Open `Check-EntraAppExpiringCredentials.ps1` and edit:

```powershell
$TenantId              = '<your-tenant-id-or-domain>'
$ClientId              = '<application-client-id-of-the-auditor-app>'
$CertificateThumbprint = '<thumbprint-uppercase-no-spaces>'
```

---

## Usage

```powershell
# Default — items expired or expiring in the next 30 days
.\Check-EntraAppExpiringCredentials.ps1

# Wider window
.\Check-EntraAppExpiringCredentials.ps1 -DaysUntilExpiration 90

# Export results to CSV
.\Check-EntraAppExpiringCredentials.ps1 -DaysUntilExpiration 60 -ExportPath "C:\Temp\expiring.csv"

# Hide already-expired entries (only show upcoming)
.\Check-EntraAppExpiringCredentials.ps1 -IncludeExpired $false
```

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `DaysUntilExpiration` | `int` | `30` | Window (in days) for flagging items as *Expiring Soon* |
| `ExportPath` | `string` | — | Path to write a CSV report |
| `IncludeExpired` | `bool` | `$true` | Whether already-expired items appear in the output |

### Output columns

`AppDisplayName`, `AppId`, `ObjectId`, `CredentialType`, `CredentialName`, `KeyId`, `StartDate`, `EndDate`, `DaysUntilExpiry`, `Status`, `Notes`

---

## Scheduling

To run unattended (e.g. weekly via Task Scheduler), install the certificate in `LocalMachine\My` and grant the service account that runs the task **read** access to the private key:

1. Run `certlm.msc` → **Personal → Certificates**.
2. Right-click the cert → **All Tasks → Manage Private Keys…**.
3. Add the service account with **Read** permission only.

Example scheduled-task action:

```text
Program:   powershell.exe
Arguments: -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Check-EntraAppExpiringCredentials.ps1" -DaysUntilExpiration 60 -ExportPath "C:\Reports\entra-expiring.csv"
```

---

## Security notes

- The thumbprint in the script is **not a secret**, but it does tell anyone reading the file which certificate to look for. The actual security relies on the private key staying protected.
- Prefer `LocalMachine\My` with a tight ACL on the private key for shared/server machines.
- Don't commit a script with real `TenantId` / `ClientId` values to a public repo. For shared environments, move the three hardcoded values into environment variables or an external config file outside source control.
- Set a calendar reminder for the auditor app's **own** certificate — ironically, this script can audit itself once it's running.
- Grant only `Application.Read.All`. The script never writes to Graph, so write permissions aren't needed.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `Certificate with thumbprint '…' was not found` | Cert isn't installed locally, or the thumbprint string has spaces / wrong case. Run `Get-ChildItem Cert:\CurrentUser\My` to confirm. |
| `AADSTS700027: Client assertion contains an invalid signature` | The certificate uploaded to Entra doesn't match the local one. Re-upload the correct `.cer`. |
| `Insufficient privileges to complete the operation` | `Application.Read.All` is missing or admin consent wasn't granted. |
| `Get-MgApplication` returns zero apps | Probably signed in with delegated/user context from a prior session. The script disconnects first, but if you ran `Connect-MgGraph` manually earlier, also close the PowerShell window and retry. |
| Module install fails | Run PowerShell as the user that needs it and try `Install-Module Microsoft.Graph -Scope CurrentUser -Force`. Behind a proxy you may need `-Proxy` / `-ProxyCredential`. |

---

## License

Provided as-is, no warranty. Adapt freely for internal use.
