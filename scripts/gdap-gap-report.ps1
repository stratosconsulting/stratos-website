#Requires -Version 7
<#
  gdap-gap-report.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — reporte de brechas por cliente

  Qué hace: para cada tenant en scripts/client-tenant-ids.json, lista los
  dispositivos NO cumplientes y los usuarios SIN MFA registrado — el detalle
  que necesitas para saber a quién contactar y subir los % que ves en el
  panel de seguridad del sitio. Esto es SOLO para tu uso — nunca se publica
  en el sitio (el panel público solo muestra el agregado, nunca detalle por
  cliente ni por persona, a propósito).

  Usa el mismo login delegado (panel-obo) y los mismos permisos que ya
  consentiste con obo-delegated-setup.ps1 — no hace falta nada nuevo en
  Azure para correr esto.

  Cómo correrlo:
    cd ~/Documents/stratos-website/scripts
    pwsh ./gdap-gap-report.ps1 -ClientId "f998ba84-ecc3-4688-98b7-c2daf8b9c1c7"

  El reporte se guarda en scripts/reports/gap-report-<fecha>.csv (esa
  carpeta está en .gitignore — nunca se sube a GitHub, es información de
  clientes).
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$OwnTenantId = "c92714b2-6c16-4cd7-a6d0-984d5dd17fa8",
    [string]$ClientTenantIdsFile = (Join-Path $PSScriptRoot "client-tenant-ids.json")
)

$ErrorActionPreference = "Stop"
$GraphScope = "https://graph.microsoft.com/.default offline_access"

function Start-DeviceCodeLogin {
    param([string]$TenantId)

    $dc = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $ClientId; scope = $GraphScope }

    Write-Host "`nPara iniciar sesión, abre $($dc.verification_uri) e ingresa el código $($dc.user_code)`n" -ForegroundColor Cyan
    Write-Host "(esperando a que inicies sesión con panel-obo@stratosconsultingpr.com...)" -ForegroundColor DarkGray

    $deadline = (Get-Date).AddSeconds($dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $dc.interval
        try {
            $token = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $ClientId
                    device_code = $dc.device_code
                }
            return $token
        } catch {
            $err = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue)
            if ($err.error -eq "authorization_pending") { continue }
            throw
        }
    }
    throw "Se venció el tiempo para iniciar sesión. Corre el script de nuevo."
}

function Get-TenantAccessToken {
    param([string]$RefreshToken, [string]$TenantId)

    Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body @{
            client_id     = $ClientId
            grant_type    = "refresh_token"
            refresh_token = $RefreshToken
            scope         = $GraphScope
        }
}

function Get-GraphAll {
    param([string]$Url, [string]$Token)
    $results = @()
    $next = $Url
    while ($next) {
        $res = Invoke-RestMethod -Uri $next -Headers @{ Authorization = "Bearer $Token" }
        if ($res.value) { $results += $res.value }
        $next = $res.'@odata.nextLink'
    }
    return $results
}

# Nombres conocidos de clientes, solo para que el reporte se lea sin tener
# que cruzar cada TenantId contra Partner Center a mano. Si agregas un
# cliente nuevo a client-tenant-ids.json y no aparece aquí, el reporte
# simplemente muestra el TenantId tal cual — no falla nada.
$TenantNames = @{
    "5129a3d7-dd67-4a2f-87f5-744adece3b2b" = "FV International Bank"
    "1ae84e9b-e89c-406d-a315-40035b7e3383" = "International Insurers Consulting Group"
    "478cae85-3c15-49e5-8813-aa8125eb7eb7" = "Iron Shield I.I."
    "5f5d7739-9092-4cbe-9f21-c4774b1ca205" = "Constructores Gilmar"
    "69bf8e2c-f4dd-4cf9-b4dd-6bf1cc9f8954" = "Caribbean Investments and Acquisitions Corp (CIAC)"
    "b35a76af-81dd-4739-832c-6d6f5bad7a42" = "Quinones Financial Group (QFG)"
    "991502e0-11b0-4dc6-b492-d7724408be7c" = "Instituto Sono Radiologico Hostos"
    "b7617723-c204-478c-a810-0bae25562b38" = "ABG Holdings"
    "d3ed6bfb-968a-490b-b03f-9e858bc21da8" = "Teddy Diaz"
    "33660702-685c-4d21-a4a0-71885bd606fc" = "Landrau Law"
    "23a173f8-5ea4-4dba-8c74-9febf0f4faed" = "Ador Storage Development LLC"
    "b1418c9d-dc76-4658-afd9-072e4fcde053" = "MAVI"
    "f7e0e809-d35d-475e-adee-76fbfabd1cff" = "Aquos Air & Water"
}
function Get-TenantLabel {
    param([string]$TenantId)
    if ($TenantNames.ContainsKey($TenantId)) { return "$($TenantNames[$TenantId]) ($TenantId)" }
    return $TenantId
}

# For a device whose compliance detail comes back as "error" (as opposed to
# a real "noncompliant" setting violation), this pulls the device's own
# sync/management info — the usual root cause for "error" is a stale or
# broken agent, not an actual policy violation, and this tells you which.
function Get-DeviceSyncInfo {
    param([string]$Token, [string]$DeviceId)

    try {
        $d = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($DeviceId)?`$select=lastSyncDateTime,managementAgent,azureADRegistered,complianceState,deviceEnrollmentType" `
            -Headers @{ Authorization = "Bearer $Token" }
        $daysSinceSync = if ($d.lastSyncDateTime) {
            # lastSyncDateTime comes back UTC; compare against UTC "now" too,
            # otherwise a local timezone offset shows up as a bogus negative
            # "days since sync".
            [math]::Round(((Get-Date).ToUniversalTime() - [datetime]$d.lastSyncDateTime).TotalDays, 1)
        } else { $null }
        $staleness = if ($null -eq $daysSinceSync) { "nunca ha hecho sync" }
                     elseif ($daysSinceSync -gt 14) { "$daysSinceSync días sin sync — probablemente offline/desinstalado" }
                     elseif ($daysSinceSync -gt 3) { "$daysSinceSync días sin sync — revisa que el dispositivo esté prendido y conectado" }
                     else { "sync reciente ($daysSinceSync días) — no es un problema de sync, hay que ver el detalle de la política (abajo)" }
        return "agente: $($d.managementAgent), $staleness"
    } catch {
        return "(no se pudo leer info de sync del dispositivo)"
    }
}

# One level deeper than the policy-level state: for a policy stuck in
# "error", this asks WHICH SETTING inside that policy is the one erroring
# (e.g. BitLocker status, firewall status, a specific password rule) —
# policy-level "error" alone doesn't say what's actually wrong.
function Get-SettingFailureDetail {
    param([string]$Token, [string]$DeviceId, [string]$PolicyStateId)

    try {
        $settings = Get-GraphAll -Token $Token `
            -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates/$PolicyStateId/settingStates"
        $failingSettings = $settings | Where-Object { $_.state -notin @('compliant', 'notApplicable') }
        if ($failingSettings.Count -eq 0) { return $null }
        return ($failingSettings | ForEach-Object { "$($_.setting): $($_.state)" }) -join ", "
    } catch {
        return $null
    }
}

# For a non-compliant device, this looks up WHICH assigned compliance
# policy is failing and why (deviceCompliancePolicyStates includes a
# human-readable displayName + the pass/fail state per assigned policy).
# No new permission needed — DeviceManagementConfiguration.Read.All,
# already consented, covers this. When any policy is stuck in "error"
# (agent/sync issue, not a real setting violation) it also pulls the
# device's sync info so you know whether it's just offline.
function Get-ComplianceFailureReason {
    param([string]$Token, [string]$DeviceId)

    try {
        $states = Get-GraphAll -Token $Token `
            -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates"
        $failing = $states | Where-Object { $_.state -notin @('compliant', 'notApplicable') }
        if ($failing.Count -eq 0) { return "(sin detalle disponible)" }
        $parts = @()
        foreach ($p in $failing) {
            $line = "$($p.displayName): $($p.state)"
            $settingDetail = Get-SettingFailureDetail -Token $Token -DeviceId $DeviceId -PolicyStateId $p.id
            if ($settingDetail) { $line += " [$settingDetail]" }
            $parts += $line
        }
        $detail = $parts -join "; "
        if ($failing | Where-Object { $_.state -eq 'error' }) {
            $detail += " | " + (Get-DeviceSyncInfo -Token $Token -DeviceId $DeviceId)
        }
        return $detail
    } catch {
        return "(no se pudo leer el detalle de la política)"
    }
}

# Fallback for tenants where reports/authenticationMethods/userRegistrationDetails
# 403s (that report needs Entra ID P1/P2 on the CLIENT tenant — several of
# ours don't have it). This reads each user's actual registered
# authentication methods instead — slower (one call per user) but doesn't
# need any premium license, just UserAuthenticationMethod.Read.All.
function Get-NoMfaUsersFallback {
    param([string]$Token)

    $users = Get-GraphAll -Token $Token `
        -Url "https://graph.microsoft.com/v1.0/users?`$filter=accountEnabled eq true&`$select=id,userPrincipalName"

    $noMfa = @()
    foreach ($u in $users) {
        try {
            $methods = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/v1.0/users/$($u.id)/authentication/methods" `
                -Headers @{ Authorization = "Bearer $Token" }
            # A password is always method #1 for every user — MFA means at
            # least one method beyond that (Authenticator, phone, FIDO2,
            # software OATH token, etc).
            $strongMethods = $methods.value | Where-Object {
                $_.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod'
            }
            if ($strongMethods.Count -eq 0) {
                $noMfa += [pscustomobject]@{ userPrincipalName = $u.userPrincipalName }
            }
        } catch {
            # One user's methods failing to read shouldn't stop the whole
            # tenant — skip them, they just won't show up either way.
        }
    }
    return $noMfa
}

if (-not (Test-Path $ClientTenantIdsFile)) {
    throw "No se encontró $ClientTenantIdsFile. Corre primero gdap-bulk-setup.ps1."
}
$clientTenantIds = Get-Content $ClientTenantIdsFile | ConvertFrom-Json

Write-Host "== STRATOS — Reporte de brechas (cumplimiento + MFA) ==" -ForegroundColor Cyan
Write-Host "`n[1/2] Iniciando sesión delegada como panel-obo@stratosconsultingpr.com..." -ForegroundColor Cyan
$login = Start-DeviceCodeLogin -TenantId $OwnTenantId
$refreshToken = $login.refresh_token
Write-Host "Login correcto.`n" -ForegroundColor Green

$reportsDir = Join-Path $PSScriptRoot "reports"
if (-not (Test-Path $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir | Out-Null }
$stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$csvPath = Join-Path $reportsDir "gap-report-$stamp.csv"
$rows = @()

Write-Host "[2/2] Revisando cada tenant..." -ForegroundColor Cyan
foreach ($tenantId in $clientTenantIds) {
    try {
        $tok = Get-TenantAccessToken -RefreshToken $refreshToken -TenantId $tenantId
        if ($tok.refresh_token) { $refreshToken = $tok.refresh_token }
        $accessToken = $tok.access_token

        $nonCompliant = @()
        $complianceKnown = $true
        try {
            $nonCompliant = Get-GraphAll -Token $accessToken `
                -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=complianceState eq 'noncompliant'&`$select=id,deviceName,userPrincipalName,complianceState"
        } catch {
            $complianceKnown = $false
            Write-Host "  ~ $(Get-TenantLabel $tenantId) — sin datos de cumplimiento (probablemente sin licencia de Intune en este tenant)" -ForegroundColor DarkGray
        }

        $noMfa = @()
        $mfaKnown = $true
        try {
            $noMfa = Get-GraphAll -Token $accessToken `
                -Url "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?`$filter=isMfaRegistered eq false&`$select=userPrincipalName"
        } catch {
            # Fallback: this report needs Entra ID P1/P2 on the client
            # tenant. Try the per-user method instead before giving up.
            try {
                $noMfa = Get-NoMfaUsersFallback -Token $accessToken
                Write-Host "  ~ $(Get-TenantLabel $tenantId) — el reporte de MFA no está disponible (sin Entra ID P1/P2), usando método alterno usuario por usuario" -ForegroundColor DarkGray
            } catch {
                $mfaKnown = $false
                Write-Host "  ~ $(Get-TenantLabel $tenantId) — no se pudo leer MFA por ningún método ($($_.Exception.Message))" -ForegroundColor DarkYellow
            }
        }

        $statusNote = @()
        if (-not $complianceKnown) { $statusNote += "cumplimiento desconocido" }
        if (-not $mfaKnown) { $statusNote += "MFA desconocido" }
        $suffix = if ($statusNote.Count -gt 0) { " ($($statusNote -join ', '))" } else { "" }
        Write-Host "  $(Get-TenantLabel $tenantId) — $($nonCompliant.Count) dispositivo(s) no cumpliente(s), $($noMfa.Count) usuario(s) sin MFA$suffix" -ForegroundColor $(if ($nonCompliant.Count -eq 0 -and $noMfa.Count -eq 0 -and $statusNote.Count -eq 0) { "Green" } else { "Yellow" })

        foreach ($d in $nonCompliant) {
            $reason = Get-ComplianceFailureReason -Token $accessToken -DeviceId $d.id
            $rows += [pscustomobject]@{
                TenantId = $tenantId; Tipo = "Dispositivo no cumpliente"
                Detalle = "$($d.deviceName) — $reason"; Usuario = $d.userPrincipalName
            }
        }
        foreach ($u in $noMfa) {
            $rows += [pscustomobject]@{
                TenantId = $tenantId; Tipo = "Usuario sin MFA"
                Detalle = ""; Usuario = $u.userPrincipalName
            }
        }
    } catch {
        Write-Host "  ✗ $(Get-TenantLabel $tenantId) — no se pudo leer ($($_.Exception.Message))" -ForegroundColor Red
    }
}

if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReporte guardado en: $csvPath ($($rows.Count) fila(s))" -ForegroundColor Green
    Write-Host "Cruza el TenantId contra tu lista de clientes en Partner Center para saber a quién contactar." -ForegroundColor DarkGray
} else {
    Write-Host "`nNo se encontraron brechas — todo cumpliente y con MFA en los tenants leídos." -ForegroundColor Green
}
