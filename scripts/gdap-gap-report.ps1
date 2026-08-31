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
                -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=complianceState eq 'noncompliant'&`$select=deviceName,userPrincipalName,complianceState"
        } catch {
            $complianceKnown = $false
            Write-Host "  ~ $tenantId — sin datos de cumplimiento (probablemente sin licencia de Intune en este tenant)" -ForegroundColor DarkGray
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
                Write-Host "  ~ $tenantId — el reporte de MFA no está disponible (sin Entra ID P1/P2), usando método alterno usuario por usuario" -ForegroundColor DarkGray
            } catch {
                $mfaKnown = $false
                Write-Host "  ~ $tenantId — no se pudo leer MFA por ningún método ($($_.Exception.Message))" -ForegroundColor DarkYellow
            }
        }

        $statusNote = @()
        if (-not $complianceKnown) { $statusNote += "cumplimiento desconocido" }
        if (-not $mfaKnown) { $statusNote += "MFA desconocido" }
        $suffix = if ($statusNote.Count -gt 0) { " ($($statusNote -join ', '))" } else { "" }
        Write-Host "  $tenantId — $($nonCompliant.Count) dispositivo(s) no cumpliente(s), $($noMfa.Count) usuario(s) sin MFA$suffix" -ForegroundColor $(if ($nonCompliant.Count -eq 0 -and $noMfa.Count -eq 0 -and $statusNote.Count -eq 0) { "Green" } else { "Yellow" })

        foreach ($d in $nonCompliant) {
            $rows += [pscustomobject]@{
                TenantId = $tenantId; Tipo = "Dispositivo no cumpliente"
                Detalle = $d.deviceName; Usuario = $d.userPrincipalName
            }
        }
        foreach ($u in $noMfa) {
            $rows += [pscustomobject]@{
                TenantId = $tenantId; Tipo = "Usuario sin MFA"
                Detalle = ""; Usuario = $u.userPrincipalName
            }
        }
    } catch {
        Write-Host "  ✗ $tenantId — no se pudo leer ($($_.Exception.Message))" -ForegroundColor Red
    }
}

if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReporte guardado en: $csvPath ($($rows.Count) fila(s))" -ForegroundColor Green
    Write-Host "Cruza el TenantId contra tu lista de clientes en Partner Center para saber a quién contactar." -ForegroundColor DarkGray
} else {
    Write-Host "`nNo se encontraron brechas — todo cumpliente y con MFA en los tenants leídos." -ForegroundColor Green
}
