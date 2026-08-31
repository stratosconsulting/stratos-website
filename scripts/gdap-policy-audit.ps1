#Requires -Version 7
<#
  gdap-policy-audit.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — inventario de políticas de cumplimiento por cliente

  Qué hace: para cada tenant en client-tenant-ids.json, lista TODAS las
  políticas de cumplimiento de Intune (no solo las que tienen dispositivos
  fallando ahora mismo) con el conteo agregado de dispositivos en cada
  estado (cumpliente, error, no-cumpliente, conflicto, pendiente, no
  aplica). Es el mapa completo para ir política por política decidiendo si
  se ajusta o se elimina — gdap-gap-report.ps1 te dice QUÉ dispositivo está
  fallando y por qué; este te dice qué políticas existen en cada tenant y
  qué tan sano está cada una en conjunto.

  Solo lectura — no modifica ni elimina nada. Los cambios reales (ajustar
  una regla, desactivar o borrar una política huérfana) se hacen a mano en
  el Intune admin center de cada cliente, después de revisar el reporte.

  Cómo correrlo:
    cd ~/Documents/stratos-website/scripts
    pwsh ./gdap-policy-audit.ps1 -ClientId "f998ba84-ecc3-4688-98b7-c2daf8b9c1c7"

  El reporte se guarda en scripts/reports/policy-audit-<fecha>.csv (esa
  carpeta está en .gitignore, nunca sube a GitHub).
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$OwnTenantId = "c92714b2-6c16-4cd7-a6d0-984d5dd17fa8",
    [string]$ClientTenantIdsFile = (Join-Path $PSScriptRoot "client-tenant-ids.json")
)

$ErrorActionPreference = "Stop"
$GraphScope = "https://graph.microsoft.com/.default offline_access"

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
            return Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body @{ grant_type = "urn:ietf:params:oauth:grant-type:device_code"; client_id = $ClientId; device_code = $dc.device_code }
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
        -Body @{ client_id = $ClientId; grant_type = "refresh_token"; refresh_token = $RefreshToken; scope = $GraphScope }
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

# Traduce el @odata.type de la política a algo legible sin tener que
# recordar los nombres internos de Graph.
function Get-PolicyPlatform {
    param([string]$OdataType)
    switch -Wildcard ($OdataType) {
        "*windows10*"  { return "Windows" }
        "*windows81*"  { return "Windows 8.1" }
        "*ios*"        { return "iOS" }
        "*macOS*"      { return "macOS" }
        "*android*"    { return "Android" }
        default        { return ($OdataType -replace '#microsoft\.graph\.', '') }
    }
}

# Heurística simple de qué hacer con cada política, para no tener que
# pensarlo desde cero por cada fila del CSV.
function Get-Recommendation {
    param($Overview, [int]$AssignedCount)
    $total = $Overview.successCount + $Overview.errorCount + $Overview.failedCount + $Overview.notApplicableCount + $Overview.pendingCount
    if ($total -eq 0) {
        return "Sin dispositivos evaluados todavía — revisar si tiene asignación (grupo vacío = candidata a eliminar si es residual)"
    }
    if ($Overview.errorCount -eq 0 -and $Overview.failedCount -eq 0) {
        return "Sana — no tocar"
    }
    $badRatio = ($Overview.errorCount + $Overview.failedCount) / $total
    if ($badRatio -ge 0.5) {
        return "Mayoría de dispositivos fallando — revisar la regla específica que falla (ver gdap-gap-report.ps1 para el detalle por dispositivo) antes de decidir si se ajusta o se elimina"
    }
    return "Algunos dispositivos fallando — normal si son pocos casos aislados, confirmar con gdap-gap-report.ps1 cuáles son"
}

if (-not (Test-Path $ClientTenantIdsFile)) {
    throw "No se encontró $ClientTenantIdsFile. Corre primero gdap-bulk-setup.ps1."
}
$clientTenantIds = Get-Content $ClientTenantIdsFile | ConvertFrom-Json

Write-Host "== STRATOS — Inventario de políticas de cumplimiento por cliente ==" -ForegroundColor Cyan
Write-Host "`n[1/2] Iniciando sesión delegada como panel-obo@stratosconsultingpr.com..." -ForegroundColor Cyan
$login = Start-DeviceCodeLogin -TenantId $OwnTenantId
$refreshToken = $login.refresh_token
Write-Host "Login correcto.`n" -ForegroundColor Green

$reportsDir = Join-Path $PSScriptRoot "reports"
if (-not (Test-Path $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir | Out-Null }
$stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$csvPath = Join-Path $reportsDir "policy-audit-$stamp.csv"
$rows = @()

Write-Host "[2/2] Revisando políticas de cada tenant..." -ForegroundColor Cyan
foreach ($tenantId in $clientTenantIds) {
    $label = Get-TenantLabel $tenantId
    try {
        $tok = Get-TenantAccessToken -RefreshToken $refreshToken -TenantId $tenantId
        if ($tok.refresh_token) { $refreshToken = $tok.refresh_token }
        $accessToken = $tok.access_token

        $policies = @()
        try {
            $policies = Get-GraphAll -Token $accessToken `
                -Url "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies?`$select=id,displayName"
        } catch {
            Write-Host "  ~ $label — sin licencia de Intune / sin acceso a políticas de cumplimiento" -ForegroundColor DarkGray
            continue
        }

        if ($policies.Count -eq 0) {
            Write-Host "  $label — 0 políticas de cumplimiento configuradas" -ForegroundColor DarkGray
            continue
        }

        foreach ($p in $policies) {
            try {
                $full = Invoke-RestMethod `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)" `
                    -Headers @{ Authorization = "Bearer $accessToken" }
                $overview = Invoke-RestMethod `
                    -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)/deviceStatusOverview" `
                    -Headers @{ Authorization = "Bearer $accessToken" }
                $platform = Get-PolicyPlatform $full.'@odata.type'
                $recommendation = Get-Recommendation -Overview $overview -AssignedCount 0

                Write-Host "  $label — [$platform] $($p.displayName): $($overview.successCount) ok / $($overview.errorCount) error / $($overview.failedCount) noncompliant / $($overview.notApplicableCount) n/a" `
                    -ForegroundColor $(if ($overview.errorCount -eq 0 -and $overview.failedCount -eq 0) { "Green" } else { "Yellow" })

                $rows += [pscustomobject]@{
                    TenantId       = $tenantId
                    Cliente        = if ($TenantNames.ContainsKey($tenantId)) { $TenantNames[$tenantId] } else { "" }
                    Politica       = $p.displayName
                    Plataforma     = $platform
                    Cumpliente     = $overview.successCount
                    Error          = $overview.errorCount
                    NoCumpliente   = $overview.failedCount
                    NoAplica       = $overview.notApplicableCount
                    Pendiente      = $overview.pendingCount
                    Recomendacion  = $recommendation
                }
            } catch {
                Write-Host "    ✗ no se pudo leer detalle de política $($p.displayName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "  ✗ $label — no se pudo leer ($($_.Exception.Message))" -ForegroundColor Red
    }
}

if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReporte guardado en: $csvPath ($($rows.Count) política(s) en total)" -ForegroundColor Green
    Write-Host "Revísalo por 'Recomendacion' — ve tenant por tenant, empezando por las filas en amarillo/rojo." -ForegroundColor DarkGray
} else {
    Write-Host "`nNo se encontraron políticas de cumplimiento en los tenants leídos." -ForegroundColor Green
}
