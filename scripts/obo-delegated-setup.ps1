#Requires -Version 7.0
<#
  obo-delegated-setup.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — Panel de Seguridad en Vivo, fase 2b
  (autenticación delegada OBO para leer datos de los tenants de clientes)

  Por qué existe este script:
  GDAP NO acepta tokens de aplicación pura (client credentials) contra un
  tenant de cliente — solo acepta tokens delegados app+usuario ("Secure
  Application Model"). Este script hace, en una sola corrida:

    1. Inicia sesión como la cuenta OBO usando "device code flow" (inicias
       sesión en una página web con un código de un solo uso — no hace
       falta contraseña en la terminal).
    2. Captura el refresh token delegado (válido para pedir tokens contra
       Graph en CUALQUIER tenant donde exista consentimiento).
    3. Para cada tenant de cliente en client-tenant-ids.json, llama a la API
       de Partner Center para darle a la app el consentimiento explícito que
       GDAP exige (esto es aparte de la asignación de rol GDAP que ya hizo
       gdap-bulk-setup.ps1 — son dos pasos distintos).
    4. Imprime el refresh token para que lo pegues UNA VEZ en la
       Application Setting GRAPH_OBO_REFRESH_TOKEN de la Function App. De
       ahí en adelante la propia función lo rota sola (ver
       api/security-stats/tokenStore.js) — este script no hay que
       correrlo de nuevo salvo que agregues un cliente GDAP nuevo.

  ANTES de correr esto, en el App Registration del panel (portal de Azure),
  necesitas cambiar dos cosas UNA SOLA VEZ:

    a) Authentication -> Advanced settings -> "Allow public client flows" = Yes
       (lo necesita el device code flow).
    b) Authentication -> Supported account types -> cambiar a
       "Accounts in any organizational directory (multitenant)"
       (lo necesita para pedir tokens contra tenants de clientes).
    c) API permissions -> Add a permission -> Microsoft Graph -> Delegated:
       Device.Read.All, DeviceManagementManagedDevices.Read.All,
       DeviceManagementConfiguration.Read.All, AuditLog.Read.All,
       SecurityAlert.Read.All, offline_access, User.Read
       -> "Grant admin consent for STRATOS".
       (Las de tipo Application que ya tienes de la fase 1 se quedan igual;
       estas Delegated son ADICIONALES, no las reemplazan.)
    d) API permissions -> Add a permission -> APIs my organization uses ->
       busca "Partner Center" (o "Partner Customer Delegated Admin API") ->
       Delegated -> user_impersonation -> Grant admin consent.
       Si no aparece Partner Center en la lista, avísame — significa que
       hay que registrar la app ahí primero, es un paso aparte.

  Uso:
    pwsh ./obo-delegated-setup.ps1 -OboUserPrincipalName "panel-obo@stratosconsultingpr.com" `
        -ClientId "<application (client) id del panel>" `
        -TenantId "<tenant id de STRATOS>"
  ---------------------------------------------------------------------------
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$OboUserPrincipalName,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [string]$ClientTenantIdsFile = (Join-Path $PSScriptRoot "client-tenant-ids.json"),

    # Scopes delegados que la función necesita — deben coincidir con lo que
    # se consintió en el paso (c) de arriba.
    [string[]]$GraphScopes = @(
        "Device.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "AuditLog.Read.All",
        "SecurityAlert.Read.All",
        "offline_access"
    )
)

$ErrorActionPreference = "Stop"

function Start-DeviceCodeLogin {
    param([string]$Scope)

    $dc = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $ClientId; scope = $Scope }

    Write-Host "`n$($dc.message)`n" -ForegroundColor Yellow
    Write-Host "(esperando a que inicies sesión como $OboUserPrincipalName con MFA...)" -ForegroundColor DarkGray

    $interval = [int]$dc.interval
    $expiresAt = (Get-Date).AddSeconds([int]$dc.expires_in)

    while ((Get-Date) -lt $expiresAt) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method Post `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body @{
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    client_id   = $ClientId
                    device_code = $dc.device_code
                }
            return $tok
        } catch {
            $err = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue)
            if ($err.error -eq "authorization_pending") { continue }
            if ($err.error -eq "authorization_declined") { throw "Inicio de sesión rechazado." }
            throw
        }
    }
    throw "Se venció el tiempo para iniciar sesión. Corre el script de nuevo."
}

function Get-TokenForScope {
    param([string]$RefreshToken, [string]$Scope)

    Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body @{
            client_id     = $ClientId
            grant_type    = "refresh_token"
            refresh_token = $RefreshToken
            scope         = $Scope
        }
}

Write-Host "== STRATOS Security Panel — OBO delegated setup ==" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1-2. Device code login como la cuenta OBO, capturamos el refresh token.
# ---------------------------------------------------------------------------
Write-Host "`n[1/3] Iniciando sesión delegada como $OboUserPrincipalName..." -ForegroundColor Yellow
$graphScope = ($GraphScopes -join " ")
$firstToken = Start-DeviceCodeLogin -Scope $graphScope

if (-not $firstToken.refresh_token) {
    throw "No se recibió refresh_token. Revisa que 'offline_access' esté en los scopes delegados consentidos (paso c de arriba)."
}

$refreshToken = $firstToken.refresh_token
Write-Host "Login delegado correcto. Refresh token capturado." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Consentimiento de la app en cada tenant de cliente vía Partner Center.
# ---------------------------------------------------------------------------
Write-Host "`n[2/3] Dando consentimiento de la app en cada tenant de cliente..." -ForegroundColor Yellow

if (-not (Test-Path $ClientTenantIdsFile)) {
    throw "No se encontró $ClientTenantIdsFile. Corre primero gdap-bulk-setup.ps1."
}
$clientTenantIds = Get-Content $ClientTenantIdsFile | ConvertFrom-Json

# Intercambiamos el refresh token por uno con scope de Partner Center — no
# hace falta iniciar sesión de nuevo, un mismo refresh token sirve para
# pedir tokens contra distintos recursos, siempre que la app tenga permiso
# delegado para ese recurso (paso d de arriba).
$pcToken = Get-TokenForScope -RefreshToken $refreshToken -Scope "https://api.partnercenter.microsoft.com/user_impersonation offline_access"
if ($pcToken.refresh_token) { $refreshToken = $pcToken.refresh_token }

$graphAppId = "00000003-0000-0000-c000-000000000000"  # ID fijo de Microsoft Graph, no cambia por tenant
$grantScope = ($GraphScopes | Where-Object { $_ -ne "offline_access" }) -join " "

$consented = @()
$consentFailed = @()

foreach ($tenantId in $clientTenantIds) {
    try {
        $body = @{
            applicationId    = $ClientId
            applicationGrants = @(
                @{ enterpriseApplicationId = $graphAppId; scope = $grantScope }
            )
        } | ConvertTo-Json -Depth 5

        Invoke-RestMethod -Method Post `
            -Uri "https://api.partnercenter.microsoft.com/v1/customers/$tenantId/applicationconsents" `
            -Headers @{ Authorization = "Bearer $($pcToken.access_token)"; "Content-Type" = "application/json" } `
            -Body $body | Out-Null

        Write-Host "  ✓ $tenantId — consentimiento otorgado." -ForegroundColor Green
        $consented += $tenantId
    } catch {
        Write-Host "  ✗ $tenantId — error: $($_.Exception.Message)" -ForegroundColor Red
        $consentFailed += $tenantId
    }
}

# ---------------------------------------------------------------------------
# Resumen + refresh token para pegar en Azure
# ---------------------------------------------------------------------------
Write-Host "`n[3/3] Resumen" -ForegroundColor Cyan
Write-Host "-----------------------------------------------------------"
Write-Host "Tenants con consentimiento otorgado: $($consented.Count)" -ForegroundColor Green
$consented | ForEach-Object { Write-Host "   - $_" -ForegroundColor Green }
if ($consentFailed.Count -gt 0) {
    Write-Host "Tenants con error (revisa el mensaje arriba; el más común es 'customer id' distinto al tenant id — avísame si ves esto):" -ForegroundColor Yellow
    $consentFailed | ForEach-Object { Write-Host "   - $_" -ForegroundColor Yellow }
}

Write-Host "`nPega este valor UNA VEZ en la Application Setting GRAPH_OBO_REFRESH_TOKEN" -ForegroundColor Cyan
Write-Host "(Static Web App -> Configuration -> Application settings). De ahí en" -ForegroundColor Cyan
Write-Host "adelante la función lo rota sola, no hay que volver a tocarlo:" -ForegroundColor Cyan
Write-Host "`n$refreshToken`n" -ForegroundColor White

Write-Host "También agrega/actualiza la Application Setting GRAPH_CLIENT_TENANT_IDS" -ForegroundColor Cyan
Write-Host "con el contenido exacto de $ClientTenantIdsFile (cópialo tal cual)." -ForegroundColor Cyan
