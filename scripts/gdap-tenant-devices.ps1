#Requires -Version 7
<#
  gdap-tenant-devices.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — inventario crudo de dispositivos de UN tenant

  Qué hace: lista TODOS los managedDevices de un tenant (sin filtrar por
  cumplimiento), con su plataforma, agente de administración y último sync.
  Sirve para diagnosticar por qué una política de cumplimiento muestra
  0 ok / 0 error / 0 noncompliant / 0 n/a a pesar de estar asignada a
  "Todos los dispositivos" — normalmente porque los dispositivos reales
  no están inscritos en Intune (MDM), o son de otra plataforma.

  Solo lectura — no modifica ni elimina nada.

  Cómo correrlo (ejemplo con QFG):
    cd ~/Documents/stratos-website/scripts
    pwsh ./gdap-tenant-devices.ps1 -ClientId "f998ba84-ecc3-4688-98b7-c2daf8b9c1c7" -TargetTenantId "b35a76af-81dd-4739-832c-6d6f5bad7a42"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TargetTenantId,

    [string]$OwnTenantId = "c92714b2-6c16-4cd7-a6d0-984d5dd17fa8"
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

Write-Host "== STRATOS — Inventario de dispositivos: $TargetTenantId ==" -ForegroundColor Cyan
Write-Host "`n[1/2] Iniciando sesión delegada como panel-obo@stratosconsultingpr.com..." -ForegroundColor Cyan
$login = Start-DeviceCodeLogin -TenantId $OwnTenantId
$refreshToken = $login.refresh_token
Write-Host "Login correcto.`n" -ForegroundColor Green

$tok = Get-TenantAccessToken -RefreshToken $refreshToken -TenantId $TargetTenantId
$accessToken = $tok.access_token

Write-Host "[2/2] Leyendo dispositivos administrados del tenant..." -ForegroundColor Cyan

$devices = Get-GraphAll -Token $accessToken `
    -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=deviceName,operatingSystem,osVersion,managementAgent,complianceState,lastSyncDateTime,azureADRegistered,deviceEnrollmentType,userPrincipalName"

if ($devices.Count -eq 0) {
    Write-Host "`nNo hay ningún dispositivo inscrito en Intune (managedDevices) en este tenant." -ForegroundColor Red
} else {
    Write-Host "`nTotal: $($devices.Count) dispositivo(s) inscritos en Intune`n" -ForegroundColor Yellow
    foreach ($d in $devices) {
        Write-Host "  $($d.deviceName)" -ForegroundColor White
        Write-Host "    OS: $($d.operatingSystem) $($d.osVersion) | Agente: $($d.managementAgent) | Inscripción: $($d.deviceEnrollmentType)"
        Write-Host "    Cumplimiento: $($d.complianceState) | Último sync: $($d.lastSyncDateTime) | Usuario: $($d.userPrincipalName)"
    }
}

Write-Host "`nListo.`n" -ForegroundColor Green
