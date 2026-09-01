#Requires -Version 7
<#
  gdap-tenant-compliance-detail.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — detalle profundo de UN tenant cliente

  Qué hace: para un solo tenant (por -TargetTenantId), imprime el detalle
  COMPLETO de cada política de cumplimiento (todos los settings que exige,
  no solo el conteo agregado) y, para cada dispositivo no-cumpliente,
  exactamente qué setting está fallando y por qué. Es el complemento de
  gdap-policy-audit.ps1 (que te dice CUÁLES políticas existen en TODOS los
  tenants) para cuando ya sabes en cuál tenant hay que meterse a fondo, sin
  esperar el barrido completo de los 13.

  Solo lectura — no modifica ni elimina nada.

  Cómo correrlo (ejemplo con ABG):
    cd ~/Documents/stratos-website/scripts
    pwsh ./gdap-tenant-compliance-detail.ps1 -ClientId "f998ba84-ecc3-4688-98b7-c2daf8b9c1c7" -TargetTenantId "b7617723-c204-478c-a810-0bae25562b38"
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

function Get-DeviceSyncInfo {
    param([string]$Token, [string]$DeviceId)
    try {
        $d = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($DeviceId)?`$select=lastSyncDateTime,managementAgent,azureADRegistered,complianceState,deviceEnrollmentType" `
            -Headers @{ Authorization = "Bearer $Token" }
        return "lastSync=$($d.lastSyncDateTime) agent=$($d.managementAgent) enrollment=$($d.deviceEnrollmentType)"
    } catch { return "(no se pudo leer sync info)" }
}

function Get-SettingFailureDetail {
    param([string]$Token, [string]$DeviceId, [string]$PolicyStateId)
    try {
        $settings = Get-GraphAll -Token $Token `
            -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates/$PolicyStateId/settingStates"
        $failingSettings = $settings | Where-Object { $_.state -notin @('compliant', 'notApplicable') }
        if ($failingSettings.Count -eq 0) { return $null }
        return ($failingSettings | ForEach-Object { "$($_.setting): $($_.state)" }) -join ", "
    } catch { return "(no se pudo leer settingStates: $($_.Exception.Message))" }
}

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
    } catch { return "(no se pudo leer deviceCompliancePolicyStates: $($_.Exception.Message))" }
}

Write-Host "== STRATOS — Detalle de cumplimiento: $TargetTenantId ==" -ForegroundColor Cyan
Write-Host "`n[1/2] Iniciando sesión delegada como panel-obo@stratosconsultingpr.com..." -ForegroundColor Cyan
$login = Start-DeviceCodeLogin -TenantId $OwnTenantId
$refreshToken = $login.refresh_token
Write-Host "Login correcto.`n" -ForegroundColor Green

$tok = Get-TenantAccessToken -RefreshToken $refreshToken -TenantId $TargetTenantId
$accessToken = $tok.access_token

Write-Host "[2/2] Leyendo políticas y dispositivos del tenant..." -ForegroundColor Cyan

$policies = Get-GraphAll -Token $accessToken `
    -Url "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"

foreach ($p in $policies) {
    Write-Host "`n=============================================================" -ForegroundColor DarkGray
    Write-Host "Política: $($p.displayName)  [$($p.'@odata.type' -replace '#microsoft\.graph\.', '')]" -ForegroundColor White
    Write-Host "-------------------------------------------------------------" -ForegroundColor DarkGray
    $skip = @('id','displayName','description','createdDateTime','lastModifiedDateTime',
              'version','roleScopeTagIds','scheduledActionsForRule','@odata.type','@odata.context')
    $p.PSObject.Properties | Where-Object { $_.Name -notin $skip } | ForEach-Object {
        Write-Host ("  {0,-45} {1}" -f $_.Name, $_.Value)
    }

    try {
        $assignments = Get-GraphAll -Token $accessToken `
            -Url "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)/assignments"
        if ($assignments.Count -eq 0) {
            Write-Host "  Asignada a: (SIN ASIGNAR -- no aplica a ningún grupo todavía)" -ForegroundColor Red
        } else {
            foreach ($a in $assignments) {
                $targetType = $a.target.'@odata.type' -replace '#microsoft\.graph\.', ''
                if ($a.target.groupId) {
                    $groupName = "(nombre no disponible)"
                    $memberNames = @()
                    try {
                        $g = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$($a.target.groupId)?`$select=displayName" -Headers @{ Authorization = "Bearer $accessToken" }
                        $groupName = $g.displayName
                        $members = Get-GraphAll -Token $accessToken -Url "https://graph.microsoft.com/v1.0/groups/$($a.target.groupId)/members?`$select=displayName,deviceId"
                        $memberNames = $members | ForEach-Object { $_.displayName }
                    } catch {
                        $groupName = "(error leyendo grupo: $($_.Exception.Message))"
                    }
                    $mode = if ($targetType -eq 'exclusionGroupAssignmentTarget') { 'EXCLUIDO' } else { 'incluido' }
                    Write-Host "  Asignada a: grupo `"$groupName`" [$mode] (id=$($a.target.groupId))" -ForegroundColor Cyan
                    if ($memberNames.Count -gt 0) {
                        Write-Host "    Miembros: $($memberNames -join ', ')" -ForegroundColor DarkGray
                    } else {
                        Write-Host "    Miembros: (vacío)" -ForegroundColor DarkGray
                    }
                } else {
                    Write-Host "  Asignada a: $targetType" -ForegroundColor Cyan
                }
            }
        }
    } catch {
        Write-Host "  (no se pudo leer assignments: $($_.Exception.Message))" -ForegroundColor Red
    }

    $overview = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)/deviceStatusOverview" `
        -Headers @{ Authorization = "Bearer $accessToken" }
    Write-Host "`n  Resumen: $($overview.successCount) ok / $($overview.errorCount) error / $($overview.failedCount) noncompliant / $($overview.notApplicableCount) n/a" -ForegroundColor Yellow

    $deviceStatuses = Get-GraphAll -Token $accessToken `
        -Url "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$($p.id)/deviceStatuses"
    $notOk = $deviceStatuses | Where-Object { $_.status -notin @('compliant', 'notApplicable') }
    foreach ($ds in $notOk) {
        Write-Host "`n  Dispositivo: $($ds.deviceDisplayName)  ($($ds.status))" -ForegroundColor Red
        # deviceComplianceDeviceStatus.deviceId is the Azure AD device id,
        # NOT the Intune managedDevice id that deviceCompliancePolicyStates
        # needs -- look the device up by name in managedDevices to get the
        # right id (and pull OS/edition info while we're there).
        try {
            $md = Get-GraphAll -Token $accessToken `
                -Url "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$($ds.deviceDisplayName)'&`$select=id,deviceName,operatingSystem,osVersion,model,manufacturer,complianceState"
            if ($md.Count -eq 0) {
                Write-Host "    (no se encontró en managedDevices por nombre)" -ForegroundColor DarkGray
                continue
            }
            $managedDeviceId = $md[0].id
            Write-Host "    OS: $($md[0].operatingSystem) $($md[0].osVersion) | Modelo: $($md[0].manufacturer) $($md[0].model)" -ForegroundColor DarkGray
        } catch {
            Write-Host "    (no se pudo resolver el managedDevice id: $($_.Exception.Message))" -ForegroundColor Red
            continue
        }
        try {
            $reason = Get-ComplianceFailureReason -Token $accessToken -DeviceId $managedDeviceId
            Write-Host "    $reason" -ForegroundColor Yellow
        } catch {
            Write-Host "    (error leyendo el detalle: $($_.Exception.Message))" -ForegroundColor Red
        }
    }
}

Write-Host "`nListo.`n" -ForegroundColor Green
