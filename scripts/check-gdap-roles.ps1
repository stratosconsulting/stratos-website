#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Partner
<#
  check-gdap-roles.ps1
  ---------------------------------------------------------------------------
  Diagnóstico rápido: para cada tenant de cliente en client-tenant-ids.json,
  muestra qué rol de Entra tiene realmente la relación GDAP activa
  (Cloud Application Administrator vs Application Administrator vs otro).

  Esto es para entender por qué algunos clientes devuelven datos completos
  del panel y otros (con login correcto, sin error) devuelven todo vacío —
  la sospecha es que el rol que les tocó no alcanza para leer Intune /
  seguridad, aunque sí alcance para el login delegado en sí.

  Uso:
    pwsh ./check-gdap-roles.ps1
  (conéctate con tu cuenta de admin de STRATOS, la misma de siempre)
  ---------------------------------------------------------------------------
#>

$ErrorActionPreference = "Stop"

$CloudAppAdminRoleId = "158c047a-c907-4556-b7ef-446551a6b5f7"
$AppAdminRoleId      = "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3"
$IntuneAdminRoleId   = "3a2c62db-5318-420d-8d74-23affee5d9d5"
$GlobalAdminRoleId   = "62e90394-69f5-4237-9190-012177145e10"
$SecurityReaderRoleId= "5d6b6bb7-de71-4623-b4af-96380a352509"

$roleNames = @{
    $CloudAppAdminRoleId  = "Cloud Application Administrator"
    $AppAdminRoleId       = "Application Administrator"
    $IntuneAdminRoleId    = "Intune Administrator"
    $GlobalAdminRoleId    = "Global Administrator"
    $SecurityReaderRoleId = "Security Reader"
}

Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "DelegatedAdminRelationship.Read.All" | Out-Null

$clientTenantIdsFile = Join-Path $PSScriptRoot "client-tenant-ids.json"
$clientTenantIds = Get-Content $clientTenantIdsFile | ConvertFrom-Json

$relationships = Get-MgTenantRelationshipDelegatedAdminRelationship -All |
    Where-Object { $_.Status -eq "active" }

Write-Host "`nRoles por tenant de cliente (los que están en client-tenant-ids.json):`n" -ForegroundColor Cyan

foreach ($tenantId in $clientTenantIds) {
    $rel = $relationships | Where-Object { $_.Customer.TenantId -eq $tenantId } | Select-Object -First 1
    if (-not $rel) {
        Write-Host "$tenantId -> no se encontró relación activa (raro, revisar)" -ForegroundColor Red
        continue
    }
    $roleIds = $rel.AccessDetails.UnifiedRoles.RoleDefinitionId
    $roleLabels = $roleIds | ForEach-Object { if ($roleNames.ContainsKey($_)) { $roleNames[$_] } else { $_ } }
    Write-Host "$tenantId ($($rel.Customer.DisplayName)) -> $($roleLabels -join ', ')"
}
