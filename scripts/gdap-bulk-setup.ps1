#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Identity.Partner
<#
  gdap-bulk-setup.ps1
  ---------------------------------------------------------------------------
  STRATOS Consulting Group — Panel de Seguridad en Vivo, fase 2 (multi-cliente)

  Qué hace este script, en una sola corrida:
    1. Crea (si no existe) el grupo de seguridad "assignable to role" que va
       a representar a la app/cuenta OBO en todas las relaciones GDAP.
    2. Se conecta a Microsoft Graph con tu cuenta de admin del tenant de
       partner de STRATOS.
    3. Lista TODAS tus relaciones GDAP activas con clientes.
    4. Para cada una, revisa si el rol requerido (Cloud Application
       Administrator o Application Administrator) YA está incluido en esa
       relación:
         - Si SÍ está incluido -> asigna el grupo de seguridad a ese rol
           automáticamente (no requiere que el cliente vuelva a aprobar
           nada).
         - Si NO está incluido -> Microsoft NO permite agregarlo a una
           relación ya activa sin que el cliente apruebe una relación
           nueva. El script simplemente lo reporta al final en una lista
           clara de "clientes pendientes" — ese es el único paso que de
           verdad requiere tocar cada cliente, y no se puede evitar.

  Qué necesitas ANTES de correr esto (una sola vez, ver el documento técnico
  sección 11 para más detalle):
    - La cuenta "on-behalf-of" (ej. panel-obo@stratosconsultingpr.com) ya
      creada, con MFA activado.
    - Haber iniciado sesión con esa cuenta al menos una vez (para el refresh
      token OBO — ver sección 11, paso 5 del documento técnico).
    - Los módulos de PowerShell instalados:
        Install-Module Microsoft.Graph -Scope CurrentUser

  Cómo correrlo:
    .\gdap-bulk-setup.ps1 -OboUserPrincipalName "panel-obo@stratosconsultingpr.com"

  Es seguro correrlo más de una vez (idempotente): si el grupo ya existe no
  lo duplica, y si una asignación ya existe no la duplica.
  ---------------------------------------------------------------------------
#>

param(
    # UPN de la cuenta on-behalf-of dedicada al panel (la que ya tiene MFA).
    [Parameter(Mandatory = $true)]
    [string]$OboUserPrincipalName,

    # Nombre del grupo de seguridad a crear/usar. No hace falta tocarlo.
    [string]$GroupDisplayName = "Stratos Security Panel - GDAP Access",

    # Roles Entra ID built-in aceptados para leer datos vía Graph en cada
    # cliente. IDs oficiales de Microsoft (estables, no cambian por tenant).
    [string]$CloudAppAdminRoleId = "158c047a-c907-4556-b7ef-446551a6b5f7",
    [string]$AppAdminRoleId      = "9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3"
)

$ErrorActionPreference = "Stop"

Write-Host "== STRATOS Security Panel — GDAP bulk setup ==" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 1. Conectar a Microsoft Graph con los permisos necesarios
# ---------------------------------------------------------------------------
Write-Host "`n[1/5] Conectando a Microsoft Graph (te va a pedir iniciar sesión)..." -ForegroundColor Yellow
Connect-MgGraph -Scopes @(
    "DelegatedAdminRelationship.ReadWrite.All",
    "Group.ReadWrite.All",
    "Directory.Read.All"
) | Out-Null
$ctx = Get-MgContext
Write-Host "Conectado como $($ctx.Account) en el tenant $($ctx.TenantId)" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Crear (o reutilizar) el grupo de seguridad "assignable to role"
# ---------------------------------------------------------------------------
Write-Host "`n[2/5] Buscando/creando el grupo de seguridad '$GroupDisplayName'..." -ForegroundColor Yellow

$group = Get-MgGroup -Filter "displayName eq '$GroupDisplayName'" -ErrorAction SilentlyContinue
if (-not $group) {
    $group = New-MgGroup -DisplayName $GroupDisplayName `
        -MailEnabled:$false `
        -MailNickname "stratos-panel-gdap" `
        -SecurityEnabled:$true `
        -IsAssignableToRole:$true `
        -Description "Grupo usado por la Azure Function del panel de seguridad para leer datos de clientes vía GDAP. No agregar personas aquí manualmente."
    Write-Host "Grupo creado: $($group.Id)" -ForegroundColor Green
} else {
    Write-Host "Grupo ya existía: $($group.Id)" -ForegroundColor Green
}

# Asegura que la cuenta OBO sea miembro del grupo
$oboUser = Get-MgUser -Filter "userPrincipalName eq '$OboUserPrincipalName'"
if (-not $oboUser) {
    throw "No se encontró el usuario '$OboUserPrincipalName'. Créalo primero (con MFA activo) antes de correr este script."
}
$existingMember = Get-MgGroupMember -GroupId $group.Id -All | Where-Object { $_.Id -eq $oboUser.Id }
if (-not $existingMember) {
    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $oboUser.Id
    Write-Host "Cuenta OBO agregada al grupo." -ForegroundColor Green
} else {
    Write-Host "Cuenta OBO ya era miembro del grupo." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 3. Listar todas las relaciones GDAP activas
# ---------------------------------------------------------------------------
Write-Host "`n[3/5] Listando relaciones GDAP activas..." -ForegroundColor Yellow

$relationships = Get-MgTenantRelationshipDelegatedAdminRelationship -All |
    Where-Object { $_.Status -eq "active" }

Write-Host "Encontradas $($relationships.Count) relaciones GDAP activas." -ForegroundColor Green

if ($relationships.Count -eq 0) {
    Write-Host "No hay relaciones GDAP activas — no hay nada más que hacer. Revisa Partner Center." -ForegroundColor Red
    return
}

# ---------------------------------------------------------------------------
# 4. Para cada relación: asignar el grupo al rol correcto si ya lo tiene
# ---------------------------------------------------------------------------
Write-Host "`n[4/5] Procesando cada cliente..." -ForegroundColor Yellow

$ready = @()
$readyTenantIds = @{}
$needsNewRequest = @()
$alreadyAssigned = @()

foreach ($rel in $relationships) {
    $customerName = $rel.Customer.DisplayName
    $includedRoleIds = $rel.AccessDetails.UnifiedRoles.RoleDefinitionId

    $roleToUse = $null
    if ($includedRoleIds -contains $CloudAppAdminRoleId) {
        $roleToUse = $CloudAppAdminRoleId
    } elseif ($includedRoleIds -contains $AppAdminRoleId) {
        $roleToUse = $AppAdminRoleId
    }

    if (-not $roleToUse) {
        Write-Host "  ✗ $customerName — no tiene Cloud/Application Administrator en esta relación. Necesita una relación GDAP NUEVA." -ForegroundColor Red
        $needsNewRequest += $customerName
        continue
    }

    # ¿Ya existe un accessAssignment para este grupo en esta relación?
    $existingAssignments = Get-MgTenantRelationshipDelegatedAdminRelationshipAccessAssignment `
        -DelegatedAdminRelationshipId $rel.Id -All -ErrorAction SilentlyContinue

    $already = $existingAssignments | Where-Object {
        $_.AccessContainer.AccessContainerId -eq $group.Id
    }

    if ($already) {
        Write-Host "  = $customerName — ya tenía la asignación, no se duplica." -ForegroundColor DarkGray
        $alreadyAssigned += $customerName
        $readyTenantIds[$customerName] = $rel.Customer.TenantId
        continue
    }

    $body = @{
        accessContainer = @{
            accessContainerId   = $group.Id
            accessContainerType = "securityGroup"
        }
        accessDetails = @{
            unifiedRoles = @(@{ roleDefinitionId = $roleToUse })
        }
    }

    try {
        New-MgTenantRelationshipDelegatedAdminRelationshipAccessAssignment `
            -DelegatedAdminRelationshipId $rel.Id `
            -BodyParameter $body | Out-Null
        Write-Host "  ✓ $customerName — asignación creada correctamente." -ForegroundColor Green
        $ready += $customerName
        $readyTenantIds[$customerName] = $rel.Customer.TenantId
    } catch {
        Write-Host "  ✗ $customerName — error al crear la asignación: $($_.Exception.Message)" -ForegroundColor Red
        $needsNewRequest += $customerName
    }
}

# ---------------------------------------------------------------------------
# 5. Resumen final
# ---------------------------------------------------------------------------
Write-Host "`n[5/5] Resumen" -ForegroundColor Cyan
Write-Host "-----------------------------------------------------------"
Write-Host "Clientes listos para leer vía Graph HOY:      $($ready.Count)" -ForegroundColor Green
$ready | ForEach-Object { Write-Host "   - $_" -ForegroundColor Green }

Write-Host "Clientes ya asignados de una corrida anterior: $($alreadyAssigned.Count)" -ForegroundColor DarkGray
$alreadyAssigned | ForEach-Object { Write-Host "   - $_" -ForegroundColor DarkGray }

Write-Host "Clientes que necesitan una relación GDAP NUEVA (acción manual, uno por uno): $($needsNewRequest.Count)" -ForegroundColor Yellow
$needsNewRequest | ForEach-Object { Write-Host "   - $_" -ForegroundColor Yellow }

Write-Host "`nGuarda el ID del grupo para configurarlo en la Azure Function: $($group.Id)" -ForegroundColor Cyan

# Deja un archivo JSON listo para pegar directo en la Application Setting
# GRAPH_CLIENT_TENANT_IDS de la Azure Function (ver security-stats-multi/index.js).
$outputFile = Join-Path $PSScriptRoot "client-tenant-ids.json"
$readyTenantIds.Values | ConvertTo-Json | Set-Content -Path $outputFile -Encoding utf8
Write-Host "IDs de tenants listos guardados en: $outputFile" -ForegroundColor Cyan
Write-Host "($($readyTenantIds.Count) tenant(s) listos para leer datos)" -ForegroundColor Cyan

Write-Host "`nListo. Corre este script de nuevo cuando agregues un cliente GDAP nuevo — es seguro repetirlo." -ForegroundColor Cyan
