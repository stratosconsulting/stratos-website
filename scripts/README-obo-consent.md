# Conectar los tenants de clientes al panel — guía rápida

Esto es lo que falta después de `gdap-bulk-setup.ps1` (que ya corriste). Ese
script asignó el rol GDAP, pero **GDAP no acepta tokens de aplicación pura
contra un tenant de cliente** — solo tokens delegados app+usuario. Este paso
cierra esa parte.

## 1. Cambios de una sola vez en el App Registration (portal de Azure)

Ve a tu app "Stratos Security Panel" en Entra ID -> App registrations:

- **Authentication -> Advanced settings** -> "Allow public client flows" = **Yes**
- **Authentication -> Supported account types** -> cambia a "Accounts in any
  organizational directory (multitenant)"
- **API permissions -> Add a permission -> Microsoft Graph -> Delegated**,
  agrega: `DeviceManagementManagedDevices.Read.All`,
  `DeviceManagementConfiguration.Read.All`, `AuditLog.Read.All`,
  `SecurityAlert.Read.All`, `offline_access` -> **Grant admin consent for STRATOS**
  (esto es ADICIONAL a los permisos de tipo Application que ya tienes de la
  fase 1 — no los reemplaza, no los borres)
- **API permissions -> Add a permission -> APIs my organization uses** ->
  busca "Partner Center" -> Delegated -> `user_impersonation` -> Grant admin
  consent. Si no aparece en la lista, dímelo antes de seguir — hace falta un
  paso aparte para registrarla ahí.

## 2. Correr el script

```powershell
cd ~/Documents/stratos-website/scripts
pwsh ./obo-delegated-setup.ps1 `
    -OboUserPrincipalName "panel-obo@stratosconsultingpr.com" `
    -ClientId "<application (client) id del panel>" `
    -TenantId "<tenant id de STRATOS>"
```

Te va a mostrar una URL y un código de un solo uso — entras con la cuenta
OBO (con MFA) y listo, no hay que pegar contraseñas en la terminal.

## 3. Crear un Storage Account para guardar el refresh token (una sola vez)

Las "Managed Functions" de Static Web Apps no exponen storage propio, así
que la función necesita uno para guardar el refresh token que va rotando.
Es rápido:

- Portal de Azure -> **Create a resource -> Storage account**
- Mismo Resource Group que tu Static Web App, nombre corto (ej.
  `stratospanelstorage`), Standard / Locally-redundant storage (LRS) —
  no hace falta nada más elegante para esto, la tabla es diminuta.
- Cuando termine de crearse, entra a ese Storage Account -> **Access keys**
  (o "Security + networking -> Access keys") -> copia el
  **Connection string** de key1.

## 4. Pegar tres cosas en Azure (una sola vez)

- El **connection string** que acabas de copiar -> pégalo en la Environment
  variable `TOKEN_STORAGE_CONNECTION_STRING`.
- El **refresh token** larguísimo que imprime el script -> pégalo en
  `GRAPH_OBO_REFRESH_TOKEN`. De ahí en adelante la función lo rota sola en
  el Storage Account — no hay que tocarlo de nuevo salvo que quede vencido
  por no usarse 90+ días.
- El contenido de `client-tenant-ids.json` -> `GRAPH_CLIENT_TENANT_IDS`
  (tal cual, como una lista JSON).

## Si algún tenant falla el consentimiento

El error más probable es que el "customer id" que usa Partner Center no
coincida exactamente con el tenant id de Entra para ese cliente (pasa en
configuraciones CSP más viejas). El script sigue con los demás tenants y al
final te dice cuáles fallaron — mándame ese mensaje de error si pasa y lo
resolvemos para ese cliente específico.

## Cuándo volver a correrlo

Solo cuando agregues un cliente GDAP nuevo a `client-tenant-ids.json`
(después de correr `gdap-bulk-setup.ps1` de nuevo). Es seguro repetirlo —
no duplica consentimientos ya otorgados.
