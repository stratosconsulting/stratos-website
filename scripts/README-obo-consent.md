# Conectar los tenants de clientes al panel — guía rápida

Esto es lo que falta después de `gdap-bulk-setup.ps1` (que ya corriste). Ese
script asignó el rol GDAP, pero **GDAP no acepta tokens de aplicación pura
contra un tenant de cliente** — solo tokens delegados app+usuario. Este paso
cierra esa parte.

> **Ya corriste esto una vez y funcionó (8 de 9 clientes).** Si solo vienes
> a agregar el permiso nuevo `Device.Read.All` (para contar todos los
> dispositivos de Microsoft 365/Entra, no solo los de Intune), ve directo a
> la sección **"Actualización: Device.Read.All"** al final — no hace falta
> repetir el Storage Account ni las variables de Azure, solo el permiso y
> volver a correr el script.

## 1. Cambios de una sola vez en el App Registration (portal de Azure)

Ve a tu app "Stratos Security Panel" en Entra ID -> App registrations:

- **Authentication -> Advanced settings** -> "Allow public client flows" = **Yes**
- **Authentication -> Supported account types** -> cambia a "Accounts in any
  organizational directory (multitenant)"
- **API permissions -> Add a permission -> Microsoft Graph -> Delegated**,
  agrega: `Device.Read.All`, `DeviceManagementManagedDevices.Read.All`,
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
(después de correr `gdap-bulk-setup.ps1` de nuevo), o cuando agregues un
permiso Graph nuevo (ver abajo). Es seguro repetirlo — no duplica
consentimientos ya otorgados.

## Actualización: Device.Read.All

Para contar TODOS los dispositivos que Microsoft 365/Entra conoce (no solo
los inscritos en Intune, ya que no todos tus clientes lo tienen), la función
ahora también necesita el permiso `Device.Read.All`. Dos pasos:

1. **En Entra ID -> App registrations -> "Stratos Security Panel" ->
   API permissions**, agrega `Device.Read.All` DOS veces:
   - **Microsoft Graph -> Delegated -> `Device.Read.All`** (para los
     tenants de clientes, vía OBO)
   - **Microsoft Graph -> Application -> `Device.Read.All`** (para tu
     propio tenant, vía la app-only credential que ya tenías desde fase 1)
   -> **Grant admin consent for STRATOS** después de agregar ambas.

2. **Vuelve a correr `obo-delegated-setup.ps1`** exactamente como la
   primera vez (mismos parámetros). El consentimiento que ya le diste a los
   8 clientes NO incluye este permiso nuevo automáticamente — hay que
   re-otorgarlo. Es seguro, no duplica nada, solo actualiza el alcance.

No hace falta tocar el Storage Account ni las Application settings de
nuevo — esas ya están bien.

## Actualización: IdentityRiskEvent.Read.All (amenazas bloqueadas)

Para que "Amenazas bloqueadas" cuente también los inicios de sesión
riesgosos que detecta Entra ID Protection (no solo las alertas formales de
Defender), la función ahora también necesita `IdentityRiskEvent.Read.All`.
Mismos dos pasos que arriba:

1. **En Entra ID -> App registrations -> "Stratos Security Panel" ->
   API permissions**, agrega `IdentityRiskEvent.Read.All` DOS veces:
   - **Microsoft Graph -> Delegated -> `IdentityRiskEvent.Read.All`**
     (para los tenants de clientes, vía OBO)
   - **Microsoft Graph -> Application -> `IdentityRiskEvent.Read.All`**
     (para tu propio tenant)
   -> **Grant admin consent for STRATOS** después de agregar ambas.

2. **Vuelve a correr `obo-delegated-setup.ps1`** exactamente como siempre.

Nota: este dato requiere que el tenant del cliente tenga Entra ID P2 (viene
incluido en Microsoft 365 E5, o se puede comprar aparte). Los clientes que
no lo tengan simplemente no aportan nada a esta parte del número — no falla
nada, el panel sigue funcionando igual con solo las alertas de Defender
para esos tenants.

## Actualización: UserAuthenticationMethod.Read.All (reporte de brechas MFA)

El reporte `gdap-gap-report.ps1` usa `reports/authenticationMethods/userRegistrationDetails`
para ver quién no tiene MFA — pero ese endpoint requiere Entra ID P1/P2 en
el tenant del cliente, y varios de tus clientes no lo tienen (por eso salió
403 en algunos). Este permiso nuevo habilita un método alterno (leer los
métodos de autenticación registrados usuario por usuario, sin depender de
ningún reporte premium) que el script usa automáticamente como respaldo
cuando el reporte normal falla.

1. **En Entra ID -> App registrations -> "Stratos Security Panel" ->
   API permissions**, agrega DOS permisos (ambos Delegated, ninguno
   necesita la versión Application — este dato solo se usa desde
   `gdap-gap-report.ps1`, que corre localmente con tu login delegado,
   nunca desde la función pública del panel):
   - **Microsoft Graph -> Delegated -> `UserAuthenticationMethod.Read.All`**
     (leer los métodos de autenticación de un usuario)
   - **Microsoft Graph -> Delegated -> `User.Read.All`** (listar los
     usuarios del tenant — sin esto el script no puede ni empezar a
     recorrer la lista de usuarios)
   -> **Grant admin consent for STRATOS** después de agregar ambos.

2. **Vuelve a correr `obo-delegated-setup.ps1`** exactamente como siempre.
