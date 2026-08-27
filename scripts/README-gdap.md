# GDAP bulk setup — guía rápida

Este script (`gdap-bulk-setup.ps1`) automatiza la parte que SÍ se puede automatizar
de conectar el panel a todos tus clientes MSP. La parte que Microsoft no permite
automatizar (aprobar una relación GDAP nueva cuando falta el rol correcto) queda
reportada al final, cliente por cliente, para que la resuelvas tú una sola vez.

## Antes de correrlo (una sola vez)

1. **Cuenta OBO con MFA.** Crea (si no existe) una cuenta dedicada, por ejemplo
   `panel-obo@stratosconsultingpr.com`, y actívale MFA. No uses tu cuenta personal.
2. **PowerShell + módulo de Graph** en tu Mac:
   ```bash
   # Si no tienes PowerShell instalado:
   brew install --cask powershell

   # Dentro de pwsh:
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```
3. Ten a mano tu login de administrador del tenant de **partner** de STRATOS
   (no el de un cliente).

## Cómo correrlo

```powershell
cd ~/Documents/stratos-website/scripts
pwsh ./gdap-bulk-setup.ps1 -OboUserPrincipalName "panel-obo@stratosconsultingpr.com"
```

Te va a pedir iniciar sesión (con MFA) — es normal, es tu propia sesión de admin,
no la de la cuenta OBO.

## Qué esperar al final

El script imprime tres listas:

- **Clientes listos hoy** — ya quedaron conectados, no hay que hacer nada más
  con ellos.
- **Ya asignados de una corrida anterior** — si vuelves a correrlo, estos no se
  duplican.
- **Clientes que necesitan una relación GDAP nueva** — aquí sí hace falta un
  paso manual por cliente: entrar a Partner Center, crear una relación GDAP
  nueva para ese cliente que SÍ incluya el rol "Cloud Application Administrator"
  o "Application Administrator", y esperar a que el cliente (o tú, si tienes
  permiso delegado) la apruebe. Una vez aprobada, vuelves a correr este mismo
  script y ese cliente pasa a la primera lista automáticamente.

También te deja un archivo `client-tenant-ids.json` en esta misma carpeta con
los IDs de tenant de todos los clientes ya listos — ese archivo es lo próximo
que conectamos a la Azure Function para que empiece a sumar sus datos.

## Es seguro repetirlo

Corre el script cada vez que:
- Apruebes una relación GDAP nueva para un cliente que había quedado pendiente.
- Firmes un cliente MSP nuevo.

No duplica el grupo ni las asignaciones ya hechas.
