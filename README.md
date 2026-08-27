# Stratos Consulting Group — Sitio Web (Prototipo)

Prototipo funcional del nuevo sitio de **STRATOS Consulting Group, LLC**, construido como una sola página (`index.html`, autocontenida — sin dependencias externas salvo Google Fonts) con navegación tipo SPA por hash de URL.

## Estado actual

- Diseño con branding real (logo, colores `#9FD3EC` / `#2370B3` / `#2F4E87`, tipografía Saira).
- Páginas: Inicio, Servicios, Industrias, Nosotros, Evaluación gratuita, Contacto, Soporte, Recursos — con el copy aprobado del brief de contenido.
- Panel de métricas de seguridad en el hero, con animación de conteo — actualmente usa **datos de muestra**.
- Formularios de Evaluación y Contacto: validan y muestran confirmación en el navegador, pero **no envían datos a ningún sistema todavía**.
- Sin versión en inglés (ES/EN toggle es solo visual por ahora).

## Cómo verlo

Abre `index.html` directamente en un navegador, o publícalo con GitHub Pages:

1. Sube este archivo a la raíz del repo (o a `/docs` si prefieres esa carpeta como fuente).
2. En **Settings → Pages**, selecciona la rama y carpeta como fuente.
3. GitHub te da una URL tipo `https://<usuario>.github.io/<repo>/` en un par de minutos.

## Pendientes para producción

1. **Panel de seguridad en vivo** — el frontend ya intenta hacer `fetch('/api/security-stats')` (ver comentario en el `<script>` al final del archivo con el contrato JSON esperado). Falta construir ese endpoint: una función serverless que se autentique contra Microsoft Graph (Intune para endpoints, Defender para amenazas) y tu app de tickets, cachee resultados, y los exponga sin exponer credenciales al navegador.
2. **Formularios reales** — conectar Evaluación y Contacto a Odoo CRM (o tu sistema de leads) para que efectivamente creen el lead.
3. **Sistema de tickets** — el botón "Crear un ticket de soporte" apunta a `#` por ahora; falta la URL final de tu portal.
4. **Versión en inglés** — el brief pide experiencias completas en ES y EN con selector, sin banderas.
5. **Dominio y hosting final** — decidir si esto reemplaza stratosconsultingpr.com (actualmente en Wix) o se reconstruye ahí.

## Estructura

Todo vive en un solo archivo (`index.html`) por simplicidad de prototipo: CSS y JS inline, logo embebido como base64. Cuando pase a desarrollo real, probablemente convenga separarlo en assets propios y, si se integra el panel en vivo, añadir el backend serverless aparte.
