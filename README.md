# Nuri Solutions – sitio web (fase 1: página + formulario general)

## Archivos
- `index.html`: página completa de una sola sección (EN/ES) con el formulario de descubrimiento.
- `assets/`: logo transparente, favicon e imagen para compartir (og-image).
- `apps-script/Code.gs`: backend en Google Apps Script (Sheets + Drive + correos).

## Cómo funciona el formulario
1. En cuanto el cliente escribe nombre, correo válido y teléfono, se guarda como interesado en la hoja **Tracker** con la etapa `Lead (form started)`, aunque no termine el formulario. Si corrige algún dato, se actualiza la misma fila.
2. Las respuestas se guardan solas en el navegador (localStorage). Si el cliente cierra la página y vuelve, las encuentra donde las dejó.
3. Al enviarlo:
   - se añade una fila en **Discovery forms**
   - el Tracker pasa a la etapa `Discovery form`
   - se crea la carpeta del cliente en Drive con el PDF
   - Nuri recibe un correo con el resumen y el PDF (si responde ese correo, le llega al cliente)
   - el cliente recibe una confirmación en su idioma con su PDF
4. Una etapa nunca retrocede de forma automática. Nuri puede cambiarla a mano con el menú desplegable.
5. Menú **Nuri Solutions** en la hoja: exporta a CSV los interesados que dieron autorización (formato listo para audiencias de Meta/Google) o todos los interesados.
6. Hay un campo oculto anti-spam (honeypot).

## Instalación (cuenta de Google nueva del sitio)
1. Crear una Google Sheet llamada `Nuri Solutions – Panel`.
2. Extensiones > Apps Script > pegar `Code.gs` y revisar `CONFIG` (correo de aviso).
3. Ejecutar `setup()` una vez y autorizar.
4. Implementar > Nueva implementación > Aplicación web. Ejecutar como: **Yo**. Acceso: **Cualquier usuario**.
5. Copiar la URL `/exec` en `NURI_CONFIG.endpoint` dentro de `index.html`.
6. Mientras `endpoint` esté vacío, el formulario funciona en **modo demo** y no envía nada.

## Publicación
- GitHub Pages: subir la carpeta a un repo, activar Pages, y en Settings > Pages poner el dominio personalizado (crea el archivo `CNAME`).
- DNS en Cloudflare: registros A a las IP de GitHub Pages (185.199.108.153, .109, .110, .111) y CNAME `www` → `ivanortizx.github.io`, en modo "DNS only" (nube gris).

## Pendiente / próximas fases
- Fotos de Nuri (sección Founder: reemplazar la "N") y portafolio real (sección Our Work).
- Contrato con firma en línea + PDF a ambos.
- Pago con Stripe (tarjeta + ACH) y webhook para marcar `Payment received`.
- Onboarding con acceso por Google o enlace mágico.
- Panel privado de Nuri.
- Antes de publicar: textos legales (política de privacidad) para la casilla de marketing, a cargo de Nuri.
