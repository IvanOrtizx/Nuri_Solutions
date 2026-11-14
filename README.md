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

## Onboarding (`/onboarding/`)
Página privada para clientes que ya firmaron y pagaron (etapa `Payment received` o posterior en el Tracker).

**Cómo entra el cliente**
- Escribe su correo y recibe un **enlace privado** que dura 30 días. No necesita contraseña.
- O Nuri, desde la hoja, selecciona la fila del cliente en Tracker y usa **Nuri Solutions > Send onboarding link to selected client**. Si el cliente aún no está en `Payment received`, el script pregunta si quiere marcarlo así.
- Opcional: botón "Continuar con Google". Para activarlo se crea un ID de cliente OAuth (tipo Web) en Google Cloud, con los orígenes autorizados `https://camott.com` y `https://www.nurisolutions.si`. El ID se pega en `CONFIG.GOOGLE_CLIENT_ID` (Code.gs) y en `ONB_CONFIG.googleClientId` (onboarding/index.html).

**Qué guarda**
- Las respuestas se guardan solas (en el navegador y en la pestaña **Onboarding**), así que el cliente puede seguir otro día o desde otro dispositivo.
- Los archivos se suben uno por uno, máximo 20 MB cada uno, a `Carpeta del cliente/Onboarding/<categoría>`. Para videos pesados el cliente puede pegar enlaces.
- Al enviar se genera el PDF y se envía a Nuri y al cliente, y el Tracker pasa a `Onboarding complete`. El cliente puede volver y actualizar sus respuestas.

**Después de cambiar Code.gs**: ejecutar `setup()` de nuevo y publicar una **nueva versión** de la implementación existente para que la URL `/exec` no cambie.

## Pendiente / próximas fases
- Fotos de Nuri (sección Founder: reemplazar la "N") y portafolio real (sección Our Work).
- Contrato con firma en línea + PDF a ambos.
- Pago con Stripe (tarjeta + ACH) y webhook para marcar `Payment received`.
- Panel privado de Nuri.
- Antes de publicar: textos legales (política de privacidad) para la casilla de marketing, a cargo de Nuri.
