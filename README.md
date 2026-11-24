# VITIMAX

Plataforma web y aplicación Android para administrar el catálogo, inventario,
pedidos, importación de productos desde Excel y el asistente de IA de VITIMAX.

## Desarrollo web

```powershell
npm install
npm run dev
```

## Configuración de Firebase

1. Copia `.env.example` como `.env`.
2. Completa las variables con la configuración de tu proyecto Firebase.
3. En Firebase activa Authentication con Email/Password, Google y Microsoft.
4. Crea Firestore y Storage, y publica sus reglas desde la raíz:

```powershell
npx firebase-tools login
npx firebase-tools use TU_PROJECT_ID
npx firebase-tools deploy --only firestore:rules,storage
```

El catálogo, perfil y preferencias se sincronizan por usuario en Firestore. Las
fotografías de productos se guardan en Storage bajo `users/{uid}/products/`.
Si Firebase no está configurado, la aplicación mantiene un modo local/invitado.

### Microsoft: por qué rechaza cuentas personales e institucionales

El botón "Continuar con Microsoft" usa el proveedor OAuth de Firebase
(`microsoft.com`). Firebase delega en un **App Registration de Azure/Entra ID**:
si ahí se cargó un Application (client) ID que no existe (por ejemplo el texto
`vitimax`), Microsoft rechaza **toda** cuenta, tanto personal (`@outlook.com`,
`@hotmail.com`) como institucional.

En este proyecto el proveedor Microsoft quedó configurado con `client_id=vitimax`,
que no es válido. Para corregirlo, en la cuenta de Azure del administrador:

1. Entra a **Azure Portal → Entra ID → Registros de aplicaciones → Nuevo registro**.
2. Nombre: `VITIMAX`.
3. En **Tipos de cuenta admitidos** elige
   **Cuentas en cualquier directorio organizativo y cuentas personales de Microsoft**
   (esto habilita `common`, necesario para aceptar ambos tipos).
4. En **URI de redireccionamiento** agrega plataforma **Web** con **ambas**
   URLs (Firebase usa la de `firebaseapp.com` en el flujo de redirección):
   - `https://vitimax-4094b.firebaseapp.com/__/auth/handler`
   - `https://vitimax-4094b.web.app/__/auth/handler`
5. En **Información general** copia el **Application (client) ID** (un UUID).
6. En **Certificados y secretos → Nuevo secreto de cliente** crea uno y copia su
   valor (solo se muestra una vez).

Luego en **Firebase Console → Authentication → Sign-in method → Microsoft**:

1. Habilita el proveedor y pega el Application ID y el secreto de Azure.
2. Guarda los cambios.
3. En **Authentication → Settings → Authorized domains** confirma que esté
   `vitimax-4094b.web.app`.

Para verificarlo sin iniciar sesión, la URL que genera Firebase debe contener un
UUID real:

```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?...&client_id=<UUID>...
```

Si sigue diciendo `client_id=vitimax`, el cambio en Firebase no se guardó.
El mismo proveedor se usa en Android e iPhone, así que no hace falta otro
cambio en el código móvil; solo recompilar los artefactos.

Mientras tanto se puede entrar con correo y contraseña o como invitado.

## Asistente de IA (BOT IA VITIMAX)

El bot funciona en **dos modos**. La app elige automáticamente el mejor
disponible y, si uno falla, prueba el siguiente antes de usar las respuestas
locales.

### Capacidades actuales

- Responde desde la primera oración, sin saludos ni relleno, y adapta el tono al usuario.
- Inyecta memoria persistente e historial completo de la conversación en el system prompt.
- Analiza Excel, CSV, PDF y texto adjunto; auditando datos reales antes de transformarlos.
- Encadena herramientas en un mismo turno (hasta 8 pasos) con `tool_choice: "auto"`.
- Genera Excel con "Resumen Ejecutivo", fórmulas reales, formato CLP, autofiltros y encabezados congelados.
- Genera PDF y Word con banner verde bosque, sumas verificadas y texto limpio (sin LaTeX ni Markdown crudo).
- Genera PowerPoint con tablas de 5 filas, columnas Nombre/Categoría/Formato/Precio/Descripción y cierre "Próximos Pasos".
- Genera páginas web completas con Tailwind, buscador, filtros, WhatsApp y vista previa en vivo.
- Descarta filas de totales y registros sin nombre real al importar, exportar o sincronizar.

### Modo directo (funciona en Android sin servidor)

Es el modo que usa el APK por defecto. La app llama directamente a OpenRouter
desde el dispositivo.

1. En `.env` completa `VITE_AI_DIRECT_KEY` con tu clave de OpenRouter.
2. Ajusta los modelos en `VITE_AI_MODELS` (el primero es el principal y los
   demás son respaldo automático):

```env
VITE_AI_DIRECT_KEY=sk-or-v1-...
VITE_AI_MODELS=openai/gpt-4o-mini,google/gemini-2.5-flash,nvidia/nemotron-3.5-super-vl:free
```

3. Vuelve a compilar (`npm run build` o `npm run android:apk`).

> **Importante:** en este modo la clave queda dentro de la app y puede ser
> extraída por alguien con conocimientos técnicos. Usa una clave de OpenRouter
> con **límite de gasto bajo** y cámbiala si la compartes públicamente. Si
> quieres ocultar la clave, usa el modo proxy.

Cada usuario también puede escribir su propia clave en
**Configuración → Clave de OpenRouter**, y así no comparte la del dueño.

### Modo proxy (recomendado para producción)

La clave se queda en el servidor y la app solo habla con tu proxy.

**En tu computadora (desarrollo):**

```powershell
Copy-Item .env.example .env
# Completa OPENROUTER_API_KEY
npm run ai-server
npm run dev
```

**En internet (para que funcione en cualquier Android):** despliega
`ai-server.cjs` en un servicio con HTTPS. El proyecto incluye `Dockerfile` y
`render.yaml` para Render, pero sirve cualquier hosting de Node.js.

1. Crea el servicio en [Render](https://render.com) apuntando a este repositorio
   (o usa el botón de Blueprint con `render.yaml`).
2. Configura las variables de entorno del servicio:
   - `OPENROUTER_API_KEY` = tu clave privada.
   - `OPENROUTER_MODEL` = `openai/gpt-4o-mini` (opcional).
   - `AI_ALLOWED_ORIGINS` = `*` o los dominios permitidos.
   - `AI_SHARED_TOKEN` = un token secreto (opcional, recomendado).
3. Copia la URL pública (por ejemplo `https://vitimax-ai.onrender.com`) en
   `VITE_AI_API_URL` y vuelve a compilar:

```env
VITE_AI_API_URL=https://vitimax-ai.onrender.com
VITE_AI_TOKEN=el-mismo-AI_SHARED_TOKEN
```

También puedes cambiar la URL y el token desde **Configuración → Servidor de
IA** dentro de la app, sin recompilar. El botón **Probar conexión IA** verifica
que el servidor responde.

### Qué puede hacer el bot

- Responder preguntas sobre el catálogo real (sin inventar productos, precios ni stock).
- Analizar imágenes, archivos de texto, CSV y JSON adjuntos.
- Generar y descargar Excel, PDF, Word y PowerPoint del catálogo.
- Crear y descargar una página web promocional de un producto.
- Ayudar con inventario, contenido para redes y trámites del negocio.
- Mantener el contexto de la conversación (historial).

Los modelos gratuitos de OpenRouter pueden responder con límites temporales
(HTTP 429). Por eso el bot intenta varios modelos en orden. Para uso intensivo,
agrega saldo a OpenRouter o usa el modo proxy con una clave propia.

## Android

El proyecto Android está en [`android/`](./android). Para regenerar el contenido
web, sincronizar Capacitor y compilar el APK:

```powershell
npm run android:apk
```

El APK queda en:

```text
android/app/build/outputs/apk/debug/app-debug.apk
```

Para instalarlo en un celular:

1. Copia el APK al teléfono (cable, Google Drive, WhatsApp, etc.).
2. Abre el archivo y permite "Instalar aplicaciones de origen desconocido".
3. Instala y abre VITIMAX.

Para publicarlo en Google Play necesitas una compilación `release` firmada con
un keystore propio. La compilación `debug` sirve para pruebas y para compartir
directamente el APK.

### Dónde se guardan los archivos en Android

La app usa un plugin nativo (`SaveToDownloadsPlugin`) que escribe el archivo
real en la carpeta pública **Descargas** del teléfono (Android 10 o superior,
sin pedir permisos). En Android 9 o anterior solicita permiso de almacenamiento
una sola vez.

- Excel: `Descargas/catalogo-vitimax-ia.xlsx`
- PDF: `Descargas/documento-vitimax.pdf`
- Word: `Descargas/documento-vitimax.docx`
- PowerPoint: `Descargas/presentacion-vitimax.pptx`
- Página web: `Descargas/pagina-<producto>.html`

Puedes abrir la app **Archivos** o **Administrador de archivos** del teléfono y
entrar a **Descargas** para comprobar que el archivo existe de verdad.

## iPhone / iPad (iOS)

El guardado ya está implementado: en iPhone el archivo se escribe en la carpeta
**Documentos** de la app y se abre la hoja de compartir de iOS para que puedas
guardarlo en **Archivos**, Drive, WhatsApp, etc. El proyecto iOS ya está creado
en [`ios/`](./ios), con los permisos de cámara y archivos configurados.

El proyecto iOS **debe compilarse en una Mac con Xcode**, porque Apple no permite
compilar apps de iPhone desde Windows. Pasos en la Mac:

1. Copia o clona este proyecto en la Mac.
2. Instala dependencias y sincroniza:

   ```bash
   npm install
   npm run ios:sync
   npx cap open ios
   ```

3. En Xcode selecciona tu equipo de firma (Signing & Capabilities) y el
   identificador `com.vitimax.app`.
4. Conecta el iPhone y ejecuta **Run**. Para instalarlo en otros iPhone usa
   TestFlight o una cuenta de desarrollador de Apple.

Para que el inicio de sesión con Google y Microsoft funcione en iPhone, agrega
el archivo `GoogleService-Info.plist` del proyecto Firebase dentro de
`ios/App/App/` (Firebase Console → Configuración del proyecto → app iOS). Sin
ese archivo, el correo y la contraseña siguen funcionando, pero los botones de
Google/Microsoft no.

Cada vez que cambie el código web, ejecuta en la Mac:

```bash
npm run ios:sync
```

## Dónde se guardan los archivos en la web

En el navegador se usa el menú de compartir del sistema (en celulares) o la
descarga normal del navegador. El archivo queda en la carpeta de descargas del
dispositivo.

## Publicar la app web (opcional)

El plan gratuito de Firebase Hosting permite publicar la versión web:

```powershell
npm run build
npx firebase-tools deploy --only hosting
```

Agrega el dominio publicado en Firebase Authentication → Settings → Authorized
domains para que funcionen Google y Microsoft.

## Seguridad

- No subas `.env` al repositorio. Está excluido por `.gitignore`.
- Nunca pongas claves privadas en variables `VITE_*` si vas a usar el modo proxy.
- En modo directo la clave sí viaja en la app: usa límites de gasto y rótala
  cuando deje de ser privada.
- Revisa diagnósticos, dosis y recomendaciones con un especialista de VITIMAX
  antes de aplicarlas en cultivos reales.
