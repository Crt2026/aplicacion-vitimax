# VITIMAX

Plataforma web y aplicación Android para administrar el catálogo, inventario,
pedidos, importación de productos desde Excel y asistencia para VITIMAX.

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

5. Para el asistente, crea el secreto del backend y despliega las funciones:

```powershell
npx firebase-tools functions:secrets:set GEMINI_API_KEY
npm --prefix functions install
npx firebase-tools deploy --only functions
```

El catálogo, perfil y preferencias se sincronizan por usuario en Firestore. Las
fotografías de productos se guardan en Storage bajo `users/{uid}/products/`.
Si Firebase no está configurado, la aplicación mantiene un modo local/invitado.
El modo local también permite importar Excel, adjuntar archivos al asistente,
leer archivos de texto/CSV/JSON y descargar páginas HTML generadas desde el
catálogo. Los análisis avanzados de PDF, Word e imágenes con Gemini requieren
Cloud Functions y, por tanto, el plan Blaze.

Para Microsoft, registra una aplicación en Microsoft Entra ID, copia el Client
ID en el proveedor Microsoft de Firebase y agrega el dominio de producción a
los dominios autorizados de Firebase Authentication. Nunca pongas secretos de
Microsoft o Gemini en `VITE_*` ni en el frontend.

No subas `.env` al repositorio. Ese archivo está excluido por `.gitignore`.

## Android

El proyecto Android está en [`android/`](./android). Para regenerar el contenido
web y sincronizar Capacitor:

```powershell
npm run build
npx cap sync android
```

El archivo `VITIMAX.apk` incluido en este repositorio corresponde a una versión
debug para pruebas.
