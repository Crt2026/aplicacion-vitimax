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
3. Activa los proveedores de autenticación que quieras utilizar.

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
