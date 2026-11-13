import type { CapacitorConfig } from '@capacitor/cli'

const config: CapacitorConfig = {
  appId: 'com.vitimax.app',
  appName: 'VITIMAX',
  webDir: 'dist',
  bundledWebRuntime: false,
  android: {
    backgroundColor: '#f6f8f6',
    // Android 15 dibuja de borde a borde; esto evita que la barra de estado
    // y la barra de navegación tapen el contenido de la app.
    adjustMarginsForEdgeToEdge: 'auto'
  },
  plugins: {
    FirebaseAuthentication: {
      providers: ['google.com', 'microsoft.com']
    }
  }
}

export default config
