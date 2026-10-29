import type { CapacitorConfig } from '@capacitor/cli'

const config: CapacitorConfig = {
  appId: 'com.vitimax.app',
  appName: 'VITIMAX',
  webDir: 'dist',
  bundledWebRuntime: false,
  android: {
    backgroundColor: '#f6f8f6'
  }
}

export default config
