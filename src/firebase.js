import { getApps, initializeApp } from 'firebase/app'
import { browserLocalPersistence, browserPopupRedirectResolver, getAuth, indexedDBLocalPersistence, initializeAuth } from 'firebase/auth'
import { getFirestore } from 'firebase/firestore'
import { getStorage } from 'firebase/storage'
import { getFunctions } from 'firebase/functions'

const config = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID
}

export const firebaseConfigured = Object.values(config).every(Boolean)
const firebaseApp = firebaseConfigured
  ? (getApps().length ? getApps()[0] : initializeApp(config))
  : null
let auth = null
if (firebaseApp) {
  try {
    auth = initializeAuth(firebaseApp, {
      persistence: [indexedDBLocalPersistence, browserLocalPersistence],
      popupRedirectResolver: browserPopupRedirectResolver
    })
  } catch (error) {
    if (error?.code !== 'auth/already-initialized') throw error
    auth = getAuth(firebaseApp)
  }
}
export const firebaseAuth = auth
export const firebaseDb = firebaseApp ? getFirestore(firebaseApp) : null
export const firebaseStorage = firebaseApp ? getStorage(firebaseApp) : null
export const firebaseFunctions = firebaseApp ? getFunctions(firebaseApp) : null
