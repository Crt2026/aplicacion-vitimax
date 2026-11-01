import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import { Document, HeadingLevel, Packer, Paragraph } from 'docx'
import pptxgen from 'pptxgenjs'
import { Html5Qrcode } from 'html5-qrcode'
import { GoogleAuthProvider, OAuthProvider, createUserWithEmailAndPassword, getRedirectResult, onAuthStateChanged, sendEmailVerification, sendPasswordResetEmail, signInWithCredential, signInWithEmailAndPassword, signInWithPopup, signOut, updateProfile } from 'firebase/auth'
import { FirebaseAuthentication } from '@capacitor-firebase/authentication'
import { doc, onSnapshot, setDoc } from 'firebase/firestore'
import { getDownloadURL, ref, uploadBytes } from 'firebase/storage'
import { httpsCallable } from 'firebase/functions'
import {
  ArrowDownToLine, ArrowUpToLine, Bell, Bot, Boxes, ChevronRight,
  CircleHelp, ClipboardList, FileSpreadsheet, Filter,
  LayoutDashboard, Leaf, Menu, MessageSquare, MoreHorizontal, PackageCheck,
  Plus, Search, Settings2, ShoppingBag, Sparkles, Sprout, Tags, Trash2, Upload,
  UserRound, Users, X, Zap, QrCode, CalendarDays
} from 'lucide-react'
import './styles.css'
import { firebaseAuth, firebaseConfigured, firebaseDb, firebaseFunctions, firebaseStorage } from './firebase'

const STORAGE_KEY = 'vitimax-products-v1'
const PROFILE_KEY = 'vitimax-profile-v1'
const SETTINGS_KEY = 'vitimax-settings-v1'
const fieldMap = {
  name: ['nombre', 'producto', 'name'],
  category: ['categoria', 'categoría', 'category'],
  price: ['precio', 'price'],
  previousPrice: ['precio anterior', 'precio_anterior', 'previousprice'],
  stock: ['stock', 'existencias', 'cantidad'],
  format: ['formato', 'tamaño', 'tamano', 'format'],
  description: ['descripcion', 'descripción', 'description'],
  usage: ['modo de uso', 'uso', 'usage'],
  recommendations: ['recomendaciones', 'recommendations'],
  features: ['caracteristicas', 'características', 'features'],
  image: ['fotografia', 'fotografía', 'imagen', 'image', 'url imagen']
}

const emptyProduct = {
  name: '', category: '', price: '', previousPrice: '', stock: '',
  format: '', description: '', usage: '', recommendations: '', features: '', image: ''
}

function normalize(value) {
  return String(value ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
}

function makeId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID()
  return `vitimax-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
}

function readProduct(row) {
  const keys = Object.keys(row)
  const find = (aliases) => {
    const key = keys.find((candidate) => aliases.includes(normalize(candidate)))
    return key ? row[key] : ''
  }
  return Object.fromEntries(Object.entries(fieldMap).map(([field, aliases]) => [field, find(aliases)]))
}

function money(value) {
  if (value === '' || value === null || value === undefined || Number.isNaN(Number(value))) return 'Pendiente'
  return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 }).format(Number(value))
}

function imageFileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(new Error('No se pudo leer la imagen.'))
    reader.readAsDataURL(file)
  })
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function escapeHtml(value) {
  return String(value || '').replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]))
}

function authErrorMessage(code) {
  return code === 'auth/invalid-credential' || code === 'auth/wrong-password' || code === 'auth/user-not-found' ? 'Correo o contraseña incorrectos.'
    : code === 'auth/email-already-in-use' ? 'Ese correo ya está registrado.'
      : code === 'auth/weak-password' ? 'La contraseña debe tener al menos 6 caracteres.'
        : code === 'auth/invalid-email' ? 'Escribe un correo electrónico válido.'
          : code === 'auth/popup-closed-by-user' ? 'Se cerró la ventana de inicio de sesión.'
            : code === 'auth/unauthorized-domain' ? 'Firebase no autoriza este dominio. Agrega el dominio actual en Authentication → Settings → Authorized domains.'
              : code === 'auth/operation-not-allowed' ? 'Activa este proveedor en Firebase → Authentication → Sign-in method.'
                  : code === 'auth/account-exists-with-different-credential' ? 'Ese correo ya existe con otro método de acceso.'
                    : code === 'auth/network-request-failed' ? (navigator.onLine
                      ? 'El celular tiene internet, pero Firebase no respondió. Revisa fecha y hora automáticas, VPN/bloqueadores e inténtalo otra vez.'
                      : 'El celular aparece sin conexión. Activa Wi-Fi o datos móviles e inténtalo otra vez.')
                      : code === 'auth/too-many-requests' ? 'Demasiados intentos. Espera unos minutos y vuelve a intentarlo.'
                        : code === 'auth/user-disabled' ? 'Esta cuenta fue deshabilitada en Firebase. Contacta al administrador.'
                        : code === 'auth/popup-blocked' ? 'El celular bloqueó la ventana de Google/Microsoft. Permite ventanas emergentes e inténtalo otra vez.'
                          : code === 'auth/popup-timeout' ? 'El inicio con Google/Microsoft tardó demasiado. Comprueba internet e inténtalo otra vez.'
                          : code === 'auth/cancelled-popup-request' || code === 'auth/redirect-cancelled-by-user' ? 'Se canceló el inicio de sesión.'
                            : code === 'auth/invalid-api-key' ? 'La clave de Firebase no es válida. Revisa la configuración del proyecto.'
                              : code === 'auth/internal-error' ? 'Firebase rechazó temporalmente el acceso. Revisa proveedores y dominios autorizados.'
                                : code === 'auth/invalid-continue-uri' || code === 'auth/unauthorized-continue-uri' ? 'Firebase bloqueó el regreso de Google/Microsoft. Agrega el dominio de la app en Authentication → Settings → Authorized domains.'
                                : code ? `Firebase no pudo completar el acceso (${code}). Revisa la configuración y vuelve a intentar.` : 'Firebase no devolvió un error identificable. Cierra y abre la aplicación nuevamente.'
}

function isNativeApp() {
  return Boolean(globalThis.Capacitor?.isNativePlatform?.())
}

async function signInWithNativeProvider(providerId) {
  const result = providerId === 'google.com'
    ? await FirebaseAuthentication.signInWithGoogle()
    : await FirebaseAuthentication.signInWithMicrosoft()
  if (!result.credential) throw new Error('El proveedor no devolvió credenciales.')
  const credential = providerId === 'google.com'
    ? GoogleAuthProvider.credential(result.credential.idToken, result.credential.accessToken)
    : new OAuthProvider(providerId).credential(result.credential.idToken, result.credential.accessToken)
  return signInWithCredential(firebaseAuth, credential)
}

function AuthGate() {
  const [user, setUser] = useState(null)
  const [guest, setGuest] = useState(() => sessionStorage.getItem('vitimax-guest') === 'true')
  const [authLoading, setAuthLoading] = useState(Boolean(firebaseAuth))
  const [authError, setAuthError] = useState('')
  useEffect(() => {
    if (!firebaseAuth) return undefined
    const loadingTimeout = setTimeout(() => {
      setAuthLoading((current) => {
        if (current) setAuthError('Firebase tardó demasiado en responder. Revisa internet, fecha/hora automática y vuelve a intentar.')
        return false
      })
    }, 12000)
    const unsubscribe = onAuthStateChanged(
      firebaseAuth,
      (nextUser) => { setUser(nextUser); setAuthLoading(false) },
      (error) => { setAuthError(authErrorMessage(error.code)); setAuthLoading(false) }
    )
    getRedirectResult(firebaseAuth).catch((error) => { setAuthError(authErrorMessage(error.code)); setAuthLoading(false) })
    return () => { clearTimeout(loadingTimeout); unsubscribe() }
  }, [])
  if (authLoading) return <div className="auth-loading"><Leaf size={32} /><strong>VITIMAX</strong><span>Verificando sesión...</span></div>
  if (guest) return <App guest onSignOut={() => { sessionStorage.removeItem('vitimax-guest'); setGuest(false) }} />
  if (!firebaseConfigured) return <AuthScreenV2 configured={false} onGuest={() => { sessionStorage.setItem('vitimax-guest', 'true'); setGuest(true) }} />
  if (!user) return <AuthScreenV2 configured initialError={authError} onAuthenticated={setUser} onGuest={() => { sessionStorage.setItem('vitimax-guest', 'true'); setGuest(true) }} />
  return <App user={user} onSignOut={() => signOut(firebaseAuth)} />
}

function AuthScreen({ configured, initialError = '', onAuthenticated, onGuest }) {
  const [mode, setMode] = useState('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [firstName, setFirstName] = useState('')
  const [lastName, setLastName] = useState('')
  const [birthDate, setBirthDate] = useState('')
  const [error, setError] = useState('')
  const [notice, setNotice] = useState('')
  const [busy, setBusy] = useState(false)
  useEffect(() => { if (initialError) setError(initialError) }, [initialError])
  async function submit(event) {
    event.preventDefault()
    setError('')
    if (!firebaseAuth) return
    setBusy(true)
    try {
      if (mode === 'reset') {
        await sendPasswordResetEmail(firebaseAuth, email.trim())
        setNotice('Te enviamos un enlace para cambiar tu contraseña. Revisa tu correo y la carpeta de spam.')
        setMode('login')
        return
      }
      if (mode === 'register' && (!firstName.trim() || !lastName.trim() || !birthDate)) {
        setError('Completa nombre, apellido y fecha de nacimiento.')
        return
      }
      const normalizedEmail = email.trim().toLowerCase()
      const result = mode === 'login'
        ? await signInWithEmailAndPassword(firebaseAuth, normalizedEmail, password)
        : await createUserWithEmailAndPassword(firebaseAuth, normalizedEmail, password)
      if (mode === 'register') {
        const fullName = `${firstName.trim()} ${lastName.trim()}`
        await updateProfile(result.user, { displayName: fullName })
        await setDoc(doc(firebaseDb, 'users', result.user.uid), { profile: { name: fullName, firstName: firstName.trim(), lastName: lastName.trim(), email: normalizedEmail, birthDate, phone: '', photo: '' } }, { merge: true })
        await sendEmailVerification(result.user)
        setNotice('Cuenta creada. Te enviamos un correo de verificación.')
      }
      onAuthenticated?.(result.user)
    } catch (authError) {
      setError(authErrorMessage(authError.code))
    } finally { setBusy(false) }
  }
  async function providerLogin(provider) {
    setError('')
    if (!firebaseAuth) return
    setBusy(true)
    try {
      provider.setCustomParameters?.({ prompt: 'select_account' })
      if (isNativeApp()) {
        await signInWithNativeProvider(provider.providerId)
      } else {
        await Promise.race([
          signInWithPopup(firebaseAuth, provider),
          new Promise((_, reject) => setTimeout(() => reject({ code: 'auth/popup-timeout' }), 20000))
        ])
      }
    } catch (authError) { setError(authErrorMessage(authError.code)) }
    finally { setBusy(false) }
  }
  return <main className="auth-screen"><section className="auth-card"><div className="auth-brand"><div className="brand-mark"><Leaf size={23} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div>{!configured ? <div className="auth-setup"><CircleHelp size={28} /><h1>Accede a VITIMAX</h1><p>El acceso de administrador necesita configurar Firebase. Mientras tanto puedes entrar como invitado para consultar la aplicación.</p><small>El modo invitado es solo de lectura: no permite editar productos, importar Excel ni cambiar la configuración.</small><button className="btn primary auth-guest" onClick={onGuest}>Entrar como invitado</button></div> : <><p className="eyebrow">ACCESO DE ADMINISTRADOR</p><h1>{mode === 'login' ? 'Iniciar sesión' : 'Crear cuenta'}</h1><p className="auth-muted">Accede al centro de gestión de VITIMAX.</p><div className="social-actions"><button className="social-btn google" disabled={busy} onClick={() => providerLogin(new GoogleAuthProvider())}><GoogleBrandIcon /> Continuar con Google</button><button className="social-btn microsoft" disabled={busy} onClick={() => providerLogin(new OAuthProvider('microsoft.com'))}><MicrosoftBrandIcon /> Continuar con Microsoft</button></div><div className="auth-divider"><span>o con correo</span></div><form onSubmit={submit} className="auth-form"><label>Correo electrónico<input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" /></label><label>Contraseña<input type="password" required minLength="6" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete={mode === 'login' ? 'current-password' : 'new-password'} /></label>{error && <p className="auth-error">{error}</p>}<button className="btn primary auth-submit" disabled={busy}>{busy ? 'Conectando...' : mode === 'login' ? 'Iniciar sesión' : 'Registrarme'}</button></form><button className="auth-switch" onClick={() => { setMode(mode === 'login' ? 'register' : 'login'); setError('') }}>{mode === 'login' ? 'Crear cuenta de administrador' : 'Ya tengo una cuenta'}</button><button className="auth-guest-link" onClick={onGuest}>Continuar como invitado</button></>}</section></main>
}

function AuthScreenV2({ configured, initialError = '', onAuthenticated, onGuest }) {
  const [mode, setMode] = useState('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [firstName, setFirstName] = useState('')
  const [lastName, setLastName] = useState('')
  const [birthDate, setBirthDate] = useState('')
  const [error, setError] = useState(initialError)
  const [notice, setNotice] = useState('')
  const [busy, setBusy] = useState(false)
  useEffect(() => { if (initialError) setError(initialError) }, [initialError])

  async function submit(event) {
    event.preventDefault()
    setError('')
    setNotice('')
    if (!firebaseAuth) return
    setBusy(true)
    try {
      if (mode === 'reset') {
        await sendPasswordResetEmail(firebaseAuth, email.trim())
        setNotice('Te enviamos un enlace para cambiar tu contraseña. Revisa tu correo y la carpeta de spam.')
        setMode('login')
        return
      }
      if (mode === 'register' && (!firstName.trim() || !lastName.trim() || !birthDate)) {
        setError('Completa nombre, apellido y fecha de nacimiento.')
        return
      }
      const normalizedEmail = email.trim().toLowerCase()
      const result = mode === 'login'
        ? await signInWithEmailAndPassword(firebaseAuth, normalizedEmail, password)
        : await createUserWithEmailAndPassword(firebaseAuth, normalizedEmail, password)
      if (mode === 'register') {
        const cleanFirstName = firstName.trim()
        const cleanLastName = lastName.trim()
        const fullName = `${cleanFirstName} ${cleanLastName}`
        await updateProfile(result.user, { displayName: fullName })
        if (firebaseDb) {
          setDoc(doc(firebaseDb, 'users', result.user.uid), { profile: { name: fullName, firstName: cleanFirstName, lastName: cleanLastName, email: normalizedEmail, birthDate, phone: '', photo: '' } }, { merge: true }).catch(() => undefined)
        }
        try {
          await sendEmailVerification(result.user)
          setNotice('Cuenta creada. Te enviamos un correo de verificación.')
        } catch {
          setNotice('Cuenta creada. No pudimos enviar el correo de verificación todavía; podrás solicitarlo después.')
        }
      }
      onAuthenticated?.(result.user)
    } catch (authError) {
      setError(authErrorMessage(authError.code))
    } finally { setBusy(false) }
  }

  async function providerLogin(provider) {
    setError('')
    if (!firebaseAuth) return
    setBusy(true)
    try {
      provider.setCustomParameters?.({ prompt: 'select_account' })
      if (isNativeApp()) {
        await signInWithNativeProvider(provider.providerId)
      } else {
        await Promise.race([
          signInWithPopup(firebaseAuth, provider),
          new Promise((_, reject) => setTimeout(() => reject({ code: 'auth/popup-timeout' }), 20000))
        ])
      }
    } catch (authError) { setError(authErrorMessage(authError.code)); setBusy(false) }
  }

  if (!configured) return <main className="auth-screen"><section className="auth-card"><div className="auth-brand"><div className="brand-mark"><Leaf size={23} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div><div className="auth-setup"><CircleHelp size={28} /><h1>Accede a VITIMAX</h1><p>El acceso de administrador necesita configurar Firebase. Mientras tanto puedes entrar como invitado para consultar la aplicación.</p><small>El modo invitado es solo de lectura.</small><button className="btn primary auth-guest" onClick={onGuest}>Entrar como invitado</button></div></section></main>

  const register = mode === 'register'
  const reset = mode === 'reset'
  const today = new Date().toISOString().slice(0, 10)
  return <main className="auth-screen"><section className="auth-card"><div className="auth-brand"><div className="brand-mark"><Leaf size={23} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div><p className="eyebrow">ACCESO DE ADMINISTRADOR</p><h1>{reset ? 'Cambiar contraseña' : register ? 'Crear cuenta' : 'Iniciar sesión'}</h1><p className="auth-muted">{reset ? 'Te enviaremos un enlace seguro a tu correo.' : 'Accede al centro de gestión de VITIMAX.'}</p>{!reset && <><div className="social-actions"><button type="button" className="social-btn google" disabled={busy} onClick={() => providerLogin(new GoogleAuthProvider())}><GoogleBrandIcon /> Continuar con Google</button><button type="button" className="social-btn microsoft" disabled={busy} onClick={() => providerLogin(new OAuthProvider('microsoft.com'))}><MicrosoftBrandIcon /> Continuar con Microsoft</button></div><div className="auth-divider"><span>o con correo</span></div></>}<form onSubmit={submit} className="auth-form">{register && <div className="auth-form-row"><label>Nombre<input type="text" required value={firstName} onChange={(e) => setFirstName(e.target.value)} autoComplete="given-name" /></label><label>Apellido<input type="text" required value={lastName} onChange={(e) => setLastName(e.target.value)} autoComplete="family-name" /></label></div>}<label>Correo electrónico<input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" /></label>{register && <label>Fecha de nacimiento<BirthdayPicker value={birthDate} onChange={setBirthDate} maxDate={today} /><small className="field-help">Calendario VITIMAX 2026: elige directamente el mes, año y día. No puede ser una fecha futura.</small></label>}{!reset && <label>Contraseña<input type="password" required minLength="6" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete={register ? 'new-password' : 'current-password'} /></label>}{error && <p className="auth-error">{error}</p>}{notice && <p className="auth-success">{notice}</p>}<button className="btn primary auth-submit" disabled={busy}>{busy ? 'Procesando...' : reset ? 'Enviar enlace' : register ? 'Crear mi cuenta' : 'Iniciar sesión'}</button></form>{!reset && !register && <button type="button" className="auth-forgot" onClick={() => { setMode('reset'); setError(''); setNotice('') }}>¿Olvidaste tu contraseña?</button>}<button type="button" className="auth-switch" onClick={() => { setMode(reset || register ? 'login' : 'register'); setError(''); setNotice('') }}>{reset || register ? 'Volver a iniciar sesión' : 'Crear cuenta de administrador'}</button><button type="button" className="auth-guest-link" onClick={onGuest}>Continuar como invitado</button></section></main>
}

function BirthdayPicker({ value, onChange, maxDate }) {
  const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
  const weekDays = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
  const minYear = 1900
  const maxYear = Number(maxDate.slice(0, 4))
  const parseDate = (date) => { const [year, month, day] = (date || '').split('-').map(Number); return year && month && day ? new Date(year, month - 1, day) : new Date(maxYear, 0, 1) }
  const [open, setOpen] = useState(false)
  const [viewDate, setViewDate] = useState(() => parseDate(value))
  const pickerRef = useRef(null)
  useEffect(() => {
    const close = (event) => { if (pickerRef.current && !pickerRef.current.contains(event.target)) setOpen(false) }
    document.addEventListener('mousedown', close)
    document.addEventListener('touchstart', close)
    return () => { document.removeEventListener('mousedown', close); document.removeEventListener('touchstart', close) }
  }, [])
  const selected = value ? parseDate(value) : null
  const firstDay = new Date(viewDate.getFullYear(), viewDate.getMonth(), 1)
  const offset = (firstDay.getDay() + 6) % 7
  const daysInMonth = new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 0).getDate()
  const cells = Array.from({ length: 42 }, (_, index) => {
    const day = index - offset + 1
    return day > 0 && day <= daysInMonth ? new Date(viewDate.getFullYear(), viewDate.getMonth(), day) : null
  })
  const formatISO = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
  const displayValue = selected ? `${String(selected.getDate()).padStart(2, '0')} de ${monthNames[selected.getMonth()]} de ${selected.getFullYear()}` : 'Selecciona tu fecha de nacimiento'
  const changeMonth = (amount) => setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() + amount, 1))
  const chooseMonth = (month) => setViewDate(new Date(viewDate.getFullYear(), Number(month), 1))
  const chooseYear = (year) => setViewDate(new Date(Number(year), viewDate.getMonth(), 1))
  const isDisabled = (date) => { const iso = formatISO(date); return iso < '1900-01-01' || iso > maxDate }
  return <div className="birthday-picker" ref={pickerRef}><button type="button" className={`date-picker-field ${value ? 'has-value' : ''}`} onClick={() => setOpen((current) => !current)} aria-haspopup="dialog" aria-expanded={open}><span>{displayValue}</span><CalendarDays size={18} /></button>{open && <div className="birthday-calendar" role="dialog" aria-label="Calendario de fecha de nacimiento"><div className="calendar-header"><button type="button" className="calendar-arrow" onClick={() => changeMonth(-1)} disabled={viewDate.getFullYear() <= minYear && viewDate.getMonth() === 0} aria-label="Mes anterior">‹</button><div className="calendar-selects"><select value={viewDate.getMonth()} onChange={(event) => chooseMonth(event.target.value)} aria-label="Mes">{monthNames.map((month, index) => <option key={month} value={index}>{month}</option>)}</select><select value={viewDate.getFullYear()} onChange={(event) => chooseYear(event.target.value)} aria-label="Año">{Array.from({ length: maxYear - minYear + 1 }, (_, index) => maxYear - index).map((year) => <option key={year} value={year}>{year}</option>)}</select></div><button type="button" className="calendar-arrow" onClick={() => changeMonth(1)} disabled={viewDate.getFullYear() >= maxYear && viewDate.getMonth() === 11} aria-label="Mes siguiente">›</button></div><div className="calendar-weekdays">{weekDays.map((day) => <span key={day}>{day}</span>)}</div><div className="calendar-grid">{cells.map((date, index) => date ? <button type="button" key={index} className={`${selected && formatISO(selected) === formatISO(date) ? 'selected' : ''} ${formatISO(date) === maxDate ? 'today' : ''}`} disabled={isDisabled(date)} onClick={() => { onChange(formatISO(date)); setViewDate(date); setOpen(false) }}>{date.getDate()}</button> : <span key={index} />)}</div><p className="calendar-note">Selecciona una fecha entre 1900 y {maxYear}.</p></div>}</div>
}

function GoogleBrandIcon() { return <svg className="provider-logo" viewBox="0 0 24 24" aria-hidden="true"><path fill="#4285F4" d="M21.35 12.27c0-.72-.06-1.42-.18-2.09H12v3.96h5.24a4.48 4.48 0 0 1-1.94 2.94v2.45h3.14c1.84-1.69 2.91-4.18 2.91-7.26Z"/><path fill="#34A853" d="M12 21.7c2.63 0 4.84-.87 6.45-2.36l-3.14-2.45c-.87.58-1.98.93-3.31.93-2.54 0-4.7-1.72-5.47-4.03H3.29v2.53A9.74 9.74 0 0 0 12 21.7Z"/><path fill="#FBBC05" d="M6.53 13.79A5.85 5.85 0 0 1 6.23 12c0-.62.11-1.23.3-1.79V7.68H3.29A9.74 9.74 0 0 0 2.25 12c0 1.57.38 3.05 1.04 4.32l3.24-2.53Z"/><path fill="#EA4335" d="M12 6.18c1.43 0 2.71.49 3.72 1.46l2.79-2.79C16.83 3.28 14.63 2.3 12 2.3a9.74 9.74 0 0 0-8.71 5.38l3.24 2.53C7.3 7.9 9.46 6.18 12 6.18Z"/></svg> }
function MicrosoftBrandIcon() { return <svg className="provider-logo ms" viewBox="0 0 24 24" aria-hidden="true"><path fill="#F25022" d="M2 2h9.5v9.5H2z"/><path fill="#7FBA00" d="M12.5 2H22v9.5h-9.5z"/><path fill="#00A4EF" d="M2 12.5h9.5V22H2z"/><path fill="#FFB900" d="M12.5 12.5H22V22h-9.5z"/></svg> }

function App({ user, onSignOut, guest = false }) {
  const [products, setProducts] = useState(() => JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'))
  const [profile, setProfile] = useState(() => {
    const saved = JSON.parse(localStorage.getItem(PROFILE_KEY) || '{}')
    return { name: saved.name || user?.displayName || (guest ? 'Invitado' : 'Administrador'), email: saved.email || user?.email || '', phone: saved.phone || '', photo: saved.photo || user?.photoURL || '' }
  })
  const [settings, setSettings] = useState(() => JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{"notifications":true,"lowStock":5,"currency":"ARS"}'))
  const [active, setActive] = useState('Resumen')
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('Todas')
  const [notice, setNotice] = useState('')
  const [showProduct, setShowProduct] = useState(false)
  const [editing, setEditing] = useState(null)
  const [assistantOpen, setAssistantOpen] = useState(false)
  const [qrOpen, setQrOpen] = useState(false)
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)
  const [overlay, setOverlay] = useState(null)
  const [importReport, setImportReport] = useState(null)
  const [messages, setMessages] = useState([{ role: 'assistant', text: 'Hola. Soy el asistente de VITIMAX. Puedo ayudarte a encontrar información del catálogo cargado o preparar una importación.' }])
  const [cloudReady, setCloudReady] = useState(!user || !firebaseDb)
  const [saving, setSaving] = useState(false)
  const fileRef = useRef(null)

  useEffect(() => {
    if (!user || !firebaseDb) return undefined
    const workspace = doc(firebaseDb, 'users', user.uid)
    setCloudReady(false)
    return onSnapshot(workspace, (snapshot) => {
      const data = snapshot.data() || {}
      if (Array.isArray(data.products)) setProducts(data.products)
      if (data.profile) setProfile((current) => ({ ...current, ...data.profile }))
      if (data.settings) setSettings((current) => ({ ...current, ...data.settings }))
      setCloudReady(true)
    }, () => { setCloudReady(true); setNotice('No se pudo sincronizar la nube; se usará una copia local.') })
  }, [user])
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(products))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { products }, { merge: true }).catch(() => setNotice('No se pudo guardar el catálogo en la nube.'))
  }, [products, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(PROFILE_KEY, JSON.stringify(profile))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { profile: { name: profile.name, email: profile.email, phone: profile.phone, photo: profile.photo } }, { merge: true }).catch(() => setNotice('No se pudo guardar el perfil en la nube.'))
  }, [profile, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { settings }, { merge: true }).catch(() => setNotice('No se pudo guardar la configuración en la nube.'))
  }, [settings, user, cloudReady])
  useEffect(() => { if (notice) { const timer = setTimeout(() => setNotice(''), 3500); return () => clearTimeout(timer) } }, [notice])

  const categories = useMemo(() => ['Todas', ...new Set(products.map((p) => p.category).filter(Boolean))], [products])
  const filtered = useMemo(() => products.filter((p) => {
    const haystack = Object.values(p).join(' ').toLowerCase()
    return haystack.includes(query.toLowerCase()) && (category === 'Todas' || p.category === category)
  }), [products, query, category])
  const lowStock = products.filter((p) => Number(p.stock) > 0 && Number(p.stock) <= 5)
  const missingData = products.filter((p) => !p.name || p.price === '' || p.stock === '')

  function navigate(item) { setActive(item); setQuery(''); setMobileMenuOpen(false); setOverlay(null) }
  function downloadTemplate() {
    const worksheet = XLSX.utils.json_to_sheet([emptyProduct])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Productos')
    XLSX.writeFile(workbook, 'plantilla-productos-vitimax.xlsx')
    setNotice('Plantilla Excel descargada')
  }
  function exportProducts() {
    const worksheet = XLSX.utils.json_to_sheet(products)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Productos')
    XLSX.writeFile(workbook, 'catalogo-vitimax.xlsx')
    setNotice('Catálogo exportado correctamente')
  }
  function generateAssistantExcel(template = false) {
    const rows = template ? [emptyProduct] : products
    const worksheet = XLSX.utils.json_to_sheet(rows)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Productos')
    XLSX.writeFile(workbook, template ? 'plantilla-productos-vitimax.xlsx' : 'catalogo-vitimax-ia.xlsx')
    setNotice(template ? 'Plantilla Excel descargada desde el asistente' : 'Excel generado y descargado desde el asistente')
  }
  async function generateAssistantFile(type, requestText) {
    const title = requestText.trim() || 'Documento VITIMAX'
    const catalogText = products.length
      ? products.map((item) => `${item.name || 'Producto pendiente'} | ${item.category || 'Categoría pendiente'} | Precio: ${money(item.price)} | Stock: ${item.stock === '' ? 'Pendiente' : item.stock}`).join('\n')
      : 'No hay productos registrados todavía.'
    if (type === 'pdf') {
      const pdf = new jsPDF()
      const lines = pdf.splitTextToSize(`${title}\n\nCatálogo VITIMAX\n\n${catalogText}`, 175)
      pdf.setFontSize(16)
      pdf.text(lines, 18, 22)
      pdf.save('documento-vitimax.pdf')
    } else if (type === 'word') {
      const document = new Document({ sections: [{ children: [
        new Paragraph({ text: title, heading: HeadingLevel.TITLE }),
        new Paragraph({ text: 'Catálogo VITIMAX', heading: HeadingLevel.HEADING_1 }),
        ...catalogText.split('\n').map((line) => new Paragraph(line))
      ] }] })
      downloadBlob(await Packer.toBlob(document), 'documento-vitimax.docx')
    } else if (type === 'ppt') {
      const presentation = new pptxgen()
      presentation.layout = 'LAYOUT_WIDE'
      const slide = presentation.addSlide()
      slide.background = { color: 'F4F7F2' }
      slide.addText(title, { x: 0.6, y: 0.7, w: 12, h: 0.7, fontSize: 24, bold: true, color: '173524' })
      slide.addText(`Catálogo VITIMAX\n\n${catalogText}`, { x: 0.7, y: 1.7, w: 11.8, h: 4.8, fontSize: 16, color: '173524', breakLine: false, fit: 'shrink' })
      downloadBlob(await presentation.write({ outputType: 'blob' }), 'presentacion-vitimax.pptx')
    } else if (type === 'code') {
      const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title><style>body{font-family:Arial,sans-serif;max-width:900px;margin:40px auto;padding:24px;color:#173524;background:#f4f7f2}section{background:#fff;padding:28px;border-radius:20px;box-shadow:0 10px 30px #17352418}li{margin:10px 0}</style></head><body><section><h1>${escapeHtml(title)}</h1><h2>Catálogo VITIMAX</h2><ul>${products.map((item) => `<li><strong>${escapeHtml(item.name || 'Producto pendiente')}</strong> — ${escapeHtml(item.description || 'Descripción pendiente.')}</li>`).join('')}</ul></section></body></html>`
      downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), 'codigo-vitimax.html')
    }
    setNotice(`${type === 'ppt' ? 'PowerPoint' : type === 'word' ? 'Word' : type.toUpperCase()} generado y descargado`)
  }
  function importExcel(event) {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (loadEvent) => {
      try {
        const workbook = XLSX.read(loadEvent.target.result, { type: 'array' })
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '' })
        const parsed = rows.map(readProduct).filter((p) => Object.values(p).some(Boolean))
        const invalid = parsed.filter((p) => !p.name).length
        const valid = parsed.filter((p) => p.name)
        const existing = new Map(products.map((p) => [normalize(p.name), p]))
        let updated = 0
        let added = 0
        valid.forEach((product) => {
          const key = normalize(product.name)
          if (!key) return
          if (existing.has(key)) { existing.set(key, { ...existing.get(key), ...product }); updated += 1 }
          else { existing.set(key, { ...product, id: makeId() }); added += 1 }
        })
        setProducts([...existing.values()])
        setImportReport({ rows: parsed.length, invalid, added, updated, missing: valid.filter((p) => p.price === '' || p.stock === '').length })
        setActive('Catálogo')
        setNotice(`${added} agregados · ${updated} actualizados · ${invalid} filas con error`)
      } catch {
        setNotice('No se pudo leer el archivo. Revisa que sea un .xlsx válido.')
      } finally { event.target.value = '' }
    }
    reader.readAsArrayBuffer(file)
  }
  async function uploadProductImage(file, productId) {
    const readLocal = () => new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(reader.result)
      reader.onerror = reject
      reader.readAsDataURL(file)
    })
    if (!firebaseStorage || !user) {
      return readLocal()
    }
    const storageRef = ref(firebaseStorage, `users/${user.uid}/products/${productId}/${file.name.replace(/[^a-zA-Z0-9._-]/g, '_')}`)
    try {
      await uploadBytes(storageRef, file, { contentType: file.type })
      return getDownloadURL(storageRef)
    } catch {
      return readLocal()
    }
  }
  async function saveProduct(product, imageFile) {
    setSaving(true)
    const id = product.id || makeId()
    let image = product.image
    try {
      if (imageFile) image = await uploadProductImage(imageFile, id)
      const clean = { ...product, id, image }
      setProducts((current) => {
        const index = current.findIndex((p) => p.id === clean.id)
        if (index < 0) return [...current, clean]
        return current.map((p, i) => i === index ? clean : p)
      })
      setShowProduct(false)
      setNotice(product.id ? 'Producto actualizado' : 'Producto creado correctamente')
    } catch {
      setNotice('No se pudo subir la imagen. Revisa la configuración de Firebase Storage.')
    } finally { setSaving(false) }
  }
  function deleteProduct(id) {
    if (window.confirm('¿Eliminar este producto del catálogo local?')) {
      setProducts((all) => all.filter((p) => p.id !== id))
      setNotice('Producto eliminado')
    }
  }
  async function saveProfile(next, photoFile) {
    let normalized = { ...next, name: next.name?.trim() || (guest ? 'Invitado' : 'Administrador'), phone: next.phone?.trim() || '', photo: next.photo?.trim() || '' }
    try {
      if (photoFile) {
        if (!photoFile.type.startsWith('image/')) throw new Error('El archivo seleccionado no es una imagen.')
        if (photoFile.size > 5 * 1024 * 1024) throw new Error('La foto debe pesar menos de 5 MB.')
        if (firebaseStorage && user && !guest) {
          try {
            const photoRef = ref(firebaseStorage, `users/${user.uid}/profile/avatar-${Date.now()}`)
            const uploaded = await uploadBytes(photoRef, photoFile, { contentType: photoFile.type })
            normalized = { ...normalized, photo: await getDownloadURL(uploaded.ref) }
          } catch (error) {
            throw new Error(error?.code === 'storage/unauthorized'
              ? 'Firebase no autorizó la foto. Inicia sesión nuevamente e inténtalo otra vez.'
              : 'No se pudo subir la foto a Firebase Storage.')
          }
        } else normalized = { ...normalized, photo: await imageFileToDataUrl(photoFile) }
      }
      if (firebaseAuth && user && !guest) await updateProfile(user, { displayName: normalized.name || null, photoURL: normalized.photo || null })
      localStorage.setItem(PROFILE_KEY, JSON.stringify(normalized))
      if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { profile: normalized }, { merge: true }).catch(() => setNotice('Perfil guardado localmente; no se pudo sincronizar la nube.'))
      setProfile(normalized)
      setOverlay(null)
      setNotice(guest ? 'Perfil guardado en este dispositivo' : 'Perfil guardado correctamente')
    } catch {
      setNotice('No se pudo actualizar el perfil de Firebase. Revisa la sesión e inténtalo de nuevo.')
    }
  }
  function saveSettings(next) { setSettings(next); setOverlay(null); setNotice('Configuración guardada correctamente') }
  function downloadProductPage(product = null, requestText = '') {
    const pageData = product || {
      name: 'VITIMAX',
      category: 'Catálogo y soluciones agrícolas',
      description: requestText || 'Soluciones VITIMAX para el cuidado y manejo responsable de cultivos.',
      price: '',
      format: 'Información bajo consulta',
      usage: 'Consulta la ficha técnica oficial antes de usar cualquier producto.',
      stock: '',
      image: ''
    }
    const slug = normalize(pageData.name || 'pagina-vitimax').replace(/\s+/g, '-') || 'pagina-vitimax'
    const html = `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${pageData.name || 'VITIMAX'}</title><style>body{font-family:Arial,sans-serif;margin:0;background:#f4f7f2;color:#173524}main{max-width:900px;margin:auto;padding:48px 24px}article{background:white;border-radius:24px;padding:32px;box-shadow:0 12px 40px #17352418}img{max-width:100%;max-height:360px;object-fit:contain;border-radius:16px}h1{font-size:clamp(2rem,5vw,4rem);margin:16px 0 8px}.tag{color:#4e8a52;text-transform:uppercase;font-weight:bold;letter-spacing:.12em}.price{font-size:1.6rem;font-weight:bold;margin:24px 0}a{display:inline-block;padding:14px 20px;border-radius:999px;background:#276b3b;color:#fff;text-decoration:none}</style></head><body><main><article>${pageData.image ? `<img src="${pageData.image}" alt="${pageData.name || 'VITIMAX'}">` : ''}<div class="tag">${pageData.category || 'VITIMAX'}</div><h1>${pageData.name || 'VITIMAX'}</h1><p>${pageData.description || 'Descripción pendiente de completar.'}</p>${pageData.price ? `<div class="price">${money(pageData.price)}</div>` : ''}<p><strong>Formato:</strong> ${pageData.format || 'Pendiente'}</p><p><strong>Modo de uso:</strong> ${pageData.usage || 'Pendiente de ficha técnica.'}</p><p><strong>Disponibilidad:</strong> ${pageData.stock === '' ? 'Pendiente' : `${pageData.stock} unidades`}</p><a href="mailto:${profile.email || ''}">Solicitar información</a></article></main></body></html>`
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `pagina-${slug}.html`
    link.click()
    URL.revokeObjectURL(url)
    setNotice('Página web creada y descargada')
  }
  async function askAssistant(text, attachmentFile = null) {
    const promptNormalized = normalize(text)
    const excelRequested = promptNormalized.includes('excel') || promptNormalized.includes('xlsx') || promptNormalized.includes('hoja de calculo') || promptNormalized.includes('hoja calculo')
    const templateRequested = excelRequested && (promptNormalized.includes('plantilla') || promptNormalized.includes('vacio') || promptNormalized.includes('vacío'))
    if (excelRequested && !guest) generateAssistantExcel(templateRequested)
    const fileType = promptNormalized.includes('pdf') ? 'pdf'
      : promptNormalized.includes('word') || promptNormalized.includes('docx') ? 'word'
        : promptNormalized.includes('powerpoint') || promptNormalized.includes('ppt') || promptNormalized.includes('presentacion') ? 'ppt'
          : promptNormalized.includes('codigo') || promptNormalized.includes('código') || promptNormalized.includes('programa') ? 'code'
            : null
    if (fileType && !guest) generateAssistantFile(fileType, text).catch((error) => setNotice(error?.message || 'No se pudo generar el archivo.'))
    const pageRequested = promptNormalized.includes('crear pagina') || promptNormalized.includes('pagina web') || promptNormalized.includes('landing')
    const pageTerms = promptNormalized.split(/\s+/).filter((term) => term.length > 3 && !['crear', 'pagina', 'web', 'promocional', 'promocion'].includes(term))
    const pageProduct = pageRequested && (products.length === 1 ? products[0] : products.find((item) => pageTerms.some((term) => normalize(item.name).includes(term))))
    if (pageRequested) downloadProductPage(pageProduct, text)
    let remoteAnswered = false
    const cloudAssistantEnabled = import.meta.env.VITE_ENABLE_CLOUD_ASSISTANT === 'true'
    if (firebaseFunctions && cloudAssistantEnabled && !guest) {
      const requestId = makeId()
      setMessages((items) => [...items, { role: 'user', text, requestId }, { role: 'assistant', text: 'Estoy consultando Gemini con los datos reales del catálogo...', requestId }])
      try {
        let attachment
        if (attachmentFile) {
          if (attachmentFile.size > 8 * 1024 * 1024) throw new Error('El archivo supera el límite de 8 MB para análisis.')
          const dataUrl = await new Promise((resolve, reject) => {
            const reader = new FileReader()
            reader.onload = () => resolve(reader.result)
            reader.onerror = reject
            reader.readAsDataURL(attachmentFile)
          })
          const [header, data] = String(dataUrl).split(',')
          attachment = { name: attachmentFile.name, mimeType: header.match(/data:(.*);base64/)?.[1] || attachmentFile.type || 'application/octet-stream', data }
        }
        const { data } = await httpsCallable(firebaseFunctions, 'vitimaxAssistant')({ message: text, products, attachment })
        setMessages((items) => items.map((item) => item.requestId === requestId ? { ...item, text: data.text } : item))
        remoteAnswered = true
      } catch (error) {
        setMessages((items) => items.filter((item) => item.requestId !== requestId))
        setNotice(error?.message || 'Gemini no respondió. Se usará el asistente local.')
      }
    }
    if (remoteAnswered) return
    const normalized = normalize(text)
    let attachmentNote = ''
    if (attachmentFile) {
      if (attachmentFile.size > 25 * 1024 * 1024) {
        attachmentNote = `El archivo "${attachmentFile.name}" supera el límite local de 25 MB.`
      } else if (attachmentFile.type.startsWith('text/') || /\.(txt|md|json|csv)$/i.test(attachmentFile.name)) {
        try {
          const content = await attachmentFile.text()
          attachmentNote = `ARCHIVO RECIBIDO: ${attachmentFile.name}\n\nPrimeros datos del archivo:\n${content.slice(0, 1800)}${content.length > 1800 ? '\n\n[Contenido recortado]' : ''}`
        } catch { attachmentNote = `Recibí "${attachmentFile.name}", pero no pude leer su contenido localmente.` }
      } else {
        attachmentNote = `ARCHIVO RECIBIDO: ${attachmentFile.name}\n\nTipo: ${attachmentFile.type || 'desconocido'}\nTamaño: ${Math.ceil(attachmentFile.size / 1024)} KB\n\nPuedo conservarlo en esta sesión. Para analizar el contenido de PDF, Word o imágenes con IA necesitas activar Gemini/Functions; mientras tanto puedo ayudarte con Excel, imágenes del catálogo y archivos de texto.`
      }
    }
    const terms = normalized.split(/\s+/).filter((term) => term.length > 2)
    const matches = products.filter((p) => {
      const haystack = normalize(Object.values(p).join(' '))
      return !terms.length || terms.some((term) => haystack.includes(term))
    })
    const low = products.filter((p) => Number(p.stock) > 0 && Number(p.stock) <= Number(settings.lowStock))
    let response = ''
    const product = matches[0]
    if (!products.length) {
      response = 'Todavía no hay productos reales cargados. Importa el Excel oficial para que pueda recomendar, crear páginas o redactar contenido sin inventar productos, precios, stock ni características.'
    } else if (normalized.includes('inventario') || normalized.includes('stock') || normalized.includes('reporte')) {
      response = `INFORME DE INVENTARIO\n\nProductos registrados: ${products.length}\nUnidades contabilizadas: ${products.reduce((sum, item) => sum + (Number(item.stock) || 0), 0)}\nStock bajo: ${low.length}\n\n${low.length ? `Requieren atención: ${low.map((item) => `${item.name || 'Sin nombre'} (${item.stock})`).join(', ')}` : 'No hay productos dentro del umbral configurado.'}`
    } else if (normalized.includes('excel') || normalized.includes('import') || normalized.includes('plantilla')) {
      response = guest
        ? 'Para generar y descargar un Excel desde el asistente debes iniciar sesión como administrador.'
        : templateRequested
        ? 'PLANTILLA EXCEL GENERADA\n\nDescargué una plantilla vacía con las columnas oficiales de VITIMAX. Puedes completarla y luego importarla desde Catálogo.'
        : products.length
          ? `EXCEL GENERADO\n\nDescargué el archivo catalogo-vitimax-ia.xlsx con los ${products.length} productos reales registrados. Puedes abrirlo, editarlo y volver a importarlo desde Catálogo.`
          : 'No hay productos cargados todavía. Descargué una plantilla Excel vacía para que puedas completar los datos oficiales.'
    } else if (normalized.includes('crear pagina') || normalized.includes('pagina web') || normalized.includes('landing') || normalized.includes('promocion')) {
      if (!product) {
        response = 'Puedo crear la estructura de una página promocional, pero necesito que indiques el nombre de un producto registrado o cargues primero el catálogo.'
      } else {
        response = `PÁGINA WEB CREADA: ${product.name || 'Producto pendiente'}\n\nARCHIVO: Se descargó una página HTML lista para abrir y publicar.\n\nPASO 1 · Objetivo\nPresentar ${product.name || 'el producto'} con información confirmada y facilitar una solicitud de contacto.\n\nPASO 2 · Público\nPersonas que buscan información sobre ${product.category || 'este tipo de producto'}. No inventé un público ni beneficios que no estén registrados.\n\nPASO 3 · Estructura\n1. Encabezado y categoría\n2. Nombre y descripción\n3. Formato, modo de uso y disponibilidad\n4. Precio registrado\n5. Botón de contacto\n\nPASO 4 · Contenido confirmado\nTítulo: ${product.name || 'Producto pendiente'}\nDescripción: ${product.description || 'Descripción pendiente de completar.'}\nInformación disponible: ${[product.category && `Categoría: ${product.category}`, product.format && `Formato: ${product.format}`, product.features && `Características: ${product.features}`].filter(Boolean).join(' · ') || 'Pendiente de completar.'}\nUso: ${product.usage || 'Pendiente de completar con la ficha técnica oficial.'}\nPrecio: ${money(product.price)}\nStock: ${product.stock === '' ? 'Pendiente' : `${product.stock} unidades registradas`}\n\nPASO 5 · Diseño\nSe generó un diseño responsive, limpio y preparado para móvil, con tarjeta de producto, jerarquía visual y botón de contacto.\n\nPASO 6 · Llamada a la acción\nSolicitar información sobre ${product.name || 'este producto'}.\n\nPASO 7 · Publicación\nAbre el archivo .html descargado en el navegador, revisa la información y súbelo a GitHub Pages, Netlify o tu hosting.\n\nNo se publicaron datos que no estén registrados. Revisa la ficha técnica antes de publicar recomendaciones.`
      }
    } else if (normalized.includes('descripcion') || normalized.includes('publicacion') || normalized.includes('redes') || normalized.includes('contenido')) {
      response = product
        ? `BORRADOR DE CONTENIDO\n\n${product.name || 'Producto pendiente'}\n\n${product.description || 'La descripción todavía está pendiente de completar.'}\n\nDatos confirmados: ${[product.category, product.format, product.features].filter(Boolean).join(' · ') || 'No hay características registradas.'}\n\nPrecio: ${money(product.price)} · Stock: ${product.stock === '' ? 'Pendiente' : product.stock}\n\nCompleta los campos pendientes antes de publicar.`
        : 'Indica el nombre de un producto registrado y crearé un borrador utilizando solamente sus datos.'
    } else if (normalized.includes('dosis') || normalized.includes('aplicar') || normalized.includes('tratamiento') || normalized.includes('sirve')) {
      response = product
        ? `INFORMACIÓN REGISTRADA DE ${product.name || 'PRODUCTO'}\n\nModo de uso: ${product.usage || 'Pendiente de ficha técnica.'}\nRecomendaciones: ${product.recommendations || 'Pendientes de ficha técnica.'}\n\nNo puedo calcular ni inventar una dosis. Confirma la etiqueta oficial y consulta a un especialista de VITIMAX antes de aplicar cualquier producto.`
        : 'Necesito el nombre de un producto registrado. No entregaré dosis ni tratamientos que no estén en una ficha técnica oficial.'
    } else {
      response = `Encontré ${matches.length} coincidencia(s) en los datos registrados. Solo recomiendo productos existentes en VITIMAX.${matches.length ? `\n\n${matches.slice(0, 5).map((item) => `• ${item.name || 'Sin nombre'} · ${money(item.price)} · Stock: ${item.stock === '' ? 'Pendiente' : item.stock}`).join('\n')}` : '\n\nNo encontré coincidencias. Prueba con el nombre, categoría, uso o problema registrado.'}`
    }
    setMessages((items) => [...items, { role: 'user', text: attachmentFile ? `${text}\n📎 ${attachmentFile.name}` : text }, { role: 'assistant', text: attachmentNote ? `${response}\n\n${attachmentNote}` : response }])
  }
  function handleQrResult(value) {
    setQrOpen(false)
    const match = products.find((product) => normalize(product.id) === normalize(value) || normalize(product.name) === normalize(value) || String(product.qr || '') === value)
    if (match) {
      setActive('Catálogo')
      setQuery(match.name)
      setMessages((items) => [...items, { role: 'assistant', text: `Encontré ${match.name} en el catálogo. La ficha muestra únicamente la información registrada en VITIMAX.` }])
      setAssistantOpen(true)
      return
    }
    askAssistant(`Producto escaneado: ${value}`)
    setAssistantOpen(true)
  }

  const nav = [
    { label: 'Resumen', icon: LayoutDashboard },
    { label: 'Catálogo', icon: ShoppingBag },
    { label: 'Inventario', icon: Boxes },
    { label: 'Pedidos', icon: ClipboardList },
    { label: 'Clientes', icon: Users },
    { label: 'Asesorías', icon: Sprout },
    { label: 'Contenido', icon: Sparkles }
  ]

  const canManage = !guest
  return <div className={guest ? 'app-shell guest-mode' : 'app-shell'}>
    <aside className={mobileMenuOpen ? 'sidebar mobile-open' : 'sidebar'}>
      <div className="brand"><div className="brand-mark"><Leaf size={22} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div>
      <div className="workspace-label">ESPACIO DE TRABAJO <span className="status-dot" /></div>
      <nav>{nav.map(({ label, icon: Icon }) => <button key={label} className={active === label ? 'nav-item active' : 'nav-item'} onClick={() => navigate(label)}><Icon size={18} /><span>{label}</span>{label === 'Asesorías' && <em>0</em>}</button>)}</nav>
      <div className="sidebar-bottom">{canManage && <button className="nav-item" onClick={() => { setOverlay('settings'); setMobileMenuOpen(false) }}><Settings2 size={18} /><span>Configuración</span></button>}<button className="profile profile-button" onClick={() => setOverlay('profile')}><div className="avatar">{profile.name?.[0] || 'V'}</div><div><strong>{profile.name || 'Invitado'}</strong><span>{profile.email || 'Acceso de invitado'}</span></div><MoreHorizontal size={17} /></button></div>
    </aside>
    <main className="main">
      <header className="topbar"><button className="mobile-menu" aria-label="Abrir menú" onClick={() => setMobileMenuOpen(true)}><Menu size={21} /></button><div className="crumb">VITIMAX <ChevronRight size={14} /> <b>{active}</b></div><div className="top-actions"><button className="icon-btn" title="Escanear código QR" aria-label="Escanear código QR" onClick={() => setQrOpen(true)}><QrCode size={18} /></button><button className="icon-btn" title="Notificaciones" aria-label="Notificaciones" onClick={() => setOverlay('notifications')}><Bell size={18} /><i /></button><button className="user-chip" onClick={() => setOverlay('profile')}><div className="avatar small">{profile.name?.[0] || 'V'}</div><span>{profile.name || 'Administrador'}</span><ChevronRight size={15} /></button></div></header>
      <div className="content">
        {active === 'Resumen' && <Dashboard products={products} lowStock={lowStock} missingData={missingData} onImport={canManage ? () => fileRef.current?.click() : undefined} onTemplate={canManage ? downloadTemplate : undefined} onNavigate={navigate} />}
        {active === 'Catálogo' && <Catalog products={filtered} allProducts={products} categories={categories} query={query} setQuery={setQuery} category={category} setCategory={setCategory} onImport={canManage ? () => fileRef.current?.click() : undefined} onExport={canManage ? exportProducts : undefined} onTemplate={canManage ? downloadTemplate : undefined} onCreate={canManage ? () => { setEditing(null); setShowProduct(true) } : undefined} onEdit={canManage ? (p) => { setEditing(p); setShowProduct(true) } : undefined} onDelete={canManage ? (id) => setProducts((all) => all.filter((p) => p.id !== id)) : undefined} readOnly={guest} />}
        {active === 'Inventario' && <Inventory products={products} lowStock={lowStock} onImport={canManage ? () => fileRef.current?.click() : undefined} readOnly={guest} />}
        {['Pedidos', 'Clientes', 'Asesorías', 'Contenido'].includes(active) && <EmptySection active={active} onNavigate={navigate} onImport={() => fileRef.current?.click()} onAssistant={() => setAssistantOpen(true)} />}
      </div>
    </main>
    <input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={importExcel} />
    {notice && <div className="toast"><PackageCheck size={18} />{notice}</div>}
    {mobileMenuOpen && <button className="drawer-backdrop" aria-label="Cerrar menú" onClick={() => setMobileMenuOpen(false)} />}
    {overlay === 'profile' && <ProfileModal profile={profile} onClose={() => setOverlay(null)} onSave={saveProfile} onSignOut={onSignOut} />}
    {overlay === 'settings' && <SettingsModal settings={settings} onClose={() => setOverlay(null)} onSave={saveSettings} />}
    {overlay === 'notifications' && <NotificationsModal lowStock={lowStock} missingData={missingData} onClose={() => setOverlay(null)} />}
    {importReport && <ImportReport report={importReport} onClose={() => setImportReport(null)} />}
    <button className="assistant-fab" onClick={() => setAssistantOpen(true)} aria-label="Abrir BOT IA VITIMAX"><img className="assistant-bot-image" src="/bot-vitimax.png" alt="" /><span>BOT IA VITIMAX</span></button>
    {assistantOpen && <Assistant messages={messages} onClose={() => setAssistantOpen(false)} onAsk={askAssistant} />}
    {qrOpen && <QrScanner onClose={() => setQrOpen(false)} onResult={handleQrResult} />}
    {showProduct && canManage && <ProductModal product={editing || emptyProduct} onClose={() => setShowProduct(false)} onSave={saveProduct} onDelete={deleteProduct} saving={saving} />}
  </div>
}

function Dashboard({ products, lowStock, missingData, onImport, onTemplate, onNavigate }) {
  const readOnly = !onImport
  return <><div className="page-heading"><div><p className="eyebrow">CENTRO DIGITAL · 01</p><h1>{readOnly ? 'Bienvenido a VITIMAX' : 'Buenos días, administrador'} <span>✦</span></h1><p className="muted">{readOnly ? 'Consulta el catálogo y la información disponible.' : 'Aquí tienes una vista general de la operación de VITIMAX.'}</p></div>{!readOnly && <div className="heading-actions"><button className="btn secondary" onClick={onTemplate}><FileSpreadsheet size={16} /> Plantilla Excel</button><button className="btn primary" onClick={onImport}><Upload size={16} /> Importar productos</button></div>}</div>
    <section className="hero-card"><div><span className="pill light"><Zap size={13} /> {readOnly ? 'Modo invitado' : 'Configuración inicial'}</span><h2>Tu centro de VITIMAX,<br /><em>listo para crecer.</em></h2><p>{readOnly ? 'Explora el catálogo, revisa el inventario y consulta al asistente con la información disponible.' : 'Carga el Excel oficial de productos para activar el catálogo, inventario y recomendaciones con datos reales.'}</p>{!readOnly && <button className="btn hero-btn" onClick={onImport}>Comenzar importación <ArrowUpToLine size={16} /></button>}</div><div className="hero-illustration"><Sprout size={110} strokeWidth={1.2} /><Leaf className="leaf-float one" size={30} /><Leaf className="leaf-float two" size={22} /></div></section>
    <div className="section-title"><h2>Vista general</h2><span>Actualizado ahora</span></div>
    <section className="metrics"><Metric icon={ShoppingBag} label="Productos en catálogo" value={products.length} detail={products.length ? 'datos registrados' : 'pendiente de importar'} color="green" /><Metric icon={Boxes} label="Unidades en stock" value={products.reduce((sum, p) => sum + (Number(p.stock) || 0), 0)} detail="según catálogo actual" color="blue" /><Metric icon={Bell} label="Alertas de inventario" value={lowStock.length} detail={lowStock.length ? 'requieren atención' : 'sin alertas'} color="orange" /><Metric icon={CircleHelp} label="Datos pendientes" value={missingData.length} detail="para completar ficha" color="purple" /></section>
    <div className="dashboard-grid"><section className="panel"><div className="panel-heading"><div><h3>Actividad del catálogo</h3><p>Productos cargados recientemente</p></div><button className="link-btn" onClick={() => onNavigate('Catálogo')}>Ver catálogo <ChevronRight size={15} /></button></div>{products.length ? products.slice(0, 5).map((p) => <div className="activity-row" key={p.id}><div className="product-thumb">{p.image ? <img src={p.image} /> : <Leaf size={19} />}</div><div><strong>{p.name || 'Sin nombre'}</strong><span>{p.category || 'Categoría pendiente'}</span></div><b>{money(p.price)}</b></div>) : <EmptyState icon={FileSpreadsheet} title="Aún no hay productos" text="Importa el Excel oficial de VITIMAX para comenzar." action="Importar Excel" onClick={onImport} />}</section><section className="panel checklist"><div className="panel-heading"><div><h3>Configuración inicial</h3><p>Deja todo listo para operar</p></div><span className="progress">1 / 5</span></div>{['Importar catálogo oficial', 'Completar información de empresa', 'Configurar categorías', 'Revisar niveles mínimos', 'Publicar catálogo'].map((item, i) => <div className={i === 0 && products.length ? 'check-row done' : 'check-row'} key={item}><span>{i === 0 && products.length ? '✓' : i + 1}</span><div><strong>{item}</strong><small>{i === 0 && products.length ? 'Completado' : 'Pendiente'}</small></div></div>)}</section></div>
  </>
}

function Metric({ icon: Icon, label, value, detail, color }) { return <div className="metric"><div className={`metric-icon ${color}`}><Icon size={19} /></div><div><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></div> }

function Catalog({ products, allProducts, categories, query, setQuery, category, setCategory, onImport, onExport, onTemplate, onCreate, onEdit, onDelete, readOnly }) {
  return <><div className="page-heading"><div><p className="eyebrow">{readOnly ? 'CONSULTA · 02' : 'GESTIÓN · 02'}</p><h1>Catálogo de productos</h1><p className="muted">La información visible proviene únicamente de los datos cargados.</p></div>{!readOnly && <button className="btn primary" onClick={onCreate}><Plus size={17} /> Nuevo producto</button>}</div><div className="toolbar"><div className="searchbox"><Search size={18} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar por nombre, uso, problema o característica..." /></div><div className="selectbox"><Filter size={16} /><select value={category} onChange={(e) => setCategory(e.target.value)}>{categories.map((c) => <option key={c}>{c}</option>)}</select></div>{!readOnly && <div className="toolbar-actions"><button className="btn secondary compact" onClick={onTemplate}><FileSpreadsheet size={15} /> Plantilla</button><button className="btn secondary compact" onClick={onImport}><Upload size={15} /> Importar</button><button className="btn secondary compact" onClick={onExport} disabled={!allProducts.length}><ArrowDownToLine size={15} /> Exportar</button></div>}</div>{products.length ? <div className="product-grid">{products.map((p) => <ProductCard key={p.id} product={p} onEdit={onEdit} onDelete={onDelete} readOnly={readOnly} />)}</div> : <div className="empty-catalog"><div className="empty-icon"><FileSpreadsheet size={28} /></div><h2>Tu catálogo está esperando los datos oficiales</h2><p>{readOnly ? 'Todavía no hay productos publicados para consultar.' : 'No se encontraron productos registrados. Importa un archivo Excel para cargar nombres, precios, stock y toda la información disponible sin duplicados.'}</p>{!readOnly && <div><button className="btn primary" onClick={onImport}><Upload size={16} /> Importar Excel</button><button className="btn secondary" onClick={onTemplate}>Descargar plantilla</button></div>}<small>{readOnly ? 'El catálogo será visible cuando el administrador cargue los datos oficiales.' : 'Formatos aceptados: .xlsx, .xls · La primera hoja será procesada'}</small></div>}</>
}

function ProductCard({ product, onEdit, onDelete, readOnly }) { return <article className="product-card"><div className="card-image">{product.image ? <img src={product.image} alt="" /> : <><Leaf size={37} /><span>Foto pendiente</span></> }<span className={Number(product.stock) > 0 ? 'stock-badge available' : 'stock-badge'}>{Number(product.stock) > 0 ? `${product.stock} en stock` : 'Stock pendiente'}</span>{!readOnly && <button className="card-menu" onClick={() => onEdit(product)} aria-label={`Editar ${product.name || 'producto'}`}><MoreHorizontal size={17} /></button>}</div><div className="card-body"><div className="category-label">{product.category || 'Categoría pendiente'}</div><h3>{product.name || 'Producto sin nombre'}</h3><p>{product.description || 'Descripción pendiente de completar.'}</p><div className="card-footer"><strong>{money(product.price)}</strong>{!readOnly && <div className="product-actions"><button className="link-btn" onClick={() => onEdit(product)}>Editar <ChevronRight size={14} /></button><button className="link-btn danger-link" onClick={() => onDelete(product.id)}><Trash2 size={14} /> Eliminar</button></div>}</div></div></article> }

function Inventory({ products, lowStock, onImport, readOnly }) { return <><div className="page-heading"><div><p className="eyebrow">OPERACIÓN · 03</p><h1>Inventario</h1><p className="muted">Consulta existencias y detecta rápidamente lo que necesita atención.</p></div>{!readOnly && <button className="btn primary" onClick={onImport}><Upload size={16} /> Actualizar por Excel</button>}</div><div className="metrics inventory-metrics"><Metric icon={Boxes} label="Total de unidades" value={products.reduce((s, p) => s + (Number(p.stock) || 0), 0)} detail="stock registrado" color="green" /><Metric icon={Bell} label="Stock bajo" value={lowStock.length} detail="umbral de 5 unidades" color="orange" /><Metric icon={Tags} label="Sin stock" value={products.filter((p) => p.stock === 0 || p.stock === '0').length} detail="productos agotados" color="purple" /></div><section className="panel table-panel"><div className="panel-heading"><div><h3>Existencias actuales</h3><p>Los valores se actualizan desde el catálogo o Excel.</p></div></div>{products.length ? <div className="table-wrap"><table><thead><tr><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock actual</th><th>Estado</th></tr></thead><tbody>{products.map((p) => <tr key={p.id}><td><strong>{p.name || 'Sin nombre'}</strong></td><td>{p.category || 'Pendiente'}</td><td>{money(p.price)}</td><td>{p.stock === '' ? 'Pendiente' : p.stock}</td><td><span className={Number(p.stock) <= 5 ? 'table-status warning' : 'table-status ok'}>{p.stock === '' ? 'Revisar' : Number(p.stock) <= 5 ? 'Stock bajo' : 'En rango'}</span></td></tr>)}</tbody></table></div> : <EmptyState icon={Boxes} title="Sin inventario registrado" text={readOnly ? 'El inventario aparecerá cuando el administrador cargue los datos.' : 'Carga el Excel oficial para visualizar existencias.'} action={readOnly ? 'Ir al catálogo' : 'Importar Excel'} onClick={readOnly ? undefined : onImport} />}</section></> }

function EmptySection({ active, onNavigate, onImport, onAssistant }) { const content = { Pedidos: ['Gestiona tus pedidos', 'Cuando se registren compras, aquí podrás seguir cada etapa del pedido.', ClipboardList], Clientes: ['Conoce a tus clientes', 'El historial de consultas y pedidos se organizará aquí.', Users], Asesorías: ['Asesorías agrícolas', 'Recibe problemas, fotografías y consultas para orientar a tus clientes.', Sprout], Contenido: ['Contenido y páginas', 'Crea páginas, promociones y publicaciones con ayuda del asistente.', Sparkles] }[active]; const Icon = content[2]; return <div className="module-empty"><div className="module-orb"><Icon size={34} /></div><p className="eyebrow">MÓDULO · {active.toUpperCase()}</p><h1>{content[0]}</h1><p>{content[1]}</p><div className="module-actions"><button className="btn primary" onClick={() => active === 'Contenido' ? onAssistant() : onImport()}>{active === 'Asesorías' ? 'Configurar formulario' : active === 'Contenido' ? 'Abrir asistente' : 'Importar datos'}</button><button className="btn secondary" onClick={() => onNavigate('Catálogo')}>Ir al catálogo</button></div><div className="coming-soon"><Zap size={15} /> Módulo preparado para conectarse a la base de datos</div></div> }

function EmptyState({ icon: Icon, title, text, action, onClick }) { return <div className="empty-state"><Icon size={28} /><h3>{title}</h3><p>{text}</p><button className="link-btn" onClick={onClick}>{action} <ChevronRight size={15} /></button></div> }

function ModalShell({ title, eyebrow, children, onClose, actions }) {
  return <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}><div className="modal"><div className="modal-head"><div><p className="eyebrow">{eyebrow}</p><h2>{title}</h2></div><button className="icon-btn" aria-label="Cerrar" onClick={onClose}><X size={19} /></button></div>{children}{actions && <div className="modal-actions">{actions}</div>}</div></div>
}

function ProfileModal({ profile, onClose, onSave, onSignOut }) {
  const [form, setForm] = useState(profile)
  const [photoFile, setPhotoFile] = useState(null)
  const [photoPreview, setPhotoPreview] = useState(profile.photo || '')
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))
  const choosePhoto = (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    if (!file.type.startsWith('image/')) return
    if (file.size > 5 * 1024 * 1024) return
    setPhotoFile(file)
    setPhotoPreview(URL.createObjectURL(file))
  }
  return <ModalShell title="Mi perfil" eyebrow="CUENTA" onClose={onClose} actions={<><button className="btn secondary" onClick={onSignOut}>Cerrar sesión</button><button className="btn secondary" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={() => onSave(form, photoFile)}>Guardar cambios</button></>}>
    <div className="profile-preview">{photoPreview ? <img className="profile-photo" src={photoPreview} alt="Vista previa de la foto de perfil" /> : <div className="avatar profile-avatar">{form.name?.[0] || 'V'}</div>}<div><strong>{form.name || 'Administrador'}</strong><span>Administrador de VITIMAX</span></div></div>
    <div className="form-grid"><label>Nombre<input value={form.name} onChange={(e) => update('name', e.target.value)} /></label><label>Correo electrónico<input type="email" value={form.email} onChange={(e) => update('email', e.target.value)} placeholder="Pendiente" readOnly /></label><label>Teléfono<input value={form.phone} onChange={(e) => update('phone', e.target.value)} placeholder="Pendiente" /></label><label>Foto de perfil (URL)<input value={form.photo} onChange={(e) => { update('photo', e.target.value); setPhotoPreview(e.target.value) }} placeholder="Opcional: URL pública" /></label><div className="full"><span className="field-label">Foto de perfil</span><div className="modal-actions"><label className="btn secondary">Elegir desde celular o computador<input type="file" accept="image/*" onChange={choosePhoto} hidden /></label><label className="btn secondary">Tomar foto<input type="file" accept="image/*" capture="user" onChange={choosePhoto} hidden /></label></div><small>JPG, PNG o WEBP · máximo 5 MB. Puedes elegir una imagen del dispositivo o tomar una foto.</small>{photoFile && <small>Foto seleccionada: {photoFile.name} · {Math.ceil(photoFile.size / 1024)} KB</small>}</div></div>
  </ModalShell>
}

function SettingsModal({ settings, onClose, onSave }) {
  const [form, setForm] = useState(settings)
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))
  return <ModalShell title="Configuración" eyebrow="PREFERENCIAS" onClose={onClose} actions={<><button className="btn secondary" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={() => onSave(form)}>Guardar configuración</button></>}>
    <div className="settings-list"><label className="setting-row"><span><strong>Notificaciones</strong><small>Mostrar alertas de inventario y actividad</small></span><input type="checkbox" checked={form.notifications} onChange={(e) => update('notifications', e.target.checked)} /></label><label>Umbral de stock bajo<input type="number" min="0" value={form.lowStock} onChange={(e) => update('lowStock', e.target.value)} /></label><label>Moneda del catálogo<select value={form.currency} onChange={(e) => update('currency', e.target.value)}><option>ARS</option><option>CLP</option><option>USD</option></select></label></div>
  </ModalShell>
}

function NotificationsModal({ lowStock, missingData, onClose }) {
  return <ModalShell title="Notificaciones" eyebrow="CENTRO DE ALERTAS" onClose={onClose}><div className="notification-list">{lowStock.length ? lowStock.map((p) => <div className="notification-row" key={p.id}><Bell size={16} /><span><strong>Stock bajo</strong><small>{p.name || 'Producto pendiente'} · {p.stock} unidades</small></span></div>) : <div className="empty-state"><Bell size={25} /><h3>Sin alertas de stock</h3><p>No hay productos dentro del umbral actual.</p></div>}{missingData.length > 0 && <div className="notification-row"><CircleHelp size={16} /><span><strong>Datos pendientes</strong><small>{missingData.length} producto(s) necesitan información</small></span></div>}</div></ModalShell>
}

function ImportReport({ report, onClose }) {
  return <ModalShell title="Resultado de importación" eyebrow="EXCEL VALIDADO" onClose={onClose}><div className="import-summary"><Metric icon={FileSpreadsheet} label="Filas leídas" value={report.rows} detail="primera hoja" color="blue" /><Metric icon={Plus} label="Agregados" value={report.added} detail="nuevos registros" color="green" /><Metric icon={PackageCheck} label="Actualizados" value={report.updated} detail="sin duplicar" color="purple" /><Metric icon={CircleHelp} label="Pendientes" value={report.missing + report.invalid} detail={`${report.invalid} con error`} color="orange" /></div><p className="modal-note"><CircleHelp size={15} /> Los productos con precio o stock faltante permanecen marcados como pendientes. No se inventaron datos.</p><div className="modal-actions"><button className="btn primary" onClick={onClose}>Continuar</button></div></ModalShell>
}

function ProductModal({ product, onClose, onSave, onDelete, saving }) {
  const [form, setForm] = useState(product)
  const [imageFile, setImageFile] = useState(null)
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))
  const remove = () => { if (!product.id) return; if (window.confirm(`¿Eliminar "${product.name || 'este producto'}"? Esta acción no se puede deshacer.`)) { onDelete(product.id); onClose() } }
  return <div className="modal-backdrop"><div className="modal"><div className="modal-head"><div><p className="eyebrow">FICHA DE PRODUCTO</p><h2>{product.id ? 'Editar producto' : 'Nuevo producto'}</h2></div><button className="icon-btn" onClick={onClose}><X size={19} /></button></div><p className="modal-note"><CircleHelp size={15} /> Los campos sin información quedarán marcados como pendientes.</p><div className="form-grid">{[['name','Nombre del producto'],['category','Categoría'],['price','Precio'],['previousPrice','Precio anterior'],['stock','Stock disponible'],['format','Formato o tamaño'],['image','URL de fotografía']].map(([key, label]) => <label key={key}>{label}<input value={form[key] ?? ''} onChange={(e) => update(key, e.target.value)} placeholder="Pendiente" /></label>)}<label className="full">Subir fotografía<input type="file" accept="image/*" onChange={(e) => setImageFile(e.target.files?.[0] || null)} />{imageFile && <small>{imageFile.name}</small>}</label><label className="full">Descripción<textarea value={form.description ?? ''} onChange={(e) => update('description', e.target.value)} placeholder="Pendiente de completar" /></label><label>Modo de uso<textarea value={form.usage ?? ''} onChange={(e) => update('usage', e.target.value)} placeholder="Pendiente" /></label><label>Recomendaciones<textarea value={form.recommendations ?? ''} onChange={(e) => update('recommendations', e.target.value)} placeholder="Pendiente" /></label></div><div className="modal-actions">{product.id && <button className="btn danger-button" onClick={remove} disabled={saving}><Trash2 size={16} /> Eliminar producto</button>}<button className="btn secondary" onClick={onClose} disabled={saving}>Cancelar</button><button className="btn primary" onClick={() => onSave(form, imageFile)} disabled={saving}><PackageCheck size={16} /> {saving ? 'Subiendo...' : product.id ? 'Guardar cambios' : 'Agregar producto'}</button></div></div></div>
}

function QrScanner({ onClose, onResult }) {
  const scannerRef = useRef(null)
  const resultRef = useRef(onResult)
  const activeRef = useRef(true)
  const [error, setError] = useState('')
  const [cameraReady, setCameraReady] = useState(false)
  useEffect(() => { resultRef.current = onResult }, [onResult])
  useEffect(() => {
    const scanner = new Html5Qrcode('vitimax-qr-reader')
    scannerRef.current = scanner
    activeRef.current = true
    const cameraError = (scannerError) => {
      const name = scannerError?.name || ''
      if (name === 'NotAllowedError' || /permission|not allowed/i.test(scannerError?.message || '')) return 'Permiso de cámara rechazado. Actívalo en los ajustes del navegador o usa “Leer imagen QR”.'
      if (name === 'NotFoundError') return 'No se encontró una cámara disponible en este dispositivo.'
      if (name === 'NotReadableError') return 'La cámara está siendo usada por otra aplicación. Ciérrala e inténtalo otra vez.'
      if (window.location.protocol !== 'https:' && !['localhost', '127.0.0.1'].includes(window.location.hostname)) return 'El lector necesita una conexión HTTPS para usar la cámara.'
      return 'No se pudo abrir la cámara. Revisa el permiso del navegador o usa “Leer imagen QR”.'
    }
    scanner.start({ facingMode: 'environment' }, { fps: 10, qrbox: { width: 220, height: 220 }, aspectRatio: 1 }, (decodedText) => {
      if (!activeRef.current) return
      activeRef.current = false
      scanner.stop().catch(() => undefined)
      resultRef.current(decodedText)
    }, () => undefined).then(() => setCameraReady(true)).catch((scannerError) => setError(cameraError(scannerError)))
    return () => {
      activeRef.current = false
      if (scanner.isScanning) scanner.stop().catch(() => undefined)
    }
  }, [])
  async function scanImage(event) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file || !scannerRef.current) return
    setError('')
    activeRef.current = false
    if (scannerRef.current.isScanning) await scannerRef.current.stop().catch(() => undefined)
    try {
      const decodedText = await scannerRef.current.scanFile(file, true)
      resultRef.current(decodedText)
    } catch {
      setError('No encontré un código QR legible en esa imagen. Prueba con una foto más clara.')
    }
  }
  return <div className="modal-backdrop"><div className="modal qr-modal"><div className="modal-head"><div><p className="eyebrow">LECTOR DE ENVASES</p><h2>Escanea el código QR</h2></div><button className="icon-btn" aria-label="Cerrar escáner" onClick={onClose}><X size={19} /></button></div><p className="modal-note"><QrCode size={15} /> Enfoca el código dentro del recuadro. Debe identificar un producto registrado en el catálogo de VITIMAX.</p><div id="vitimax-qr-reader" className="qr-reader" />{cameraReady && <p className="qr-status">Cámara lista · apunta al código QR</p>}{error && <p className="auth-error">{error}</p>}<label className="qr-file-btn">Leer imagen QR<input type="file" accept="image/*" capture="environment" onChange={scanImage} /></label><small className="field-help qr-help">Si la cámara no funciona, toma o selecciona una foto del código QR.</small><button className="btn secondary qr-close" onClick={onClose}>Cerrar escáner</button></div></div>
}

function AssistantIcon() {
  return <span className="assistant-logo" aria-hidden="true"><svg viewBox="0 0 32 32" role="img"><path className="assistant-antenna" d="M16 5V2m0 0-2 2m2-2 2 2" /><rect x="5" y="7" width="22" height="18" rx="6" /><path d="M10 15h.01M22 15h.01M11 20c3 2 7 2 10 0" /><path className="assistant-ear" d="M5 14H2m25 0h3" /></svg><b>IA</b></span>
}

function Assistant({ messages, onClose, onAsk }) { const [text, setText] = useState(''); const [attachment, setAttachment] = useState(null); const submit = (e) => { e.preventDefault(); if (!text.trim() && !attachment) return; onAsk(text.trim() || 'Analiza el archivo adjunto y dime qué información adicional necesitas.', attachment); setText(''); setAttachment(null) }; return <div className="assistant-panel"><div className="assistant-head"><div className="assistant-title"><div className="bot-avatar"><img className="assistant-bot-image panel-bot-image" src="/bot-vitimax.png" alt="Robot BOT IA VITIMAX" /></div><div><strong>BOT IA VITIMAX</strong><span>Gemini · Catálogo · Archivos · Páginas</span></div></div><button className="icon-btn" onClick={onClose}><X size={18} /></button></div><div className="assistant-body">{messages.map((m, i) => <div className={`chat ${m.role}`} key={i}><div>{m.text}</div></div>)}</div><div className="quick-prompts"><button onClick={() => onAsk('Crea paso a paso una página web promocional usando el producto registrado que corresponda.')}>Crear página web</button><button onClick={() => onAsk('Genera y descarga un Excel con mi catálogo actual.')}>Generar Excel</button><button onClick={() => onAsk('Genera y descarga un PDF con mi catálogo actual.')}>Generar PDF</button><button onClick={() => onAsk('Genera y descarga un documento Word con mi catálogo actual.')}>Generar Word</button><button onClick={() => onAsk('Genera y descarga una presentación PowerPoint con mi catálogo actual.')}>Generar PowerPoint</button><button onClick={() => onAsk('Genera y descarga el código de una página web.')}>Generar código</button><button onClick={() => onAsk('Ayúdame a organizar y analizar mi Excel de productos.')}>Ayuda con Excel</button></div><form className="assistant-input" onSubmit={submit}><input value={text} onChange={(e) => setText(e.target.value)} placeholder="Pregunta o pide crear algo..." /><label className="assistant-photo" title="Subir archivo">📎<input type="file" accept="image/*,.pdf,.txt,.md,.json,.csv,.xlsx,.xls,.doc,.docx" hidden onChange={(e) => setAttachment(e.target.files?.[0] || null)} /></label><button type="submit"><ArrowUpToLine size={17} /></button></form>{attachment && <small className="assistant-file">Archivo listo: {attachment.name} ({Math.ceil(attachment.size / 1024)} KB)</small>}<small>Los archivos para análisis deben pesar menos de 8 MB. Verifica diagnósticos y aplicaciones con un especialista.</small></div> }

createRoot(document.getElementById('root')).render(<AuthGate />)
