import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import { Document, HeadingLevel, Packer, Paragraph } from 'docx'
import pptxgen from 'pptxgenjs'
import { Html5Qrcode } from 'html5-qrcode'
import { GoogleAuthProvider, OAuthProvider, createUserWithEmailAndPassword, getRedirectResult, onAuthStateChanged, sendEmailVerification, sendPasswordResetEmail, signInWithCredential, signInWithEmailAndPassword, signInWithPopup, signOut, updateProfile } from 'firebase/auth'
import { FirebaseAuthentication } from '@capacitor-firebase/authentication'
import { Capacitor, registerPlugin } from '@capacitor/core'
import { Directory, Filesystem } from '@capacitor/filesystem'
import { Share } from '@capacitor/share'
import { doc, onSnapshot, setDoc } from 'firebase/firestore'
import { getDownloadURL, ref, uploadBytes } from 'firebase/storage'
import { httpsCallable } from 'firebase/functions'
import {
  Activity, AlertCircle, AlertTriangle, ArrowDownToLine, ArrowUpRight, ArrowUpToLine, BadgeCheck, BarChart3, Bell, BookOpen, Bot, Boxes, Brain, Camera,
  Check, CheckCircle2, ChevronDown, ChevronRight, CircleHelp, ClipboardCheck, ClipboardList, Clock, Cloud, Download, Eye, EyeOff,
  FileBarChart2, FileSpreadsheet, FileText, Filter, Flashlight, Globe, HeartHandshake, ImagePlus, Info, Layers, Loader2, Lock,
  LayoutDashboard, Leaf, ListChecks, LogOut, Mail, MapPin, Megaphone, Menu, MessageCircle, MessageSquare, MoreHorizontal,
  PackageCheck, Palette, Paperclip, PenLine, Phone, PieChart, Plus, QrCode, ReceiptText, RotateCcw, Search, Send,
  Settings2, ShieldCheck, ShoppingBag, SlidersHorizontal, Sparkles, Sprout, Star, SwitchCamera, Tags, Target, Trash2,
  TrendingUp, Truck, Upload, UserPlus, UserRound, Users, UsersRound, Wallet, Wheat, X, Zap, CalendarDays, ExternalLink
} from 'lucide-react'
import './styles.css'
import { firebaseAuth, firebaseConfigured, firebaseDb, firebaseFunctions, firebaseStorage } from './firebase'
import {
  ASSISTANT_TOOLS,
  DIRECT_AI_SYSTEM_PROMPT,
  EXCEL_SYSTEM_PROMPT,
  MEMORY_CATEGORIES,
  MEMORY_STORAGE_KEY,
  SALES_SYSTEM_PROMPT,
  buildDemoProducts,
  buildMemorySection,
  buildSalesSheets,
  buildWorkbook,
  extractPdfText,
  mergeImportedProducts,
  mergeImportedRecords,
  mergeWorkbookSheets,
  normalizeImportedProducts,
  normalizeImportedRecords,
  normalizeMemories,
  normalizeMemoryRecord,
  normalizeProduct,
  normalizeProducts,
  parseJsonBlock,
  parseWorkbookImport,
  productToRow,
  simulateSales,
  summarizeSimulation,
  workbookBlob
} from './ai-data.js'

const STORAGE_KEY = 'vitimax-products-v1'
const PROFILE_KEY = 'vitimax-profile-v1'
const SETTINGS_KEY = 'vitimax-settings-v1'
const CHAT_KEY = 'vitimax-assistant-chat-v1'
const MODULES_KEY = 'vitimax-modules-v1'
const EMPTY_MODULES = { pedidos: [], clientes: [], asesorias: [], contenido: [], ventas: [], plantas: [], proveedores: [] }

function readLocalJson(key, fallback) {
  try {
    const value = JSON.parse(localStorage.getItem(key) || 'null')
    return value ?? fallback
  } catch {
    return fallback
  }
}

// El historial guardado puede tener una consulta pendiente si la app se cerró
// antes de que la IA respondiera. Al reabrir se marca para reenviarla.
function restoreChatMessages(saved) {
  if (!Array.isArray(saved)) return []
  return saved
    .filter((item) => item && typeof item === 'object' && typeof item.text === 'string')
    .map((item) => item.pending ? { ...item, pending: false, text: 'La consulta anterior no se completó. Vuelve a enviarla, por favor.' } : item)
}

function persistableChat(messages) {
  return messages.slice(-150).map(({ file, blob, ...rest }) => {
    const copy = { ...rest, text: String(rest.text || '').slice(0, 10000) }
    if (Array.isArray(copy.tools)) {
      copy.tools = copy.tools.slice(-6).map((tool) => {
        if (typeof tool.result === 'string') {
          return { role: 'tool', name: tool.name, arguments: tool.arguments, result: tool.result.slice(0, 1500) }
        }
        const { registered, file: toolFile, ...result } = (tool.result && typeof tool.result === 'object') ? tool.result : {}
        return {
          role: 'tool',
          name: tool.name,
          arguments: tool.arguments,
          result: JSON.stringify(result).slice(0, 1500)
        }
      })
    }
    return copy
  })
}

// Firestore limita cada documento a 1 MB: se guarda la parte más reciente
// del chat que quepa con margen, sin perder la copia local completa.
function cloudChat(messages) {
  const list = []
  let size = 0
  for (const item of persistableChat(messages).reverse()) {
    const length = JSON.stringify(item).length
    if (size + length > 600000) break
    list.unshift(item)
    size += length
  }
  return list
}
const emptyProduct = {
  name: '', category: '', price: '', previousPrice: '', stock: '',
  format: '', description: '', usage: '', recommendations: '', features: '', image: ''
}

function normalize(value) {
  return String(value ?? '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
}

function greetingForNow() {
  const hour = new Date().getHours()
  if (hour < 12) return 'Buenos días'
  if (hour < 20) return 'Buenas tardes'
  return 'Buenas noches'
}

function makeId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID()
  return `vitimax-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
}

function money(value) {
  if (value === '' || value === null || value === undefined || Number.isNaN(Number(value))) return 'Pendiente'
  return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 }).format(Number(value))
}

// Un producto importado puede tener 0 en precio/stock porque el Excel venía incompleto.
// pendingFields evita mostrarlo como "sin stock" o precio real cuando en realidad falta el dato.
function productPending(product, field) {
  if (!product) return false
  if (Array.isArray(product.pendingFields)) return product.pendingFields.includes(field)
  const value = product[field]
  return value === '' || value === null || value === undefined
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
  document.body.appendChild(link)
  link.click()
  link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '')
    reader.onerror = () => reject(reader.error || new Error('No se pudo leer el archivo generado.'))
    reader.readAsDataURL(blob)
  })
}

// Guarda archivos de verdad en cada plataforma:
// - Android: plugin nativo SaveToDownloads -> carpeta pública "Descargas".
// - iPhone/iPad: Documents de la app + hoja de compartir (Guardar en Archivos).
// - Web: compartir archivos del sistema o descarga normal del navegador.
const SaveToDownloads = registerPlugin('SaveToDownloads')
const DevicePermissions = registerPlugin('DevicePermissions')

function safeFileName(filename) {
  return String(filename || 'archivo-vitimax').replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim() || 'archivo-vitimax'
}

async function saveBlob(blob, filename) {
  const name = safeFileName(filename)
  const mimeType = blob.type || 'application/octet-stream'
  const platform = Capacitor.getPlatform()

  if (platform === 'android') {
    try {
      // En Android 9 o anterior se pide el permiso de almacenamiento aquí.
      // En Android 10 o superior no hace falta y el plugin lo omite.
      const permission = await SaveToDownloads.ensurePermission().catch(() => ({ granted: true }))
      if (permission && permission.granted === false) {
        throw new Error(permission.message || 'Permiso de almacenamiento denegado. Actívalo en Ajustes > Aplicaciones > VITIMAX > Permisos.')
      }
      const data = await blobToBase64(blob)
      const saved = await SaveToDownloads.save({ fileName: name, mimeType, data })
      return { filename: name, location: saved?.path || `Descargas/${name}`, platform, blob, uri: saved?.uri, mimeType, native: true }
    } catch (nativeError) {
      // Respaldo: guarda en la caché de la app y abre el menú de compartir de Android.
      try {
        const data = await blobToBase64(blob)
        const written = await Filesystem.writeFile({ path: name, data, directory: Directory.Cache, recursive: true })
        await Share.share({ title: name, text: `Archivo generado por VITIMAX: ${name}`, url: written.uri, dialogTitle: `Guardar ${name}` })
        return { filename: name, location: 'Menú de compartir de Android (elige Archivos o Descargas)', platform, blob, uri: written.uri, mimeType, shared: true }
      } catch (shareError) {
        throw new Error(nativeError?.message || shareError?.message || 'No se pudo guardar el archivo en el teléfono.')
      }
    }
  }

  if (platform === 'ios') {
    const data = await blobToBase64(blob)
    await Filesystem.writeFile({ path: name, data, directory: Directory.Documents, recursive: true })
    let location = `Documentos/${name}`
    try {
      const uri = await Filesystem.getUri({ path: name, directory: Directory.Documents })
      await Share.share({
        title: name,
        text: `Archivo generado por VITIMAX: ${name}`,
        url: uri.uri,
        dialogTitle: `Guardar o compartir ${name}`
      })
      location = `Documentos/${name} (también puedes guardarlo en Archivos)`
    } catch {
      // El archivo ya quedó guardado en Documentos aunque se cancele el menú.
    }
    return { filename: name, location, platform, blob }
  }

  try {
    const file = new File([blob], name, { type: mimeType })
    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({ files: [file], title: name })
      return { filename: name, location: 'Compartido desde el navegador', platform, blob }
    }
  } catch (shareError) {
    if (shareError?.name === 'AbortError') return { filename: name, location: 'Guardado cancelado', platform, blob }
  }

  downloadBlob(blob, name)
  return { filename: name, location: 'Descargas del navegador', platform, blob }
}

function escapeHtml(value) {
  return String(value || '').replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]))
}

const LAST_EMAIL_KEY = 'vitimax-last-email'
// Huella del certificado con el que se firma el APK que entrega VITIMAX.
// Debe estar registrada en Firebase (Authentication → Google → huella SHA-1)
// para que el acceso con Google funcione dentro de la app Android.
const ANDROID_APP_SHA1 = 'F9:01:BC:C4:43:34:CE:CF:C6:DF:D8:66:C1:36:D7:14:E0:A1:02:12'

function scopedKey(base, uid) {
  return uid ? `${base}:${uid}` : base
}

function providerLabel(user) {
  const id = user?.providerData?.[0]?.providerId || user?.providerId || ''
  if (id === 'google.com') return 'Google'
  if (id === 'microsoft.com') return 'Microsoft'
  if (id === 'password') return 'Correo y contraseña'
  return id ? id.replace('.com', '') : 'Cuenta VITIMAX'
}

function hasPasswordProvider(user) {
  return Boolean(user?.providerData?.some((item) => item.providerId === 'password'))
}

// Mensaje claro cuando Google rechaza la app por falta de la huella SHA-1.
function isGoogleConfigError(error) {
  const code = String(error?.code ?? '')
  const message = String(error?.message ?? '')
  const blob = `${code} ${message}`.toLowerCase()
  return blob.includes('developer_error') || blob.includes('apiexception: 10') || blob.includes(': 10:')
}

async function compressImage(file, maxSide = 640, quality = 0.85) {
  const dataUrl = await imageFileToDataUrl(file)
  const image = await new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('No se pudo procesar la imagen. Prueba con otra foto.'))
    img.src = dataUrl
  })
  const scale = Math.min(1, maxSide / Math.max(image.width || 1, image.height || 1))
  const width = Math.max(1, Math.round((image.width || 1) * scale))
  const height = Math.max(1, Math.round((image.height || 1) * scale))
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) return dataUrl
  context.imageSmoothingEnabled = true
  context.imageSmoothingQuality = 'high'
  context.drawImage(image, 0, 0, width, height)
  try {
    return canvas.toDataURL('image/jpeg', quality)
  } catch {
    return dataUrl
  }
}

function passwordStrength(value) {
  const text = String(value || '')
  if (!text) return 0
  let score = 1
  if (text.length >= 10) score += 1
  if ((/[A-Z]/.test(text) && /[a-z]/.test(text)) || /\d/.test(text) || /[^A-Za-z0-9]/.test(text)) score += 1
  return Math.min(3, score)
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
                                : code === 'auth/requires-recent-login' ? 'Por seguridad debes volver a iniciar sesión para completar este cambio.'
                                  : code === 'auth/credential-already-in-use' ? 'Esa cuenta de Google/Microsoft ya está vinculada a otro usuario.'
                                    : code === 'auth/operation-not-supported-in-this-environment' ? 'Este acceso no está disponible en este dispositivo. Usa correo y contraseña o inténtalo desde el navegador.'
                                      : code === 'auth/missing-email' ? 'Escribe tu correo electrónico.'
                                        : code ? `Firebase no pudo completar el acceso (${code}). Revisa la configuración y vuelve a intentar.` : 'Firebase no devolvió un error identificable. Cierra y abre la aplicación nuevamente.'
}

function isNativeApp() {
  return Boolean(globalThis.Capacitor?.isNativePlatform?.())
}

// Permisos del teléfono: notificaciones, cámara y archivos/fotos.
async function requestDevicePermissions() {
  if (Capacitor.getPlatform() !== 'android') return null
  try {
    return await DevicePermissions.requestAll()
  } catch {
    return null
  }
}

async function sendDeviceNotification(title, body) {
  if (Capacitor.getPlatform() !== 'android') return false
  try {
    await DevicePermissions.notify({ title, body })
    return true
  } catch {
    return false
  }
}

function openDeviceSettings() {
  if (Capacitor.getPlatform() !== 'android') return
  DevicePermissions.openAppSettings().catch(() => {})
}

function permissionsMessage(status) {
  if (!status) return 'No se pudo comprobar el estado de los permisos.'
  if (status.allGranted) return 'Permisos activados: notificaciones, cámara y archivos.'
  const missing = []
  if (!status.notifications) missing.push('notificaciones')
  if (!status.camera) missing.push('cámara')
  if (!status.storage || !status.media) missing.push('archivos y fotos')
  return `Falta activar: ${missing.join(', ')}. Puedes hacerlo en Ajustes de la aplicación.`
}

const ENV_AI_URL = (import.meta.env.VITE_AI_API_URL || '').trim()
const ENV_AI_TOKEN = (import.meta.env.VITE_AI_TOKEN || '').trim()

function normalizeAiUrl(value) {
  return String(value || '').trim().replace(/\/+$/, '')
}

function isLocalAiUrl(url) {
  return /^https?:\/\/(127\.0\.0\.1|localhost|0\.0\.0\.0)(:\d+)?(\/|$)/i.test(url) || /^https?:\/\/\[::1\]/i.test(url)
}

function resolveAiUrl(settings) {
  const url = normalizeAiUrl(settings?.aiServerUrl || ENV_AI_URL)
  if (!url) return ''
  if (isLocalAiUrl(url)) {
    // 127.0.0.1/localhost solo sirven si la app corre en la misma máquina.
    const pageHost = typeof window !== 'undefined' ? window.location.hostname : ''
    const pageIsLocal = ['localhost', '127.0.0.1', '0.0.0.0', '[::1]'].includes(pageHost)
    if (isNativeApp() || !pageIsLocal) return ''
  }
  return url
}

function resolveAiToken(settings) {
  return String(settings?.aiToken || ENV_AI_TOKEN || '').trim()
}

async function fileToAttachment(file) {
  if (!file) return undefined
  if (file.size > 8 * 1024 * 1024) throw new Error('El archivo supera el límite de 8 MB para análisis.')
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
  const [header, data] = String(dataUrl).split(',')
  const attachment = {
    name: file.name,
    mimeType: header.match(/data:(.*);base64/)?.[1] || file.type || 'application/octet-stream',
    data
  }
  // Excel/CSV: se convierten a texto para que la IA pueda analizarlos de verdad.
  if (/\.(xlsx|xls|csv)$/i.test(file.name) || /sheet|csv|excel/i.test(attachment.mimeType)) {
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' })
      attachment.text = workbook.SheetNames.slice(0, 5).map((name) => {
        const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[name])
        return `HOJA "${name}":\n${csv.slice(0, 14_000)}${csv.length > 14_000 ? '\n[Contenido recortado]' : ''}`
      }).join('\n\n')
    } catch {
      attachment.text = ''
    }
  } else if (attachment.mimeType === 'application/pdf' || /\.pdf$/i.test(file.name)) {
    try {
      const extracted = await extractPdfText(file, 30)
      attachment.text = extracted.text
        ? `${extracted.text.slice(0, 22_000)}${extracted.text.length > 22_000 ? '\n[Contenido recortado]' : ''}`
        : ''
      attachment.pages = extracted.pages
      attachment.scanned = !extracted.text
    } catch (error) {
      attachment.text = ''
      attachment.pdfError = error?.message || 'No se pudo extraer el texto del PDF.'
    }
  } else if (attachment.mimeType.startsWith('text/') || /\.(txt|md|json|xml|html|css|js|ts)$/i.test(file.name)) {
    try { attachment.text = await file.text() } catch { attachment.text = '' }
  }
  return attachment
}

async function fetchWithTimeout(url, options, timeoutMs = 45_000) {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fetch(url, { ...options, signal: controller.signal })
  } finally {
    clearTimeout(timer)
  }
}

function aiErrorMessage(error, fallback) {
  if (error?.name === 'AbortError') return 'La IA tardó demasiado en responder. Se conservó la respuesta local.'
  const message = String(error?.message || '')
  if (/429|rate.?limit|too many requests|saturad/i.test(message)) return 'La IA gratuita está saturada en este momento. Espera unos segundos y vuelve a enviar el mensaje.'
  if (/401|invalid.*api.*key|no auth|unauthorized/i.test(message)) return 'La clave de IA no es válida. Revisa la clave en Configuración.'
  if (/402|insufficient|credit|quota/i.test(message)) return 'La cuenta de IA no tiene crédito disponible en este momento.'
  if (/403|forbidden/i.test(message)) return 'El servicio de IA bloqueó la solicitud. Inténtalo de nuevo en unos minutos.'
  if (/network|failed to fetch|fetch failed|conexión|offline/i.test(message)) return 'No hay conexión a internet o el servicio de IA no responde. Revisa Wi-Fi o datos móviles.'
  return message || fallback
}

const DIRECT_AI_MODELS = (import.meta.env.VITE_AI_MODELS || 'google/gemini-2.5-flash,anthropic/claude-3.5-sonnet,openai/gpt-4o-mini')
  .split(',')
  .map((item) => item.trim())
  .filter(Boolean)
  .slice(0, 3) // OpenRouter acepta como máximo 3 modelos en el arreglo "models".

function resolveDirectKey(settings) {
  return String(settings?.aiApiKey || import.meta.env.VITE_AI_DIRECT_KEY || '').trim()
}

function attachmentText(attachment) {
  if (!attachment) return ''
  if (typeof attachment.text === 'string' && attachment.text.trim()) return attachment.text
  const isText = attachment.mimeType?.startsWith('text/') || /\.(txt|md|json|csv|xml|js|ts|html|css)$/i.test(attachment.name || '')
  if (!isText || typeof attachment.data !== 'string') return ''
  try {
    const binary = atob(attachment.data)
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0))
    return new TextDecoder().decode(bytes)
  } catch {
    return ''
  }
}

function attachmentContentParts(attachment) {
  const parts = []
  if (!attachment) return parts
  const fileText = attachmentText(attachment)
  if (fileText) {
    parts.push({
      type: 'text',
      text: `ARCHIVO ADJUNTO "${attachment.name}":\n${fileText.slice(0, 24_000)}${fileText.length > 24_000 ? '\n[Contenido recortado]' : ''}`
    })
  }
  if (attachment.mimeType?.startsWith('image/')) {
    parts.push({ type: 'image_url', image_url: { url: `data:${attachment.mimeType};base64,${attachment.data}` } })
  } else if (attachment.mimeType === 'application/pdf' && attachment.data) {
    // OpenRouter acepta PDF como archivo adjunto nativo en los modelos compatibles.
    parts.push({ type: 'file', file: { filename: attachment.name || 'documento.pdf', file_data: `data:application/pdf;base64,${attachment.data}` } })
  }
  return parts
}

async function askOpenRouterChatOnce({ messages, apiKey, tools, toolChoice, maxTokens = 1500, temperature = 0.7, responseFormat }) {
  const body = {
    models: DIRECT_AI_MODELS,
    messages,
    temperature,
    max_tokens: maxTokens
  }
  if (tools?.length) {
    body.tools = tools
    body.tool_choice = toolChoice || 'auto'
  }
  if (responseFormat) body.response_format = responseFormat
  const send = () => fetchWithTimeout('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': typeof window !== 'undefined' ? window.location.origin : 'https://vitimax.app',
      'X-Title': 'VITIMAX'
    },
    body: JSON.stringify(body)
  }, 120_000)
  let response = await send()
  let payload = await response.json().catch(() => ({}))
  // Algunos modelos gratuitos no aceptan function calling: reintentamos sin herramientas.
  if (!response.ok && body.tools && /tool|function|unsupported|not support|invalid.*parameter/i.test(String(payload?.error?.message || ''))) {
    delete body.tools
    delete body.tool_choice
    response = await send()
    payload = await response.json().catch(() => ({}))
  }
  // Si el modelo no acepta PDF como archivo, se reintenta avisando que el PDF no se pudo adjuntar.
  if (!response.ok && JSON.stringify(body.messages).includes('"type":"file"') && /file|pdf|content|unsupported|invalid/i.test(String(payload?.error?.message || ''))) {
    body.messages = body.messages.map((item) => Array.isArray(item.content)
      ? {
        ...item,
        content: item.content.map((part) => part.type === 'file'
          ? { type: 'text', text: `ARCHIVO PDF ADJUNTO: ${part.file?.filename || 'documento.pdf'}. El texto del PDF ya está incluido en el mensaje; usa ese contenido para responder.` }
          : part)
      }
      : item)
    response = await send()
    payload = await response.json().catch(() => ({}))
  }
  if (!response.ok) throw new Error(payload?.error?.message || `La IA rechazó la solicitud (${response.status}).`)
  const message = payload.choices?.[0]?.message || {}
  const content = String(message.content || message.reasoning || '').trim()
  const toolCalls = Array.isArray(message.tool_calls) ? message.tool_calls : []
  if (!content && !toolCalls.length) throw new Error('La IA devolvió una respuesta vacía.')
  return { content, toolCalls, model: payload.model || '' }
}

async function askOpenRouterChat(options) {
  try {
    return await askOpenRouterChatOnce(options)
  } catch (error) {
    const message = String(error?.message || '')
    const retryable = /429|rate|limit|saturad|temporal|502|503|504|timeout|tardó|fetch/i.test(message)
    if (!retryable) throw error
    await new Promise((resolve) => setTimeout(resolve, 2500))
    return askOpenRouterChatOnce(options)
  }
}

async function askOpenRouterDirect({ message, products, history, attachment, apiKey, systemPrompt, maxTokens = 1500, tools, toolChoice, temperature = 0.7 }) {
  const catalog = JSON.stringify((products || []).map(({ id, name, category, price, stock, format, description, usage, recommendations, features }) => ({
    id, name, category, price, stock, format, description, usage, recommendations, features
  })))
  const content = [{
    type: 'text',
    text: `CATÁLOGO REAL DE VITIMAX:\n${catalog}\n\nCONSULTA DEL USUARIO:\n${message}`
  }, ...attachmentContentParts(attachment)]
  return askOpenRouterChat({
    messages: [
      { role: 'system', content: systemPrompt || DIRECT_AI_SYSTEM_PROMPT },
      ...(Array.isArray(history) ? history : []).map((item) => ({ role: item.role, content: item.content })),
      { role: 'user', content }
    ],
    apiKey,
    tools,
    toolChoice,
    maxTokens,
    temperature
  })
}

async function askOpenRouterMessages({ messages, apiKey, tools, toolChoice, maxTokens = 1500, responseFormat, temperature = 0.7 }) {
  return askOpenRouterChat({ messages, apiKey, tools, toolChoice, maxTokens, responseFormat, temperature })
}

const IMPORT_AI_SCHEMA_PROMPT = `Eres el importador universal de VITIMAX. Recibes las primeras filas de un Excel o CSV irregular y debes mapearlas al esquema de la aplicación.
Devuelve SOLO un objeto JSON válido con esta forma exacta (omite las claves que no tengan datos y usa [] cuando no existan registros):
{
  "productos": [{"id": "", "nombre": "", "categoria": "", "precio": 0, "precioAnterior": 0, "stock": 0, "formato": "", "descripcion": "", "uso": "", "recomendaciones": "", "caracteristicas": "", "imagen": ""}],
  "clientes": [{"id": "", "nombre": "", "telefono": "", "email": "", "localidad": "", "notas": ""}],
  "plantas": [{"id": "", "nombre": "", "cientifico": "", "tipo": "", "clima": "", "suelo": "", "riego": "", "temporada": "", "notas": ""}],
  "ventas": [{"id": "", "fecha": "", "cliente": "", "producto": "", "cantidad": 0, "precioUnitario": 0, "total": 0, "estado": "", "canal": ""}],
  "asesorias": [{"id": "", "cliente": "", "cultivo": "", "consulta": "", "estado": "", "prioridad": "", "recomendacion": ""}],
  "proveedores": [{"id": "", "proveedor": "", "contacto": "", "telefono": "", "email": "", "ciudad": "", "productos": "", "condiciones": ""}],
  "pedidos": [{"id": "", "fecha": "", "cliente": "", "producto": "", "cantidad": 0, "total": 0, "estado": "", "notas": ""}]
}
Reglas:
- Usa los datos reales de las filas. NO inventes nombres, precios ni stock.
- Si un producto no trae precio o stock, usa 0; VITIMAX lo marcará como "Pendiente de completar".
- Limpia símbolos de moneda ($, puntos, comas, CLP) y entrega números.
- Reconoce encabezados por sinónimos (nombre/producto/artículo, precio/valor/costo, stock/cantidad/existencias, categoría/tipo/familia, código/SKU, descripción/dosis/uso).
- No agregues explicaciones, markdown ni texto fuera del JSON.`

async function importWorkbookWithAi(rawText, fileName, settings) {
  if (!rawText) return null
  const prompt = `${IMPORT_AI_SCHEMA_PROMPT}\n\nARCHIVO: ${fileName}\n\nFILAS DETECTADAS:\n${rawText}`
  const directKey = resolveDirectKey(settings)
  if (directKey) {
    try {
      const result = await askOpenRouterMessages({
        messages: [
          { role: 'system', content: 'Respondes únicamente con JSON válido, sin markdown ni explicaciones.' },
          { role: 'user', content: prompt }
        ],
        apiKey: directKey,
        maxTokens: 3200,
        responseFormat: { type: 'json_object' },
        temperature: 0.2
      })
      const parsed = parseJsonBlock(result.content)
      if (parsed) return parsed
    } catch { /* se intenta el servidor opcional */ }
  }
  const aiUrl = resolveAiUrl(settings)
  if (aiUrl) {
    try {
      const response = await fetchWithTimeout(`${aiUrl}/api/assistant`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: prompt, history: [], maxTokens: 3200 })
      }, 120_000)
      const data = await response.json().catch(() => ({}))
      if (response.ok) {
        const parsed = parseJsonBlock(data.text || data.reply || '')
        if (parsed) return parsed
      }
    } catch { /* sin IA disponible */ }
  }
  return null
}

const GENERATION_SYSTEM_PROMPT = `Eres el generador de entregables de VITIMAX. Produces archivos reales, completos y listos para usar.
Reglas:
- Escribe siempre en español neutro.
- Si el catálogo tiene productos, úsalos como fuente principal y respeta sus datos reales.
- Si el catálogo está vacío o el usuario pide datos de ejemplo, productos, plantas, tierra, abonos o precios, GENERA contenido completo y realista en lugar de negarte o dejar el archivo vacío. Márcalo como "Generado por IA (ejemplo)".
- Nunca presentes un dato generado como información oficial de VITIMAX.
- Devuelve únicamente el contenido pedido, sin explicaciones antes ni después.
- PROHIBIDO incluir frases meta como "no tengo acceso", "parece que", "como modelo", "aquí tienes", "no puedo", disculpas, aclaraciones o comentarios sobre el proceso. Si falta información, escribe "Pendiente" dentro del propio contenido.
- El archivo debe contener solo el entregable final: nada de conversación, notas ni advertencias fuera del contenido solicitado.`

const PAGE_SYSTEM_PROMPT = `${GENERATION_SYSTEM_PROMPT}
- Para páginas web devuelve SOLO un documento HTML5 completo y autónomo (DOCTYPE, head, style y body).
- Todo el CSS y el JavaScript deben ir embebidos: no uses enlaces a CDN ni archivos externos.
- Diseño moderno, responsive, accesible y optimizado para móvil. Incluye una llamada a la acción.
- No uses bloques de markdown; la respuesta debe empezar con <!doctype html>.`

const CODE_SYSTEM_PROMPT = `${GENERATION_SYSTEM_PROMPT}
- Para código devuelve SOLO un bloque de código markdown con el lenguaje indicado, por ejemplo \`\`\`html.
- El código debe ser completo, ejecutable y sin marcadores de posición sin explicar.
- No agregues texto fuera del bloque de código.`

const DOCUMENT_SYSTEM_PROMPT = `${GENERATION_SYSTEM_PROMPT}
- Para documentos devuelve texto plano bien estructurado, sin markdown decorativo.
- Usa líneas que empiecen con "# " para el título, "## " para secciones y "- " para viñetas.
- Para presentaciones separa cada diapositiva con una línea que contenga exactamente ---
- La primera línea de cada diapositiva será su título.
- El documento debe empezar directamente con el título y contener solo información útil para el lector final.`

function extractHtml(text) {
  const source = String(text || '')
  const fenced = source.match(/```(?:html)?\s*([\s\S]*?)```/i)
  const candidate = (fenced ? fenced[1] : source).trim()
  const start = candidate.search(/<!doctype html|<html[\s>]/i)
  if (start === -1) return ''
  const end = candidate.toLowerCase().lastIndexOf('</html>')
  const html = end === -1 ? candidate.slice(start) : candidate.slice(start, end + 7)
  return /<body[\s>]/i.test(html) || /<h1[\s>]/i.test(html) ? html : ''
}

function extractCodeBlock(text, fallbackLanguage = 'txt') {
  const source = String(text || '').trim()
  const fenced = source.match(/```([a-z0-9+#.-]*)\s*\n?([\s\S]*?)```/i)
  if (fenced) return { language: (fenced[1] || fallbackLanguage).toLowerCase(), code: fenced[2].trim() }
  return { language: fallbackLanguage, code: source }
}

const CODE_EXTENSIONS = {
  html: 'html', htm: 'html', css: 'css', js: 'js', javascript: 'js', jsx: 'jsx', ts: 'ts', typescript: 'ts', tsx: 'tsx',
  json: 'json', python: 'py', py: 'py', php: 'php', java: 'java', kotlin: 'kt', swift: 'swift', c: 'c', cpp: 'cpp',
  'c++': 'cpp', cs: 'cs', csharp: 'cs', go: 'go', rust: 'rs', ruby: 'rb', sql: 'sql', bash: 'sh', shell: 'sh', sh: 'sh',
  powershell: 'ps1', ps1: 'ps1', xml: 'xml', yaml: 'yml', yml: 'yml', markdown: 'md', md: 'md', txt: 'txt'
}

async function signInWithNativeProvider(providerId) {
  const result = providerId === 'google.com'
    ? await FirebaseAuthentication.signInWithGoogle()
    : await FirebaseAuthentication.signInWithMicrosoft()
  if (!result.credential) throw new Error('El proveedor no devolvió credenciales.')
  const credential = providerId === 'google.com'
    ? GoogleAuthProvider.credential(result.credential.idToken, result.credential.accessToken)
    : new OAuthProvider(providerId).credential(result.credential.idToken, result.credential.accessToken)
  try {
    return await signInWithCredential(firebaseAuth, credential)
  } catch (error) {
    // Si la sesión de Firebase no se pudo crear, cerramos también la sesión nativa
    // para que el próximo intento empiece limpio.
    try { await FirebaseAuthentication.signOut() } catch { /* sin sesión nativa */ }
    throw error
  }
}

function AuthGate() {
  const [user, setUser] = useState(null)
  const [guest, setGuest] = useState(() => sessionStorage.getItem('vitimax-guest') === 'true')
  const [authLoading, setAuthLoading] = useState(Boolean(firebaseAuth))
  const [authError, setAuthError] = useState('')
  const [sessionNotice, setSessionNotice] = useState('')
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
  async function handleSignOut() {
    setAuthError('')
    setSessionNotice('Sesión cerrada correctamente. Puedes volver a iniciar sesión cuando quieras.')
    try { sessionStorage.removeItem('vitimax-guest') } catch { /* sessionStorage puede no estar disponible */ }
    setGuest(false)
    try {
      if (Capacitor.isNativePlatform() && Capacitor.isPluginAvailable('FirebaseAuthentication')) {
        await FirebaseAuthentication.signOut()
      }
    } catch { /* el plugin nativo puede no tener sesión activa */ }
    try {
      if (firebaseAuth) await signOut(firebaseAuth)
    } catch { /* la sesión local se limpia igualmente */ }
    setUser(null)
  }
  function enterAsGuest() {
    try { sessionStorage.setItem('vitimax-guest', 'true') } catch { /* sessionStorage puede no estar disponible */ }
    setGuest(true)
  }
  if (authLoading) return <div className="auth-loading"><span className="auth-loading-mark"><BrandMark size={40} /></span><strong>VITIMAX</strong><span>Verificando sesión...</span></div>
  if (guest) return <App guest onSignOut={handleSignOut} />
  if (!firebaseConfigured) return <AuthScreenV2 configured={false} onGuest={enterAsGuest} />
  if (!user) return <AuthScreenV2 configured initialError={authError} initialNotice={sessionNotice} onAuthenticated={setUser} onGuest={enterAsGuest} />
  return <App user={user} onSignOut={handleSignOut} />
}

function BrandMark({ size = 24 }) {
  return <svg width={size} height={size} viewBox="0 0 1024 1024" fill="none" aria-hidden="true" focusable="false">
    <path d="M318 316C424 336 490 424 514 744C420 664 340 534 318 316Z" fill="currentColor" />
    <path d="M706 316C600 336 534 424 510 744C604 664 684 534 706 316Z" fill="currentColor" opacity=".68" />
    <path d="M512 744V780" stroke="currentColor" strokeWidth="26" strokeLinecap="round" />
  </svg>
}

const SECTION_META = {
  Resumen: { icon: LayoutDashboard, tone: 'lime' },
  Informe: { icon: FileBarChart2, tone: 'teal' },
  Catálogo: { icon: ShoppingBag, tone: 'green' },
  Inventario: { icon: Boxes, tone: 'blue' },
  Pedidos: { icon: ClipboardList, tone: 'amber' },
  Clientes: { icon: UsersRound, tone: 'sky' },
  Ventas: { icon: TrendingUp, tone: 'green' },
  Plantas: { icon: Leaf, tone: 'mint' },
  Proveedores: { icon: Truck, tone: 'slate' },
  Asesorías: { icon: Sprout, tone: 'mint' },
  Contenido: { icon: Sparkles, tone: 'violet' },
  Configuración: { icon: Settings2, tone: 'slate' }
}

function statusTone(value) {
  const v = normalize(value)
  if (['entregado', 'entregada', 'resuelta', 'resuelto', 'publicado', 'activo', 'completado', 'en rango'].includes(v)) return 'green'
  if (['cancelado', 'cancelada', 'vencido', 'error', 'sin stock'].includes(v)) return 'red'
  if (['confirmado', 'enviado', 'en revision', 'en proceso', 'en camino'].includes(v)) return 'blue'
  if (['pendiente', 'abierta', 'borrador', 'stock bajo', 'revisar'].includes(v)) return 'amber'
  return 'slate'
}

function SectionHeader({ icon: Icon, tone = 'green', eyebrow, title, description, actions, aside }) {
  return <div className="page-heading section-heading">
    <div className="heading-main">
      <span className={`section-icon ${tone}`}><Icon size={23} strokeWidth={1.9} /></span>
      <div className="heading-copy">
        <p className="eyebrow">{eyebrow}</p>
        <h1>{title}{aside}</h1>
        {description && <p className="muted">{description}</p>}
      </div>
    </div>
    {actions && <div className="heading-actions">{actions}</div>}
  </div>
}

const MODULE_CONFIG = {
  Pedidos: {
    key: 'pedidos', icon: ClipboardList, tone: 'amber', eyebrow: 'OPERACIÓN · 04',
    title: 'Pedidos', singular: 'pedido', addLabel: 'Nuevo pedido', listTitle: 'Pedidos registrados',
    description: 'Registra cada pedido y sigue su estado desde la consulta hasta la entrega.',
    emptyTitle: 'Todavía no hay pedidos',
    emptyText: 'Crea el primer pedido o impórtalos desde el catálogo cuando empieces a vender.',
    features: [[ReceiptText, 'Pedido completo', 'Cliente, producto, cantidad, total y notas.'], [Truck, 'Seguimiento', 'Estados de pendiente a entregado.'], [Wallet, 'Control de importes', 'Suma automática de los pedidos registrados.']],
    fields: [
      { key: 'cliente', label: 'Cliente', required: true, placeholder: 'Nombre del cliente' },
      { key: 'producto', label: 'Producto', required: true, placeholder: 'Producto solicitado' },
      { key: 'cantidad', label: 'Cantidad', type: 'number', min: 1, placeholder: '1' },
      { key: 'total', label: 'Total', type: 'number', min: 0, placeholder: '0' },
      { key: 'estado', label: 'Estado', type: 'select', options: ['Pendiente', 'Confirmado', 'Enviado', 'Entregado', 'Cancelado'], default: 'Pendiente' },
      { key: 'notas', label: 'Notas', type: 'textarea', full: true, placeholder: 'Detalles del pedido (opcional)' }
    ],
    statusKey: 'estado',
    primary: (r) => r.cliente || 'Cliente sin nombre',
    secondary: (r) => [r.producto, r.cantidad ? `${r.cantidad} u.` : ''].filter(Boolean).join(' · ') || 'Producto pendiente',
    meta: (r) => r.total !== '' && r.total !== undefined ? money(r.total) : 'Total pendiente',
    stats: (rows) => [
      { icon: ClipboardList, label: 'Pedidos', value: rows.length, detail: 'registrados', color: 'orange' },
      { icon: Clock, label: 'Pendientes', value: rows.filter((r) => r.estado === 'Pendiente').length, detail: 'por confirmar', color: 'orange' },
      { icon: Truck, label: 'En proceso', value: rows.filter((r) => ['Confirmado', 'Enviado'].includes(r.estado)).length, detail: 'confirmados o enviados', color: 'blue' },
      { icon: Wallet, label: 'Importe total', value: money(rows.reduce((s, r) => s + (Number(r.total) || 0), 0)), detail: 'suma de pedidos', color: 'green' }
    ]
  },
  Clientes: {
    key: 'clientes', icon: UsersRound, tone: 'sky', eyebrow: 'RELACIONES · 05',
    title: 'Clientes', singular: 'cliente', addLabel: 'Nuevo cliente', listTitle: 'Directorio de clientes',
    description: 'Guarda los datos de tus clientes y ten su información siempre a mano.',
    emptyTitle: 'Tu directorio está vacío',
    emptyText: 'Añade tu primer cliente para organizar contactos, consultas y pedidos.',
    features: [[UserPlus, 'Ficha de cliente', 'Nombre, teléfono, correo y localidad.'], [Phone, 'Contacto rápido', 'Datos listos para llamar o escribir.'], [MessageCircle, 'Historial', 'Relaciona consultas y pedidos con cada cliente.']],
    fields: [
      { key: 'nombre', label: 'Nombre', required: true, placeholder: 'Nombre y apellido' },
      { key: 'telefono', label: 'Teléfono', placeholder: '+54 9 11 ...' },
      { key: 'email', label: 'Correo', type: 'email', placeholder: 'cliente@correo.com' },
      { key: 'localidad', label: 'Localidad', placeholder: 'Ciudad o zona' },
      { key: 'notas', label: 'Notas', type: 'textarea', full: true, placeholder: 'Preferencias, cultivos o información útil' }
    ],
    statusKey: null,
    primary: (r) => r.nombre || 'Cliente sin nombre',
    secondary: (r) => [r.localidad, r.telefono, r.email].filter(Boolean).join(' · ') || 'Sin datos de contacto',
    meta: (r) => r.localidad || 'Sin localidad',
    stats: (rows) => [
      { icon: UsersRound, label: 'Clientes', value: rows.length, detail: 'en el directorio', color: 'blue' },
      { icon: Phone, label: 'Con teléfono', value: rows.filter((r) => r.telefono).length, detail: 'contacto directo', color: 'green' },
      { icon: Mail, label: 'Con correo', value: rows.filter((r) => r.email).length, detail: 'para novedades', color: 'purple' },
      { icon: MapPin, label: 'Localidades', value: new Set(rows.map((r) => r.localidad).filter(Boolean)).size, detail: 'zonas registradas', color: 'orange' }
    ]
  },
  Ventas: {
    key: 'ventas', icon: TrendingUp, tone: 'green', eyebrow: 'COMERCIAL · 08',
    title: 'Ventas', singular: 'venta', addLabel: 'Nueva venta', listTitle: 'Ventas registradas',
    description: 'Controla ingresos, productos vendidos y el canal de cada operación.',
    emptyTitle: 'Todavía no hay ventas',
    emptyText: 'Registra la primera venta o impórtala desde un Excel con hojas de Ventas.',
    features: [[TrendingUp, 'Ingresos claros', 'Total por venta y suma acumulada.'], [ReceiptText, 'Detalle completo', 'Cliente, producto, cantidad y precio.'], [Wallet, 'Canales', 'Compara ventas por medio o vendedor.']],
    fields: [
      { key: 'fecha', label: 'Fecha', type: 'date' },
      { key: 'cliente', label: 'Cliente', required: true, placeholder: 'Nombre del cliente' },
      { key: 'producto', label: 'Producto', required: true, placeholder: 'Producto vendido' },
      { key: 'cantidad', label: 'Cantidad', type: 'number', min: 0, placeholder: '1' },
      { key: 'precioUnitario', label: 'Precio unitario', type: 'number', min: 0, placeholder: '0' },
      { key: 'total', label: 'Total', type: 'number', min: 0, placeholder: '0' },
      { key: 'estado', label: 'Estado', type: 'select', options: ['Pagada', 'Pendiente', 'Anulada'], default: 'Pagada' },
      { key: 'canal', label: 'Canal', placeholder: 'Local, online, terreno...' }
    ],
    statusKey: 'estado',
    primary: (r) => r.cliente || 'Cliente sin nombre',
    secondary: (r) => [r.producto, r.cantidad ? `${r.cantidad} u.` : '', r.canal].filter(Boolean).join(' · ') || 'Producto pendiente',
    meta: (r) => r.total !== '' && r.total !== undefined ? money(r.total) : 'Total pendiente',
    stats: (rows) => [
      { icon: TrendingUp, label: 'Ventas', value: rows.length, detail: 'registradas', color: 'green' },
      { icon: Wallet, label: 'Ingresos', value: money(rows.reduce((sum, r) => sum + (Number(r.total) || 0), 0)), detail: 'suma total', color: 'blue' },
      { icon: PackageCheck, label: 'Unidades', value: rows.reduce((sum, r) => sum + (Number(r.cantidad) || 0), 0), detail: 'productos vendidos', color: 'purple' },
      { icon: Clock, label: 'Pendientes', value: rows.filter((r) => r.estado === 'Pendiente').length, detail: 'por cobrar', color: 'orange' }
    ]
  },
  Plantas: {
    key: 'plantas', icon: Leaf, tone: 'mint', eyebrow: 'CULTIVOS · 09',
    title: 'Plantas', singular: 'planta', addLabel: 'Nueva planta', listTitle: 'Plantas y cultivos',
    description: 'Organiza especies, clima, suelo, riego y temporada de cada cultivo.',
    emptyTitle: 'Sin plantas registradas',
    emptyText: 'Agrega una planta o importa una hoja de Plantas desde Excel.',
    features: [[Leaf, 'Ficha botánica', 'Nombre científico y familia.'], [Sprout, 'Manejo', 'Suelo, riego y temporada.'], [BookOpen, 'Notas', 'Cuidados y observaciones.']],
    fields: [
      { key: 'nombre', label: 'Planta / cultivo', required: true, placeholder: 'Ej.: Tomate' },
      { key: 'cientifico', label: 'Nombre científico', placeholder: 'Solanum lycopersicum' },
      { key: 'tipo', label: 'Tipo', placeholder: 'Hortaliza, frutal, aromática...' },
      { key: 'clima', label: 'Clima', placeholder: 'Templado, cálido...' },
      { key: 'suelo', label: 'Suelo recomendado', placeholder: 'Franco, drenado...' },
      { key: 'riego', label: 'Riego', placeholder: 'Frecuencia y forma de riego' },
      { key: 'temporada', label: 'Temporada', placeholder: 'Primavera - verano' },
      { key: 'notas', label: 'Notas', type: 'textarea', full: true, placeholder: 'Cuidados, plagas o información útil' }
    ],
    statusKey: null,
    primary: (r) => r.nombre || 'Planta sin nombre',
    secondary: (r) => [r.cientifico, r.tipo, r.clima].filter(Boolean).join(' · ') || 'Ficha pendiente',
    meta: (r) => r.temporada || 'Sin temporada',
    stats: (rows) => [
      { icon: Leaf, label: 'Plantas', value: rows.length, detail: 'registradas', color: 'green' },
      { icon: Tags, label: 'Tipos', value: new Set(rows.map((r) => r.tipo).filter(Boolean)).size, detail: 'categorías', color: 'blue' },
      { icon: Sprout, label: 'Con nombre científico', value: rows.filter((r) => r.cientifico).length, detail: 'fichas botánicas', color: 'purple' },
      { icon: CalendarDays, label: 'Temporadas', value: new Set(rows.map((r) => r.temporada).filter(Boolean)).size, detail: 'épocas distintas', color: 'orange' }
    ]
  },
  Proveedores: {
    key: 'proveedores', icon: Truck, tone: 'slate', eyebrow: 'ABASTECIMIENTO · 10',
    title: 'Proveedores', singular: 'proveedor', addLabel: 'Nuevo proveedor', listTitle: 'Proveedores registrados',
    description: 'Guarda contactos, productos y condiciones de tus proveedores.',
    emptyTitle: 'Sin proveedores registrados',
    emptyText: 'Agrega un proveedor o importa una hoja de Proveedores desde Excel.',
    features: [[Truck, 'Contacto directo', 'Teléfono, correo y ciudad.'], [Boxes, 'Productos', 'Rubro o insumos que entrega.'], [Wallet, 'Condiciones', 'Pagos, plazos y entregas.']],
    fields: [
      { key: 'proveedor', label: 'Proveedor', required: true, placeholder: 'Nombre o empresa' },
      { key: 'contacto', label: 'Contacto', placeholder: 'Persona de contacto' },
      { key: 'telefono', label: 'Teléfono', placeholder: '+56 9 ...' },
      { key: 'email', label: 'Correo', type: 'email', placeholder: 'proveedor@correo.com' },
      { key: 'ciudad', label: 'Ciudad', placeholder: 'Ciudad o comuna' },
      { key: 'productos', label: 'Productos / rubro', placeholder: 'Fertilizantes, sustratos...' },
      { key: 'condiciones', label: 'Condiciones', type: 'textarea', full: true, placeholder: 'Pago, plazos, entregas' }
    ],
    statusKey: null,
    primary: (r) => r.proveedor || 'Proveedor sin nombre',
    secondary: (r) => [r.contacto, r.telefono, r.ciudad].filter(Boolean).join(' · ') || 'Sin datos de contacto',
    meta: (r) => r.productos || 'Rubro pendiente',
    stats: (rows) => [
      { icon: Truck, label: 'Proveedores', value: rows.length, detail: 'registrados', color: 'green' },
      { icon: Phone, label: 'Con teléfono', value: rows.filter((r) => r.telefono).length, detail: 'contacto directo', color: 'blue' },
      { icon: Mail, label: 'Con correo', value: rows.filter((r) => r.email).length, detail: 'para pedidos', color: 'purple' },
      { icon: MapPin, label: 'Ciudades', value: new Set(rows.map((r) => r.ciudad).filter(Boolean)).size, detail: 'zonas cubiertas', color: 'orange' }
    ]
  },
  Asesorías: {
    key: 'asesorias', icon: Sprout, tone: 'mint', eyebrow: 'ACOMPAÑAMIENTO · 06',
    title: 'Asesorías', singular: 'asesoría', addLabel: 'Nueva asesoría', listTitle: 'Consultas y asesorías',
    description: 'Recibe consultas, registra el cultivo y da seguimiento a cada recomendación.',
    emptyTitle: 'Sin asesorías registradas',
    emptyText: 'Registra la primera consulta para acompañar a tus clientes con información ordenada.',
    features: [[HeartHandshake, 'Consulta cercana', 'Cliente, cultivo y problema detectado.'], [Sprout, 'Seguimiento', 'Estados de abierta a resuelta.'], [Bot, 'Apoyo del BOT IA', 'Ideas y respuestas con el asistente VITIMAX.']],
    fields: [
      { key: 'cliente', label: 'Cliente', required: true, placeholder: 'Nombre del cliente' },
      { key: 'cultivo', label: 'Cultivo', placeholder: 'Ej.: tomate, vid, cítricos' },
      { key: 'consulta', label: 'Consulta', type: 'textarea', full: true, placeholder: 'Describe el problema o la consulta' },
      { key: 'estado', label: 'Estado', type: 'select', options: ['Abierta', 'En revisión', 'Resuelta'], default: 'Abierta' },
      { key: 'prioridad', label: 'Prioridad', type: 'select', options: ['Normal', 'Alta', 'Urgente'], default: 'Normal' }
    ],
    statusKey: 'estado',
    primary: (r) => r.cliente || 'Cliente sin nombre',
    secondary: (r) => [r.cultivo, r.consulta].filter(Boolean).join(' · ') || 'Consulta pendiente de describir',
    meta: (r) => r.prioridad || 'Normal',
    stats: (rows) => [
      { icon: Sprout, label: 'Asesorías', value: rows.length, detail: 'registradas', color: 'green' },
      { icon: MessageCircle, label: 'Abiertas', value: rows.filter((r) => r.estado === 'Abierta').length, detail: 'por atender', color: 'orange' },
      { icon: Clock, label: 'En revisión', value: rows.filter((r) => r.estado === 'En revisión').length, detail: 'en seguimiento', color: 'blue' },
      { icon: CheckCircle2, label: 'Resueltas', value: rows.filter((r) => r.estado === 'Resuelta').length, detail: 'completadas', color: 'purple' }
    ]
  },
  Contenido: {
    key: 'contenido', icon: Sparkles, tone: 'violet', eyebrow: 'COMUNICACIÓN · 07',
    title: 'Contenido', singular: 'contenido', addLabel: 'Nuevo contenido', listTitle: 'Páginas y publicaciones',
    description: 'Organiza páginas, promociones y publicaciones creadas con ayuda del asistente.',
    emptyTitle: 'Aún no hay contenido',
    emptyText: 'Crea una página, promoción o publicación y mantenla organizada en un solo lugar.',
    features: [[PenLine, 'Borradores', 'Guarda ideas antes de publicarlas.'], [Megaphone, 'Promociones', 'Prepara campañas y novedades.'], [Bot, 'Creación con IA', 'Pide textos y páginas al BOT VITIMAX.']],
    fields: [
      { key: 'titulo', label: 'Título', required: true, placeholder: 'Título de la página o publicación' },
      { key: 'tipo', label: 'Tipo', type: 'select', options: ['Página', 'Publicación', 'Promoción', 'Guía'], default: 'Página' },
      { key: 'estado', label: 'Estado', type: 'select', options: ['Borrador', 'En revisión', 'Publicado'], default: 'Borrador' },
      { key: 'descripcion', label: 'Descripción', type: 'textarea', full: true, placeholder: 'Resumen o texto principal' }
    ],
    statusKey: 'estado',
    primary: (r) => r.titulo || 'Contenido sin título',
    secondary: (r) => [r.tipo, r.descripcion].filter(Boolean).join(' · ') || 'Sin descripción',
    meta: (r) => r.tipo || 'Página',
    stats: (rows) => [
      { icon: Sparkles, label: 'Contenidos', value: rows.length, detail: 'creados', color: 'purple' },
      { icon: PenLine, label: 'Borradores', value: rows.filter((r) => r.estado === 'Borrador').length, detail: 'en preparación', color: 'orange' },
      { icon: Eye, label: 'Publicados', value: rows.filter((r) => r.estado === 'Publicado').length, detail: 'visibles', color: 'green' },
      { icon: Layers, label: 'Tipos', value: new Set(rows.map((r) => r.tipo).filter(Boolean)).size, detail: 'formatos usados', color: 'blue' }
    ]
  }
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
  return <main className="auth-screen"><section className="auth-card"><div className="auth-brand"><div className="brand-mark"><BrandMark size={22} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div>{!configured ? <div className="auth-setup"><CircleHelp size={28} /><h1>Accede a VITIMAX</h1><p>El acceso de administrador necesita configurar Firebase. Mientras tanto puedes entrar como invitado para consultar la aplicación.</p><small>El modo invitado es solo de lectura: no permite editar productos, importar Excel ni cambiar la configuración.</small><button className="btn primary auth-guest" onClick={onGuest}>Entrar como invitado</button></div> : <><p className="eyebrow">ACCESO DE ADMINISTRADOR</p><h1>{mode === 'login' ? 'Iniciar sesión' : 'Crear cuenta'}</h1><p className="auth-muted">Accede al centro de gestión de VITIMAX.</p><div className="social-actions"><button className="social-btn google" disabled={busy} onClick={() => providerLogin(new GoogleAuthProvider())}><GoogleBrandIcon /> Continuar con Google</button><button className="social-btn microsoft" disabled={busy} onClick={() => providerLogin(new OAuthProvider('microsoft.com'))}><MicrosoftBrandIcon /> Continuar con Microsoft</button></div><div className="auth-divider"><span>o con correo</span></div><form onSubmit={submit} className="auth-form"><label>Correo electrónico<input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" /></label><label>Contraseña<input type="password" required minLength="6" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete={mode === 'login' ? 'current-password' : 'new-password'} /></label>{error && <p className="auth-error">{error}</p>}<button className="btn primary auth-submit" disabled={busy}>{busy ? 'Conectando...' : mode === 'login' ? 'Iniciar sesión' : 'Registrarme'}</button></form><button className="auth-switch" onClick={() => { setMode(mode === 'login' ? 'register' : 'login'); setError('') }}>{mode === 'login' ? 'Crear cuenta de administrador' : 'Ya tengo una cuenta'}</button><button className="auth-guest-link" onClick={onGuest}>Continuar como invitado</button></>}</section></main>
}

function AuthScreenV2({ configured, initialError = '', initialNotice = '', onAuthenticated, onGuest }) {
  const [mode, setMode] = useState('login')
  const [email, setEmail] = useState(() => { try { return localStorage.getItem(LAST_EMAIL_KEY) || '' } catch { return '' } })
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [firstName, setFirstName] = useState('')
  const [lastName, setLastName] = useState('')
  const [birthDate, setBirthDate] = useState('')
  const [error, setError] = useState(initialError)
  const [notice, setNotice] = useState(initialNotice)
  const [busy, setBusy] = useState(false)
  useEffect(() => { if (initialError) setError(initialError) }, [initialError])
  useEffect(() => { if (initialNotice) setNotice(initialNotice) }, [initialNotice])

  function rememberEmail(value) {
    try { if (value) localStorage.setItem(LAST_EMAIL_KEY, value) } catch { /* almacenamiento no disponible */ }
  }

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
      rememberEmail(normalizedEmail)
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
    setNotice('')
    if (!firebaseAuth) return
    setBusy(true)
    try {
      provider.setCustomParameters?.({ prompt: 'select_account' })
      const result = isNativeApp()
        ? await signInWithNativeProvider(provider.providerId)
        : await Promise.race([
          signInWithPopup(firebaseAuth, provider),
          new Promise((_, reject) => setTimeout(() => reject({ code: 'auth/popup-timeout' }), 20000))
        ])
      if (result?.user?.email) rememberEmail(result.user.email)
      onAuthenticated?.(result?.user || null)
    } catch (authError) {
      if (isGoogleConfigError(authError)) {
        setError(`Google rechazó esta instalación porque falta autorizar el teléfono en Firebase. Mientras se activa, puedes entrar con correo y contraseña o con Microsoft. Huella SHA-1 que debes registrar: ${ANDROID_APP_SHA1}`)
      } else {
        setError(authErrorMessage(authError.code))
      }
    } finally { setBusy(false) }
  }

  if (!configured) return <main className="auth-screen"><section className="auth-card"><div className="auth-brand"><div className="brand-mark"><BrandMark size={22} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div><div className="auth-setup"><CircleHelp size={28} /><h1>Accede a VITIMAX</h1><p>El acceso de administrador necesita configurar Firebase. Mientras tanto puedes entrar como invitado para consultar la aplicación.</p><small>El modo invitado es solo de lectura.</small><button className="btn primary auth-guest" onClick={onGuest}>Entrar como invitado</button></div></section></main>

  const register = mode === 'register'
  const reset = mode === 'reset'
  const today = new Date().toISOString().slice(0, 10)
  const strength = passwordStrength(password)
  const strengthLabel = ['', 'Débil', 'Buena', 'Fuerte'][strength]
  function switchMode(next) { setMode(next); setError(''); setNotice('') }
  return <main className="auth-screen"><section className="auth-card">
    <div className="auth-brand"><div className="brand-mark"><BrandMark size={22} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div>
    {!reset && <div className="auth-tabs" role="tablist" aria-label="Acceso a VITIMAX">
      <button type="button" role="tab" aria-selected={!register} className={register ? '' : 'active'} onClick={() => switchMode('login')}>Iniciar sesión</button>
      <button type="button" role="tab" aria-selected={register} className={register ? 'active' : ''} onClick={() => switchMode('register')}>Crear cuenta</button>
    </div>}
    <div className="auth-heading">
      <p className="eyebrow">{reset ? 'RECUPERAR ACCESO' : 'ACCESO DE ADMINISTRADOR'}</p>
      <h1>{reset ? 'Cambiar contraseña' : register ? 'Crea tu cuenta' : 'Bienvenido de nuevo'}</h1>
      <p className="auth-muted">{reset ? 'Escribe tu correo y te enviaremos un enlace seguro para crear una nueva contraseña.' : register ? 'Usa tu cuenta real de Google, Microsoft o tu correo electrónico.' : 'Entra con tu cuenta real de Google, Microsoft o tu correo electrónico.'}</p>
    </div>
    {!reset && <>
      <div className="social-actions">
        <button type="button" className="social-btn google" disabled={busy} onClick={() => providerLogin(new GoogleAuthProvider())}><GoogleBrandIcon /><span>Continuar con Google</span></button>
        <button type="button" className="social-btn microsoft" disabled={busy} onClick={() => providerLogin(new OAuthProvider('microsoft.com'))}><MicrosoftBrandIcon /><span>Continuar con Microsoft</span></button>
      </div>
      <p className="auth-provider-note">Se abrirá tu cuenta real. Si es tu primera vez, se crea tu cuenta VITIMAX automáticamente.</p>
      <div className="auth-divider"><span>o con tu correo</span></div>
    </>}
    <form onSubmit={submit} className="auth-form">
      {register && <div className="auth-form-row">
        <label>Nombre<span className="auth-input"><UserRound size={16} /><input type="text" required value={firstName} onChange={(e) => setFirstName(e.target.value)} autoComplete="given-name" placeholder="Tu nombre" /></span></label>
        <label>Apellido<span className="auth-input"><UserRound size={16} /><input type="text" required value={lastName} onChange={(e) => setLastName(e.target.value)} autoComplete="family-name" placeholder="Tu apellido" /></span></label>
      </div>}
      <label>Correo electrónico<span className="auth-input"><Mail size={16} /><input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" inputMode="email" autoCapitalize="none" spellCheck={false} placeholder="tucorreo@ejemplo.com" /></span></label>
      {register && <label>Fecha de nacimiento<BirthdayPicker value={birthDate} onChange={setBirthDate} maxDate={today} /><small className="field-help">Elige el día en el calendario. No puede ser una fecha futura.</small></label>}
      {!reset && <label>Contraseña<span className="auth-input"><Lock size={16} /><input type={showPassword ? 'text' : 'password'} required minLength="6" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete={register ? 'new-password' : 'current-password'} placeholder={register ? 'Crea una contraseña' : 'Tu contraseña'} /><button type="button" className="password-toggle" onClick={() => setShowPassword((current) => !current)} aria-label={showPassword ? 'Ocultar contraseña' : 'Mostrar contraseña'}>{showPassword ? <EyeOff size={17} /> : <Eye size={17} />}</button></span>
        {register && password && <span className={`password-strength s${strength}`}><i /><i /><i /><em>{strengthLabel}</em></span>}
        {register && <small className="field-help">Mínimo 6 caracteres. Usa mayúsculas y números para una clave más segura.</small>}
      </label>}
      {error && <p className="auth-error" role="alert"><AlertCircle size={15} /><span>{error}</span></p>}
      {notice && <p className="auth-success" role="status"><CheckCircle2 size={15} /><span>{notice}</span></p>}
      <button className="btn primary auth-submit" disabled={busy}>{busy ? <><Loader2 className="auth-spin" size={16} /> Procesando…</> : reset ? 'Enviar enlace' : register ? 'Crear mi cuenta' : 'Iniciar sesión'}</button>
    </form>
    {!reset && !register && <button type="button" className="auth-forgot" onClick={() => switchMode('reset')}>¿Olvidaste tu contraseña?</button>}
    {reset && <button type="button" className="auth-switch" onClick={() => switchMode('login')}>Volver a iniciar sesión</button>}
    <div className="auth-guest-row"><span>¿Solo quieres mirar?</span><button type="button" className="auth-guest-link" onClick={onGuest}>Entrar como invitado</button></div>
    {register && <p className="auth-terms">Tu cuenta queda protegida por Firebase y puedes cerrar sesión cuando quieras.</p>}
  </section></main>
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

function GoogleBrandIcon() {
  // Logotipo oficial de Google con los cuatro colores de marca.
  // Trazado nítido y proporcionado para que se vea bien en cualquier pantalla.
  return <svg className="provider-logo google" viewBox="0 0 24 24" aria-hidden="true" focusable="false">
    <path fill="#4285F4" d="M21.35 12.27c0-.72-.06-1.42-.18-2.09H12v3.96h5.24a4.48 4.48 0 0 1-1.94 2.94v2.45h3.14c1.84-1.69 2.91-4.18 2.91-7.26Z" />
    <path fill="#34A853" d="M12 21.7c2.63 0 4.84-.87 6.45-2.36l-3.14-2.45c-.87.58-1.98.93-3.31.93-2.54 0-4.7-1.72-5.47-4.03H3.29v2.53A9.74 9.74 0 0 0 12 21.7Z" />
    <path fill="#FBBC05" d="M6.53 13.79A5.85 5.85 0 0 1 6.23 12c0-.62.11-1.23.3-1.79V7.68H3.29A9.74 9.74 0 0 0 2.25 12c0 1.57.38 3.05 1.04 4.32l3.24-2.53Z" />
    <path fill="#EA4335" d="M12 6.18c1.43 0 2.71.49 3.72 1.46l2.79-2.79C16.83 3.28 14.63 2.3 12 2.3a9.74 9.74 0 0 0-8.71 5.38l3.24 2.53C7.3 7.9 9.46 6.18 12 6.18Z" />
  </svg>
}
function MicrosoftBrandIcon() { return <svg className="provider-logo ms" viewBox="0 0 23 23" aria-hidden="true" focusable="false"><rect x="1" y="1" width="10" height="10" fill="#F25022" /><rect x="12" y="1" width="10" height="10" fill="#7FBA00" /><rect x="1" y="12" width="10" height="10" fill="#00A4EF" /><rect x="12" y="12" width="10" height="10" fill="#FFB900" /></svg> }

function App({ user, onSignOut, guest = false }) {
  // Cada cuenta guarda su copia local con su propio uid, así al cerrar sesión
  // y entrar con otra cuenta no se mezclan el perfil ni los datos.
  const uid = user?.uid || ''
  const localKey = (base) => scopedKey(base, uid)
  const [products, setProducts] = useState(() => readLocalJson(localKey(STORAGE_KEY), []))
  const [profile, setProfile] = useState(() => {
    const saved = readLocalJson(localKey(PROFILE_KEY), {})
    return {
      name: saved.name || user?.displayName || (guest ? 'Invitado' : 'Administrador'),
      email: saved.email || user?.email || '',
      phone: saved.phone || '',
      photo: saved.photo || user?.photoURL || ''
    }
  })
  const [settings, setSettings] = useState(() => readLocalJson(localKey(SETTINGS_KEY), { notifications: true, lowStock: 5, currency: 'ARS', aiServerUrl: '', aiToken: '' }))
  const [records, setRecords] = useState(() => ({ ...EMPTY_MODULES, ...readLocalJson(localKey(MODULES_KEY), {}) }))
  const [moduleModal, setModuleModal] = useState(null)
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
  const [messages, setMessages] = useState(() => {
    const saved = readLocalJson(localKey(CHAT_KEY), [{ role: 'assistant', text: `${greetingForNow()}, ${profile.name || 'administrador'}. Soy VITIMAX IA: puedo ayudarte con código, estudios, cálculos, redacción, negocios y toda la gestión de VITIMAX. También genero Excel, PDF, Word, PowerPoint y páginas web reales. ¿Qué resolvemos?` }])
    return restoreChatMessages(saved)
  })
  const [memories, setMemories] = useState(() => normalizeMemories(readLocalJson(localKey(MEMORY_STORAGE_KEY), [])))
  const [cloudReady, setCloudReady] = useState(!user || !firebaseDb)
  const [saving, setSaving] = useState(false)
  const fileRef = useRef(null)
  const messagesRef = useRef(messages)
  const memoriesRef = useRef(memories)
  const cloudSeededRef = useRef(false)

  useEffect(() => {
    if (!user || !firebaseDb) return undefined
    const workspace = doc(firebaseDb, 'users', user.uid)
    setCloudReady(false)
    cloudSeededRef.current = false
    return onSnapshot(workspace, (snapshot) => {
      const data = snapshot.data() || {}
      if (Array.isArray(data.products)) setProducts(data.products)
      if (data.profile) setProfile((current) => ({ ...current, ...data.profile }))
      if (data.settings) setSettings((current) => ({ ...current, ...data.settings }))
      if (data.modules) setRecords((current) => ({ ...current, ...data.modules }))
      // Al abrir la app se recupera el historial y la memoria guardados en la nube
      // solo si este dispositivo todavía no tiene una copia local más completa.
      if (!cloudSeededRef.current) {
        cloudSeededRef.current = true
        if (Array.isArray(data.assistantChat) && data.assistantChat.length > messagesRef.current.length) {
          const restored = restoreChatMessages(data.assistantChat)
          if (restored.length) { messagesRef.current = restored; setMessages(restored) }
        }
        if (Array.isArray(data.aiMemory) && data.aiMemory.length) {
          const restored = normalizeMemories([...memoriesRef.current, ...data.aiMemory])
          if (restored.length > memoriesRef.current.length) { memoriesRef.current = restored; setMemories(restored) }
        }
      }
      setCloudReady(true)
    }, () => { setCloudReady(true); setNotice('No se pudo sincronizar la nube; se usará una copia local.') })
  }, [user])
  useEffect(() => {
    localStorage.setItem(localKey(STORAGE_KEY), JSON.stringify(products))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { products }, { merge: true }).catch(() => setNotice('No se pudo guardar el catálogo en la nube.'))
  }, [products, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(localKey(PROFILE_KEY), JSON.stringify(profile))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { profile: { name: profile.name, email: profile.email, phone: profile.phone, photo: profile.photo } }, { merge: true }).catch(() => setNotice('No se pudo guardar el perfil en la nube.'))
  }, [profile, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(localKey(SETTINGS_KEY), JSON.stringify(settings))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { settings }, { merge: true }).catch(() => setNotice('No se pudo guardar la configuración en la nube.'))
  }, [settings, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(localKey(MODULES_KEY), JSON.stringify(records))
    if (user && firebaseDb && cloudReady) setDoc(doc(firebaseDb, 'users', user.uid), { modules: records }, { merge: true }).catch(() => setNotice('No se pudieron guardar los módulos en la nube.'))
  }, [records, user, cloudReady])
  useEffect(() => { messagesRef.current = messages }, [messages])
  useEffect(() => { memoriesRef.current = memories }, [memories])
  useEffect(() => {
    // Los archivos generados contienen Blob y no se pueden guardar en localStorage.
    // Se conserva el texto del chat (incluida la traza de herramientas) y se omiten los archivos.
    const stored = persistableChat(messages)
    localStorage.setItem(localKey(CHAT_KEY), JSON.stringify(stored))
    if (user && firebaseDb && cloudReady) {
      setDoc(doc(firebaseDb, 'users', user.uid), { assistantChat: cloudChat(messages), assistantChatUpdatedAt: new Date().toISOString() }, { merge: true })
        .catch(() => { /* la copia local sigue disponible */ })
    }
  }, [messages, user, cloudReady])
  useEffect(() => {
    localStorage.setItem(localKey(MEMORY_STORAGE_KEY), JSON.stringify(memories))
    if (user && firebaseDb && cloudReady) {
      setDoc(doc(firebaseDb, 'users', user.uid), { aiMemory: memories, aiMemoryUpdatedAt: new Date().toISOString() }, { merge: true })
        .catch(() => setNotice('No se pudo guardar la memoria de la IA en la nube; se conserva en este dispositivo.'))
    }
  }, [memories, user, cloudReady])
  useEffect(() => { if (notice) { const timer = setTimeout(() => setNotice(''), 3500); return () => clearTimeout(timer) } }, [notice])
  useEffect(() => {
    // En Android se piden todos los permisos la primera vez que se abre la app.
    if (Capacitor.getPlatform() !== 'android') return undefined
    const key = 'vitimax-permissions-asked-v1'
    if (localStorage.getItem(key)) return undefined
    localStorage.setItem(key, '1')
    let alive = true
    requestDevicePermissions().then((status) => {
      if (alive && status) setNotice(permissionsMessage(status))
    })
    return () => { alive = false }
  }, [])

  const categories = useMemo(() => ['Todas', ...new Set(products.map((p) => p.category).filter(Boolean))], [products])
  const filtered = useMemo(() => products.filter((p) => {
    const haystack = Object.values(p).join(' ').toLowerCase()
    return haystack.includes(query.toLowerCase()) && (category === 'Todas' || p.category === category)
  }), [products, query, category])
  const lowStock = products.filter((p) => !productPending(p, 'stock') && Number(p.stock) > 0 && Number(p.stock) <= 5)
  const missingData = products.filter((p) => !p.name || productPending(p, 'price') || productPending(p, 'stock'))

  function navigate(item) { setActive(item); setQuery(''); setMobileMenuOpen(false); setOverlay(null) }
  function saveModuleRecord(moduleKey, values, existing) {
    setRecords((current) => {
      const list = current[moduleKey] || []
      if (existing) return { ...current, [moduleKey]: list.map((row) => row.id === existing.id ? { ...row, ...values } : row) }
      return { ...current, [moduleKey]: [{ id: makeId(), createdAt: new Date().toISOString(), ...values }, ...list] }
    })
    setModuleModal(null)
    setNotice(existing ? 'Registro actualizado correctamente' : 'Registro guardado correctamente')
  }
  function deleteModuleRecord(moduleKey, id) {
    setRecords((current) => ({ ...current, [moduleKey]: (current[moduleKey] || []).filter((row) => row.id !== id) }))
    setNotice('Registro eliminado')
  }
  // --- Memoria persistente del BOT IA ---------------------------------------
  function applyMemories(next) {
    memoriesRef.current = next
    setMemories(next)
    return next
  }
  function saveMemory(categoria, dato) {
    const record = normalizeMemoryRecord({ categoria, contenido: dato })
    if (!record) return null
    const current = memoriesRef.current
    const next = normalizeMemories([
      record,
      ...current.filter((item) => normalize(item.contenido) !== normalize(record.contenido))
    ])
    applyMemories(next)
    return record
  }
  function updateMemory(id, categoria, contenido) {
    const current = memoriesRef.current
    const next = normalizeMemories(current.map((item) => item.id === id
      ? { ...item, categoria: categoria || item.categoria, contenido: String(contenido || item.contenido).slice(0, 1200), fecha_actualizacion: new Date().toISOString() }
      : item))
    applyMemories(next)
    return next.find((item) => item.id === id) || null
  }
  function forgetMemory(keyOrId, nuevoDato) {
    const needle = normalize(keyOrId)
    if (!needle) return { removed: 0, updated: 0 }
    const current = memoriesRef.current
    const match = current.find((item) => item.id === keyOrId)
      || current.find((item) => normalize(`${item.categoria} ${item.contenido}`).includes(needle))
    if (!match) return { removed: 0, updated: 0 }
    if (String(nuevoDato || '').trim()) {
      const updated = { ...match, contenido: String(nuevoDato).trim().slice(0, 1200), fecha_actualizacion: new Date().toISOString() }
      applyMemories(normalizeMemories(current.map((item) => item.id === match.id ? updated : item)))
      return { removed: 0, updated: 1, record: updated }
    }
    applyMemories(current.filter((item) => item.id !== match.id))
    return { removed: 1, updated: 0, record: match }
  }
  function clearMemories() {
    applyMemories([])
    setNotice('Memoria de la IA borrada')
  }
  async function downloadTemplate() {
    const worksheet = XLSX.utils.json_to_sheet([emptyProduct])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Productos')
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    try {
      const saved = await saveBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'plantilla-productos-vitimax.xlsx')
      setNotice(`Plantilla Excel guardada en ${saved.location || 'el dispositivo'}`)
    } catch (error) {
      setNotice(error?.message || 'No se pudo guardar la plantilla Excel.')
    }
  }
  async function exportProducts() {
    const worksheet = XLSX.utils.json_to_sheet(products)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Productos')
    const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    try {
      const saved = await saveBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'catalogo-vitimax.xlsx')
      setNotice(`Catálogo exportado en ${saved.location || 'el dispositivo'}`)
    } catch (error) {
      setNotice(error?.message || 'No se pudo exportar el catálogo.')
    }
  }
  async function generateAssistantExcel(template = false, aiSheets = null, aiProducts = null, filename = '', options = {}) {
    const baseProducts = Array.isArray(aiProducts) && aiProducts.length ? aiProducts : products
    const sheets = template
      ? mergeWorkbookSheets([], [], records, { allowDemo: false }).map((sheet) => ({ ...sheet, rows: [] }))
      : mergeWorkbookSheets(aiSheets, baseProducts, records, { allowDemo: !products.length, ...options })
    const workbook = buildWorkbook(sheets)
    const blob = workbookBlob(workbook)
    const name = filename || (template ? 'plantilla-vitimax-7-hojas.xlsx' : 'catalogo-vitimax-ia.xlsx')
    const saved = await saveBlob(blob, name)
    setNotice(template
      ? `Plantilla Excel con ${sheets.length} hojas guardada en ${saved.location || 'el dispositivo'}`
      : `Excel generado con ${sheets.length} hojas y guardado en ${saved.location || 'el dispositivo'}`)
    return { ...saved, filename: name, blob, sheets }
  }
  function generateDemoProducts() {
    const demoProducts = buildDemoProducts()
    setProducts((current) => {
      const existing = new Set(current.map((product) => normalize(product.name)))
      return [...current, ...demoProducts.filter((product) => !existing.has(normalize(product.name)))]
    })
    setActive('Catálogo')
    setNotice(`${demoProducts.length} productos DEMO generados y agregados al catálogo`)
    return demoProducts.length
  }
  function registerProductsInApp(list, replace = false) {
    const incoming = normalizeProducts(list, { generated: true })
    if (!incoming.length) return { added: 0, updated: 0, total: products.length }
    if (replace) {
      const marked = incoming.map((product) => ({ ...product, source: 'Generado por IA (ejemplo)' }))
      setProducts(marked)
      setActive('Catálogo')
      setNotice(`${marked.length} productos generados por IA reemplazaron el catálogo`)
      return { added: marked.length, updated: 0, total: marked.length }
    }
    let added = 0
    let updated = 0
    const map = new Map(products.map((product) => [normalize(product.name), product]))
    incoming.forEach((product) => {
      const key = normalize(product.name)
      if (!key) return
      const existing = map.get(key)
      if (existing) {
        map.set(key, { ...existing, ...product, id: existing.id, source: existing.source || product.source })
        updated += 1
      } else {
        // Todo producto nuevo que llega desde la IA se marca como generado.
        map.set(key, { ...product, source: 'Generado por IA (ejemplo)' })
        added += 1
      }
    })
    const next = [...map.values()]
    setProducts(next)
    setActive('Catálogo')
    setNotice(`${added} productos agregados · ${updated} actualizados en el catálogo`)
    return { added, updated, total: next.length }
  }
  async function generateAssistantFile(type, requestText, aiText = '', filenameOverride = '') {
    const title = String(requestText || '').trim().slice(0, 120) || 'Documento VITIMAX'
    const catalogSource = products.length ? products : buildDemoProducts()
    const catalogText = catalogSource.map((item) => [
      item.name || 'Producto pendiente',
      `Categoría: ${item.category || 'Pendiente'}`,
      `Formato: ${item.format || 'Pendiente'}`,
      `Precio: ${money(item.price)}`,
      `Stock: ${productPending(item, 'stock') ? 'Pendiente' : item.stock}`,
      `Descripción: ${item.description || 'Pendiente'}`
    ].join(' | ')).join('\n')
    const body = String(aiText || '').trim() || `Catálogo VITIMAX\n\n${catalogText}\n\n${products.length ? '' : 'Datos de ejemplo generados por IA. No son información oficial de VITIMAX.'}`
    let filename = filenameOverride || 'documento-vitimax'
    let saved = null
    if (type === 'pdf') {
      const pdf = new jsPDF()
      pdf.setFontSize(16)
      pdf.text(title, 18, 20)
      pdf.setFontSize(11)
      const lines = pdf.splitTextToSize(body, 175)
      let y = 32
      lines.forEach((line) => {
        if (y > 282) {
          pdf.addPage()
          y = 22
        }
        pdf.text(line, 18, y)
        y += 6.5
      })
      filename = filenameOverride || 'documento-vitimax.pdf'
      saved = await saveBlob(pdf.output('blob'), filename)
    } else if (type === 'word') {
      const children = [new Paragraph({ text: title, heading: HeadingLevel.TITLE })]
      body.split('\n').forEach((raw) => {
        const line = raw.trimEnd()
        if (!line.trim()) return
        if (line.startsWith('### ')) children.push(new Paragraph({ text: line.slice(4), heading: HeadingLevel.HEADING_2 }))
        else if (line.startsWith('## ')) children.push(new Paragraph({ text: line.slice(3), heading: HeadingLevel.HEADING_1 }))
        else if (line.startsWith('# ')) children.push(new Paragraph({ text: line.slice(2), heading: HeadingLevel.HEADING_1 }))
        else if (/^[-*•]\s+/.test(line)) children.push(new Paragraph({ text: line.replace(/^[-*•]\s+/, ''), bullet: { level: 0 } }))
        else children.push(new Paragraph(line))
      })
      const document = new Document({ sections: [{ children }] })
      filename = filenameOverride || 'documento-vitimax.docx'
      saved = await saveBlob(await Packer.toBlob(document), filename)
    } else if (type === 'ppt') {
      const presentation = new pptxgen()
      presentation.layout = 'LAYOUT_WIDE'
      const chunks = body.split(/\n\s*---\s*\n/).map((chunk) => chunk.trim()).filter(Boolean)
      const slides = chunks.length ? chunks : [body]
      slides.forEach((chunk, index) => {
        const slide = presentation.addSlide()
        slide.background = { color: index === 0 ? '173524' : 'F4F7F2' }
        const rows = chunk.split('\n').map((line) => line.trim()).filter(Boolean)
        const heading = (rows.shift() || `VITIMAX · ${index + 1}`).replace(/^#+\s*/, '')
        slide.addText(heading, { x: 0.7, y: 0.6, w: 12, h: 0.9, fontSize: 26, bold: true, color: index === 0 ? 'FFFFFF' : '173524' })
        const bullets = rows.map((line) => ({ text: line.replace(/^[-*•]\s+/, ''), options: { bullet: true, breakLine: true } }))
        if (bullets.length) {
          slide.addText(bullets, { x: 0.8, y: 1.7, w: 11.6, h: 4.7, fontSize: 16, color: index === 0 ? 'E8F2E9' : '173524', fit: 'shrink' })
        }
      })
      filename = filenameOverride || 'presentacion-vitimax.pptx'
      saved = await saveBlob(await presentation.write({ outputType: 'blob' }), filename)
    } else if (type === 'code') {
      const extracted = aiText ? extractCodeBlock(aiText) : { language: 'html', code: '' }
      const code = extracted.code || `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title><style>body{font-family:Arial,sans-serif;max-width:900px;margin:40px auto;padding:24px;color:#173524;background:#f4f7f2}section{background:#fff;padding:28px;border-radius:20px;box-shadow:0 10px 30px #17352418}li{margin:10px 0}</style></head><body><section><h1>${escapeHtml(title)}</h1><h2>Catálogo VITIMAX</h2><ul>${catalogSource.map((item) => `<li><strong>${escapeHtml(item.name || 'Producto pendiente')}</strong> — ${escapeHtml(item.description || 'Descripción pendiente.')}</li>`).join('')}</ul></section></body></html>`
      const extension = CODE_EXTENSIONS[extracted.language] || 'txt'
      filename = filenameOverride || `codigo-vitimax.${extension}`
      saved = await saveBlob(new Blob([code], { type: 'text/plain;charset=utf-8' }), filename)
    }
    setNotice(`${type === 'ppt' ? 'PowerPoint' : type === 'word' ? 'Word' : type.toUpperCase()} generado y guardado`)
    return saved || { filename, location: '' }
  }
  async function saveHtmlDocument({ titulo, html, filename }) {
    const body = String(html || '').trim()
    if (!body) throw new Error('La página HTML está vacía.')
    const slug = normalize(titulo || 'pagina-vitimax').replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '') || 'pagina-vitimax'
    const base = String(filename || '').trim().replace(/\.html?$/i, '') || `pagina-${slug}`
    const name = `${base}.html`
    const page = /<html[\s>]/i.test(body)
      ? body
      : `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(titulo || 'VITIMAX')}</title><style>body{font-family:Arial,sans-serif;max-width:960px;margin:40px auto;padding:24px;color:#173524;background:#f4f7f2;line-height:1.6}article{background:#fff;padding:32px;border-radius:20px;box-shadow:0 10px 30px #17352418}h1,h2{color:#173524}code,pre{background:#eef3ec;padding:2px 6px;border-radius:6px}</style></head><body><article>${body}</article></body></html>`
    const saved = await saveBlob(new Blob([page], { type: 'text/html;charset=utf-8' }), name)
    setNotice('Página HTML creada y guardada en el dispositivo')
    return saved
  }
  async function importExcel(event) {
    const file = event.target.files?.[0]
    if (!file) return
    const reset = () => { event.target.value = '' }
    if (file.size > 25 * 1024 * 1024) {
      setNotice(`El archivo "${file.name}" supera el límite de 25 MB.`)
      reset()
      return
    }
    const isCsv = /\.csv$/i.test(file.name) || /csv/i.test(file.type)
    try {
      setNotice(`Analizando "${file.name}"…`)
      const workbook = isCsv
        ? XLSX.read(await file.text(), { type: 'string', cellDates: true })
        : XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellDates: true })
      if (!workbook.SheetNames.length) throw new Error('El archivo no contiene hojas.')

      let parsed = parseWorkbookImport(workbook)
      let aiUsed = false

      // Fallback: si no se reconoció ninguna tabla estándar, el BOT IA mapea las filas crudas.
      if (parsed.needsAi) {
        setNotice('Estructura irregular detectada · el BOT IA está analizando las filas…')
        const aiRaw = await importWorkbookWithAi(parsed.rawText, file.name, settings)
        if (aiRaw) {
          const aiProducts = normalizeImportedProducts(aiRaw.productos || aiRaw.products || [])
          const aiRecords = {}
          ;['clientes', 'plantas', 'ventas', 'asesorias', 'proveedores', 'pedidos'].forEach((kind) => {
            aiRecords[kind] = normalizeImportedRecords(kind, aiRaw[kind] || [])
          })
          const recordTotal = Object.values(aiRecords).reduce((sum, list) => sum + list.length, 0)
          if (aiProducts.length || recordTotal) {
            aiUsed = true
            parsed = {
              ...parsed,
              products: aiProducts,
              records: Object.fromEntries(Object.keys(parsed.records).map((kind) => [kind, aiRecords[kind] || []])),
              pending: aiProducts.filter((product) => product.status === 'Pendiente de completar').length
            }
          }
        }
      }

      const productMerge = mergeImportedProducts(products, parsed.products)
      const nextRecords = { ...records }
      const recordTotals = { added: 0, updated: 0 }
      Object.entries(parsed.records).forEach(([kind, list]) => {
        if (!list?.length) return
        const merged = mergeImportedRecords(nextRecords[kind] || [], list, kind)
        nextRecords[kind] = merged.records
        recordTotals.added += merged.added
        recordTotals.updated += merged.updated
      })

      if (!parsed.sheets.length && !parsed.products.length && !Object.values(parsed.records).some((list) => list?.length)) {
        throw new Error(parsed.needsAi
          ? 'La estructura es irregular y el BOT IA no pudo analizarla. Revisa la conexión o configura una clave de IA en Configuración.'
          : 'No se reconoció ninguna tabla. Revisa que el archivo tenga encabezados como Nombre, Producto, Precio o Stock.')
      }

      setProducts(productMerge.products)
      setRecords(nextRecords)
      setImportReport({
        fileName: file.name,
        rows: parsed.rows || parsed.rawRows || 0,
        added: productMerge.added,
        updated: productMerge.updated,
        missing: parsed.pending,
        invalid: parsed.invalid,
        sheets: parsed.sheets,
        irregular: parsed.irregularSheets,
        recordsAdded: recordTotals.added,
        recordsUpdated: recordTotals.updated,
        aiUsed
      })

      const sectionByKind = { clientes: 'Clientes', plantas: 'Plantas', ventas: 'Ventas', asesorias: 'Asesorías', proveedores: 'Proveedores', pedidos: 'Pedidos' }
      const importedKind = Object.entries(parsed.records).find(([, list]) => list?.length)?.[0]
      setActive(parsed.products.length ? 'Catálogo' : sectionByKind[importedKind] || 'Resumen')

      const pieces = []
      if (productMerge.added) pieces.push(`${productMerge.added} producto(s) agregado(s)`)
      if (productMerge.updated) pieces.push(`${productMerge.updated} actualizado(s)`)
      if (recordTotals.added) pieces.push(`${recordTotals.added} registro(s) agregado(s)`)
      if (parsed.pending) pieces.push(`${parsed.pending} pendiente(s) de completar`)
      if (aiUsed) pieces.push('mapeado por el BOT IA')
      setNotice(pieces.length ? pieces.join(' · ') : `El archivo "${file.name}" se leyó, pero no aportó registros nuevos.`)
    } catch (error) {
      setNotice(`No se pudo importar "${file.name}": ${error?.message || 'formato no reconocido'}.`)
    } finally { reset() }
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
    let photoStoredLocally = false
    try {
      if (photoFile) {
        if (!photoFile.type.startsWith('image/')) throw new Error('El archivo seleccionado no es una imagen.')
        if (photoFile.size > 8 * 1024 * 1024) throw new Error('La foto debe pesar menos de 8 MB.')
        // Comprimimos la foto antes de subirla: así se ve nítida en el teléfono
        // y también funciona como respaldo si Storage no está disponible.
        const compressed = await compressImage(photoFile)
        if (firebaseStorage && user && !guest) {
          try {
            const photoRef = ref(firebaseStorage, `users/${user.uid}/profile/avatar-${Date.now()}.jpg`)
            const uploaded = await uploadBytes(photoRef, photoFile, { contentType: photoFile.type || 'image/jpeg' })
            normalized = { ...normalized, photo: await getDownloadURL(uploaded.ref) }
          } catch {
            // Firebase Storage puede no estar activado o las reglas no estar publicadas.
            // Guardamos la foto comprimida dentro del perfil para que igual se vea.
            normalized = { ...normalized, photo: compressed }
            photoStoredLocally = true
          }
        } else normalized = { ...normalized, photo: compressed }
      }
      if (firebaseAuth && user && !guest) {
        try { await updateProfile(user, { displayName: normalized.name || null, photoURL: normalized.photo || null }) } catch { /* la foto ya quedó guardada en el perfil */ }
      }
      localStorage.setItem(localKey(PROFILE_KEY), JSON.stringify(normalized))
      if (user && firebaseDb && cloudReady) {
        try { await setDoc(doc(firebaseDb, 'users', user.uid), { profile: normalized }, { merge: true }) } catch { photoStoredLocally = true }
      }
      setProfile(normalized)
      setOverlay(null)
      setNotice(photoStoredLocally && !guest
        ? 'Perfil y foto guardados. La foto quedó en tu cuenta VITIMAX.'
        : guest ? 'Perfil guardado en este dispositivo' : 'Perfil y foto guardados correctamente')
    } catch (error) {
      const message = error?.message || 'No se pudo actualizar el perfil. Revisa la sesión e inténtalo de nuevo.'
      setNotice(message)
      throw new Error(message)
    }
  }
  function saveSettings(next) { setSettings(next); setOverlay(null); setNotice('Configuración guardada correctamente') }
  async function resendVerificationEmail() {
    if (!firebaseAuth || !user || guest) { setNotice('Inicia sesión con una cuenta real para verificar tu correo.'); return }
    try {
      await sendEmailVerification(user)
      setNotice('Te enviamos un correo de verificación. Revisa tu bandeja y la carpeta de spam.')
    } catch (error) {
      setNotice(error?.code === 'auth/too-many-requests'
        ? 'Ya enviamos varios correos. Espera unos minutos antes de volver a intentar.'
        : 'No se pudo enviar el correo de verificación. Revisa tu conexión e inténtalo otra vez.')
    }
  }
  async function sendPasswordResetToUser() {
    const target = user?.email || profile.email
    if (!firebaseAuth || !target || guest) { setNotice('Esta cuenta no usa contraseña de VITIMAX.'); return }
    try {
      await sendPasswordResetEmail(firebaseAuth, target)
      setNotice(`Te enviamos un enlace para cambiar la contraseña a ${target}.`)
    } catch {
      setNotice('No se pudo enviar el enlace de contraseña. Revisa tu conexión e inténtalo otra vez.')
    }
  }
  function confirmSignOut() {
    setOverlay(null)
    setMobileMenuOpen(false)
    onSignOut?.()
  }
  async function downloadProductPage(product = null, requestText = '', aiHtml = '') {
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
    const safe = (value, fallback = '') => escapeHtml(value || fallback)
    const safeImage = pageData.image && /^(https?:|data:image\/)/i.test(pageData.image) ? escapeHtml(pageData.image) : ''
    const template = `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${safe(pageData.name, 'VITIMAX')}</title><style>body{font-family:Arial,sans-serif;margin:0;background:#f4f7f2;color:#173524}main{max-width:900px;margin:auto;padding:48px 24px}article{background:white;border-radius:24px;padding:32px;box-shadow:0 12px 40px #17352418}img{max-width:100%;max-height:360px;object-fit:contain;border-radius:16px}h1{font-size:clamp(2rem,5vw,4rem);margin:16px 0 8px}.tag{color:#4e8a52;text-transform:uppercase;font-weight:bold;letter-spacing:.12em}.price{font-size:1.6rem;font-weight:bold;margin:24px 0}a{display:inline-block;padding:14px 20px;border-radius:999px;background:#276b3b;color:#fff;text-decoration:none}</style></head><body><main><article>${safeImage ? `<img src="${safeImage}" alt="${safe(pageData.name, 'VITIMAX')}">` : ''}<div class="tag">${safe(pageData.category, 'VITIMAX')}</div><h1>${safe(pageData.name, 'VITIMAX')}</h1><p>${safe(pageData.description, 'Descripción pendiente de completar.')}</p>${pageData.price ? `<div class="price">${safe(money(pageData.price))}</div>` : ''}<p><strong>Formato:</strong> ${safe(pageData.format, 'Pendiente')}</p><p><strong>Modo de uso:</strong> ${safe(pageData.usage, 'Pendiente de ficha técnica.')}</p><p><strong>Disponibilidad:</strong> ${productPending(pageData, 'stock') ? 'Pendiente' : `${safe(pageData.stock)} unidades`}</p><a href="mailto:${safe(profile.email)}">Solicitar información</a></article></main></body></html>`
    const html = aiHtml || template
    const filename = `pagina-${slug}.html`
    const saved = await saveBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), filename)
    setNotice('Página web creada y guardada en el dispositivo')
    return saved
  }
  async function askAssistant(text, attachmentFile = null) {
    const promptNormalized = normalize(text)
    const analysisIntent = /analiz|revis|resum|explic|que dice|contenido/.test(promptNormalized)
    const explicitGeneration = /genera|generar|crea|crear|descarga|descargar|hazme|elabora|prepara|plantilla|exporta|simula|simular/.test(promptNormalized)
    const analysisOnly = Boolean(attachmentFile) && analysisIntent && !explicitGeneration
    const excelHelp = (promptNormalized.includes('ayuda') || promptNormalized.includes('ayudame') || promptNormalized.includes('organizar') || promptNormalized.includes('analizar')) && !explicitGeneration
    const excelRequested = !excelHelp && !analysisOnly && /excel|xlsx|hoja de calculo|hoja calculo|planilla/.test(promptNormalized)
    const templateRequested = excelRequested && /plantilla|vacio|vacío/.test(promptNormalized)
    const fileType = analysisOnly ? null
      : /pdf/.test(promptNormalized) ? 'pdf'
        : /word|docx/.test(promptNormalized) ? 'word'
          : /powerpoint|ppt|presentacion|diapositiva/.test(promptNormalized) ? 'ppt'
            : /codigo|programa/.test(promptNormalized) ? 'code'
              : null
    const pageRequested = !fileType && !analysisOnly && /crear pagina|pagina web|sitio web|landing|pagina promocional/.test(promptNormalized)
    const salesRequested = !fileType && !pageRequested && (/simul|proyecc|pronostico|pronóstico/.test(promptNormalized)
      || (/venta/.test(promptNormalized) && /recomend|estrategia|analiz|plan|mejorar|aumentar/.test(promptNormalized)))
    const pageTerms = promptNormalized.split(/\s+/).filter((term) => term.length > 3 && !['crear', 'pagina', 'web', 'promocional', 'promocion', 'sitio'].includes(term))
    const pageProduct = pageRequested && (products.length === 1 ? products[0] : products.find((item) => pageTerms.some((term) => normalize(item.name).includes(term))))
    const cloudAssistantEnabled = import.meta.env.VITE_ENABLE_CLOUD_ASSISTANT === 'true'
    const aiUrl = resolveAiUrl(settings)
    const aiToken = resolveAiToken(settings)
    const directKey = resolveDirectKey(settings)
    const modes = []
    if (aiUrl) modes.push('proxy')
    if (directKey) modes.push('direct')
    if (firebaseFunctions && cloudAssistantEnabled && !guest) modes.push('cloud')
    let workingProducts = products
    let attachment
    try {
      attachment = await fileToAttachment(attachmentFile)
    } catch (error) {
      setNotice(error?.message || 'No se pudo leer el archivo adjunto.')
      return
    }
    let attachmentNote = ''
    if (attachmentFile) {
      if (attachmentFile.size > 25 * 1024 * 1024) attachmentNote = `El archivo "${attachmentFile.name}" supera el límite local de 25 MB.`
      else if (attachment?.text) attachmentNote = `ARCHIVO ANALIZADO: ${attachmentFile.name}\n\n${attachment.text.slice(0, 1800)}${attachment.text.length > 1800 ? '\n\n[Contenido recortado]' : ''}`
      else if (attachment?.mimeType === 'application/pdf') attachmentNote = `PDF adjunto "${attachmentFile.name}" sin texto extraíble (posible escaneo). Adjunta una imagen de la página o pega el texto para analizarlo.`
      else if (/\.(docx?|pptx?)$/i.test(attachmentFile.name)) attachmentNote = `El archivo "${attachmentFile.name}" se adjuntó, pero Word/PowerPoint no se pueden leer directamente. Adjúntalo como PDF o pega el texto.`
      else if (attachment?.mimeType?.startsWith('image/')) attachmentNote = `Imagen adjunta "${attachmentFile.name}" lista para análisis.`
      else attachmentNote = `Archivo adjunto "${attachmentFile.name}" (${Math.ceil(attachmentFile.size / 1024)} KB).`
    }
    const userText = attachmentFile ? `${text}\n📎 ${attachmentFile.name}` : text
    const history = messages
      .filter((item) => !item.pending && item.text !== '…' && typeof item.text === 'string' && item.text.trim())
      .slice(-24)
      .flatMap((item) => {
        const rows = [{ role: item.role === 'tool' ? 'assistant' : item.role, content: item.text }]
        if (Array.isArray(item.tools)) {
          for (const tool of item.tools.slice(-6)) {
            const detail = typeof tool.result === 'string' ? tool.result : JSON.stringify(tool.result || {})
            rows.push({ role: 'assistant', content: `[Herramienta ${tool.name || 'IA'}] ${detail.slice(0, 500)}` })
          }
        }
        return rows
      })
    // La memoria se lee justo antes de cada consulta para incluir lo que la IA
    // acaba de guardar en el mismo turno o en sesiones anteriores.
    const memorySection = () => buildMemorySection(memoriesRef.current)
    const withMemory = (base) => `${base || DIRECT_AI_SYSTEM_PROMPT}${memorySection()}`
    const catalogForAi = (list = workingProducts) => JSON.stringify(list.map(({ id, name, category, price, stock, format, description, usage, recommendations, features, source }) => ({
      id, name, category, price, stock, format, description, usage, recommendations, features, source
    })))
    const recordsForAi = () => JSON.stringify({
      clientes: (records.clientes || []).slice(0, 30),
      pedidos: (records.pedidos || []).slice(0, 30),
      asesorias: (records.asesorias || []).slice(0, 30),
      contenido: (records.contenido || []).slice(0, 10)
    })
    const mergeProducts = (current, incoming) => {
      const map = new Map(current.map((product) => [normalize(product.name), product]))
      incoming.forEach((product) => {
        const key = normalize(product.name)
        if (!key) return
        const existing = map.get(key)
        map.set(key, existing ? { ...existing, ...product, id: existing.id } : product)
      })
      return [...map.values()]
    }
    const toFileExtra = (saved) => {
      if (!saved?.blob) return null
      const filename = saved.filename || saved.name || 'archivo-vitimax'
      return {
        name: filename,
        filename,
        blob: saved.blob,
        location: saved.location,
        uri: saved.uri,
        mimeType: saved.mimeType || saved.blob.type,
        native: saved.native,
        sheets: saved.sheets
      }
    }
    const productsFromSheets = (sheets) => {
      const sheet = (sheets || []).find((item) => normalize(item.name).startsWith('producto'))
      if (!sheet) return []
      const columns = sheet.columns || []
      return normalizeProducts((sheet.rows || []).map((row) => Object.fromEntries(columns.map((column, index) => [column, row[index]]))), { generated: true })
    }
    // Si piden simular ventas y el catálogo está vacío, se generan productos de ejemplo
    // para que la simulación y las recomendaciones tengan datos reales sobre los que trabajar.
    const salesDemoRequested = salesRequested && !workingProducts.length
      && /simul|proyecc|pronostico|ejemplo|demo|genera|crea|recomend/.test(promptNormalized)
    if (salesDemoRequested) {
      const demo = buildDemoProducts()
      if (demo.length) {
        registerProductsInApp(demo)
        workingProducts = demo
      }
    }
    const low = workingProducts.filter((product) => Number(product.stock) > 0 && Number(product.stock) <= Number(settings.lowStock))
    const terms = promptNormalized.split(/\s+/).filter((term) => term.length > 2)
    const matches = workingProducts.filter((product) => {
      const haystack = normalize(Object.values(product).join(' '))
      return !terms.length || terms.some((term) => haystack.includes(term))
    })
    const product = matches[0]
    const simulated = salesRequested ? simulateSales(workingProducts, { periodoDias: 30 }) : null
    let response = ''
    if (salesRequested && simulated) {
      response = `${summarizeSimulation(simulated)}${salesDemoRequested ? '\n\n📊 Generé productos de ejemplo porque el catálogo estaba vacío. Están marcados como "Generado por IA (ejemplo)".' : ''}`
    } else if (!workingProducts.length) {
      response = `El catálogo está vacío, pero no estoy bloqueado. Puedo generar productos, precios en CLP, plantas, tierra, abonos y datos de ejemplo, y crear Excel/PDF/Word/PowerPoint.\n\nPrueba con:\n• "Crea un catálogo de 10 productos con precios en CLP y genera el Excel con las 7 hojas"\n• "Genera un PDF con una guía de suelos y fertilizantes"\n• "Simula ventas de 30 días y dame recomendaciones"`
    } else if (promptNormalized.includes('inventario') || promptNormalized.includes('stock') || promptNormalized.includes('reporte')) {
      response = `INFORME DE INVENTARIO\n\nProductos registrados: ${workingProducts.length}\nUnidades contabilizadas: ${workingProducts.reduce((sum, item) => sum + (Number(item.stock) || 0), 0)}\nStock bajo: ${low.length}\n\n${low.length ? `Requieren atención: ${low.map((item) => `${item.name || 'Sin nombre'} (${item.stock})`).join(', ')}` : 'No hay productos dentro del umbral configurado.'}`
    } else if (promptNormalized.includes('descripcion') || promptNormalized.includes('publicacion') || promptNormalized.includes('redes') || promptNormalized.includes('contenido')) {
      response = product
        ? `BORRADOR DE CONTENIDO\n\n${product.name || 'Producto pendiente'}\n\n${product.description || 'La descripción todavía está pendiente de completar.'}\n\nDatos disponibles: ${[product.category, product.format, product.features].filter(Boolean).join(' · ') || 'Sin características registradas.'}\n\nPrecio: ${money(product.price)} · Stock: ${productPending(product, 'stock') ? 'Pendiente' : product.stock}`
        : 'Puedo redactar el contenido que necesites. Si quieres que use un producto del catálogo, escribe su nombre.'
    } else if (promptNormalized.includes('dosis') || promptNormalized.includes('aplicar') || promptNormalized.includes('tratamiento') || promptNormalized.includes('sirve')) {
      response = product
        ? `INFORMACIÓN DE ${product.name || 'PRODUCTO'}\n\nModo de uso: ${product.usage || 'Pendiente de ficha técnica.'}\nRecomendaciones: ${product.recommendations || 'Pendientes de ficha técnica.'}\n\nVerifica siempre la etiqueta oficial y consulta a un especialista antes de aplicar cualquier producto.`
        : 'Puedo darte una orientación general o generar una ficha de ejemplo. Confirma siempre la dosis con la etiqueta oficial y un especialista.'
    } else {
      response = `Encontré ${matches.length} coincidencia(s) en el catálogo.${matches.length ? `\n\n${matches.slice(0, 5).map((item) => `• ${item.name || 'Sin nombre'} · ${money(item.price)} · Stock: ${productPending(item, 'stock') ? 'Pendiente' : item.stock}`).join('\n')}` : '\n\nNo encontré coincidencias exactas. Puedo buscar en el catálogo, generar datos de ejemplo o crear el archivo que necesites.'}`
    }
    const localFallback = attachmentNote ? `${response}\n\n${attachmentNote}` : response
    const generateLocalFiles = async () => {
      const created = []
      if (excelRequested) {
        try {
          const saved = await generateAssistantExcel(templateRequested, [], [], '')
          created.push(saved)
          if (!workingProducts.length) {
            const fromSheet = productsFromSheets(saved.sheets)
            if (fromSheet.length) { registerProductsInApp(fromSheet); workingProducts = fromSheet }
          }
        } catch (error) { setNotice(error?.message || 'No se pudo generar el archivo Excel.') }
      }
      if (fileType) {
        try { created.push(await generateAssistantFile(fileType, text)) } catch (error) { setNotice(error?.message || 'No se pudo generar el archivo.') }
      }
      if (pageRequested) {
        try { created.push(await downloadProductPage(pageProduct, text)) } catch (error) { setNotice(error?.message || 'No se pudo generar la página web.') }
      }
      return created
    }
    const requestId = makeId()
    setMessages((items) => [...items, { role: 'user', text: userText }, { role: 'assistant', text: 'El asistente está pensando…', requestId, pending: true }])
    const replace = (nextText, extra = {}) => setMessages((items) => items.map((item) => item.requestId === requestId ? { ...item, text: nextText, pending: false, ...extra } : item))
    const finish = (messageText, created = [], registered = [], tools = []) => {
      const last = [...created].reverse().find((item) => item?.blob)
      const regNote = registered?.length ? `\n\n✅ ${registered.length} productos generados por IA quedaron guardados en el catálogo de VITIMAX.` : ''
      const textOut = `${String(messageText || 'Listo.').trim()}${regNote}`
      const extra = {}
      if (last) extra.file = toFileExtra(last)
      if (tools?.length) extra.tools = tools
      replace(textOut, extra)
    }
    const executeTool = async (name, args = {}) => {
      if (name === 'crear_excel') {
        const aiProducts = normalizeProducts(args.productos || [], { generated: true })
        const saved = await generateAssistantExcel(false, args.sheets || [], aiProducts, args.filename || '')
        let registered = []
        if (aiProducts.length) {
          registerProductsInApp(aiProducts, Boolean(args.reemplazar))
          workingProducts = args.reemplazar ? aiProducts : mergeProducts(workingProducts, aiProducts)
          registered = aiProducts
        }
        return {
          ok: true,
          archivo: saved.filename,
          ubicacion: saved.location,
          hojas: (saved.sheets || []).map((sheet) => sheet.name),
          resumen: args.resumen || '',
          productos_guardados: registered.length,
          file: toFileExtra(saved),
          registered
        }
      }
      if (name === 'registrar_productos_en_app') {
        const list = normalizeProducts(args.productos || [], { generated: true })
        if (!list.length) return { ok: false, error: 'No se recibieron productos válidos.' }
        const result = registerProductsInApp(list, Boolean(args.reemplazar))
        workingProducts = args.reemplazar ? list : mergeProducts(workingProducts, list)
        return { ok: true, agregados: result.added, actualizados: result.updated, total: workingProducts.length, registered: list }
      }
      if (name === 'simular_ventas') {
        const list = normalizeProducts(args.productos || [], { generated: true })
        const simulation = simulateSales(list.length ? list : workingProducts, {
          periodoDias: args.periodoDias,
          crecimientoMensual: args.crecimientoMensual
        })
        return { ok: true, ...simulation, resumen: summarizeSimulation(simulation) }
      }
      if (name === 'crear_documento') {
        const saved = await generateAssistantFile(args.tipo || 'pdf', args.titulo || text, args.contenido || '', args.filename || '')
        return { ok: true, archivo: saved.filename, ubicacion: saved.location, file: toFileExtra(saved) }
      }
      if (name === 'crear_pdf' || name === 'crear_word' || name === 'crear_ppt') {
        const tipo = name === 'crear_word' ? 'word' : name === 'crear_ppt' ? 'ppt' : 'pdf'
        const saved = await generateAssistantFile(tipo, args.titulo || text, args.contenido || '', args.filename || '')
        return { ok: true, archivo: saved.filename, ubicacion: saved.location, file: toFileExtra(saved) }
      }
      if (name === 'crear_html') {
        const saved = await saveHtmlDocument({ titulo: args.titulo || text, html: args.html || args.contenido || '', filename: args.filename || '' })
        return { ok: true, archivo: saved.filename, ubicacion: saved.location, file: toFileExtra(saved) }
      }
      if (name === 'guardar_memoria') {
        const record = saveMemory(args.categoria, args.dato || args.contenido || args.dato_a_recordar)
        if (!record) return { ok: false, error: 'El dato para recordar está vacío.' }
        return { ok: true, guardado: record, total: memoriesRef.current.length, mensaje: 'Dato guardado en la memoria persistente de VITIMAX.' }
      }
      if (name === 'olvidar_memoria') {
        const result = forgetMemory(args.clave_o_id || args.id || args.clave, args.nuevo_dato)
        if (!result.removed && !result.updated) return { ok: false, error: `No encontré un recuerdo que coincida con "${args.clave_o_id || args.clave || ''}".` }
        return { ok: true, ...result, total: memoriesRef.current.length, mensaje: result.updated ? 'Recuerdo actualizado.' : 'Recuerdo eliminado.' }
      }
      return { ok: false, error: `Herramienta no reconocida: ${name}` }
    }
    const callDirectAgent = async ({ message, systemPrompt = DIRECT_AI_SYSTEM_PROMPT, maxTokens = 3000, conversation = history, attachment: fileAttachment = attachment, tools = ASSISTANT_TOOLS, toolChoice = 'auto', temperature = 0.7 }) => {
      const messagesForAi = [
        { role: 'system', content: withMemory(systemPrompt) },
        ...conversation.map((item) => ({ role: item.role, content: item.content })),
        {
          role: 'user',
          content: [
            { type: 'text', text: `CATÁLOGO ACTUAL (${workingProducts.length} productos):\n${catalogForAi()}\n\nMÓDULOS Y REGISTROS:\n${recordsForAi()}\n\nSOLICITUD DEL USUARIO:\n${message}` },
            ...attachmentContentParts(fileAttachment)
          ]
        }
      ]
      const created = []
      const registered = []
      const toolTrace = []
      let lastText = ''
      for (let step = 0; step < 5; step += 1) {
        const result = await askOpenRouterMessages({
          messages: messagesForAi,
          apiKey: directKey,
          tools,
          toolChoice: step === 0 ? toolChoice : 'auto',
          maxTokens,
          temperature
        })
        lastText = result.content
        if (!result.toolCalls.length) return { text: result.content, created, registered, tools: toolTrace }
        messagesForAi.push({ role: 'assistant', content: result.content || '', tool_calls: result.toolCalls })
        for (const call of result.toolCalls) {
          let args = {}
          try { args = JSON.parse(call.function?.arguments || '{}') } catch { args = {} }
          let toolResult
          try {
            toolResult = await executeTool(call.function?.name, args)
          } catch (error) {
            toolResult = { ok: false, error: error?.message || 'No se pudo ejecutar la herramienta.' }
          }
          if (toolResult?.file) created.push(toolResult.file)
          if (Array.isArray(toolResult?.registered)) registered.push(...toolResult.registered)
          toolTrace.push({ role: 'tool', name: call.function?.name || 'herramienta', arguments: args, result: toolResult })
          messagesForAi.push({
            role: 'tool',
            tool_call_id: call.id || makeId('tool'),
            content: JSON.stringify(toolResult).slice(0, 14_000)
          })
        }
      }
      return { text: lastText || 'Listo. Ejecuté las acciones solicitadas.', created, registered, tools: toolTrace }
    }
    const callProxyOrCloud = async ({ message, systemPrompt, maxTokens = 2000, conversation = history, attachment: fileAttachment = attachment }) => {
      const remoteModes = modes.filter((mode) => mode !== 'direct')
      let lastError = null
      for (const mode of remoteModes) {
        try {
          let answer = ''
          if (mode === 'proxy') {
            const httpResponse = await fetchWithTimeout(`${aiUrl}/api/assistant`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json', ...(aiToken ? { 'X-Vitimax-Token': aiToken } : {}) },
              body: JSON.stringify({ message, products: workingProducts, attachment: fileAttachment, history: conversation, system: withMemory(systemPrompt), memories: memoriesRef.current, maxTokens })
            }, 90_000)
            const data = await httpResponse.json().catch(() => ({}))
            if (!httpResponse.ok) throw new Error(data.error || 'La IA no respondió.')
            answer = data.text
          } else {
            const request = httpsCallable(firebaseFunctions, 'vitimaxAssistant')({ message, products: workingProducts, attachment: fileAttachment, system: withMemory(systemPrompt), memories: memoriesRef.current, maxTokens })
            const timeout = new Promise((_, reject) => setTimeout(() => reject(new Error('La IA remota tardó demasiado.')), 30_000))
            const { data } = await Promise.race([request, timeout])
            answer = data.text
          }
          if (!answer || !String(answer).trim()) throw new Error('La IA devolvió una respuesta vacía.')
          return String(answer).trim()
        } catch (error) {
          lastError = error
        }
      }
      throw lastError || new Error('La IA no respondió.')
    }
    const generateText = async ({ message, systemPrompt, maxTokens = 3000, attachment: fileAttachment = attachment }) => {
      let lastError = null
      if (directKey) {
        try {
          const result = await askOpenRouterDirect({ message, products: workingProducts, history: [], attachment: fileAttachment, apiKey: directKey, systemPrompt: withMemory(systemPrompt), maxTokens })
          if (result.content) return result.content
        } catch (error) { lastError = error }
      }
      try {
        return await callProxyOrCloud({ message, systemPrompt, maxTokens, conversation: [], attachment: fileAttachment })
      } catch (error) { lastError = error }
      throw lastError || new Error('La IA no respondió.')
    }
    const workbookFromData = async (data) => {
      const aiProducts = normalizeProducts(data.productos || data.products || [], { generated: true })
      const saved = await generateAssistantExcel(false, data.sheets || data.hojas || [], aiProducts, data.filename || '')
      if (aiProducts.length) {
        registerProductsInApp(aiProducts, Boolean(data.reemplazar))
        workingProducts = data.reemplazar ? aiProducts : mergeProducts(workingProducts, aiProducts)
      }
      return { saved, text: data.resumen || data.summary || '', registered: aiProducts }
    }
    const generateWorkbookWithAI = async (requestText, fileAttachment) => {
      const instruction = `Genera el Excel completo para esta solicitud: "${requestText}".\n\nCATÁLOGO ACTUAL (${workingProducts.length} productos):\n${catalogForAi()}\n\nMÓDULOS Y REGISTROS:\n${recordsForAi()}\n\nInstrucciones obligatorias:\n- Usa la herramienta crear_excel.\n- Llena SIEMPRE las 7 hojas: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores.\n- Si el catálogo está vacío o el usuario pide ejemplos, genera datos completos y realistas (IDs, categorías, precios CLP, stock, dosis, descripciones) y márcalos como "Generado por IA (ejemplo)".\n- No uses "-" como dato: deja "" si falta.\n- Incluye en "productos" los productos que deban quedar guardados en el catálogo de la app.`
      if (directKey) {
        try {
          const result = await callDirectAgent({
            message: instruction,
            systemPrompt: EXCEL_SYSTEM_PROMPT,
            maxTokens: 6000,
            conversation: [],
            attachment: fileAttachment,
            tools: [ASSISTANT_TOOLS[0]],
            toolChoice: { type: 'function', function: { name: 'crear_excel' } },
            temperature: 0.35
          })
          if (result.created.length) return { saved: result.created[0], text: result.text, registered: result.registered }
          const data = parseJsonBlock(result.text)
          if (data?.sheets) return workbookFromData(data)
        } catch { /* se intenta el siguiente modo */ }
      }
      try {
        const answer = await callProxyOrCloud({ message: instruction, systemPrompt: EXCEL_SYSTEM_PROMPT, maxTokens: 6000, conversation: [], attachment: fileAttachment })
        const data = parseJsonBlock(answer)
        if (data?.sheets) return workbookFromData(data)
      } catch { /* respaldo local */ }
      const saved = await generateAssistantExcel(false, [], [], '')
      return { saved, text: '', registered: [] }
    }
    const runJsonAction = async (action) => {
      if (!action || typeof action !== 'object') return null
      const kind = normalize(action.accion || action.action || action.tipo || action.type || '')
      if (kind.includes('guardar_memoria') || action.guardar_memoria) {
        const record = saveMemory(action.categoria, action.dato || action.contenido)
        if (!record) return null
        return {
          text: `Listo, lo recordaré: ${record.contenido}`,
          created: [],
          registered: [],
          tools: [{ role: 'tool', name: 'guardar_memoria', arguments: action, result: { ok: true, guardado: record } }]
        }
      }
      if (kind.includes('olvidar_memoria') || action.olvidar_memoria) {
        const result = forgetMemory(action.clave_o_id || action.id || action.clave, action.nuevo_dato)
        return {
          text: result.updated ? 'Actualicé ese recuerdo en mi memoria.' : result.removed ? 'Listo, olvidé ese dato.' : 'No encontré un recuerdo que coincida.',
          created: [],
          registered: [],
          tools: [{ role: 'tool', name: 'olvidar_memoria', arguments: action, result }]
        }
      }
      if (action.sheets || action.hojas || kind.includes('excel')) {
        const built = await workbookFromData({
          ...action,
          sheets: action.sheets || action.hojas || [],
          productos: action.productos || action.products || []
        })
        return { text: built.text || `Excel generado con ${(built.saved.sheets || []).length} hojas.`, created: [built.saved], registered: built.registered }
      }
      if (kind.includes('registrar') || (action.productos && !action.sheets)) {
        const list = normalizeProducts(action.productos || action.products || [], { generated: true })
        if (!list.length) return null
        const result = registerProductsInApp(list, Boolean(action.reemplazar))
        workingProducts = action.reemplazar ? list : mergeProducts(workingProducts, list)
        return { text: action.resumen || `Guardé ${result.added} productos en el catálogo de VITIMAX.`, created: [], registered: list }
      }
      if (kind.includes('simular') || action.simulacion) {
        const simulation = simulateSales(workingProducts, {
          periodoDias: action.periodoDias,
          crecimientoMensual: action.crecimientoMensual
        })
        return { text: summarizeSimulation(simulation), created: [], registered: [] }
      }
      if (kind.includes('html') || action.html) {
        const saved = await saveHtmlDocument({ titulo: action.titulo || text, html: action.html || action.contenido || '', filename: action.filename || '' })
        return { text: action.resumen || 'Página web generada.', created: [saved], registered: [] }
      }
      if (kind.includes('pdf') || kind.includes('word') || kind.includes('ppt') || kind.includes('documento') || action.documento) {
        const doc = action.documento || {}
        const tipo = kind.includes('word') ? 'word' : kind.includes('ppt') ? 'ppt' : kind.includes('pdf') ? 'pdf' : doc.tipo || action.tipoArchivo || 'pdf'
        const saved = await generateAssistantFile(tipo, doc.titulo || action.titulo || text, doc.contenido || action.contenido || '', doc.filename || action.filename || '')
        return { text: action.resumen || 'Documento generado.', created: [saved], registered: [] }
      }
      return null
    }
    const stripActionJson = (value) => String(value || '').replace(/```(?:json)?\s*\{[\s\S]*?\}\s*```/gi, '').trim()
    void (async () => {
      try {
        if (!modes.length) {
          const localCreated = await generateLocalFiles()
          finish(localFallback, localCreated)
          return
        }
        if (templateRequested) {
          const saved = await generateAssistantExcel(true)
          finish(`✅ Plantilla Excel generada\n\nArchivo: ${saved.filename}\n${saved.location ? `Guardado en: ${saved.location}` : 'Se guardó en tu dispositivo.'}\n\nIncluye las 7 hojas con encabezados formateados y columnas autoajustadas.`, [saved])
          return
        }
        if (salesRequested && simulated) {
          let advice = ''
          try {
            advice = await generateText({
              message: `Analiza esta simulación de ventas y entrega recomendaciones comerciales concretas.\n\nSIMULACIÓN:\n${JSON.stringify(simulated)}\n\nSolicitud original: ${text}`,
              systemPrompt: SALES_SYSTEM_PROMPT,
              maxTokens: 2200,
              attachment: undefined
            })
          } catch { advice = '' }
          const created = []
          if (excelRequested || /archivo|descarg|planilla/.test(promptNormalized)) {
            try {
              const sheets = buildSalesSheets(simulated, workingProducts, records)
              const saved = await generateAssistantExcel(false, sheets, [], 'simulacion-ventas-vitimax.xlsx', { keepExtra: true })
              created.push(saved)
            } catch (error) { setNotice(error?.message || 'No se pudo generar el Excel de ventas.') }
          }
          const demoNote = salesDemoRequested ? '\n\n📊 El catálogo estaba vacío, así que generé productos de ejemplo marcados como "Generado por IA (ejemplo)" para poder simular. Puedes reemplazarlos por tus productos reales.' : ''
          finish(`${advice ? `${summarizeSimulation(simulated)}\n\n${advice}` : summarizeSimulation(simulated)}${demoNote}`, created)
          return
        }
        if (excelRequested) {
          // Si el usuario pide exportar SU catálogo real, se genera directo desde los datos
          // guardados, sin pedirle a la IA que invente hojas ni productos.
          const wantsOwnCatalog = /mi catalogo|mis productos|mi inventario|catalogo actual|mis datos|productos reales/.test(promptNormalized)
          if (wantsOwnCatalog && products.length) {
            const saved = await generateAssistantExcel(false, [], workingProducts, 'catalogo-vitimax.xlsx')
            const count = (saved.sheets || []).length || 7
            finish(`✅ Excel generado con tu catálogo real\n\nArchivo: ${saved.filename}\n${saved.location ? `Guardado en: ${saved.location}` : 'Se guardó en tu dispositivo.'}\n\nHojas incluidas: ${count}.\nProductos exportados: ${products.length}.\nClientes, pedidos y asesorías se incluyen si están registrados.`, [saved])
            return
          }
          const { saved } = await generateWorkbookWithAI(text, attachment)
          if (!workingProducts.length) {
            const fromSheet = productsFromSheets(saved.sheets)
            if (fromSheet.length) { registerProductsInApp(fromSheet); workingProducts = fromSheet }
          }
          const count = (saved.sheets || []).length || 7
          const productCount = workingProducts.length
          finish(`✅ Excel generado\n\nArchivo: ${saved.filename}\n${saved.location ? `Guardado en: ${saved.location}` : 'Se guardó en tu dispositivo.'}\n\nHojas incluidas: ${count}.${productCount ? `\nProductos en el catálogo: ${productCount}.` : '\nEl archivo incluye datos generados por IA claramente marcados como ejemplo.'}`, [saved])
          return
        }
        if (pageRequested || fileType) {
          const systemPrompt = pageRequested ? PAGE_SYSTEM_PROMPT : fileType === 'code' ? CODE_SYSTEM_PROMPT : DOCUMENT_SYSTEM_PROMPT
          const realData = catalogForAi()
          const instruction = pageRequested
            ? `Crea una página web completa para esta solicitud: "${text}".\nCatálogo actual: ${realData}.\nProducto principal: ${JSON.stringify(pageProduct || {})}.\nCorreo de contacto: ${profile.email || 'no configurado'}.\nSi no hay producto, crea una página general de VITIMAX con contenido de ejemplo claramente marcado.`
            : fileType === 'code'
              ? `Genera el código completo para esta solicitud: "${text}".\nSi el código usa datos de VITIMAX, usa estos datos: ${realData}.`
              : `Redacta el contenido completo para esta solicitud: "${text}".\nDatos de VITIMAX: ${realData}.\nSi el catálogo está vacío y se piden ejemplos, genera datos completos y márcalos como "Generado por IA (ejemplo)".\nTipo de archivo: ${fileType.toUpperCase()}.`
          let aiAnswer = await generateText({ message: instruction, systemPrompt, maxTokens: 4500 })
          let saved = null
          let kind = ''
          if (pageRequested) {
            let html = extractHtml(aiAnswer)
            if (!html) {
              aiAnswer = await generateText({
                message: `${instruction}\n\nDevuelve SOLO el HTML completo, empezando con <!doctype html> y terminando con </html>. Sin explicaciones.`,
                systemPrompt: PAGE_SYSTEM_PROMPT,
                maxTokens: 4500
              })
              html = extractHtml(aiAnswer)
            }
            if (!html) throw new Error('La IA no devolvió un HTML válido.')
            saved = await downloadProductPage(pageProduct, text, html)
            kind = 'Página web'
          } else {
            saved = await generateAssistantFile(fileType, text, aiAnswer)
            kind = fileType === 'pdf' ? 'PDF' : fileType === 'word' ? 'Word' : fileType === 'ppt' ? 'PowerPoint' : 'Código'
          }
          const generated = kind === 'Página web' ? 'generada' : 'generado'
          finish(`✅ ${kind} ${generated} con IA\n\nArchivo: ${saved.filename}\n${saved.location ? `Guardado en: ${saved.location}` : 'Se guardó en tu dispositivo.'}\n\nPuedes pedirme cambios y lo vuelvo a generar.`, [saved])
          return
        }
        if (directKey) {
          const result = await callDirectAgent({ message: text, systemPrompt: DIRECT_AI_SYSTEM_PROMPT, maxTokens: 3500, conversation: history })
          if (!result.created.length) {
            const action = parseJsonBlock(result.text)
            if (action) {
              const handled = await runJsonAction(action)
              if (handled) { finish(handled.text, handled.created, handled.registered, handled.tools); return }
            }
          }
          const answer = stripActionJson(result.text) || (result.created.length ? `✅ Archivo generado: ${result.created[result.created.length - 1].filename}` : 'Listo.')
          finish(answer, result.created, result.registered, result.tools)
          return
        }
        const answer = await callProxyOrCloud({ message: text })
        const action = parseJsonBlock(answer)
        if (action) {
          const handled = await runJsonAction(action)
          if (handled) { finish(handled.text, handled.created, handled.registered, handled.tools); return }
        }
        finish(stripActionJson(answer) || 'Listo.')
      } catch (error) {
        let localCreated = []
        try { localCreated = await generateLocalFiles() } catch { /* el aviso ya se muestra */ }
        finish(localFallback, localCreated)
        setNotice(aiErrorMessage(error, 'La IA no respondió. Se conservó la respuesta local.'))
      }
    })()
  }
  async function downloadMessageFile(file, index) {
    if (!file?.blob) return
    try {
      const saved = await saveBlob(file.blob, file.name)
      if (typeof index === 'number') {
        setMessages((items) => items.map((item, i) => i === index ? { ...item, file: { ...file, uri: saved.uri, mimeType: saved.mimeType || file.mimeType, location: saved.location, native: saved.native } } : item))
      }
      setNotice(`Archivo guardado en: ${saved.location || saved.filename}`)
    } catch (error) {
      setNotice(error?.message || 'No se pudo guardar el archivo. Revisa el permiso de almacenamiento.')
    }
  }
  async function openMessageFile(file) {
    if (!file?.uri) return
    try {
      await SaveToDownloads.openFile({ uri: file.uri, mimeType: file.mimeType || file.blob?.type || 'application/octet-stream' })
    } catch (error) {
      setNotice(error?.message || 'No se pudo abrir el archivo. Búscalo en la carpeta Descargas.')
    }
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
    setNotice(`El código "${value}" no está en el catálogo. Consultando al BOT IA…`)
    askAssistant(`Producto escaneado: ${value}`)
    setAssistantOpen(true)
  }

  const nav = Object.entries(SECTION_META).filter(([label]) => label !== 'Configuración').map(([label, meta]) => ({ label, ...meta }))
  const moduleCounts = {
    Pedidos: (records.pedidos || []).filter((row) => row.estado === 'Pendiente').length,
    Clientes: (records.clientes || []).length,
    Ventas: (records.ventas || []).length,
    Plantas: (records.plantas || []).length,
    Proveedores: (records.proveedores || []).length,
    Asesorías: (records.asesorias || []).filter((row) => row.estado === 'Abierta').length,
    Contenido: (records.contenido || []).filter((row) => row.estado === 'Borrador').length
  }
  const ActiveIcon = SECTION_META[active]?.icon || LayoutDashboard
  const activeTone = SECTION_META[active]?.tone || 'green'
  const canManage = !guest
  return <div className={guest ? 'app-shell guest-mode' : 'app-shell'}>
    <aside className={mobileMenuOpen ? 'sidebar mobile-open' : 'sidebar'}>
      <div className="brand"><div className="brand-mark"><BrandMark size={21} /></div><div><strong>VITIMAX</strong><span>Centro de gestión</span></div></div>
      <div className="workspace-label">ESPACIO DE TRABAJO <span className="status-dot" /></div>
      <nav>{nav.map(({ label, icon: Icon, tone }) => <button key={label} className={active === label ? 'nav-item active' : 'nav-item'} onClick={() => navigate(label)}><span className={`nav-icon ${tone}`}><Icon size={17} /></span><span>{label}</span>{moduleCounts[label] > 0 && <em>{moduleCounts[label]}</em>}</button>)}</nav>
      <div className="sidebar-bottom">{canManage && <button className="nav-item" onClick={() => { setOverlay('settings'); setMobileMenuOpen(false) }}><span className="nav-icon slate"><Settings2 size={17} /></span><span>Configuración</span></button>}<button className="nav-item signout-item" onClick={() => { setOverlay(null); setMobileMenuOpen(false); onSignOut?.() }}><span className="nav-icon rose"><LogOut size={17} /></span><span>Cerrar sesión</span></button><button className="profile profile-button" onClick={() => setOverlay('profile')}><Avatar name={profile.name} photo={profile.photo} /><div><strong>{profile.name || 'Invitado'}</strong><span>{profile.email || 'Acceso de invitado'}</span></div><MoreHorizontal size={17} /></button></div>
    </aside>
    <main className="main">
      <header className="topbar"><button className="mobile-menu" aria-label="Abrir menú" onClick={() => setMobileMenuOpen(true)}><Menu size={21} /></button><div className="crumb"><span className={`crumb-icon ${activeTone}`}><ActiveIcon size={14} /></span> VITIMAX <ChevronRight size={14} /> <b>{active}</b></div><div className="top-actions"><button className="icon-btn" title="Escanear código QR" aria-label="Escanear código QR" onClick={() => setQrOpen(true)}><QrCode size={18} /></button><button className="icon-btn" title="Notificaciones" aria-label="Notificaciones" onClick={() => setOverlay('notifications')}><Bell size={18} /><i /></button><button className="user-chip" onClick={() => setOverlay('profile')}><Avatar name={profile.name} photo={profile.photo} className="small" /><span>{profile.name || 'Administrador'}</span><ChevronRight size={15} /></button></div></header>
      <div className="content">
        {active === 'Resumen' && <Dashboard products={products} lowStock={lowStock} missingData={missingData} onImport={canManage ? () => fileRef.current?.click() : undefined} onTemplate={canManage ? downloadTemplate : undefined} onNavigate={navigate} records={records} />}
        {active === 'Informe' && <ApplicationReport products={products} lowStock={lowStock} missingData={missingData} guest={guest} />}
        {active === 'Catálogo' && <Catalog products={filtered} allProducts={products} categories={categories} query={query} setQuery={setQuery} category={category} setCategory={setCategory} onImport={canManage ? () => fileRef.current?.click() : undefined} onExport={canManage ? exportProducts : undefined} onTemplate={canManage ? downloadTemplate : undefined} onCreate={canManage ? () => { setEditing(null); setShowProduct(true) } : undefined} onEdit={canManage ? (p) => { setEditing(p); setShowProduct(true) } : undefined} onDelete={canManage ? (id) => setProducts((all) => all.filter((p) => p.id !== id)) : undefined} readOnly={guest} />}
        {active === 'Inventario' && <Inventory products={products} lowStock={lowStock} onImport={canManage ? () => fileRef.current?.click() : undefined} readOnly={guest} />}
        {['Pedidos', 'Clientes', 'Asesorías', 'Contenido', 'Ventas', 'Plantas', 'Proveedores'].includes(active) && <ModulePage config={MODULE_CONFIG[active]} rows={records[MODULE_CONFIG[active].key] || []} onCreate={canManage ? () => setModuleModal({ module: active, record: null }) : undefined} onEdit={canManage ? (record) => setModuleModal({ module: active, record }) : undefined} onDelete={canManage ? (id) => deleteModuleRecord(MODULE_CONFIG[active].key, id) : undefined} readOnly={guest} onAssistant={active === 'Asesorías' || active === 'Contenido' ? () => setAssistantOpen(true) : undefined} />}
      </div>
    </main>
    <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" hidden onChange={importExcel} />
    {notice && <div className="toast"><PackageCheck size={18} />{notice}</div>}
    {mobileMenuOpen && <button className="drawer-backdrop" aria-label="Cerrar menú" onClick={() => setMobileMenuOpen(false)} />}
    {overlay === 'profile' && <ProfileModal profile={profile} user={user} guest={guest} onClose={() => setOverlay(null)} onSave={saveProfile} onSignOut={confirmSignOut} onResendVerification={resendVerificationEmail} onPasswordReset={sendPasswordResetToUser} />}
    {overlay === 'settings' && <SettingsModal settings={settings} onClose={() => setOverlay(null)} onSave={saveSettings} />}
    {overlay === 'notifications' && <NotificationsModal lowStock={lowStock} missingData={missingData} onClose={() => setOverlay(null)} />}
    {importReport && <ImportReport report={importReport} onClose={() => setImportReport(null)} />}
    <button className="assistant-fab" onClick={() => setAssistantOpen(true)} aria-label="Abrir BOT IA VITIMAX"><img className="assistant-bot-image" src="/bot-vitimax.png" alt="" /><span>BOT IA VITIMAX</span></button>
    {assistantOpen && <Assistant messages={messages} onClose={() => setAssistantOpen(false)} onAsk={askAssistant} onDownload={downloadMessageFile} onOpen={openMessageFile} aiReady={Boolean(resolveAiUrl(settings) || resolveDirectKey(settings) || (firebaseFunctions && import.meta.env.VITE_ENABLE_CLOUD_ASSISTANT === 'true' && !guest))} memories={memories} onSaveMemory={saveMemory} onDeleteMemory={(id) => forgetMemory(id)} onClearMemories={clearMemories} />}
    {qrOpen && <QrScanner onClose={() => setQrOpen(false)} onResult={handleQrResult} />}
    {moduleModal && canManage && <ModuleFormModal config={MODULE_CONFIG[moduleModal.module]} record={moduleModal.record} onClose={() => setModuleModal(null)} onSave={(values) => saveModuleRecord(MODULE_CONFIG[moduleModal.module].key, values, moduleModal.record)} />}
    {showProduct && canManage && <ProductModal product={editing || emptyProduct} onClose={() => setShowProduct(false)} onSave={saveProduct} onDelete={deleteProduct} saving={saving} />}
  </div>
}

function Dashboard({ products, lowStock, missingData, onImport, onTemplate, onNavigate, records = EMPTY_MODULES }) {
  const readOnly = !onImport
  const totalStock = products.reduce((sum, p) => sum + (Number(p.stock) || 0), 0)
  const categories = new Set(products.map((p) => p.category).filter(Boolean))
  const steps = [
    { title: 'Importar catálogo oficial', done: products.length > 0, detail: products.length ? 'Completado' : 'Pendiente' },
    { title: 'Completar información', done: products.length > 0 && missingData.length === 0, detail: missingData.length ? `${missingData.length} ficha(s) pendientes` : products.length ? 'Completado' : 'Pendiente' },
    { title: 'Configurar categorías', done: categories.size > 0, detail: categories.size ? `${categories.size} categoría(s)` : 'Pendiente' },
    { title: 'Revisar niveles mínimos', done: products.length > 0 && lowStock.length === 0, detail: lowStock.length ? `${lowStock.length} alerta(s)` : products.length ? 'Completado' : 'Pendiente' },
    { title: 'Publicar catálogo', done: products.length > 0, detail: products.length ? 'Completado' : 'Pendiente' }
  ]
  const doneSteps = steps.filter((step) => step.done).length
  const quick = [
    { label: 'Catálogo', icon: ShoppingBag, tone: 'green', text: 'Productos y precios', count: products.length, unit: 'productos' },
    { label: 'Inventario', icon: Boxes, tone: 'blue', text: 'Existencias y alertas', count: lowStock.length, unit: 'alertas' },
    { label: 'Pedidos', icon: ClipboardList, tone: 'amber', text: 'Seguimiento de ventas', count: (records.pedidos || []).length, unit: 'pedidos' },
    { label: 'Clientes', icon: UsersRound, tone: 'sky', text: 'Directorio de contactos', count: (records.clientes || []).length, unit: 'clientes' },
    { label: 'Asesorías', icon: Sprout, tone: 'mint', text: 'Consultas y seguimiento', count: (records.asesorias || []).filter((r) => r.estado === 'Abierta').length, unit: 'abiertas' },
    { label: 'Contenido', icon: Sparkles, tone: 'violet', text: 'Páginas y promociones', count: (records.contenido || []).length, unit: 'contenidos' }
  ]
  return <>
    <SectionHeader icon={LayoutDashboard} tone="lime" eyebrow="CENTRO DIGITAL · 01" title={readOnly ? 'Bienvenido a VITIMAX' : 'Buenos días, administrador'} aside={<span className="heading-spark"><Sparkles size={20} /></span>} description={readOnly ? 'Consulta el catálogo y la información disponible.' : 'Aquí tienes una vista general de la operación de VITIMAX.'} actions={!readOnly && <><button className="btn secondary" onClick={onTemplate}><FileSpreadsheet size={16} /> Plantilla Excel</button><button className="btn primary" onClick={onImport}><Upload size={16} /> Importar productos</button></>} />
    <section className="hero-card hero-v2">
      <div className="hero-copy">
        <span className="pill light"><Zap size={13} /> {readOnly ? 'Modo invitado' : 'Centro de trabajo activo'}</span>
        <h2>Tu centro de VITIMAX,<br /><em>listo para crecer.</em></h2>
        <p>{readOnly ? 'Explora el catálogo, revisa el inventario y consulta al asistente con la información disponible.' : 'Gestiona catálogo, inventario, pedidos, clientes y asesorías desde un solo lugar, con ayuda del BOT IA.'}</p>
        <div className="hero-actions">
          {!readOnly && <button className="btn hero-btn" onClick={onImport}>Comenzar importación <ArrowUpToLine size={16} /></button>}
          <button className="btn hero-ghost" onClick={() => onNavigate('Catálogo')}>Ver catálogo <ArrowUpRight size={16} /></button>
        </div>
        <div className="hero-stats">
          <span><strong>{products.length}</strong> productos</span>
          <span><strong>{totalStock}</strong> unidades</span>
          <span><strong>{lowStock.length}</strong> alertas</span>
        </div>
      </div>
      <div className="hero-visual">
        <div className="hero-brand"><BrandMark size={104} /></div>
        <span className="hero-chip chip-a"><PackageCheck size={14} /> Catálogo real</span>
        <span className="hero-chip chip-b"><Bot size={14} /> BOT IA</span>
        <span className="hero-chip chip-c"><ShieldCheck size={14} /> Guardado seguro</span>
      </div>
    </section>
    <div className="section-title"><h2>Accesos rápidos</h2><span>Todo tu centro de trabajo</span></div>
    <section className="quick-grid">{quick.map(({ label, icon: Icon, tone, text, count, unit }) => <button className={`quick-card ${tone}`} key={label} onClick={() => onNavigate(label)}><span className={`quick-icon ${tone}`}><Icon size={21} /></span><div><strong>{label}</strong><span>{text}</span></div><em>{count} {unit}</em><ArrowUpRight className="quick-arrow" size={16} /></button>)}</section>
    <div className="section-title"><h2>Vista general</h2><span>Actualizado ahora</span></div>
    <section className="metrics"><Metric icon={ShoppingBag} label="Productos en catálogo" value={products.length} detail={products.length ? 'datos registrados' : 'pendiente de importar'} color="green" /><Metric icon={Boxes} label="Unidades en stock" value={totalStock} detail="según catálogo actual" color="blue" /><Metric icon={Bell} label="Alertas de inventario" value={lowStock.length} detail={lowStock.length ? 'requieren atención' : 'sin alertas'} color="orange" /><Metric icon={CircleHelp} label="Datos pendientes" value={missingData.length} detail="para completar ficha" color="purple" /></section>
    <div className="dashboard-grid"><section className="panel"><div className="panel-heading"><div><h3>Actividad del catálogo</h3><p>Productos cargados recientemente</p></div><button className="link-btn" onClick={() => onNavigate('Catálogo')}>Ver catálogo <ChevronRight size={15} /></button></div>{products.length ? products.slice(0, 5).map((p) => <div className="activity-row" key={p.id}><div className="product-thumb">{p.image ? <img src={p.image} /> : <Leaf size={19} />}</div><div><strong>{p.name || 'Sin nombre'}</strong><span>{p.category || 'Categoría pendiente'}</span></div><b>{money(p.price)}</b></div>) : <EmptyState icon={FileSpreadsheet} title="Aún no hay productos" text="Importa el Excel oficial de VITIMAX para comenzar." action="Importar Excel" onClick={onImport} />}</section><section className="panel checklist"><div className="panel-heading"><div><h3>Configuración inicial</h3><p>Deja todo listo para operar</p></div><span className="progress">{doneSteps} / {steps.length}</span></div>{steps.map((step, i) => <div className={step.done ? 'check-row done' : 'check-row'} key={step.title}><span>{step.done ? <Check size={13} /> : i + 1}</span><div><strong>{step.title}</strong><small>{step.detail}</small></div></div>)}</section></div>
  </>
}

function ApplicationReport({ products, lowStock, missingData, guest }) {
  const withImages = products.filter((product) => product.image).length
  const complete = products.length - missingData.length
  const cloudAssistant = import.meta.env.VITE_ENABLE_CLOUD_ASSISTANT === 'true'
  const completion = products.length ? Math.round((complete / products.length) * 100) : 0
  const totalStock = products.reduce((sum, product) => sum + (Number(product.stock) || 0), 0)
  const [pdfState, setPdfState] = useState('idle')
  const categoryRows = useMemo(() => {
    const map = new Map()
    products.forEach((product) => {
      const name = product.category || 'Sin categoría'
      const entry = map.get(name) || { name, count: 0, stock: 0 }
      entry.count += 1
      entry.stock += Number(product.stock) || 0
      map.set(name, entry)
    })
    return [...map.values()].sort((a, b) => b.count - a.count).slice(0, 6)
  }, [products])
  const maxCategory = Math.max(1, ...categoryRows.map((row) => row.count))
  const statusRows = [
    { label: 'Modo de acceso', value: guest ? 'Invitado · solo lectura' : 'Administrador autenticado', tone: guest ? 'amber' : 'green' },
    { label: 'Asistente IA', value: cloudAssistant && !guest ? 'Configurado' : 'Disponible en el dispositivo', tone: 'green' },
    { label: 'Importación Excel', value: 'Disponible', tone: 'green' },
    { label: 'Guardado de archivos', value: Capacitor.getPlatform() === 'android' ? 'Carpeta Descargas' : 'Descarga del navegador', tone: 'blue' }
  ]
  async function downloadReport() {
    setPdfState('loading')
    try {
      const pdf = new jsPDF()
      const width = pdf.internal.pageSize.getWidth()
      pdf.setFillColor(18, 61, 43)
      pdf.rect(0, 0, width, 34, 'F')
      pdf.setTextColor(216, 245, 95)
      pdf.setFontSize(18)
      pdf.text('VITIMAX', 18, 15)
      pdf.setTextColor(255, 255, 255)
      pdf.setFontSize(11)
      pdf.text('Informe general de la aplicacion', 18, 24)
      pdf.setTextColor(23, 60, 44)
      let y = 48
      const line = (label, value) => {
        if (y > 272) { pdf.addPage(); y = 24 }
        pdf.setFontSize(10)
        pdf.setTextColor(110, 132, 120)
        pdf.text(String(label), 18, y)
        pdf.setTextColor(23, 60, 44)
        pdf.text(String(value), 78, y)
        y += 8
      }
      pdf.setFontSize(13)
      pdf.text('Resumen operativo', 18, y)
      y += 10
      line('Productos analizados', products.length)
      line('Fichas completas', `${complete} de ${products.length} (${completion}%)`)
      line('Unidades en stock', totalStock)
      line('Alertas de stock', lowStock.length)
      line('Datos pendientes', missingData.length)
      line('Imagenes cargadas', withImages)
      line('Modo de acceso', guest ? 'Invitado (solo lectura)' : 'Administrador')
      line('Generado', new Date().toLocaleString('es-AR'))
      if (categoryRows.length) {
        y += 6
        pdf.setFontSize(13)
        pdf.text('Categorias del catalogo', 18, y)
        y += 10
        categoryRows.forEach((row) => line(row.name, `${row.count} producto(s) · ${row.stock} unidades`))
      }
      if (missingData.length) {
        y += 6
        pdf.setFontSize(13)
        pdf.text('Fichas con datos pendientes', 18, y)
        y += 10
        missingData.slice(0, 20).forEach((product) => line(product.name || 'Sin nombre', product.category || 'Categoria pendiente'))
      }
      pdf.setFontSize(9)
      pdf.setTextColor(130, 150, 138)
      pdf.text('Documento generado por VITIMAX', 18, 288)
      const saved = await saveBlob(pdf.output('blob'), 'informe-vitimax.pdf')
      setPdfState('ok')
      return saved
    } catch (error) {
      setPdfState('error')
      return null
    }
  }
  return <>
    <SectionHeader icon={FileBarChart2} tone="teal" eyebrow="CONTROL · 00" title="Informe de la aplicación" aside={<span className="heading-spark"><Sparkles size={20} /></span>} description="Resumen operativo, calidad del catálogo e integración del asistente IA." actions={<button className="btn primary" onClick={downloadReport} disabled={pdfState === 'loading'}><ArrowDownToLine size={16} /> {pdfState === 'loading' ? 'Generando…' : pdfState === 'ok' ? 'Informe descargado' : pdfState === 'error' ? 'Reintentar PDF' : 'Descargar PDF'}</button>} />
    <section className="metrics"><Metric icon={ShoppingBag} label="Productos analizados" value={products.length} detail="registros actuales" color="green" /><Metric icon={ClipboardCheck} label="Fichas completas" value={`${completion}%`} detail={`${complete} de ${products.length || 0}`} color="blue" /><Metric icon={Bell} label="Alertas activas" value={lowStock.length + missingData.length} detail="stock o datos pendientes" color="orange" /><Metric icon={FileSpreadsheet} label="Imágenes cargadas" value={withImages} detail="productos con fotografía" color="purple" /></section>
    <div className="dashboard-grid report-grid">
      <section className="panel completion-panel">
        <div className="panel-heading"><div><h3>Calidad del catálogo</h3><p>Qué porcentaje de fichas está completo.</p></div><span className={`table-status ${completion >= 80 ? 'green' : completion >= 50 ? 'amber' : 'red'}`}>{completion >= 80 ? 'Muy bien' : completion >= 50 ? 'Mejorable' : 'A revisar'}</span></div>
        <div className="completion-body">
          <div className="completion-ring" style={{ background: `conic-gradient(#1f7451 ${completion}%, #e7efe9 0)` }}><div><strong>{completion}%</strong><span>completo</span></div></div>
          <div className="completion-legend">
            <div><span className="legend-dot green" /><strong>{complete}</strong> fichas completas</div>
            <div><span className="legend-dot amber" /><strong>{missingData.length}</strong> con datos pendientes</div>
            <div><span className="legend-dot blue" /><strong>{withImages}</strong> con imagen cargada</div>
            <div><span className="legend-dot slate" /><strong>{totalStock}</strong> unidades en stock</div>
          </div>
        </div>
      </section>
      <section className="panel"><div className="panel-heading"><div><h3>Estado de la plataforma</h3><p>Indicadores configurados en esta instalación.</p></div></div><div className="report-list">{statusRows.map((row) => <div key={row.label}><strong>{row.label}</strong><span className={`table-status ${row.tone}`}>{row.value}</span></div>)}</div></section>
      <section className="panel"><div className="panel-heading"><div><h3>Categorías del catálogo</h3><p>Distribución de productos por categoría.</p></div><span className="progress">{categoryRows.length}</span></div>{categoryRows.length ? <div className="category-bars">{categoryRows.map((row) => <div className="category-bar" key={row.name}><div className="category-bar-head"><span>{row.name}</span><b>{row.count}</b></div><div className="bar-track"><span style={{ width: `${Math.round((row.count / maxCategory) * 100)}%` }} /></div></div>)}</div> : <div className="empty-state"><Tags size={24} /><h3>Sin categorías</h3><p>Importa el catálogo para ver la distribución.</p></div>}</section>
      <section className="panel"><div className="panel-heading"><div><h3>Recomendaciones</h3><p>Acciones sugeridas según los datos actuales.</p></div></div><div className="report-list"><div><strong>Completar fichas</strong><span>{missingData.length ? `Revisar ${missingData.length} producto(s) con campos pendientes.` : 'No hay fichas pendientes.'}</span></div><div><strong>Revisar inventario</strong><span>{lowStock.length ? `${lowStock.length} producto(s) tienen stock bajo.` : 'No hay alertas de stock bajo.'}</span></div><div><strong>Configurar IA</strong><span>Activa la clave de OpenRouter en Configuración para respuestas en cualquier red.</span></div></div></section>
    </div>
  </>
}

function Metric({ icon: Icon, label, value, detail, color = 'green' }) {
  return <div className={`metric metric-${color}`}><div className={`metric-icon ${color}`}><Icon size={19} /></div><div className="metric-copy"><span>{label}</span><strong>{value}</strong><small>{detail}</small></div></div>
}

function Catalog({ products, allProducts, categories, query, setQuery, category, setCategory, onImport, onExport, onTemplate, onCreate, onEdit, onDelete, readOnly }) {
  const totalStock = products.reduce((sum, product) => sum + (Number(product.stock) || 0), 0)
  return <>
    <SectionHeader icon={ShoppingBag} tone="green" eyebrow={readOnly ? 'CONSULTA · 02' : 'GESTIÓN · 02'} title="Catálogo de productos" description="La información visible proviene únicamente de los datos cargados." actions={!readOnly && <button className="btn primary" onClick={onCreate}><Plus size={17} /> Nuevo producto</button>} />
    <div className="toolbar catalog-toolbar">
      <div className="searchbox"><Search size={18} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar por nombre, uso, problema o característica..." /></div>
      <div className="selectbox"><Filter size={16} /><select value={category} onChange={(e) => setCategory(e.target.value)}>{categories.map((c) => <option key={c}>{c}</option>)}</select></div>
      {!readOnly && <div className="toolbar-actions"><button className="btn secondary compact" onClick={onTemplate}><FileSpreadsheet size={15} /> Plantilla</button><button className="btn secondary compact" onClick={onImport}><Upload size={15} /> Importar</button><button className="btn secondary compact" onClick={onExport} disabled={!allProducts.length}><ArrowDownToLine size={15} /> Exportar</button></div>}
    </div>
    <div className="catalog-summary">
      <span><ShoppingBag size={14} /> <b>{allProducts.length}</b> productos</span>
      <span><Boxes size={14} /> <b>{totalStock}</b> unidades visibles</span>
      <span><Tags size={14} /> <b>{Math.max(0, categories.length - 1)}</b> categorías</span>
    </div>
    {products.length ? <div className="product-grid">{products.map((p) => <ProductCard key={p.id} product={p} onEdit={onEdit} onDelete={onDelete} readOnly={readOnly} />)}</div> : <div className="empty-catalog"><div className="empty-icon"><FileSpreadsheet size={28} /></div><h2>{query || category !== 'Todas' ? 'No encontramos productos con ese filtro' : 'Tu catálogo está esperando los datos oficiales'}</h2><p>{readOnly ? 'Todavía no hay productos publicados para consultar.' : query || category !== 'Todas' ? 'Prueba con otra búsqueda o vuelve a ver todas las categorías.' : 'No se encontraron productos registrados. Importa un archivo Excel para cargar nombres, precios, stock y toda la información disponible sin duplicados.'}</p>{!readOnly && !query && category === 'Todas' && <div><button className="btn primary" onClick={onImport}><Upload size={16} /> Importar Excel</button><button className="btn secondary" onClick={onTemplate}>Descargar plantilla</button></div>}<small>{readOnly ? 'El catálogo será visible cuando el administrador cargue los productos.' : 'Los datos nunca se inventan: solo se muestra la información importada.'}</small></div>}
  </>
}

function ProductCard({ product, onEdit, onDelete, readOnly }) {
  const stock = Number(product.stock) || 0
  const stockPending = productPending(product, 'stock')
  const pricePending = productPending(product, 'price')
  const stockLabel = stockPending ? 'Stock pendiente' : stock <= 0 ? 'Sin stock' : stock <= 5 ? `${stock} · stock bajo` : `${stock} en stock`
  const stockClass = stockPending ? 'pending' : stock <= 0 ? 'out' : stock <= 5 ? 'low' : 'available'
  return <article className="product-card">
    <div className="card-image">
      {product.image ? <img src={product.image} alt="" /> : <><Leaf size={37} /><span>Foto pendiente</span></>}
      <span className={`stock-badge ${stockClass}`}>{stockLabel}</span>
      {!readOnly && <button className="card-menu" onClick={() => onEdit(product)} aria-label={`Editar ${product.name || 'producto'}`}><MoreHorizontal size={17} /></button>}
    </div>
    <div className="card-body">
      <div className="category-label"><Tags size={12} /> {product.category || 'Categoría pendiente'}</div>
      <h3>{product.name || 'Producto sin nombre'}</h3>
      {product.status === 'Pendiente de completar' && <span className="pending-badge">Pendiente de completar</span>}
      <p>{product.description || 'Descripción pendiente de completar.'}</p>
      <div className="card-footer">
        <strong className={pricePending ? 'price-pending' : undefined}>{pricePending ? 'Precio pendiente' : money(product.price)}</strong>
        {!readOnly && <div className="product-actions"><button className="link-btn" onClick={() => onEdit(product)}>Editar <ChevronRight size={14} /></button><button className="link-btn danger-link" onClick={() => onDelete(product.id)}><Trash2 size={14} /> Eliminar</button></div>}
      </div>
    </div>
  </article>
}

function Inventory({ products, lowStock, onImport, readOnly }) {
  const [search, setSearch] = useState('')
  const [view, setView] = useState('Todos')
  const totalStock = products.reduce((sum, product) => sum + (Number(product.stock) || 0), 0)
  const outOfStock = products.filter((product) => !productPending(product, 'stock') && Number(product.stock) <= 0)
  const maxStock = Math.max(1, ...products.map((product) => Number(product.stock) || 0))
  const views = ['Todos', 'Stock bajo', 'Sin stock', 'En rango']
  const filtered = products.filter((product) => {
    const stock = Number(product.stock) || 0
    const pending = productPending(product, 'stock')
    if (view === 'Stock bajo' && !(stock > 0 && stock <= 5)) return false
    if (view === 'Sin stock' && !(!pending && stock <= 0)) return false
    if (view === 'En rango' && (pending || stock <= 5)) return false
    if (!search.trim()) return true
    return `${product.name || ''} ${product.category || ''}`.toLowerCase().includes(search.trim().toLowerCase())
  })
  return <>
    <SectionHeader icon={Boxes} tone="blue" eyebrow="OPERACIÓN · 03" title="Inventario" description="Consulta existencias y detecta rápidamente lo que necesita atención." actions={!readOnly && <button className="btn primary" onClick={onImport}><Upload size={16} /> Actualizar por Excel</button>} />
    <section className="metrics inventory-metrics">
      <Metric icon={Boxes} label="Total de unidades" value={totalStock} detail="stock registrado" color="green" />
      <Metric icon={Bell} label="Stock bajo" value={lowStock.length} detail="entre 1 y 5 unidades" color="orange" />
      <Metric icon={AlertTriangle} label="Sin stock" value={outOfStock.length} detail="productos agotados" color="purple" />
      <Metric icon={PackageCheck} label="Salud del stock" value={products.length ? `${Math.round(((products.length - lowStock.length - outOfStock.length) / products.length) * 100)}%` : '0%'} detail="fichas en rango" color="blue" />
    </section>
    <section className="panel table-panel">
      <div className="panel-heading"><div><h3>Existencias actuales</h3><p>Los valores se actualizan desde el catálogo o Excel.</p></div></div>
      {products.length ? <>
        <div className="module-toolbar inventory-toolbar">
          <div className="searchbox"><Search size={17} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar producto o categoría..." /></div>
          <div className="chip-row">{views.map((item) => <button key={item} className={view === item ? 'chip active' : 'chip'} onClick={() => setView(item)}>{item}</button>)}</div>
        </div>
        <div className="table-wrap"><table><thead><tr><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock actual</th><th>Estado</th></tr></thead><tbody>{filtered.map((p) => {
          const stock = Number(p.stock) || 0
          const pending = productPending(p, 'stock')
          const tone = pending ? 'amber' : stock <= 0 ? 'red' : stock <= 5 ? 'amber' : 'green'
          const label = pending ? 'Revisar' : stock <= 0 ? 'Sin stock' : stock <= 5 ? 'Stock bajo' : 'En rango'
          return <tr key={p.id}>
            <td><strong>{p.name || 'Sin nombre'}</strong></td>
            <td>{p.category || 'Pendiente'}</td>
            <td>{productPending(p, 'price') ? 'Pendiente' : money(p.price)}</td>
            <td><div className="stock-cell"><span>{pending ? 'Pendiente' : stock}</span><div className="bar-track"><span className={`bar-${tone}`} style={{ width: `${pending ? 0 : Math.max(4, Math.round((stock / maxStock) * 100))}%` }} /></div></div></td>
            <td><span className={`table-status ${tone}`}>{label}</span></td>
          </tr>
        })}</tbody></table></div>
        {!filtered.length && <div className="empty-state"><Search size={24} /><h3>Sin resultados</h3><p>Prueba con otra búsqueda o cambia el filtro de estado.</p></div>}
      </> : <EmptyState icon={Boxes} title="Sin inventario registrado" text={readOnly ? 'El inventario aparecerá cuando el administrador cargue los datos.' : 'Carga el Excel oficial para visualizar existencias.'} action={readOnly ? undefined : 'Importar Excel'} onClick={readOnly ? undefined : onImport} />}
    </section>
  </>
}

function ModulePage({ config, rows, onCreate, onEdit, onDelete, readOnly, onAssistant }) {
  const [search, setSearch] = useState('')
  const [status, setStatus] = useState('Todos')
  const statusField = config.statusKey ? config.fields.find((field) => field.key === config.statusKey) : null
  const statuses = statusField ? ['Todos', ...statusField.options] : []
  const filtered = rows.filter((row) => {
    if (status !== 'Todos' && row[config.statusKey] !== status) return false
    if (!search.trim()) return true
    return Object.values(row).join(' ').toLowerCase().includes(search.trim().toLowerCase())
  })
  return <>
    <SectionHeader icon={config.icon} tone={config.tone} eyebrow={config.eyebrow} title={config.title} description={config.description} actions={<>{onAssistant && <button className="btn secondary" onClick={onAssistant}><Bot size={16} /> Abrir BOT IA</button>}{!readOnly && <button className="btn primary" onClick={onCreate}><Plus size={16} /> {config.addLabel}</button>}</>} />
    <section className="metrics">{config.stats(rows).map(({ icon, label, value, detail, color }) => <Metric key={label} icon={icon} label={label} value={value} detail={detail} color={color} />)}</section>
    {rows.length ? <section className="panel module-panel">
      <div className="panel-heading"><div><h3>{config.listTitle}</h3><p>Busca, filtra y actualiza cada registro cuando lo necesites.</p></div><span className="progress">{filtered.length} / {rows.length}</span></div>
      <div className="module-toolbar">
        <div className="searchbox"><Search size={17} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder={`Buscar ${config.title.toLowerCase()}...`} /></div>
        {statuses.length > 1 && <div className="chip-row">{statuses.map((item) => <button key={item} className={status === item ? 'chip active' : 'chip'} onClick={() => setStatus(item)}>{item}</button>)}</div>}
      </div>
      <div className="record-list">{filtered.length ? filtered.map((record) => <RecordRow key={record.id} config={config} record={record} onEdit={onEdit} onDelete={onDelete} readOnly={readOnly} />) : <div className="empty-state"><Search size={24} /><h3>Sin resultados</h3><p>Prueba con otra búsqueda o cambia el filtro de estado.</p></div>}</div>
    </section> : <ModuleEmpty config={config} readOnly={readOnly} onCreate={onCreate} onAssistant={onAssistant} />}
  </>
}

function RecordRow({ config, record, onEdit, onDelete, readOnly }) {
  const Icon = config.icon
  const tone = config.statusKey ? statusTone(record[config.statusKey]) : config.tone
  return <div className="record-row">
    <span className={`record-icon ${config.tone}`}><Icon size={18} /></span>
    <div className="record-main">
      <div className="record-title"><strong>{config.primary(record)}</strong>{config.statusKey && record[config.statusKey] && <span className={`table-status ${tone}`}>{record[config.statusKey]}</span>}</div>
      <small>{config.secondary(record)}</small>
    </div>
    <div className="record-meta">{config.meta(record)}</div>
    {!readOnly && <div className="record-actions"><button className="icon-btn small" onClick={() => onEdit(record)} aria-label="Editar registro"><PenLine size={15} /></button><button className="icon-btn small danger" onClick={() => onDelete(record.id)} aria-label="Eliminar registro"><Trash2 size={15} /></button></div>}
  </div>
}

function ModuleEmpty({ config, readOnly, onCreate, onAssistant }) {
  const Icon = config.icon
  return <div className="module-empty module-empty-v2">
    <div className={`module-orb ${config.tone}`}><Icon size={34} /></div>
    <p className="eyebrow">MÓDULO · {config.title.toUpperCase()}</p>
    <h1>{config.emptyTitle}</h1>
    <p>{config.emptyText}</p>
    <div className="module-actions">{!readOnly && <button className="btn primary" onClick={onCreate}><Plus size={16} /> {config.addLabel}</button>}{onAssistant && <button className="btn secondary" onClick={onAssistant}><Bot size={16} /> Abrir BOT IA</button>}</div>
    <div className="feature-grid">{config.features.map(([FeatureIcon, title, text]) => <div className="feature-card" key={title}><span className={`feature-icon ${config.tone}`}><FeatureIcon size={18} /></span><div><strong>{title}</strong><small>{text}</small></div></div>)}</div>
  </div>
}

function ModuleFormModal({ config, record, onClose, onSave }) {
  const [form, setForm] = useState(() => {
    const base = {}
    config.fields.forEach((field) => { base[field.key] = record?.[field.key] ?? field.default ?? '' })
    return base
  })
  const [error, setError] = useState('')
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))
  function submit(event) {
    event.preventDefault()
    const missing = config.fields.filter((field) => field.required && !String(form[field.key] ?? '').trim())
    if (missing.length) return setError(`Completa los campos obligatorios: ${missing.map((field) => field.label).join(', ')}.`)
    setError('')
    onSave(form)
  }
  return <ModalShell title={record ? `Editar ${config.singular}` : config.addLabel} eyebrow={config.title.toUpperCase()} onClose={onClose}>
    <form className="module-form" onSubmit={submit}>
      <div className="form-grid">
        {config.fields.map((field) => <label key={field.key} className={field.full ? 'full' : ''}>
          <span className="field-label">{field.label}{field.required && <i>*</i>}</span>
          {field.type === 'textarea'
            ? <textarea rows={3} value={form[field.key]} onChange={(e) => update(field.key, e.target.value)} placeholder={field.placeholder} />
            : field.type === 'select'
              ? <select value={form[field.key]} onChange={(e) => update(field.key, e.target.value)}>{field.options.map((option) => <option key={option}>{option}</option>)}</select>
              : <input type={field.type || 'text'} min={field.min} value={form[field.key]} onChange={(e) => update(field.key, e.target.value)} placeholder={field.placeholder} />}
        </label>)}
      </div>
      {error && <p className="auth-error">{error}</p>}
      <div className="modal-actions"><button type="button" className="btn secondary" onClick={onClose}>Cancelar</button><button type="submit" className="btn primary">{record ? 'Guardar cambios' : `Crear ${config.singular}`}</button></div>
    </form>
  </ModalShell>
}

function EmptyState({ icon: Icon, title, text, action, onClick }) { return <div className="empty-state"><Icon size={28} /><h3>{title}</h3><p>{text}</p>{action && <button className="link-btn" onClick={onClick}>{action} <ChevronRight size={15} /></button>}</div> }

function ModalShell({ title, eyebrow, children, onClose, actions, className = '' }) {
  return <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}><div className={`modal${className ? ` ${className}` : ''}`}><div className="modal-head"><div><p className="eyebrow">{eyebrow}</p><h2>{title}</h2></div><button className="icon-btn" aria-label="Cerrar" onClick={onClose}><X size={19} /></button></div>{children}{actions && <div className="modal-actions">{actions}</div>}</div></div>
}

function profileInitials(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean)
  if (!parts.length) return 'V'
  return `${parts[0][0] || ''}${parts[1]?.[0] || ''}`.toUpperCase() || 'V'
}

function Avatar({ name, photo, className = '', alt = '' }) {
  const src = String(photo || '').trim()
  const [broken, setBroken] = useState(false)
  useEffect(() => { setBroken(false) }, [src])
  const classes = className ? `avatar ${className}` : 'avatar'
  if (src && !broken) {
    return <div className={`${classes} has-photo`}><img src={src} alt={alt} onError={() => setBroken(true)} /></div>
  }
  return <div className={classes}>{profileInitials(name)}</div>
}

function ProfileModal({ profile, user, guest = false, onClose, onSave, onSignOut, onResendVerification, onPasswordReset }) {
  const [form, setForm] = useState(profile)
  const [photoFile, setPhotoFile] = useState(null)
  const [photoPreview, setPhotoPreview] = useState(String(profile.photo || '').trim())
  const [photoBroken, setPhotoBroken] = useState(false)
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const fileInputRef = useRef(null)
  const previewUrlRef = useRef('')
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))

  useEffect(() => () => { if (previewUrlRef.current) URL.revokeObjectURL(previewUrlRef.current) }, [])

  function setPreview(value) {
    if (previewUrlRef.current) {
      URL.revokeObjectURL(previewUrlRef.current)
      previewUrlRef.current = ''
    }
    setPhotoBroken(false)
    setPhotoPreview(value)
  }

  function choosePhoto(event) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    if (!file.type.startsWith('image/')) {
      setError('Selecciona una imagen JPG, PNG o WEBP.')
      return
    }
    if (file.size > 8 * 1024 * 1024) {
      setError('La foto supera el límite de 8 MB.')
      return
    }
    setError('')
    setPhotoFile(file)
    // Importante: revocar la vista previa anterior antes de crear la nueva,
    // si no se revoca la URL recién creada y la foto no se muestra.
    const objectUrl = URL.createObjectURL(file)
    setPreview(objectUrl)
    previewUrlRef.current = objectUrl
  }

  function removePhoto() {
    setPhotoFile(null)
    setPreview('')
    update('photo', '')
  }

  async function save() {
    if (saving) return
    setSaving(true)
    setError('')
    try {
      await onSave(form, photoFile)
    } catch (saveError) {
      setError(saveError?.message || 'No se pudo guardar el perfil. Inténtalo de nuevo.')
    } finally {
      setSaving(false)
    }
  }

  function signOut() {
    if (window.confirm('¿Cerrar sesión en VITIMAX? Podrás volver a entrar con tu cuenta cuando quieras.')) onSignOut()
  }

  const showPhoto = Boolean(photoPreview) && !photoBroken
  const emailVerified = Boolean(user?.emailVerified)
  const passwordAccount = hasPasswordProvider(user)
  return <ModalShell className="profile-modal" title="Mi perfil" eyebrow="CUENTA" onClose={onClose} actions={<>
    <button className="btn secondary profile-signout" onClick={signOut}><LogOut size={16} /> Cerrar sesión</button>
    <button className="btn secondary" onClick={onClose} disabled={saving}>Cancelar</button>
    <button className="btn primary" onClick={save} disabled={saving}>{saving ? 'Guardando…' : 'Guardar cambios'}</button>
  </>}>
    <div className="profile-body">
      <div className="profile-hero">
        <div className="profile-hero-cover" aria-hidden="true"><span /></div>
        <div className="profile-hero-body">
          <button type="button" className="profile-photo-button" onClick={() => fileInputRef.current?.click()} title="Cambiar foto de perfil" aria-label="Cambiar foto de perfil">
            {showPhoto
              ? <img className="profile-photo profile-photo-lg" src={photoPreview} alt="" onError={() => setPhotoBroken(true)} />
              : <span className="avatar profile-avatar profile-avatar-lg">{profileInitials(form.name)}</span>}
            <span className="profile-photo-badge"><Camera size={15} /></span>
          </button>
          <div className="profile-hero-info">
            <strong>{form.name || 'Administrador'}</strong>
            <span>{form.email || 'Correo pendiente'}</span>
            <em><BadgeCheck size={13} /> {guest ? 'Acceso de invitado' : `${providerLabel(user)} · Cuenta VITIMAX`}</em>
          </div>
        </div>
        <div className="profile-photo-actions">
          <button type="button" className="btn secondary" onClick={() => fileInputRef.current?.click()}><ImagePlus size={16} /> Elegir foto</button>
          <label className="btn secondary"><Camera size={16} /> Tomar foto<input type="file" accept="image/*" capture="user" onChange={choosePhoto} hidden /></label>
          {(photoPreview || photoFile) && <button type="button" className="btn secondary profile-remove-photo" onClick={removePhoto}><Trash2 size={16} /> Quitar</button>}
        </div>
        {!photoPreview && <p className="profile-photo-hint">Agrega una foto para personalizar tu perfil. Se guarda en tu cuenta y aparece en toda la aplicación.</p>}
        {photoBroken && <p className="profile-photo-note">La foto guardada no se pudo cargar. Elige una nueva o quítala.</p>}
        <input ref={fileInputRef} type="file" accept="image/*" hidden onChange={choosePhoto} />
      </div>
      {error && <p className="auth-error profile-error">{error}</p>}
      <p className="profile-section-title"><UserRound size={14} /> Datos personales</p>
      <div className="form-grid profile-fields">
        <label>Nombre completo<input value={form.name} onChange={(e) => update('name', e.target.value)} placeholder="Tu nombre" autoComplete="name" /></label>
        <label>Teléfono<input value={form.phone || ''} onChange={(e) => update('phone', e.target.value)} placeholder="Opcional" inputMode="tel" autoComplete="tel" /></label>
      </div>
      <p className="profile-section-title profile-section-title-spaced"><ShieldCheck size={14} /> Cuenta y seguridad</p>
      <div className="profile-account-card">
        <span className="profile-account-icon">{providerLabel(user) === 'Google' ? <GoogleBrandIcon /> : providerLabel(user) === 'Microsoft' ? <MicrosoftBrandIcon /> : <Mail size={17} />}</span>
        <div>
          <small>{guest ? 'Modo invitado' : `Acceso con ${providerLabel(user)}`}</small>
          <strong>{form.email || 'Sin correo registrado'}</strong>
        </div>
        {!guest && <em className={emailVerified ? 'verified' : 'pending'}>{emailVerified ? 'Verificado' : 'Sin verificar'}</em>}
      </div>
      {!guest && <div className="profile-account-actions">
        {!emailVerified && <button type="button" className="btn secondary" onClick={onResendVerification}><Mail size={15} /> Verificar mi correo</button>}
        {passwordAccount && <button type="button" className="btn secondary" onClick={onPasswordReset}><ShieldCheck size={15} /> Cambiar contraseña</button>}
        <button type="button" className="btn secondary profile-signout-inline" onClick={signOut}><LogOut size={15} /> Cerrar sesión</button>
      </div>}
      <details className="profile-url-details">
        <summary>Usar enlace de imagen</summary>
        <label>URL de la foto<input value={form.photo || ''} onChange={(e) => { const value = e.target.value; update('photo', value); setPhotoFile(null); setPreview(value.trim()) }} placeholder="https://..." inputMode="url" /></label>
        <small>También puedes pegar el enlace público de una imagen. Si el enlace no funciona, se mostrarán tus iniciales.</small>
      </details>
    </div>
  </ModalShell>
}

function SettingsModal({ settings, onClose, onSave }) {
  const [form, setForm] = useState(settings)
  const [test, setTest] = useState({ state: 'idle', message: '' })
  const [saveTest, setSaveTest] = useState({ state: 'idle', message: '' })
  const [diag, setDiag] = useState(null)
  const [perm, setPerm] = useState({ state: 'idle', message: '' })
  useEffect(() => {
    let alive = true
    async function run() {
      const platform = Capacitor.getPlatform()
      const available = Capacitor.isPluginAvailable('SaveToDownloads')
      let native = null
      if (available) {
        try { native = await SaveToDownloads.diagnose() } catch (error) { native = { error: String(error?.message || error) } }
      }
      if (alive) setDiag({ platform, available, native })
    }
    run()
    return () => { alive = false }
  }, [])
  const update = (key, value) => setForm((current) => ({ ...current, [key]: value }))
  async function testSave() {
    setSaveTest({ state: 'loading', message: 'Guardando archivo de prueba…' })
    try {
      const saved = await saveBlob(new Blob([`Prueba de guardado VITIMAX\nFecha: ${new Date().toLocaleString('es-AR')}\nSi ves este archivo, la descarga funciona de verdad.\n`], { type: 'text/plain;charset=utf-8' }), 'prueba-vitimax.txt')
      setSaveTest({ state: 'ok', message: `Archivo guardado en: ${saved.location || saved.filename}` })
    } catch (error) {
      setSaveTest({ state: 'error', message: error?.message || 'No se pudo guardar el archivo de prueba. Revisa el permiso de almacenamiento.' })
    }
  }
  async function requestPermissions() {
    setPerm({ state: 'loading', message: 'Solicitando permisos…' })
    const status = await requestDevicePermissions()
    setPerm({ state: status?.allGranted ? 'ok' : 'error', message: permissionsMessage(status) })
  }
  async function testNotification() {
    setPerm({ state: 'loading', message: 'Enviando notificación de prueba…' })
    const ok = await sendDeviceNotification('VITIMAX', 'Las notificaciones de VITIMAX funcionan correctamente.')
    setPerm({ state: ok ? 'ok' : 'error', message: ok ? 'Notificación enviada. Revisa la barra de notificaciones del teléfono.' : 'No se pudo enviar la notificación. Activa el permiso de notificaciones.' })
  }
  async function testConnection() {
    const url = resolveAiUrl(form)
    const key = String(form.aiApiKey || import.meta.env.VITE_AI_DIRECT_KEY || '').trim()
    if (!url && !key) return setTest({ state: 'error', message: 'Configura la URL del servidor o una clave de OpenRouter.' })
    setTest({ state: 'loading', message: 'Probando conexión…' })
    const errors = []
    if (url) {
      try {
        const response = await fetchWithTimeout(`${url}/health`, {}, 15_000)
        const data = await response.json().catch(() => ({}))
        if (!response.ok) throw new Error(data.error || `El servidor respondió ${response.status}.`)
        setTest({ state: 'ok', message: `Conectado · ${data.provider || 'IA'} · ${data.model || ''}${data.keyConfigured === false ? ' · falta la clave en el servidor' : ''}` })
        return
      } catch (error) {
        errors.push(error)
      }
    }
    if (key) {
      try {
        const response = await fetchWithTimeout('https://openrouter.ai/api/v1/chat/completions', {
          method: 'POST',
          headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json', 'X-Title': 'VITIMAX' },
          body: JSON.stringify({ models: DIRECT_AI_MODELS, messages: [{ role: 'user', content: 'Responde solo: OK' }], max_tokens: 20 })
        }, 45_000)
        const data = await response.json().catch(() => ({}))
        if (!response.ok) throw new Error(data?.error?.message || `OpenRouter respondió ${response.status}.`)
        setTest({ state: 'ok', message: `Conectado · OpenRouter · ${data.model || DIRECT_AI_MODELS[0]}` })
        return
      } catch (error) {
        errors.push(error)
      }
    }
    setTest({ state: 'error', message: aiErrorMessage(errors[errors.length - 1], 'No se pudo conectar con la IA.') })
  }
  const android = Capacitor.getPlatform() === 'android'
  return <ModalShell title="Configuración" eyebrow="PREFERENCIAS" onClose={onClose} className="settings-modal" actions={<><button className="btn secondary" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={() => onSave(form)}><Check size={16} /> Guardar configuración</button></>}>
    <div className="settings-groups">
      <section className="settings-group">
        <div className="settings-group-head"><span className="settings-group-icon green"><SlidersHorizontal size={17} /></span><div><h3>General</h3><p>Preferencias básicas de la aplicación.</p></div></div>
        <label className="setting-row"><span><strong>Notificaciones</strong><small>Mostrar alertas de inventario y actividad</small></span><span className="switch"><input type="checkbox" checked={Boolean(form.notifications)} onChange={(e) => { update('notifications', e.target.checked); if (e.target.checked && android) requestPermissions() }} /><i /></span></label>
        <div className="form-grid">
          <label>Umbral de stock bajo<input type="number" min="0" value={form.lowStock} onChange={(e) => update('lowStock', e.target.value)} /><small>Se avisa cuando el stock llega a este número.</small></label>
          <label>Moneda del catálogo<select value={form.currency} onChange={(e) => update('currency', e.target.value)}><option>ARS</option><option>CLP</option><option>USD</option></select><small>Formato usado en precios y totales.</small></label>
        </div>
      </section>
      <section className="settings-group">
        <div className="settings-group-head"><span className="settings-group-icon violet"><Bot size={17} /></span><div><h3>Asistente IA</h3><p>Conecta el BOT para usarlo desde cualquier red.</p></div></div>
        <label className="full">Servidor de IA (URL)<input value={form.aiServerUrl ?? ''} onChange={(e) => update('aiServerUrl', e.target.value)} placeholder={ENV_AI_URL || 'https://tu-servidor-de-ia.com'} /><small>En Android usa la IP de tu PC en la misma red WiFi (ej. http://192.168.0.10:8787) o una URL HTTPS pública. Si dejas 127.0.0.1/localhost, la app usará el modo directo con OpenRouter y no fallará por conexión local.</small></label>
        <label className="full">Token del servidor (opcional)<input value={form.aiToken ?? ''} onChange={(e) => update('aiToken', e.target.value)} placeholder="Solo si tu servidor lo pide" /></label>
        <label className="full">Clave de OpenRouter (modo directo)<input type="password" value={form.aiApiKey ?? ''} onChange={(e) => update('aiApiKey', e.target.value)} placeholder="sk-or-v1-..." /><small>Opcional. Permite usar la IA en este dispositivo sin servidor. Se guarda solo en este equipo.</small></label>
        <div className="settings-actions">
          <button type="button" className="btn secondary" onClick={testConnection} disabled={test.state === 'loading'}>{test.state === 'loading' ? <RotateCcw size={15} className="spin" /> : <Zap size={15} />} {test.state === 'loading' ? 'Probando…' : 'Probar conexión IA'}</button>
        </div>
        {test.message && <small className={test.state === 'error' ? 'auth-error' : 'field-help'}>{test.message}</small>}
      </section>
      {android && <section className="settings-group">
        <div className="settings-group-head"><span className="settings-group-icon blue"><ShieldCheck size={17} /></span><div><h3>Permisos y dispositivo</h3><p>Notificaciones, cámara y archivos del teléfono.</p></div></div>
        <div className="settings-actions">
          <button type="button" className="btn secondary" onClick={requestPermissions} disabled={perm.state === 'loading'}><ShieldCheck size={15} /> {perm.state === 'loading' ? 'Solicitando…' : 'Activar todos los permisos'}</button>
          <button type="button" className="btn secondary" onClick={testNotification} disabled={perm.state === 'loading'}><Bell size={15} /> Probar notificación</button>
          <button type="button" className="btn secondary" onClick={openDeviceSettings}><Settings2 size={15} /> Ajustes de la app</button>
        </div>
        {perm.message && <small className={perm.state === 'error' ? 'auth-error' : 'field-help'}>{perm.message}</small>}
        <small className="field-help">VITIMAX usa notificaciones, cámara (código QR), archivos y fotos. En Android 10 o superior los archivos se guardan en Descargas sin permiso especial.</small>
      </section>}
      <section className="settings-group">
        <div className="settings-group-head"><span className="settings-group-icon amber"><Download size={17} /></span><div><h3>Archivos y descargas</h3><p>Comprueba que guardar archivos funciona de verdad.</p></div></div>
        <div className="settings-actions">
          <button type="button" className="btn secondary" onClick={testSave} disabled={saveTest.state === 'loading'}><Download size={15} /> {saveTest.state === 'loading' ? 'Guardando…' : 'Probar guardado en Descargas'}</button>
        </div>
        {saveTest.message && <small className={saveTest.state === 'error' ? 'auth-error' : 'field-help'}>{saveTest.message}</small>}
        <small className="field-help">En Android 10 o superior no se necesita permiso. En Android 9 o anterior se pedirá permiso de almacenamiento una sola vez.</small>
        {diag && <small className="field-help">Diagnóstico: plataforma {diag.platform} · plugin {diag.available ? 'disponible' : 'no disponible'}{diag.native && !diag.native.error ? ` · Android ${diag.native.android || '?'} (SDK ${diag.native.sdk ?? '?'}) · MediaStore ${diag.native.mediaStore ? 'sí' : 'no'} · permiso ${diag.native.permission || '?'}` : diag.native?.error ? ` · error: ${diag.native.error}` : ''}</small>}
      </section>
      <section className="settings-group about-group">
        <div className="settings-group-head"><span className="settings-group-icon slate"><Info size={17} /></span><div><h3>Acerca de VITIMAX</h3><p>Centro de gestión agrícola digital.</p></div></div>
        <div className="about-grid"><div><span>Versión</span><strong>1.1.7</strong></div><div><span>Plataforma</span><strong>{Capacitor.getPlatform() === 'web' ? 'Web / PWA' : Capacitor.getPlatform() === 'android' ? 'Android' : 'iOS'}</strong></div><div><span>Guardado</span><strong>Local + nube</strong></div></div>
      </section>
    </div>
  </ModalShell>
}

function NotificationsModal({ lowStock, missingData, onClose }) {
  return <ModalShell title="Notificaciones" eyebrow="CENTRO DE ALERTAS" onClose={onClose}><div className="notification-list">{lowStock.length ? lowStock.map((p) => <div className="notification-row" key={p.id}><Bell size={16} /><span><strong>Stock bajo</strong><small>{p.name || 'Producto pendiente'} · {p.stock} unidades</small></span></div>) : <div className="empty-state"><Bell size={25} /><h3>Sin alertas de stock</h3><p>No hay productos dentro del umbral actual.</p></div>}{missingData.length > 0 && <div className="notification-row"><CircleHelp size={16} /><span><strong>Datos pendientes</strong><small>{missingData.length} producto(s) necesitan información</small></span></div>}</div></ModalShell>
}

function ImportReport({ report, onClose }) {
  const sheets = report.sheets || []
  const irregular = report.irregular || []
  return <ModalShell title="Resultado de importación" eyebrow="EXCEL VALIDADO" onClose={onClose}>
    <div className="import-summary">
      <Metric icon={FileSpreadsheet} label="Filas leídas" value={report.rows} detail={`${sheets.length} hoja(s)`} color="blue" />
      <Metric icon={Plus} label="Productos" value={report.added} detail={`${report.updated} actualizados`} color="green" />
      <Metric icon={PackageCheck} label="Registros" value={report.recordsAdded || 0} detail={`${report.recordsUpdated || 0} actualizados`} color="purple" />
      <Metric icon={CircleHelp} label="Pendientes" value={report.missing + report.invalid} detail={`${report.invalid} sin nombre`} color="orange" />
    </div>
    {sheets.length > 0 && <div className="import-sheets">
      <h4>Hojas procesadas</h4>
      {sheets.map((sheet) => <div className="import-sheet-row" key={sheet.name}>
        <span><FileSpreadsheet size={15} /> {sheet.name}</span>
        <small>{sheet.kind} · encabezado en fila {sheet.headerRow} · {sheet.imported} registro(s)</small>
      </div>)}
    </div>}
    {irregular.length > 0 && <p className="modal-note"><AlertTriangle size={15} /> {irregular.length} hoja(s) sin tabla reconocible: {irregular.map((item) => item.name).join(', ')}.</p>}
    {report.aiUsed && <p className="modal-note"><Bot size={15} /> La estructura era irregular y el BOT IA mapeó las filas al esquema VITIMAX.</p>}
    <p className="modal-note"><CircleHelp size={15} /> Los productos con precio o stock faltante se guardaron con 0 y quedaron como "Pendiente de completar". No se inventaron datos.</p>
    <div className="modal-actions"><button className="btn primary" onClick={onClose}>Continuar</button></div>
  </ModalShell>
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
  const [status, setStatus] = useState('starting')
  const [facing, setFacing] = useState('environment')
  const [manual, setManual] = useState('')
  const [torchOn, setTorchOn] = useState(false)
  const [torchAvailable, setTorchAvailable] = useState(false)

  useEffect(() => { resultRef.current = onResult }, [onResult])

  function cameraError(scannerError) {
    const name = scannerError?.name || ''
    if (name === 'NotAllowedError' || /permission|not allowed/i.test(scannerError?.message || '')) return 'Permiso de cámara rechazado. Actívalo en los ajustes del teléfono o escribe el código manualmente.'
    if (name === 'NotFoundError') return 'No se encontró una cámara disponible en este dispositivo.'
    if (name === 'NotReadableError') return 'La cámara está siendo usada por otra aplicación. Ciérrala e inténtalo otra vez.'
    if (window.location.protocol !== 'https:' && !['localhost', '127.0.0.1'].includes(window.location.hostname)) return 'El lector necesita una conexión HTTPS para usar la cámara.'
    return 'No se pudo abrir la cámara. Revisa el permiso o escribe el código manualmente.'
  }

  async function stopScanner() {
    const scanner = scannerRef.current
    scannerRef.current = null
    if (!scanner) return
    try { if (scanner.isScanning) await scanner.stop() } catch { /* la cámara ya estaba detenida */ }
    try { scanner.clear() } catch { /* el contenedor ya estaba limpio */ }
  }

  async function startCamera(mode) {
    await stopScanner()
    setStatus('starting')
    setError('')
    setTorchOn(false)
    setTorchAvailable(false)
    const scanner = new Html5Qrcode('vitimax-qr-reader')
    scannerRef.current = scanner
    try {
      await scanner.start(
        { facingMode: mode },
        { fps: 10, qrbox: { width: 230, height: 230 }, aspectRatio: 1 },
        (decodedText) => {
          if (!activeRef.current) return
          activeRef.current = false
          scanner.stop().catch(() => undefined)
          resultRef.current(decodedText)
        },
        () => undefined
      )
      setStatus('ready')
      try {
        const capabilities = scanner.getRunningTrackCapabilities?.()
        setTorchAvailable(Boolean(capabilities?.torch))
      } catch { setTorchAvailable(false) }
    } catch (scannerError) {
      setStatus('error')
      setError(cameraError(scannerError))
    }
  }

  useEffect(() => {
    activeRef.current = true
    startCamera('environment')
    return () => {
      activeRef.current = false
      stopScanner()
    }
    // El escáner se inicia una sola vez; girar la cámara usa startCamera directamente.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  async function flipCamera() {
    const next = facing === 'environment' ? 'user' : 'environment'
    setFacing(next)
    activeRef.current = true
    await startCamera(next)
  }

  async function toggleTorch() {
    const scanner = scannerRef.current
    if (!scanner?.applyVideoConstraints) return
    try {
      await scanner.applyVideoConstraints({ advanced: [{ torch: !torchOn }] })
      setTorchOn((value) => !value)
    } catch {
      setError('Este teléfono no permite activar la linterna desde la app.')
    }
  }

  async function scanImage(event) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setError('')
    activeRef.current = false
    if (scannerRef.current?.isScanning) await scannerRef.current.stop().catch(() => undefined)
    let scanner = scannerRef.current
    if (!scanner) {
      scanner = new Html5Qrcode('vitimax-qr-reader')
      scannerRef.current = scanner
    }
    try {
      const decodedText = await scanner.scanFile(file, true)
      resultRef.current(decodedText)
    } catch {
      setError('No encontré un código QR legible en esa imagen. Prueba con una foto más clara o escribe el código manualmente.')
      setStatus('error')
    }
  }

  function submitManual(event) {
    event.preventDefault()
    const value = manual.trim()
    if (!value) return
    resultRef.current(value)
  }

  return <div className="modal-backdrop">
    <div className="modal qr-modal">
      <div className="modal-head">
        <div><p className="eyebrow">LECTOR DE ENVASES</p><h2>Escanea el código QR</h2></div>
        <button className="icon-btn" aria-label="Cerrar escáner" onClick={onClose}><X size={19} /></button>
      </div>
      <p className="modal-note"><QrCode size={15} /> Enfoca el código dentro del recuadro. Debe identificar un producto registrado en VITIMAX.</p>
      <div className="qr-stage">
        <div id="vitimax-qr-reader" className="qr-reader" />
        {status === 'ready' && <div className="qr-frame" aria-hidden="true"><span /><span /><span /><span /><i /></div>}
        {status === 'starting' && <div className="qr-loading"><span className="qr-spinner" /> Abriendo cámara…</div>}
      </div>
      <div className="qr-status-row">
        <span className={`qr-chip ${status}`}>{status === 'ready' ? 'Cámara lista' : status === 'error' ? 'Cámara no disponible' : 'Iniciando…'}</span>
        <button type="button" className="qr-tool" onClick={flipCamera} title="Cambiar cámara"><SwitchCamera size={16} /> Girar</button>
        {torchAvailable && <button type="button" className={`qr-tool${torchOn ? ' active' : ''}`} onClick={toggleTorch} title="Linterna"><Flashlight size={16} /> Luz</button>}
        {status === 'error' && <button type="button" className="qr-tool" onClick={() => { activeRef.current = true; startCamera(facing) }}><RotateCcw size={16} /> Reintentar</button>}
      </div>
      {error && <p className="auth-error">{error}</p>}
      <div className="qr-actions">
        <label className="qr-file-btn"><ImagePlus size={16} /> Leer imagen QR<input type="file" accept="image/*" capture="environment" onChange={scanImage} /></label>
      </div>
      <form className="qr-manual" onSubmit={submitManual}>
        <label>Código o ID del producto<input value={manual} onChange={(e) => setManual(e.target.value)} placeholder="Ej: VIT-001" /></label>
        <button type="submit" className="btn primary" disabled={!manual.trim()}><Search size={15} /> Buscar</button>
      </form>
      <small className="field-help qr-help">Si la cámara no funciona, escribe el código o toma una foto del QR.</small>
      <button className="btn secondary qr-close" onClick={onClose}>Cerrar escáner</button>
    </div>
  </div>
}

function AssistantIcon() {
  return <span className="assistant-logo" aria-hidden="true"><svg viewBox="0 0 32 32" role="img"><path className="assistant-antenna" d="M16 5V2m0 0-2 2m2-2 2 2" /><rect x="5" y="7" width="22" height="18" rx="6" /><path d="M10 15h.01M22 15h.01M11 20c3 2 7 2 10 0" /><path className="assistant-ear" d="M5 14H2m25 0h3" /></svg><b>IA</b></span>
}

function Assistant({ messages, onClose, onAsk, onDownload, onOpen, aiReady, memories = [], onSaveMemory, onDeleteMemory, onClearMemories }) {
  const [text, setText] = useState('')
  const [attachment, setAttachment] = useState(null)
  const [attachError, setAttachError] = useState('')
  const [memoryOpen, setMemoryOpen] = useState(false)
  const [memoryCategory, setMemoryCategory] = useState(MEMORY_CATEGORIES[0])
  const [memoryText, setMemoryText] = useState('')
  const fileInputRef = useRef(null)
  const bodyRef = useRef(null)

  useEffect(() => {
    const body = bodyRef.current
    if (body) body.scrollTop = body.scrollHeight
  }, [messages, memoryOpen])

  function addMemory(event) {
    event.preventDefault()
    const value = memoryText.trim()
    if (!value || !onSaveMemory) return
    onSaveMemory(memoryCategory, value)
    setMemoryText('')
  }

  function memoryDate(value) {
    const date = new Date(value)
    return Number.isNaN(date.getTime()) ? '' : date.toLocaleDateString('es-CL', { day: '2-digit', month: 'short', year: 'numeric' })
  }

  function chooseAttachment(event) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    if (file.size > 8 * 1024 * 1024) {
      setAttachError('El archivo supera el límite de 8 MB.')
      return
    }
    setAttachError('')
    setAttachment(file)
  }

  function submit(event) {
    event.preventDefault()
    if (!text.trim() && !attachment) return
    onAsk(text.trim() || 'Analiza el archivo adjunto y dime qué información adicional necesitas.', attachment)
    setText('')
    setAttachment(null)
    setAttachError('')
  }

  const busy = messages.some((m) => m.pending)
  const sizeKb = attachment ? Math.max(1, Math.ceil(attachment.size / 1024)) : 0

  return <div className="assistant-panel">
    <div className="assistant-head">
      <div className="assistant-title">
        <div className="bot-avatar"><img className="assistant-bot-image panel-bot-image" src="/bot-vitimax.png" alt="Robot BOT IA VITIMAX" /></div>
        <div><strong>BOT IA VITIMAX</strong><span>{aiReady ? 'IA conectada · Catálogo · Archivos · Memoria' : 'IA no configurada · respuestas locales'}</span></div>
      </div>
      <div className="assistant-head-actions">
        <button type="button" className={`assistant-memory-btn${memoryOpen ? ' active' : ''}`} onClick={() => setMemoryOpen((current) => !current)} aria-expanded={memoryOpen} title="Ver lo que la IA recuerda">
          <Brain size={14} /><span>Memoria</span>{memories.length > 0 && <b>{memories.length}</b>}
        </button>
        <button className="icon-btn" aria-label="Cerrar asistente" onClick={onClose}><X size={18} /></button>
      </div>
    </div>
    {memoryOpen ? <div className="assistant-memory">
      <div className="assistant-memory-head">
        <div><strong>Memoria de la IA</strong><span>{memories.length ? `${memories.length} recuerdo(s) guardado(s)` : 'Aún no hay recuerdos guardados'}</span></div>
        {memories.length > 0 && onClearMemories && <button type="button" className="assistant-memory-clear" onClick={onClearMemories}><Trash2 size={13} /> Borrar todo</button>}
      </div>
      <form className="assistant-memory-form" onSubmit={addMemory}>
        <select value={memoryCategory} onChange={(event) => setMemoryCategory(event.target.value)} aria-label="Categoría del recuerdo">
          {MEMORY_CATEGORIES.map((item) => <option key={item} value={item}>{item}</option>)}
        </select>
        <input value={memoryText} onChange={(event) => setMemoryText(event.target.value)} placeholder="Ej: Siempre aplica 19% de IVA" aria-label="Nota para recordar" maxLength={1200} />
        <button type="submit" disabled={!memoryText.trim()}><Plus size={14} /> Guardar</button>
      </form>
      <div className="assistant-memory-list">
        {!memories.length && <p className="assistant-memory-empty">Pídele al BOT “recuerda que mi nombre es…” o “siempre aplica 19% de IVA”. También puedes agregar una nota manual aquí.</p>}
        {memories.map((item) => <article className="assistant-memory-item" key={item.id}>
          <div>
            <span className="assistant-memory-tag">{item.categoria}</span>
            <p>{item.contenido}</p>
            {memoryDate(item.fecha_actualizacion) && <small>{memoryDate(item.fecha_actualizacion)}</small>}
          </div>
          {onDeleteMemory && <button type="button" onClick={() => onDeleteMemory(item.id)} aria-label={`Borrar recuerdo: ${item.contenido.slice(0, 40)}`} title="Borrar recuerdo"><Trash2 size={14} /></button>}
        </article>)}
      </div>
      <small className="assistant-memory-note">La IA usa estos datos para personalizar sus respuestas y los conserva al cerrar la app.</small>
    </div> : <>
    <div className="assistant-body" ref={bodyRef}>
      {messages.map((m, i) => <div className={`chat ${m.role}${m.pending ? ' pending' : ''}`} key={i}>
        <div>
          {m.text}
          {m.file && onDownload && <div className="chat-file-actions">
            <button type="button" className="chat-download" onClick={() => onDownload(m.file, i)}><ArrowDownToLine size={14} /><span>Descargar {m.file.name}</span></button>
            {m.file.native && m.file.uri && onOpen && <button type="button" className="chat-open" onClick={() => onOpen(m.file)}><ExternalLink size={14} /><span>Abrir archivo</span></button>}
          </div>}
        </div>
      </div>)}
    </div>
    <div className="quick-prompts">
      <button onClick={() => onAsk('Crea un catálogo de 10 productos con precios en CLP, stock y dosis de ejemplo, y genera el Excel con las 7 hojas.')}>Crear catálogo + Excel</button>
      <button onClick={() => onAsk('Simula las ventas de los próximos 30 días y dame recomendaciones para vender más.')}>Simular ventas</button>
      <button onClick={() => onAsk('Genera una guía completa de suelos, fertilizantes y plantas en PDF, con datos de ejemplo si hace falta.')}>Guía de suelos PDF</button>
      <button onClick={() => onAsk('Crea paso a paso una página web promocional usando el producto registrado que corresponda.')}>Crear página web</button>
      <button onClick={() => onAsk('Genera y descarga un Excel con mi catálogo actual.')}>Generar Excel</button>
      <button onClick={() => onAsk('Genera y descarga un PDF con mi catálogo actual.')}>Generar PDF</button>
      <button onClick={() => onAsk('Genera y descarga un documento Word con mi catálogo actual.')}>Generar Word</button>
      <button onClick={() => onAsk('Genera y descarga una presentación PowerPoint con mi catálogo actual.')}>Generar PowerPoint</button>
      <button onClick={() => onAsk('Genera y descarga el código de una página web.')}>Generar código</button>
    </div>
    {attachment && <div className="assistant-attachment">
      <span className="assistant-attachment-icon"><FileText size={16} /></span>
      <span className="assistant-attachment-name">{attachment.name}<small>{sizeKb} KB · listo para enviar</small></span>
      <button type="button" aria-label="Quitar archivo" title="Quitar archivo" onClick={() => setAttachment(null)}><X size={15} /></button>
    </div>}
    {attachError && <p className="assistant-attach-error">{attachError}</p>}
    <form className="assistant-input" onSubmit={submit}>
      <input value={text} onChange={(e) => setText(e.target.value)} placeholder="Pregunta, programa, analiza o pide un archivo…" />
      <button type="button" className="assistant-attach" title="Adjuntar archivo" aria-label="Adjuntar archivo" onClick={() => fileInputRef.current?.click()}><Paperclip size={18} /></button>
      <input ref={fileInputRef} type="file" accept="image/*,.pdf,.txt,.md,.json,.csv,.xlsx,.xls,.doc,.docx,.ppt,.pptx" hidden onChange={chooseAttachment} />
      <button type="submit" className="assistant-send" disabled={busy} aria-label="Enviar mensaje"><ArrowUpToLine size={17} /></button>
    </form>
    <small>Adjunta imágenes, PDF, Excel/CSV o texto (máximo 8 MB). Excel/CSV se analizan como texto y los PDF se leen directamente. Verifica diagnósticos y aplicaciones con un especialista.</small>
    </>}
  </div>
}

createRoot(document.getElementById('root')).render(<AuthGate />)
