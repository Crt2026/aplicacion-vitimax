const { onCall, HttpsError } = require('firebase-functions/v2/https')
const admin = require('firebase-admin')

admin.initializeApp()
const vertexModel = 'gemini-2.0-flash-001'
const vertexLocation = 'us-central1'
const DAILY_REQUEST_LIMIT = 20
const MAX_MESSAGE_LENGTH = 4000
const MAX_CATALOG_LENGTH = 120000
const MAX_ATTACHMENT_LENGTH = 8_000_000
const allowedAttachmentTypes = new Set([
  'image/jpeg',
  'image/png',
  'image/webp',
  'application/pdf',
  'text/plain',
  'text/markdown',
  'text/csv',
  'application/json'
])

const SPANISH_ONLY_RULE = 'IDIOMA OBLIGATORIO: Responde SIEMPRE en español natural y profesional (Chile). Está estrictamente prohibido escribir pensamientos internos, notas técnicas, nombres de herramientas o explicaciones en inglés en tus respuestas al usuario.'

const systemInstruction = `${SPANISH_ONLY_RULE}

Eres el BOT IA oficial de VITIMAX. Ayudas a clientes y administradores con catálogo, inventario, ventas, Excel, PDF, Word, PowerPoint, publicaciones, creación de páginas web, análisis de archivos y cualquier consulta general.
Reglas obligatorias:
- Si el catálogo tiene productos, úsalos como fuente principal y respeta sus datos reales.
- Si el catálogo está VACÍO y el usuario pide un catálogo, productos de ejemplo, plantas, tierra, abonos, fertilizantes, precios, stock o dosis, GENERA datos completos y realistas (ID, categoría, precio en CLP, stock, formato, descripción, dosis, recomendaciones). No te niegues y no entregues un archivo vacío.
- Marca siempre los datos inventados como "Generado por IA (ejemplo)" y no los presentes como información oficial de VITIMAX.
- Puedes combinar el catálogo real con datos nuevos solo si el usuario lo pide explícitamente.
- Para fotografías de plantas, ofrece una orientación preliminar, explica la incertidumbre y recomienda revisión de un especialista de VITIMAX antes de aplicar productos.
- No confirmes diagnósticos ni dosis si no aparecen en una ficha técnica oficial.
- Cuando pidan una página web, guía paso a paso: objetivo, público, producto, contenido confirmado, llamada a la acción y estructura final.
- Si el usuario pide un archivo o una acción, responde además con un bloque \`\`\`json: {"accion":"crear_excel","filename":"catalogo-vitimax.xlsx","resumen":"...","productos":[...],"sheets":[{"name":"Productos","columns":[...],"rows":[[...]]}]} incluyendo las 7 hojas (Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores); {"accion":"registrar_productos_en_app","productos":[...]}; {"accion":"simular_ventas","periodoDias":30}; o {"accion":"crear_documento","documento":{"tipo":"pdf|word|ppt","titulo":"...","contenido":"..."}}. En Excel, Productos y Clientes son la fuente de IDs/nombres: Inventario, Ventas y Asesorias los repiten exactamente (Ventas y Asesorias incluyen "ID Cliente", Ventas incluye "ID Producto"); Plantas usa PLT-001… y Proveedores PRV-001… en la columna A; ninguna fecha de Ventas supera la fecha actual; fórmulas en inglés con coma y formato CLP "$#,##0".
- No uses "-" como dato: deja el campo vacío o escribe "Pendiente". Precios en CLP como números.
- Responde en español, de forma clara, práctica y accionable.

${SPANISH_ONLY_RULE}`

const REASONING_START_PATTERNS = [
  /^\s*(?:analysis|reasoning|thinking|thoughts?|plan|notas? internas?|pensamiento(?:s)? internos?)\s*[:：]/i,
  /^\s*(?:the user is asking|looking at (?:the )?context|let me (?:create|start|use|check|prepare|build|write|make|review|try)|i (?:need to|should|must|will|have to|can|am going to)|i'll|we (?:need to|should|must|will)|now i (?:need|have|should)|for (?:the|this|excel|pdf|the catalog|the file)|since (?:the|this|only)|based on (?:the )?context|according to (?:the )?context|the (?:context|catalog|user|app|file) (?:is|has|shows|says)|so i|first,? i|next,? i|my plan|to do this)\b/i
]
const ENGLISH_HINT_RE = /\b(?:the|and|is|are|to|of|for|with|this|that|i|we|my|have|has|need|should|will|can|from|user|context|catalog|file|app|so|now|let|based|according|since|only|also|then|but|if|it|its|as|at|on|in|by|files|both|create|use|check)\b/gi
const SPANISH_HINT_RE = /\b(?:el|la|los|las|de|del|que|y|en|con|por|para|es|son|al|un|una|se|su|sus|te|tu|necesito|debo|voy|puedo|archivo|archivos|catálogo|catalogo|producto|productos|cliente|clientes|venta|ventas|informe|excel|pdf|word|resumen|total|precio|stock|fecha|generé|generar|creé|crear|lista|datos)\b/gi

// Nunca se leen campos reasoning / reasoning_content: solo texto visible.
function sanitizeAssistantText(value) {
  const source = typeof value === 'string' ? value : ''
  if (!source) return ''
  const segments = source.split(/(```[\s\S]*?```)/g)
  const cleaned = segments.map((segment, index) => {
    if (index % 2 === 1) return segment
    let text = segment
      .replace(/<\s*think(?:\s[^>]*)?\s*>[\s\S]*?(?:<\s*\/\s*think\s*>|$)/gi, '')
      .replace(/<\s*\/\s*think\s*>/gi, '')
      .replace(/<\|(?:\s*thinking|\s*analysis|\s*reasoning)\s*\|>/gi, '')
    return text.split('\n').filter((line) => {
      const trimmed = line.trim()
      if (!trimmed) return true
      if (REASONING_START_PATTERNS.some((pattern) => pattern.test(trimmed))) return false
      const english = (trimmed.match(ENGLISH_HINT_RE) || []).length
      const spanish = (trimmed.match(SPANISH_HINT_RE) || []).length
      return !(english >= 4 && english > spanish * 2)
    }).join('\n').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n')
  })
  return cleaned.join('').trim()
}

const TEMPORAL_MARKER = 'RELOJ Y CONTEXTO TEMPORAL EN VIVO'

// La app envía el system prompt ya construido con el reloj del dispositivo.
// Este respaldo evita fechas fijas si una versión anterior del cliente no lo envía.
function buildCloudSystem(system, temporal) {
  if (typeof system === 'string' && system.trim() && system.includes(TEMPORAL_MARKER)) return system.slice(0, 100_000)
  const clock = temporal && typeof temporal === 'object' ? temporal : {}
  const when = new Date()
  const fecha = String(clock.fechaCompleta || when.toLocaleDateString('es-CL', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })).slice(0, 120)
  const hora = String(clock.hora24 || when.toLocaleTimeString('es-CL', { hour12: false })).slice(0, 12)
  const zona = String(clock.zonaHoraria || 'America/Santiago').slice(0, 80)
  const franja = ['Buenos días', 'Buenas tardes', 'Buenas noches'].includes(clock.franja) ? clock.franja : 'Buenas noches'
  const fallback = `\n\n### ${TEMPORAL_MARKER}\nHoy es ${fecha}, son las ${hora} (${franja}). Zona horaria: ${zona}. Usa esta referencia exacta para saludar, calcular plazos y fechar entregables; no uses fechas memorizadas.`
  const base = typeof system === 'string' && system.trim() ? system.slice(0, 100_000) : systemInstruction
  return base.includes(TEMPORAL_MARKER) ? base : `${base}${fallback}`
}

exports.vitimaxAssistant = onCall(
  { region: vertexLocation, enforceAppCheck: false },
  async (request) => {
    if (!request.auth) throw new HttpsError('unauthenticated', 'Debes iniciar sesión para usar la IA.')
    const { message, products = [], attachment, system, temporal, maxTokens = 6000 } = request.data || {}
    if (typeof message !== 'string' || !message.trim()) {
      throw new HttpsError('invalid-argument', 'La consulta está vacía.')
    }
    if (message.trim().length > MAX_MESSAGE_LENGTH) {
      throw new HttpsError('invalid-argument', 'La consulta supera el límite de 4.000 caracteres.')
    }
    if (!Array.isArray(products) || products.length > 200) {
      throw new HttpsError('invalid-argument', 'El catálogo enviado no es válido.')
    }

    const catalog = JSON.stringify(products.map(({ id, name, category, price, stock, format, description, usage, recommendations, features }) => ({
      id, name, category, price, stock, format, description, usage, recommendations, features
    })))
    if (catalog.length > MAX_CATALOG_LENGTH) {
      throw new HttpsError('invalid-argument', 'El catálogo supera el límite permitido para la consulta.')
    }
    const parts = [{ text: `CATÁLOGO REAL DE VITIMAX:\n${catalog}\n\nCONSULTA DEL USUARIO:\n${message.trim()}` }]
    if (attachment) {
      const name = String(attachment.name || 'archivo').slice(0, 120)
      // Excel/CSV y texto llegan ya convertidos a texto desde la app.
      if (typeof attachment.text === 'string' && attachment.text.trim()) {
        if (attachment.text.length > MAX_ATTACHMENT_LENGTH) {
          throw new HttpsError('invalid-argument', 'El archivo adjunto supera el límite de 8 MB.')
        }
        parts.push({ text: `ARCHIVO ADJUNTO "${name}":\n${attachment.text.slice(0, 24_000)}${attachment.text.length > 24_000 ? '\n[Contenido recortado]' : ''}` })
      } else {
        if (typeof attachment.data !== 'string' || typeof attachment.mimeType !== 'string') {
          throw new HttpsError('invalid-argument', 'El archivo adjunto no es válido.')
        }
        if (!allowedAttachmentTypes.has(attachment.mimeType)) {
          throw new HttpsError('invalid-argument', 'Ese tipo de archivo no está permitido para el asistente.')
        }
        if (attachment.data.length > MAX_ATTACHMENT_LENGTH) {
          throw new HttpsError('invalid-argument', 'El archivo adjunto supera el límite de 8 MB.')
        }
        parts.push({ text: `ARCHIVO ADJUNTO: ${name} (${attachment.mimeType})` })
        if (attachment.mimeType.startsWith('image/') || attachment.mimeType === 'application/pdf') {
          parts.push({ inlineData: { mimeType: attachment.mimeType, data: attachment.data } })
        }
      }
    }

    const today = new Date().toISOString().slice(0, 10)
    const usageRef = admin.firestore()
      .collection('users')
      .doc(request.auth.uid)
      .collection('assistantUsage')
      .doc(today)
    await admin.firestore().runTransaction(async (transaction) => {
      const usage = (await transaction.get(usageRef)).data() || {}
      const count = Number(usage.count) || 0
      if (count >= DAILY_REQUEST_LIMIT) {
        throw new HttpsError('resource-exhausted', `Alcanzaste el límite de ${DAILY_REQUEST_LIMIT} consultas de IA por día.`)
      }
      transaction.set(usageRef, {
        count: count + 1,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      }, { merge: true })
    })

    const accessToken = await admin.app().options.credential.getAccessToken()
    const projectId = process.env.GCLOUD_PROJECT || process.env.GCP_PROJECT
    if (!projectId || !accessToken?.access_token) {
      throw new HttpsError('failed-precondition', 'Google Cloud no está configurado para la IA.')
    }
    const endpoint = `https://${vertexLocation}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${vertexLocation}/publishers/google/models/${vertexModel}:generateContent`
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken.access_token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: buildCloudSystem(system, temporal) }] },
        contents: [{ role: 'user', parts }],
        generationConfig: { temperature: 0.2, maxOutputTokens: Math.min(Math.max(Number(maxTokens) || 6000, 100), 6000) }
      })
    })
    if (!response.ok) {
      const errorText = await response.text()
      console.error('Vertex AI request failed', response.status, errorText.slice(0, 500))
      throw new HttpsError('unavailable', 'Google Gemini no pudo responder. Inténtalo nuevamente.')
    }
    const result = await response.json()
    const text = sanitizeAssistantText(result.candidates?.[0]?.content?.parts?.map((part) => part.text || '').join(''))
    if (typeof text !== 'string' || !text.trim()) {
      throw new HttpsError('unavailable', 'Google Gemini devolvió una respuesta vacía.')
    }
    return { text }
  }
)
