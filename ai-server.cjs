const http = require('node:http')
const fs = require('node:fs')
const path = require('node:path')

const envPath = path.join(process.cwd(), '.env')
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, 'utf8').split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z][A-Z0-9_]*)\s*=\s*(.*)\s*$/)
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2].replace(/^['"]|['"]$/g, '')
  }
}

const port = Number(process.env.PORT || process.env.AI_SERVER_PORT || 8787)
const host = process.env.AI_SERVER_HOST || '0.0.0.0'
const provider = (process.env.AI_PROVIDER || (process.env.OPENROUTER_API_KEY ? 'openrouter' : 'gemini')).toLowerCase()
const apiKey = provider === 'openrouter' ? process.env.OPENROUTER_API_KEY : process.env.GEMINI_API_KEY
const model = provider === 'openrouter'
  ? (process.env.OPENROUTER_MODEL || 'openai/gpt-4o-mini')
  : (process.env.GEMINI_MODEL || 'gemini-2.0-flash')
const fallbackModels = (process.env.OPENROUTER_FALLBACK_MODELS || 'google/gemini-2.5-flash,openrouter/free')
  .split(',')
  .map((item) => item.trim())
  .filter(Boolean)
const sharedToken = (process.env.AI_SHARED_TOKEN || '').trim()
const upstreamTimeout = Number(process.env.AI_TIMEOUT_MS || 60_000)
const rateLimitWindow = 60_000
const rateLimitMax = Number(process.env.AI_RATE_LIMIT_PER_MINUTE || 30)
const maxBodyBytes = 12 * 1024 * 1024
const maxHistoryItems = 12
const maxMessageLength = 4000
const maxMemories = 100

const rawOrigins = (process.env.AI_ALLOWED_ORIGINS || '*')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean)
const allowAllOrigins = rawOrigins.length === 0 || rawOrigins.includes('*')

const systemInstruction = `Eres el BOT IA oficial de VITIMAX, una plataforma para administrar catálogo, inventario, pedidos, clientes, asesorías y soluciones agrícolas.
Responde siempre en español, de forma clara, cercana y práctica. Puedes ayudar con CUALQUIER tema o tarea: catálogo, inventario, ventas, Excel, PDF, Word, PowerPoint, páginas web, marketing, análisis de archivos, agricultura, negocios, cálculos, redacción, código y consultas generales.

REGLA PRINCIPAL SOBRE LOS DATOS:
- Si el catálogo tiene productos, úsalos como fuente principal y cítalos con sus datos reales.
- Si el catálogo está VACÍO y el usuario pide crear un catálogo, productos de ejemplo, tierra, abonos, fertilizantes, plantas, precios, stock, dosis o cualquier dato concreto, GENERA datos completos, coherentes y realistas (ID, categoría, precio en CLP, stock, formato, descripción, dosis y recomendaciones). NO te niegues, NO respondas solo que falta el catálogo y NO entregues un archivo vacío.
- Si el usuario pide inventar, simular, ejemplos o datos DEMO, hazlo y dilo claramente: "Datos generados por IA (ejemplo)".
- No presentes datos generados como si fueran información oficial de VITIMAX. Separa siempre "Datos reales del catálogo" de "Datos generados por IA".
- Puedes combinar el catálogo real con datos nuevos solo cuando el usuario lo pida explícitamente.

ACCIONES EN LA APP (bloques JSON):
Cuando el usuario pida un archivo o una acción, responde con un bloque \`\`\`json y el resto del texto que quieras mostrar:
- Excel / catálogo / planilla: {"accion":"crear_excel","filename":"catalogo-vitimax.xlsx","resumen":"...","productos":[{...}],"sheets":[{"name":"Productos","columns":[...],"rows":[[...]]}]}. Incluye SIEMPRE las 7 hojas: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores.
- Guardar productos en el catálogo: {"accion":"registrar_productos_en_app","productos":[{...}],"reemplazar":false}
- Simular ventas: {"accion":"simular_ventas","periodoDias":30,"crecimientoMensual":0}
- Documento PDF/Word/PowerPoint: {"accion":"crear_documento","documento":{"tipo":"pdf|word|ppt","titulo":"...","contenido":"# Título\\n## Sección\\n- Punto"}}
- Guardar un recuerdo persistente: {"accion":"guardar_memoria","categoria":"Negocio|Preferencias|Instrucciones|Datos personales|Clientes|Productos|Finanzas|Otros","dato":"..."}
- Olvidar o actualizar un recuerdo: {"accion":"olvidar_memoria","clave_o_id":"...","nuevo_dato":"..."}

MEMORIA PERSISTENTE:
- Si el usuario dice "recuerda que...", "mi nombre es...", "siempre aplica...", "ten en cuenta...", guarda el dato con la acción guardar_memoria.
- Si dice que un dato cambió o ya no aplica, usa olvidar_memoria (con nuevo_dato para actualizarlo).
- Aplica la memoria disponible de forma natural, sin repetirla robóticamente.
- No guardes contraseñas, claves, tokens ni datos de tarjetas.

REGLAS DE TRABAJO:
- No uses "-" como dato. Si falta un dato, deja el campo vacío o escribe "Pendiente".
- Precios en CLP como números, sin símbolos ni separadores. Stock y cantidades como números.
- Para fotografías de plantas o cultivos, ofrece orientación preliminar, explica la incertidumbre y recomienda revisión de un especialista antes de aplicar productos.
- No confirmes diagnósticos fitosanitarios definitivos ni dosis oficiales si no están en una ficha técnica; si son ejemplos, indícalo.
- Usa listas y pasos cuando ayuden. Ten en cuenta el historial de la conversación.`

const rateBuckets = new Map()

function isRateLimited(ip) {
  const now = Date.now()
  const bucket = (rateBuckets.get(ip) || []).filter((time) => now - time < rateLimitWindow)
  bucket.push(now)
  rateBuckets.set(ip, bucket)
  return bucket.length > rateLimitMax
}

setInterval(() => {
  const now = Date.now()
  for (const [ip, bucket] of rateBuckets) {
    if (!bucket.some((time) => now - time < rateLimitWindow)) rateBuckets.delete(ip)
  }
}, 60_000).unref?.()

function corsHeaders(origin) {
  if (allowAllOrigins) {
    return {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Content-Type, X-Vitimax-Token',
      'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
      'Access-Control-Max-Age': '86400'
    }
  }
  const allowed = origin && rawOrigins.includes(origin)
  return {
    ...(allowed ? { 'Access-Control-Allow-Origin': origin, Vary: 'Origin' } : {}),
    'Access-Control-Allow-Headers': 'Content-Type, X-Vitimax-Token',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Max-Age': '86400'
  }
}

function sendJson(request, response, status, body) {
  response.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    ...corsHeaders(request.headers.origin)
  })
  response.end(JSON.stringify(body))
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    let body = ''
    request.on('data', (chunk) => {
      body += chunk
      if (body.length > maxBodyBytes) {
        request.destroy()
        reject(new Error('La solicitud supera el límite de 12 MB.'))
      }
    })
    request.on('end', () => resolve(body))
    request.on('error', reject)
  })
}

function sanitizeHistory(history) {
  if (!Array.isArray(history)) return []
  return history
    .filter((item) => item && (item.role === 'user' || item.role === 'assistant') && typeof item.content === 'string' && item.content.trim())
    .slice(-maxHistoryItems)
    .map((item) => ({ role: item.role, content: item.content.trim().slice(0, maxMessageLength) }))
}

function sanitizeMemories(memories) {
  if (!Array.isArray(memories)) return []
  const seen = new Set()
  const list = []
  for (const item of memories) {
    if (!item || typeof item !== 'object') continue
    const contenido = String(item.contenido ?? item.content ?? item.dato ?? '').trim().slice(0, 1200)
    if (!contenido) continue
    const categoria = String(item.categoria || item.category || 'Otros').trim().slice(0, 40) || 'Otros'
    const key = `${categoria.toLowerCase()}|${contenido.toLowerCase()}`
    if (seen.has(key)) continue
    seen.add(key)
    list.push({ categoria, contenido })
    if (list.length >= maxMemories) break
  }
  return list
}

function memoryBlock(memories) {
  const list = sanitizeMemories(memories)
  if (!list.length) return ''
  const lines = []
  let total = 0
  for (const item of list) {
    const line = `- [${item.categoria}] ${item.contenido.slice(0, 280)}`
    if (total + line.length > 7000) break
    lines.push(line)
    total += line.length + 1
  }
  if (!lines.length) return ''
  return `

### MEMORIA PERSISTENTE DEL USUARIO Y NEGOCIO
Datos que el usuario pidió recordar entre sesiones. Úsalos de forma natural, sin repetirlos robóticamente ni contradecirlos.
${lines.join('\n')}`
}

function attachmentParts(attachment) {
  const parts = []
  if (!attachment || typeof attachment !== 'object') return parts
  const name = String(attachment.name || 'archivo').slice(0, 160)
  const mimeType = typeof attachment.mimeType === 'string' ? attachment.mimeType : 'application/octet-stream'
  // Excel/CSV y texto ya vienen convertidos a texto desde la app.
  if (typeof attachment.text === 'string' && attachment.text.trim()) {
    parts.push({ kind: 'text', text: `ARCHIVO ADJUNTO "${name}":\n${attachment.text.slice(0, 24_000)}${attachment.text.length > 24_000 ? '\n[Contenido recortado]' : ''}` })
    return parts
  }
  if (typeof attachment.data !== 'string') return parts
  if (mimeType.startsWith('image/')) {
    parts.push({ kind: 'image', mimeType, data: attachment.data })
    return parts
  }
  if (mimeType === 'application/pdf') {
    parts.push({ kind: 'file', mimeType, data: attachment.data, name })
    return parts
  }
  const isText = mimeType.startsWith('text/') || /(json|csv|xml|javascript)/i.test(mimeType) || /\.(txt|md|json|csv|xml|js|ts|html|css)$/i.test(name)
  if (isText) {
    let text = ''
    try {
      text = Buffer.from(attachment.data, 'base64').toString('utf8')
    } catch {
      text = ''
    }
    if (text) parts.push({ kind: 'text', text: `ARCHIVO ADJUNTO "${name}":\n${text.slice(0, 20_000)}${text.length > 20_000 ? '\n[Contenido recortado]' : ''}` })
    return parts
  }
  parts.push({ kind: 'text', text: `ARCHIVO ADJUNTO: "${name}" (${mimeType}). No puedo leer este formato directamente; pide al usuario que lo convierta a texto, CSV, Excel, PDF o imagen.` })
  return parts
}

async function askOpenRouter({ catalog, message, history, attachment, signal, system, memories, maxTokens = 1500 }) {
  const content = []
  const attachmentItems = attachmentParts(attachment)
  const textBlocks = [`CATÁLOGO REAL DE VITIMAX:\n${catalog}`, `CONSULTA DEL USUARIO:\n${message}`]
  for (const item of attachmentItems) {
    if (item.kind === 'text') textBlocks.push(item.text)
  }
  content.push({ type: 'text', text: textBlocks.join('\n\n') })
  for (const item of attachmentItems) {
    if (item.kind === 'image') content.push({ type: 'image_url', image_url: { url: `data:${item.mimeType};base64,${item.data}` } })
    if (item.kind === 'file') content.push({ type: 'file', file: { filename: item.name, file_data: `data:${item.mimeType};base64,${item.data}` } })
  }
  const messages = [
    { role: 'system', content: `${system || systemInstruction}${memoryBlock(memories)}` },
    ...history.map((item) => ({ role: item.role, content: item.content })),
    { role: 'user', content }
  ]
  const modelList = [model, ...fallbackModels].filter(Boolean).slice(0, 3)
  const result = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    signal,
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
      'HTTP-Referer': process.env.AI_PUBLIC_URL || 'https://vitimax.app',
      'X-Title': 'VITIMAX'
    },
    body: JSON.stringify({ models: modelList, messages, temperature: 0.3, max_tokens: Math.min(Math.max(Number(maxTokens) || 1500, 100), 6000) })
  })
  const payload = await result.json().catch(() => ({}))
  if (!result.ok) {
    const reason = payload?.error?.message || `HTTP ${result.status}`
    const error = new Error(`OpenRouter rechazó la solicitud: ${reason}`)
    error.status = result.status
    throw error
  }
  return { text: payload.choices?.[0]?.message?.content?.trim(), model: payload.model || model }
}

async function askGemini({ catalog, message, history, attachment, signal, system, memories, maxTokens = 1500 }) {
  const parts = [{ text: `CATÁLOGO REAL DE VITIMAX:\n${catalog}\n\nCONSULTA DEL USUARIO:\n${message}` }]
  for (const item of attachmentParts(attachment)) {
    if (item.kind === 'text') parts.push({ text: item.text })
    if (item.kind === 'image') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
    if (item.kind === 'file') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
  }
  const contents = history.map((item) => ({
    role: item.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: item.content }]
  }))
  contents.push({ role: 'user', parts })
  const result = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      system_instruction: { parts: [{ text: `${system || systemInstruction}${memoryBlock(memories)}` }] },
      contents,
      generationConfig: { temperature: 0.3, maxOutputTokens: Math.min(Math.max(Number(maxTokens) || 1500, 100), 8000) }
    })
  })
  const payload = await result.json().catch(() => ({}))
  if (!result.ok) {
    const reason = payload?.error?.message || `HTTP ${result.status}`
    const error = new Error(`Google Gemini rechazó la solicitud: ${reason}`)
    error.status = result.status
    throw error
  }
  return { text: payload.candidates?.[0]?.content?.parts?.map((part) => part.text || '').join('').trim(), model }
}

const server = http.createServer(async (request, response) => {
  if (request.method === 'OPTIONS') {
    response.writeHead(204, corsHeaders(request.headers.origin))
    return response.end()
  }

  if (request.method === 'GET' && (request.url === '/health' || request.url === '/')) {
    return sendJson(request, response, 200, {
      ok: true,
      service: 'vitimax-ai',
      provider,
      model,
      keyConfigured: Boolean(apiKey),
      rateLimitPerMinute: rateLimitMax
    })
  }

  if (request.method !== 'POST' || request.url !== '/api/assistant') {
    return sendJson(request, response, 404, { error: 'Ruta no encontrada.' })
  }

  if (!apiKey) {
    return sendJson(request, response, 503, {
      error: provider === 'openrouter'
        ? 'Falta OPENROUTER_API_KEY en el archivo .env o en las variables del servidor.'
        : 'Falta GEMINI_API_KEY en el archivo .env o en las variables del servidor.'
    })
  }

  if (sharedToken && request.headers['x-vitimax-token'] !== sharedToken) {
    return sendJson(request, response, 401, { error: 'Token de acceso al asistente inválido.' })
  }

  const ip = request.socket.remoteAddress || 'unknown'
  if (isRateLimited(ip)) {
    return sendJson(request, response, 429, { error: 'Demasiadas consultas seguidas. Espera un minuto e inténtalo de nuevo.' })
  }

  try {
    const body = JSON.parse(await readBody(request))
    const { message, products = [], attachment, history = [], system, memories = [], maxTokens } = body || {}
    if (typeof message !== 'string' || !message.trim()) {
      return sendJson(request, response, 400, { error: 'La consulta está vacía.' })
    }
    if (message.trim().length > maxMessageLength) {
      return sendJson(request, response, 400, { error: `La consulta supera el límite de ${maxMessageLength} caracteres.` })
    }
    if (!Array.isArray(products) || products.length > 200) {
      return sendJson(request, response, 400, { error: 'El catálogo no es válido.' })
    }

    const catalog = JSON.stringify(products.map(({ id, name, category, price, stock, format, description, usage, recommendations, features }) => ({
      id, name, category, price, stock, format, description, usage, recommendations, features
    })))
    const safeHistory = sanitizeHistory(history)
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), upstreamTimeout)
    let result
    try {
      result = await (provider === 'openrouter' ? askOpenRouter : askGemini)({
        catalog,
        message: message.trim(),
        history: safeHistory,
        attachment,
        signal: controller.signal,
        system: typeof system === 'string' && system.trim() ? system.slice(0, 4000) : undefined,
        memories: sanitizeMemories(memories),
        maxTokens
      })
    } finally {
      clearTimeout(timer)
    }

    const text = result?.text
    if (!text) {
      return sendJson(request, response, 502, { error: `${provider === 'openrouter' ? 'OpenRouter' : 'Google Gemini'} devolvió una respuesta vacía.` })
    }
    return sendJson(request, response, 200, { text, provider, model: result.model || model })
  } catch (error) {
    if (error?.name === 'AbortError') {
      return sendJson(request, response, 504, { error: 'La IA tardó demasiado en responder. Inténtalo de nuevo.' })
    }
    console.error('[vitimax-ai]', error?.message || error)
    const status = Number(error?.status)
    const message = status === 401 || status === 403
      ? 'La clave de la IA fue rechazada. Revisa OPENROUTER_API_KEY / GEMINI_API_KEY.'
      : status === 402
        ? 'La cuenta de OpenRouter no tiene saldo disponible.'
        : status === 429
          ? 'La IA alcanzó su límite temporal. Espera un momento e inténtalo otra vez.'
          : error?.message?.startsWith('La solicitud')
            ? error.message
            : 'No se pudo procesar la consulta con la IA. Revisa la configuración del servidor.'
    return sendJson(request, response, status && status >= 400 ? 502 : 400, { error: message })
  }
})

server.listen(port, host, () => {
  console.log(`VITIMAX ${provider} server listening on http://${host}:${port}`)
  console.log(`Modelo: ${model} · Clave configurada: ${apiKey ? 'sí' : 'no'}`)
})
