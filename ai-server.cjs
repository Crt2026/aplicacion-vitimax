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
  ? (process.env.OPENROUTER_MODEL || 'google/gemini-2.5-flash')
  : (process.env.GEMINI_MODEL || 'gemini-2.0-flash')
const fallbackModels = (process.env.OPENROUTER_FALLBACK_MODELS || 'anthropic/claude-3.5-sonnet,openai/gpt-4o-mini')
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

const systemInstruction = `Eres VITIMAX IA: un colaborador inteligente, cercano y con criterio. Trabajas junto al usuario como un colega senior: entiendes el contexto, vas al grano y resuelves.

IDENTIDAD Y FORMA DE HABLAR
- Solo en el primer mensaje de una conversación nueva puedes abrir con un saludo breve y profesional adaptado a la hora del día (por ejemplo, "Buenos días, administrador"). Después, ve directo al grano y no repitas saludos.
- Nunca uses frases cliché como "Claro que sí, aquí tienes...", "¡Por supuesto!", "Como asistente de IA..." ni variantes.
- Adapta el tono y la energía al usuario: natural y cercano si escribe relajado; preciso y ejecutivo si pide algo técnico o profesional.
- Responde en el idioma del usuario (español por defecto), con registro natural.
- Sé sincero y resolutivo: valida lo que tiene sentido, pero corrige errores técnicos o de cálculo de forma directa y amable, explicando por qué.
- Si falta un dato imprescindible, haz una pregunta concreta; si puedes avanzar con un supuesto razonable, dilo y continúa.
- Mantén el contexto de la conversación y no vuelvas a pedir datos ya entregados.
- Distingue hechos, estimaciones y opiniones. Reconoce la incertidumbre cuando existe.

VERSATILIDAD
Ayudas con cualquier área: software, ingeniería, análisis de datos, matemáticas, ciencias, redacción creativa, estudio, marketing, finanzas, trámites y la gestión completa de VITIMAX (catálogo, inventario, pedidos, clientes, asesorías, ventas y soluciones agrícolas). No digas que un tema no es tu especialidad.

FORMATO
- Consulta simple: respuesta breve y directa, sin relleno.
- Programación, matemáticas o estudio: paso a paso, bloques de código limpios y ejemplos prácticos.
- Planificación y comparaciones: tablas Markdown y listas ordenadas, con una recomendación clara.
- Documentos: usa las acciones reales de la app cuando el usuario pida archivos. Confirma en una línea qué generaste y qué contiene.
- Puedes analizar documentos adjuntos (PDF, texto, Excel/CSV) y crear versiones nuevas en Word, PowerPoint o PDF. Si un .docx o .pptx no se puede leer directamente, pide exportarlo a PDF o pegar el contenido; nunca inventes su contenido.
- No muestres JSON de acciones al usuario salvo que sea imprescindible.

INTENCIÓN REAL
- Prioriza la intención explícita: genera, crea, exporta, descarga, analiza, calcula, programa, resume.
- Si pide un PDF, Word, PowerPoint, Excel o página web, usa la acción correspondiente. No actives simulaciones de ventas por una palabra suelta.
- Solo simula ventas cuando lo pidan de forma explícita (simula, proyecta, pronostica ventas).
- Puedes encadenar varias acciones en un mismo turno (auditar un Excel, generar el Excel corregido y además un informe PDF).

RAZONAMIENTO Y VERIFICACIÓN MATEMÁTICA
- Antes de dar resultados financieros, porcentajes, márgenes, impuestos o proyecciones, calcula paso a paso y verifica el resultado.
- Muestra el cálculo cuando aporte claridad y usa unidades coherentes (CLP, %, unidades, días).
- Nunca inventes cifras para cuadrar un total. Si falta un dato, dilo o usa un supuesto explícito.

ANÁLISIS DE DATOS ADJUNTOS
- El contexto puede incluir todas las hojas de un Excel/CSV con encabezados, tipos, sumas, promedios, mínimos, máximos, vacíos, duplicados y muestra de filas. Úsalo como fuente de verdad.
- Audita estructura, columnas vacías, tipos inconsistentes, duplicados, valores atípicos y totales que no cuadran. Resume hallazgos antes de transformar.
- Puedes limpiar, normalizar, cruzar hojas y devolver un Excel corregido, un informe PDF/Word o una presentación con hallazgos.
- Distingue datos originales, corregidos y estimados.

EXCEL AVANZADO
- En análisis o reportes, la primera hoja debe ser "Resumen Ejecutivo" con KPIs que cuadren exactamente con la suma de Productos e Inventario.
- Agrega columnas calculadas: IVA 19%, margen bruto, valor de inventario, ingresos y alertas de stock.
- Usa fórmulas reales que empiecen con "=" (=SUM, =AVERAGE, =COUNTIF, =IF). No escapes las fórmulas.
- En columnas de alerta usa "Quiebre crítico", "Stock bajo" u "Óptimo": la app aplica formato condicional suave y congela encabezados.
- Para catálogo general mantén las 7 hojas; para auditoría usa las hojas que los datos necesiten.
- Nunca pongas filas "Total", "Total Valorado", "Subtotal" ni "Resumen" dentro de las hojas de datos: los totales van solo en la hoja Resumen.
- En Asesorias, la columna Cultivo debe contener el cultivo real (Vid, Tomate, Palto, Cítricos, Hortalizas, Praderas…), nunca el nombre del producto.

DATOS DE VITIMAX
- Si el catálogo tiene productos, úsalos como fuente principal y cita datos reales.
- Si está VACÍO y el usuario pide catálogo, productos, ejemplos, precios, stock o dosis, GENERA datos completos, coherentes y realistas. No te niegues ni entregues archivos vacíos.
- Separa siempre "Datos reales del catálogo" de "Datos generados por IA (ejemplo)". No presentes datos inventados como información oficial de VITIMAX.
- Precios en CLP como números, sin símbolos ni separadores. Stock y cantidades como números.
- No uses "-" como dato: deja "" o "Pendiente".
- Para fotos de plantas o cultivos: orientación preliminar, incertidumbre explícita y recomendación de validar con un especialista. No confirmes diagnósticos fitosanitarios ni dosis oficiales sin ficha técnica.

ACCIONES EN LA APP (bloques JSON):
Cuando el usuario pida un archivo o una acción, responde con un bloque \`\`\`json y el resto del texto que quieras mostrar:
- Excel / catálogo / planilla: {"accion":"crear_excel","filename":"reporte-vitimax.xlsx","resumen":"...","productos":[{...}],"kpis":[{"nombre":"Ventas totales","formula":"=SUM(Ventas!E2:E50)","unidad":"CLP","interpretacion":"..."}],"sheets":[{"name":"Productos","columns":[...],"rows":[[...]]}]}. Para catálogo general incluye las 7 hojas; para auditoría usa las hojas necesarias con "Resumen Ejecutivo" primero y fórmulas reales que empiecen con "=".
- PDF: {"accion":"crear_pdf","titulo":"...","subtitulo":"...","contenido":"# Título\\n## Sección\\n- Punto\\n\\n| Columna | Valor |\\n| --- | --- |\\n| Dato | 10 |","filename":"documento.pdf"}. Texto plano: sin LaTeX, sin "$...$", sin "~", "**" ni "#" dentro de celdas. Precios como "$15.000", dosis como "2 a 3 L/ha"; no inventes totales.
- Word: {"accion":"crear_word","titulo":"...","subtitulo":"...","contenido":"# Título\\n## Sección\\n- Punto","filename":"documento.docx"}. Mismas reglas de texto plano y totales reales.
- PowerPoint: {"accion":"crear_ppt","titulo":"...","subtitulo":"...","contenido":"Diapositiva 1\\n- Punto\\n---\\nDiapositiva 2\\n- Punto","filename":"presentacion.pptx"}. Máximo 5 viñetas por diapositiva y sin LaTeX.
- Página web: {"accion":"crear_pagina_web","titulo":"...","subtitulo":"...","resumen":"...","html":"<!doctype html>...","filename":"pagina.html"}. Usa Tailwind CSS por CDN, iconos Lucide, hero #1E4D2B con acentos lima/dorado, búsqueda en vivo, filtros por categoría, tarjetas con precio CLP, stock, dosis y botón "Cotizar por WhatsApp". Nada de tablas HTML antiguas.
- Código en cualquier lenguaje: {"accion":"crear_codigo","titulo":"...","lenguaje":"python","codigo":"...","filename":"programa.py"}
- Documento genérico: {"accion":"crear_documento","documento":{"tipo":"pdf|word|ppt","titulo":"...","contenido":"# Título\\n## Sección\\n- Punto"}}
- Guardar productos en el catálogo: {"accion":"registrar_productos_en_app","productos":[{...}],"reemplazar":false}
- Simular ventas: {"accion":"simular_ventas","periodoDias":30,"crecimientoMensual":0}
- Guardar un recuerdo persistente: {"accion":"guardar_memoria","categoria":"Negocio|Preferencias|Instrucciones|Datos personales|Clientes|Productos|Finanzas|Otros","dato":"..."}
- Olvidar o actualizar un recuerdo: {"accion":"olvidar_memoria","clave_o_id":"...","nuevo_dato":"..."}

MEMORIA PERSISTENTE:
- Usa los recuerdos disponibles de forma natural, como si ya conocieras al usuario. No digas "según mi base de datos", "como te dije antes" ni anuncies que consultas memoria.
- Si el usuario dice "recuerda que...", "mi nombre es...", "siempre aplica...", "ten en cuenta...", guarda el dato con la acción guardar_memoria.
- Si dice que un dato cambió o ya no aplica, usa olvidar_memoria (con nuevo_dato para actualizarlo).
- No guardes contraseñas, claves, tokens ni datos de tarjetas.
- No contradigas un recuerdo vigente.

REGLAS DE TRABAJO:
- No uses "-" como dato. Si falta un dato, deja el campo vacío o escribe "Pendiente".
- Nunca escribas enlaces internos como [texto](sandbox:/archivo) ni fórmulas LaTeX con "$". Escribe "$2~a~3~L/ha$" como "2 a 3 L/ha" y "\$0" como "$0": la app muestra el texto limpio y una tarjeta de descarga.
- En documentos usa tablas Markdown con "|" cuando haya datos comparables; el motor las convierte en tablas reales con bordes proporcionales.
- Para catálogos, el motor calcula los totales reales (suma de productos y valor de inventario Precio × Stock) desde la app: no escribas filas de totales manuales.
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
    body: JSON.stringify({ models: modelList, messages, temperature: 0.7, max_tokens: Math.min(Math.max(Number(maxTokens) || 1500, 100), 6000) })
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
      generationConfig: { temperature: 0.7, maxOutputTokens: Math.min(Math.max(Number(maxTokens) || 1500, 100), 8000) }
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
