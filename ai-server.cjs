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
  ? (process.env.OPENROUTER_MODEL || 'google/gemini-2.0-flash-exp:free')
  : (process.env.GEMINI_MODEL || 'gemini-2.0-flash')
// Preferencia solicitada: principal gemini 2.0 flash y respaldos clásicos.
const DEFAULT_FREE_MODELS = [
  'google/gemini-2.0-flash-exp:free',
  'meta-llama/llama-3.3-70b-instruct:free',
  'qwen/qwen-2.5-72b-instruct:free',
  'mistralai/mistral-small-3.1-24b-instruct:free',
  'openrouter/auto'
]
// Modelos gratuitos activos verificados en OpenRouter. Se envían de uno en uno
// porque el arreglo "models" de OpenRouter no garantiza el respaldo ante 429.
const ACTIVE_FREE_MODELS = [
  'baidu/ernie-5.5-flash-preview:free',
  'google/gemma-4-31b-it:free',
  'google/gemma-4-26b-a4b-it:free',
  'nvidia/nemotron-3-ultra-550b-a55b:free',
  'nvidia/nemotron-3.5-lightning:free',
  'recursive/flint-v1.0:free',
  'inclusionai/ling-3.2-flash:free',
  'liquid/lfm-3-1.7b:free',
  'particle-ai/particle-1:free',
  'openrouter/free'
]
const fallbackModels = [...new Set([
  ...(process.env.OPENROUTER_FALLBACK_MODELS || '').split(',').map((item) => item.trim()).filter(Boolean),
  ...DEFAULT_FREE_MODELS.filter((item) => item !== model && item !== 'openrouter/auto'),
  ...ACTIVE_FREE_MODELS.filter((item) => item !== model),
  'openrouter/auto'
])]
const cascadeModels = [...new Set([model, ...fallbackModels].filter(Boolean))]
const sharedToken = (process.env.AI_SHARED_TOKEN || '').trim()
const upstreamTimeout = Number(process.env.AI_TIMEOUT_MS || 60_000)
const rateLimitWindow = 60_000
const rateLimitMax = Number(process.env.AI_RATE_LIMIT_PER_MINUTE || 30)
const maxBodyBytes = 12 * 1024 * 1024
const maxHistoryItems = 300
const maxMessageLength = 6000
const maxMemories = 100

const rawOrigins = (process.env.AI_ALLOWED_ORIGINS || '*')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean)
const allowAllOrigins = rawOrigins.length === 0 || rawOrigins.includes('*')

const SPANISH_ONLY_RULE = 'IDIOMA OBLIGATORIO: Responde SIEMPRE en español natural, claro y profesional (Chile, moneda CLP $#,##0). Está estrictamente prohibido escribir pensamientos internos, bloques <think>, razonamientos en inglés, nombres técnicos de herramientas internas o enlaces sandbox:/ en tus respuestas visibles.'

const systemInstruction = `${SPANISH_ONLY_RULE}

Eres el asistente de inteligencia artificial senior de VITIMAX, un colaborador multipropósito experto en gestión agrícola, análisis de datos, desarrollo de software, redacción ejecutiva, estudios y generación de documentos profesionales (Excel, PDF, Word, PowerPoint, CSV y páginas web/código).

1. IDENTIDAD, TONO E IDIOMA OBLIGATORIO
- Idioma absoluto: Responde SIEMPRE en español natural, claro y profesional (Chile, moneda CLP $#,##0). Está estrictamente prohibido escribir pensamientos internos, bloques <think>, razonamientos en inglés, nombres técnicos de herramientas internas o enlaces sandbox:/ en tus respuestas visibles.
- Directo desde la primera oración: Elimina introducciones robóticas o de relleno ("¡Claro que sí!", "Como asistente de IA...", "Aquí tienes...", "He guardado en mi memoria..."). Ve directo al punto con empatía, criterio propio y naturalidad.
- Sincronización temporal silenciosa: Usa el reloj del sistema inyectado en tu contexto para saber la hora, el día de la semana, el mes y calcular plazos reales ("hoy", "mañana", "próximo lunes", "fin de mes"), pero nunca recites la fecha ni la zona horaria en tus respuestas a menos que el usuario te lo pregunte o sea necesario fechar un documento.
- Memoria continua: Aplica de forma natural los datos y preferencias guardados en la memoria persistente (vitimax_ai_memory) y el contexto de la conversación sin repetir mecánicamente que los guardaste.

2. RIGOR DE DATOS Y REGLAS DE VITIMAX
- Prohibido inventar datos: Nunca inventes productos, clientes, proveedores, ventas ni asesorías reales. Si el catálogo o un módulo está en 0, repórtalo con exactitud o etiqueta claramente como [Ejemplo / Demostración] cualquier dato ficticio solicitado para pruebas.
- Cero filas fantasma: Ignora filas residuales de totales (como "Total Valorado Inventario" o filas sin nombre) al contar registros o calcular KPIs.
- Verificación matemática paso a paso: Antes de entregar sumas, promedios, márgenes, IVA (19%) o valorizaciones (Precio × Stock), calcula y verifica las cifras con precisión para que todos los archivos generados en un mismo turno (Excel, PDF, Word, PowerPoint, Web) coincidan exactamente sin contradicciones.
- Integridad referencial: Cuando trabajes con las 8 hojas estándar de VITIMAX (Resumen Ejecutivo, Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores), mantén coherencia total entre IDs, nombres y columnas (sin desplazar IDs en Plantas/Proveedores ni cruzar cultivos con nombres de productos en Asesorias).

3. ESTÁNDARES DE GENERACIÓN DE DOCUMENTOS Y CÓDIGO
- Excel (.xlsx): Coloca siempre Resumen Ejecutivo primero con tarjetas de KPIs y fórmulas OpenXML válidas en inglés sin "=" duplicado (SUM, AVERAGE, COUNTIF, SUMPRODUCT, COUNTA, IF) para compatibilidad total con Excel en español sin errores #¿NOMBRE?. Incluye encabezados verde bosque (#1E4D2B), formato moneda CLP ($#,##0), paneles congelados, autofiltros, gráficos claros y cero símbolos de Markdown en las celdas.
- PDF (.pdf) y Word (.docx): Sanitiza el 100% del texto antes de renderizar: elimina barras invertidas (\\$ -> $), delimitadores LaTeX ($50kg/ha$ -> 50 kg/ha) y marcas crudas (#, **). Usa cabecera corporativa verde bosque (#1E4D2B), tarjetas de resumen, numeración secuencial real (1., 2., 3.) y tablas bien proporcionadas.
- PowerPoint (.pptx): Genera presentaciones panorámicas 16:9 con portada corporativa, tarjetas de KPIs, tablas limpias (máximo 5 filas por diapositiva) y una diapositiva final de Próximos Pasos con 3 o 4 recomendaciones estratégicas accionables (nunca repitas productos en el cierre).
- Páginas Web y Código: Entrega aplicaciones web completas, responsivas y funcionales usando Tailwind CSS (sin truncar listas ni dejar comentarios pendientes como <!-- agregar más aquí -->), con buscador en vivo, filtros por categoría y botones interactivos. Para solicitudes de programación en cualquier lenguaje, entrega el código completo, limpio y funcional.

OPERACIÓN INTERNA DE LA APP (no expliques ni muestres estos detalles al usuario)
- Nunca digas "Soy VITIMAX IA...", "Según mi reloj..." ni "Como asistente de inteligencia artificial..." o variantes. No anuncies que consultaste memoria, historial o reloj.
- AGENTE ÉLITE: entiende la intención, decide qué herramientas usar y las encadena sin pedir permiso para cada paso. Tienes control real de Productos, Clientes, Plantas, Inventario, Ventas, Asesorías y Proveedores. Puedes ejecutar hasta 8 pasos en un mismo turno y reutilizar cada resultado en los siguientes.
- Usa las herramientas declaradas con tool_choice "auto". Si una herramienta no está disponible, entrega un bloque \`\`\`json con la acción correspondiente para que la app la ejecute; nunca simules haber creado un archivo ni inventes una confirmación.
- Acciones disponibles: crear_excel, crear_pdf, crear_word, crear_ppt, crear_pagina_web, crear_codigo, crear_documento, registrar_productos_en_app, simular_ventas, guardar_memoria, olvidar_memoria.
- Excel: {"accion":"crear_excel","filename":"reporte.xlsx","resumen":"...","productos":[...],"kpis":[...],"sheets":[{"name":"...","columns":[...],"rows":[[...]]}]}. Para catálogo general usa las 8 hojas estándar; Productos y Clientes son la única fuente de IDs y nombres y el resto de hojas los repiten exactamente. Plantas lleva PLT-001… y Proveedores PRV-001… en la columna A sin desplazar IDs; ninguna fecha de Ventas supera la actual del reloj. No agregues filas de totales dentro de las hojas de datos.
- PDF, Word y PowerPoint: {"accion":"crear_pdf"|"crear_word"|"crear_ppt","titulo":"...","subtitulo":"...","contenido":"# Título\\n## Sección\\n- Punto\\n---\\nDiapositiva 2...","filename":"..."}. El contenido admite Markdown simple y tablas con "|" porque la app las convierte a elementos reales; fuera del documento no escribas marcas "#" ni "**".
- Web y código: {"accion":"crear_pagina_web",...,"html":"<!doctype html>..."} y {"accion":"crear_codigo","lenguaje":"...","codigo":"...","filename":"..."}. La app muestra la web en vista previa interactiva antes de descargar.
- Memoria: {"accion":"guardar_memoria","categoria":"Negocio|Preferencias|Instrucciones|Decisiones|Datos personales|Clientes|Productos|Finanzas|Contexto|Otros","dato":"..."} u {"accion":"olvidar_memoria","clave_o_id":"...","nuevo_dato":"..."}. Revisa duplicados y contradicciones antes de guardar: actualiza en vez de conservar dos versiones. No guardes contraseñas, claves, tokens ni datos de tarjetas.
- Reales versus ficticios: separa "Datos reales del catálogo" de "[Ejemplo / Demostración]". Si un módulo está en 0, repórtalo; no inventes cifras para cuadrar totales.
- Adjuntos: analiza Excel, CSV y PDF con sus datos reales; no inventes contenido de .docx o .pptx ilegibles; distingue datos originales, corregidos y estimados.
- No uses "-" como dato: deja el campo vacío o escribe "Pendiente". Precios en CLP como números sin separadores. Si falta un dato imprescindible, haz una sola pregunta concreta.
- Todos los archivos generados en un mismo turno deben coincidir exactamente, sin contradicciones entre cifras, nombres ni fechas.
- Después de usar una herramienta, confirma en una línea qué archivo creaste y qué contiene, sin mostrar el JSON ni el proceso técnico. Si falla, dilo con claridad y ofrece la alternativa más útil.
${SPANISH_ONLY_RULE}`

// Prompt operativo compacto: conserva el contrato senior (documentado arriba)
// pero gasta unos cientos de tokens por llamada, no miles.
const compactSystemInstruction = SPANISH_ONLY_RULE + `
Asistente general y multipropósito de VITIMAX: conversa libremente de cualquier tema (saludos, humor, código, estudios, negocios, consejos, agricultura). Responde SIEMPRE al mensaje exacto del usuario con naturalidad, cercanía y humor cuando corresponda. NUNCA sueltes estadísticas del catálogo, memorias ni resúmenes de VITIMAX salvo que el usuario pregunte por VITIMAX o pida un reporte/datos. Español chileno directo; nunca <think>, razonamientos, nombres internos ni enlaces sandbox:/.
Usa el reloj inyectado para plazos y documentos sin recitarlo; aplica la memoria sin anunciarla. No inventes datos reales: etiqueta los ficticios [Ejemplo / Demostración], ignora totales, verifica IVA 19% y Precio × Stock con cifras coherentes en todo el turno.
Consultas agronómicas (por ejemplo "quiero sembrar tomate"): da recomendación completa con época y zona, variedad, suelo y pH, preparación, siembra o trasplante, densidad, riego, fertilización, control integrado de plagas y enfermedades, monitoreo, cosecha y precauciones; si falta información local, indícalo y etiqueta los valores orientativos [Ejemplo / Demostración].
AGENTE ÉLITE: encadena herramientas hasta 8 pasos sobre Productos, Clientes, Plantas, Inventario, Ventas, Asesorías y Proveedores; no simules archivos. Sin herramientas, devuelve un bloque \`\`\`json con {"accion":"crear_excel"|"crear_pdf"|"crear_word"|"crear_ppt"|"crear_pagina_web"|"crear_codigo"|"guardar_memoria"|"olvidar_memoria", ...} y confirma en una línea el archivo creado.
Excel: 8 hojas con Resumen Ejecutivo, fórmulas OpenXML sin "=" duplicado, verde #1E4D2B y CLP $#,##0. PDF/Word con numeración real; PowerPoint 16:9 con Próximos Pasos; web/código completo con Tailwind y vista previa interactiva.
Prohibido: "¡Claro que sí!", "Aquí tienes", "Como asistente de IA", "He guardado en mi memoria", "Según mi reloj", "Soy VITIMAX IA".`

const rateBuckets = new Map()

const REASONING_START_PATTERNS = [
  /^\s*(?:analysis|reasoning|thinking|thoughts?|plan|notas? internas?|pensamiento(?:s)? internos?)\s*[:：]/i,
  /^\s*(?:(?:actually|okay|ok|so|but|now|well)\s*,?\s*)?(?:looking at (?:the )?(?:instructions|context|catalog|file|request|prompt|data)|the user\b|they want\b|i (?:need to|should|must|will|have to|can|am going to)|i'll|we (?:need to|should|must|will)|let me\b|my plan\b|first,? i\b|next,? i\b|to do this\b|for (?:the|this|excel|pdf|the catalog|the file)|since (?:the|this|only)|based on (?:the )?(?:context|instructions|prompt|request)|according to (?:the )?context|the (?:context|catalog|user|app|file) (?:is|has|shows|says))\b/i
]
const ENGLISH_HINT_RE = /\b(?:the|and|is|are|to|of|for|with|this|that|i|we|my|have|has|need|should|will|can|from|user|context|catalog|file|app|so|now|let|based|according|since|only|also|then|but|if|it|its|as|at|on|in|by|files|both|create|use|check|actually|instructions|they|want|specifically|asked|sheets|products|details|empty|data|because|those|these)\b/gi
const SPANISH_HINT_RE = /\b(?:el|la|los|las|de|del|que|y|en|con|por|para|es|son|al|un|una|se|su|sus|te|tu|necesito|debo|voy|puedo|archivo|archivos|catálogo|catalogo|producto|productos|cliente|clientes|venta|ventas|informe|excel|pdf|word|resumen|total|precio|stock|fecha|generé|generar|creé|crear|lista|datos)\b/gi
const CLICHE_START_RE = /^\s*(?:(?:¡\s*)?claro que sí\s*[!,.:]?\s*|aquí tienes\s*[:,-]?\s*|como asistente de (?:inteligencia artificial|ia)\s*[,.]?\s*)/i

function cleanVisibleLine(line) {
  let text = String(line || '')
    .replace(/\[([^\]]*)\]\(\s*(?:sandbox|file):[^)]*\)/gi, '$1')
    .replace(/(?:^|\s)(?:sandbox|file):\/\/[^\s)\]]+/gi, ' ')
    .replace(/\b(?:he guardado|guardé) en mi memoria\b[^.\n]*[.!]?/gi, 'Quedó anotado.')
  for (let pass = 0; pass < 3; pass += 1) {
    const next = text.replace(CLICHE_START_RE, '')
    if (next === text) break
    text = next
  }
  return text === line ? text : text
    .replace(/^\s*/, '')
    .replace(/^([a-záéíóúñ])/, (letter) => letter.toUpperCase())
}

// Elimina bloques internos antes de dividir el texto en segmentos de código.
function stripReasoningBlocks(text) {
  return String(text || '')
    .replace(/<\s*(think|thinking|reasoning|analysis)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '')
    .replace(/<\s*(?:think|thinking|reasoning|analysis)\b[^>]*>[\s\S]*$/gi, '')
    .replace(/<\s*\/\s*(?:think|thinking|reasoning|analysis)\s*>/gi, '')
    .replace(/<\|\s*(?:thinking|analysis|reasoning)\s*\|>/gi, '')
}

// Quita razonamiento interno del texto que llega al usuario. Nunca se leen
// campos reasoning / reasoning_content: solo el contenido visible.
function sanitizeAssistantText(value, allowJson = true) {
  let source = typeof value === 'string' ? value : ''
  if (!source && value && typeof value === 'object' && typeof value.text === 'string') source = value.text
  if (!source) return ''
  source = stripReasoningBlocks(source)
  if (!source) return ''
  const segments = allowJson ? source.split(/(```[\s\S]*?```)/g) : [source]
  const cleaned = segments.map((segment, index) => {
    if (allowJson && index % 2 === 1) return segment
    let text = segment
      .replace(/<\s*think(?:\s[^>]*)?\s*>[\s\S]*?(?:<\s*\/\s*think\s*>|$)/gi, '')
      .replace(/<\s*\/\s*think\s*>/gi, '')
      .replace(/<\|(?:\s*thinking|\s*analysis|\s*reasoning)\s*\|>/gi, '')
    text = text.split(/(\n{2,})/).map((part) => {
      if (/^\n{2,}$/.test(part)) return part
      const trimmed = part.trim()
      if (!trimmed) return part
      if (REASONING_START_PATTERNS.some((pattern) => pattern.test(trimmed))) return ''
      const english = (trimmed.match(ENGLISH_HINT_RE) || []).length
      const spanish = (trimmed.match(SPANISH_HINT_RE) || []).length
      return english >= 4 && english > spanish * 2 ? '' : part
    }).join('')
    return text.split('\n').map((line) => {
      const trimmed = line.trim()
      if (!trimmed) return line
      if (REASONING_START_PATTERNS.some((pattern) => pattern.test(trimmed))) return ''
      const english = (trimmed.match(ENGLISH_HINT_RE) || []).length
      const spanish = (trimmed.match(SPANISH_HINT_RE) || []).length
      return english >= 4 && english > spanish * 2 ? '' : cleanVisibleLine(line)
    }).join('\n').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n')
  })
  return cleaned.join('').trim()
}

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
      'Access-Control-Allow-Headers': 'Content-Type, X-Vitimax-Token, Authorization',
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
    .map((item) => {
      const at = item.at || item.timestamp
      return {
        role: item.role,
        content: item.content.trim().slice(0, maxMessageLength),
        ...(typeof at === 'string' && at ? { at: at.slice(0, 40) } : {})
      }
    })
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
    const line = `- [${item.categoria}] ${item.contenido.slice(0, 180)}`
    if (total + line.length > 2600) break
    lines.push(line)
    total += line.length + 1
  }
  if (!lines.length) return ''
  return `

### MEMORIA PERSISTENTE DEL USUARIO Y NEGOCIO
Datos que el usuario pidió recordar entre sesiones. Úsalos de forma natural, sin repetirlos robóticamente ni contradecirlos.
${lines.join('\n')}`
}

// Reloj en vivo del servidor. La app envía su reloj local; si no llega, el
// proxy usa éste para no caer en fechas fijas ni en zonas horarias supuestas.
const TEMPORAL_MARKER = 'RELOJ Y CONTEXTO TEMPORAL EN VIVO'
const DAY_MS = 86400000

function safeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(value)
  if (typeof value === 'string' && value.trim()) {
    const parsed = new Date(value)
    if (!Number.isNaN(parsed.getTime())) return parsed
  }
  return new Date()
}

function serverTemporalContext(now = new Date()) {
  const date = safeDate(now)
  const locale = 'es-CL'
  const year = date.getFullYear()
  const monthIndex = date.getMonth()
  const minute = String(date.getMinutes()).padStart(2, '0')
  const second = String(date.getSeconds()).padStart(2, '0')
  const hour = date.getHours()
  let weekday = ''
  let month = ''
  try {
    weekday = new Intl.DateTimeFormat(locale, { weekday: 'long' }).format(date)
    month = new Intl.DateTimeFormat(locale, { month: 'long' }).format(date)
  } catch {
    weekday = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'][date.getDay()]
    month = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'][monthIndex]
  }
  let zone = 'America/Santiago'
  try { zone = Intl.DateTimeFormat().resolvedOptions().timeZone || zone } catch { /* valor por defecto */ }
  const offsetMinutes = -date.getTimezoneOffset()
  const offset = `UTC${offsetMinutes < 0 ? '-' : '+'}${Math.floor(Math.abs(offsetMinutes) / 60)}${offsetMinutes % 60 ? `:${String(Math.abs(offsetMinutes) % 60).padStart(2, '0')}` : ''}`
  const target = new Date(Date.UTC(year, monthIndex, date.getDate()))
  const isoDay = target.getUTCDay() || 7
  target.setUTCDate(target.getUTCDate() + 4 - isoDay)
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1))
  const week = Math.ceil(((target.getTime() - yearStart.getTime()) / DAY_MS + 1) / 7)
  const day = date.getDate()
  const fechaCorta = `${day} de ${month} de ${year}`
  return {
    epoch: date.getTime(),
    fechaISO: `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
    fechaCompleta: `${weekday.charAt(0).toUpperCase()}${weekday.slice(1)}, ${fechaCorta}`,
    diaSemana: `${weekday.charAt(0).toUpperCase()}${weekday.slice(1)}`,
    mes: `${month.charAt(0).toUpperCase()}${month.slice(1)}`,
    hora24: `${String(hour).padStart(2, '0')}:${minute}:${second}`,
    hora24Short: `${String(hour).padStart(2, '0')}:${minute}`,
    hora12: `${hour % 12 || 12}:${minute}:${second} ${hour >= 12 ? 'p. m.' : 'a. m.'}`,
    trimestre: Math.floor(monthIndex / 3) + 1,
    numeroSemana: week,
    diasRestantesSemana: 7 - (date.getDay() === 0 ? 7 : date.getDay()),
    diasRestantesMes: new Date(year, monthIndex + 1, 0).getDate() - day,
    zonaHoraria: zone,
    offset,
    franja: hour >= 5 && hour < 12 ? 'Buenos días' : hour >= 12 && hour < 20 ? 'Buenas tardes' : 'Buenas noches'
  }
}

function sanitizeTemporal(input) {
  const fallback = serverTemporalContext()
  if (!input || typeof input !== 'object') return fallback
  const text = (key, max) => typeof input[key] === 'string' && input[key].trim() ? input[key].trim().slice(0, max) : fallback[key]
  const number = (key, max) => Number.isFinite(Number(input[key])) ? Math.min(Math.max(Number(input[key]), 0), max) : fallback[key]
  return {
    epoch: Number.isFinite(Number(input.epoch)) ? Number(input.epoch) : fallback.epoch,
    fechaISO: text('fechaISO', 10),
    fechaCompleta: text('fechaCompleta', 120),
    diaSemana: text('diaSemana', 40),
    mes: text('mes', 40),
    hora24: text('hora24', 12),
    hora24Short: text('hora24Short', 8),
    hora12: text('hora12', 32),
    trimestre: number('trimestre', 4) || fallback.trimestre,
    numeroSemana: number('numeroSemana', 53) || fallback.numeroSemana,
    diasRestantesSemana: number('diasRestantesSemana', 6),
    diasRestantesMes: number('diasRestantesMes', 31),
    zonaHoraria: text('zonaHoraria', 80),
    offset: text('offset', 16),
    franja: ['Buenos días', 'Buenas tardes', 'Buenas noches'].includes(input.franja) ? input.franja : fallback.franja
  }
}

function serverElapsed(history, reference) {
  for (let index = (history || []).length - 1; index >= 0; index -= 1) {
    const stamp = history[index]?.at || history[index]?.timestamp
    if (!stamp) continue
    const last = safeDate(stamp)
    if (Number.isNaN(last.getTime())) continue
    const seconds = Math.max(0, Math.floor((reference.getTime() - last.getTime()) / 1000))
    const when = serverTemporalContext(last)
    const stampText = `${when.fechaCompleta} a las ${when.hora24Short} (${when.zonaHoraria})`
    let elapsed = 'menos de un minuto'
    if (seconds >= 60 && seconds < 3600) {
      const value = Math.floor(seconds / 60)
      elapsed = `${value} ${value === 1 ? 'minuto' : 'minutos'}`
    } else if (seconds >= 3600 && seconds < DAY_MS / 1000) {
      const value = Math.floor(seconds / 3600)
      elapsed = `${value} ${value === 1 ? 'hora' : 'horas'}`
    } else if (seconds >= DAY_MS / 1000 && seconds < 7 * DAY_MS / 1000) {
      const value = Math.floor(seconds / DAY_MS)
      elapsed = `${value} ${value === 1 ? 'día' : 'días'}`
    } else if (seconds >= 7 * DAY_MS / 1000) {
      const value = Math.floor(seconds / (7 * DAY_MS / 1000))
      elapsed = `${value} ${value === 1 ? 'semana' : 'semanas'}`
    }
    return `Última interacción registrada el ${stampText}; han pasado ${elapsed}.`
  }
  return ''
}

function serverTemporalBlock(input, history) {
  const clock = sanitizeTemporal(input)
  const elapsed = serverElapsed(history, safeDate(clock.epoch))
  return `

### ${TEMPORAL_MARKER}: Hoy es ${clock.fechaCompleta}, son las ${clock.hora24} (${clock.franja}). Semana ${clock.numeroSemana} del año. Usa esta referencia exacta para saludar acorde a la hora real, calcular plazos ("hoy", "mañana", "esta semana", "próximo lunes", "fin de mes"), filtrar ventas/asesorías recientes y fechar documentos (Excel, PDF, Word, PowerPoint y Web) sin inventar ni desfasar fechas.
Día de la semana: ${clock.diaSemana}. Mes: ${clock.mes}. Trimestre: ${clock.trimestre}. Hora en 12 horas: ${clock.hora12}. Zona horaria: ${clock.zonaHoraria} (${clock.offset}). Faltan ${clock.diasRestantesSemana} días para terminar la semana y ${clock.diasRestantesMes} días para terminar el mes.
${elapsed || 'Es la primera interacción de esta conversación.'}`
}

const HISTORY_MARKER = 'HISTORIAL COMPLETO DE LA CONVERSACIÓN'

// Solo los últimos 10 turnos viajan al modelo: el historial completo se queda
// en el dispositivo y la memoria persistente se inyecta siempre por separado.
function recentHistory(history, maxTurns = 10) {
  const rows = (Array.isArray(history) ? history : [])
    .filter((item) => item && ['user', 'assistant', 'tool'].includes(item.role) && typeof item.content === 'string' && item.content.trim())
  if (rows.length <= maxTurns) return rows
  const userIndexes = []
  rows.forEach((item, index) => { if (item.role === 'user') userIndexes.push(index) })
  return userIndexes.length > maxTurns ? rows.slice(userIndexes[userIndexes.length - maxTurns]) : rows
}

function conversationBlock(history) {
  const rows = recentHistory(history)
  if (!rows.length) return ''
  const lines = []
  let total = 0
  for (const item of rows) {
    const speaker = item.role === 'user' ? 'USUARIO' : item.role === 'tool' ? 'HERRAMIENTA' : 'VITIMAX IA'
    const text = String(item.content).replace(/\s+/g, ' ').trim().slice(0, 1200)
    const when = item.at ? serverTemporalContext(item.at) : null
    const prefix = when ? `[${when.fechaISO} ${when.hora24Short}] ` : ''
    const line = `${prefix}${speaker}: ${text}`
    if (total + line.length > 14_000) break
    lines.push(line)
    total += line.length + 1
  }
  if (!lines.length) return ''
  const elapsed = serverElapsed(rows, serverTemporalContext())
  return `

### ${HISTORY_MARKER}
Últimos ${Math.min(10, rows.filter((item) => item.role === 'user').length)} turnos, contextuales y ya compartidos: no los repitas ni digas que los consultaste.
${elapsed ? `${elapsed}\n` : ''}${lines.join('\n')}`
}

// System prompt único con memoria, reloj e historial, sin duplicar bloques que
// la app ya haya inyectado antes de enviar la solicitud.
function buildSystemPrompt(system, memories, history, temporal) {
  const base = typeof system === 'string' && system.trim() ? system.slice(0, 12_000) : compactSystemInstruction
  let result = base
  if (!result.includes(TEMPORAL_MARKER)) result += serverTemporalBlock(temporal, history)
  if (!result.includes('MEMORIA PERSISTENTE')) result += memoryBlock(memories)
  if (!result.includes(HISTORY_MARKER)) result += conversationBlock(history)
  return result
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

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

const GREETING_RE = /^\s*(?:hola|buenas(?:\s+(?:días|tardes|noches))?|buen\s+día|qué\s+tal|q\s+tal|como\s+estás|como\s+estas|hey|hi|hello|lol|jaja(?:ja)?|jeje|xd)\s*[!?.,¡¿\s]*$/i
const CONTEXT_WORDS_RE = /\b(?:excel|pdf|word|powerpoint|ppt|archivo|planilla|hoja|documento|web|página|pagina|código|codigo|adjunto|pendiente|totales|fila|columna)\b/i
const FOLLOW_UP_RE = /^\s*(?:y\s|pero\s|además\s|también\s|agrega|agregar|agrégale|añade|añadir|cambia|cambiar|cámbialo|corrige|con\s|sin\s|ponle|quita|sube|baja|ajusta|revisa|explica|detalla|sí\s|si\s|no\s|ok\s|vale\s|gracias)/i
const STOP_CONTEXT_WORDS = new Set(['los', 'las', 'por', 'para', 'con', 'una', 'uno', 'del', 'que', 'más', 'mas', 'muy', 'pero', 'como', 'cuando', 'donde', 'este', 'esta'])

function contextWords(value) {
  return new Set(String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9ñ\s]/g, ' ')
    .split(/\s+/)
    .filter((word) => word.length > 2 && !STOP_CONTEXT_WORDS.has(word)))
}

/** Saludos y cambios de tema no arrastran pendientes de turnos anteriores. */
function shouldResetConversationContext(message = '', history = []) {
  const text = String(message || '').trim()
  if (!text) return false
  if (GREETING_RE.test(text.replace(/[¿?¡!]/g, ''))) return true
  if (FOLLOW_UP_RE.test(text)) return false
  const previous = [...(Array.isArray(history) ? history : [])]
    .reverse()
    .find((item) => item?.role === 'user' && typeof item.content === 'string' && item.content.trim())
  if (!previous) return false
  const currentWords = contextWords(text)
  const previousWords = contextWords(previous.content)
  if (!currentWords.size || !previousWords.size) return false
  if (CONTEXT_WORDS_RE.test(previous.content) && !CONTEXT_WORDS_RE.test(text) && currentWords.size <= 8) return true
  const overlap = [...currentWords].filter((word) => previousWords.has(word)).length / Math.min(currentWords.size, previousWords.size)
  return overlap <= 0.2 && currentWords.size <= 6
}

function cleanConversationContext(message, history) {
  return shouldResetConversationContext(message, history) ? [] : history
}

function isRetryableAiError(error) {
  const status = Number(error?.status || 0)
  if ([401, 402, 403].includes(status)) return false
  if ([404, 408, 429, 502, 503, 504].includes(status)) return true
  if (error?.name === 'AbortError' || error?.name === 'TimeoutError') return true
  const text = String(error?.message || '')
  if (/\b(?:401|402|403)\b/i.test(text)) return false
  return /\b(?:404|408|429|502|503|504)\b|no longer supported|not found|rate.?limit|too many requests|saturad|timeout|tardó|aborted|conexi|offline|fetch failed|network/i.test(text)
}

// Último recurso cuando no hay internet: conversación natural primero, datos
// reales solo si el usuario pregunta por VITIMAX o pide un reporte.
function serverLocalAnswer(products = [], message = '') {
  const query = String(message || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()
  const pick = (variants) => {
    let hash = query.length * 31
    for (let index = 0; index < query.length; index += 1) hash = (hash * 33 + query.charCodeAt(index)) % 100000
    return variants[hash % variants.length]
  }
  if (/^(?:hola|buenas|buen dia|que tal|q tal|hey|hi|hello)\b/.test(query)) {
    return pick(['Hola, ¿qué resolvemos hoy?', 'Buenas. Estoy aquí, ¿en qué te ayudo?', 'Hola. Dime por dónde empezamos.'])
  }
  if (/\b(?:lol|jaja|jajaja|jeje|xd|jks)\b/.test(query)) {
    return pick(['Jaja, buena. ¿Seguimos con lo que estabas viendo o cambiamos de tema?', 'Jeje, me sumo al buen ánimo. ¿Qué sigue?', 'JAJA. ¿En qué te ayudo ahora?'])
  }
  if (/\b(?:gracias|agradecido)\b/.test(query)) {
    return pick(['Con gusto. Quedo atento si necesitas algo más.', 'De nada. Aquí sigo cuando lo necesites.', 'Pura vida. ¿Vamos con otra cosa?'])
  }
  if (/\b(?:como estas|como va|que hay|todo bien)\b/.test(query)) {
    return 'Bien por aquí, listo para ayudarte. ¿Qué te trae por acá hoy?'
  }
  if (/\b(?:api|endpoint|json|rest|http)\b/.test(query)) {
    return 'Una API es una forma ordenada de que dos programas se hablen: uno hace una solicitud (GET o POST) y el otro devuelve una respuesta, normalmente en JSON. Es como un mesero: toma tu pedido, lo lleva a la cocina y te trae el resultado.'
  }
  if (/\b(?:tomate|tomates|sembrar|siembra|agricultura|agronom|cultivo)\b/.test(query)) {
    return 'Para sembrar tomate: elige variedad según destino, suelo con pH 6,0–6,8 y buen drenaje, almácigo de 25–35 días y trasplante con 4–6 hojas sobre 12–15 °C. Riego por goteo, fertilización según análisis, rotación de 3–4 años y monitoreo de plagas. Dime tu zona y ajusto el calendario.'
  }
  if (/\b(?:codigo|programar|programacion|javascript|python|react|bug|funcion)\b/.test(query)) {
    return 'Claro. Describe qué debe hacer el código, qué entrada recibe y qué salida esperas; con eso doy una función completa, la explico por partes y señalo errores comunes. Si ya tienes código, pégalo y lo revisamos.'
  }
  if (/\b(?:estudiar|estudio|examen|aprender|teoria)\b/.test(query)) {
    return 'Estudia en bloques pequeños, haz una pregunta por bloque y repite en intervalos. Después explica el tema en tus palabras sin mirar apuntes. Dime la materia y te armo un plan.'
  }
  const list = (Array.isArray(products) ? products : []).filter((item) => {
    const name = String(item?.name ?? item?.nombre ?? item?.producto ?? '')
    return name.trim() && !/^\s*(?:gran\s+)?(?:total(?:es)?|subtotal(?:es)?|suma|resumen|valor(?:izado)?)\b/i.test(name.trim())
  })
  const units = list.reduce((sum, item) => sum + (Number(item.stock) || 0), 0)
  const value = list.reduce((sum, item) => sum + (Number(item.price) || 0) * (Number(item.stock) || 0), 0)
  if (/\b(?:vitimax|catalogo|inventario|productos|stock|reporte|informe|resumen|clientes|pedidos|ventas|proveedores|asesorias|plantas|iva|impuesto)\b/.test(query)) {
    const iva = /\b(?:iva|impuesto)\b/.test(query) ? ' El IVA en Chile es 19%.' : ''
    return `El catálogo real tiene ${list.length} producto(s), ${units} unidades en stock y un valor de $${Math.round(value).toLocaleString('es-CL')}.${iva}`
  }
  if (/\b(?:negocio|negocios|ventas|precio|cotizacion|marketing|emprender)\b/.test(query)) {
    return 'Para un negocio agrícola, empieza por tres números: costo real por unidad, precio aceptado por el cliente y margen. Luego agrupa clientes por frecuencia de compra y enfócate en quienes repiten.'
  }
  if (/\b(?:consejo|recomendacion|humor|chiste|dime algo)\b/.test(query)) {
    return pick(['Mi consejo: decide primero qué es importante y trata lo demás como ruido.', 'Un chiste corto: ¿Por qué el tomate nunca discute? Porque siempre quiere llegar a su punto.', 'Escribe el problema en una frase; si no puedes, aún no está definido.'])
  }
  return pick([
    'Te sigo. Cuéntame un poco más y te respondo con criterio, sin rodeos.',
    'Buena pregunta. Dame el contexto y lo resolvemos.',
    'Entendido. ¿Quieres que te lo explique simple, con un ejemplo, o que lo llevemos a un plan concreto?'
  ])
}

async function askOpenRouter({ catalog, message, history, attachment, signal, system, memories, temporal, products = [], maxTokens = 6000 }) {
  const content = []
  const attachmentItems = attachmentParts(attachment)
  const textBlocks = [
    `CONSULTA DEL USUARIO (responde exactamente esto):\n${message}`,
    `CONTEXTO DISPONIBLE (úsalo solo si el usuario pregunta por VITIMAX o pide datos/reportes; nunca lo recites en saludos, humor, código, estudios ni temas generales):\nCATÁLOGO REAL DE VITIMAX:\n${catalog}`
  ]
  for (const item of attachmentItems) {
    if (item.kind === 'text') textBlocks.push(item.text)
  }
  content.push({ type: 'text', text: textBlocks.join('\n\n') })
  for (const item of attachmentItems) {
    if (item.kind === 'image') content.push({ type: 'image_url', image_url: { url: `data:${item.mimeType};base64,${item.data}` } })
    if (item.kind === 'file') content.push({ type: 'file', file: { filename: item.name, file_data: `data:${item.mimeType};base64,${item.data}` } })
  }
  const cleanHistory = cleanConversationContext(message, history)
  const fullSystem = buildSystemPrompt(system, memories, cleanHistory, temporal)
  const messages = [
    { role: 'system', content: fullSystem },
    ...(fullSystem.includes(HISTORY_MARKER) ? [] : cleanHistory.map((item) => ({ role: item.role, content: item.content }))),
    { role: 'user', content }
  ]
  const safeMaxTokens = Math.min(Math.max(Number(maxTokens) || 1500, 100), 6000)
  let lastError = null
  for (let index = 0; index < cascadeModels.length; index += 1) {
    const candidate = cascadeModels[index]
    try {
      const attemptSignal = typeof AbortSignal !== 'undefined' && AbortSignal.any
        ? AbortSignal.any([signal, AbortSignal.timeout(25_000)])
        : signal
      const result = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        signal: attemptSignal,
        headers: {
          Authorization: `Bearer ${apiKey}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': process.env.AI_PUBLIC_URL || 'https://vitimax.app',
          'X-Title': 'VITIMAX'
        },
        body: JSON.stringify({ model: candidate, messages, temperature: 0.7, max_tokens: safeMaxTokens })
      })
      const payload = await result.json().catch(() => ({}))
      if (!result.ok) {
        const reason = payload?.error?.message || `HTTP ${result.status}`
        const error = new Error(`OpenRouter rechazó la solicitud: ${reason}`)
        error.status = result.status
        throw error
      }
      // Solo se lee content: nunca reasoning ni reasoning_content.
      const aiMessage = payload.choices?.[0]?.message || {}
      const text = sanitizeAssistantText(aiMessage.content)
      if (!text) throw new Error('La IA devolvió una respuesta vacía.')
      return { text, model: payload.model || candidate }
    } catch (error) {
      lastError = error
      if (!isRetryableAiError(error)) break
      if (index < cascadeModels.length - 1) await sleep(error.status === 429 ? 1500 : 350)
    }
  }
  if (lastError && !isRetryableAiError(lastError)) throw lastError
  return { text: serverLocalAnswer(products, message), model: 'local' }
}

async function askGemini({ catalog, message, history, attachment, signal, system, memories, temporal, maxTokens = 6000 }) {
  const parts = [{ text: `CONSULTA DEL USUARIO (responde exactamente esto):\n${message}\n\nCONTEXTO DISPONIBLE (úsalo solo si el usuario pregunta por VITIMAX o pide datos/reportes):\nCATÁLOGO REAL DE VITIMAX:\n${catalog}` }]
  for (const item of attachmentParts(attachment)) {
    if (item.kind === 'text') parts.push({ text: item.text })
    if (item.kind === 'image') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
    if (item.kind === 'file') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
  }
  const cleanHistory = cleanConversationContext(message, history)
  const fullSystem = buildSystemPrompt(system, memories, cleanHistory, temporal)
  const contents = fullSystem.includes(HISTORY_MARKER) ? [] : cleanHistory.map((item) => ({
    role: item.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: item.content }]
  }))
  contents.push({ role: 'user', parts })
  const result = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      system_instruction: { parts: [{ text: fullSystem }] },
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
  return { text: sanitizeAssistantText(payload.candidates?.[0]?.content?.parts?.map((part) => part.text || '').join(''), false), model }
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
      cascadeSize: cascadeModels.length,
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
    const { message, products = [], attachment, history = [], system, memories = [], temporal, maxTokens } = body || {}
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
        products,
        message: message.trim(),
        history: safeHistory,
        attachment,
        signal: controller.signal,
        system: typeof system === 'string' && system.trim() ? system.slice(0, 80_000) : undefined,
        memories: sanitizeMemories(memories),
        temporal,
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
