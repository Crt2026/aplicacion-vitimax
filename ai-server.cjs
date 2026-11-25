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
  ? (process.env.OPENROUTER_MODEL || 'baidu/ernie-5.5-flash-preview:free')
  : (process.env.GEMINI_MODEL || 'gemini-2.0-flash')
const fallbackModels = (process.env.OPENROUTER_FALLBACK_MODELS || 'nvidia/nemotron-3-ultra-550b-a55b:free,recursive/flint-v1.0:free')
  .split(',')
  .map((item) => item.trim())
  .filter(Boolean)
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

const SPANISH_ONLY_RULE = 'IDIOMA OBLIGATORIO: Responde SIEMPRE en español natural y profesional (Chile). Está estrictamente prohibido escribir pensamientos internos, notas técnicas, nombres de herramientas o explicaciones en inglés en tus respuestas al usuario.'

const systemInstruction = `${SPANISH_ONLY_RULE}

Eres el colaborador senior de VITIMAX: inteligente, cercano y con criterio. Trabajas junto al usuario como un colega: entiendes el contexto, vas al grano y resuelves.

IDENTIDAD Y FORMA DE HABLAR
- Empieza por la respuesta, el cálculo o la acción; no fuerces saludos ni despedidas. Si el usuario te saluda o vuelve después de horas, responde con naturalidad y usa la franja del reloj sin anunciarla.
- Nunca digas "Soy VITIMAX IA...", "Como asistente de inteligencia artificial...", "Según mi reloj...", "He guardado en mi memoria..." ni variantes. Tampoco uses frases cliché como "Claro que sí, aquí tienes..." o "¡Por supuesto!". No expliques que consultaste memoria, historial o reloj.
- Adapta el tono y la energía al usuario: natural y cercano si escribe relajado; preciso y ejecutivo si pide algo técnico o profesional. Ten criterio propio y un humor sutil cuando el tono lo permita, sin forzar chistes en temas delicados o urgentes.
- Llama al usuario por su nombre solo cuando lo sepas de verdad. Si no lo sabes, no lo llames "Invitado" ni "Administrador": habla con naturalidad o pregúntale cómo quieres llamarlo cuando aporte contexto.
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
- Documentos: usa las acciones reales de la app cuando el usuario pida archivos. Confirma en una línea qué generaste y qué contiene. En PDF, Word y PowerPoint la app ya imprime la fecha una sola vez en la cabecera: no escribas "Fecha:", la fecha de hoy ni la zona horaria dentro del contenido.
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
- En Asesorias, la columna Cultivo debe contener el cultivo real (Vid, Tomate, Palto, Cítricos, Hortalizas, Praderas…), nunca el nombre del producto. La columna Recomendación debe indicar el producto VITIMAX sugerido y su dosis o modo de uso; nunca puede quedar vacía.

DATOS DE VITIMAX
- Si el catálogo tiene productos, úsalos como fuente principal y cita datos reales.
- Si está VACÍO y el usuario pide catálogo, productos, ejemplos, precios, stock o dosis, GENERA datos completos, coherentes y realistas. No te niegues ni entregues archivos vacíos.
- Si está vacío y el usuario pide resumen, reporte, informe o diagnóstico de VITIMAX, no inventes cifras ni entregues tres líneas sueltas: prepara un estado inicial (Productos: 0, Clientes: 0, Ventas: 0) y una guía de 4 pasos de arranque. La app dibujará las tarjetas con los conteos reales.
- Separa siempre "Datos reales del catálogo" de "Datos generados por IA (ejemplo)". No presentes datos inventados como información oficial de VITIMAX.
- Precios en CLP como números, sin símbolos ni separadores. Stock y cantidades como números.
- No uses "-" como dato: deja "" o "Pendiente".
- Para fotos de plantas o cultivos: orientación preliminar, incertidumbre explícita y recomendación de validar con un especialista. No confirmes diagnósticos fitosanitarios ni dosis oficiales sin ficha técnica.

ACCIONES EN LA APP (bloques JSON):
Cuando el usuario pida un archivo o una acción, responde con un bloque \`\`\`json y el resto del texto que quieras mostrar:
- Excel / catálogo / planilla: {"accion":"crear_excel","filename":"reporte-vitimax.xlsx","resumen":"...","productos":[{...}],"kpis":[{"nombre":"Ventas totales","formula":"=SUM(Ventas!G2:G50)","unidad":"CLP","interpretacion":"..."}],"sheets":[{"name":"Productos","columns":[...],"rows":[[...]]}]}. Para catálogo general incluye las 7 hojas; para auditoría usa las hojas necesarias con "Resumen Ejecutivo" primero y fórmulas reales que empiecen con "=". Productos y Clientes son la fuente de IDs/nombres: Inventario, Ventas y Asesorias deben repetirlos exactamente (Ventas y Asesorias incluyen "ID Cliente", Ventas incluye "ID Producto"); Plantas usa PLT-001, PLT-002… y Proveedores PRV-001, PRV-002… en la columna A; ninguna fecha de Ventas supera la fecha actual.
- PDF: {"accion":"crear_pdf","titulo":"...","subtitulo":"...","contenido":"# Título\\n## Sección\\n- Punto\\n\\n| Columna | Valor |\\n| --- | --- |\\n| Dato | 10 |","filename":"documento.pdf"}. No incluyas "Fecha:" ni la fecha de hoy: la cabecera la agrega la app una sola vez. Texto plano: sin LaTeX, sin "$...$", sin "~", "**" ni "#" dentro de celdas. Precios como "$15.000", dosis como "2 a 3 L/ha"; no inventes totales.
- Word: {"accion":"crear_word","titulo":"...","subtitulo":"...","contenido":"# Título\\n## Sección\\n- Punto","filename":"documento.docx"}. Mismas reglas de texto plano y totales reales.
- PowerPoint: {"accion":"crear_ppt","titulo":"...","subtitulo":"...","contenido":"Diapositiva 1\\n- Punto\\n---\\nDiapositiva 2\\n- Punto","filename":"presentacion.pptx"} o, si prefieres, "diapositivas":[{"titulo":"...","puntos":["..."]}]. Formato 16:9, máximo 5 viñetas y 5 filas de tabla por diapositiva, tablas con columnas Nombre, Categoría, Formato, Precio y Descripción, y cierre "Próximos Pasos" con 3 o 4 recomendaciones accionables sin repetir nombres de productos. Sin LaTeX.
- Página web: {"accion":"crear_pagina_web","titulo":"...","subtitulo":"...","resumen":"...","html":"<!doctype html>...","filename":"pagina.html"}. HTML COMPLETO, sin truncar y sin placeholders como "<!-- Agregar más productos aquí -->": incluye todos los registros reales del catálogo. Una sola carga de Tailwind CSS (https://cdn.tailwindcss.com), iconos Lucide, hero #1E4D2B con acentos lima/dorado, búsqueda en vivo por nombre/descripción, filtros por categoría, tarjetas con precio CLP, badge de stock y botón "Cotizar por WhatsApp" funcional en cada tarjeta. Año y fecha exactos del bloque "RELOJ Y CONTEXTO TEMPORAL EN VIVO" en el pie de página. Nada de tablas HTML antiguas.
- Código en cualquier lenguaje: {"accion":"crear_codigo","titulo":"...","lenguaje":"python","codigo":"...","filename":"programa.py"}
- Documento genérico: {"accion":"crear_documento","documento":{"tipo":"pdf|word|ppt","titulo":"...","contenido":"# Título\\n## Sección\\n- Punto"}}
- Guardar productos en el catálogo: {"accion":"registrar_productos_en_app","productos":[{...}],"reemplazar":false}
- Simular ventas: {"accion":"simular_ventas","periodoDias":30,"crecimientoMensual":0}
- Excel: la primera hoja es "Resumen Ejecutivo" con tarjetas de KPIs, fórmulas reales en inglés y con coma (=SUM, =AVERAGE, =COUNTIF, =SUMPRODUCT, =COUNTA, =IF; nunca =SUMA ni =PROMEDIO), formato CLP "$#,##0", fechas dd-mm-yyyy, columnas autoajustadas, encabezados congelados y autofiltros. Cero Markdown en celdas.
- PDF y Word: banner corporativo verde bosque #1E4D2B, tarjetas con sumas reales verificadas y tablas proporcionales. La fecha se imprime una sola vez en la cabecera; no la repitas en el cuerpo. Sanitización absoluta: sin "\\$", sin LaTeX como "$50kg/ha$" y sin marcas crudas "#" ni "**".
- Guardar o actualizar un recuerdo persistente: {"accion":"guardar_memoria","categoria":"Negocio|Preferencias|Instrucciones|Decisiones|Datos personales|Clientes|Productos|Finanzas|Contexto|Otros","dato":"..."}
- Olvidar o actualizar un recuerdo: {"accion":"olvidar_memoria","clave_o_id":"...","nuevo_dato":"..."}
- Nunca pidas contraseñas, claves ni tokens.

MEMORIA PERSISTENTE:
- Usa recuerdos, preferencias, reglas de cálculo, decisiones y contexto de forma natural, como si ya conocieras al usuario. No digas "según mi base de datos", "como te dije antes", "he guardado en mi memoria" ni anuncies que consultas memoria.
- Si el usuario dice "recuerda que...", "mi nombre es...", "siempre aplica...", "ten en cuenta...", guarda el dato con la acción guardar_memoria.
- Si dice que un dato cambió o ya no aplica, usa olvidar_memoria (con nuevo_dato para actualizarlo) o guardar_memoria con el dato nuevo, sin conservar la versión anterior.
- No guardes contraseñas, claves, tokens ni datos de tarjetas.
- Confirma el cambio con una frase breve y natural, sin decir que lo guardaste en tu memoria.

TIEMPO Y CONTINUIDAD:
- Cada solicitud incluye un bloque "### RELOJ Y CONTEXTO TEMPORAL EN VIVO" con fecha, hora, zona horaria, día, semana, trimestre y tiempo desde el último mensaje, leídos del dispositivo al enviar. Es la única referencia válida: no uses fechas memorizadas ni supongas otro año. Es información interna: no la recites en el saludo ni en respuestas normales.
- Resuelve "hoy", "mañana", "ayer", "esta semana", "próximo lunes", "fin de mes" y cualquier plazo con esa fecha y zona horaria. Para ventas o asesorías recientes, filtra con las fechas reales de los registros; si un registro no tiene fecha, dilo en vez de inventarla.
- Menciona la fecha, la hora o la zona horaria solo cuando el usuario pregunte por ellas o por un plazo. La app fecha los entregables con esa fecha local en español y no debes repetirla dentro del contenido; el pie de una web lleva el año real del reloj.
- Si el usuario vuelve horas después o al día siguiente, retoma con naturalidad y saluda con la franja exacta (Buenos días / Buenas tardes / Buenas noches). En respuestas normales la primera oración sigue siendo la respuesta, no un saludo ni un dato del reloj.

REGLAS DE TRABAJO:
- No uses "-" como dato. Si falta un dato, deja el campo vacío o escribe "Pendiente".
- Nunca escribas enlaces internos como [texto](sandbox:/archivo) ni fórmulas LaTeX con "$". Escribe "$2~a~3~L/ha$" como "2 a 3 L/ha" y "\$0" como "$0": la app muestra el texto limpio y una tarjeta de descarga.
- En documentos usa tablas Markdown con "|" cuando haya datos comparables; el motor las convierte en tablas reales con bordes proporcionales.
- Para catálogos, el motor calcula los totales reales (suma de productos y valor de inventario Precio × Stock) desde la app: no escribas filas de totales manuales.
- Usa listas y pasos cuando ayuden. Ten en cuenta el historial de la conversación.

${SPANISH_ONLY_RULE}`

const rateBuckets = new Map()

const REASONING_START_PATTERNS = [
  /^\s*(?:analysis|reasoning|thinking|thoughts?|plan|notas? internas?|pensamiento(?:s)? internos?)\s*[:：]/i,
  /^\s*(?:the user is asking|looking at (?:the )?context|let me (?:create|start|use|check|prepare|build|write|make|review|try)|i (?:need to|should|must|will|have to|can|am going to)|i'll|we (?:need to|should|must|will)|now i (?:need|have|should)|for (?:the|this|excel|pdf|the catalog|the file)|since (?:the|this|only)|based on (?:the )?context|according to (?:the )?context|the (?:context|catalog|user|app|file) (?:is|has|shows|says)|so i|first,? i|next,? i|my plan|to do this)\b/i
]
const ENGLISH_HINT_RE = /\b(?:the|and|is|are|to|of|for|with|this|that|i|we|my|have|has|need|should|will|can|from|user|context|catalog|file|app|so|now|let|based|according|since|only|also|then|but|if|it|its|as|at|on|in|by|files|both|create|use|check)\b/gi
const SPANISH_HINT_RE = /\b(?:el|la|los|las|de|del|que|y|en|con|por|para|es|son|al|un|una|se|su|sus|te|tu|necesito|debo|voy|puedo|archivo|archivos|catálogo|catalogo|producto|productos|cliente|clientes|venta|ventas|informe|excel|pdf|word|resumen|total|precio|stock|fecha|generé|generar|creé|crear|lista|datos)\b/gi

// Quita razonamiento interno del texto que llega al usuario. Nunca se leen
// campos reasoning / reasoning_content: solo el contenido visible.
function sanitizeAssistantText(value, allowJson = true) {
  let source = typeof value === 'string' ? value : ''
  if (!source && value && typeof value === 'object' && typeof value.text === 'string') source = value.text
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

function conversationBlock(history) {
  const rows = (Array.isArray(history) ? history : [])
    .filter((item) => item && ['user', 'assistant', 'tool'].includes(item.role) && typeof item.content === 'string' && item.content.trim())
    .slice(-300)
  if (!rows.length) return ''
  const lines = []
  let total = 0
  for (const item of rows) {
    const speaker = item.role === 'user' ? 'USUARIO' : item.role === 'tool' ? 'HERRAMIENTA' : 'VITIMAX IA'
    const text = String(item.content).replace(/\s+/g, ' ').trim().slice(0, 4000)
    const when = item.at ? serverTemporalContext(item.at) : null
    const prefix = when ? `[${when.fechaISO} ${when.hora24Short}] ` : ''
    const line = `${prefix}${speaker}: ${text}`
    if (total + line.length > 60_000) break
    lines.push(line)
    total += line.length + 1
  }
  if (!lines.length) return ''
  const elapsed = serverElapsed(rows, serverTemporalContext())
  return `

### ${HISTORY_MARKER}
Continúa esta conversación con naturalidad. Es contexto ya compartido: no lo repitas, no lo enumeres y no digas que lo estás consultando.
${elapsed ? `${elapsed}\n` : ''}${lines.join('\n')}`
}

// System prompt único con memoria, reloj e historial, sin duplicar bloques que
// la app ya haya inyectado antes de enviar la solicitud.
function buildSystemPrompt(system, memories, history, temporal) {
  const base = typeof system === 'string' && system.trim() ? system.slice(0, 100_000) : systemInstruction
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

async function askOpenRouter({ catalog, message, history, attachment, signal, system, memories, temporal, maxTokens = 6000 }) {
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
  const fullSystem = buildSystemPrompt(system, memories, history, temporal)
  const messages = [
    { role: 'system', content: fullSystem },
    ...(fullSystem.includes(HISTORY_MARKER) ? [] : history.map((item) => ({ role: item.role, content: item.content }))),
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
  // Solo se lee content: nunca reasoning ni reasoning_content.
  const aiMessage = payload.choices?.[0]?.message || {}
  return { text: sanitizeAssistantText(aiMessage.content), model: payload.model || model }
}

async function askGemini({ catalog, message, history, attachment, signal, system, memories, temporal, maxTokens = 6000 }) {
  const parts = [{ text: `CATÁLOGO REAL DE VITIMAX:\n${catalog}\n\nCONSULTA DEL USUARIO:\n${message}` }]
  for (const item of attachmentParts(attachment)) {
    if (item.kind === 'text') parts.push({ text: item.text })
    if (item.kind === 'image') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
    if (item.kind === 'file') parts.push({ inline_data: { mime_type: item.mimeType, data: item.data } })
  }
  const fullSystem = buildSystemPrompt(system, memories, history, temporal)
  const contents = fullSystem.includes(HISTORY_MARKER) ? [] : history.map((item) => ({
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
