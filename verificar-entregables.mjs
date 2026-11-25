import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import {
  DIRECT_AI_SYSTEM_PROMPT,
  buildConversationSection,
  buildDemoProducts,
  buildMemorySection,
  buildWelcomeMessage,
  createExcelBlob,
  filterJunkProducts,
  isJunkProductName,
  buildTemporalSection,
  formatElapsedSince,
  formatSpanishDate,
  getTemporalContext,
  mergeWorkbookSheets,
  normalizeImportedRecords,
  resolveDisplayName
} from './src/ai-data.js'
import { sanitizeText } from './src/text-utils.js'
import {
  buildDocxDocument,
  extractFencedCode,
  buildHtmlPage,
  buildPdfDocument,
  buildPptxDocument
} from './src/doc-engine.js'

const out = path.join(process.cwd(), 'verificacion-entregables')
fs.mkdirSync(out, { recursive: true })

async function saveBlob(blob, filename) {
  const bytes = Buffer.from(await blob.arrayBuffer())
  fs.writeFileSync(path.join(out, filename), bytes)
  return bytes.length
}

const demo = buildDemoProducts()
const extra = [
  { id: 'D-07', name: 'VITIMAX Raíz 7', category: 'Sustratos', price: 12990, stock: 12, format: 'Saco 50 kg', description: 'Sustrato estructurado para enraizamiento.' },
  { id: 'D-08', name: 'VITIMAX Hoja Verde', category: 'Nutrición', price: 8990, stock: 18, format: 'Botella 1 L', description: 'Nutriente foliar de aplicación segura.' },
  { id: 'D-09', name: 'VITIMAX Fruto Plus', category: 'Nutrición', price: 14990, stock: 7, format: 'Botella 5 L', description: 'Refuerzo nutricional en etapa de fruto.' },
  { id: 'D-10', name: 'VITIMAX Suelo Activo', category: 'Suelos', price: 9990, stock: 25, format: 'Saco 25 kg', description: 'Mejorador biológico de suelo.' }
]
const products = filterJunkProducts([
  ...demo,
  ...extra,
  { id: 'JUNK-1', name: 'Total Valorado Inventario', price: 999999, stock: 99 },
  { id: 'JUNK-2', name: 'Total', price: 1, stock: 1 }
])
assert.equal(products.length, 10, 'Deben quedar exactamente 10 productos reales')
assert.ok(isJunkProductName('Total Valorado Inventario'))
assert.ok(isJunkProductName('Total'))
const clients = normalizeImportedRecords('clientes', [
  { Nombre: 'Juan Pérez', Teléfono: '+56911111111' },
  { Nombre: 'Total', Teléfono: '+56922222222' },
  { Teléfono: '+56933333333' }
])
assert.equal(clients.length, 1, 'Los registros sin nombre real o de resumen se descartan')
console.log(`Productos reales: ${products.length} (demo: ${demo.length} + 4) · registros limpios: ${clients.length}`)

// Sanitización absoluta de LaTeX, escapes y Markdown.
assert.equal(sanitizeText('$50kg/ha$'), '50 kg/ha')
assert.equal(sanitizeText('\\$0'), '$0')
assert.equal(sanitizeText('**negrita** y `#`'), 'negrita y')
console.log('Sanitización de texto: OK')

// Memoria + historial inyectados en el system prompt.
const memoryPrompt = buildMemorySection([{ categoria: 'Negocio', contenido: 'El IVA es 19%' }])
const historyPrompt = buildConversationSection([
  { role: 'user', content: 'Hola' },
  { role: 'assistant', content: 'Dime qué necesitas' }
])
assert.match(memoryPrompt, /MEMORIA PERSISTENTE/)
assert.match(historyPrompt, /HISTORIAL COMPLETO DE LA CONVERSACIÓN/)
const aiDataCode = fs.readFileSync('src/ai-data.js', 'utf8')
const mainCode = fs.readFileSync('src/main.jsx', 'utf8')
const serverCode = fs.readFileSync('ai-server.cjs', 'utf8')
const stylesCode = fs.readFileSync('src/styles.css', 'utf8')
for (const rule of ['Próximos Pasos', 'COUNTIF', 'verde bosque']) {
  assert.ok(aiDataCode.includes(rule) || serverCode.includes(rule), `Falta la regla ${rule} en el prompt`)
}
assert.match(mainCode, /tool_choice:\s*"auto"|tool_choice\s*=\s*"auto"/)
assert.match(serverCode, /HISTORY_MARKER/)
assert.ok(!serverCode.includes("fullSystem.includes('HISTORIAL DE CONVERSACIÓN')"), 'El proxy no debe duplicar el historial')
const fenced = extractFencedCode('```js\nconsole.log(1)\n```\n```js\n' + 'x'.repeat(500) + '\n```')
assert.equal(fenced.code.length, 500, 'Se conserva el bloque de código completo más largo')
console.log('Memoria + historial en system prompt · código completo: OK')

// Reloj en vivo: franjas exactas, fecha local, semana, zona y elapsed time.
const morning = getTemporalContext(new Date(2026, 10, 23, 5, 0, 0))
const afternoon = getTemporalContext(new Date(2026, 10, 23, 12, 0, 0))
const night = getTemporalContext(new Date(2026, 10, 23, 20, 0, 0))
const midnight = getTemporalContext(new Date(2026, 10, 23, 0, 0, 0))
assert.equal(morning.franja, 'Buenos días')
assert.equal(afternoon.franja, 'Buenas tardes')
assert.equal(night.franja, 'Buenas noches')
assert.equal(midnight.franja, 'Buenas noches')
assert.equal(morning.hora24, '05:00:00')
assert.equal(afternoon.hora24, '12:00:00')
assert.equal(night.hora24, '20:00:00')
assert.equal(morning.fechaCompleta, 'Lunes, 23 de noviembre de 2026')
assert.equal(morning.numeroSemana, 48)
assert.equal(morning.diaSemana, 'Lunes')
assert.ok(morning.zonaHoraria && /^UTC[+-]\d/.test(morning.offset), 'Falta la zona horaria local')
assert.match(morning.hora12, /a\. m\.|p\. m\./)
const temporalSection = buildTemporalSection(morning, [])
assert.match(temporalSection, /### RELOJ Y CONTEXTO TEMPORAL EN VIVO: Hoy es/)
assert.match(temporalSection, /Hoy es Lunes, 23 de noviembre de 2026, son las 05:00:00 \(Buenos días\)/)
assert.match(temporalSection, /Semana 48 del año/)
assert.match(temporalSection, /Es la primera interacción de esta conversación\./)
const elapsedNow = new Date(2026, 10, 23, 12, 0, 0)
const elapsed = formatElapsedSince(new Date(2026, 10, 23, 10, 0, 0).toISOString(), elapsedNow)
assert.match(elapsed, /han pasado 2 horas\./)
assert.match(buildTemporalSection(afternoon, [{ role: 'user', content: 'Hola', at: new Date(2026, 10, 23, 10, 0, 0).toISOString() }]), /han pasado 2 horas\./)
assert.match(buildConversationSection([{ role: 'user', content: 'Hola', at: new Date(2026, 10, 23, 10, 0, 0).toISOString() }]), /\[2026-11-23 10:00\] USUARIO: Hola/)
assert.match(formatSpanishDate(new Date(2026, 10, 23, 12, 0, 0)), /23 de noviembre de 2026/)
assert.ok(mainCode.includes('buildTemporalSection(getTemporalContext(), source)'), 'El reloj no se inyecta en cada llamada IA')
assert.ok(mainCode.includes('temporal: getTemporalContext()'), 'El proxy no recibe el reloj del dispositivo')
assert.ok(serverCode.includes('serverTemporalBlock('), 'El proxy no construye el bloque temporal')
assert.ok(serverCode.includes('TEMPORAL_MARKER'), 'El proxy no evita duplicar el bloque temporal')
assert.ok(mainCode.includes('baidu/ernie-5.5-flash-preview:free'), 'El BOT no usa un modelo gratuito y estable por defecto')
console.log('Reloj local: 4 franjas, semana 48, zona horaria y elapsed time: OK')

// Reloj en pantalla: tick reactivo de 1 segundo y dígitos sin salto horizontal.
assert.ok(mainCode.includes('const [now, setNow] = useState(() => new Date())'), 'Falta el estado reactivo now')
assert.ok(mainCode.includes('setInterval(() => setNow(new Date()), 1000)'), 'El reloj no avanza cada segundo')
assert.ok(mainCode.includes('return () => clearInterval(id)'), 'Falta limpiar el intervalo del reloj')
assert.ok(mainCode.includes('const temporal = getTemporalContext(now)'), 'El reloj visible no se deriva del estado now')
assert.ok(!mainCode.includes('useTemporalClock'), 'Quedó un reloj cacheado fuera del render')
assert.ok(!/temporal = getTemporalContext\(\)/.test(mainCode), 'Un componente aún usa un reloj estático por defecto')
assert.ok(mainCode.includes('className="temporal-clock"'), 'Falta la clase de dígitos tabulares en la hora')
assert.match(stylesCode, /\.temporal-clock\{[^}]*font-variant-numeric:tabular-nums/)
console.log('Reloj en pantalla: tick de 1s, derivado de now y dígitos tabulares: OK')

// Interfaz del chat: cabecera responsiva, nombres de archivo intactos y
// sugerencias sin barra de scroll visible.
assert.ok(mainCode.includes('className="assistant-subline"'), 'Falta la línea secundaria compacta de la cabecera')
assert.ok(mainCode.includes('className="assistant-tz"'), 'Falta ocultar la zona horaria en móvil')
assert.ok(mainCode.includes('placeholder="Pregunta o pide crear algo..."'), 'El placeholder del chat no se acortó')
assert.ok(mainCode.includes('CHAT_FILE_TOKEN'), 'Falta proteger los nombres de archivo en el Markdown del chat')
assert.ok(mainCode.includes("(?:pdf|xlsx|docx|pptx|html|csv)"), 'La protección no cubre todas las extensiones')
assert.ok(mainCode.includes('restoreChatFileTokens'), 'No se restauran los nombres de archivo protegidos')
assert.match(stylesCode, /\.assistant-head \.assistant-title\{min-width:0;flex:1\}/)
assert.match(stylesCode, /\.assistant-head \.assistant-tz\{display:none\}/)
assert.match(stylesCode, /\.quick-prompts\{[^}]*scrollbar-width:none/)
assert.match(stylesCode, /\.quick-prompts::-webkit-scrollbar\{display:none\}/)
console.log('Chat UI: cabecera sin recortes, archivos con "_" y sugerencias sin scrollbar: OK')

// Saludo inicial exacto por franja, sin fecha, hora, zona ni nombres falsos.
assert.equal(buildWelcomeMessage('', morning), 'Buenos días. ¿En qué te ayudo hoy?')
assert.equal(buildWelcomeMessage('Cristóbal', afternoon), 'Buenas tardes, Cristóbal. ¿Qué resolvemos hoy?')
assert.equal(buildWelcomeMessage('Cristóbal', night), 'Buenas noches, Cristóbal. ¿En qué trabajamos ahora?')
for (const clock of [morning, afternoon, night]) {
  const greeting = buildWelcomeMessage('Invitado', clock)
  assert.ok(!/invitado|administrador|\d|UTC|America\//i.test(greeting), `Saludo con relleno: ${greeting}`)
}
assert.equal(resolveDisplayName({ profileName: 'Invitado', displayName: 'Administrador' }), '')
assert.equal(resolveDisplayName({
  profileName: 'Invitado',
  memories: [{ contenido: 'Mi nombre es Cristóbal Suárez' }]
}), 'Cristóbal')
assert.ok(!mainCode.includes("profile.name || 'administrador'"), 'El saludo todavía usa un nombre falso')
assert.match(buildTemporalSection(morning, []), /No menciones la fecha, la hora exacta ni la zona horaria/)
for (const phrase of ['Soy VITIMAX IA', 'Como asistente de inteligencia artificial', 'Según mi reloj', 'He guardado en mi memoria']) {
  assert.ok(DIRECT_AI_SYSTEM_PROMPT.includes(phrase), `Falta prohibir "${phrase}" en el prompt`)
  assert.ok(serverCode.includes(phrase), `Falta prohibir "${phrase}" en el proxy`)
}
assert.match(DIRECT_AI_SYSTEM_PROMPT, /humor sutil/)
assert.match(DIRECT_AI_SYSTEM_PROMPT, /no escribas "Fecha:"/i)
console.log('Saludo exacto, tono humano y reloj silencioso: OK')

// Excel multi-hoja con KPIs y fórmulas reales.
const sheets = mergeWorkbookSheets([], products, {}, { allowDemo: true })
const excelBlob = await createExcelBlob(sheets, { products })
const excelBytes = await saveBlob(excelBlob, '01-resumen-ejecutivo.xlsx')
console.log(`Excel: ${excelBytes} bytes · ${sheets.length} hojas`)

// PDF con contenido sucio y fecha repetida para verificar limpieza.
const reportDate = 'Lunes, 23 de noviembre de 2026'
const dirty = `# Informe VITIMAX\n\nFecha: ${reportDate}\n\n## Dosis\n$50kg/ha$ y **negrita** con # raw\n\n| Producto | Precio |\n| --- | --- |\n| Ejemplo | $15.000 |`
const pdf = buildPdfDocument({ title: 'Informe ejecutivo', subtitle: `${reportDate} · Datos reales`, content: dirty, catalog: products, isDemo: false, dateLabel: reportDate })
const pdfBytes = typeof pdf?.output === 'function'
  ? Buffer.from(await pdf.output('arraybuffer'))
  : Buffer.from(await pdf.arrayBuffer())
fs.writeFileSync(path.join(out, '02-informe.pdf'), pdfBytes)
console.log(`PDF: ${pdfBytes.length} bytes`)

// Catálogo vacío: estado inicial real + guía de 4 pasos, sin datos de ejemplo.
const emptyPdf = buildPdfDocument({
  title: 'Estado inicial de VITIMAX',
  subtitle: 'Diagnóstico de arranque con datos reales',
  content: '# Estado inicial de VITIMAX\n\nConteos reales del centro de gestión.\n\n## Arranque rápido\n\n1. **Carga tu catálogo.** Importa el Excel oficial de VITIMAX.\n2. **Agrega tus clientes.** Registra contactos reales con nombre y teléfono.\n3. **Registra tus primeras ventas.** Asocia cada venta a un cliente y un producto.\n4. **Publica y ajusta.** Revisa la web, compártela por WhatsApp y corrige con datos reales.',
  catalog: [],
  isDemo: false,
  dateLabel: reportDate,
  emptyState: { productos: 0, clientes: 0, ventas: 0 }
})
const emptyPdfBytes = typeof emptyPdf?.output === 'function'
  ? Buffer.from(await emptyPdf.output('arraybuffer'))
  : Buffer.from(await emptyPdf.arrayBuffer())
fs.writeFileSync(path.join(out, '06-estado-inicial.pdf'), emptyPdfBytes)
console.log(`PDF estado inicial: ${emptyPdfBytes.length} bytes`)

// Word ejecutivo.
const wordBlob = await buildDocxDocument({ title: 'Informe ejecutivo', subtitle: 'Datos reales', content: dirty, catalog: products, isDemo: false })
const wordBytes = await saveBlob(wordBlob, '03-informe.docx')
console.log(`Word: ${wordBytes} bytes`)

// PowerPoint con tablas de catálogo (5 filas por slide) y cierre.
const pptBlob = await buildPptxDocument({ title: 'Catálogo VITIMAX', subtitle: 'Presentación ejecutiva', content: '# Catálogo\n- Revisar precios\n- Validar stock', catalog: products, isDemo: false })
const pptBytes = await saveBlob(pptBlob, '04-presentacion.pptx')
console.log(`PowerPoint: ${pptBytes} bytes`)

// Página web completa.
const html = buildHtmlPage({ title: 'Catálogo VITIMAX', subtitle: 'Soluciones agrícolas', html: '', catalog: products, phone: '+56912345678', isDemo: false, dateLabel: 'Lunes, 23 de noviembre de 2026' })
fs.writeFileSync(path.join(out, '05-catalogo.html'), html)
assert.equal((html.match(/cdn\.tailwindcss\.com/g) || []).length, 1, 'Una sola carga de Tailwind')
for (const product of products) {
  assert.ok(html.includes(product.name), `Falta el producto ${product.name} en la web`)
}
assert.match(html, /Cotizar por WhatsApp|wa\.me/i, 'Falta el botón de WhatsApp')
assert.ok(html.includes('Lunes, 23 de noviembre de 2026'), 'La web no usa la fecha sincronizada')
console.log(`Web: ${html.length} bytes · 10/10 productos · Tailwind único · WhatsApp OK`)

console.log('Todos los entregables se generaron sin errores.')
