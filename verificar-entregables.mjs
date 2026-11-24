import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import {
  buildConversationSection,
  buildDemoProducts,
  buildMemorySection,
  createExcelBlob,
  filterJunkProducts,
  isJunkProductName,
  mergeWorkbookSheets,
  normalizeImportedRecords
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
for (const rule of ['Próximos Pasos', 'COUNTIF', 'verde bosque']) {
  assert.ok(aiDataCode.includes(rule) || serverCode.includes(rule), `Falta la regla ${rule} en el prompt`)
}
assert.match(mainCode, /tool_choice:\s*"auto"|tool_choice\s*=\s*"auto"/)
assert.match(serverCode, /HISTORY_MARKER/)
assert.ok(!serverCode.includes("fullSystem.includes('HISTORIAL DE CONVERSACIÓN')"), 'El proxy no debe duplicar el historial')
const fenced = extractFencedCode('```js\nconsole.log(1)\n```\n```js\n' + 'x'.repeat(500) + '\n```')
assert.equal(fenced.code.length, 500, 'Se conserva el bloque de código completo más largo')
console.log('Memoria + historial en system prompt · código completo: OK')

// Excel multi-hoja con KPIs y fórmulas reales.
const sheets = mergeWorkbookSheets([], products, {}, { allowDemo: true })
const excelBlob = await createExcelBlob(sheets, { products })
const excelBytes = await saveBlob(excelBlob, '01-resumen-ejecutivo.xlsx')
console.log(`Excel: ${excelBytes} bytes · ${sheets.length} hojas`)

// PDF con contenido sucio para verificar limpieza.
const dirty = '# Informe VITIMAX\n\n## Dosis\n$50kg/ha$ y **negrita** con # raw\n\n| Producto | Precio |\n| --- | --- |\n| Ejemplo | $15.000 |'
const pdf = buildPdfDocument({ title: 'Informe ejecutivo', subtitle: 'Datos reales', content: dirty, catalog: products, isDemo: false })
const pdfBytes = typeof pdf?.output === 'function'
  ? Buffer.from(await pdf.output('arraybuffer'))
  : Buffer.from(await pdf.arrayBuffer())
fs.writeFileSync(path.join(out, '02-informe.pdf'), pdfBytes)
console.log(`PDF: ${pdfBytes.length} bytes`)

// Word ejecutivo.
const wordBlob = await buildDocxDocument({ title: 'Informe ejecutivo', subtitle: 'Datos reales', content: dirty, catalog: products, isDemo: false })
const wordBytes = await saveBlob(wordBlob, '03-informe.docx')
console.log(`Word: ${wordBytes} bytes`)

// PowerPoint con tablas de catálogo (5 filas por slide) y cierre.
const pptBlob = await buildPptxDocument({ title: 'Catálogo VITIMAX', subtitle: 'Presentación ejecutiva', content: '# Catálogo\n- Revisar precios\n- Validar stock', catalog: products, isDemo: false })
const pptBytes = await saveBlob(pptBlob, '04-presentacion.pptx')
console.log(`PowerPoint: ${pptBytes} bytes`)

// Página web completa.
const html = buildHtmlPage({ title: 'Catálogo VITIMAX', subtitle: 'Soluciones agrícolas', html: '', catalog: products, phone: '+56912345678', isDemo: false })
fs.writeFileSync(path.join(out, '05-catalogo.html'), html)
assert.equal((html.match(/cdn\.tailwindcss\.com/g) || []).length, 1, 'Una sola carga de Tailwind')
for (const product of products) {
  assert.ok(html.includes(product.name), `Falta el producto ${product.name} en la web`)
}
assert.match(html, /Cotizar por WhatsApp|wa\.me/i, 'Falta el botón de WhatsApp')
console.log(`Web: ${html.length} bytes · 10/10 productos · Tailwind único · WhatsApp OK`)

console.log('Todos los entregables se generaron sin errores.')
