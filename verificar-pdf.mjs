import fs from 'node:fs'
import assert from 'node:assert/strict'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import { parseDocument, inlineSegments, buildPdfDocument } from './src/doc-engine.js'

async function readPdf(filename) {
  const data = new Uint8Array(fs.readFileSync(filename))
  const doc = await pdfjs.getDocument({ data, useSystemFonts: true }).promise
  let text = ''
  for (let page = 1; page <= doc.numPages; page += 1) {
    const content = await (await doc.getPage(page)).getTextContent()
    text += `${content.items.map((item) => item.str).join(' ')} `
  }
  return { doc, text: text.replace(/\s+/g, ' ') }
}

const report = await readPdf('verificacion-entregables/02-informe.pdf')
assert.ok(!report.text.includes('**'), 'El PDF no debe contener **')
assert.ok(!report.text.includes('$50kg/ha$'), 'El PDF no debe contener LaTeX')
assert.ok(!report.text.includes('\\$'), 'El PDF no debe contener dólares escapados')
assert.ok(report.text.includes('50 kg/ha'), 'El PDF debe contener la dosis en texto plano')
assert.ok(report.text.includes('$15.000'), 'El PDF debe conservar precios en CLP')
assert.ok(report.text.includes('23 de noviembre de 2026'), 'El PDF debe usar la fecha sincronizada')
assert.ok(!/Fecha\s*:/i.test(report.text), 'El PDF no debe repetir la fecha con la etiqueta "Fecha:"')
assert.equal(report.text.split('23 de noviembre de 2026').length - 1, 1, 'La fecha debe aparecer una sola vez')
console.log(`PDF páginas: ${report.doc.numPages} · texto limpio y fecha única: OK`)
console.log(report.text.slice(0, 300))

const empty = await readPdf('verificacion-entregables/06-estado-inicial.pdf')
const lower = empty.text.toLowerCase()
assert.ok(lower.includes('estado inicial'), 'Falta el estado inicial del catálogo vacío')
for (const label of ['productos', 'clientes', 'ventas']) {
  assert.ok(lower.includes(label), `Falta la tarjeta de ${label}`)
}
assert.ok(lower.includes('arranque rápido'), 'Falta la guía de arranque rápido')
for (const step of ['carga tu catálogo', 'agrega tus clientes', 'registra tus primeras ventas', 'publica y ajusta']) {
  assert.ok(lower.includes(step), `Falta el paso "${step}"`)
}
for (const label of ['2\\.', '3\\.', '4\\.']) {
  assert.match(empty.text, new RegExp(label), 'La guía debe numerar los pasos del 1 al 4')
}
assert.ok(!empty.text.includes('Ejemplo'), 'El estado inicial no debe inventar productos de ejemplo')
assert.ok(!empty.text.includes('$15.000'), 'El estado inicial no debe mostrar precios ajenos')
assert.equal(empty.text.split('23 de noviembre de 2026').length - 1, 1, 'La fecha debe aparecer una sola vez')
console.log(`PDF estado inicial: ${empty.doc.numPages} página(s) · 3 tarjetas + 4 pasos: OK`)

// Listas numeradas: las líneas en blanco entre ítems no deben reiniciar el
// contador, y el título en negrita debe quedar separado por un espacio limpio.
const listContent = [
  '# Guía de arranque',
  '',
  '1. **Carga tu catálogo.**Importa el Excel oficial con precios reales.',
  '',
  '2. **Agrega tus clientes.**Registra nombre, teléfono y localidad.',
  '',
  '3. **Registra tus ventas.**Asocia cada venta a cliente y producto.',
  '',
  '4. **Publica y ajusta.**Corrige precios y stock con datos reales.'
].join('\n')
const numbers = parseDocument(listContent).filter((block) => block.type === 'numbers')
assert.equal(numbers.length, 1, 'La lista con líneas en blanco debe ser un solo bloque')
assert.equal(numbers[0].items.length, 4, 'Deben conservarse los 4 ítems')
assert.equal(numbers[0].start, 1, 'La numeración debe arrancar en 1')
assert.equal(
  inlineSegments('**Carga tu catálogo.**Importa el Excel.').map((part) => part.text).join(''),
  'Carga tu catálogo. Importa el Excel.',
  'El título en negrita debe llevar un espacio antes de la descripción'
)
const numberedPdf = buildPdfDocument({
  title: 'Numeración secuencial',
  content: listContent,
  dateLabel: 'Martes, 24 de noviembre de 2026'
})
const numberedBytes = typeof numberedPdf?.output === 'function'
  ? Buffer.from(await numberedPdf.output('arraybuffer'))
  : Buffer.from(await numberedPdf.arrayBuffer())
fs.writeFileSync('verificacion-entregables/07-numeracion.pdf', numberedBytes)
const numbered = await readPdf('verificacion-entregables/07-numeracion.pdf')
const sequence = [...numbered.text.matchAll(/\b([1-4])\.\s/g)].map((match) => match[1])
assert.deepEqual(sequence, ['1', '2', '3', '4'], `Numeración incorrecta en el PDF: ${sequence.join(',')}`)
assert.ok(numbered.text.includes('catálogo. Importa'), 'Falta el espacio limpio tras el título en negrita')
console.log('PDF listas numeradas: secuencia 1-4 y espaciado del título en negrita: OK')
