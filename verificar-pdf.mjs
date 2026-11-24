import fs from 'node:fs'
import assert from 'node:assert/strict'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const data = new Uint8Array(fs.readFileSync('verificacion-entregables/02-informe.pdf'))
const doc = await pdfjs.getDocument({ data, useSystemFonts: true }).promise
let text = ''
for (let page = 1; page <= doc.numPages; page += 1) {
  const content = await (await doc.getPage(page)).getTextContent()
  text += `${content.items.map((item) => item.str).join(' ')} `
}
assert.ok(!text.includes('**'), 'El PDF no debe contener **')
assert.ok(!text.includes('$50kg/ha$'), 'El PDF no debe contener LaTeX')
assert.ok(!text.includes('\\$'), 'El PDF no debe contener dólares escapados')
assert.ok(text.includes('50 kg/ha'), 'El PDF debe contener la dosis en texto plano')
assert.ok(text.includes('$15.000'), 'El PDF debe conservar preios en CLP')
assert.ok(text.includes('23 de noviembre de 2026'), 'El PDF debe usar la fecha sincronizada')
console.log(`PDF páginas: ${doc.numPages} · texto limpio verificado`)
console.log(text.slice(0, 420).replace(/\s+/g, ' '))
