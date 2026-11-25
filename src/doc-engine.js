import { jsPDF } from 'jspdf'
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  PageBreak,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType
} from 'docx'
import pptxgen from 'pptxgenjs'
import {
  computeCatalogTotals,
  formatClp,
  parseNumberLoose,
  sanitizeText
} from './text-utils.js'
import { formatSpanishDate } from './temporal.js'

// ---------------------------------------------------------------------------
// Paleta corporativa VITIMAX
// ---------------------------------------------------------------------------
export const BRAND = {
  forest: '1E4D2B',
  deep: '12321F',
  mid: '2F8058',
  lime: 'D8F55F',
  zebra: 'F4F7F4',
  totals: 'E6EFE8',
  ink: '24382B',
  muted: '5F7969',
  line: 'DDE7DA',
  white: 'FFFFFF'
}

const rgb = (hex) => [
  parseInt(String(hex).slice(0, 2), 16),
  parseInt(String(hex).slice(2, 4), 16),
  parseInt(String(hex).slice(4, 6), 16)
]

function todayLabel() {
  return formatSpanishDate(new Date())
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

// La fecha se imprime una sola vez en la cabecera verde. Si el modelo la
// repite en el subtítulo o en el cuerpo, se elimina aquí sin tocar fechas
// históricas distintas a la del día.
function stripHeaderDate(value, date) {
  let text = String(value ?? '')
  const dateOnly = String(date || '').replace(/^\s*[^,]+,\s*/, '').trim()
  const values = [date, dateOnly].filter(Boolean)
  if (!values.length) return text
  for (const item of values) {
    const escaped = escapeRegExp(item)
    text = text.replace(new RegExp(`^\\s*fecha\\s*[:\\-]\\s*${escaped}\\s*`, 'i'), '')
    text = text.replace(new RegExp(escaped, 'gi'), '')
  }
  text = text.replace(/^\s*fecha\s*[:\-]\s*/i, '')
  text = text.replace(/^\s*[·|,;–—-]+\s*/, '')
  return text.replace(/[ \t]{2,}/g, ' ').trim()
}

function sanitizeBlockDates(blocks, date) {
  return blocks.map((block) => {
    const next = { ...block }
    if (typeof next.text === 'string') next.text = stripHeaderDate(next.text, date)
    if (Array.isArray(next.items)) next.items = next.items.map((item) => stripHeaderDate(item, date)).filter(Boolean)
    if (Array.isArray(next.header)) next.header = next.header.map((cell) => stripHeaderDate(cell, date))
    if (Array.isArray(next.rows)) next.rows = next.rows.map((row) => Array.isArray(row) ? row.map((cell) => stripHeaderDate(cell, date)) : row)
    if (typeof next.code === 'string') next.code = next.code.split('\n').map((line) => stripHeaderDate(line, date)).join('\n')
    return next
  }).filter((block) => {
    if (block.type === 'table') return (block.rows || []).length > 0
    if (['heading', 'paragraph', 'quote'].includes(block.type)) return Boolean(String(block.text || '').trim())
    if (['bullets', 'numbers'].includes(block.type)) return (block.items || []).length > 0
    return true
  })
}

// ---------------------------------------------------------------------------
// 1. Limpieza de Markdown y enlaces internos
// ---------------------------------------------------------------------------

// Quita enlaces internos tipo [texto](sandbox:/...) y deja solo el texto.
// También elimina rutas sueltas sandbox:/... que algunos modelos escriben.
export function stripSandboxLinks(value) {
  return String(value ?? '')
    .replace(/\[([^\]]*)\]\(\s*(?:sandbox|file|javascript|data):[^)]*\)/gi, '$1')
    .replace(/\[([^\]]*)\]\(([^)]*)\)/g, '$1')
    .replace(/\(\s*(?:sandbox|file):[^)]*\)/gi, '')
    .replace(/(?:sandbox|file):\/?[^\s)]+/gi, '')
    .replace(/[ \t]{2,}/g, ' ')
}

// Limpia marcas Markdown y LaTeX dentro de una línea, conservando el contenido.
function cleanInline(value) {
  return sanitizeText(value).replace(/\s+/g, ' ')
}

// Devuelve segmentos con negrita real: [{ text, bold }]
export function inlineSegments(value) {
  const source = stripSandboxLinks(String(value ?? '')).replace(/\r\n?/g, '\n')
  const parts = []
  const pattern = /(\*\*|__)([\s\S]+?)\1/g
  let last = 0
  let match = pattern.exec(source)
  while (match) {
    if (match.index > last) parts.push({ text: cleanInline(source.slice(last, match.index)), bold: false })
    parts.push({ text: cleanInline(match[2]), bold: true })
    last = match.index + match[0].length
    match = pattern.exec(source)
  }
  if (last < source.length) parts.push({ text: cleanInline(source.slice(last)), bold: false })
  const cleaned = parts
    .map((part) => ({ ...part, text: part.text.replace(/\s+/g, ' ') }))
    .filter((part) => part.text.length)
  if (!cleaned.length) return []
  // Recorta espacios al inicio y al final del conjunto, no entre segmentos.
  cleaned[0] = { ...cleaned[0], text: cleaned[0].text.replace(/^\s+/, '') }
  const lastIndex = cleaned.length - 1
  cleaned[lastIndex] = { ...cleaned[lastIndex], text: cleaned[lastIndex].text.replace(/\s+$/, '') }
  return cleaned.filter((part) => part.text.length)
}

export function inlineText(value) {
  return inlineSegments(value).map((part) => part.text).join('')
}

export function plainText(value) {
  return inlineText(value).replace(/\s+/g, ' ').trim()
}

// ---------------------------------------------------------------------------
// 2. Parser de bloques (encabezados, párrafos, listas, tablas, código)
// ---------------------------------------------------------------------------

function splitTableRow(line) {
  return line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((cell) => cell.trim())
}

function isTableSeparator(line) {
  return /^\s*\|?[\s:|-]*-[\s:|-]*\|[\s:|-]*\|?\s*$/.test(line)
}

export function parseDocument(value) {
  const lines = stripSandboxLinks(String(value ?? '')).replace(/\r\n?/g, '\n').split('\n')
  const blocks = []
  let paragraph = []
  let list = null
  let fence = null

  const flushParagraph = () => {
    if (!paragraph.length) return
    const text = paragraph.join(' ').trim()
    paragraph = []
    if (text) blocks.push({ type: 'paragraph', text })
  }
  const flushList = () => {
    if (!list) return
    if (list.items.length) blocks.push(list)
    list = null
  }
  const flushAll = () => {
    flushParagraph()
    flushList()
  }

  for (let index = 0; index < lines.length; index += 1) {
    const raw = lines[index]
    const line = raw.replace(/\s+$/, '')

    if (fence) {
      if (/^\s*```/.test(line)) {
        blocks.push({ type: 'code', language: fence.language, code: fence.lines.join('\n').replace(/\s+$/, '') })
        fence = null
      } else {
        fence.lines.push(raw)
      }
      continue
    }

    const fenceStart = line.match(/^\s*```([a-zA-Z0-9+#.-]*)\s*$/)
    if (fenceStart) {
      flushAll()
      fence = { language: (fenceStart[1] || '').toLowerCase(), lines: [] }
      continue
    }

    if (!line.trim()) {
      flushAll()
      continue
    }

    const heading = line.match(/^\s*(#{1,6})\s+(.*)$/)
    if (heading) {
      flushAll()
      blocks.push({ type: 'heading', level: heading[1].length, text: heading[2].trim() })
      continue
    }

    if (/^\s*(?:---+|\*\*\*+|___+)\s*$/.test(line)) {
      flushAll()
      blocks.push({ type: 'divider' })
      continue
    }

    if (/^\s*\|.*\|\s*$/.test(line) && index + 1 < lines.length && isTableSeparator(lines[index + 1])) {
      flushAll()
      const header = splitTableRow(line)
      const rows = []
      index += 2
      while (index < lines.length && /^\s*\|.*\|\s*$/.test(lines[index])) {
        const cells = splitTableRow(lines[index])
        if (cells.some((cell) => cell.length)) rows.push(cells)
        index += 1
      }
      index -= 1
      blocks.push({ type: 'table', header, rows })
      continue
    }

    const bullet = line.match(/^\s*[-*+•]\s+(.*)$/)
    if (bullet) {
      flushParagraph()
      if (!list || list.type !== 'bullets') {
        flushList()
        list = { type: 'bullets', items: [] }
      }
      list.items.push(bullet[1].trim())
      continue
    }

    const numbered = line.match(/^\s*(\d{1,3})[.)]\s+(.*)$/)
    if (numbered) {
      flushParagraph()
      if (!list || list.type !== 'numbers') {
        flushList()
        list = { type: 'numbers', items: [] }
      }
      list.items.push(numbered[2].trim())
      continue
    }

    const quote = line.match(/^\s*>\s?(.*)$/)
    if (quote) {
      flushAll()
      blocks.push({ type: 'quote', text: quote[1].trim() })
      continue
    }

    flushList()
    paragraph.push(line.trim())
  }

  if (fence) blocks.push({ type: 'code', language: fence.language, code: fence.lines.join('\n').replace(/\s+$/, '') })
  flushAll()
  return blocks
}

function normalizeTitle(value) {
  return plainText(value).toLowerCase().replace(/[^a-z0-9áéíóúñü ]/gi, '').trim()
}

// ---------------------------------------------------------------------------
// Detección de documentos de catálogo y limpieza de filas de totales
// ---------------------------------------------------------------------------

const TOTALS_ROW_PATTERN = /^(gran\s+)?(total|totales|subtotal|suma|resumen|total\s+general|total\s+valorado|valor\s+total|valorizado)\b/i
const TOTALS_HEADING_PATTERN = /^(totales|total\s+general|resumen\s+de\s+totales|valor\s+total|total\s+valorado|valorizado)\b/i

function isTotalsRow(row) {
  return Array.isArray(row) && TOTALS_ROW_PATTERN.test(String(row[0] ?? '').trim())
}

export function isCatalogDocument(title, subtitle, blocks = []) {
  const haystack = `${title} ${subtitle}`.toLowerCase()
  if (/cat[aá]logo|inventario|productos|tienda|precios|lista de precios|vivero|agro|fertiliz|bioestimul/.test(haystack)) return true
  return blocks.some((block) => {
    if (block.type !== 'table') return false
    const header = (block.header || []).map((cell) => plainText(cell).toLowerCase())
    const hasProduct = header.some((cell) => /producto|nombre|descripci[oó]n/.test(cell))
    const hasMoney = header.some((cell) => /precio|valor|stock|inventario|cantidad/.test(cell))
    return hasProduct && hasMoney
  })
}

// Elimina filas fantasma de totales y tablas que quedan vacías.
function prepareDocumentBlocks(blocks, catalogMode) {
  if (!catalogMode) return blocks
  const output = []
  blocks.forEach((block) => {
    if (block.type === 'table') {
      const rows = (block.rows || []).filter((row) => !isTotalsRow(row))
      if (!rows.length) return
      output.push({ ...block, rows })
      return
    }
    if (block.type === 'heading' && TOTALS_HEADING_PATTERN.test(plainText(block.text))) return
    output.push(block)
  })
  return output
}

function drawTableWidths(doc, header, rows, contentW) {
  const pad = 4
  const weights = header.map((label, index) => {
    const key = plainText(label).toLowerCase()
    const longest = rows.reduce((max, row) => Math.max(max, String(row[index] ?? '').length), 0)
    let semantic = 16
    if (/^(id|c[oó]digo|sku|n[°º]|#)$/.test(key)) semantic = 6
    else if (/precio|valor|monto|total|stock|cantidad|unidades|iva|clp/.test(key)) semantic = 13
    else if (/nombre|producto/.test(key)) semantic = 26
    else if (/descripci[oó]n|dosis|uso|modo|recomendaci|caracter|detalle|notas/.test(key)) semantic = 34
    else if (/categor|formato|presentaci|estado|fecha|cultivo/.test(key)) semantic = 15
    const content = Math.min(40, Math.max(6, longest * 0.72 + 6))
    return Math.max(semantic, content)
  })
  const totalWeight = weights.reduce((sum, weight) => sum + weight, 0) || 1
  let widths = weights.map((weight) => (weight / totalWeight) * contentW)

  // Ancho mínimo para que encabezados como "Precio (CLP)" no se partan.
  const mins = header.map((label, index) => {
    doc.setFont('helvetica', 'bold')
    doc.setFontSize(8.8)
    const headerWidth = doc.getTextWidth(String(label || '')) + pad
    const longestWord = rows.reduce((max, row) => {
      doc.setFont('helvetica', 'normal')
      doc.setFontSize(8.6)
      const words = String(row[index] ?? '').split(/\s+/)
      return Math.max(max, ...words.map((word) => doc.getTextWidth(word) + pad))
    }, 0)
    return Math.min(contentW * 0.45, Math.max(12, headerWidth, Math.min(longestWord, contentW * 0.3)))
  })

  for (let pass = 0; pass < 4; pass += 1) {
    const deficit = widths.reduce((sum, width, index) => sum + Math.max(0, mins[index] - width), 0)
    if (deficit <= 0.01) break
    const surplus = widths.reduce((sum, width, index) => sum + Math.max(0, width - mins[index]), 0)
    if (surplus <= 0.01) break
    widths = widths.map((width, index) => {
      if (width < mins[index]) return mins[index]
      return Math.max(mins[index], width - deficit * ((width - mins[index]) / surplus))
    })
  }

  const sum = widths.reduce((total, width) => total + width, 0)
  if (sum > contentW) widths = widths.map((width) => (width * contentW) / sum)
  return widths
}

// ---------------------------------------------------------------------------
// 3. PDF editorial profesional
// ---------------------------------------------------------------------------

function layoutRich(doc, segments, maxWidth, size, baseBold = false) {
  const lines = []
  let line = []
  let width = 0
  const pushLine = () => {
    while (line.length && /^\s*$/.test(line[line.length - 1].text)) line.pop()
    lines.push(line.length ? line : [{ text: '', bold: baseBold }])
    line = []
    width = 0
  }

  for (const segment of segments) {
    const tokens = segment.text.split(/(\s+)/).filter((token) => token.length)
    for (const token of tokens) {
      doc.setFont('helvetica', segment.bold || baseBold ? 'bold' : 'normal')
      doc.setFontSize(size)
      const tokenWidth = doc.getTextWidth(token)
      if (/^\s+$/.test(token)) {
        if (line.length) {
          line.push({ text: ' ', bold: segment.bold || baseBold })
          width += tokenWidth
        }
        continue
      }
      if (width + tokenWidth > maxWidth && line.length) pushLine()
      line.push({ text: token, bold: segment.bold || baseBold })
      width += tokenWidth
    }
  }
  if (line.length) pushLine()
  return lines.length ? lines : [[{ text: '', bold: baseBold }]]
}

export function buildPdfDocument({ title = 'Documento VITIMAX', subtitle = '', content = '', dateLabel = '', catalog = [], isDemo = false, emptyState = null } = {}) {
  const doc = new jsPDF({ unit: 'mm', format: 'a4', compress: true })
  const pageW = doc.internal.pageSize.getWidth()
  const pageH = doc.internal.pageSize.getHeight()
  const margin = 18
  const contentW = pageW - margin * 2
  const bottom = pageH - 22
  const date = dateLabel || todayLabel()
  const cleanTitle = plainText(title) || 'Documento VITIMAX'
  const cleanSubtitle = stripHeaderDate(plainText(subtitle), date)
  const rawBlocks = sanitizeBlockDates(parseDocument(content), date)
  const catalogMode = Array.isArray(catalog) && catalog.length > 0
    && isCatalogDocument(cleanTitle, cleanSubtitle, rawBlocks)
  const blocks = prepareDocumentBlocks(rawBlocks, catalogMode)
  const totals = catalogMode ? computeCatalogTotals(catalog) : null
  let y = 0

  function newPage() {
    doc.addPage()
    y = margin
    doc.setFont('helvetica', 'normal')
    doc.setFontSize(7.5)
    doc.setTextColor(...rgb(BRAND.muted))
    doc.text(`VITIMAX · ${cleanTitle}`.slice(0, 95), margin, 11)
    doc.setDrawColor(...rgb(BRAND.line))
    doc.setLineWidth(0.2)
    doc.line(margin, 13, pageW - margin, 13)
    y = 20
  }

  function ensureSpace(height) {
    if (y + height > bottom) newPage()
  }

  function drawRich(segments, x, maxWidth, { size = 10.5, color = BRAND.ink, lineHeight = 5.4, bold = false, spacing = 2.4 } = {}) {
    const lines = layoutRich(doc, segments, maxWidth, size, bold)
    const height = lines.length * lineHeight
    ensureSpace(height + spacing)
    let cursorY = y
    for (const line of lines) {
      let cursorX = x
      for (const piece of line) {
        doc.setFont('helvetica', piece.bold ? 'bold' : 'normal')
        doc.setFontSize(size)
        doc.setTextColor(...rgb(color))
        doc.text(piece.text, cursorX, cursorY)
        cursorX += doc.getTextWidth(piece.text)
      }
      cursorY += lineHeight
    }
    y = cursorY + spacing
  }

  function drawHeading(block) {
    const text = plainText(block.text)
    if (!text) return
    if (block.level === 1) {
      ensureSpace(22)
      doc.setFont('helvetica', 'bold')
      doc.setFontSize(17)
      doc.setTextColor(...rgb(BRAND.forest))
      const lines = doc.splitTextToSize(text, contentW)
      lines.forEach((line) => {
        ensureSpace(8)
        doc.text(line, margin, y)
        y += 8
      })
      doc.setDrawColor(...rgb(BRAND.lime))
      doc.setLineWidth(0.8)
      doc.line(margin, y - 1.5, margin + Math.min(contentW, 46), y - 1.5)
      y += 4.5
    } else if (block.level === 2) {
      ensureSpace(16)
      doc.setFont('helvetica', 'bold')
      doc.setFontSize(13)
      doc.setTextColor(...rgb(BRAND.forest))
      const lines = doc.splitTextToSize(text, contentW)
      lines.forEach((line) => {
        ensureSpace(6.5)
        doc.text(line, margin, y)
        y += 6.5
      })
      y += 2.2
    } else {
      ensureSpace(12)
      doc.setFont('helvetica', 'bold')
      doc.setFontSize(11.2)
      doc.setTextColor(...rgb(BRAND.mid))
      const lines = doc.splitTextToSize(text, contentW)
      lines.forEach((line) => {
        ensureSpace(5.6)
        doc.text(line, margin, y)
        y += 5.6
      })
      y += 1.6
    }
  }

  function drawBullet(item, index, numbered) {
    const segments = inlineSegments(item)
    if (!segments.length) return
    const indent = 6
    const lines = layoutRich(doc, segments, contentW - indent, 10.5)
    const height = lines.length * 5.2
    ensureSpace(height + 1.6)
    doc.setFont('helvetica', 'bold')
    doc.setFontSize(10.5)
    doc.setTextColor(...rgb(BRAND.mid))
    doc.text(numbered ? `${index + 1}.` : '•', margin, y)
    let cursorY = y
    for (const line of lines) {
      let cursorX = margin + indent
      for (const piece of line) {
        doc.setFont('helvetica', piece.bold ? 'bold' : 'normal')
        doc.setFontSize(10.5)
        doc.setTextColor(...rgb(BRAND.ink))
        doc.text(piece.text, cursorX, cursorY)
        cursorX += doc.getTextWidth(piece.text)
      }
      cursorY += 5.2
    }
    y = cursorY + 1.6
  }

  function drawTable(block) {
    const header = (block.header || []).map((cell) => plainText(cell))
    const rows = (block.rows || []).map((row) => header.map((_, index) => plainText(row[index] || '')))
    if (!header.length) return
    const widths = drawTableWidths(doc, header, rows, contentW)
    const cellPad = 2
    const lineHeight = 4.1
    const rowHeight = (cells, size) => {
      doc.setFont('helvetica', 'normal')
      doc.setFontSize(size)
      const maxLines = cells.reduce((max, cell, index) => {
        const lines = doc.splitTextToSize(cell || ' ', Math.max(8, widths[index] - cellPad * 2))
        return Math.max(max, lines.length)
      }, 1)
      return maxLines * lineHeight + cellPad * 2
    }

    const drawRow = (cells, options) => {
      const height = rowHeight(cells, options.size)
      ensureSpace(height + 1)
      let cursorX = margin
      cells.forEach((cell, index) => {
        const width = widths[index]
        doc.setFillColor(...rgb(options.fill || BRAND.white))
        doc.rect(cursorX, y, width, height, 'F')
        doc.setDrawColor(...rgb(BRAND.line))
        doc.setLineWidth(0.2)
        doc.rect(cursorX, y, width, height, 'S')
        doc.setFont('helvetica', options.bold ? 'bold' : 'normal')
        doc.setFontSize(options.size)
        doc.setTextColor(...rgb(options.color || BRAND.ink))
        const lines = doc.splitTextToSize(cell || ' ', Math.max(8, width - cellPad * 2))
        let textY = y + cellPad + lineHeight - 0.6
        lines.forEach((line) => {
          doc.text(line, cursorX + cellPad, textY)
          textY += lineHeight
        })
        cursorX += width
      })
      y += height
    }

    const headerHeight = rowHeight(header, 8.8)
    ensureSpace(headerHeight + 6)
    drawRow(header, { fill: BRAND.forest, color: BRAND.white, bold: true, size: 8.8 })
    rows.forEach((row, index) => {
      drawRow(row, { fill: index % 2 ? BRAND.zebra : BRAND.white, size: 8.6 })
    })
    y += 4
  }

  function drawCode(block) {
    const code = String(block.code || '').replace(/\t/g, '  ')
    const lines = code.split('\n')
    doc.setFont('courier', 'normal')
    doc.setFontSize(8.6)
    const wrapped = []
    lines.forEach((line) => {
      const parts = doc.splitTextToSize(line || ' ', contentW - 8)
      parts.forEach((part) => wrapped.push(part))
    })
    const height = wrapped.length * 4 + 5
    ensureSpace(height)
    doc.setFillColor(...rgb(BRAND.zebra))
    doc.rect(margin, y, contentW, height, 'F')
    doc.setDrawColor(...rgb(BRAND.line))
    doc.setLineWidth(0.2)
    doc.rect(margin, y, contentW, height, 'S')
    let cursorY = y + 5
    doc.setTextColor(...rgb(BRAND.deep))
    wrapped.forEach((line) => {
      doc.text(line, margin + 4, cursorY)
      cursorY += 4
    })
    y += height + 3
  }

  function drawSectionLabel(label) {
    ensureSpace(10)
    doc.setFont('helvetica', 'bold')
    doc.setFontSize(8)
    doc.setTextColor(...rgb(BRAND.mid))
    doc.text(String(label || '').toUpperCase(), margin, y)
    y += 4.6
  }

  function drawKpiCards(items) {
    if (!items.length) return
    const gap = 3.5
    const cardW = (contentW - gap * (items.length - 1)) / items.length
    const cardH = 17
    ensureSpace(cardH + 5)
    items.forEach((item, index) => {
      const x = margin + index * (cardW + gap)
      doc.setFillColor(...rgb(BRAND.zebra))
      doc.rect(x, y, cardW, cardH, 'F')
      doc.setFillColor(...rgb(BRAND.forest))
      doc.rect(x, y, 1.1, cardH, 'F')
      doc.setFont('helvetica', 'bold')
      doc.setFontSize(6.8)
      doc.setTextColor(...rgb(BRAND.muted))
      doc.text(String(item.label || '').toUpperCase(), x + 3.6, y + 4.8)
      doc.setFont('helvetica', 'bold')
      doc.setFontSize(12)
      doc.setTextColor(...rgb(BRAND.deep))
      doc.text(String(item.value || ''), x + 3.6, y + 12.4, { maxWidth: cardW - 5 })
    })
    y += cardH + 5
  }

  function drawTotalsSection(catalogTotals) {
    if (!catalogTotals) return
    ensureSpace(46)
    drawSectionLabel('Totales del catálogo')
    drawTable({
      header: ['Concepto', 'Valor'],
      rows: [
        ['Productos registrados', String(catalogTotals.count)],
        ['Unidades en stock', String(catalogTotals.units)],
        ['Valor de inventario (Precio × Stock)', formatClp(catalogTotals.inventoryValue)],
        ['Suma de precios de lista', formatClp(catalogTotals.priceSum)],
        ['Precio promedio', formatClp(catalogTotals.avgPrice)]
      ]
    })
  }

  // Portada / banner corporativo en la primera página
  doc.setFont('helvetica', 'bold')
  doc.setFontSize(22)
  const titleLines = doc.splitTextToSize(cleanTitle, contentW).slice(0, 4)
  const bannerH = 44 + Math.max(0, titleLines.length - 1) * 8.4
  doc.setFillColor(...rgb(BRAND.forest))
  doc.rect(0, 0, pageW, bannerH, 'F')
  doc.setFillColor(...rgb(BRAND.mid))
  doc.rect(pageW * 0.66, 0, pageW * 0.34, bannerH, 'F')
  doc.setFillColor(...rgb(BRAND.lime))
  doc.rect(0, bannerH, pageW, 1.6, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(9)
  doc.setTextColor(...rgb(BRAND.white))
  doc.text('VITIMAX', margin, 14)
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(7.5)
  doc.setTextColor(...rgb(BRAND.lime))
  doc.text('CENTRO DE GESTIÓN', margin + 20, 14)

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(titleLines.length > 2 ? 20 : 24)
  doc.setTextColor(...rgb(BRAND.white))
  let titleY = 26
  titleLines.forEach((line) => {
    doc.text(line, margin, titleY)
    titleY += 8.4
  })

  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.setTextColor(...rgb(BRAND.totals))
  doc.text(`${date}${cleanSubtitle ? ` · ${cleanSubtitle}` : ''}`.slice(0, 120), margin, bannerH - 6)

  y = bannerH + 12

  // Catálogo vacío: tarjetas con los conteos reales en lugar de datos de ejemplo.
  if (emptyState && !catalogMode) {
    const safeCount = (value) => Math.max(0, Number(value) || 0)
    drawSectionLabel('Estado inicial')
    drawKpiCards([
      { label: 'Productos', value: String(safeCount(emptyState.productos)) },
      { label: 'Clientes', value: String(safeCount(emptyState.clientes)) },
      { label: 'Ventas', value: String(safeCount(emptyState.ventas)) }
    ])
  }

  // La página 1 nunca queda vacía: banner + KPIs reales + primera tabla.
  if (catalogMode && totals) {
    drawSectionLabel('Resumen ejecutivo')
    drawKpiCards([
      { label: 'Productos', value: String(totals.count) },
      { label: 'Unidades en stock', value: String(totals.units) },
      { label: 'Valor inventario', value: formatClp(totals.inventoryValue) },
      { label: 'Precio promedio', value: formatClp(totals.avgPrice) }
    ])
  }

  const skipTitle = blocks.length && blocks[0].type === 'heading' && blocks[0].level === 1
    && normalizeTitle(blocks[0].text) === normalizeTitle(cleanTitle)
  let startIndex = skipTitle ? 1 : 0
  // Evita duplicar el "Resumen Ejecutivo" que ya se dibujó como tarjetas.
  if (catalogMode && blocks[startIndex]?.type === 'heading' && /resumen ejecutivo/i.test(plainText(blocks[startIndex].text))) {
    startIndex += 1
  }

  for (let index = startIndex; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'heading') drawHeading(block)
    else if (block.type === 'paragraph') drawRich(inlineSegments(block.text), margin, contentW)
    else if (block.type === 'quote') drawRich(inlineSegments(block.text), margin + 2, contentW - 4, { color: BRAND.deep, size: 10.2, bold: true })
    else if (block.type === 'bullets') block.items.forEach((item, i) => drawBullet(item, i, false))
    else if (block.type === 'numbers') block.items.forEach((item, i) => drawBullet(item, i, true))
    else if (block.type === 'table') drawTable(block)
    else if (block.type === 'code') drawCode(block)
    else if (block.type === 'divider') {
      ensureSpace(8)
      doc.setDrawColor(...rgb(BRAND.line))
      doc.setLineWidth(0.3)
      doc.line(margin, y, pageW - margin, y)
      y += 5
    }
  }

  // Totales calculados desde los datos reales de la app, nunca desde la IA.
  if (catalogMode && totals) drawTotalsSection(totals)

  // Pie de página en todas las hojas
  const total = doc.getNumberOfPages()
  for (let page = 1; page <= total; page += 1) {
    doc.setPage(page)
    doc.setDrawColor(...rgb(BRAND.line))
    doc.setLineWidth(0.2)
    doc.line(margin, pageH - 16, pageW - margin, pageH - 16)
    doc.setFont('helvetica', 'normal')
    doc.setFontSize(7.5)
    doc.setTextColor(...rgb(BRAND.muted))
    doc.text(`VITIMAX · Centro de Gestión · Página ${page} de ${total}`, pageW / 2, pageH - 11, { align: 'center' })
  }

  return doc.output('blob')
}

// ---------------------------------------------------------------------------
// 4. Word ejecutivo formal
// ---------------------------------------------------------------------------

function docxRuns(value, { size = 22, color = BRAND.ink, bold = false } = {}) {
  const segments = inlineSegments(value)
  if (!segments.length) return [new TextRun({ text: '', size, color })]
  return segments.map((segment) => new TextRun({
    text: segment.text,
    bold: bold || segment.bold,
    size,
    color,
    font: 'Calibri'
  }))
}

function docxTableCell(value, { header = false, zebra = false, width = 0 } = {}) {
  return new TableCell({
    shading: {
      type: ShadingType.CLEAR,
      color: 'auto',
      fill: header ? BRAND.forest : zebra ? BRAND.zebra : BRAND.white
    },
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 90, bottom: 90, left: 120, right: 120 },
    children: [new Paragraph({
      spacing: { after: 0 },
      children: docxRuns(value, {
        size: header ? 19 : 19,
        color: header ? BRAND.white : BRAND.ink,
        bold: header
      })
    })]
  })
}

function docxKpiCell(label, value) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: BRAND.totals },
    width: { size: 25, type: WidthType.PERCENTAGE },
    margins: { top: 150, bottom: 150, left: 160, right: 160 },
    children: [
      new Paragraph({
        spacing: { after: 50 },
        children: [new TextRun({ text: String(label || '').toUpperCase(), bold: true, size: 14, color: BRAND.mid, characterSpacing: 20 })]
      }),
      new Paragraph({
        spacing: { after: 0 },
        children: [new TextRun({ text: String(value || ''), bold: true, size: 30, color: BRAND.deep })]
      })
    ]
  })
}

function docxKpiTable(totals) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
      left: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
      right: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line }
    },
    rows: [new TableRow({
      children: [
        docxKpiCell('Productos', totals.count),
        docxKpiCell('Unidades en stock', totals.units),
        docxKpiCell('Valor inventario', formatClp(totals.inventoryValue)),
        docxKpiCell('Precio promedio', formatClp(totals.avgPrice))
      ]
    })]
  })
}

function docxTableWidths(header, rows) {
  const weights = header.map((label, index) => {
    const key = plainText(label).toLowerCase()
    let base = 16
    if (/^(id|c[oó]digo|sku|n[°º])$/.test(key)) base = 8
    else if (/precio|valor|monto|total|stock|cantidad|unidades|iva|clp/.test(key)) base = 14
    else if (/nombre|producto/.test(key)) base = 24
    else if (/descripci[oó]n|dosis|uso|modo|recomendaci|caracter|detalle/.test(key)) base = 30
    const longest = rows.reduce((max, row) => Math.max(max, String(row[index] ?? '').length), 0)
    return Math.max(base, Math.min(38, longest * 0.7 + 6))
  })
  const total = weights.reduce((sum, weight) => sum + weight, 0) || 1
  const widths = weights.map((weight) => Math.max(7, Math.round((weight / total) * 100)))
  const sum = widths.reduce((acc, width) => acc + width, 0)
  widths[widths.length - 1] += 100 - sum
  return widths
}

export function buildDocxDocument({ title = 'Documento VITIMAX', subtitle = '', content = '', dateLabel = '', catalog = [], isDemo = false } = {}) {
  const cleanTitle = plainText(title) || 'Documento VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const rawBlocks = parseDocument(content)
  const catalogMode = Array.isArray(catalog) && catalog.length > 0
    && isCatalogDocument(cleanTitle, cleanSubtitle, rawBlocks)
  const blocks = prepareDocumentBlocks(rawBlocks, catalogMode)
  const totals = catalogMode ? computeCatalogTotals(catalog) : null
  const date = dateLabel || todayLabel()
  const children = []

  // Membrete / portada corporativa
  children.push(new Paragraph({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: BRAND.forest },
    indent: { left: 200, right: 200 },
    spacing: { before: 0, after: 160 },
    children: [
      new TextRun({ text: 'VITIMAX', bold: true, size: 22, color: BRAND.white, characterSpacing: 80 }),
      new TextRun({ text: '   CENTRO DE GESTIÓN', size: 16, color: BRAND.lime, characterSpacing: 40 })
    ]
  }))
  children.push(new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text: cleanTitle, bold: true, size: 54, color: BRAND.deep })]
  }))
  children.push(new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 14, color: BRAND.lime, space: 8 } },
    spacing: { after: 200 },
    children: [new TextRun({ text: cleanSubtitle || 'Documento ejecutivo', size: 24, color: BRAND.muted })]
  }))
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: date, size: 19, color: BRAND.muted })]
  }))
  children.push(new Paragraph({ children: [new PageBreak()] }))

  // Tarjetas de resumen ejecutivo con datos reales del catálogo.
  if (catalogMode && totals) {
    children.push(new Paragraph({
      spacing: { after: 140 },
      children: [new TextRun({ text: 'Resumen ejecutivo', bold: true, size: 30, color: BRAND.forest })]
    }))
    children.push(docxKpiTable(totals))
    children.push(new Paragraph({ spacing: { after: 200 }, children: [] }))
  }

  const skipTitle = blocks.length && blocks[0].type === 'heading' && blocks[0].level === 1
    && normalizeTitle(blocks[0].text) === normalizeTitle(cleanTitle)
  let startIndex = skipTitle ? 1 : 0
  if (catalogMode && blocks[startIndex]?.type === 'heading' && /resumen ejecutivo/i.test(plainText(blocks[startIndex].text))) {
    startIndex += 1
  }

  for (let index = startIndex; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'heading') {
      const level = block.level === 1 ? HeadingLevel.HEADING_1 : block.level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3
      const size = block.level === 1 ? 34 : block.level === 2 ? 27 : 23
      const color = block.level === 3 ? BRAND.mid : BRAND.forest
      children.push(new Paragraph({
        heading: level,
        spacing: { before: block.level === 1 ? 320 : 240, after: 120 },
        border: block.level === 1
          ? { bottom: { style: BorderStyle.SINGLE, size: 10, color: BRAND.lime, space: 6 } }
          : undefined,
        children: docxRuns(block.text, { size, color, bold: true })
      }))
      continue
    }
    if (block.type === 'paragraph') {
      children.push(new Paragraph({
        spacing: { after: 140, line: 300 },
        alignment: AlignmentType.JUSTIFIED,
        children: docxRuns(block.text, { size: 22 })
      }))
      continue
    }
    if (block.type === 'quote') {
      children.push(new Paragraph({
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: BRAND.totals },
        border: { left: { style: BorderStyle.SINGLE, size: 26, color: BRAND.forest, space: 12 } },
        indent: { left: 260, right: 140 },
        spacing: { before: 140, after: 180, line: 290 },
        children: docxRuns(block.text, { size: 21, color: BRAND.deep, bold: true })
      }))
      continue
    }
    if (block.type === 'bullets') {
      block.items.forEach((item) => children.push(new Paragraph({
        bullet: { level: 0 },
        spacing: { after: 70, line: 290 },
        children: docxRuns(item, { size: 22 })
      })))
      continue
    }
    if (block.type === 'numbers') {
      block.items.forEach((item, itemIndex) => children.push(new Paragraph({
        indent: { left: 460, hanging: 280 },
        spacing: { after: 70, line: 290 },
        children: [
          new TextRun({ text: `${itemIndex + 1}. `, bold: true, size: 22, color: BRAND.mid }),
          ...docxRuns(item, { size: 22 })
        ]
      })))
      continue
    }
    if (block.type === 'table') {
      const header = block.header || []
      const rows = block.rows || []
      if (!header.length) continue
      const widths = docxTableWidths(header, rows)
      children.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.AUTOFIT,
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
          left: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
          right: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
          insideVertical: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line }
        },
        rows: [
          new TableRow({
            tableHeader: true,
            children: header.map((cell, cellIndex) => docxTableCell(cell, { header: true, width: widths[cellIndex] }))
          }),
          ...rows.map((row, rowIndex) => new TableRow({
            children: header.map((_, cellIndex) => docxTableCell(row[cellIndex] || '', { zebra: rowIndex % 2 === 1, width: widths[cellIndex] }))
          }))
        ]
      }))
      children.push(new Paragraph({ spacing: { after: 180 }, children: [] }))
      continue
    }
    if (block.type === 'code') {
      const lines = String(block.code || '').split('\n')
      children.push(new Paragraph({
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'F1F5F1' },
        border: { left: { style: BorderStyle.SINGLE, size: 18, color: BRAND.mid, space: 10 } },
        indent: { left: 240 },
        spacing: { before: 120, after: 160 },
        children: lines.map((line, lineIndex) => new TextRun({
          text: `${line}${lineIndex < lines.length - 1 ? '\n' : ''}`,
          font: 'Consolas',
          size: 17,
          color: BRAND.deep
        }))
      }))
      continue
    }
    if (block.type === 'divider') {
      children.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND.line, space: 6 } },
        spacing: { before: 140, after: 140 },
        children: []
      }))
    }
  }

  // Totales reales calculados desde el catálogo de la app.
  if (catalogMode && totals) {
    children.push(new Paragraph({
      spacing: { before: 320, after: 140 },
      children: [new TextRun({ text: 'Totales del catálogo', bold: true, size: 27, color: BRAND.forest })]
    }))
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
        left: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
        right: { style: BorderStyle.SINGLE, size: 4, color: BRAND.line },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line },
        insideVertical: { style: BorderStyle.SINGLE, size: 2, color: BRAND.line }
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            docxTableCell('Concepto', { header: true, width: 34 }),
            docxTableCell('Valor', { header: true, width: 66 })
          ]
        }),
        ...[
          ['Productos registrados', String(totals.count)],
          ['Unidades en stock', String(totals.units)],
          ['Valor de inventario (Precio × Stock)', formatClp(totals.inventoryValue)],
          ['Suma de precios de lista', formatClp(totals.priceSum)],
          ['Precio promedio', formatClp(totals.avgPrice)]
        ].map((row, rowIndex) => new TableRow({
          children: [
            docxTableCell(row[0], { zebra: rowIndex % 2 === 1, width: 34 }),
            docxTableCell(row[1], { zebra: rowIndex % 2 === 1, width: 66 })
          ]
        }))
      ]
    }))
  }

  const footer = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({ text: 'VITIMAX · Centro de Gestión · Página ', size: 16, color: BRAND.muted }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: BRAND.muted }),
        new TextRun({ text: ' de ', size: 16, color: BRAND.muted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: BRAND.muted })
      ]
    })]
  })

  const document = new Document({
    creator: 'VITIMAX',
    title: cleanTitle,
    description: cleanSubtitle,
    styles: {
      default: {
        document: {
          run: { font: 'Calibri', size: 22, color: BRAND.ink },
          paragraph: { spacing: { line: 300 } }
        }
      }
    },
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 900, left: 1100, right: 1100 } } },
      footers: { default: footer },
      children
    }]
  })

  return Packer.toBlob(document)
}

// ---------------------------------------------------------------------------
// 5. PowerPoint 16:9 de nivel presentación
// ---------------------------------------------------------------------------

function deckFromBlocks(blocks, deckTitle) {
  const slides = []
  let current = null
  let started = false

  const push = () => {
    if (current && (current.bullets.length || current.paragraphs.length || current.table)) slides.push(current)
    current = null
  }

  for (const block of blocks) {
    if (block.type === 'heading') {
      const text = plainText(block.text)
      if (!text) continue
      if (!started && block.level === 1 && normalizeTitle(text) === normalizeTitle(deckTitle)) {
        started = true
        continue
      }
      if (block.level <= 2) {
        push()
        current = { title: text, bullets: [], paragraphs: [], table: null }
        started = true
        continue
      }
      if (!current) current = { title: deckTitle || 'VITIMAX', bullets: [], paragraphs: [], table: null }
      current.paragraphs.push({ text, bold: true })
      continue
    }
    if (!current) current = { title: deckTitle || 'VITIMAX', bullets: [], paragraphs: [], table: null }
    if (block.type === 'bullets') current.bullets.push(...block.items.map((item) => plainText(item)).filter(Boolean))
    else if (block.type === 'numbers') current.bullets.push(...block.items.map((item, index) => `${index + 1}. ${plainText(item)}`).filter(Boolean))
    else if (block.type === 'paragraph') {
      const text = plainText(block.text)
      if (text) current.paragraphs.push({ text })
    } else if (block.type === 'quote') {
      const text = plainText(block.text)
      if (text) current.paragraphs.push({ text, bold: true })
    } else if (block.type === 'table') {
      current.table = {
        header: (block.header || []).map((cell) => plainText(cell)),
        rows: (block.rows || []).map((row) => (block.header || []).map((_, index) => plainText(row[index] || '')))
      }
    }
  }
  push()
  return slides
}

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

function bytesFromBase64(base64) {
  const clean = String(base64 || '').replace(/^data:[^;]+;base64,/i, '').replace(/\s+/g, '')
  const binary = atob(clean)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

// pptxgenjs puede devolver Blob, ArrayBuffer, Uint8Array o base64 según el
// entorno (navegador, Android WebView, Node). Siempre se entrega un Blob con
// el MIME correcto de PowerPoint, sin rutas exclusivas de Node.
async function pptxWriteBlob(pptx) {
  let lastError = null
  for (const outputType of ['blob', 'arraybuffer', 'base64']) {
    try {
      const result = await pptx.write({ outputType })
      if (!result) continue
      if (typeof Blob !== 'undefined' && result instanceof Blob) {
        return new Blob([result], { type: PPTX_MIME })
      }
      if (result instanceof ArrayBuffer) {
        return new Blob([new Uint8Array(result)], { type: PPTX_MIME })
      }
      if (ArrayBuffer.isView(result)) {
        return new Blob([result], { type: PPTX_MIME })
      }
      if (typeof result === 'string') {
        return new Blob([bytesFromBase64(result)], { type: PPTX_MIME })
      }
    } catch (error) {
      lastError = error
    }
  }
  throw lastError || new Error('No se pudo generar el archivo PowerPoint.')
}

export async function buildPptxDocument({ title = 'Presentación VITIMAX', subtitle = '', content = '', dateLabel = '', catalog = [], isDemo = false } = {}) {
  const deckTitle = plainText(title) || 'Presentación VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const date = dateLabel || todayLabel()
  const rawBlocks = parseDocument(content)
  const catalogMode = Array.isArray(catalog) && catalog.length > 0
    && isCatalogDocument(deckTitle, cleanSubtitle, rawBlocks)
  const blocks = prepareDocumentBlocks(rawBlocks, catalogMode)
  const totals = catalogMode ? computeCatalogTotals(catalog) : null
  let slides = deckFromBlocks(blocks, deckTitle)

  if (!slides.length) {
    const fallback = blocks.filter((block) => block.type === 'paragraph' || block.type === 'bullets')
    slides = [{
      title: deckTitle,
      bullets: fallback.flatMap((block) => block.type === 'bullets' ? block.items.map(plainText) : [plainText(block.text)]).filter(Boolean).slice(0, 5),
      paragraphs: [],
      table: null
    }]
  }

  // Máximo 5 viñetas por diapositiva para que el texto nunca se desborde.
  const expanded = []
  slides.forEach((slide) => {
    if (slide.bullets.length > 5) {
      for (let index = 0; index < slide.bullets.length; index += 5) {
        expanded.push({
          ...slide,
          title: index === 0 ? slide.title : `${slide.title} (${Math.floor(index / 5) + 1})`,
          bullets: slide.bullets.slice(index, index + 5),
          paragraphs: index === 0 ? slide.paragraphs : [],
          table: index === 0 ? slide.table : null
        })
      }
    } else {
      expanded.push(slide)
    }
  })

  // Tablas limpias del catálogo: columnas exactas y máximo 5 filas por slide.
  // Se eliminan las tablas genéricas para no repetir datos ni columnas raras.
  if (catalogMode && Array.isArray(catalog) && catalog.length) {
    const junkName = /^(?:gran\s+)?(?:total(?:es)?|subtotal(?:es)?|suma|resumen|valor(?:izado)?)(?:\s+(?:general|valorado|inventario|productos|stock|clp|ejecutivo|de|del|la|el))*\s*$/i
    const productRows = catalog
      .filter((item) => item && !junkName.test(String(item.name ?? item.nombre ?? item.producto ?? '').trim()))
      .map((item) => [
        plainText(item.name ?? item.nombre ?? item.producto ?? ''),
        plainText(item.category ?? item.categoria ?? ''),
        plainText(item.format ?? item.formato ?? ''),
        formatClp(item.price ?? item.precio ?? 0),
        plainText(item.description ?? item.descripcion ?? '').slice(0, 120)
      ])
    const productSlides = []
    for (let index = 0; index < productRows.length; index += 5) {
      productSlides.push({
        title: `Productos VITIMAX · ${Math.floor(index / 5) + 1}`,
        bullets: [],
        paragraphs: [],
        table: {
          header: ['Nombre', 'Categoría', 'Formato', 'Precio', 'Descripción'],
          rows: productRows.slice(index, index + 5)
        }
      })
    }
    const withoutGenericTables = expanded.filter((slide) => !slide.table)
    expanded.length = 0
    expanded.push(...withoutGenericTables, ...productSlides)
  }

  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'VITIMAX'
  pptx.company = 'VITIMAX'
  pptx.subject = deckTitle
  pptx.title = deckTitle
  pptx.theme = { headFontFace: 'Segoe UI Semibold', bodyFontFace: 'Segoe UI', lang: 'es-CL' }

  const coverSubtitle = cleanSubtitle
    || expanded[0]?.paragraphs?.[0]?.text
    || 'Centro de Gestión · Informe ejecutivo'

  // Portada
  const cover = pptx.addSlide()
  cover.background = { color: BRAND.deep }
  cover.addShape(pptx.ShapeType.ellipse, {
    x: 8.9, y: -2.1, w: 7.2, h: 7.2,
    fill: { color: BRAND.forest }, line: { color: BRAND.forest }
  })
  cover.addShape(pptx.ShapeType.rect, {
    x: 0, y: 6.95, w: 13.333, h: 0.55,
    fill: { color: BRAND.forest }, line: { color: BRAND.forest }
  })
  cover.addText('VITIMAX', {
    x: 0.85, y: 0.62, w: 6, h: 0.45,
    fontSize: 16, bold: true, color: BRAND.lime, charSpacing: 5
  })
  cover.addText(deckTitle, {
    x: 0.85, y: 1.95, w: 10.9, h: 2.1,
    fontSize: 40, bold: true, color: BRAND.white, valign: 'top', fit: 'shrink'
  })
  cover.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 4.2, w: 1.6, h: 0.09,
    fill: { color: BRAND.lime }, line: { color: BRAND.lime }
  })
  cover.addText(coverSubtitle.slice(0, 180), {
    x: 0.85, y: 4.5, w: 9.4, h: 1,
    fontSize: 16, color: 'CFE3D4', fit: 'shrink'
  })
  cover.addText(date, {
    x: 0.85, y: 6.5, w: 7, h: 0.35,
    fontSize: 10, color: 'A9C6B4'
  })
  cover.addText('Centro de Gestión', {
    x: 9.4, y: 6.5, w: 3.1, h: 0.35,
    fontSize: 10, color: 'A9C6B4', align: 'right'
  })

  // Quita el prefijo "1." / "1)" cuando la propia diapositiva ya dibuja el número.
  const stripOrder = (text) => String(text || '').replace(/^\s*\d{1,3}[.)]\s+/, '')

  // La portada es la página 1; el numerador se lleva aquí para que el orden
  // siga siendo correcto aunque se inserten tablas de catálogo.
  let pageCounter = 2
  const addFooter = (slide) => {
    slide.addText('VITIMAX · Centro de Gestión', {
      x: 0.6, y: 7.12, w: 5, h: 0.25,
      fontSize: 8, color: '8AA79A'
    })
    slide.addText(String(pageCounter), {
      x: 12.1, y: 7.12, w: 0.7, h: 0.25,
      fontSize: 8, color: '8AA79A', align: 'right'
    })
    pageCounter += 1
  }

  // Tarjetas de resumen ejecutivo con los datos reales del catálogo.
  if (catalogMode && totals) {
    const summary = pptx.addSlide()
    summary.background = { color: 'F7FAF7' }
    summary.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.333, h: 0.42,
      fill: { color: BRAND.forest }, line: { color: BRAND.forest }
    })
    summary.addText('VITIMAX', {
      x: 0.55, y: 0.06, w: 3, h: 0.3,
      fontSize: 9, bold: true, color: BRAND.white, charSpacing: 3
    })
    summary.addText('Resumen Ejecutivo', {
      x: 0.6, y: 0.75, w: 12.1, h: 0.8,
      fontSize: 28, bold: true, color: BRAND.deep, valign: 'top', fit: 'shrink'
    })
    summary.addShape(pptx.ShapeType.rect, {
      x: 0.62, y: 1.62, w: 1.15, h: 0.07,
      fill: { color: BRAND.lime }, line: { color: BRAND.lime }
    })
    const cards = [
      { label: 'PRODUCTOS', value: String(totals.count) },
      { label: 'UNIDADES EN STOCK', value: String(totals.units) },
      { label: 'VALOR INVENTARIO', value: formatClp(totals.inventoryValue) },
      { label: 'PRECIO PROMEDIO', value: formatClp(totals.avgPrice) }
    ]
    const gap = 0.32
    const cardW = (12.1 - gap * (cards.length - 1)) / cards.length
    cards.forEach((card, index) => {
      const x = 0.62 + index * (cardW + gap)
      summary.addShape(pptx.ShapeType.roundRect, {
        x, y: 2.15, w: cardW, h: 2.35,
        fill: { color: BRAND.white },
        line: { color: BRAND.line, width: 0.75 },
        rectRadius: 0.06
      })
      summary.addShape(pptx.ShapeType.rect, {
        x, y: 2.15, w: cardW, h: 0.12,
        fill: { color: BRAND.mid }, line: { color: BRAND.mid }
      })
      summary.addText(card.label, {
        x: x + 0.28, y: 2.55, w: cardW - 0.56, h: 0.4,
        fontSize: 9, bold: true, color: BRAND.muted, charSpacing: 1
      })
      summary.addText(card.value, {
        x: x + 0.28, y: 3.05, w: cardW - 0.56, h: 0.9,
        fontSize: 25, bold: true, color: BRAND.deep, valign: 'top', fit: 'shrink'
      })
    })
    summary.addText('Datos calculados desde el catálogo real de VITIMAX.', {
      x: 0.62, y: 5.0, w: 12.1, h: 0.4,
      fontSize: 11, color: BRAND.muted
    })
    addFooter(summary, 2)
  }

  expanded.forEach((item, index) => {
    const slide = pptx.addSlide()
    slide.background = { color: 'F7FAF7' }
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.333, h: 0.42,
      fill: { color: BRAND.forest }, line: { color: BRAND.forest }
    })
    slide.addText('VITIMAX', {
      x: 0.55, y: 0.06, w: 3, h: 0.3,
      fontSize: 9, bold: true, color: BRAND.white, charSpacing: 3
    })
    slide.addText('Centro de Gestión', {
      x: 9.7, y: 0.06, w: 3.1, h: 0.3,
      fontSize: 8.5, color: 'CFE3D4', align: 'right'
    })
    slide.addText(item.title, {
      x: 0.6, y: 0.68, w: 12.1, h: 0.95,
      fontSize: 26, bold: true, color: BRAND.deep, valign: 'top', fit: 'shrink'
    })
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.62, y: 1.6, w: 1.15, h: 0.07,
      fill: { color: BRAND.lime }, line: { color: BRAND.lime }
    })

    let bodyY = 1.9
    // La introducción solo se dibuja si después vienen viñetas o una tabla.
    // Si la diapositiva es solo texto, se dibuja una única vez más abajo.
    const hasTable = Boolean(item.table && item.table.header.length)
    if (item.paragraphs.length && (item.bullets.length || hasTable)) {
      const intro = item.paragraphs.slice(0, 3).map((paragraph) => paragraph.text).join('\n')
      slide.addText(intro, {
        x: 0.62, y: bodyY, w: 12.1, h: 0.75,
        fontSize: 13, color: BRAND.muted, fit: 'shrink', valign: 'top'
      })
      bodyY += 0.85
    }

    if (hasTable) {
      const header = item.table.header
      const rows = item.table.rows.slice(0, 5)
      const tableRows = [
        header.map((cell) => ({
          text: cell,
          options: { fill: { color: BRAND.forest }, color: BRAND.white, bold: true, fontSize: 11 }
        })),
        ...rows.map((row, rowIndex) => row.map((cell) => ({
          text: cell,
          options: {
            fill: { color: rowIndex % 2 ? BRAND.zebra : BRAND.white },
            color: BRAND.ink,
            fontSize: 10.5
          }
        })))
      ]
      const weights = header.map((label, columnIndex) => {
        const key = plainText(label).toLowerCase()
        if (/^(id|c[oó]digo|sku|n[°º])$/.test(key)) return 6
        if (/precio|valor|monto|total|stock|cantidad|unidades|iva|clp/.test(key)) return 13
        if (/nombre|producto/.test(key)) return 26
        if (/descripci[oó]n|dosis|uso|modo|recomendaci|caracter|detalle/.test(key)) return 32
        const longest = rows.reduce((max, row) => Math.max(max, String(row[columnIndex] ?? '').length), 0)
        return Math.max(14, Math.min(34, longest * 0.7 + 6))
      })
      const totalWeight = weights.reduce((sum, weight) => sum + weight, 0) || 1
      const colW = weights.map((weight) => (weight / totalWeight) * 12.1)
      slide.addTable(tableRows, {
        x: 0.62,
        y: bodyY,
        w: 12.1,
        colW,
        border: { type: 'solid', color: BRAND.line, pt: 1 },
        align: 'left',
        valign: 'middle',
        margin: 6,
        autoPage: false
      })
      addFooter(slide, index + (catalogMode ? 3 : 2))
      return
    }

    const bullets = item.bullets.slice(0, 5)
    if (!bullets.length) {
      slide.addText(item.paragraphs.map((paragraph) => paragraph.text).join('\n') || 'Contenido de la presentación.', {
        x: 0.7, y: bodyY + 0.2, w: 11.9, h: 4.4,
        fontSize: 18, color: BRAND.ink, fit: 'shrink', valign: 'top'
      })
      addFooter(slide, index + (catalogMode ? 3 : 2))
      return
    }

    if (bullets.length <= 4) {
      const cols = bullets.length === 1 ? 1 : 2
      const rows = Math.ceil(bullets.length / cols)
      const gap = 0.32
      const availableH = 6.75 - bodyY
      const cardH = Math.min(2.05, (availableH - gap * (rows - 1)) / rows)
      const cardW = (12.1 - gap * (cols - 1)) / cols
      bullets.forEach((bullet, bulletIndex) => {
        const col = bulletIndex % cols
        const row = Math.floor(bulletIndex / cols)
        const x = 0.62 + col * (cardW + gap)
        const y = bodyY + row * (cardH + gap)
        slide.addShape(pptx.ShapeType.roundRect, {
          x, y, w: cardW, h: cardH,
          fill: { color: BRAND.white },
          line: { color: BRAND.line, width: 0.75 },
          rectRadius: 0.06
        })
        slide.addShape(pptx.ShapeType.rect, {
          x, y, w: 0.1, h: cardH,
          fill: { color: BRAND.mid }, line: { color: BRAND.mid }
        })
        slide.addShape(pptx.ShapeType.ellipse, {
          x: x + 0.28, y: y + 0.28, w: 0.5, h: 0.5,
          fill: { color: BRAND.lime }, line: { color: BRAND.lime }
        })
        slide.addText(String(bulletIndex + 1), {
          x: x + 0.28, y: y + 0.28, w: 0.5, h: 0.5,
          fontSize: 12, bold: true, color: BRAND.deep, align: 'center', valign: 'middle'
        })
        slide.addText(stripOrder(bullet), {
          x: x + 0.28, y: y + 0.9, w: cardW - 0.55, h: cardH - 1.05,
          fontSize: 13.5, color: BRAND.ink, valign: 'top', fit: 'shrink'
        })
      })
    } else {
      const startY = bodyY + 0.1
      const rowH = Math.min(0.86, (6.75 - startY) / bullets.length)
      bullets.forEach((bullet, bulletIndex) => {
        const y = startY + bulletIndex * rowH
        slide.addShape(pptx.ShapeType.ellipse, {
          x: 0.68, y: y + 0.18, w: 0.2, h: 0.2,
          fill: { color: BRAND.mid }, line: { color: BRAND.mid }
        })
        slide.addText(bullet, {
          x: 1.02, y, w: 11.5, h: rowH - 0.08,
          fontSize: 14, color: BRAND.ink, valign: 'middle', fit: 'shrink'
        })
      })
    }
    addFooter(slide, index + (catalogMode ? 3 : 2))
  })

  // Totales reales del catálogo antes del cierre.
  if (catalogMode && totals) {
    const totalsSlide = pptx.addSlide()
    totalsSlide.background = { color: 'F7FAF7' }
    totalsSlide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 13.333, h: 0.42,
      fill: { color: BRAND.forest }, line: { color: BRAND.forest }
    })
    totalsSlide.addText('VITIMAX', {
      x: 0.55, y: 0.06, w: 3, h: 0.3,
      fontSize: 9, bold: true, color: BRAND.white, charSpacing: 3
    })
    totalsSlide.addText('Totales del catálogo', {
      x: 0.6, y: 0.75, w: 12.1, h: 0.8,
      fontSize: 28, bold: true, color: BRAND.deep, valign: 'top', fit: 'shrink'
    })
    totalsSlide.addShape(pptx.ShapeType.rect, {
      x: 0.62, y: 1.62, w: 1.15, h: 0.07,
      fill: { color: BRAND.lime }, line: { color: BRAND.lime }
    })
    const rows = [
      ['Productos registrados', String(totals.count)],
      ['Unidades en stock', String(totals.units)],
      ['Valor de inventario (Precio × Stock)', formatClp(totals.inventoryValue)],
      ['Suma de precios de lista', formatClp(totals.priceSum)],
      ['Precio promedio', formatClp(totals.avgPrice)]
    ]
    totalsSlide.addTable([
      [
        { text: 'Concepto', options: { fill: { color: BRAND.forest }, color: BRAND.white, bold: true, fontSize: 13 } },
        { text: 'Valor', options: { fill: { color: BRAND.forest }, color: BRAND.white, bold: true, fontSize: 13 } }
      ],
      ...rows.map((row, rowIndex) => row.map((cell) => ({
        text: cell,
        options: {
          fill: { color: rowIndex % 2 ? BRAND.zebra : BRAND.white },
          color: BRAND.ink,
          fontSize: 13
        }
      })))
    ], {
      x: 0.62,
      y: 2.1,
      w: 12.1,
      colW: [7.4, 4.7],
      border: { type: 'solid', color: BRAND.line, pt: 1 },
      align: 'left',
      valign: 'middle',
      margin: 8,
      autoPage: false
    })
    addFooter(totalsSlide, expanded.length + 3)
  }

  // Cierre con próximos pasos
  const closing = pptx.addSlide()
  closing.background = { color: BRAND.deep }
  closing.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: BRAND.lime }, line: { color: BRAND.lime }
  })
  closing.addText('CIERRE EJECUTIVO', {
    x: 0.9, y: 0.9, w: 8, h: 0.4,
    fontSize: 12, bold: true, color: BRAND.lime, charSpacing: 3
  })
  closing.addText('Próximos Pasos', {
    x: 0.9, y: 1.45, w: 9, h: 1.1,
    fontSize: 40, bold: true, color: BRAND.white
  })
  const lastBullets = expanded[expanded.length - 1]?.bullets || []
  const productNames = (Array.isArray(catalog) ? catalog : [])
    .map((item) => String(item?.name ?? item?.nombre ?? item?.producto ?? '').trim().toLowerCase())
    .filter((name) => name.length > 3)
  const mentionsProduct = (text) => productNames.some((name) => String(text).toLowerCase().includes(name))
  const actionable = (text) => /^(?:revisar|validar|definir|publicar|medir|asignar|confirmar|actualizar|programar|capacitar|documentar|priorizar|implementar|monitorear|preparar|comunicar|analizar|ajustar)\b/i.test(text)
  const genericSteps = [
    'Validar precios y existencias con el equipo comercial antes de publicar.',
    'Definir responsables y una fecha concreta para la siguiente revisión.',
    'Medir consultas, conversiones y rotación en 30 días y ajustar el plan.',
    'Documentar los cambios y dejar registrada la decisión final.'
  ]
  const usefulSteps = lastBullets
    .map(stripOrder)
    .filter((text) => text && actionable(text) && !mentionsProduct(text))
  // En catálogos siempre se usan pasos genéricos: ningún nombre de producto
  // se repite en el cierre.
  const nextSteps = (catalogMode ? genericSteps : (usefulSteps.length >= 3 ? usefulSteps : genericSteps)).slice(0, 4)
  nextSteps.forEach((step, index) => {
    const y = 3.0 + index * 1.0
    closing.addShape(pptx.ShapeType.ellipse, {
      x: 0.95, y: y + 0.06, w: 0.42, h: 0.42,
      fill: { color: BRAND.lime }, line: { color: BRAND.lime }
    })
    closing.addText(String(index + 1), {
      x: 0.95, y: y + 0.06, w: 0.42, h: 0.42,
      fontSize: 12, bold: true, color: BRAND.deep, align: 'center', valign: 'middle'
    })
    closing.addText(step, {
      x: 1.6, y, w: 10.8, h: 0.6,
      fontSize: 15, color: 'DCEBE1', valign: 'middle', fit: 'shrink'
    })
  })
  closing.addText('VITIMAX · Centro de Gestión', {
    x: 0.9, y: 6.85, w: 6, h: 0.3,
    fontSize: 10, color: '8AA79A'
  })

  return pptxWriteBlob(pptx)
}

// ---------------------------------------------------------------------------
// 6a. Página de catálogo moderna (Tailwind + Lucide + WhatsApp)
// ---------------------------------------------------------------------------

const CATALOG_PAGE_HINT = /cat[aá]logo|tienda|productos|inventario|lista de precios|precios|vivero|agr[ií]cola|fertiliz|bioestimul/i
const TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'
const LUCIDE_CDN = '<script src="https://unpkg.com/lucide@latest"></script>'

function currentYear() {
  return new Date().getFullYear() || 2026
}

// Marcas típicas de una respuesta cortada por límite de tokens o de un
// placeholder que la IA dejó en el HTML. Nunca deben llegar al usuario.
function isTruncatedHtml(html = '') {
  const source = String(html || '')
  if (!source.trim()) return false
  if (/<!--\s*(?:agregar|añadir|continuar|resto|m[aá]s productos|todo|pendiente|aqu[ií]|m[aá]s secciones)/i.test(source)) return true
  if (/(?:agregar|añadir|incluir)\s+m[aá]s\s+(?:productos|contenido|secciones|tarjetas|items)/i.test(source)) return true
  if (/\[(?:continuar[aá]|truncad[oa]|omitid[oa]|resto del c[oó]digo|corte)\]/i.test(source)) return true
  if (/(?:c[oó]digo|contenido|html)\s+(?:truncad[oa]|omitid[oa]|cortad[oa])/i.test(source)) return true
  if (/<html[\s>]/i.test(source) && !/<\/html>/i.test(source)) return true
  if (/<body[\s>]/i.test(source) && !/<\/body>/i.test(source)) return true
  return false
}

function catalogCoverage(html = '', catalog = []) {
  const names = (Array.isArray(catalog) ? catalog : [])
    .map((product) => String(product?.name || '').trim().toLowerCase())
    .filter(Boolean)
  if (!names.length) return 1
  const haystack = String(html || '').toLowerCase()
  const found = names.filter((name) => haystack.includes(name)).length
  return found / names.length
}

function catalogPageIsComplete(html = '', catalog = []) {
  const source = String(html || '')
  if (!source || isTruncatedHtml(source)) return false
  if (!hasModernCatalogUi(source)) return false
  if (!/(?:wa\.me|whatsapp)/i.test(source)) return false
  if (!/(?:stock|disponible|agotado|sin stock)/i.test(source)) return false
  if (!/(?:data-category|filtrar|categor[ií]a)/i.test(source)) return false
  // Todos los productos reales deben aparecer (mínimo 80% del catálogo).
  return catalogCoverage(source, catalog) >= 0.8
}

export function isCatalogPageRequest({ title = '', subtitle = '', source = '', catalog = [] } = {}) {
  if (!Array.isArray(catalog) || !catalog.length) return false
  const haystack = `${title} ${subtitle} ${String(source || '').slice(0, 600)}`
  return CATALOG_PAGE_HINT.test(haystack)
}

// Detecta si el HTML que devolvió la IA ya es una interfaz moderna
// (Tailwind + grilla/búsqueda) y no una tabla antigua.
export function hasModernCatalogUi(html = '') {
  const source = String(html || '')
  if (!source) return false
  const tailwind = /cdn\.tailwindcss\.com|tailwindcss/i.test(source)
  const search = /type=["']search["']|<input[^>]+(?:buscar|busca|search|filtro)/i.test(source)
  const grid = /grid-template-columns|class=["'][^"']*\bgrid\b/i.test(source)
  return Boolean(grid && (tailwind || search))
}

export function ensureModernAssets(html = '') {
  let page = ensureCharset(ensureViewport(String(html || '')))
  // Una sola carga limpia de Tailwind: se eliminan duplicados y se deja la
  // primera aparición, reemplazada por la URL oficial.
  let tailwindSeen = false
  page = page.replace(/<script\b[^>]*\bsrc=["'][^"']*cdn\.tailwindcss\.com[^"']*["'][^>]*>\s*<\/script>\s*/gi, () => {
    if (tailwindSeen) return ''
    tailwindSeen = true
    return `${TAILWIND_CDN}\n`
  })
  if (!tailwindSeen) {
    page = /<\/head>/i.test(page) ? page.replace(/<\/head>/i, `${TAILWIND_CDN}\n</head>`) : `${TAILWIND_CDN}\n${page}`
  }
  if (!/lucide/i.test(page)) {
    page = /<\/head>/i.test(page) ? page.replace(/<\/head>/i, `${LUCIDE_CDN}\n</head>`) : `${LUCIDE_CDN}\n${page}`
  }
  if (/lucide/i.test(page) && !/createIcons\s*\(/.test(page)) {
    const init = '<script>window.addEventListener("DOMContentLoaded",function(){if(window.lucide)window.lucide.createIcons();});</script>'
    page = /<\/body>/i.test(page) ? page.replace(/<\/body>/i, `${init}\n</body>`) : `${page}\n${init}`
  }
  // Pie de página con el año actual (2026), incluso si la IA lo omitió.
  page = page.replace(/<footer\b[\s\S]*?<\/footer>/i, (footer) => {
    if (/©\s*\d{4}/.test(footer)) return footer.replace(/©\s*\d{4}/g, `© ${currentYear()}`)
    if (/\b(?:19|20)\d{2}\b/.test(footer)) return footer.replace(/\b(?:19|20)\d{2}\b/g, String(currentYear()))
    return footer.replace(/<\/footer>/i, `<p>© ${currentYear()} VITIMAX · Centro de Gestión</p></footer>`)
  })
  return page
}

function whatsappUrl(phone, message) {
  const digits = String(phone || '').replace(/\D/g, '')
  const text = encodeURIComponent(message)
  return digits ? `https://wa.me/${digits}?text=${text}` : `https://wa.me/?text=${text}`
}

function prepareCatalogProducts(catalog = [], phone = '') {
  const totalsLabel = /^(?:gran\s+)?(?:total(?:es)?|subtotal(?:es)?|suma|resumen|valor(?:izado)?)(?:\s+(?:general|valorado|inventario|productos|stock|clp|ejecutivo|de|del|la|el))*\s*$/i
  return (Array.isArray(catalog) ? catalog : []).filter((item) => {
    if (!item) return false
    const name = String(item.name ?? item.nombre ?? item.producto ?? '').trim()
    return !name || !totalsLabel.test(name)
  }).map((item) => {
    const name = sanitizeText(item.name) || 'Producto VITIMAX'
    const category = sanitizeText(item.category) || 'General'
    const format = sanitizeText(item.format)
    const description = sanitizeText(item.description)
    const usage = sanitizeText(item.usage)
    const price = parseNumberLoose(item.price)
    const previousPrice = parseNumberLoose(item.previousPrice)
    const stockPending = item.stock === '' || item.stock === null || item.stock === undefined
    const stock = stockPending ? null : Math.max(0, Math.round(parseNumberLoose(item.stock)))
    const image = String(item.image || '').trim()
    return {
      name,
      category,
      format,
      description,
      usage,
      price,
      priceLabel: formatClp(price),
      previousPrice: previousPrice > price ? previousPrice : 0,
      previousPriceLabel: previousPrice > price ? formatClp(previousPrice) : '',
      stock,
      stockPending,
      image: /^(https?:|data:image\/)/i.test(image) ? image : '',
      whatsapp: whatsappUrl(phone, `Hola VITIMAX, quiero cotizar ${name}${format ? ` (${format})` : ''}. ¿Está disponible?`)
    }
  })
}

const catalogPageScript = (payload) => String.raw`
const PRODUCTS = ${payload};
const state = { q: '', category: 'Todos' };
const grid = document.getElementById('catalog-grid');
const empty = document.getElementById('empty-state');
const countEl = document.getElementById('result-count');
const filters = document.getElementById('category-filters');
const searchInput = document.getElementById('catalog-search');

const esc = (value) => String(value == null ? '' : value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const categories = ['Todos'].concat(Array.from(new Set(PRODUCTS.map((p) => p.category).filter(Boolean))));

function stockBadge(product) {
  if (product.stockPending) return '<span class="inline-flex items-center gap-1.5 rounded-full bg-slate-100 px-2.5 py-1 text-[11px] font-bold text-slate-600"><span class="h-1.5 w-1.5 rounded-full bg-slate-400"></span>Stock pendiente</span>';
  if (!product.stock) return '<span class="inline-flex items-center gap-1.5 rounded-full bg-red-50 px-2.5 py-1 text-[11px] font-bold text-red-700"><span class="h-1.5 w-1.5 rounded-full bg-red-500"></span>Sin stock</span>';
  if (product.stock <= 5) return '<span class="inline-flex items-center gap-1.5 rounded-full bg-amber-50 px-2.5 py-1 text-[11px] font-bold text-amber-800"><span class="h-1.5 w-1.5 rounded-full bg-amber-500"></span>Últimas ' + product.stock + ' u.</span>';
  return '<span class="inline-flex items-center gap-1.5 rounded-full bg-emerald-50 px-2.5 py-1 text-[11px] font-bold text-emerald-800"><span class="h-1.5 w-1.5 rounded-full bg-emerald-500"></span>Stock: ' + product.stock + ' u.</span>';
}

function card(product) {
  const media = product.image
    ? '<img src="' + esc(product.image) + '" alt="' + esc(product.name) + '" class="h-44 w-full object-cover" loading="lazy">'
    : '<div class="flex h-44 w-full items-center justify-center bg-gradient-to-br from-[#1E4D2B] via-[#2F8058] to-[#12321F]"><i data-lucide="sprout" class="h-14 w-14 text-[#D8F55F]"></i></div>';
  const previous = product.previousPriceLabel ? '<span class="text-sm text-slate-400 line-through">' + esc(product.previousPriceLabel) + '</span>' : '';
  const dose = product.usage ? '<div class="mt-3 flex items-start gap-2 rounded-xl bg-[#F4F7F4] px-3 py-2 text-xs leading-relaxed text-[#5F7969]"><i data-lucide="droplets" class="mt-0.5 h-4 w-4 shrink-0 text-[#2F8058]"></i><span><span class="font-semibold text-[#1E4D2B]">Dosis: </span>' + esc(product.usage) + '</span></div>' : '';
  const description = product.description ? '<p class="mt-2 line-clamp-3 text-sm leading-relaxed text-slate-600">' + esc(product.description) + '</p>' : '';
  const format = product.format ? '<span class="mt-1 inline-flex items-center gap-1 text-xs font-medium text-slate-500"><i data-lucide="package" class="h-3.5 w-3.5"></i>' + esc(product.format) + '</span>' : '';
  return '<article class="group flex flex-col overflow-hidden rounded-2xl border border-[#DDE7DA] bg-white shadow-sm transition duration-300 hover:-translate-y-1 hover:shadow-xl">' +
    media +
    '<div class="flex flex-1 flex-col p-5">' +
      '<div class="flex items-center justify-between gap-2">' +
        '<span class="inline-flex items-center rounded-full bg-[#EAF3D0] px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide text-[#1E4D2B]">' + esc(product.category) + '</span>' +
        stockBadge(product) +
      '</div>' +
      '<h3 class="mt-3 text-lg font-bold leading-snug text-[#12321F]">' + esc(product.name) + '</h3>' +
      format +
      description + dose +
      '<div class="mt-auto pt-5">' +
        '<div class="flex items-end justify-between gap-3">' +
          '<div><p class="text-[11px] font-semibold uppercase tracking-wide text-slate-400">Precio</p><p class="text-2xl font-extrabold text-[#1E4D2B]">' + esc(product.priceLabel) + '</p>' + previous + '</div>' +
        '</div>' +
        '<a href="' + esc(product.whatsapp) + '" target="_blank" rel="noopener" class="mt-4 flex w-full items-center justify-center gap-2 rounded-xl bg-[#25D366] px-4 py-3 text-sm font-bold text-white shadow-sm transition hover:bg-[#1EBE5D]"><i data-lucide="message-circle" class="h-4 w-4"></i>Cotizar por WhatsApp</a>' +
      '</div>' +
    '</div>' +
  '</article>';
}

function matches(product) {
  if (state.category !== 'Todos' && product.category !== state.category) return false;
  if (!state.q) return true;
  const hay = (product.name + ' ' + product.category + ' ' + product.format + ' ' + product.description + ' ' + product.usage).toLowerCase();
  return hay.indexOf(state.q) > -1;
}

function renderFilters() {
  filters.innerHTML = categories.map(function (category) {
    const active = category === state.category;
    return '<button type="button" data-category="' + esc(category) + '" class="rounded-full border px-4 py-2 text-sm font-semibold transition ' + (active ? 'border-[#1E4D2B] bg-[#1E4D2B] text-white shadow' : 'border-[#DDE7DA] bg-white text-[#24382B] hover:border-[#2F8058] hover:text-[#1E4D2B]') + '">' + esc(category) + '</button>';
  }).join('');
}

function render() {
  const list = PRODUCTS.filter(matches);
  grid.innerHTML = list.map(card).join('');
  empty.classList.toggle('hidden', list.length > 0);
  countEl.textContent = list.length + (list.length === 1 ? ' producto' : ' productos');
  if (window.lucide) window.lucide.createIcons();
}

filters.addEventListener('click', function (event) {
  const button = event.target.closest('button[data-category]');
  if (!button) return;
  state.category = button.getAttribute('data-category');
  renderFilters();
  render();
});

searchInput.addEventListener('input', function (event) {
  state.q = String(event.target.value || '').trim().toLowerCase();
  render();
});

renderFilters();
render();
`

export function buildModernCatalogPage({ title = 'Catálogo VITIMAX', subtitle = '', catalog = [], phone = '', isDemo = false, dateLabel = '' } = {}) {
  const cleanTitle = sanitizeText(title) || 'Catálogo VITIMAX'
  const cleanSubtitle = sanitizeText(subtitle) || 'Productos, precios y stock actualizados de VITIMAX.'
  const products = prepareCatalogProducts(catalog, phone)
  const totals = computeCatalogTotals(catalog)
  const categories = new Set(products.map((product) => product.category).filter(Boolean))
  const data = JSON.stringify(products)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029')
  const stats = [
    { label: 'Productos', value: String(totals.count) },
    { label: 'Categorías', value: String(categories.size) },
    { label: 'Unidades en stock', value: String(totals.units) },
    { label: 'Valor inventario', value: formatClp(totals.inventoryValue) }
  ]

  return `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="${escapeHtml(cleanSubtitle)}">
<title>${escapeHtml(cleanTitle)} · VITIMAX</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
${TAILWIND_CDN}
<script>
tailwind.config = {
  theme: {
    extend: {
      colors: {
        forest: '#1E4D2B', deep: '#12321F', mid: '#2F8058', lime: '#D8F55F',
        gold: '#E9C46A', paper: '#F4F7F4', ink: '#24382B'
      },
      fontFamily: { sans: ['Inter', 'system-ui', 'sans-serif'] }
    }
  }
}
</script>
${LUCIDE_CDN}
<style>
  html { scroll-behavior: smooth }
  body { font-family: Inter, system-ui, -apple-system, 'Segoe UI', sans-serif }
  .hero-glow { background: radial-gradient(60% 80% at 85% 0%, rgba(216,245,95,.28), transparent 60%), radial-gradient(50% 70% at 10% 100%, rgba(233,196,106,.18), transparent 65%) }
  ::selection { background: #D8F55F; color: #12321F }
</style>
</head>
<body class="bg-[#F4F7F4] text-[#24382B] antialiased">
<header class="relative overflow-hidden bg-[#1E4D2B] text-white">
  <div class="absolute inset-0 hero-glow"></div>
  <div class="relative mx-auto max-w-7xl px-4 py-14 sm:px-6 sm:py-20 lg:px-8">
    <div class="flex flex-wrap items-center gap-3">
      <span class="inline-flex items-center gap-2 rounded-full border border-white/20 bg-white/10 px-3 py-1 text-[11px] font-bold uppercase tracking-[0.2em] text-[#D8F55F]"><i data-lucide="leaf" class="h-3.5 w-3.5"></i>VITIMAX · Centro de Gestión</span>
      ${isDemo ? '<span class="rounded-full bg-[#E9C46A] px-3 py-1 text-[11px] font-bold uppercase tracking-wide text-[#12321F]">Datos de ejemplo</span>' : ''}
    </div>
    <h1 class="mt-5 max-w-3xl text-4xl font-extrabold leading-tight tracking-tight sm:text-5xl">${escapeHtml(cleanTitle)}</h1>
    <p class="mt-4 max-w-2xl text-base leading-relaxed text-[#D6E8DC] sm:text-lg">${escapeHtml(cleanSubtitle)}</p>
    <div class="mt-8 max-w-xl">
      <label for="catalog-search" class="sr-only">Buscar producto</label>
      <div class="relative">
        <i data-lucide="search" class="pointer-events-none absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#5F7969]"></i>
        <input id="catalog-search" type="search" autocomplete="off" placeholder="Buscar por producto, categoría o dosis…" class="w-full rounded-2xl border-0 bg-white py-4 pl-12 pr-4 text-base text-[#12321F] shadow-2xl shadow-black/20 outline-none ring-2 ring-transparent transition placeholder:text-slate-400 focus:ring-[#D8F55F]">
      </div>
    </div>
    <dl class="mt-10 grid grid-cols-2 gap-3 sm:grid-cols-4 sm:gap-4">
      ${stats.map((item) => `<div class="rounded-2xl border border-white/15 bg-white/10 p-4 backdrop-blur"><dt class="text-[11px] font-bold uppercase tracking-wide text-[#D8F55F]">${escapeHtml(item.label)}</dt><dd class="mt-1 text-xl font-extrabold sm:text-2xl">${escapeHtml(item.value)}</dd></div>`).join('')}
    </dl>
  </div>
</header>

${isDemo ? '<div class="border-b border-amber-200 bg-amber-50 px-4 py-3 text-center text-sm font-medium text-amber-900">Datos de ejemplo generados por IA. No son información oficial de VITIMAX.</div>' : ''}

<div class="sticky top-0 z-30 border-b border-[#DDE7DA] bg-white/95 backdrop-blur">
  <div class="mx-auto flex max-w-7xl flex-col gap-3 px-4 py-4 sm:px-6 lg:flex-row lg:items-center lg:justify-between lg:px-8">
    <div id="category-filters" class="flex flex-wrap gap-2"></div>
    <p id="result-count" class="text-sm font-semibold text-[#5F7969]"></p>
  </div>
</div>

<main class="mx-auto max-w-7xl px-4 py-10 sm:px-6 lg:px-8">
  <div id="catalog-grid" class="grid gap-6 sm:grid-cols-2 lg:grid-cols-3"></div>
  <div id="empty-state" class="hidden rounded-3xl border border-dashed border-[#DDE7DA] bg-white px-6 py-16 text-center">
    <i data-lucide="search-x" class="mx-auto h-12 w-12 text-[#2F8058]"></i>
    <h2 class="mt-4 text-xl font-bold text-[#12321F]">Sin resultados</h2>
    <p class="mt-2 text-sm text-[#5F7969]">Prueba con otra palabra o cambia la categoría.</p>
  </div>
</main>

<footer class="border-t border-[#DDE7DA] bg-white">
  <div class="mx-auto flex max-w-7xl flex-col items-center justify-between gap-3 px-4 py-8 text-sm text-[#5F7969] sm:flex-row sm:px-6 lg:px-8">
    <p class="font-semibold text-[#1E4D2B]">VITIMAX · Centro de Gestión</p>
    <p>© ${currentYear()} VITIMAX · Catálogo generado el ${escapeHtml(dateLabel || todayLabel())}</p>
  </div>
</footer>

<script>${catalogPageScript(data)}</script>
</body>
</html>`
}

function wrapModernFragment(body, title, subtitle, dateLabel = '') {
  return ensureModernAssets(`<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="${escapeHtml(subtitle || title)}">
${dateLabel ? `<meta name="generated-at" content="${escapeHtml(dateLabel)}">` : ''}
<title>${escapeHtml(title)} · VITIMAX</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
${TAILWIND_CDN}
</head>
<body class="bg-[#F4F7F4] text-[#24382B] antialiased">
${body}
</body>
</html>`)
}

// ---------------------------------------------------------------------------
// 6. Página web autocontenida
// ---------------------------------------------------------------------------

const BASE_PAGE_CSS = `
  :root{--forest:#1E4D2B;--deep:#12321F;--mid:#2F8058;--lime:#D8F55F;--paper:#F4F7F4;--ink:#24382B;--muted:#5F7969;--line:#DDE7DA}
  *{box-sizing:border-box}
  body{margin:0;font-family:"Segoe UI",system-ui,-apple-system,sans-serif;color:var(--ink);background:var(--paper);line-height:1.65;-webkit-font-smoothing:antialiased}
  .vitimax-shell{max-width:1080px;margin:0 auto;padding:clamp(20px,5vw,56px)}
  header.vitimax-hero{background:linear-gradient(135deg,var(--deep),var(--forest));color:#fff;border-radius:24px;padding:clamp(28px,6vw,64px);box-shadow:0 24px 60px #12321f26}
  header.vitimax-hero small{display:inline-block;letter-spacing:.22em;font-weight:700;font-size:.72rem;color:var(--lime);text-transform:uppercase}
  header.vitimax-hero h1{margin:.5rem 0 1rem;font-size:clamp(1.8rem,5vw,3.2rem);line-height:1.1}
  header.vitimax-hero p{margin:0;color:#d6e8dc;max-width:62ch}
  main.vitimax-content{background:#fff;border:1px solid var(--line);border-radius:24px;padding:clamp(22px,5vw,48px);margin-top:22px;box-shadow:0 18px 44px #12321f12}
  h2,h3{color:var(--forest);line-height:1.2}
  a{color:var(--mid)}
  img{max-width:100%;height:auto;border-radius:16px}
  table{width:100%;border-collapse:collapse;margin:18px 0}
  th{background:var(--forest);color:#fff;text-align:left;padding:11px 12px}
  td{padding:10px 12px;border-bottom:1px solid var(--line)}
  tr:nth-child(even) td{background:var(--paper)}
  button,.btn{background:var(--forest);color:#fff;border:0;border-radius:12px;padding:12px 22px;font-weight:700;cursor:pointer;transition:transform .2s ease,box-shadow .2s ease}
  button:hover,.btn:hover{transform:translateY(-2px);box-shadow:0 10px 24px #12321f33}
  footer.vitimax-foot{text-align:center;color:var(--muted);font-size:.82rem;padding:26px 0 8px}
  @media(max-width:640px){.vitimax-shell{padding:16px}header.vitimax-hero{border-radius:18px}main.vitimax-content{border-radius:18px}}
`

function ensureViewport(html) {
  if (/<meta[^>]+name=["']viewport["']/i.test(html)) return html
  return html.replace(/<head([^>]*)>/i, `<head$1><meta name="viewport" content="width=device-width, initial-scale=1">`)
}

function ensureCharset(html) {
  if (/<meta[^>]+charset=/i.test(html)) return html
  return html.replace(/<head([^>]*)>/i, `<head$1><meta charset="utf-8">`)
}

function inlineHtml(value) {
  return inlineSegments(value)
    .map((segment) => (segment.bold ? `<strong>${escapeHtml(segment.text)}</strong>` : escapeHtml(segment.text)))
    .join('')
}

// Si el modelo devuelve Markdown en lugar de HTML, se convierte a HTML real
// para que la página nunca muestre "#", "**", "|" ni viñetas sin formato.
function blocksToHtml(markdown) {
  return parseDocument(markdown).map((block) => {
    if (block.type === 'heading') {
      const level = Math.min(Math.max(block.level + 1, 2), 6)
      return `<h${level}>${inlineHtml(block.text)}</h${level}>`
    }
    if (block.type === 'paragraph') return `<p>${inlineHtml(block.text)}</p>`
    if (block.type === 'bullets') return `<ul>${block.items.map((item) => `<li>${inlineHtml(item)}</li>`).join('')}</ul>`
    if (block.type === 'numbers') return `<ol>${block.items.map((item) => `<li>${inlineHtml(item)}</li>`).join('')}</ol>`
    if (block.type === 'table') {
      const head = block.header.map((cell) => `<th>${inlineHtml(cell)}</th>`).join('')
      const body = block.rows.map((row) => `<tr>${block.header.map((_, index) => `<td>${inlineHtml(row[index] || '')}</td>`).join('')}</tr>`).join('')
      return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`
    }
    if (block.type === 'quote') return `<blockquote>${inlineHtml(block.text)}</blockquote>`
    if (block.type === 'code') return `<pre><code>${escapeHtml(block.code)}</code></pre>`
    return ''
  }).filter(Boolean).join('\n')
}

const HTML_TAG = /<(?:!doctype|html|head|body|div|section|article|h[1-6]|p|ul|ol|li|table|thead|tbody|tr|td|th|header|main|footer|style|script|span|a|img|br|strong|em|button|form|input|label|nav)\b/i

export function buildHtmlPage({ title = 'Página VITIMAX', subtitle = '', html = '', catalog = [], phone = '', isDemo = false, dateLabel = '' } = {}) {
  const cleanTitle = plainText(title) || 'Página VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const source = stripSandboxLinks(String(html || '')).trim()
  const catalogRequest = isCatalogPageRequest({ title: cleanTitle, subtitle: cleanSubtitle, source, catalog })

  // Catálogo real: se construye una página moderna con búsqueda, filtros y
  // tarjetas aunque la IA haya devuelto una tabla antigua, Markdown, un HTML
  // truncado o un catálogo incompleto.
  if (catalogRequest) {
    if (/<html[\s>]/i.test(source) && catalogPageIsComplete(source, catalog)) {
      return ensureModernAssets(source)
    }
    if (source && !/<html[\s>]/i.test(source) && catalogPageIsComplete(source, catalog)) {
      const fragment = /<body[\s>]/i.test(source)
        ? source.replace(/^[\s\S]*?<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '')
        : source
      return wrapModernFragment(fragment, cleanTitle, cleanSubtitle, dateLabel)
    }
    return buildModernCatalogPage({ title: cleanTitle, subtitle: cleanSubtitle, catalog, phone, isDemo, dateLabel })
  }

  if (!source) throw new Error('La página HTML está vacía.')

  if (/<html[\s>]/i.test(source)) {
    let page = ensureCharset(ensureViewport(source))
    if (!/<title[\s>]/i.test(page)) {
      page = page.replace(/<head([^>]*)>/i, `<head$1><title>${escapeHtml(cleanTitle)}</title>`)
    }
    return page
  }

  const body = !HTML_TAG.test(source)
    ? blocksToHtml(source)
    : /<body[\s>]/i.test(source)
      ? source.replace(/^[\s\S]*?<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '')
      : source

  return `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="${escapeHtml(cleanSubtitle || cleanTitle)}">
<title>${escapeHtml(cleanTitle)}</title>
<style>${BASE_PAGE_CSS}</style>
</head>
<body>
<div class="vitimax-shell">
  <header class="vitimax-hero">
    <small>VITIMAX · Centro de Gestión</small>
    <h1>${escapeHtml(cleanTitle)}</h1>
    ${cleanSubtitle ? `<p>${escapeHtml(cleanSubtitle)}</p>` : ''}
  </header>
  <main class="vitimax-content">
${body}
  </main>
  <footer class="vitimax-foot">© ${currentYear()} VITIMAX · Generado el ${escapeHtml(dateLabel || todayLabel())}</footer>
</div>
</body>
</html>`
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[character]))
}

export { escapeHtml }

// ---------------------------------------------------------------------------
// 7. Utilidades de código
// ---------------------------------------------------------------------------

export const CODE_EXTENSIONS = {
  html: 'html', htm: 'html', css: 'css', js: 'js', javascript: 'js', jsx: 'jsx', ts: 'ts', typescript: 'ts', tsx: 'tsx',
  json: 'json', python: 'py', py: 'py', php: 'php', java: 'java', kotlin: 'kt', swift: 'swift', c: 'c', cpp: 'cpp',
  'c++': 'cpp', cs: 'cs', csharp: 'cs', 'c#': 'cs', go: 'go', rust: 'rs', ruby: 'rb', sql: 'sql', bash: 'sh', shell: 'sh', sh: 'sh',
  powershell: 'ps1', ps1: 'ps1', xml: 'xml', yaml: 'yml', yml: 'yml', markdown: 'md', md: 'md', txt: 'txt', vue: 'vue', dart: 'dart'
}

export const CODE_MIME = {
  html: 'text/html;charset=utf-8',
  css: 'text/css;charset=utf-8',
  js: 'text/javascript;charset=utf-8',
  jsx: 'text/javascript;charset=utf-8',
  ts: 'text/typescript;charset=utf-8',
  tsx: 'text/typescript;charset=utf-8',
  json: 'application/json;charset=utf-8',
  py: 'text/x-python;charset=utf-8',
  php: 'application/x-httpd-php;charset=utf-8',
  sql: 'application/sql;charset=utf-8',
  sh: 'application/x-sh;charset=utf-8',
  ps1: 'text/plain;charset=utf-8',
  md: 'text/markdown;charset=utf-8'
}

export function detectLanguage(value) {
  const text = String(value || '').toLowerCase()
  const pairs = [
    ['python', /python|\.py\b|django|flask|pandas/],
    ['javascript', /javascript|\.js\b|node\.?js|react|vue|express/],
    ['typescript', /typescript|\.ts\b/],
    ['sql', /sql|mysql|postgres|sqlite|consulta a la base/],
    ['php', /php|laravel/],
    ['csharp', /c#|csharp|\.net|asp\.net/],
    ['java', /java(?!script)|spring/],
    ['kotlin', /kotlin/],
    ['swift', /swift/],
    ['go', /\bgo\b|golang/],
    ['rust', /rust|cargo/],
    ['html', /html|pagina web|página web|landing|sitio web/],
    ['css', /css|estilos/],
    ['bash', /bash|shell|terminal/],
    ['powershell', /powershell/]
  ]
  for (const [language, pattern] of pairs) if (pattern.test(text)) return language
  return 'txt'
}

export function extractFencedCode(value, fallbackLanguage = 'txt') {
  const source = String(value || '')
  const blocks = []
  const fenced = /```([a-zA-Z0-9+#.-]*)\s*\n?([\s\S]*?)```/g
  let match
  while ((match = fenced.exec(source))) {
    blocks.push({ language: (match[1] || fallbackLanguage).toLowerCase(), code: match[2].trim() })
  }
  // Si el modelo devuelve varios bloques, se conserva el completo más largo:
  // nunca se entrega un fragmento truncado del código.
  if (blocks.length) {
    return blocks.reduce((best, block) => block.code.length > best.code.length ? block : best)
  }
  const open = source.match(/```([a-zA-Z0-9+#.-]*)\s*\n?([\s\S]*)$/)
  if (open) return { language: (open[1] || fallbackLanguage).toLowerCase(), code: open[2].trim() }
  return { language: fallbackLanguage, code: source.trim() }
}
