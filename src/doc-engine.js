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
  try {
    return new Date().toLocaleDateString('es-CL', { day: 'numeric', month: 'long', year: 'numeric' })
  } catch {
    return new Date().toISOString().slice(0, 10)
  }
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

// Limpia marcas Markdown dentro de una línea, conservando el contenido.
function cleanInline(value) {
  return String(value ?? '')
    .replace(/\$\$([^$]*)\$\$/g, '$1')
    .replace(/\$([^$\n]+)\$/g, '$1')
    .replace(/`([^`]*)`/g, '$1')
    .replace(/\*\*/g, '')
    .replace(/__+/g, '')
    .replace(/\*/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]{2,}/g, ' ')
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

export function buildPdfDocument({ title = 'Documento VITIMAX', subtitle = '', content = '', dateLabel = '' } = {}) {
  const doc = new jsPDF({ unit: 'mm', format: 'a4', compress: true })
  const pageW = doc.internal.pageSize.getWidth()
  const pageH = doc.internal.pageSize.getHeight()
  const margin = 18
  const contentW = pageW - margin * 2
  const bottom = pageH - 22
  const blocks = parseDocument(content)
  const cleanTitle = plainText(title) || 'Documento VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const date = dateLabel || todayLabel()
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
    const weights = header.map((label, index) => {
      const longest = rows.reduce((max, row) => Math.max(max, String(row[index] || '').length), 0)
      return Math.min(46, Math.max(10, Math.max(label.length, longest)))
    })
    const total = weights.reduce((sum, weight) => sum + weight, 0) || 1
    const widths = weights.map((weight) => (weight / total) * contentW)
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

  const skipTitle = blocks.length && blocks[0].type === 'heading' && blocks[0].level === 1
    && normalizeTitle(blocks[0].text) === normalizeTitle(cleanTitle)
  let startIndex = skipTitle ? 1 : 0

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

export function buildDocxDocument({ title = 'Documento VITIMAX', subtitle = '', content = '', dateLabel = '' } = {}) {
  const blocks = parseDocument(content)
  const cleanTitle = plainText(title) || 'Documento VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const date = dateLabel || todayLabel()
  const children = []

  // Membrete / portada corporativa
  children.push(new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({ text: 'VITIMAX', bold: true, size: 22, color: BRAND.forest, characterSpacing: 80 }),
      new TextRun({ text: '   CENTRO DE GESTIÓN', size: 16, color: BRAND.muted, characterSpacing: 40 })
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

  const skipTitle = blocks.length && blocks[0].type === 'heading' && blocks[0].level === 1
    && normalizeTitle(blocks[0].text) === normalizeTitle(cleanTitle)
  const startIndex = skipTitle ? 1 : 0

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
      const width = Math.floor(100 / header.length)
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
            children: header.map((cell) => docxTableCell(cell, { header: true, width }))
          }),
          ...rows.map((row, rowIndex) => new TableRow({
            children: header.map((_, cellIndex) => docxTableCell(row[cellIndex] || '', { zebra: rowIndex % 2 === 1, width }))
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

export async function buildPptxDocument({ title = 'Presentación VITIMAX', subtitle = '', content = '', dateLabel = '' } = {}) {
  const deckTitle = plainText(title) || 'Presentación VITIMAX'
  const date = dateLabel || todayLabel()
  const blocks = parseDocument(content)
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

  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'VITIMAX'
  pptx.company = 'VITIMAX'
  pptx.subject = deckTitle
  pptx.title = deckTitle
  pptx.theme = { headFontFace: 'Segoe UI Semibold', bodyFontFace: 'Segoe UI', lang: 'es-CL' }

  const coverSubtitle = plainText(subtitle)
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

  const addFooter = (slide, index) => {
    slide.addText('VITIMAX · Centro de Gestión', {
      x: 0.6, y: 7.12, w: 5, h: 0.25,
      fontSize: 8, color: '8AA79A'
    })
    slide.addText(String(index + 2), {
      x: 12.1, y: 7.12, w: 0.7, h: 0.25,
      fontSize: 8, color: '8AA79A', align: 'right'
    })
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
      const rows = item.table.rows.slice(0, 8)
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
      slide.addTable(tableRows, {
        x: 0.62,
        y: bodyY,
        w: 12.1,
        colW: header.map(() => 12.1 / header.length),
        border: { type: 'solid', color: BRAND.line, pt: 1 },
        align: 'left',
        valign: 'middle',
        margin: 6,
        autoPage: false
      })
      addFooter(slide, index)
      return
    }

    const bullets = item.bullets.slice(0, 5)
    if (!bullets.length) {
      slide.addText(item.paragraphs.map((paragraph) => paragraph.text).join('\n') || 'Contenido de la presentación.', {
        x: 0.7, y: bodyY + 0.2, w: 11.9, h: 4.4,
        fontSize: 18, color: BRAND.ink, fit: 'shrink', valign: 'top'
      })
      addFooter(slide, index)
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
    addFooter(slide, index)
  })

  // Cierre con próximos pasos
  const closing = pptx.addSlide()
  closing.background = { color: BRAND.deep }
  closing.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.18, h: 7.5,
    fill: { color: BRAND.lime }, line: { color: BRAND.lime }
  })
  closing.addText('PRÓXIMOS PASOS', {
    x: 0.9, y: 0.9, w: 8, h: 0.4,
    fontSize: 12, bold: true, color: BRAND.lime, charSpacing: 3
  })
  closing.addText('Gracias', {
    x: 0.9, y: 1.45, w: 9, h: 1.1,
    fontSize: 40, bold: true, color: BRAND.white
  })
  const lastBullets = expanded[expanded.length - 1]?.bullets || []
  const nextSteps = (lastBullets.length ? lastBullets : [
    'Revisar el documento y validar los datos clave.',
    'Definir responsables y fechas para la ejecución.',
    'Medir resultados y ajustar el plan en el próximo periodo.'
  ]).slice(0, 3).map(stripOrder)
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

  return pptx.write({ outputType: 'blob' })
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

export function buildHtmlPage({ title = 'Página VITIMAX', subtitle = '', html = '' } = {}) {
  const cleanTitle = plainText(title) || 'Página VITIMAX'
  const cleanSubtitle = plainText(subtitle)
  const source = stripSandboxLinks(String(html || '')).trim()
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
  <footer class="vitimax-foot">Generado por VITIMAX · ${escapeHtml(todayLabel())}</footer>
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
  const fenced = source.match(/```([a-zA-Z0-9+#.-]*)\s*\n?([\s\S]*?)```/)
  if (fenced) return { language: (fenced[1] || fallbackLanguage).toLowerCase(), code: fenced[2].trim() }
  return { language: fallbackLanguage, code: source.trim() }
}
