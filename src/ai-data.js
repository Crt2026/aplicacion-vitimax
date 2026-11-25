import * as XLSX from 'xlsx'
import { parseNumberLoose, sanitizeText } from './text-utils.js'
import {
  buildTemporalSection,
  buildWelcomeMessage,
  formatElapsedSince,
  formatSpanishDate,
  getTemporalContext,
  isWelcomeMessage,
  lastInteractionAt,
  resolveDisplayName
} from './temporal.js'

// El reloj local se expone desde aquí para que la app, los prompts y los
// documentos compartan exactamente la misma referencia de fecha y hora.
export {
  buildTemporalSection,
  buildWelcomeMessage,
  formatElapsedSince,
  formatSpanishDate,
  getTemporalContext,
  isWelcomeMessage,
  lastInteractionAt,
  resolveDisplayName
}

// Extrae texto de un PDF en el propio dispositivo (web, Android e iPhone).
// Así el BOT puede analizar el contenido aunque el modelo no acepte PDF nativo.
export async function extractPdfText(file, maxPages = 30) {
  const data = new Uint8Array(await file.arrayBuffer())
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  try {
    if (pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
      const worker = await import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url')
      pdfjs.GlobalWorkerOptions.workerSrc = worker.default
    }
  } catch {
    // Sin worker: pdf.js funciona en el hilo principal igualmente.
  }
  const loading = pdfjs.getDocument({
    data,
    isEvalSupported: false,
    useWorkerFetch: false,
    disableFontFace: true
  })
  const doc = await loading.promise
  const totalPages = doc.numPages
  const chunks = []
  const limit = Math.min(totalPages, maxPages)
  for (let index = 1; index <= limit; index += 1) {
    const page = await doc.getPage(index)
    const content = await page.getTextContent()
    const text = content.items
      .map((item) => item.str || '')
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim()
    if (text) chunks.push(`PÁGINA ${index}: ${text}`)
    page.cleanup?.()
  }
  try { await doc.destroy() } catch { /* nada */ }
  return {
    text: chunks.join('\n\n'),
    pages: totalPages,
    readPages: limit
  }
}

// ---------------------------------------------------------------------------
// Contrato de datos del BOT IA VITIMAX
// El BOT puede generar hojas, columnas y filas completas y la app las convierte
// en archivos reales (Excel, PDF, Word, PowerPoint) sin depender del catálogo.
// ---------------------------------------------------------------------------

export const SHEET_NAMES = ['Productos', 'Clientes', 'Plantas', 'Inventario', 'Ventas', 'Asesorias', 'Proveedores']

export const PRODUCT_FIELDS = [
  'id', 'name', 'category', 'price', 'previousPrice', 'stock', 'format',
  'description', 'usage', 'recommendations', 'features', 'image', 'source'
]

export const PRODUCT_COLUMNS = [
  'ID', 'Nombre', 'Categoría', 'Precio CLP', 'Precio anterior', 'Stock',
  'Formato', 'Descripción', 'Modo de uso / dosis', 'Recomendaciones', 'Características', 'Origen'
]

const SHEET_COLUMNS = {
  Productos: PRODUCT_COLUMNS,
  Clientes: ['ID', 'Nombre', 'Teléfono', 'Correo', 'Localidad', 'Notas'],
  Plantas: ['ID', 'Planta / Cultivo', 'Nombre científico', 'Tipo', 'Clima', 'Suelo recomendado', 'Riego', 'Temporada', 'Notas'],
  Inventario: ['ID', 'Producto', 'Categoría', 'Stock actual', 'Stock mínimo', 'Estado', 'Precio CLP', 'Valor inventario CLP'],
  Ventas: ['ID', 'Fecha', 'Cliente', 'Producto', 'Cantidad', 'Precio unitario CLP', 'Total CLP', 'Estado', 'Canal', 'ID Cliente', 'ID Producto'],
  Asesorias: ['ID', 'Cliente', 'Cultivo', 'Consulta', 'Estado', 'Prioridad', 'Recomendación', 'ID Cliente'],
  Proveedores: ['ID', 'Proveedor', 'Contacto', 'Teléfono', 'Correo', 'Ciudad', 'Productos', 'Condiciones']
}

export function makeId(prefix = 'vitimax') {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID()
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
}

export function cleanCell(value) {
  if (value === null || value === undefined) return ''
  if (typeof value === 'number' || typeof value === 'boolean') return value
  const raw = String(value).replace(/\u00a0/g, ' ').trim()
  // El usuario pidió quitar los "-" que aparecían como dato vacío.
  if (/^[-–—]+$/.test(raw)) return ''
  if (/^(n\/?a|sin dato|no aplica)$/i.test(raw)) return ''
  // Las fórmulas se conservan tal cual: el escritor avanzado las convierte en
  // fórmulas reales de Excel. Si se escaparan, el archivo perdería el cálculo.
  if (raw.startsWith('=')) return raw
  // Limpieza compartida: quita Markdown, LaTeX, "~", "**", "#" y escapes.
  const text = sanitizeText(
    raw
      .replace(/\[([^\]]*)\]\(\s*(?:sandbox|file):[^)]*\)/gi, '$1')
      .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
  )
  return text.replace(/\s+/g, ' ').trim()
}

export function toNumber(value) {
  return parseNumberLoose(value)
}

function pick(raw, keys) {
  const entries = Object.entries(raw).map(([key, value]) => [key.toLowerCase().trim(), value])
  for (const key of keys) {
    const target = String(key).toLowerCase().trim()
    const found = entries.find(([entryKey, value]) => entryKey === target && value !== undefined && value !== null && value !== '')
    if (found) return found[1]
  }
  return ''
}

export function normalizeProduct(raw = {}, options = {}) {
  const product = {
    id: String(pick(raw, ['id', 'ID', 'codigo', 'código']) || makeId('producto')),
    name: cleanCell(pick(raw, ['name', 'nombre', 'producto', 'Name'])),
    category: cleanCell(pick(raw, ['category', 'categoria', 'categoría', 'rubro', 'tipo'])),
    price: toNumber(pick(raw, ['price', 'precio', 'precio_clp', 'precio CLP', 'valor'])),
    previousPrice: toNumber(pick(raw, ['previousPrice', 'precio_anterior', 'precio anterior', 'precioAntiguo'])),
    stock: toNumber(pick(raw, ['stock', 'existencias', 'cantidad', 'unidades', 'inventario'])),
    format: cleanCell(pick(raw, ['format', 'formato', 'presentacion', 'presentación', 'tamaño', 'tamano'])),
    description: cleanCell(pick(raw, ['description', 'descripcion', 'descripción', 'detalle'])),
    usage: cleanCell(pick(raw, ['usage', 'uso', 'modo_de_uso', 'modo de uso', 'modo de uso / dosis', 'modo de uso/dosis', 'dosis y modo de uso', 'dosis', 'aplicacion', 'aplicación'])),
    recommendations: cleanCell(pick(raw, ['recommendations', 'recomendaciones', 'consejos'])),
    features: cleanCell(pick(raw, ['features', 'caracteristicas', 'características', 'beneficios'])),
    image: cleanCell(pick(raw, ['image', 'imagen', 'foto', 'fotografia', 'fotografía', 'url imagen'])),
    source: cleanCell(pick(raw, ['source', 'origen', 'generadoPorIA', 'generado_por_ia'])) || (options.generated ? 'Generado por IA (ejemplo)' : '')
  }
  return product
}

export function normalizeProducts(list, options = {}) {
  if (!Array.isArray(list)) return []
  return list
    .map((item) => normalizeProduct(item, options))
    .filter((item) => (item.name || item.description || item.price) && !isJunkProductName(item.name))
}

// ---------------------------------------------------------------------------
// Importador universal de Excel/CSV
// Detecta encabezados aunque no estén en la fila 1, mapea columnas por
// sinónimos, recorre todas las hojas y cruza Inventario con Productos.
// ---------------------------------------------------------------------------

export function normalizeHeader(value) {
  return String(value ?? '')
    .replace(/\u00a0/g, ' ')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
}

const PRODUCT_ALIASES = {
  id: ['id', 'id producto', 'producto id', 'codigo', 'codigo producto', 'cod producto', 'cod', 'sku', 'referencia', 'ref'],
  name: ['nombre', 'nombre producto', 'nombre del producto', 'producto', 'articulo', 'item', 'insumo', 'titulo', 'denominacion', 'descripcion corta'],
  category: ['categoria', 'categoria producto', 'tipo', 'familia', 'rubro', 'linea', 'linea producto', 'grupo'],
  previousPrice: ['precio anterior', 'precio antes', 'precio normal', 'precio regular', 'precio lista', 'precio de lista', 'precio antiguo', 'precioantiguo', 'precioanterior'],
  price: ['precio', 'precio venta', 'precio de venta', 'precio unitario', 'valor', 'valor unitario', 'costo', 'pvp', 'monto', 'precio clp', 'precio lista', 'precio publico'],
  stock: ['stock', 'stock actual', 'stock disponible', 'cantidad', 'cantidad disponible', 'existencias', 'unidades', 'inventario', 'saldo', 'disponible', 'stock fisico'],
  format: ['formato', 'presentacion', 'tamano', 'tamano envase', 'medida', 'contenido', 'envase', 'volumen', 'peso', 'capacidad'],
  description: ['descripcion', 'descripcion producto', 'detalle', 'caracteristicas', 'dosis', 'modo de uso', 'modo uso', 'modouso', 'uso', 'informacion', 'observaciones', 'composicion', 'beneficios'],
  usage: ['modo de uso', 'modo uso', 'modouso', 'dosis', 'dosis recomendada', 'aplicacion', 'instrucciones', 'uso recomendado', 'forma de uso'],
  recommendations: ['recomendaciones', 'recomendacion', 'consejos', 'observacion', 'notas'],
  features: ['caracteristicas', 'beneficios', 'propiedades', 'atributos', 'composicion'],
  image: ['imagen', 'foto', 'fotografia', 'url imagen', 'imagen url', 'link imagen', 'foto url', 'url foto'],
  source: ['origen', 'fuente', 'source', 'generado por ia', 'generadopor ia', 'generadoPorIA']
}

const RECORD_ALIASES = {
  clientes: {
    id: ['id', 'id cliente', 'codigo', 'codigo cliente', 'rut', 'documento'],
    nombre: ['nombre', 'nombre cliente', 'cliente', 'razon social', 'nombre completo', 'contacto'],
    telefono: ['telefono', 'fono', 'celular', 'movil', 'whatsapp', 'tel', 'numero'],
    email: ['email', 'correo', 'correo electronico', 'e mail', 'mail'],
    localidad: ['localidad', 'ciudad', 'comuna', 'region', 'direccion', 'zona', 'ubicacion'],
    notas: ['notas', 'observaciones', 'comentarios', 'detalle']
  },
  plantas: {
    id: ['id', 'id planta', 'codigo', 'codigo planta'],
    nombre: ['nombre', 'planta', 'cultivo', 'nombre planta', 'nombre cultivo', 'especie'],
    cientifico: ['nombre cientifico', 'nombrecientifico', 'cientifico', 'especie cientifica', 'taxonomia'],
    tipo: ['tipo', 'categoria', 'familia', 'grupo', 'clasificacion'],
    clima: ['clima', 'clima recomendado', 'ambiente'],
    suelo: ['suelo', 'suelo recomendado', 'tipo suelo', 'sustrato'],
    riego: ['riego', 'riego recomendado', 'frecuencia riego'],
    temporada: ['temporada', 'epoca', 'epoca de siembra', 'siembra', 'cosecha'],
    notas: ['notas', 'observaciones', 'detalle', 'cuidados']
  },
  ventas: {
    id: ['id', 'id venta', 'codigo', 'codigo venta', 'folio', 'boleta', 'factura'],
    fecha: ['fecha', 'fecha venta', 'dia', 'date'],
    cliente: ['cliente', 'nombre cliente', 'comprador', 'razon social'],
    producto: ['producto', 'articulo', 'item', 'insumo', 'descripcion'],
    cantidad: ['cantidad', 'unidades', 'qty', 'cant'],
    precioUnitario: ['precio unitario', 'preciounitario', 'precio', 'valor unitario', 'p unitario', 'precio venta'],
    total: ['total', 'monto', 'importe', 'total clp', 'valor total', 'subtotal'],
    estado: ['estado', 'situacion', 'status'],
    canal: ['canal', 'medio', 'origen', 'vendedor', 'punto de venta']
  },
  asesorias: {
    id: ['id', 'id asesoria', 'codigo', 'codigo asesoria', 'ticket'],
    cliente: ['cliente', 'nombre cliente', 'solicitante', 'contacto'],
    cultivo: ['cultivo', 'planta', 'especie', 'producto'],
    consulta: ['consulta', 'problema', 'descripcion', 'detalle', 'motivo', 'requerimiento', 'mensaje'],
    estado: ['estado', 'situacion', 'status'],
    prioridad: ['prioridad', 'urgencia', 'nivel'],
    recomendacion: ['recomendacion', 'respuesta', 'solucion', 'sugerencia', 'notas']
  },
  proveedores: {
    id: ['id', 'id proveedor', 'codigo', 'codigo proveedor', 'rut'],
    proveedor: ['proveedor', 'nombre', 'nombre proveedor', 'razon social', 'empresa', 'contacto'],
    contacto: ['contacto', 'nombre contacto', 'vendedor', 'representante', 'persona'],
    telefono: ['telefono', 'fono', 'celular', 'movil', 'whatsapp', 'tel'],
    email: ['email', 'correo', 'correo electronico', 'mail'],
    ciudad: ['ciudad', 'comuna', 'localidad', 'region', 'direccion'],
    productos: ['productos', 'rubro', 'categoria', 'insumos', 'linea'],
    condiciones: ['condiciones', 'condiciones de pago', 'plazo', 'notas', 'observaciones']
  },
  pedidos: {
    id: ['id', 'id pedido', 'codigo', 'folio', 'numero'],
    fecha: ['fecha', 'dia', 'date'],
    cliente: ['cliente', 'nombre cliente', 'comprador'],
    producto: ['producto', 'articulo', 'item', 'insumo'],
    cantidad: ['cantidad', 'unidades', 'qty', 'cant'],
    total: ['total', 'monto', 'importe', 'valor total'],
    estado: ['estado', 'situacion', 'status'],
    notas: ['notas', 'observaciones', 'detalle']
  }
}

const ALL_ALIASES = (() => {
  const merged = { ...PRODUCT_ALIASES }
  Object.values(RECORD_ALIASES).forEach((aliases) => {
    Object.entries(aliases).forEach(([field, list]) => {
      merged[field] = Array.from(new Set([...(merged[field] || []), ...list]))
    })
  })
  return merged
})()

const FIELD_RULES = {
  previousPrice: { require: /anterior|antes|normal|regular|lista|antiguo/ },
  price: { reject: /anterior|antes|normal|regular|lista|antiguo/ },
  stock: { reject: /minimo|minima|min/ }
}

function hasValue(value) {
  if (value === null || value === undefined) return false
  if (typeof value === 'number') return Number.isFinite(value)
  if (typeof value === 'boolean') return true
  if (value instanceof Date) return !Number.isNaN(value.getTime())
  const text = String(value).replace(/\u00a0/g, ' ').trim()
  if (!text || /^[-–—]+$/.test(text)) return false
  if (/^(n\/?a|sin dato|no aplica|pendiente|sin informacion|sin información)$/i.test(text)) return false
  if (/^[\s$.,;:/\\|-]*$/.test(text)) return false
  if (/^(clp|ars|usd|peso|pesos|monto|precio|valor)$/i.test(text)) return false
  return /[a-z0-9]/i.test(text.normalize('NFD').replace(/[\u0300-\u036f]/g, ''))
}

function readNumber(value) {
  return hasValue(value) ? toNumber(value) : 0
}

function formatDateCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return getTemporalContext(value).fechaISO
  if (typeof value === 'number' && value > 20000 && value < 90000) {
    const date = XLSX.SSF.parse_date_code(value)
    if (date) return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`
  }
  return cleanCell(value)
}

function aliasScore(header, alias) {
  if (!header || !alias) return 0
  if (header === alias) return 1000 + alias.length
  const headerWords = header.split(' ').filter(Boolean)
  const aliasWords = alias.split(' ').filter(Boolean)
  if (` ${header} `.includes(` ${alias} `)) return 600 + alias.length
  if (aliasWords.length > 1 && aliasWords.every((word) => headerWords.includes(word))) return 520 + alias.length
  if (header.length >= 4 && alias.includes(header)) return 320 + header.length
  const overlap = aliasWords.filter((word) => headerWords.includes(word)).length
  if (overlap) return 100 + overlap * 20 + Math.min(header.length, alias.length)
  return 0
}

function fieldAllowed(field, header) {
  const rule = FIELD_RULES[field]
  if (!rule) return true
  if (rule.reject && rule.reject.test(header)) return false
  if (rule.require && !rule.require.test(header)) return false
  return true
}

export function matchColumns(headerRow, aliases) {
  const headers = headerRow.map((cell) => normalizeHeader(cell))
  const mapping = {}
  const used = new Set()
  const reuse = new Set(['usage', 'features'])
  const order = ['id', 'name', 'previousPrice', 'price', 'stock', 'category', 'format', 'description', 'usage', 'recommendations', 'features', 'image']
  const fields = order.filter((field) => aliases[field])
  Object.keys(aliases).forEach((field) => { if (!fields.includes(field)) fields.push(field) })
  fields.forEach((field) => {
    let best = { index: -1, score: 0 }
    headers.forEach((header, index) => {
      if (!header || (used.has(index) && !reuse.has(field))) return
      if (!fieldAllowed(field, header)) return
      aliases[field].forEach((alias) => {
        const score = aliasScore(header, normalizeHeader(alias))
        if (score > best.score) best = { index, score }
      })
    })
    if (best.index >= 0) {
      mapping[field] = best.index
      used.add(best.index)
    }
  })
  return mapping
}

export function findHeaderRow(matrix, aliases, maxRows = 15) {
  const limit = Math.min(matrix.length, maxRows)
  let best = null
  for (let index = 0; index < limit; index += 1) {
    const row = matrix[index] || []
    const nonEmpty = row.filter(hasValue).length
    if (nonEmpty < 1) continue
    const mapping = matchColumns(row, aliases)
    const recognized = Object.keys(mapping).length
    const strong = Boolean(mapping.name || mapping.producto || mapping.precio || mapping.price || mapping.stock || mapping.id)
    if (recognized < 2 && !(recognized === 1 && strong && nonEmpty <= 2)) continue
    const score = recognized * 10 + Math.min(nonEmpty, 10) + (mapping.name ? 6 : 0) + (mapping.price || mapping.stock ? 4 : 0)
    if (!best || score > best.score) best = { index, mapping, recognized, score, headers: row }
  }
  return best
}

function aliasesForKind(kind) {
  if (kind === 'productos' || kind === 'inventario') return PRODUCT_ALIASES
  return RECORD_ALIASES[kind] || ALL_ALIASES
}

function detectKindByName(name) {
  const value = normalizeHeader(name)
  if (!value) return ''
  if (/pedido/.test(value)) return 'pedidos'
  if (/proveedor|supplier/.test(value)) return 'proveedores'
  if (/asesor|consulta/.test(value)) return 'asesorias'
  if (/cliente|directorio/.test(value)) return 'clientes'
  if (/planta|cultivo|vivero/.test(value)) return 'plantas'
  // "inventario" contiene la cadena "venta": debe evaluarse antes que ventas.
  if (/inventario|existencia|bodega|stock/.test(value)) return 'inventario'
  if (/venta|factura|boleta/.test(value)) return 'ventas'
  if (/producto|catalogo|articulo|insumo|precio/.test(value)) return 'productos'
  return ''
}

function inferKind(mapping) {
  if (mapping.proveedor) return 'proveedores'
  if (mapping.consulta || mapping.cultivo) return 'asesorias'
  if (mapping.cliente && mapping.producto) return 'ventas'
  if (mapping.cliente) return 'clientes'
  if (mapping.cientifico || mapping.planta) return 'plantas'
  if (mapping.name && (mapping.price || mapping.stock || mapping.category)) return 'productos'
  if (mapping.name) return 'productos'
  return ''
}

function rowObject(row, mapping) {
  const raw = {}
  Object.entries(mapping).forEach(([field, index]) => { raw[field] = row[index] ?? '' })
  return raw
}

function productFromRaw(raw, origin = 'Importado desde Excel') {
  const pendingFields = []
  const pricePresent = hasValue(raw.price)
  const stockPresent = hasValue(raw.stock)
  if (!pricePresent) pendingFields.push('price')
  if (!stockPresent) pendingFields.push('stock')
  let id = cleanCell(raw.id)
  let name = cleanCell(raw.name)
  if (!name && id) name = `Producto ${id}`
  if (!name) return null
  // Filas de totales importadas por error desde Excel ("Total Valorado
  // Inventario", "Total", "Subtotal"...): nunca son productos reales.
  if (isJunkProductName(name)) return null
  if (!id) id = makeId('producto')
  const source = cleanCell(raw.source) || origin
  const product = {
    id,
    name,
    category: cleanCell(raw.category),
    price: readNumber(raw.price),
    previousPrice: readNumber(raw.previousPrice),
    stock: readNumber(raw.stock),
    format: cleanCell(raw.format),
    description: cleanCell(raw.description),
    usage: cleanCell(raw.usage) || cleanCell(raw.description),
    recommendations: cleanCell(raw.recommendations),
    features: cleanCell(raw.features),
    image: cleanCell(raw.image),
    source,
    status: pendingFields.length ? 'Pendiente de completar' : 'Completo',
    pendingFields
  }
  return product
}

function parseProductRows(rows, mapping, origin) {
  const products = []
  let invalid = 0
  rows.forEach((row) => {
    if (!row.some(hasValue)) return
    const product = productFromRaw(rowObject(row, mapping), origin)
    if (!product) { invalid += 1; return }
    products.push(product)
  })
  return { products, invalid }
}

function recordFromRaw(kind, raw) {
  const record = {}
  const aliases = RECORD_ALIASES[kind] || {}
  const numeric = new Set(['cantidad', 'precioUnitario', 'total'])
  Object.keys(aliases).forEach((field) => {
    if (field === 'id') return
    const value = raw[field]
    record[field] = numeric.has(field) ? (hasValue(value) ? toNumber(value) : '') : cleanCell(value)
  })
  if (record.fecha !== undefined) record.fecha = formatDateCell(raw.fecha)
  const primary = kind === 'proveedores' ? 'proveedor' : kind === 'plantas' ? 'nombre' : 'cliente'
  if (!record[primary] && primary !== 'nombre' && record.nombre) record[primary] = record.nombre
  // No se importan filas de resumen ni registros sin un nombre real: son
  // residuos de Excel que ensucian Clientes, Ventas, Asesorías, etc.
  if (!record[primary] || isJunkProductName(record[primary])) return null
  record.id = cleanCell(raw.id) || makeId(kind.slice(0, 3))
  return record
}

function parseRecordRows(rows, mapping, kind) {
  const records = []
  rows.forEach((row) => {
    if (!row.some(hasValue)) return
    const record = recordFromRaw(kind, rowObject(row, mapping))
    if (record) records.push(record)
  })
  return records
}

function dedupeProducts(list) {
  const byId = new Map()
  const byName = new Map()
  const output = []
  list.forEach((product) => {
    const idKey = normalizeHeader(product.id)
    const nameKey = normalizeHeader(product.name)
    const existingIndex = byId.has(idKey) ? byId.get(idKey) : byName.has(nameKey) ? byName.get(nameKey) : -1
    if (existingIndex >= 0) {
      const current = output[existingIndex]
      const pending = new Set(current.pendingFields || [])
      const incomingPending = new Set(product.pendingFields || [])
      ;['price', 'stock'].forEach((field) => {
        if (pending.has(field) && !incomingPending.has(field)) pending.delete(field)
      })
      const merged = { ...current }
      Object.entries(product).forEach(([field, value]) => {
        if (['id', 'status', 'pendingFields'].includes(field)) return
        if (field === 'price' || field === 'stock') {
          // Solo se conserva el valor previo cuando la fila entrante no aporta el dato.
          if (incomingPending.has(field) && !pending.has(field)) merged[field] = current[field]
          else merged[field] = value
          return
        }
        if (field === 'source') {
          if (!current.source || current.source === 'Inventario Excel') merged[field] = value
          return
        }
        if (value !== '' && value !== null && value !== undefined) merged[field] = value
      })
      merged.pendingFields = [...pending]
      merged.status = merged.pendingFields.length ? 'Pendiente de completar' : 'Completo'
      output[existingIndex] = merged
    } else {
      if (idKey) byId.set(idKey, output.length)
      if (nameKey) byName.set(nameKey, output.length)
      output.push(product)
    }
  })
  return output
}

export function parseWorkbookImport(workbook, options = {}) {
  const maxHeaderRows = Number(options.maxHeaderRows) || 15
  const result = {
    products: [],
    records: { clientes: [], plantas: [], ventas: [], asesorias: [], proveedores: [], pedidos: [] },
    sheets: [],
    rows: 0,
    rawRows: 0,
    invalid: 0,
    pending: 0,
    irregularSheets: [],
    needsAi: false,
    rawText: ''
  }
  if (!workbook?.SheetNames?.length) return result
  let rawRows = 0
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) return
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false })
    if (!matrix.length) return
    rawRows += matrix.length
    let kind = detectKindByName(sheetName)
    let aliases = aliasesForKind(kind || '')
    let header = findHeaderRow(matrix, aliases, maxHeaderRows)
    if (!header) {
      aliases = ALL_ALIASES
      header = findHeaderRow(matrix, aliases, maxHeaderRows)
    }
    if (!header) {
      result.irregularSheets.push({ name: sheetName, rows: matrix.length })
      return
    }
    let mapping = header.mapping
    if (!kind) kind = inferKind(mapping)
    if (!kind) {
      result.irregularSheets.push({ name: sheetName, rows: matrix.length })
      return
    }
    if (kind === 'productos' && !mapping.name) {
      const retry = findHeaderRow(matrix, ALL_ALIASES, maxHeaderRows)
      if (retry?.mapping?.name) { header = retry; mapping = retry.mapping }
    }
    const dataRows = matrix.slice(header.index + 1)
    const nonEmptyRows = dataRows.filter((row) => row.some(hasValue))
    let imported = 0
    let invalid = 0
    if (kind === 'productos' || kind === 'inventario') {
      const parsed = parseProductRows(dataRows, mapping, kind === 'inventario' ? 'Inventario Excel' : 'Importado desde Excel')
      result.products.push(...parsed.products)
      imported = parsed.products.length
      invalid = parsed.invalid
    } else {
      const parsed = parseRecordRows(dataRows, mapping, kind)
      result.records[kind].push(...parsed)
      imported = parsed.length
    }
    result.rows += nonEmptyRows.length
    result.invalid += invalid
    result.sheets.push({
      name: sheetName,
      kind,
      headerRow: header.index + 1,
      recognized: header.recognized,
      imported,
      irregular: false
    })
  })
  result.products = dedupeProducts(result.products)
  result.pending = result.products.filter((product) => product.status === 'Pendiente de completar').length
  const recordCount = Object.values(result.records).reduce((sum, list) => sum + list.length, 0)
  result.needsAi = !result.products.length && !recordCount && rawRows > 0
  result.rawRows = rawRows
  result.rawText = workbookToRawText(workbook)
  return result
}

export function workbookToRawText(workbook, maxChars = 12000) {
  if (!workbook?.SheetNames?.length) return ''
  const chunks = []
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) return
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false })
    if (!matrix.length) return
    const lines = matrix.slice(0, 40).map((row) => row.map((cell) => {
      if (cell instanceof Date) return formatDateCell(cell)
      return cleanCell(cell)
    }).join(' | ')).filter((line) => line.replace(/[|\s]/g, ''))
    if (lines.length) chunks.push(`HOJA "${sheetName}":\n${lines.join('\n')}`)
  })
  return chunks.join('\n\n').slice(0, maxChars)
}

export function mergeImportedProducts(existing = [], incoming = []) {
  const list = Array.isArray(existing) ? [...existing] : []
  const byId = new Map()
  const byName = new Map()
  const refreshIndexes = () => {
    byId.clear()
    byName.clear()
    list.forEach((product, index) => {
      const idKey = normalizeHeader(product.id)
      const nameKey = normalizeHeader(product.name)
      if (idKey) byId.set(idKey, index)
      if (nameKey) byName.set(nameKey, index)
    })
  }
  refreshIndexes()
  let added = 0
  let updated = 0
  incoming.forEach((product) => {
    if (!product?.name) return
    const idKey = normalizeHeader(product.id)
    const nameKey = normalizeHeader(product.name)
    const existingIndex = byId.has(idKey) ? byId.get(idKey) : byName.has(nameKey) ? byName.get(nameKey) : -1
    if (existingIndex < 0) {
      list.push(product)
      added += 1
      refreshIndexes()
      return
    }
    const previous = list[existingIndex]
    const merged = { ...previous }
    const pendingFields = []
    Object.entries(product).forEach(([field, value]) => {
      if (['id', 'status', 'pendingFields'].includes(field)) return
      if (field === 'price' || field === 'stock') {
        const incomingPending = product.pendingFields?.includes(field)
        const previousPending = previous.pendingFields?.includes(field) || !hasValue(previous[field])
        if (incomingPending) {
          if (previousPending) {
            merged[field] = 0
            pendingFields.push(field)
          } else {
            merged[field] = previous[field]
          }
        } else {
          merged[field] = value
        }
        return
      }
      if (field === 'source' && previous.source && product.source === 'Importado desde Excel') return
      if (value !== '' && value !== null && value !== undefined) merged[field] = value
    })
    merged.pendingFields = pendingFields
    merged.status = pendingFields.length ? 'Pendiente de completar' : 'Completo'
    list[existingIndex] = merged
    updated += 1
    refreshIndexes()
  })
  return { products: list, added, updated }
}

function recordKey(kind, record) {
  // Se priorizan claves naturales para que un Excel se cruce con registros creados
  // manualmente aunque el ID sea un UUID interno o un RUT distinto.
  const id = normalizeHeader(record.id)
  const natural = (value) => normalizeHeader(value)
  if (kind === 'proveedores') {
    const name = natural(record.proveedor)
    return name ? `name:${name}` : `id:${id}`
  }
  if (kind === 'plantas') {
    const name = natural(record.nombre)
    return name ? `name:${name}` : `id:${id}`
  }
  if (kind === 'ventas' || kind === 'pedidos') {
    return `sale:${natural(record.fecha)}|${natural(record.cliente)}|${natural(record.producto)}|${natural(record.total)}`
  }
  if (kind === 'asesorias') {
    return `case:${natural(record.cliente)}|${natural(record.cultivo)}|${natural(record.consulta)}`
  }
  const name = natural(record.nombre || record.cliente)
  if (name) return `name:${name}`
  const email = natural(record.email)
  if (email) return `email:${email}`
  const phone = natural(record.telefono)
  if (phone) return `phone:${phone}`
  return `id:${id}`
}

export function mergeImportedRecords(existing = [], incoming = [], kind = 'clientes') {
  const list = Array.isArray(existing) ? [...existing] : []
  const indexByKey = new Map()
  list.forEach((record, index) => indexByKey.set(recordKey(kind, record), index))
  let added = 0
  let updated = 0
  incoming.forEach((record) => {
    if (!record) return
    const key = recordKey(kind, record)
    const index = indexByKey.get(key)
    if (index === undefined) {
      list.push(record)
      indexByKey.set(key, list.length - 1)
      added += 1
      return
    }
    const merged = { ...list[index], ...record, id: list[index].id || record.id }
    list[index] = merged
    updated += 1
  })
  return { records: list, added, updated }
}

export function normalizeImportedProducts(list = []) {
  if (!Array.isArray(list)) return []
  return list.map((item) => {
    const mapping = matchColumns(Object.keys(item || {}), PRODUCT_ALIASES)
    const raw = {}
    Object.entries(mapping).forEach(([field, index]) => { raw[field] = Object.values(item)[index] })
    return productFromRaw(raw, 'Importado desde Excel (IA)')
  }).filter(Boolean)
}

export function normalizeImportedRecords(kind, list = []) {
  if (!Array.isArray(list)) return []
  const aliases = RECORD_ALIASES[kind] || {}
  return list.map((item) => {
    const mapping = matchColumns(Object.keys(item || {}), aliases)
    const raw = {}
    Object.entries(mapping).forEach(([field, index]) => { raw[field] = Object.values(item)[index] })
    return recordFromRaw(kind, raw)
  }).filter(Boolean)
}

export function productToRow(product) {
  return [
    cleanCell(product.id),
    cleanCell(product.name),
    cleanCell(product.category),
    toNumber(product.price) || '',
    toNumber(product.previousPrice) || '',
    product.stock === '' || product.stock === null || product.stock === undefined ? '' : toNumber(product.stock),
    cleanCell(product.format),
    cleanCell(product.description),
    cleanCell(product.usage),
    cleanCell(product.recommendations),
    cleanCell(product.features),
    cleanCell(product.source || (product.generated ? 'Generado por IA' : 'Catálogo VITIMAX'))
  ]
}

export function buildDemoProducts() {
  const base = [
    ['VITIMAX VerdeRaíz DEMO', 'Bioestimulantes', 8990, 9990, 36, '1 litro', 'Bioestimulante de ejemplo para desarrollo radicular.', 'Aplicar según ficha técnica del fabricante. Dosis de ejemplo: 2 ml por litro de agua.', 'Producto DEMO: no usar en cultivos reales sin validación.', 'Ejemplo generado por IA'],
    ['VITIMAX FrutoPlus DEMO', 'Fertilizantes', 12490, '', 24, '500 ml', 'Fertilizante foliar de ejemplo para etapa de fructificación.', 'Dosis de ejemplo: 1,5 ml por litro de agua cada 15 días.', 'Producto DEMO: validar con un especialista antes de aplicar.', 'Ejemplo generado por IA'],
    ['VITIMAX HojaSana DEMO', 'Cuidado foliar', 7490, '', 48, '1 litro', 'Solución foliar de ejemplo para mantener el follaje limpio.', 'Dosis de ejemplo: 2 ml por litro de agua en pulverización foliar.', 'Producto DEMO: no reemplaza un diagnóstico fitosanitario.', 'Ejemplo generado por IA'],
    ['VITIMAX TierraActiva DEMO', 'Suelos', 15990, 17990, 12, '2 kg', 'Acondicionador de suelo de ejemplo para mejorar la estructura.', 'Dosis de ejemplo: 100 g por metro cuadrado e incorporar al suelo.', 'Producto DEMO: ajustar según análisis de suelo.', 'Ejemplo generado por IA'],
    ['VITIMAX CosechaPro DEMO', 'Nutrición vegetal', 10990, '', 30, '1 litro', 'Nutrición vegetal de ejemplo para acompañar la etapa productiva.', 'Dosis de ejemplo: 3 ml por litro de agua cada 20 días.', 'Producto DEMO: consultar ficha técnica oficial.', 'Ejemplo generado por IA'],
    ['VITIMAX RaízFuerte DEMO', 'Bioestimulantes', 9490, '', 19, '250 ml', 'Presentación pequeña de ejemplo para propagación y trasplante.', 'Dosis de ejemplo: 1 ml por litro de agua al trasplantar.', 'Producto DEMO: no usar en cultivos reales sin validación.', 'Ejemplo generado por IA']
  ]
  return base.map(([name, category, price, previousPrice, stock, format, description, usage, recommendations, features]) => ({
    id: makeId('demo'),
    name, category, price, previousPrice, stock, format, description, usage, recommendations, features,
    image: '', source: 'Generado por IA (ejemplo)'
  }))
}

function demoRecords() {
  return {
    clientes: [
      { id: makeId('cli'), nombre: 'Cliente de ejemplo · Agro Los Aromos', telefono: '+56 9 1111 1111', email: 'contacto@ejemplo.cl', localidad: 'Rancagua', notas: 'Registro de ejemplo generado por IA.' },
      { id: makeId('cli'), nombre: 'Cliente de ejemplo · Huerto El Maitén', telefono: '+56 9 2222 2222', email: 'huerto@ejemplo.cl', localidad: 'Talca', notas: 'Registro de ejemplo generado por IA.' }
    ],
    pedidos: [
      { id: makeId('ped'), cliente: 'Agro Los Aromos', producto: 'VITIMAX VerdeRaíz DEMO', cantidad: 4, total: 35960, estado: 'Confirmado', notas: 'Pedido de ejemplo generado por IA.' },
      { id: makeId('ped'), cliente: 'Huerto El Maitén', producto: 'VITIMAX FrutoPlus DEMO', cantidad: 2, total: 24980, estado: 'Pendiente', notas: 'Pedido de ejemplo generado por IA.' }
    ],
    asesorias: [
      { id: makeId('ase'), cliente: 'Agro Los Aromos', cultivo: 'Tomate', consulta: 'Ejemplo: evaluar plan de nutrición para la temporada.', estado: 'Abierta', prioridad: 'Normal', recomendacion: 'Revisar análisis de suelo antes de definir dosis.' },
      { id: makeId('ase'), cliente: 'Huerto El Maitén', cultivo: 'Vid', consulta: 'Ejemplo: seguimiento de desarrollo radicular.', estado: 'En revisión', prioridad: 'Alta', recomendacion: 'Validar producto y dosis con un especialista.' }
    ],
    contenido: []
  }
}

function plantRows() {
  return [
    ['Planta de ejemplo · Tomate', 'Solanum lycopersicum', 'Hortaliza', 'Templado / invernadero', 'Franco, rico en materia orgánica', 'Riego regular sin encharcar', 'Primavera - verano', 'Datos de ejemplo generados por IA.'],
    ['Planta de ejemplo · Vid', 'Vitis vinifera', 'Frutal', 'Mediterráneo', 'Franco arenoso con buen drenaje', 'Riego controlado', 'Primavera - otoño', 'Datos de ejemplo generados por IA.'],
    ['Planta de ejemplo · Cítrico', 'Citrus sinensis', 'Frutal', 'Cálido', 'Franco con buen drenaje', 'Riego profundo y espaciado', 'Todo el año', 'Datos de ejemplo generados por IA.'],
    ['Planta de ejemplo · Lechuga', 'Lactuca sativa', 'Hortaliza', 'Templado', 'Suelto y húmedo', 'Riego frecuente y ligero', 'Otoño - primavera', 'Datos de ejemplo generados por IA.'],
    ['Planta de ejemplo · Papa', 'Solanum tuberosum', 'Tubérculo', 'Templado frío', 'Suelto, profundo y drenado', 'Riego moderado', 'Primavera', 'Datos de ejemplo generados por IA.']
  ]
}

function supplierRows() {
  return [
    ['Proveedor de ejemplo · Insumos del Valle', 'Área comercial', '+56 9 3333 3333', 'ventas@ejemplo.cl', 'Santiago', 'Fertilizantes y bioestimulantes', 'Pago a 30 días · entrega 5 días'],
    ['Proveedor de ejemplo · AgroSur', 'Distribuidor', '+56 9 4444 4444', 'contacto@agrosur-ejemplo.cl', 'Temuco', 'Suelos y enmiendas', 'Pago contra entrega · entrega 7 días'],
    ['Proveedor de ejemplo · Semillas y Raíces', 'Vivero', '+56 9 5555 5555', 'vivero@ejemplo.cl', 'Curicó', 'Plantas y material vegetal', 'Pedido mínimo 10 unidades']
  ]
}

function recordRows(list, columns, mapper) {
  return (list || []).map((item, index) => mapper(item, index)).filter((row) => row.some((cell) => cell !== ''))
}

export function defaultSheetRows(name, products = [], records = {}, options = {}) {
  const demo = Boolean(options.allowDemo)
  const safeProducts = products.length ? products : demo ? buildDemoProducts() : []
  const safeRecords = demo ? { ...demoRecords(), ...records } : records
  if (name === 'Productos') {
    return { columns: PRODUCT_COLUMNS, rows: safeProducts.filter((product) => !isJunkProductName(product?.name)).map(productToRow) }
  }
  if (name === 'Clientes') {
    return {
      columns: SHEET_COLUMNS.Clientes,
      rows: recordRows(safeRecords.clientes, SHEET_COLUMNS.Clientes, (item) => [
        cleanCell(item.id), cleanCell(item.nombre || item.cliente), cleanCell(item.telefono), cleanCell(item.email),
        cleanCell(item.localidad), cleanCell(item.notas)
      ])
    }
  }
  if (name === 'Plantas') {
    return { columns: SHEET_COLUMNS.Plantas, rows: demo ? plantRows() : [] }
  }
  if (name === 'Inventario') {
    return {
      columns: SHEET_COLUMNS.Inventario,
      rows: safeProducts.map((product) => {
        const stock = toNumber(product.stock)
        const price = toNumber(product.price)
        return [
          cleanCell(product.id), cleanCell(product.name), cleanCell(product.category), stock,
          options.minStock ?? 5, stock <= 0 ? 'Sin stock' : stock <= (options.minStock ?? 5) ? 'Stock bajo' : 'Disponible',
          price, price * stock
        ]
      })
    }
  }
  if (name === 'Ventas') {
    const sales = Array.isArray(records.pedidos) && records.pedidos.length
      ? records.pedidos
      : demo
        ? safeProducts.slice(0, 4).map((product, index) => ({
          id: makeId('venta'),
          fecha: new Date(Date.now() - index * 3 * 86400000).toISOString().slice(0, 10),
          cliente: `Cliente de ejemplo ${index + 1}`,
          producto: product.name,
          cantidad: 1 + (index % 3),
          precioUnitario: toNumber(product.price),
          total: toNumber(product.price) * (1 + (index % 3)),
          estado: 'Confirmado',
          canal: index % 2 ? 'WhatsApp' : 'Directo'
        }))
        : []
    return {
      columns: SHEET_COLUMNS.Ventas,
      rows: recordRows(sales, SHEET_COLUMNS.Ventas, (item) => [
        cleanCell(item.id), cleanCell(item.fecha || item.createdAt?.slice?.(0, 10)), cleanCell(item.cliente),
        cleanCell(item.producto), toNumber(item.cantidad), toNumber(item.precioUnitario ?? item.precio),
        toNumber(item.total), cleanCell(item.estado || 'Pendiente'), cleanCell(item.canal || 'Directo'),
        cleanCell(item.idCliente || item.clienteId || ''), cleanCell(item.idProducto || item.productoId || '')
      ])
    }
  }
  if (name === 'Asesorias') {
    const rawAdvisories = Array.isArray(records.asesorias) && records.asesorias.length
      ? records.asesorias
      : demo
        ? safeProducts.slice(0, 3).map((product, index) => ({
          id: makeId('asesoria'),
          cliente: `Cliente de ejemplo ${index + 1}`,
          cultivo: cropForProduct(product),
          consulta: `Ejemplo: plan de uso para ${product.name || 'producto VITIMAX'}.`,
          estado: index % 2 ? 'En revisión' : 'Abierta',
          prioridad: index === 0 ? 'Alta' : 'Normal',
          recomendacion: product.recommendations || 'Validar dosis con un especialista antes de aplicar.'
        }))
        : []
    const advisories = fixAsesoriaRecords(rawAdvisories, safeProducts)
    return {
      columns: SHEET_COLUMNS.Asesorias,
      rows: recordRows(advisories, SHEET_COLUMNS.Asesorias, (item) => [
        cleanCell(item.id), cleanCell(item.cliente), cleanCell(item.cultivo), cleanCell(item.consulta),
        cleanCell(item.estado), cleanCell(item.prioridad), cleanCell(item.recomendacion),
        cleanCell(item.idCliente || item.clienteId || '')
      ])
    }
  }
  if (name === 'Proveedores') {
    return { columns: SHEET_COLUMNS.Proveedores, rows: demo ? supplierRows() : [] }
  }
  return { columns: [], rows: [] }
}

// ---------------------------------------------------------------------------
// Cultivos reales para Asesorías (nunca el nombre del producto)
// ---------------------------------------------------------------------------

const CROP_RULES = [
  { crop: 'Vid', pattern: /vid\b|uva|vi[nñ]edo|parronal|vitis/i },
  { crop: 'Palto', pattern: /palto|aguacate|persea/i },
  { crop: 'Tomate', pattern: /tomate|tomatera|solanum lycopersicum/i },
  { crop: 'Cítricos', pattern: /c[ií]trico|naranja|lim[oó]n|mandarin|pomelo|citrus/i },
  { crop: 'Arándanos', pattern: /ar[aá]ndano|blueberr/i },
  { crop: 'Frutales', pattern: /frutal|manzano|peral|cerezo|durazno|kiwi|frambuesa|ciruelo|nogal/i },
  { crop: 'Hortalizas', pattern: /hortaliza|lechuga|pimiento|zanahoria|cebolla|papa|repollo|br[oó]coli/i },
  { crop: 'Praderas', pattern: /pradera|pastizal|forraj|alfalfa|ballica/i },
  { crop: 'Césped', pattern: /c[eé]sped|green|golf/i },
  { crop: 'Ornamentales', pattern: /ornamental|flor|rosa|jard[ií]n|vivero/i }
]

export function cropForProduct(product = {}, extraText = '') {
  const haystack = [
    product.name, product.category, product.description,
    product.usage, product.features, product.recommendations, extraText
  ].filter(Boolean).join(' ')
  const rule = CROP_RULES.find((item) => item.pattern.test(haystack))
  return rule ? rule.crop : 'Cultivo general'
}

const PRODUCT_LIKE = /vitimax|bioestimul|fertiliz|enmienda|nutrici[oó]n|foliar|radicular|correct|abono|algas|amino/i

function productResolver(products = []) {
  const list = (products || []).filter((product) => product && product.name && !isJunkProductName(product.name))
  const byName = new Map()
  list.forEach((product) => byName.set(normalizeHeader(product.name), product))
  return {
    list,
    byName,
    find(value, context = '') {
      const key = normalizeHeader(value)
      if (key && byName.has(key)) return byName.get(key)
      const haystack = normalizeHeader(`${value} ${context}`)
      const byProduct = list.find((product) => {
        const nameKey = normalizeHeader(product.name)
        return nameKey && haystack.includes(nameKey)
      })
      if (byProduct) return byProduct
      const byCategory = list.find((product) => {
        const categoryKey = normalizeHeader(product.category)
        return categoryKey && haystack.includes(categoryKey)
      })
      return byCategory || null
    }
  }
}

function recommendationText(product, fallback = '') {
  if (!product) {
    return cleanCell(fallback) || 'Recomendación pendiente: validar dosis con un especialista VITIMAX.'
  }
  const detail = cleanCell(product.recommendations) || cleanCell(product.usage) || cleanCell(product.description)
  return detail
    ? `Se sugiere ${product.name}: ${detail}`
    : `Se sugiere ${product.name} según la dosis indicada en su ficha técnica.`
}

function suggestedProduct(resolver, matched, crop, context) {
  return matched
    || resolver.list.find((product) => cropForProduct(product, context) === crop)
    || resolver.list[0]
    || null
}

// Corrige una fila de Asesorías: "Cultivo" muestra el cultivo afectado (nunca
// el producto) y "Recomendación" siempre indica el producto VITIMAX sugerido.
function fixAsesoriaRows(columns, rows, products = []) {
  const cols = Array.isArray(columns) ? columns : []
  if (!rows.length) return rows
  let cropIndex = cols.findIndex((column) => /cultivo|planta|especie/i.test(normalizeHeader(column)))
  let recIndex = cols.findIndex((column) => /recomendacion|respuesta|solucion|sugerencia/i.test(normalizeHeader(column)))
  if (cropIndex < 0) { cropIndex = cols.length; cols.push('Cultivo') }
  if (recIndex < 0) { recIndex = cols.length; cols.push('Recomendación') }
  const resolver = productResolver(products)
  return rows.map((row) => {
    const next = Array.isArray(row) ? [...row] : []
    const context = next.map((cell) => cleanCell(cell)).join(' ')
    const currentCrop = cleanCell(next[cropIndex])
    const currentRec = cleanCell(next[recIndex])
    const matched = resolver.find(currentCrop, context)
    const cropIsProduct = Boolean(matched)
      || resolver.byName.has(normalizeHeader(currentCrop))
      || (PRODUCT_LIKE.test(currentCrop) && !CROP_RULES.some((rule) => rule.pattern.test(currentCrop)))
    let crop = currentCrop
    if (!crop || cropIsProduct) crop = matched ? cropForProduct(matched, context) : cropForProduct({}, context)
    next[cropIndex] = crop || 'Cultivo general'
    if (!currentRec) next[recIndex] = recommendationText(suggestedProduct(resolver, matched, next[cropIndex], context), context)
    return next
  })
}

/**
 * Versión para registros de Asesorías en memoria/nube: garantiza cultivo real
 * y recomendación con el producto VITIMAX sugerido.
 */
export function fixAsesoriaRecords(records = [], products = []) {
  if (!Array.isArray(records) || !records.length) return records
  const resolver = productResolver(products)
  return records.map((record) => {
    if (!record) return record
    const next = { ...record }
    const context = Object.values(next).map((value) => cleanCell(value)).join(' ')
    const currentCrop = cleanCell(next.cultivo)
    const matched = resolver.find(currentCrop, context)
    const cropIsProduct = Boolean(matched)
      || resolver.byName.has(normalizeHeader(currentCrop))
      || (PRODUCT_LIKE.test(currentCrop) && !CROP_RULES.some((rule) => rule.pattern.test(currentCrop)))
    let crop = currentCrop
    if (!crop || cropIsProduct) crop = matched ? cropForProduct(matched, context) : cropForProduct({}, context)
    next.cultivo = crop || 'Cultivo general'
    if (!cleanCell(next.recomendacion)) {
      next.recomendacion = recommendationText(suggestedProduct(resolver, matched, next.cultivo, context), context)
    }
    return next
  })
}

// ---------------------------------------------------------------------------
// Limpieza de hojas: sin filas fantasma, sin duplicados y con totales reales
// ---------------------------------------------------------------------------

const TOTALS_ROW = /^(gran\s+)?(total|totales|subtotal|suma|resumen|total\s+general|total\s+valorado|total\s+inventario|valor\s+total|valorizado|valor\s+inventario)\b/i

// Nombre de fila/registro que corresponde a un total y no a un producto real.
// Cualquier nombre que empiece por "TOTAL" (o equivalente de total) es una
// fila de resumen, no un producto: se elimina aunque tenga texto añadido.
const JUNK_PRODUCT_NAME = /^\s*(?:gran\s+)?(?:total(?:es)?|subtotal(?:es)?|suma|resumen|valor(?:izado)?)\b/i

/**
 * Detecta filas de totales que llegan desde Excel (por ejemplo
 * "Total Valorado Inventario" o "Total") para no contarlas como productos.
 */
export function isJunkProductName(value) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim()
  if (!text) return false
  return JUNK_PRODUCT_NAME.test(text)
}

/** Filtra productos basura (totales) del catálogo en memoria o en la nube. */
export function filterJunkProducts(list = []) {
  return (Array.isArray(list) ? list : []).filter((product) => {
    if (!product) return false
    const name = product.name ?? product.nombre ?? product.producto ?? product.title ?? ''
    return !isJunkProductName(name)
  })
}

function isTotalsRow(row) {
  if (!Array.isArray(row)) return false
  // La fila de totales puede traer la etiqueta en cualquier columna, no solo
  // en la primera. Solo se descartan etiquetas cortas y autónomas.
  return row.some((cell) => {
    const text = String(cell ?? '').replace(/\s+/g, ' ').trim()
    if (!text || text.length > 48) return false
    return JUNK_PRODUCT_NAME.test(text)
  })
}

function sanitizeSheetRows(rows, { keepTotals = false } = {}) {
  const seen = new Set()
  return (Array.isArray(rows) ? rows : []).filter((row) => {
    const cells = Array.isArray(row) ? row : []
    if (!cells.some((cell) => String(cell ?? '').trim() !== '')) return false
    if (!keepTotals && isTotalsRow(cells)) return false
    const key = cells.map((cell) => String(cell ?? '').trim().toLowerCase()).join('|')
    if (key.replace(/\|/g, '')) {
      if (seen.has(key)) return false
      seen.add(key)
    }
    return true
  })
}

function sheetHeaderIndex(columns, pattern) {
  return (columns || []).findIndex((column) => pattern.test(normalizeHeader(column)))
}

function formulaSheetName(name = '') {
  const clean = String(name || '')
    .replace(/[\\/?*[\]:]/g, ' ')
    .replace(/'/g, "''")
    .trim()
    .slice(0, 31) || 'Hoja'
  return `'${clean}'`
}

function rangeRef(sheet, columnIndex, rowCount) {
  if (!sheet || columnIndex < 0 || !rowCount) return ''
  const letter = excelColumnLetter(columnIndex + 1)
  return `${formulaSheetName(sheet.name)}!${letter}2:${letter}${rowCount + 1}`
}

function workbookMetrics(sheets) {
  const find = (pattern) => sheets.find((sheet) => pattern.test(String(sheet?.name || '').trim()))
  const productsSheet = find(/^productos$/i) || find(/producto/i)
  const inventorySheet = find(/^inventario$/i) || find(/inventario/i)
  if (!productsSheet && !inventorySheet) return null

  const productRows = sanitizeSheetRows(productsSheet?.rows || [])
  const inventoryRows = sanitizeSheetRows(inventorySheet?.rows || [])
  const productName = sheetHeaderIndex(productsSheet?.columns, /nombre|producto|articulo|item|descripcion/)
  const productPrice = sheetHeaderIndex(productsSheet?.columns, /precio(?!.*anterior)/)
  const productStock = sheetHeaderIndex(productsSheet?.columns, /stock|cantidad|unidades|existencia/)
  const inventoryStock = sheetHeaderIndex(inventorySheet?.columns, /stock|cantidad|unidades|existencia/)
  const inventoryPrice = sheetHeaderIndex(inventorySheet?.columns, /precio/)
  const inventoryValue = sheetHeaderIndex(inventorySheet?.columns, /valor.*inventario|valorizado|valor\s+total|valor/)

  const prices = productPrice >= 0
    ? productRows.map((row) => toNumber(row[productPrice])).filter((value) => value > 0)
    : []
  const priceSum = prices.reduce((sum, value) => sum + value, 0)

  let units = 0
  if (inventorySheet && inventoryStock >= 0) {
    units = inventoryRows.reduce((sum, row) => sum + toNumber(row[inventoryStock]), 0)
  } else if (productStock >= 0) {
    units = productRows.reduce((sum, row) => sum + toNumber(row[productStock]), 0)
  }

  let inventoryTotal = 0
  if (inventorySheet && inventoryValue >= 0) {
    inventoryTotal = inventoryRows.reduce((sum, row) => sum + toNumber(row[inventoryValue]), 0)
  } else if (inventorySheet && inventoryPrice >= 0 && inventoryStock >= 0) {
    inventoryTotal = inventoryRows.reduce((sum, row) => sum + toNumber(row[inventoryPrice]) * toNumber(row[inventoryStock]), 0)
  } else if (productPrice >= 0 && productStock >= 0) {
    inventoryTotal = productRows.reduce((sum, row) => sum + toNumber(row[productPrice]) * toNumber(row[productStock]), 0)
  }

  const count = productRows.length || inventoryRows.length
  if (!count && !units && !inventoryTotal) return null

  // Fórmulas reales sobre el rango exacto de datos (sin filas de totales).
  const formulas = {}
  if (inventorySheet && inventoryStock >= 0) {
    formulas.units = `=SUM(${rangeRef(inventorySheet, inventoryStock, inventoryRows.length)})`
  } else if (productsSheet && productStock >= 0) {
    formulas.units = `=SUM(${rangeRef(productsSheet, productStock, productRows.length)})`
  }
  if (inventorySheet && inventoryValue >= 0) {
    formulas.inventoryTotal = `=SUM(${rangeRef(inventorySheet, inventoryValue, inventoryRows.length)})`
  } else if (inventorySheet && inventoryPrice >= 0 && inventoryStock >= 0) {
    formulas.inventoryTotal = `=SUMPRODUCT(${rangeRef(inventorySheet, inventoryPrice, inventoryRows.length)},${rangeRef(inventorySheet, inventoryStock, inventoryRows.length)})`
  } else if (productsSheet && productPrice >= 0 && productStock >= 0) {
    formulas.inventoryTotal = `=SUMPRODUCT(${rangeRef(productsSheet, productPrice, productRows.length)},${rangeRef(productsSheet, productStock, productRows.length)})`
  }
  if (productsSheet && productPrice >= 0 && productRows.length) {
    formulas.avgPrice = `=IFERROR(AVERAGE(${rangeRef(productsSheet, productPrice, productRows.length)}),0)`
  }
  if (productsSheet && productName >= 0 && productRows.length) {
    formulas.count = `=COUNTA(${rangeRef(productsSheet, productName, productRows.length)})`
  } else if (productsSheet && productRows.length) {
    formulas.count = `=ROWS(${formulaSheetName(productsSheet.name)}!A2:A${productRows.length + 1})`
  }

  return {
    count,
    units,
    inventoryTotal,
    priceSum,
    avgPrice: prices.length ? priceSum / prices.length : 0,
    formulas
  }
}

function reconcileSummarySheets(sheets) {
  const summary = sheets.filter((sheet) => /resumen|kpi|dashboard|indicador/i.test(String(sheet?.name || '')))
  if (!summary.length) return sheets
  const metrics = workbookMetrics(sheets)
  if (!metrics) return sheets
  const unitsFormula = metrics.formulas.units || ''
  const stockRange = unitsFormula.startsWith('=SUM(') ? unitsFormula.slice(5, -1) : ''
  // El orden importa: "Stock bajo" y "Estado" se resuelven antes que los
  // patrones genéricos de stock/inventario para no sobrescribir sus fórmulas.
  const definitions = [
    { pattern: /generad|fecha|actualiz/, label: 'Fecha de generación', value: formatSpanishDate(new Date()) },
    { pattern: /bajo|reposici[oó]n|cr[ií]tico/, label: 'Productos con stock bajo (5 o menos)', value: '', formula: stockRange ? `=COUNTIF(${stockRange},"<=5")` : '' },
    { pattern: /estado|operativ/, label: 'Estado del inventario', value: '', formula: unitsFormula ? `=IF(${unitsFormula}>0,"Operativo","Revisar existencias")` : '' },
    { pattern: /unidad|stock/, label: 'Unidades en stock', value: metrics.units, formula: metrics.formulas.units },
    { pattern: /inventario|valorizado|valor/, label: 'Valor inventario CLP', value: metrics.inventoryTotal, formula: metrics.formulas.inventoryTotal },
    { pattern: /precio|promedio/, label: 'Precio promedio CLP', value: metrics.avgPrice, formula: metrics.formulas.avgPrice },
    { pattern: /producto|cantidad/, label: 'Productos registrados', value: metrics.count, formula: metrics.formulas.count }
  ]
  return sheets.map((sheet) => {
    if (!summary.includes(sheet)) return sheet
    const columns = Array.isArray(sheet.columns) ? [...sheet.columns] : []
    if (columns.length < 2) return sheet
    const rows = (sheet.rows || []).map((row) => [...row])
    const used = new Set()
    definitions.forEach((definition) => {
      const value = definition.formula || definition.value
      const index = rows.findIndex((row, rowIndex) => (
        !used.has(rowIndex) && definition.pattern.test(normalizeHeader(row[0]))
      ))
      if (index >= 0) {
        used.add(index)
        // Si la IA ya escribió una fórmula real, se conserva intacta.
        if (!String(rows[index][1] ?? '').trim().startsWith('=')) rows[index][1] = value
        return
      }
      const next = new Array(columns.length).fill('')
      next[0] = definition.label
      next[1] = value
      rows.push(next)
    })
    return { ...sheet, rows }
  })
}

function normalizeAiSheet(raw) {
  if (!raw || typeof raw !== 'object') return null
  const name = cleanCell(raw.name || raw.nombre || raw.hoja || raw.sheet)
  let columns = raw.columns || raw.columnas || raw.headers || raw.encabezados
  let rows = raw.rows || raw.filas || raw.data || raw.datos
  if (!Array.isArray(rows)) rows = []
  if (!Array.isArray(columns)) {
    const first = rows.find((row) => row && typeof row === 'object' && !Array.isArray(row))
    columns = first ? Object.keys(first) : []
  }
  columns = columns.map((column) => cleanCell(column)).filter(Boolean)
  rows = rows.map((row) => {
    if (Array.isArray(row)) return columns.map((_, index) => cleanCell(row[index]))
    if (row && typeof row === 'object') return columns.map((column) => cleanCell(row[column]))
    return []
  }).filter((row) => row.some((cell) => cell !== ''))
  return { name, columns, rows }
}

// Todo libro generado por el asistente debe traer un Resumen Ejecutivo con
// fórmulas reales, incluso cuando el modelo solo pidió el catálogo de 7 hojas
// y no envió KPIs. Las referencias se calculan desde las columnas de la hoja
// Productos que realmente se va a escribir.
function ensureSummarySheet(sheets, products = []) {
  if (sheets.some((sheet) => /resumen|kpi|dashboard|indicador/i.test(sheet?.name || ''))) return sheets
  const productSheet = sheets.find((sheet) => /producto/i.test(sheet?.name || '')) || sheets[0] || {}
  const headers = (productSheet.columns || []).map((header) => String(header).toLowerCase())
  const indexOf = (pattern, fallback) => {
    const found = headers.findIndex((header) => pattern.test(header))
    return found >= 0 ? found : fallback
  }
  const nameIndex = indexOf(/nombre|producto|descripci/i, 0)
  const priceIndex = indexOf(/precio|valor|monto/i, 2)
  const stockIndex = indexOf(/stock|existencia|cantidad|unidades/i, 3)
  const lastRow = Math.max((productSheet.rows?.length || products.length || 0) + 1, 11)
  const range = (index) => {
    const letter = excelColumnLetter(index + 1)
    return `${letter}2:${letter}${lastRow}`
  }
  const nameRange = range(nameIndex)
  const priceRange = range(priceIndex)
  const stockRange = range(stockIndex)
  const summary = {
    name: 'Resumen Ejecutivo',
    columns: ['KPI', 'Valor', 'Unidad', 'Interpretación'],
    rows: [
      ['Productos registrados', `=COUNTA(Productos!${nameRange})`, 'unidades', 'Filas reales del catálogo VITIMAX'],
      ['Unidades en stock', `=SUM(Productos!${stockRange})`, 'unidades', 'Existencias disponibles'],
      ['Valor de inventario', `=SUMPRODUCT(Productos!${priceRange},Productos!${stockRange})`, 'CLP', 'Suma de precio por unidad en stock'],
      ['Precio promedio', `=AVERAGE(Productos!${priceRange})`, 'CLP', 'Promedio de precios del catálogo'],
      ['Productos con stock bajo (5 o menos)', `=COUNTIF(Productos!${stockRange},"<=5")`, 'unidades', 'Filas que requieren reposición'],
      ['Estado del inventario', `=IF(SUM(Productos!${stockRange})>0,"Operativo","Revisar existencias")`, 'estado', 'Lectura automática del stock total']
    ]
  }
  return [summary, ...sheets]
}

// ---------------------------------------------------------------------------
// Coherencia referencial: Productos y Clientes son la única fuente de
// IDs/nombres. Inventario, Ventas y Asesorías se reescriben desde esas hojas
// para que nunca aparezcan IDs de un esquema con nombres de otro.
// ---------------------------------------------------------------------------

function workbookSheetKey(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '')
}

function workbookCanonicalName(value) {
  return SHEET_NAMES.find((name) => workbookSheetKey(name) === workbookSheetKey(value)) || null
}

function columnIndex(columns, pattern, fallback = -1) {
  const found = (columns || []).findIndex((column) => pattern.test(normalizeHeader(column)))
  return found >= 0 ? found : fallback
}

function sequentialId(prefix, index) {
  return `${prefix}-${String(index + 1).padStart(3, '0')}`
}

function ensureColumn(sheet, header) {
  const columns = Array.isArray(sheet.columns) ? sheet.columns : []
  if (!columns.some((column) => normalizeHeader(column) === normalizeHeader(header))) columns.push(header)
  sheet.columns = columns
  return columns.findIndex((column) => normalizeHeader(column) === normalizeHeader(header))
}

// Ninguna venta generada puede quedar en el futuro respecto al reloj local.
function clampSaleDate(value) {
  const raw = cleanCell(value)
  if (!raw) return ''
  const today = getTemporalContext().fechaISO
  let iso = ''
  if (raw instanceof Date) iso = raw.toISOString().slice(0, 10)
  else {
    const head = String(raw).slice(0, 10)
    if (/^\d{4}-\d{2}-\d{2}$/.test(head)) iso = head
    else {
      const parsed = new Date(raw)
      if (!Number.isNaN(parsed.getTime())) iso = parsed.toISOString().slice(0, 10)
    }
  }
  if (!iso) return raw
  return iso > today ? today : iso
}

function catalogFromSheet(sheet, prefix) {
  const columns = Array.isArray(sheet?.columns) ? sheet.columns : []
  const rows = Array.isArray(sheet?.rows) ? sheet.rows : []
  const indices = {
    id: columnIndex(columns, /^id$|identificador|codigo/, 0),
    name: columnIndex(columns, /nombre|cliente|producto|razon social|proveedor/i, 1),
    category: columnIndex(columns, /categoria/i, -1),
    price: columnIndex(columns, /precio/i, -1),
    stock: columnIndex(columns, /stock|existencia/i, -1)
  }
  const entries = []
  const byId = new Map()
  const byName = new Map()
  const usedIds = new Set()
  rows.forEach((raw, position) => {
    const row = Array.isArray(raw) ? [...raw] : []
    const name = cleanCell(row[indices.name])
    if (!name || isJunkProductName(name)) return
    let id = cleanCell(row[indices.id])
    // Los IDs cortos tipo P01/C01 mezclados con nombres VITIMAX rompen la
    // coherencia: se regeneran con el prefijo canónico de la hoja.
    if (/^[pc]-?\d{1,2}$/i.test(id)) id = ''
    if (!id || usedIds.has(id)) {
      id = sequentialId(prefix, entries.length)
      while (usedIds.has(id)) id = sequentialId(prefix, entries.length + usedIds.size)
    }
    usedIds.add(id)
    row[indices.id] = id
    row[indices.name] = name
    raw.splice(0, raw.length, ...row)
    const entry = { id, name, row, position, indices }
    entries.push(entry)
    byId.set(normalizeHeader(id), entry)
    byName.set(normalizeHeader(name), entry)
  })
  sheet.rows = rows
  return { sheet, rows, columns, entries, indices, byId, byName }
}

function ensureCatalogFromReferences(sheet, prefix, references) {
  if (!Array.isArray(sheet.rows)) sheet.rows = []
  if (sheet.rows.some((row) => cleanCell(row[0]) || cleanCell(row[1]))) return catalogFromSheet(sheet, prefix)
  const names = []
  for (const { sheet: reference, indices } of references) {
    for (const row of Array.isArray(reference?.rows) ? reference.rows : []) {
      for (const index of indices) {
        const name = cleanCell(row[index])
        if (!name || isJunkProductName(name) || /^\d{4}-\d{2}-\d{2}/.test(name)) continue
        if (!names.some((item) => normalizeHeader(item) === normalizeHeader(name))) names.push(name)
      }
    }
  }
  const columns = Array.isArray(sheet.columns) && sheet.columns.length ? sheet.columns : ['ID', 'Nombre']
  sheet.columns = columns
  sheet.rows = names.map((name, index) => {
    const row = Array(columns.length).fill('')
    row[0] = sequentialId(prefix, index)
    row[1] = name
    return row
  })
  return catalogFromSheet(sheet, prefix)
}

function resolveCatalogEntry(catalog, value, fallbackIndex = 0) {
  const key = normalizeHeader(value)
  const found = (key && (catalog.byName.get(key) || catalog.byId.get(key))) || null
  if (found) return found
  return catalog.entries[fallbackIndex] || catalog.entries[0] || null
}

function normalizeSequentialSheet(sheet, prefix) {
  if (!Array.isArray(sheet.columns)) sheet.columns = []
  if (!sheet.columns.length || normalizeHeader(sheet.columns[0]) !== 'id') sheet.columns = ['ID', ...sheet.columns]
  const valueCount = Math.max(sheet.columns.length - 1, 0)
  sheet.rows = (Array.isArray(sheet.rows) ? sheet.rows : []).map((raw, index) => {
    const row = Array.isArray(raw) ? [...raw] : []
    const first = cleanCell(row[0])
    const firstIsId = Boolean(first) && /^(?:(?:plt|pla|plant|prv|prov|id)(?:[-\s]*\d+)?(?:[-\s]|$)|p-?\d+$|[0-9a-f]{8}-)/i.test(first)
    const values = firstIsId ? row.slice(1, 1 + valueCount) : row.slice(0, valueCount)
    while (values.length < valueCount) values.push('')
    return [sequentialId(prefix, index), ...values]
  }).filter((row) => row.slice(1).some((cell) => cleanCell(cell)))
}

function rebuildInventorySheet(sheet, products) {
  const idIndex = columnIndex(sheet.columns, /^id$/, 0)
  const nameIndex = columnIndex(sheet.columns, /producto|nombre/i, 1)
  const categoryIndex = columnIndex(sheet.columns, /categoria/i, 2)
  const stockIndex = columnIndex(sheet.columns, /stock actual|stock|existencia/i, 3)
  const minimumIndex = columnIndex(sheet.columns, /stock minimo|minimo/i, 4)
  const statusIndex = columnIndex(sheet.columns, /estado/i, 5)
  const priceIndex = columnIndex(sheet.columns, /precio/i, 6)
  const valueIndex = columnIndex(sheet.columns, /valor/i, 7)
  const existing = new Map()
  for (const row of Array.isArray(sheet.rows) ? sheet.rows : []) {
    const entry = resolveCatalogEntry(products, cleanCell(row[idIndex]), -1)
      || resolveCatalogEntry(products, cleanCell(row[nameIndex]), -1)
    if (entry) existing.set(entry.id, row)
  }
  sheet.rows = products.entries.map((entry) => {
    const row = Array(sheet.columns.length).fill('')
    const previous = existing.get(entry.id) || []
    const product = entry.row
    const price = cleanCell(previous[priceIndex]) !== '' ? previous[priceIndex] : (entry.indices.price >= 0 ? toNumber(product[entry.indices.price]) : '')
    const stock = cleanCell(previous[stockIndex]) !== '' ? previous[stockIndex] : (entry.indices.stock >= 0 ? toNumber(product[entry.indices.stock]) : '')
    const minimum = cleanCell(previous[minimumIndex]) !== '' ? previous[minimumIndex] : 5
    const status = cleanCell(previous[statusIndex])
      || (Number(stock) <= 0 ? 'Sin stock' : Number(stock) <= Number(minimum) ? 'Stock bajo' : 'Disponible')
    row[idIndex] = entry.id
    row[nameIndex] = entry.name
    row[categoryIndex] = entry.indices.category >= 0 ? cleanCell(product[entry.indices.category]) : ''
    row[stockIndex] = stock
    row[minimumIndex] = minimum
    row[statusIndex] = status
    row[priceIndex] = price
    row[valueIndex] = cleanCell(previous[valueIndex]) || (Number.isFinite(Number(price)) && Number.isFinite(Number(stock)) ? Number(price) * Number(stock) : '')
    return row
  })
}

function enforceWorkbookCoherence(sheets) {
  const byName = new Map()
  for (const sheet of sheets) {
    const canonical = workbookCanonicalName(sheet?.name)
    if (!canonical) continue
    sheet.name = canonical
    byName.set(canonical, sheet)
  }
  if (!byName.size) return sheets

  const productSheet = byName.get('Productos') || { name: 'Productos', columns: [...PRODUCT_COLUMNS], rows: [] }
  const clientSheet = byName.get('Clientes') || { name: 'Clientes', columns: [...SHEET_COLUMNS.Clientes], rows: [] }
  if (!byName.has('Productos')) sheets.push(productSheet)
  if (!byName.has('Clientes')) sheets.push(clientSheet)

  const inventorySheet = byName.get('Inventario')
  const salesSheet = byName.get('Ventas')
  const advisorySheet = byName.get('Asesorias')
  const inventoryProductIndex = columnIndex(inventorySheet?.columns || [], /producto|nombre/i, 1)
  const salesProductIndex = columnIndex(salesSheet?.columns || [], /producto/i, 3)
  const salesClientIndex = columnIndex(salesSheet?.columns || [], /cliente/i, 2)
  const advisoryClientIndex = columnIndex(advisorySheet?.columns || [], /cliente/i, 1)
  const products = ensureCatalogFromReferences(productSheet, 'PRO', [
    inventorySheet && { sheet: inventorySheet, indices: [inventoryProductIndex] },
    salesSheet && { sheet: salesSheet, indices: [salesProductIndex] }
  ].filter(Boolean))
  const clients = ensureCatalogFromReferences(clientSheet, 'CLI', [
    salesSheet && { sheet: salesSheet, indices: [salesClientIndex] },
    advisorySheet && { sheet: advisorySheet, indices: [advisoryClientIndex] }
  ].filter(Boolean))

  if (byName.has('Plantas')) normalizeSequentialSheet(byName.get('Plantas'), 'PLT')
  if (byName.has('Proveedores')) normalizeSequentialSheet(byName.get('Proveedores'), 'PRV')
  if (inventorySheet) rebuildInventorySheet(inventorySheet, products)

  if (salesSheet) {
    const clientIdIndex = ensureColumn(salesSheet, 'ID Cliente')
    const productIdIndex = ensureColumn(salesSheet, 'ID Producto')
    const dateIndex = columnIndex(salesSheet.columns, /fecha|date/i, 1)
    const clientIndex = columnIndex(salesSheet.columns, /cliente/i, 2)
    const productIndex = columnIndex(salesSheet.columns, /producto/i, 3)
    salesSheet.rows = (Array.isArray(salesSheet.rows) ? salesSheet.rows : []).map((row) => {
      const next = Array.isArray(row) ? [...row] : []
      while (next.length < salesSheet.columns.length) next.push('')
      const client = resolveCatalogEntry(clients, next[clientIndex])
      const product = resolveCatalogEntry(products, next[productIndex])
      if (client) { next[clientIndex] = client.name; next[clientIdIndex] = client.id }
      if (product) { next[productIndex] = product.name; next[productIdIndex] = product.id }
      next[dateIndex] = clampSaleDate(next[dateIndex])
      return next
    })
  }

  if (advisorySheet) {
    const clientIdIndex = ensureColumn(advisorySheet, 'ID Cliente')
    const clientIndex = columnIndex(advisorySheet.columns, /cliente/i, 1)
    advisorySheet.rows = (Array.isArray(advisorySheet.rows) ? advisorySheet.rows : []).map((row) => {
      const next = Array.isArray(row) ? [...row] : []
      while (next.length < advisorySheet.columns.length) next.push('')
      const client = resolveCatalogEntry(clients, next[clientIndex])
      if (client) { next[clientIndex] = client.name; next[clientIdIndex] = client.id }
      return next
    })
  }

  return sheets
}

export function mergeWorkbookSheets(aiSheets, products = [], records = {}, options = {}) {
  const provided = new Map()
  if (Array.isArray(aiSheets)) {
    aiSheets.forEach((raw) => {
      const sheet = normalizeAiSheet(raw)
      if (!sheet || !sheet.name) return
      const canonical = workbookCanonicalName(sheet.name)
      if (canonical) sheet.name = canonical
      provided.set(workbookSheetKey(sheet.name), sheet)
    })
  }
  const canonicalKeys = new Set(SHEET_NAMES.map(workbookSheetKey))
  const extras = [...provided.entries()].filter(([key]) => !canonicalKeys.has(key))
  const hasCanonical = [...provided.keys()].some((key) => canonicalKeys.has(key))
  // Si la IA entregó solo hojas propias (auditoría, KPIs, datos limpios), se
  // respeta ese libro tal cual y no se rellenan las 7 hojas de VITIMAX.
  if (provided.size && !hasCanonical) {
    return enforceWorkbookCoherence(ensureSummarySheet([...provided.values()], products))
  }
  const allowDemo = options.allowDemo ?? !products.length
  const merged = SHEET_NAMES.map((name) => {
    const ai = provided.get(workbookSheetKey(name))
    if (ai && (ai.columns.length || ai.rows.length)) {
      const fallback = defaultSheetRows(name, products, records, { ...options, allowDemo })
      const columns = ai.columns.length ? ai.columns : fallback.columns
      // Si la IA dejó la hoja vacía y el catálogo no tiene datos, se completa con el ejemplo.
      const rows = ai.rows.length ? ai.rows : (allowDemo ? fallback.rows : [])
      return { name, columns, rows }
    }
    const fallback = defaultSheetRows(name, products, records, { ...options, allowDemo })
    return { name, columns: fallback.columns.length ? fallback.columns : SHEET_COLUMNS[name] || [], rows: fallback.rows }
  })
  // Las hojas de resumen/KPIs van primero; el resto de hojas propias al final.
  const summaryExtras = extras.filter(([, sheet]) => /resumen|kpi|dashboard|indicador/i.test(sheet.name))
  const otherExtras = extras.filter(([, sheet]) => !/resumen|kpi|dashboard|indicador/i.test(sheet.name))
  return enforceWorkbookCoherence(ensureSummarySheet([
    ...summaryExtras.map(([, sheet]) => sheet),
    ...merged,
    ...otherExtras.map(([, sheet]) => sheet)
  ], products))
}

function safeSheetName(name, index) {
  const cleaned = String(name || `Hoja ${index + 1}`).replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31)
  return cleaned || `Hoja ${index + 1}`
}

const EXCEL_COLORS = {
  headerFill: 'FF1E4D2B',
  headerText: 'FFFFFFFF',
  zebraFill: 'FFF4F7F4',
  border: 'FFDDE7DA',
  totalsFill: 'FFE6EFE8',
  kpiFill: 'FFEDF5EE',
  alertFill: 'FFF8D7D7',
  alertText: 'FF9B1C1C',
  okFill: 'FFE3F4E4',
  okText: 'FF1E6B34'
}

const NUMERIC_HEADER = /(precio|stock|cantidad|unidades|total|subtotal|iva|margen|valor|monto|ingreso|costo|utilidad|d[ií]as|porcentaje|rotaci[oó]n|tasa|inventario|existencia)/i
const MONEY_HEADER = /(precio|valor|monto|ingreso|costo|utilidad|margen|total|subtotal|iva|inventario)/i
const PERCENT_HEADER = /(porcentaje|%|tasa|iva|margen|rotaci[oó]n)/i
const DATE_HEADER = /(fecha|date|d[ií]a|periodo|mes|año|ano)/i
const ALERT_HEADER = /(alerta|estado|situaci[oó]n|disponibilidad)/i
const STOCK_HEADER = /(stock|inventario|existencia)/i
const MARGIN_HEADER = /(margen|utilidad|rentabilidad|resultado)/i
const ALERT_BAD = ['quiebre', 'crítico', 'critico', 'agotado', 'sin stock', 'bajo', 'reponer', 'negativo', 'pérdida', 'perdida', 'vencido', 'atrasado']
const ALERT_GOOD = ['óptimo', 'optimo', 'ok', 'disponible', 'normal', 'completo', 'suficiente', 'positivo', 'vigente', 'al día', 'al dia', 'saludable']

function excelColumnLetter(index) {
  let value = Number(index) || 1
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters || 'A'
}

// OpenXML guarda las funciones en inglés, aunque la interfaz esté en español.
// Normalizamos cualquier fórmula que llegue con nombres locales o ";" antes de
// entregarla a ExcelJS. ExcelJS recibe la fórmula SIN el "=" inicial.
const EXCEL_FUNCTION_TRANSLATIONS = [
  [/CONTAR\.SI\.CONJUNTO/gi, 'SUMIFS'], [/CONTAR\.SI/gi, 'COUNTIF'],
  [/SUMA\.SI/gi, 'SUMIF'], [/PROMEDIO\.SI/gi, 'AVERAGEIF'],
  [/CONTARA/gi, 'COUNTA'], [/CONTAR/gi, 'COUNT'], [/SUMA/gi, 'SUM'],
  [/PROMEDIO/gi, 'AVERAGE'], [/REDONDEAR/gi, 'ROUND'], [/SI\.ERROR/gi, 'IFERROR'],
  [/(^|[^A-Z])SI(?=[($])/gi, '$1IF']
]

export function translateExcelFormula(formula) {
  let text = String(formula || '').trim().replace(/^=/, '')
  for (const [pattern, replacement] of EXCEL_FUNCTION_TRANSLATIONS) text = text.replace(pattern, replacement)
  return text.replace(/;/g, ',')
}

function excelCellValue(value, header) {
  if (value === null || value === undefined) return null
  if (value instanceof Date) return value
  if (typeof value === 'number' || typeof value === 'boolean') return value
  const text = String(value).replace(/\u00a0/g, ' ').trim()
  if (!text) return null
  // Una cadena que empieza con "=" es una fórmula real, no texto.
  // ExcelJS necesita un "result" para serializar la fórmula; Excel la
  // recalcula al abrir el libro (fullCalcOnLoad). El "=" se elimina aquí:
  // cell.value = { formula: 'SUM(Productos!D2:D11)', result: 0 }.
  if (text.startsWith('=')) return { formula: translateExcelFormula(text), result: 0 }
  if (PERCENT_HEADER.test(header) && /%\s*$/.test(text)) {
    const percent = toNumber(text)
    if (Number.isFinite(percent)) return percent / 100
  }
  if (DATE_HEADER.test(header) && /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(text)) {
    const parsed = new Date(text)
    if (!Number.isNaN(parsed.getTime())) return parsed
  }
  if (NUMERIC_HEADER.test(header) && /^-?[$€]?\s?\d[\d.,]*$/.test(text)) {
    const numeric = toNumber(text)
    if (Number.isFinite(numeric)) return numeric
  }
  return text
}

function excelColumnMeta(columns, rows) {
  const isDateString = (value) => typeof value === 'string' && /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/.test(value.trim())
  // La fila de totales no define el tipo de la columna (su etiqueta puede estar
  // en cualquier celda, no solo en la primera).
  const dataRows = rows.filter((row) => !isTotalsRow(row))
  const sampleRows = dataRows.length ? dataRows : rows
  return columns.map((header, index) => {
    const values = sampleRows.map((row) => row[index]).filter((value) => value !== '' && value !== null && value !== undefined)
    const numeric = values.length > 0 && values.every((value) => {
      if (typeof value === 'number') return true
      if (value instanceof Date) return true
      if (typeof value === 'string' && value.trim().startsWith('=')) return true
      // Una columna "Fecha" con texto ISO también se trata como fecha numérica.
      if (DATE_HEADER.test(header) && isDateString(value)) return true
      return typeof value === 'string' && /^-?[$€]?\s?\d[\d.,]*$/.test(value.trim())
    })
    const decimals = numeric && values.some((value) => {
      const text = String(value).trim()
      return /[.,]\d{1,2}$/.test(text) && !/[.,]\d{3}$/.test(text)
    })
    let numFmt = null
    if (numeric && PERCENT_HEADER.test(header)) numFmt = '0,0%'
    else if (numeric && MONEY_HEADER.test(header)) numFmt = decimals ? '$#,##0.00' : '$#,##0'
    else if (numeric && DATE_HEADER.test(header)) numFmt = 'dd-mm-yyyy'
    else if (numeric) numFmt = decimals ? '#,##0.00' : '#,##0'
    return { numeric, numFmt }
  })
}

function excelWidths(columns, rows) {
  return columns.map((header, index) => {
    const longest = rows.reduce((max, row) => {
      const value = row[index]
      const text = typeof value === 'object' && value?.formula ? `=${value.formula}` : String(value ?? '')
      return Math.max(max, text.length)
    }, 0)
    const width = Math.max(14, Math.min(56, Math.max(String(header || '').length + 2, longest + 2) + 5))
    return { width }
  })
}

function applyExcelConditionalFormatting(worksheet, columns, rows) {
  if (!rows.length) return
  const lastRow = rows.length + 1
  columns.forEach((header, index) => {
    const label = String(header || '')
    const letter = excelColumnLetter(index + 1)
    const ref = `${letter}2:${letter}${lastRow}`
    const values = rows.map((row) => row[index]).filter((value) => value !== '' && value !== null && value !== undefined)
    if (!values.length) return
    const numeric = values.every((value) => typeof value === 'number'
      || (typeof value === 'string' && value.trim().startsWith('='))
      || (typeof value === 'string' && /^-?[$€]?\s?\d[\d.,]*$/.test(value.trim())))
    const alertStyle = {
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: EXCEL_COLORS.alertFill } },
      font: { color: { argb: EXCEL_COLORS.alertText } }
    }
    const okStyle = {
      fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: EXCEL_COLORS.okFill } },
      font: { color: { argb: EXCEL_COLORS.okText } }
    }
    const rules = []
    if (ALERT_HEADER.test(label)) {
      if (numeric) {
        rules.push({ type: 'expression', formulae: [`$${letter}2<=0`], priority: 1, style: alertStyle })
        rules.push({ type: 'expression', formulae: [`$${letter}2>0`], priority: 2, style: okStyle })
      } else {
        let priority = 1
        ALERT_BAD.forEach((text) => rules.push({ type: 'containsText', operator: 'containsText', text, priority: priority++, style: alertStyle }))
        ALERT_GOOD.forEach((text) => rules.push({ type: 'containsText', operator: 'containsText', text, priority: priority++, style: okStyle }))
      }
    } else if (numeric && MARGIN_HEADER.test(label)) {
      rules.push({ type: 'expression', formulae: [`$${letter}2<0`], priority: 1, style: alertStyle })
      rules.push({ type: 'expression', formulae: [`$${letter}2>=0`], priority: 2, style: okStyle })
    } else if (numeric && STOCK_HEADER.test(label)) {
      rules.push({ type: 'expression', formulae: [`$${letter}2<=0`], priority: 1, style: alertStyle })
      rules.push({ type: 'expression', formulae: [`$${letter}2>0`], priority: 2, style: okStyle })
    }
    if (rules.length) worksheet.addConditionalFormatting({ ref, rules })
  })
}

/**
 * Escribe un .xlsx real con ExcelJS: fórmulas activas, formato condicional,
 * encabezados congelados, autofiltro, anchos y estilos profesionales.
 */
export async function createExcelBlob(sheets, options = {}) {
  const ExcelJS = (await import('exceljs')).default
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'VITIMAX IA'
  workbook.created = new Date()
  workbook.modified = new Date()
  // Fuerza a Excel a recalcular todas las fórmulas al abrir el archivo.
  workbook.calcProperties.fullCalcOnLoad = true
  const used = new Set()
  const rawList = Array.isArray(sheets) && sheets.length ? sheets : [{ name: 'Hoja 1', columns: [], rows: [] }]
  // Coherencia referencial final (por si el libro viene de otra ruta) y resumen
  // ejecutivo recalculado desde Productos/Inventario reales.
  const list = reconcileSummarySheets(enforceWorkbookCoherence(rawList))
  list.forEach((sheet, index) => {
    let name = safeSheetName(sheet.name, index)
    let suffix = 1
    while (used.has(name.toLowerCase())) {
      name = `${safeSheetName(sheet.name, index).slice(0, 27)} ${suffix++}`
    }
    used.add(name.toLowerCase())
    const isSummarySheet = /resumen|kpi|dashboard|indicador/i.test(name)
    const columns = (sheet.columns || []).map((column) => cleanCell(column))
    let rows = (sheet.rows || []).map((row) => columns.map((_, i) => cleanCell(row[i])))
    // Sin filas fantasma de totales ni duplicados dentro de las hojas de datos.
    rows = sanitizeSheetRows(rows, { keepTotals: isSummarySheet })
    if (/asesor/i.test(name)) rows = fixAsesoriaRows(columns, rows, options.products || [])
    const worksheet = workbook.addWorksheet(name, {
      views: [{ state: 'frozen', ySplit: 1 }],
      properties: { defaultRowHeight: 18 }
    })
    const meta = excelColumnMeta(columns, rows)
    worksheet.columns = excelWidths(columns, rows)
    const headerRow = worksheet.addRow(columns)
    headerRow.height = 28
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: EXCEL_COLORS.headerText }, size: 11 }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.headerFill } }
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
      cell.border = { bottom: { style: 'thin', color: { argb: EXCEL_COLORS.headerFill } } }
    })
    rows.forEach((row, rowIndex) => {
      const excelRow = worksheet.addRow(row.map((value, columnIndex) => excelCellValue(value, columns[columnIndex] || '')))
      excelRow.height = 20
      const isTotals = /^(gran\s+)?(total|totales|subtotal|suma|resumen)\b/i.test(String(row[0] || '').trim())
      excelRow.eachCell((cell, columnIndex) => {
        cell.alignment = { vertical: 'top', wrapText: true }
        cell.border = { bottom: { style: 'hair', color: { argb: EXCEL_COLORS.border } } }
        if (rowIndex % 2) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.zebraFill } }
        const format = meta[columnIndex - 1]?.numFmt
        if (format && !cell.formula) cell.numFmt = format
        if (cell.formula) cell.numFmt = format || '#,##0'
        // En el resumen, cada métrica usa su formato: CLP para dinero, entero para unidades.
        if (isSummarySheet && columnIndex > 1 && (typeof cell.value === 'number' || cell.formula)) {
          cell.numFmt = /valor|precio|monto|ingreso|costo|total/i.test(String(row[0] || ''))
            ? '$#,##0'
            : '#,##0'
        }
        if (isTotals) {
          cell.font = { bold: true, color: { argb: EXCEL_COLORS.headerFill } }
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.totalsFill } }
          cell.border = {
            ...cell.border,
            top: { style: 'double', color: { argb: EXCEL_COLORS.headerFill } }
          }
        }
      })
    })
    if (columns.length && rows.length) {
      worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: columns.length }
      }
    }
    if (isSummarySheet) {
      worksheet.getColumn(1).font = { bold: true }
      if (columns.length > 1) {
        worksheet.getColumn(2).font = { bold: true, size: 12, color: { argb: EXCEL_COLORS.headerFill.slice(2) } }
        worksheet.getColumn(2).alignment = { horizontal: 'right' }
      }
      for (let rowIndex = 2; rowIndex <= rows.length + 1; rowIndex += 1) {
        worksheet.getCell(rowIndex, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: EXCEL_COLORS.kpiFill } }
      }
    }
    applyExcelConditionalFormatting(worksheet, columns, rows)
  })
  const buffer = await workbook.xlsx.writeBuffer()
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

// ---------------------------------------------------------------------------
// Auditoría de libros adjuntos: todas las hojas, tipos, estadísticas y calidad
// ---------------------------------------------------------------------------

function numericCell(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string') return null
  const text = value.trim()
  if (!text || text.startsWith('=')) return null
  if (!/^-?[$€]?\s?\d[\d.,]*%?$/.test(text)) return null
  const numeric = toNumber(text)
  return Number.isFinite(numeric) ? numeric : null
}

function detectExcelType(values) {
  let numbers = 0
  let dates = 0
  let texts = 0
  let booleans = 0
  values.forEach((value) => {
    if (!hasValue(value)) return
    if (value instanceof Date) { dates += 1; return }
    if (typeof value === 'boolean') { booleans += 1; return }
    if (typeof value === 'number') { numbers += 1; return }
    const text = String(value).trim()
    if (/^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(text) || /^\d{4}-\d{2}-\d{2}/.test(text)) { dates += 1; return }
    if (numericCell(value) !== null) { numbers += 1; return }
    texts += 1
  })
  if (numbers && !dates && !texts && !booleans) return 'número'
  if (dates && !numbers && !texts && !booleans) return 'fecha'
  if (texts && !numbers && !dates && !booleans) return 'texto'
  if (booleans && !numbers && !dates && !texts) return 'booleano'
  const total = numbers + dates + texts + booleans
  if (!total) return 'vacío'
  return 'mixto'
}

function formatStat(value) {
  if (!Number.isFinite(value)) return '0'
  return String(Number(value.toFixed(2)))
}

function columnStatistics(values) {
  const present = values.filter((value) => hasValue(value))
  const numericValues = present.map(numericCell).filter((value) => value !== null)
  const unique = new Set(present.map((value) => value instanceof Date ? value.toISOString().slice(0, 10) : String(value).trim().toLowerCase()))
  const stats = {
    type: detectExcelType(present),
    present: present.length,
    empty: values.length - present.length,
    unique: unique.size
  }
  if (numericValues.length && numericValues.length === present.length) {
    stats.sum = numericValues.reduce((sum, value) => sum + value, 0)
    stats.average = stats.sum / numericValues.length
    stats.min = Math.min(...numericValues)
    stats.max = Math.max(...numericValues)
  } else if (present.length) {
    const counts = new Map()
    present.forEach((value) => {
      const key = value instanceof Date ? value.toISOString().slice(0, 10) : String(value).trim().slice(0, 60)
      counts.set(key, (counts.get(key) || 0) + 1)
    })
    stats.top = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 3)
  }
  return stats
}

function analyzeSheet(matrix) {
  if (!matrix.length) return null
  const sampleLimit = Math.min(matrix.length, 12)
  let headerIndex = 0
  let bestScore = -1
  for (let index = 0; index < sampleLimit; index += 1) {
    const score = matrix[index].filter(hasValue).length
    if (score > bestScore) { bestScore = score; headerIndex = index }
  }
  const rawHeaders = matrix[headerIndex] || []
  const headers = rawHeaders.map((header, index) => cleanCell(header) || `Columna ${index + 1}`)
  const dataRows = matrix.slice(headerIndex + 1).filter((row) => row.some(hasValue))
  const columns = headers.map((header, index) => ({
    header,
    ...columnStatistics(dataRows.map((row) => row[index]))
  }))
  const signatures = new Map()
  dataRows.forEach((row) => {
    const key = headers.map((_, index) => String(row[index] ?? '').trim().toLowerCase()).join('|')
    if (!key.replace(/\|/g, '')) return
    signatures.set(key, (signatures.get(key) || 0) + 1)
  })
  const duplicateRows = [...signatures.values()].reduce((sum, count) => sum + Math.max(0, count - 1), 0)
  return { headers, columns, dataRows, duplicateRows }
}

/**
 * Resume todas las hojas de un libro con encabezados, tipos y estadísticas.
 * El texto resultante se inyecta en el contexto de la IA para que pueda
 * auditar, limpiar, cruzar y transformar los datos reales del archivo.
 */
export function analyzeWorkbook(workbook, options = {}) {
  const fileName = options.fileName || 'libro.xlsx'
  if (!workbook?.SheetNames?.length) return { text: '', sheets: [] }
  const maxSheets = Math.min(workbook.SheetNames.length, Number(options.maxSheets) || 30)
  const perSheetChars = Number(options.perSheetChars) || 5200
  const totalChars = Number(options.totalChars) || 30000
  const sheets = []
  const chunks = []
  workbook.SheetNames.slice(0, maxSheets).forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) return
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false, raw: true })
    const analysis = analyzeSheet(matrix)
    if (!analysis) {
      sheets.push({ name: sheetName, rows: 0, columns: [] })
      chunks.push(`HOJA ${sheetIndex + 1}: "${sheetName}" — vacía.`)
      return
    }
    const { headers, columns, dataRows, duplicateRows } = analysis
    sheets.push({
      name: sheetName,
      rows: dataRows.length,
      columns: columns.map((column) => ({ header: column.header, type: column.type, empty: column.empty, unique: column.unique }))
    })
    const lines = [
      `HOJA ${sheetIndex + 1}: "${sheetName}" — ${dataRows.length} filas de datos · ${headers.length} columnas`,
      `Encabezados: ${headers.join(' | ')}`,
      `Tipos: ${columns.map((column) => `${column.header}=${column.type}`).join(', ')}`,
      'Estadísticas por columna:'
    ]
    columns.forEach((column) => {
      const parts = [`${column.header}: ${column.type}`, `vacíos ${column.empty}`, `únicos ${column.unique}`]
      if (Number.isFinite(column.sum)) {
        parts.push(`suma ${formatStat(column.sum)}`, `promedio ${formatStat(column.average)}`, `mín ${formatStat(column.min)}`, `máx ${formatStat(column.max)}`)
      } else if (column.top?.length) {
        parts.push(`frecuentes ${column.top.map(([value, count]) => `${value} (${count})`).join(', ')}`)
      }
      lines.push(`- ${parts.join(' · ')}`)
    })
    if (duplicateRows) lines.push(`Filas duplicadas detectadas: ${duplicateRows}`)
    if (dataRows.length) {
      lines.push('Muestra de datos:')
      dataRows.slice(0, 8).forEach((row) => {
        lines.push(row.slice(0, headers.length).map((cell) => {
          if (cell instanceof Date) return cell.toISOString().slice(0, 10)
          return String(cell ?? '').replace(/\s+/g, ' ').trim().slice(0, 70)
        }).join(' | '))
      })
    }
    const text = lines.join('\n')
    chunks.push(text.length > perSheetChars ? `${text.slice(0, perSheetChars)}\n[Hoja recortada]` : text)
  })
  if (workbook.SheetNames.length > maxSheets) chunks.push(`[${workbook.SheetNames.length - maxSheets} hojas adicionales no incluidas]`)
  const header = `LIBRO "${fileName}" — ${workbook.SheetNames.length} hoja(s), ${sheets.reduce((sum, sheet) => sum + (sheet.rows || 0), 0)} fila(s) de datos.`
  const text = `${header}\n\n${chunks.join('\n\n')}`.slice(0, totalChars)
  return { text, sheets }
}

export function parseJsonBlock(text) {  const source = String(text || '').trim()
  if (!source) return null
  const candidates = []
  const fenced = source.match(/```(?:json)?\s*([\s\S]*?)```/i)
  if (fenced) candidates.push(fenced[1])
  candidates.push(source)
  for (const candidate of candidates) {
    const start = candidate.indexOf('{')
    const end = candidate.lastIndexOf('}')
    if (start === -1 || end <= start) continue
    try {
      const parsed = JSON.parse(candidate.slice(start, end + 1))
      if (parsed && typeof parsed === 'object') return parsed
    } catch { /* siguiente candidato */ }
  }
  return null
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value))
}

export function simulateSales(products = [], options = {}) {
  const days = clamp(Number(options.periodoDias) || 30, 7, 365)
  const growth = clamp(Number(options.crecimientoMensual) || 0, -50, 200) / 100
  const list = (products.length ? products : buildDemoProducts()).map((product) => normalizeProduct(product))
  const rows = list.map((product, index) => {
    const price = toNumber(product.price)
    const stock = toNumber(product.stock)
    const baseVelocity = 0.12 + ((index * 13) % 35) / 100
    const units = Math.max(0, Math.min(stock || 20, Math.round((stock || 20) * baseVelocity * (days / 30) * (1 + growth))))
    const revenue = units * price
    const margin = Math.round(revenue * 0.4)
    return {
      producto: product.name,
      categoria: product.category,
      precio: price,
      stock,
      unidadesVendidas: units,
      ingresos: revenue,
      margenEstimado: margin,
      rotacion: stock ? Math.round((units / stock) * 100) : 0
    }
  })
  const totalIngresos = rows.reduce((sum, row) => sum + row.ingresos, 0)
  const totalUnidades = rows.reduce((sum, row) => sum + row.unidadesVendidas, 0)
  const top = [...rows].sort((a, b) => b.ingresos - a.ingresos).slice(0, 3)
  const lowRotation = rows.filter((row) => row.rotacion < 20)
  const lowStock = rows.filter((row) => row.stock > 0 && row.stock <= 5)
  const noStock = rows.filter((row) => row.stock <= 0)
  const recommendations = []
  if (top.length) recommendations.push(`Prioriza la venta de ${top.map((row) => row.producto).join(', ')}: concentran la mayor proyección de ingresos.`)
  if (lowStock.length) recommendations.push(`Repón stock de ${lowStock.map((row) => row.producto).join(', ')} para no perder ventas.`)
  if (noStock.length) recommendations.push(`Activa una alerta de reposición para ${noStock.map((row) => row.producto).join(', ')}: hoy no tienen stock.`)
  if (lowRotation.length) recommendations.push(`Prepara una promoción o combina ${lowRotation.map((row) => row.producto).join(', ')} para mejorar su rotación.`)
  recommendations.push('Registra cada venta en el módulo Pedidos para que el BOT pueda comparar la proyección con la realidad.')
  recommendations.push('Revisa precios y margen por categoría cada mes antes de ajustar la estrategia comercial.')
  return {
    periodoDias: days,
    crecimientoMensual: Math.round(growth * 100),
    moneda: 'CLP',
    supuestos: 'Proyección de ejemplo basada en stock, precio y una rotación estimada. No es una promesa de ventas.',
    filas: rows,
    totales: {
      productos: rows.length,
      unidadesVendidas: totalUnidades,
      ingresos: totalIngresos,
      margenEstimado: rows.reduce((sum, row) => sum + row.margenEstimado, 0)
    },
    recomendaciones: recommendations
  }
}

export function salesSheetRows(simulation) {
  return simulation.filas.map((row) => [
    cleanCell(row.producto), cleanCell(row.categoria), row.precio, row.stock, row.unidadesVendidas,
    row.ingresos, row.margenEstimado, `${row.rotacion}%`
  ])
}

export function buildSalesSheets(simulation, products = [], records = {}) {
  const base = mergeWorkbookSheets([], products, records, { allowDemo: !products.length })
  const withSales = base.map((sheet) => sheet.name === 'Ventas'
    ? {
      name: 'Ventas',
      columns: ['Producto', 'Categoría', 'Precio CLP', 'Stock', 'Unidades proyectadas', 'Ingresos CLP', 'Margen estimado CLP', 'Rotación'],
      rows: salesSheetRows(simulation)
    }
    : sheet)
  withSales.push({
    name: 'Recomendaciones',
    columns: ['N°', 'Recomendación comercial'],
    rows: simulation.recomendaciones.map((text, index) => [index + 1, cleanCell(text)])
  })
  return withSales
}

export function summarizeSimulation(simulation) {
  const format = (value) => new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', maximumFractionDigits: 0 }).format(Number(value) || 0)
  const top = [...simulation.filas].sort((a, b) => b.ingresos - a.ingresos).slice(0, 5)
  return [
    `SIMULACIÓN DE VENTAS · ${simulation.periodoDias} días`,
    '',
    `Productos analizados: ${simulation.totales.productos}`,
    `Unidades proyectadas: ${simulation.totales.unidadesVendidas}`,
    `Ingresos proyectados: ${format(simulation.totales.ingresos)}`,
    `Margen estimado: ${format(simulation.totales.margenEstimado)}`,
    '',
    'TOP PRODUCTOS POR INGRESOS',
    ...top.map((row) => `• ${row.producto}: ${row.unidadesVendidas} u. · ${format(row.ingresos)} · rotación ${row.rotacion}%`),
    '',
    'RECOMENDACIONES',
    ...simulation.recomendaciones.map((text) => `• ${text}`),
    '',
    simulation.supuestos
  ].join('\n')
}

// ---------------------------------------------------------------------------
// Prompts del BOT
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Memoria persistente del BOT IA
// ---------------------------------------------------------------------------
export const MEMORY_STORAGE_KEY = 'vitimax_ai_memory'
export const MEMORY_CATEGORIES = ['Negocio', 'Preferencias', 'Instrucciones', 'Decisiones', 'Datos personales', 'Clientes', 'Productos', 'Finanzas', 'Contexto', 'Otros']

function memoryId() {
  return `mem-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

export function normalizeMemoryRecord(raw = {}) {
  const contenido = String(raw.contenido ?? raw.content ?? raw.dato ?? raw.texto ?? '').trim()
  if (!contenido) return null
  return {
    id: String(raw.id || raw.clave || memoryId()).slice(0, 80),
    categoria: String(raw.categoria || raw.category || 'Otros').trim().slice(0, 40) || 'Otros',
    contenido: contenido.slice(0, 1200),
    fecha_actualizacion: raw.fecha_actualizacion || raw.updatedAt || raw.fecha || new Date().toISOString()
  }
}

function memoryTextKey(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
}

// Firmas que identifican un mismo hecho aunque cambie la redacción. Así una
// actualización reemplaza al recuerdo anterior en vez de contradecirlo.
function memorySignatures(record) {
  const text = record.contenido
  const signatures = [`texto:${memoryTextKey(text)}`]
  if (/\b(?:mi nombre|me llamo|nombre del usuario)\b/i.test(text)) signatures.push('identidad')
  if (/\biva\b/i.test(text)) signatures.push('regla-iva')
  if (/\bmoneda\b/i.test(text) && /\b(?:clp|peso|pesos|usd|dolar|dólar|ars)\b/i.test(text)) signatures.push('regla-moneda')
  if (/\b(?:whatsapp|numero de telefono|número de teléfono)\b/i.test(text)) signatures.push('contacto-principal')
  if (/\b(?:razon social|empresa|negocio)\s+(?:se llama|es)\b/i.test(text)) signatures.push('identidad-negocio')
  return signatures
}

// Conserva la versión más reciente de cada recuerdo y descarta duplicados o
// versiones contradictorias del mismo dato.
export function normalizeMemories(list) {
  if (!Array.isArray(list)) return []
  const records = list
    .map((item) => normalizeMemoryRecord(item))
    .filter(Boolean)
    .sort((a, b) => String(b.fecha_actualizacion).localeCompare(String(a.fecha_actualizacion)))
  const seen = new Set()
  const kept = []
  for (const record of records) {
    const signatures = memorySignatures(record)
    if (signatures.some((signature) => seen.has(signature))) continue
    signatures.forEach((signature) => seen.add(signature))
    kept.push(record)
  }
  return kept
}

export function upsertMemoryRecord(list, raw) {
  const record = normalizeMemoryRecord(raw)
  if (!record) return { records: normalizeMemories(list), record: null }
  return { records: normalizeMemories([record, ...(Array.isArray(list) ? list : [])]), record }
}

// Bloque que se inyecta en el system prompt antes de cada consulta a la IA.
export function buildMemorySection(memories = []) {
  const list = normalizeMemories(memories)
  if (!list.length) return ''
  const lines = []
  let total = 0
  // Resumen breve y siempre presente: la memoria viaja en cada llamada, así
  // que se recorta para no consumir el límite de tokens por minuto del modelo.
  for (const item of list) {
    const line = `- [${item.categoria}] ${item.contenido.slice(0, 180)}`
    if (total + line.length > 2600) break
    lines.push(line)
    total += line.length + 1
  }
  if (!lines.length) return ''
  return `

### MEMORIA PERSISTENTE DEL USUARIO Y NEGOCIO
Úsalos de forma natural: NO los repitas, NO los contradigas y NO anuncies que consultaste memoria. Si un dato cambió, reemplaza la versión anterior.
${lines.join('\n')}

Para "recuerda que...", "mi nombre es..." o cambios, usa guardar_memoria / olvidar_memoria; sin herramientas, devuelve un bloque \`\`\`json con {"accion":"guardar_memoria","categoria":"...","dato":"..."}. Nunca guardes claves ni datos de tarjetas.`
}

/**
 * Contexto silencioso de los últimos turnos. Se limita a 10 turnos (mensaje
 * del usuario + respuestas/ herramientas posteriores) y a 14 000 caracteres
 * para no disparar el límite de tokens por minuto de los modelos gratuitos.
 */
export function buildConversationSection(history = []) {
  const rows = limitRecentHistory(history, 10)
  if (!rows.length) return ''
  const lines = []
  let total = 0
  rows.forEach((item) => {
    const speaker = item.role === 'user' ? 'USUARIO' : item.role === 'tool' ? 'HERRAMIENTA' : 'VITIMAX IA'
    const text = String(item.content).replace(/\s+/g, ' ').trim().slice(0, 1200)
    const stamp = item.at || item.timestamp || item.createdAt
    const when = stamp ? getTemporalContext(stamp) : null
    const prefix = when ? `[${when.fechaISO} ${when.hora24Short}] ` : ''
    const line = `${prefix}${speaker}: ${text}`
    if (total + line.length > 14_000) return
    lines.push(line)
    total += line.length + 1
  })
  if (!lines.length) return ''
  const elapsed = formatElapsedSince(lastInteractionAt(rows), getTemporalContext().epoch)
  return `

### HISTORIAL COMPLETO DE LA CONVERSACIÓN
Últimos ${Math.min(10, rows.filter((item) => item.role === 'user').length)} turnos, contextuales y ya compartidos: no los repitas ni digas que los consultaste.
${elapsed ? `${elapsed}\n` : ''}${lines.join('\n')}`
}

// Regla literal que va al inicio y al final de todos los system prompt.
export const SPANISH_ONLY_RULE = 'IDIOMA OBLIGATORIO: Responde SIEMPRE en español natural, claro y profesional (Chile, moneda CLP $#,##0). Está estrictamente prohibido escribir pensamientos internos, bloques <think>, razonamientos en inglés, nombres técnicos de herramientas internas o enlaces sandbox:/ en tus respuestas visibles.'

// Últimos N turnos: se conserva el mensaje del usuario que inicia cada turno y
// todo lo que le sigue (respuesta, llamadas a herramientas y resultados).
export function limitRecentHistory(history = [], maxTurns = 10) {
  const rows = (Array.isArray(history) ? history : [])
    .filter((item) => item && ['user', 'assistant', 'tool'].includes(item.role) && typeof item.content === 'string' && item.content.trim())
  if (rows.length <= maxTurns) return rows
  const userIndexes = []
  rows.forEach((item, index) => { if (item.role === 'user') userIndexes.push(index) })
  if (userIndexes.length <= maxTurns) return rows
  return rows.slice(userIndexes[userIndexes.length - maxTurns])
}

// Preferencia solicitada: principal gemini 2.0 flash y respaldos clásicos.
// Si OpenRouter ya no publica alguno de estos IDs, el 404 se trata como error
// temporal de ruta y la cascada continúa con los modelos gratuitos activos.
export const DEFAULT_FREE_MODELS = [
  'google/gemini-2.0-flash-exp:free',
  'meta-llama/llama-3.3-70b-instruct:free',
  'qwen/qwen-2.5-72b-instruct:free',
  'mistralai/mistral-small-3.1-24b-instruct:free'
]

// Modelos gratuitos activos verificados en OpenRouter, como respaldo final.
export const VERIFIED_FREE_MODELS = [
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

// "openrouter/auto" se conserva como último recurso solicitado, después de
// agotar los modelos gratuitos explícitos.
export const FREE_AI_MODELS = [...new Set([...DEFAULT_FREE_MODELS, ...VERIFIED_FREE_MODELS, 'openrouter/auto'])]

/** Cascada sin duplicados: primero los configurados, luego los gratuitos activos. */
export function aiModelCascade(configured = []) {
  const list = [...(Array.isArray(configured) ? configured : []), ...FREE_AI_MODELS]
  return [...new Set(list.map((item) => String(item || '').trim()).filter(Boolean))]
}

/** Solo se rota ante fallos temporales; 401/402/403 son definitivos. */
export function isRetryableAiError(error) {
  const status = Number(error?.status || error?.statusCode || 0)
  if ([401, 402, 403].includes(status)) return false
  // 404 = modelo retirado de OpenRouter: se salta y se prueba el siguiente.
  if ([404, 408, 429, 502, 503, 504].includes(status)) return true
  if (error?.name === 'AbortError' || error?.name === 'TimeoutError') return true
  const text = String(error?.message || '')
  if (/\b(?:401|402|403)\b/i.test(text)) return false
  return /\b(?:404|408|429|502|503|504)\b|no longer supported|not found|rate.?limit|too many requests|saturad|timeout|tardó|fetch failed|network|aborted|conexi|offline/i.test(text)
}

const formatClpLocal = (value) => `$${Math.round(Number(value) || 0).toLocaleString('es-CL')}`

function localVariantIndex(query, variants, history = []) {
  let hash = history.length * 31
  for (let index = 0; index < query.length; index += 1) hash = (hash * 33 + query.charCodeAt(index)) % 100000
  return hash % variants.length
}

/**
 * Motor local inteligente: responde al mensaje exacto cuando no hay internet.
 * Las estadísticas del catálogo solo aparecen si el usuario las pide.
 */
export function buildLocalAssistantAnswer({ products = [], modules = {}, message = '', history = [] } = {}) {
  const query = String(message || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()
  const pick = (variants) => variants[localVariantIndex(query, variants, history)]

  if (/^(?:hola|buenas|buen dia|que tal|q tal|hey|hi|hello)\b/.test(query)) {
    return pick([
      'Hola, ¿qué resolvemos hoy?',
      'Buenas. Estoy aquí, ¿en qué te ayudo?',
      'Hola. Dime por dónde empezamos.'
    ])
  }
  if (/\b(?:lol|jaja|jajaja|jeje|xd|jks|risa)\b/.test(query)) {
    return pick([
      'Jaja, buena. ¿Seguimos con lo que estabas viendo o cambiamos de tema?',
      'Jeje, me sumo al buen ánimo. ¿Qué sigue?',
      'JAJA. Con una sonrisa se trabaja mejor. ¿En qué te ayudo ahora?'
    ])
  }
  if (/\b(?:gracias|gracias por todo|mil gracias|agradecido)\b/.test(query)) {
    return pick([
      'Con gusto. Quedo atento si necesitas algo más.',
      'De nada. Aquí sigo cuando lo necesites.',
      'Pura vida. ¿Vamos con otra cosa?'
    ])
  }
  if (/\b(?:como estas|como va|que hay|todo bien)\b/.test(query)) {
    return pick([
      'Bien por aquí, listo para ayudarte. ¿Tú cómo estás?',
      'Todo bien. ¿Qué te trae por acá hoy?',
      'Bien, gracias por preguntar. ¿En qué te puedo ayudar?'
    ])
  }
  if (/\b(?:api|endpoint|json|rest|http)\b/.test(query)) {
    return 'Una API es una forma ordenada de que dos programas se hablen: uno hace una solicitud (por ejemplo, GET o POST) y el otro devuelve una respuesta, normalmente en JSON. Piensa en ella como un mesero: toma tu pedido, lo lleva a la cocina y te trae el resultado. Si quieres, te explico con un ejemplo en JavaScript o en Python.'
  }
  if (/\b(?:tomate|tomates|sembrar|siembra|agricultura|agronom|cultivo|planta)\b/.test(query)) {
    return 'Para sembrar tomate en Chile: elige variedad según destino (Limachino o Corazón de Buey para fresco; Roma para industria), asegura pH 6,0–6,8 y buen drenaje, haz almácigo 25–35 días y trasplanta con 4–6 hojas cuando el suelo supere 12–15 °C. Usa 25.000–35.000 plantas/ha [Ejemplo / Demostración], riego por goteo sin mojar el follaje, fertilización fraccionada con N, P y K según análisis de suelo, rotación de 3–4 años sin solanáceas y monitoreo de mosca blanca, tizón y pudrición apical. Si me dices tu zona y si es invernadero o campo, ajusto el calendario.'
  }
  if (/\b(?:codigo|código|programar|programacion|javascript|python|react|bug|funcion|función)\b/.test(query)) {
    return 'Claro. Para avanzar rápido: describe qué debe hacer el código, qué entrada recibe y qué salida esperas; con eso puedo darte una función completa, explicarla por partes y señalar los errores comunes. Si ya tienes código, pégalo y lo revisamos línea por línea.'
  }
  if (/\b(?:estudiar|estudio|examen|aprender|teoria|teoría|resumen de estudio)\b/.test(query)) {
    return 'Para estudiar mejor: divide el tema en bloques pequeños, haz una pregunta por bloque, repite en intervalos (hoy, mañana, en 3 días) y termina explicándolo en tus propias palabras sin mirar apuntes. Si me dices la materia y el tema, te armo un plan de estudio y un resumen claro.'
  }
  if (/\b(?:vitimax|catalogo|catálogo|inventario|productos|stock|reporte|informe|resumen|clientes|pedidos|ventas|proveedores|asesorias|plantas)\b/.test(query)) {
    return buildLocalCatalogSummary(products, modules, query)
  }
  if (/\b(?:negocio|negocios|ventas|cliente|clientes|precio|cotizacion|cotización|marketing|emprender)\b/.test(query)) {
    return 'Para un negocio agrícola, empieza por tres números: costo real por unidad, precio que el cliente acepta y margen resultante. Después separa clientes por frecuencia de compra y enfócate en quienes repiten. Si quieres, armamos juntos una cotización o un plan de seguimiento de clientes.'
  }
  if (/\b(?:consejo|consejos|recomendacion|recomendación|qué hago|que hago|dime algo|humor|chiste)\b/.test(query)) {
    return pick([
      'Mi consejo: decide primero qué es lo importante, y trata lo demás como ruido. Si quieres, lo aplicamos a tu caso concreto.',
      'Un chiste corto: ¿Por qué el tomate nunca discute? Porque siempre quiere llegar a su punto. Si quieres otro, voy por otro.',
      'Consejo práctico: escribe el problema en una frase. Si no puedes escribirlo en una frase, aún no está definido.'
    ])
  }
  return pick([
    'Te sigo. Cuéntame un poco más y te respondo con criterio, sin rodeos.',
    'Buena pregunta. Dame el contexto (para qué lo necesitas y qué ya sabes) y lo resolvemos.',
    'Entendido. ¿Quieres que te lo explique simple, con un ejemplo, o que lo llevemos a un plan concreto?'
  ])
}

function buildLocalCatalogSummary(products = [], modules = {}, query = '') {
  const list = filterJunkProducts(products)
  const units = list.reduce((sum, item) => sum + (Number(item.stock) || 0), 0)
  const value = list.reduce((sum, item) => sum + (Number(item.price) || 0) * (Number(item.stock) || 0), 0)
  const low = list.filter((item) => (Number(item.stock) || 0) <= 5).length
  const parts = []
  if (/iva|impuesto/.test(query)) parts.push('El IVA en Chile es 19%.')
  parts.push(`El catálogo real tiene ${list.length} producto(s), ${units} unidades en stock y un valor de ${formatClpLocal(value)}.`)
  if (low) parts.push(`${low} producto(s) tienen stock bajo (5 o menos).`)
  const moduleKeys = [
    ['clientes', 'clientes registrados'], ['ventas', 'ventas registradas'],
    ['plantas', 'plantas registradas'], ['proveedores', 'proveedores registrados'],
    ['pedidos', 'pedidos registrados'], ['asesorias', 'asesorías registradas']
  ]
  for (const [key, label] of moduleKeys) {
    if (new RegExp(key).test(query) && Array.isArray(modules?.[key])) {
      parts.push(`Hay ${modules[key].length} ${label}.`)
    }
  }
  return parts.join(' ')
}

/** Compatibilidad: antes solo resumía datos; ahora responde al mensaje real. */
export function buildLocalDataSummary(args = {}) {
  return buildLocalAssistantAnswer(args)
}

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

/** Saludos breves y cambios de tema no deben arrastrar pendientes anteriores. */
export function shouldResetConversationContext(message = '', history = []) {
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

/**
 * Reenumera las listas numeradas del chat antes de parsearlas: algunos modelos
 * escriben "1." en cada renglón y el renderizador terminaba mostrando 1, 1, 1.
 * Los números explícitos consecutivos se respetan; los duplicados se corrigen.
 */
export function normalizeChatListNumbers(value) {
  const lines = String(value ?? '').replace(/\r\n?/g, '\n').split('\n')
  let fence = false
  let counter = 0
  return lines.map((line) => {
    if (/^\s*```/.test(line)) { fence = !fence; counter = 0; return line }
    if (fence) return line
    const match = line.match(/^(\s*(?:\*\*|__)?\s*)(\d{1,3})(\s*[.)-]\s*)/)
    if (match) {
      const explicit = Number(match[2])
      counter = counter ? counter + 1 : (explicit > 0 ? explicit : 1)
      return `${match[1]}${counter}${match[3]}${line.slice(match[0].length)}`
    }
    if (line.trim()) counter = 0
    return line
  }).join('\n')
}

function extractAssistantContent(value) {
  if (typeof value === 'string') return value
  if (Array.isArray(value)) {
    return value.map((part) => {
      if (typeof part === 'string') return part
      if (part && typeof part === 'object' && typeof part.text === 'string') return part.text
      return ''
    }).filter(Boolean).join('\n')
  }
  // Nunca se leen campos reasoning / reasoning_content: solo texto visible.
  if (value && typeof value === 'object' && typeof value.text === 'string') return value.text
  return ''
}

const REASONING_START_PATTERNS = [
  /^\s*(?:analysis|reasoning|thinking|thoughts?|plan|notas? internas?|pensamiento(?:s)? internos?)\s*[:：]/i,
  /^\s*(?:(?:actually|okay|ok|so|but|now|well)\s*,?\s*)?(?:looking at (?:the )?(?:instructions|context|catalog|file|request|prompt|data)|the user\b|they want\b|i (?:need to|should|must|will|have to|can|am going to)|i'll|we (?:need to|should|must|will)|let me\b|my plan\b|first,? i\b|next,? i\b|to do this\b|for (?:the|this|excel|pdf|the catalog|the file)|since (?:the|this|only)|based on (?:the )?(?:context|instructions|prompt|request)|according to (?:the )?context|the (?:context|catalog|user|app|file) (?:is|has|shows|says))\b/i
]

const ENGLISH_HINT_RE = /\b(?:the|and|is|are|to|of|for|with|this|that|i|we|my|have|has|need|should|will|can|from|user|context|catalog|file|app|so|now|let|based|according|since|only|also|then|but|if|it|its|as|at|on|in|by|files|both|create|use|check|actually|instructions|they|want|specifically|asked|sheets|products|details|empty|data|because|those|these)\b/gi
const SPANISH_HINT_RE = /\b(?:el|la|los|las|de|del|que|y|en|con|por|para|es|son|al|un|una|se|su|sus|te|tu|necesito|debo|voy|puedo|archivo|archivos|catálogo|catalogo|producto|productos|cliente|clientes|venta|ventas|informe|excel|pdf|word|resumen|total|precio|stock|fecha|generé|generar|creé|crear|lista|datos)\b/gi

function startsWithReasoning(text) {
  return REASONING_START_PATTERNS.some((pattern) => pattern.test(text))
}

function isReasoningLine(line) {
  const trimmed = String(line || '').trim()
  if (!trimmed) return false
  if (startsWithReasoning(trimmed)) return true
  const english = (trimmed.match(ENGLISH_HINT_RE) || []).length
  const spanish = (trimmed.match(SPANISH_HINT_RE) || []).length
  return english >= 4 && english > spanish * 2
}

const CLICHE_START_RE = /^\s*(?:(?:¡\s*)?claro que sí\s*[!,.:]?\s*|aquí tienes\s*[:,-]?\s*|como asistente de (?:inteligencia artificial|ia)\s*[,.]?\s*)/i

// Limpia frases de relleno y enlaces internos del texto que ve el usuario.
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

// Se aplica al texto completo, antes de separar los bloques de código: así un
// <think> que envuelve un JSON tampoco puede llegar al usuario.
function stripReasoningBlocks(text) {
  return String(text || '')
    .replace(/<\s*(think|thinking|reasoning|analysis)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '')
    .replace(/<\s*(?:think|thinking|reasoning|analysis)\b[^>]*>[\s\S]*$/gi, '')
    .replace(/<\s*\/\s*(?:think|thinking|reasoning|analysis)\s*>/gi, '')
    .replace(/<\|\s*(?:thinking|analysis|reasoning)\s*\|>/gi, '')
}

export function sanitizeAssistantText(value) {
  const raw = extractAssistantContent(value)
  if (!raw) return ''
  const source = stripReasoningBlocks(raw)
  if (!source) return ''
  const segments = source.split(/(```[\s\S]*?```)/g)
  const cleaned = segments.map((segment, index) => {
    // Los bloques de código (JSON de acciones incluido) se conservan intactos.
    if (index % 2 === 1) return segment
    let text = segment
      .replace(/<\s*think(?:\s[^>]*)?\s*>[\s\S]*?(?:<\s*\/\s*think\s*>|$)/gi, '')
      .replace(/<\s*\/\s*think\s*>/gi, '')
      .replace(/<\|\s*(?:thinking|analysis|reasoning)\s*\|>/gi, '')
    text = text.split(/(\n{2,})/).map((part) => {
      if (/^\n{2,}$/.test(part)) return part
      return startsWithReasoning(part.trim()) ? '' : part
    }).join('')
    text = text.split('\n').map((line) => (isReasoningLine(line) ? '' : cleanVisibleLine(line))).join('\n')
    return text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n')
  })
  return cleaned.join('').trim()
}

export function assistantTextFromResponse(response = {}) {
  const source = response && typeof response === 'object' ? response : {}
  const message = source.message && typeof source.message === 'object' ? source.message : source
  const delta = source.delta && typeof source.delta === 'object' ? source.delta : null
  const content = message.content !== undefined ? message.content : delta?.content
  return sanitizeAssistantText(content)
}

export function assistantTextFromOpenRouter(payload) {
  return assistantTextFromResponse({ message: payload?.choices?.[0]?.message || {} })
}

export function assistantTextFromStreamChunk(chunk) {
  let payload = chunk
  if (typeof chunk === 'string') {
    const data = chunk.split(/\r?\n/).map((line) => line.trim()).find((line) => line.startsWith('data:'))
    try { payload = JSON.parse(data ? data.slice(5).trim() : chunk) } catch { return '' }
  }
  return assistantTextFromResponse({ delta: payload?.choices?.[0]?.delta || {} })
}

export async function parseOpenRouterStreamResponse(response) {
  if (!response?.body || typeof response.body.getReader !== 'function') {
    const payload = await response?.json?.().catch(() => ({})) || {}
    return { content: assistantTextFromOpenRouter(payload), toolCalls: [], model: payload.model || '' }
  }
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  let rawContent = ''
  let model = ''
  const toolCalls = []
  let stop = false
  while (!stop) {
    const { value, done } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() || ''
    for (const line of lines) {
      const trimmed = line.trim()
      if (!trimmed.startsWith('data:')) continue
      const data = trimmed.slice(5).trim()
      if (data === '[DONE]') { stop = true; break }
      let payload
      try { payload = JSON.parse(data) } catch { continue }
      if (payload?.model) model = payload.model
      const delta = payload?.choices?.[0]?.delta || {}
      rawContent += assistantTextFromResponse({ delta })
      for (const tool of Array.isArray(delta.tool_calls) ? delta.tool_calls : []) {
        const index = Number.isInteger(tool.index) ? tool.index : 0
        const current = toolCalls[index] || { id: '', type: 'function', function: { name: '', arguments: '' } }
        if (tool.id) current.id = tool.id
        if (tool.type) current.type = tool.type
        if (tool.function?.name) current.function.name += tool.function.name
        if (tool.function?.arguments) current.function.arguments += tool.function.arguments
        toolCalls[index] = current
      }
    }
  }
  const calls = toolCalls.filter((call) => call && call.function?.name)
  const content = sanitizeAssistantText(rawContent)
  if (!content && !calls.length) throw new Error('La IA devolvió una respuesta vacía.')
  return { content, toolCalls: calls, model }
}

export const DIRECT_AI_SYSTEM_PROMPT = `${SPANISH_ONLY_RULE}

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

// Versión operativa enviada en cada llamada. Conserva las reglas críticas del
// prompt senior completo (que sigue siendo el contrato documental) pero cabe
// en unos cientos de tokens: evita agotar el límite por minuto de los modelos
// gratuitos y permite que la cascada de respaldo responda más rápido.
export const COMPACT_AI_SYSTEM_PROMPT = SPANISH_ONLY_RULE + `
Asistente general y multipropósito de VITIMAX: conversa libremente de cualquier tema (saludos, humor, código, estudios, negocios, consejos, agricultura). Responde SIEMPRE al mensaje exacto del usuario con naturalidad, cercanía y humor cuando corresponda. NUNCA sueltes estadísticas del catálogo, memorias ni resúmenes de VITIMAX salvo que el usuario pregunte por VITIMAX o pida un reporte/datos. Español chileno, directo y sin rodeos; nunca escribas <think>, razonamientos, nombres de herramientas internos ni enlaces sandbox:/.
Usa el reloj inyectado para plazos y fechas de documentos sin recitarlo; aplica la memoria persistente sin anunciarla. No inventes datos reales: etiqueta los ficticios como [Ejemplo / Demostración], ignora filas de totales, verifica IVA 19%, Precio × Stock y que todas las cifras del turno coincidan.
Consultas agronómicas (por ejemplo "quiero sembrar tomate"): entrega una recomendación completa y accionable con época y zona, variedad, suelo y pH, preparación, siembra o trasplante, densidad, riego, fertilización, control integrado de plagas y enfermedades, monitoreo, cosecha y precauciones; si falta información local, indícalo y usa [Ejemplo / Demostración] solo para valores orientativos.
AGENTE ÉLITE: elige y encadena herramientas (tool_choice "auto") hasta 8 pasos sobre Productos, Clientes, Plantas, Inventario, Ventas, Asesorías y Proveedores; no simules crear archivos. Si no hay herramientas, devuelve un bloque \`\`\`json con {"accion":"crear_excel"|"crear_pdf"|"crear_word"|"crear_ppt"|"crear_pagina_web"|"crear_codigo"|"guardar_memoria"|"olvidar_memoria", ...}. Nunca muestres ese JSON al usuario: confirma en una línea el archivo creado.
Excel: 8 hojas estándar con Resumen Ejecutivo primero, fórmulas OpenXML en inglés sin "=" duplicado, verde #1E4D2B y CLP $#,##0. PDF/Word: texto sanitizado y numeración real. PowerPoint: 16:9 con Próximos Pasos. Web/código: Tailwind completo, vista previa interactiva y código completo y funcional.
No digas "¡Claro que sí!", "Aquí tienes", "Como asistente de IA", "He guardado en mi memoria", "Según mi reloj" ni "Soy VITIMAX IA".`

export const EXCEL_SYSTEM_PROMPT = `${SPANISH_ONLY_RULE}
Eres el generador de Excel de VITIMAX. Creas planillas reales, completas y listas para usar, con fórmulas activas y formato profesional.
Debes llamar a la herramienta crear_excel (o devolver SOLO un objeto JSON si la herramienta no está disponible) con esta forma:
{
  "filename": "reporte-vitimax.xlsx",
  "resumen": "Descripción breve del contenido",
  "productos": [ { "name": "...", "category": "...", "price": 9990, "stock": 20, "format": "...", "description": "...", "usage": "...", "recommendations": "...", "features": "..." } ],
  "kpis": [
    { "nombre": "Ventas totales", "formula": "=SUM(Ventas!G2:G50)", "unidad": "CLP", "interpretacion": "Ingresos del periodo" }
  ],
  "sheets": [
    { "name": "Productos", "columns": ["ID","Nombre","Categoría","Precio CLP","Costo","IVA 19%","Margen bruto","Stock","Valor inventario","Alerta"], "rows": [["PRO-001","VITIMAX VerdeRaíz","Bioestimulantes",9990,6000,"=D2*0.19","=D2-E2-F2",20,"=D2*H2","=IF(H2<=0,\"Quiebre crítico\",IF(H2<=5,\"Stock bajo\",\"Óptimo\"))"]] }
  ]
}
Reglas obligatorias:
- Para un catálogo general incluye las 8 hojas estándar: "Resumen Ejecutivo", Productos, Clientes, Plantas, Inventario, Ventas, Asesorias y Proveedores. Para una auditoría, análisis o reporte, usa solo las hojas que los datos necesiten.
- La primera hoja de un análisis debe ser "Resumen Ejecutivo" con KPIs, gráfico y fórmulas OpenXML reales en inglés sin "=" duplicado (=SUM, =AVERAGE, =COUNTIF, =SUMPRODUCT, =COUNTA, =IF) que referencien las demás hojas, con formato CLP "$#,##0".
- Las fórmulas deben empezar con "=" y usar referencias reales de fila/columna. No escapes las fórmulas ni escribas el resultado calculado a mano cuando deba ser una fórmula.
- Agrega columnas calculadas cuando aporten: IVA 19%, margen bruto, valor de inventario, ingresos, utilidad y alertas de stock.
- En las columnas de alerta usa "Quiebre crítico", "Stock bajo" u "Óptimo" para que la app aplique formato condicional suave en rojo/verde.
- La app congela automáticamente la fila de encabezados y ajusta el ancho de columnas; no necesitas indicarlo.
- Si el catálogo real está vacío y el usuario pide un catálogo general, productos de ejemplo, plantas, tierra, abonos o precios, GENERA datos completos y realistas. No dejes las hojas vacías y no te niegues.
- Si hay catálogo real, úsalo como base; puedes combinarlo con datos nuevos solo si el usuario lo pidió.
- Precios en CLP como números, sin símbolos ni separadores. Stock y cantidades como números.
- No uses "-" como valor: deja "" si falta un dato.
- En la columna Origen o Notas indica "Catálogo VITIMAX" para datos reales y "[Ejemplo / Demostración]" para datos inventados.
- Incluye en "productos" los productos que deban quedar guardados en la app. Si el usuario pide crear el catálogo desde cero, inclúyelos.
- No agregues explicaciones fuera del JSON ni de la llamada a la herramienta.

${SPANISH_ONLY_RULE}`

export const SALES_SYSTEM_PROMPT = `${SPANISH_ONLY_RULE}
Eres el analista comercial de VITIMAX. Recibes una simulación de ventas en JSON y debes entregar recomendaciones concretas en español.
Reglas:
- Explica primero el panorama en 3 a 5 líneas.
- Luego entrega recomendaciones numeradas y accionables: stock, precios, promociones, combos, clientes, canales y seguimiento.
- Distingue claramente que es una proyección de ejemplo y no una promesa de ventas.
- No inventes datos que no estén en la simulación; si falta información, dilo.
- Termina con 3 próximos pasos medibles.

${SPANISH_ONLY_RULE}`

// ---------------------------------------------------------------------------
// Herramientas (function calling de OpenRouter)
// ---------------------------------------------------------------------------

const productSchema = {
  type: 'object',
  properties: {
    name: { type: 'string', description: 'Nombre del producto' },
    category: { type: 'string' },
    price: { type: 'number', description: 'Precio en CLP' },
    previousPrice: { type: 'number' },
    stock: { type: 'number' },
    format: { type: 'string' },
    description: { type: 'string' },
    usage: { type: 'string', description: 'Modo de uso o dosis de ejemplo' },
    recommendations: { type: 'string' },
    features: { type: 'string' },
    image: { type: 'string' }
  },
  required: ['name']
}

export const ASSISTANT_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'crear_excel',
      description: 'Crea un archivo Excel .xlsx real con hojas, columnas, filas y fórmulas activas. Úsalo para catálogo, inventario, ventas, auditoría de datos, KPIs, clientes, plantas, asesorías, proveedores o cualquier planilla. La app aplica encabezados congelados, estilos, autofiltro y formato condicional suave.',
      parameters: {
        type: 'object',
        properties: {
          filename: { type: 'string', description: 'Nombre del archivo .xlsx' },
          resumen: { type: 'string', description: 'Resumen breve del contenido' },
          productos: { type: 'array', items: productSchema, description: 'Productos que deben guardarse en la app. Inclúyelos si el usuario quiere verlos en el catálogo o pidió crear un catálogo.' },
          kpis: {
            type: 'array',
            description: 'Indicadores para la hoja "Resumen Ejecutivo". Usa fórmulas reales que referencien las hojas de datos.',
            items: {
              type: 'object',
              properties: {
                nombre: { type: 'string', description: 'Nombre del KPI' },
                formula: { type: 'string', description: 'Fórmula Excel que empieza con "=", en inglés y con coma, por ejemplo =SUM(Ventas!G2:G50)' },
                valor: { type: 'string', description: 'Valor literal si no aplica una fórmula' },
                unidad: { type: 'string', description: 'CLP, %, unidades, días…' },
                interpretacion: { type: 'string', description: 'Lectura breve del indicador' }
              },
              required: ['nombre']
            }
          },
          sheets: {
            type: 'array',
            description: 'Hojas del libro. Para catálogo general usa las 7 hojas: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores. Productos y Clientes son la fuente de IDs/nombres: Inventario, Ventas y Asesorias deben repetirlos exactamente (Ventas y Asesorias incluyen "ID Cliente"; Ventas incluye "ID Producto"). Plantas usa IDs PLT-001, PLT-002… en la columna A; Proveedores usa PRV-001, PRV-002… Ninguna fecha de Ventas supera la fecha actual. Para auditorías usa las hojas necesarias y agrega "Resumen Ejecutivo" primero. Las celdas con fórmulas deben empezar con "=" y usar funciones inglesas con coma.',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                columns: { type: 'array', items: { type: 'string' } },
                rows: { type: 'array', items: { type: 'array' } }
              },
              required: ['name', 'columns', 'rows']
            }
          }
        },
        required: ['sheets']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'registrar_productos_en_app',
      description: 'Guarda productos en el catálogo real de VITIMAX (estado de la app y Firestore) para que aparezcan en el panel y catálogo sin subir un Excel.',
      parameters: {
        type: 'object',
        properties: {
          productos: { type: 'array', items: productSchema },
          reemplazar: { type: 'boolean', description: 'Si es true reemplaza el catálogo actual; por defecto agrega y actualiza por nombre.' }
        },
        required: ['productos']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'simular_ventas',
      description: 'Calcula una proyección de ventas con los productos del catálogo o los productos indicados y devuelve totales, rotación y recomendaciones.',
      parameters: {
        type: 'object',
        properties: {
          periodoDias: { type: 'number', description: 'Periodo de la simulación en días (por defecto 30)' },
          crecimientoMensual: { type: 'number', description: 'Crecimiento estimado en porcentaje' },
          productos: { type: 'array', items: productSchema, description: 'Productos a simular; si se omite usa el catálogo actual' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_documento',
      description: 'Crea un archivo PDF, Word o PowerPoint real con contenido estructurado. Alternativa genérica a crear_pdf, crear_word y crear_ppt.',
      parameters: {
        type: 'object',
        properties: {
          tipo: { type: 'string', enum: ['pdf', 'word', 'ppt'] },
          titulo: { type: 'string' },
          subtitulo: { type: 'string', description: 'Subtítulo o bajada ejecutiva (opcional)' },
          contenido: { type: 'string', description: 'Contenido completo en texto plano. Usa "# " para títulos, "## " para secciones, "- " para viñetas y --- para separar diapositivas. Nunca uses LaTeX, "$...$", "~", "**" ni "#" dentro de celdas o frases.' },
          filename: { type: 'string' }
        },
        required: ['tipo', 'contenido']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_pdf',
      description: 'Crea un archivo PDF real, con diseño limpio y contenido completo. Úsalo cuando el usuario pida un PDF, informe, propuesta, guía o documento imprimible.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título del documento' },
          subtitulo: { type: 'string', description: 'Subtítulo o bajada ejecutiva (opcional). No incluyas la fecha: la cabecera la agrega la app.' },
          contenido: { type: 'string', description: 'Contenido completo en texto plano. Usa "# " para títulos, "## " para secciones, "- " para viñetas y tablas Markdown con "|". No incluyas "Fecha:" ni la fecha de hoy: la cabecera la agrega la app una sola vez. Sin LaTeX, sin "$...$", sin "~", "**" ni "#" dentro de celdas. Precios como "$15.000" y dosis como "2 a 3 L/ha"; no inventes totales.' },
          filename: { type: 'string', description: 'Nombre del archivo .pdf (opcional)' }
        },
        required: ['contenido']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_word',
      description: 'Crea un archivo Word .docx real, editable y con formato. Úsalo cuando el usuario pida un documento Word, carta, contrato, informe editable o apunte.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título del documento' },
          subtitulo: { type: 'string', description: 'Subtítulo o bajada ejecutiva (opcional)' },
          contenido: { type: 'string', description: 'Contenido completo en texto plano. Usa "# " para títulos, "## " para secciones, "- " para viñetas y tablas Markdown con "|". Sin LaTeX, sin "$...$", sin "~", "**" ni "#" dentro de celdas. Precios como "$15.000" y dosis como "2 a 3 L/ha"; no inventes totales.' },
          filename: { type: 'string', description: 'Nombre del archivo .docx (opcional)' }
        },
        required: ['contenido']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_ppt',
      description: 'Crea una presentación PowerPoint .pptx real con diapositivas. Úsalo cuando el usuario pida una presentación, pitch, clase o resumen visual.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título de la presentación' },
          subtitulo: { type: 'string', description: 'Subtítulo de portada (opcional)' },
          contenido: { type: 'string', description: 'Contenido completo en texto plano. Separa cada diapositiva con una línea ---. La primera línea de cada bloque es el título de la diapositiva y las demás son viñetas (máximo 5 por diapositiva). Sin LaTeX, sin "$...$", sin "~", "**" ni "#" dentro de celdas.' },
          diapositivas: {
            type: 'array',
            description: 'Alternativa a contenido: una entrada por diapositiva con su título y sus viñetas (máximo 5).',
            items: {
              type: 'object',
              properties: {
                titulo: { type: 'string' },
                puntos: { type: 'array', items: { type: 'string' } }
              },
              required: ['titulo', 'puntos']
            }
          },
          filename: { type: 'string', description: 'Nombre del archivo .pptx (opcional)' }
        },
        required: ['titulo']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_pagina_web',
      description: 'Crea una página web HTML5 moderna y responsive con una sola carga de Tailwind CSS por CDN e iconos Lucide. Para catálogos: hero #1E4D2B, acentos lima/dorado, búsqueda en vivo por nombre/descripción, filtros por categoría, tarjetas con precio CLP, badge de stock y botón "Cotizar por WhatsApp" funcional en cada tarjeta, con todos los registros reales y año 2026 en el pie. La app permite verla en vivo antes de descargarla.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título de la página' },
          subtitulo: { type: 'string', description: 'Bajada o descripción breve (opcional)' },
          resumen: { type: 'string', description: 'Resumen breve de lo que hace la página' },
          html: { type: 'string', description: 'HTML COMPLETO, sin truncar y sin placeholders como "<!-- Agregar más productos aquí -->". Incluye <!doctype html>, <head> con una sola carga de Tailwind CSS (https://cdn.tailwindcss.com) y Lucide para iconos, y <body> con hero #1E4D2B, búsqueda en vivo por nombre/descripción, filtros por categoría y una tarjeta por cada producto real con precio CLP, badge de stock y botón "Cotizar por WhatsApp". Pie de página con año 2026. Nada de tablas HTML antiguas.' },
          filename: { type: 'string', description: 'Nombre del archivo .html (opcional)' }
        },
        required: ['html']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_codigo',
      description: 'Entrega código completo y ejecutable en cualquier lenguaje (Python, JavaScript, TypeScript, SQL, PHP, Java, HTML, etc.). El código se muestra en el chat para copiar y se descarga como archivo real con la extensión correcta.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título o descripción del programa' },
          lenguaje: { type: 'string', description: 'Lenguaje del código: python, javascript, typescript, sql, php, java, html, css, bash, etc.' },
          descripcion: { type: 'string', description: 'Qué debe hacer el código (opcional)' },
          codigo: { type: 'string', description: 'Código completo y ejecutable, sin bloques Markdown dentro del texto.' },
          filename: { type: 'string', description: 'Nombre del archivo con extensión (por ejemplo analisis.py)' }
        },
        required: ['codigo']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'crear_html',
      description: 'Alias de crear_pagina_web: crea una página web HTML moderna con Tailwind CSS por CDN, iconos Lucide, búsqueda y filtros cuando sea un catálogo, responsive y lista para abrir en el navegador.',
      parameters: {
        type: 'object',
        properties: {
          titulo: { type: 'string', description: 'Título de la página' },
          subtitulo: { type: 'string', description: 'Bajada o descripción breve (opcional)' },
          html: { type: 'string', description: 'HTML completo. Incluye <!doctype html>, <head> con Tailwind CSS por CDN y Lucide para iconos, y <body> con el contenido. Para catálogos: búsqueda en vivo, filtros por categoría y tarjetas con precio CLP, stock, dosis y WhatsApp.' },
          filename: { type: 'string', description: 'Nombre del archivo .html (opcional)' }
        },
        required: ['html']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'guardar_memoria',
      description: 'Guarda o actualiza un dato importante del usuario o del negocio para futuras sesiones: preferencias, reglas de cálculo, decisiones, contexto, nombre, moneda, clientes, productos e instrucciones. Si el dato cambió, reemplaza la versión anterior sin conservar contradicciones.',
      parameters: {
        type: 'object',
        properties: {
          categoria: { type: 'string', enum: MEMORY_CATEGORIES, description: 'Categoría del recuerdo' },
          dato: { type: 'string', description: 'Dato breve y claro que debe recordarse, por ejemplo "El IVA siempre es 19%" o "El usuario se llama Juan Pérez".' }
        },
        required: ['categoria', 'dato']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'olvidar_memoria',
      description: 'Elimina un recuerdo persistente o lo actualiza cuando el usuario indica que un dato cambió o ya no aplica.',
      parameters: {
        type: 'object',
        properties: {
          clave_o_id: { type: 'string', description: 'ID del recuerdo o texto/clave que permita identificarlo (por ejemplo "IVA" o "mi nombre").' },
          nuevo_dato: { type: 'string', description: 'Si se indica, actualiza el recuerdo con este nuevo contenido en lugar de eliminarlo.' }
        },
        required: ['clave_o_id']
      }
    }
  },
]
