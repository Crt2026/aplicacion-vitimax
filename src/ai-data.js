import * as XLSX from 'xlsx'

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
  Ventas: ['ID', 'Fecha', 'Cliente', 'Producto', 'Cantidad', 'Precio unitario CLP', 'Total CLP', 'Estado', 'Canal'],
  Asesorias: ['ID', 'Cliente', 'Cultivo', 'Consulta', 'Estado', 'Prioridad', 'Recomendación'],
  Proveedores: ['ID', 'Proveedor', 'Contacto', 'Teléfono', 'Correo', 'Ciudad', 'Productos', 'Condiciones']
}

const HEADER_FILL = '173524'
const HEADER_TEXT = 'FFFFFF'
const SUBTLE_FILL = 'EAF2E7'

export function makeId(prefix = 'vitimax') {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID()
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
}

export function cleanCell(value) {
  if (value === null || value === undefined) return ''
  if (typeof value === 'number' || typeof value === 'boolean') return value
  let text = String(value).replace(/\u00a0/g, ' ').trim()
  // El usuario pidió quitar los "-" que aparecían como dato vacío.
  if (/^[-–—]+$/.test(text)) return ''
  if (/^(n\/?a|sin dato|no aplica)$/i.test(text)) return ''
  if (/^=/.test(text)) text = `'${text}`
  return text
}

export function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0
  if (value === null || value === undefined) return 0
  let text = String(value).trim()
  if (!text || /^[-–—]+$/.test(text)) return 0
  text = text.replace(/[^\d.,-]/g, '')
  if (!text) return 0
  const lastComma = text.lastIndexOf(',')
  const lastDot = text.lastIndexOf('.')
  if (lastComma > -1 && lastDot > -1) {
    if (lastComma > lastDot) text = text.replace(/\./g, '').replace(',', '.')
    else text = text.replace(/,/g, '')
  } else if (lastComma > -1) {
    const decimals = text.length - lastComma - 1
    text = decimals === 3 ? text.replace(/,/g, '') : text.replace(',', '.')
  } else if (lastDot > -1) {
    const decimals = text.length - lastDot - 1
    if (decimals === 3) text = text.replace(/\./g, '')
  }
  const number = Number(text)
  return Number.isFinite(number) ? number : 0
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
    .filter((item) => item.name || item.description || item.price)
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
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
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
  if (!record[primary]) {
    const fallback = record.nombre || record.email || record.telefono || record.producto || record.consulta
    if (fallback) record[primary] = fallback
  }
  if (!record[primary] && !Object.values(record).some(hasValue)) return null
  if (!record[primary]) record[primary] = 'Sin nombre'
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
    return { columns: PRODUCT_COLUMNS, rows: safeProducts.map(productToRow) }
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
          fecha: `2026-0${(index % 9) + 1}-1${index % 9}`,
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
        toNumber(item.total), cleanCell(item.estado || 'Pendiente'), cleanCell(item.canal || 'Directo')
      ])
    }
  }
  if (name === 'Asesorias') {
    const advisories = Array.isArray(records.asesorias) && records.asesorias.length
      ? records.asesorias
      : demo
        ? safeProducts.slice(0, 3).map((product, index) => ({
          id: makeId('asesoria'),
          cliente: `Cliente de ejemplo ${index + 1}`,
          cultivo: product.category || 'Cultivo general',
          consulta: `Ejemplo: plan de uso para ${product.name || 'producto VITIMAX'}.`,
          estado: index % 2 ? 'En revisión' : 'Abierta',
          prioridad: index === 0 ? 'Alta' : 'Normal',
          recomendacion: product.recommendations || 'Validar dosis con un especialista antes de aplicar.'
        }))
        : []
    return {
      columns: SHEET_COLUMNS.Asesorias,
      rows: recordRows(advisories, SHEET_COLUMNS.Asesorias, (item) => [
        cleanCell(item.id), cleanCell(item.cliente), cleanCell(item.cultivo), cleanCell(item.consulta),
        cleanCell(item.estado), cleanCell(item.prioridad), cleanCell(item.recomendacion)
      ])
    }
  }
  if (name === 'Proveedores') {
    return { columns: SHEET_COLUMNS.Proveedores, rows: demo ? supplierRows() : [] }
  }
  return { columns: [], rows: [] }
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

export function mergeWorkbookSheets(aiSheets, products = [], records = {}, options = {}) {
  const provided = new Map()
  if (Array.isArray(aiSheets)) {
    aiSheets.forEach((raw) => {
      const sheet = normalizeAiSheet(raw)
      if (!sheet || !sheet.name) return
      provided.set(sheet.name.toLowerCase().replace(/\s+/g, ''), sheet)
    })
  }
  const allowDemo = options.allowDemo ?? !products.length
  const merged = SHEET_NAMES.map((name) => {
    const ai = provided.get(name.toLowerCase().replace(/\s+/g, ''))
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
  if (options.keepExtra) {
    const canonical = new Set(SHEET_NAMES.map((name) => name.toLowerCase().replace(/\s+/g, '')))
    provided.forEach((sheet, key) => {
      if (!canonical.has(key)) merged.push({ name: sheet.name, columns: sheet.columns, rows: sheet.rows })
    })
  }
  return merged
}

function safeSheetName(name, index) {
  const cleaned = String(name || `Hoja ${index + 1}`).replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31)
  return cleaned || `Hoja ${index + 1}`
}

function columnWidths(columns, rows) {
  return columns.map((column, index) => {
    const longest = rows.reduce((max, row) => Math.max(max, String(row[index] ?? '').length), 0)
    const width = Math.max(String(column || '').length + 2, Math.min(52, longest + 2), 12)
    return { wch: width }
  })
}

export function buildWorkbook(sheets) {
  const workbook = XLSX.utils.book_new()
  const used = new Set()
  sheets.forEach((sheet, index) => {
    let name = safeSheetName(sheet.name, index)
    let suffix = 1
    while (used.has(name.toLowerCase())) {
      name = `${safeSheetName(sheet.name, index).slice(0, 27)} ${suffix++}`
    }
    used.add(name.toLowerCase())
    const columns = (sheet.columns || []).map((column) => cleanCell(column))
    const rows = (sheet.rows || []).map((row) => columns.map((_, i) => cleanCell(row[i])))
    const data = [columns, ...rows]
    const worksheet = XLSX.utils.aoa_to_sheet(data)
    worksheet['!cols'] = columnWidths(columns, rows)
    worksheet['!rows'] = [{ hpt: 26 }]
    if (columns.length && rows.length) {
      worksheet['!autofilter'] = {
        ref: XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: columns.length - 1, r: rows.length } })
      }
    }
    // Encabezado con color corporativo (SheetJS lo conserva en Excel compatible).
    columns.forEach((_, columnIndex) => {
      const cell = worksheet[XLSX.utils.encode_cell({ c: columnIndex, r: 0 })]
      if (!cell) return
      cell.s = {
        font: { bold: true, color: { rgb: HEADER_TEXT }, sz: 11 },
        fill: { patternType: 'solid', fgColor: { rgb: HEADER_FILL } },
        alignment: { vertical: 'center', horizontal: 'left', wrapText: true }
      }
    })
    rows.forEach((_, rowIndex) => {
      columns.forEach((__, columnIndex) => {
        const cell = worksheet[XLSX.utils.encode_cell({ c: columnIndex, r: rowIndex + 1 })]
        if (!cell) return
        cell.s = {
          alignment: { vertical: 'top', wrapText: true },
          fill: rowIndex % 2 ? { patternType: 'solid', fgColor: { rgb: 'F7FAF5' } } : { patternType: 'solid', fgColor: { rgb: SUBTLE_FILL } }
        }
      })
    })
    XLSX.utils.book_append_sheet(workbook, worksheet, name)
  })
  return workbook
}

export function workbookBlob(workbook) {
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', cellStyles: true })
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
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

export const DIRECT_AI_SYSTEM_PROMPT = `Eres el BOT IA oficial de VITIMAX, una plataforma para administrar catálogo, inventario, pedidos, clientes, asesorías y soluciones agrícolas.
Responde siempre en español, de forma clara, cercana y práctica. Puedes ayudar con CUALQUIER tema o tarea: catálogo, inventario, ventas, Excel, PDF, Word, PowerPoint, páginas web, marketing, análisis de archivos, agricultura, negocios, cálculos, redacción, código y consultas generales.

REGLA PRINCIPAL SOBRE LOS DATOS:
- Si el catálogo tiene productos, úsalos como fuente principal y cítalos con sus datos reales.
- Si el catálogo está VACÍO y el usuario pide crear un catálogo, productos de ejemplo, tierra, abonos, fertilizantes, plantas, precios, stock, dosis o cualquier dato concreto, GENERA datos completos, coherentes y realistas (ID, categoría, precio en CLP, stock, formato, descripción, dosis y recomendaciones). NO te niegues, NO respondas solo que falta el catálogo y NO entregues un archivo vacío.
- Si el usuario pide inventar, simular, ejemplos o datos DEMO, hazlo y dilo claramente: "Datos generados por IA (ejemplo)".
- No presentes datos generados como si fueran información oficial de VITIMAX. Separa siempre "Datos reales del catálogo" de "Datos generados por IA".
- Puedes combinar el catálogo real con datos nuevos solo cuando el usuario lo pida explícitamente.

HERRAMIENTAS DISPONIBLES:
- crear_excel: crea un archivo .xlsx real con hojas, columnas y filas. Úsala cuando pidan Excel, catálogo, inventario, ventas, clientes, plantas, asesorías, proveedores o cualquier planilla. Cuando pidan un catálogo general o de plantas/productos, llena SIEMPRE las 7 hojas: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias y Proveedores.
- registrar_productos_en_app: guarda en el catálogo real de VITIMAX los productos generados para que aparezcan en el panel y en el catálogo de la app. Úsala cuando el usuario quiera que los productos queden guardados o cuando pida crear un catálogo desde cero.
- simular_ventas: calcula una proyección de ventas y devuelve recomendaciones comerciales.
- crear_documento: crea un PDF, Word o PowerPoint real con el contenido estructurado que indiques.

REGLAS DE TRABAJO:
- Si el usuario pide un archivo, usa la herramienta correspondiente y confirma al final qué archivo creaste y cuántas hojas o productos incluye.
- Si generas productos y el usuario quiere usarlos en la app, inclúyelos en "productos" de crear_excel y/o llama a registrar_productos_en_app.
- No uses "-" como dato. Si falta un dato, deja el campo vacío o escribe "Pendiente".
- Para fotografías de plantas o cultivos, ofrece orientación preliminar, explica la incertidumbre y recomienda revisión de un especialista antes de aplicar productos.
- No confirmes diagnósticos fitosanitarios definitivos ni dosis oficiales si no están en una ficha técnica; si son ejemplos, indícalo.
- Usa listas y pasos cuando ayuden. Ten en cuenta el historial de la conversación.
- Si no puedes usar herramientas, responde con un bloque \`\`\`json con la acción solicitada (por ejemplo {"accion":"crear_excel","sheets":[...]}) para que la app lo ejecute.`

export const EXCEL_SYSTEM_PROMPT = `Eres el generador de Excel de VITIMAX. Creas planillas reales, completas y listas para usar.
Debes llamar a la herramienta crear_excel (o devolver SOLO un objeto JSON si la herramienta no está disponible) con esta forma:
{
  "filename": "catalogo-vitimax.xlsx",
  "resumen": "Descripción breve del contenido",
  "productos": [ { "name": "...", "category": "...", "price": 9990, "stock": 20, "format": "...", "description": "...", "usage": "...", "recommendations": "...", "features": "..." } ],
  "sheets": [
    { "name": "Productos", "columns": ["ID","Nombre","Categoría","Precio CLP","Stock","Formato","Descripción","Modo de uso / dosis","Recomendaciones","Características","Origen"], "rows": [["...","...",9990,20,"...","...","...","...","...","...","..."]] }
  ]
}
Reglas obligatorias:
- Incluye SIEMPRE las 7 hojas con estos nombres exactos: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores.
- Si el catálogo real está vacío y el usuario pide un catálogo general, productos de ejemplo, plantas, tierra, abonos o precios, GENERA datos completos y realistas. No dejes las hojas vacías y no te niegues.
- Si hay catálogo real, úsalo como base; puedes combinarlo con datos nuevos solo si el usuario lo pidió.
- Precios en CLP como números, sin símbolos ni separadores. Stock y cantidades como números.
- No uses "-" como valor: deja "" si falta un dato.
- En la columna Origen o Notas indica "Catálogo VITIMAX" para datos reales y "Generado por IA (ejemplo)" para datos inventados.
- Incluye en "productos" los productos que deban quedar guardados en la app. Si el usuario pide crear el catálogo desde cero, inclúyelos.
- No agregues explicaciones fuera del JSON ni de la llamada a la herramienta.`

export const SALES_SYSTEM_PROMPT = `Eres el analista comercial de VITIMAX. Recibes una simulación de ventas en JSON y debes entregar recomendaciones concretas en español.
Reglas:
- Explica primero el panorama en 3 a 5 líneas.
- Luego entrega recomendaciones numeradas y accionables: stock, precios, promociones, combos, clientes, canales y seguimiento.
- Distingue claramente que es una proyección de ejemplo y no una promesa de ventas.
- No inventes datos que no estén en la simulación; si falta información, dilo.
- Termina con 3 próximos pasos medibles.`

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
      description: 'Crea un archivo Excel .xlsx real con hojas, columnas y filas generadas. Úsalo para catálogo, inventario, ventas, clientes, plantas, asesorías, proveedores o cualquier planilla.',
      parameters: {
        type: 'object',
        properties: {
          filename: { type: 'string', description: 'Nombre del archivo .xlsx' },
          resumen: { type: 'string', description: 'Resumen breve del contenido' },
          productos: { type: 'array', items: productSchema, description: 'Productos que deben guardarse en la app. Inclúyelos si el usuario quiere verlos en el catálogo o pidió crear un catálogo.' },
          sheets: {
            type: 'array',
            description: 'Hojas del libro. Para catálogo general usa las 7 hojas: Productos, Clientes, Plantas, Inventario, Ventas, Asesorias, Proveedores.',
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
      description: 'Crea un archivo PDF, Word o PowerPoint real con contenido estructurado.',
      parameters: {
        type: 'object',
        properties: {
          tipo: { type: 'string', enum: ['pdf', 'word', 'ppt'] },
          titulo: { type: 'string' },
          contenido: { type: 'string', description: 'Contenido completo. Usa "# " para títulos, "## " para secciones, "- " para viñetas y --- para separar diapositivas.' },
          filename: { type: 'string' }
        },
        required: ['tipo', 'contenido']
      }
    }
  }
]
