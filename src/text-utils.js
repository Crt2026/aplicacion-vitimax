// ---------------------------------------------------------------------------
// Utilidades de texto compartidas por PDF, Word, PowerPoint, Excel y web.
// Limpian LaTeX, escapes, Markdown y formato de números CLP.
// ---------------------------------------------------------------------------

// Convierte una expresión LaTeX sencilla a texto legible.
function latexToText(value) {
  return String(value ?? '')
    .replace(/\\(?:times|cdot)\b/gi, '·')
    .replace(/\\div\b/gi, '÷')
    .replace(/\\pm\b/gi, '±')
    .replace(/\\approx\b/gi, '≈')
    .replace(/\\(?:leq|le)\b/gi, '≤')
    .replace(/\\(?:geq|ge)\b/gi, '≥')
    .replace(/\\(?:rightarrow|to)\b/gi, '→')
    .replace(/\\(?:text|mathrm|mathbf|mathit|textbf|textit|operatorname|left|right|quad|qquad)\b/gi, ' ')
    .replace(/\\([%&_#])/g, '$1')
    .replace(/\\\$/g, '$')
    .replace(/\\\{/g, '{')
    .replace(/\\\}/g, '}')
    .replace(/\\[a-zA-Z]+\*?/g, ' ')
    .replace(/[{}]/g, '')
    .replace(/~/g, ' ')
    .replace(/\^/g, '')
    .replace(/_/g, ' ')
    .replace(/[ \t]{2,}/g, ' ')
    .trim()
}

/**
 * Limpia una cadena para mostrarla en documentos:
 * - `$2~a~3~L/ha$` -> `2 a 3 L/ha`
 * - `\$0` -> `$0`
 * - `**texto**`, `# título`, backticks y `~` sueltos -> texto limpio
 * - Conserva los precios sueltos como `$15.000`.
 */
export function sanitizeText(value) {
  let text = String(value ?? '')
  if (!text) return ''
  text = text.replace(/\u00a0/g, ' ').replace(/\r\n?/g, '\n')

  // Protege el dólar escapado (\$) antes de interpretar matemáticas.
  text = text.replace(/\\\$/g, '\u0000VITIMAX_DOLLAR\u0000')

  // Matemática en bloque: $$...$$
  text = text.replace(/\$\$([\s\S]*?)\$\$/g, (_, inner) => latexToText(inner))

  // Matemática en línea: $...$ solo si contiene marcas LaTeX reales.
  // Así no se destruyen precios como "$15.000 y $25.000".
  text = text.replace(/\$([^$\n]+)\$/g, (full, inner) => (
    /[~^_\\{}]/.test(inner) ? latexToText(inner) : full
  ))

  text = text.replace(/\u0000VITIMAX_DOLLAR\u0000/g, '$')

  // Escapes LaTeX sueltos y comandos que hayan quedado fuera de $...$.
  text = text.replace(/\\([%&_#])/g, '$1')
  text = text.replace(/\\([{}])/g, '$1')
  text = text
    .replace(/\\(?:times|cdot)\b/gi, '·')
    .replace(/\\div\b/gi, '÷')
    .replace(/\\pm\b/gi, '±')
    .replace(/\\approx\b/gi, '≈')
    .replace(/\\(?:leq|le)\b/gi, '≤')
    .replace(/\\(?:geq|ge)\b/gi, '≥')
    .replace(/\\(?:rightarrow|to)\b/gi, '→')
    .replace(/\\(?:text|mathrm|mathbf|mathit|textbf|textit|operatorname|left|right|quad|qquad)\b/gi, ' ')
    .replace(/\\[a-zA-Z]+\*?/g, ' ')
    .replace(/[{}]/g, '')

  // Markdown: negritas, itálicas, tachado, código y tildes de LaTeX.
  text = text
    .replace(/`([^`]*)`/g, '$1')
    .replace(/\*\*|__|~~/g, '')
    .replace(/(^|\s)[*_](\S[^*_]*)[*_](?=\s|$)/g, '$1$2')
    .replace(/[*_]/g, '')
    .replace(/~/g, ' ')

  // Encabezados y numerales de Markdown sueltos.
  text = text
    .replace(/(^|\s)#{1,6}\s+/g, '$1')
    .replace(/#/g, '')

  return text.replace(/[ \t]{2,}/g, ' ').replace(/\s+$/gm, '').trim()
}

/**
 * Convierte un valor monetario o numérico escrito por la IA a número.
 * Acepta "$15.000", "15.000", "15,5", "(1.200)" y números reales.
 */
export function parseNumberLoose(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0
  if (value === null || value === undefined) return 0
  let text = String(value).trim()
  if (!text) return 0
  const negative = /^\(.*\)$/.test(text) || /^-/.test(text)
  text = text.replace(/[^\d.,]/g, '')
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
  if (!Number.isFinite(number)) return 0
  return negative ? -Math.abs(number) : number
}

/** Formatea un valor como moneda chilena: $2.750.000 */
export function formatClp(value, { decimals = 0 } = {}) {
  const number = parseNumberLoose(value)
  const fixed = decimals ? Number(number.toFixed(decimals)) : Math.round(number)
  try {
    return `$${fixed.toLocaleString('es-CL', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    })}`
  } catch {
    return `$${String(fixed).replace(/\B(?=(\d{3})+(?!\d))/g, '.')}`
  }
}

/** Suma real de productos y valor de inventario (Precio × Stock). */
export function computeCatalogTotals(catalog = []) {
  const list = Array.isArray(catalog) ? catalog.filter(Boolean) : []
  let units = 0
  let inventoryValue = 0
  let priceSum = 0
  let priced = 0
  let lowStock = 0
  list.forEach((item) => {
    const price = parseNumberLoose(item?.price)
    const stock = parseNumberLoose(item?.stock)
    if (price > 0) {
      priceSum += price
      priced += 1
    }
    if (Number.isFinite(stock)) {
      units += stock
      if (stock <= 5) lowStock += 1
      if (price > 0) inventoryValue += price * stock
    }
  })
  return {
    count: list.length,
    units,
    inventoryValue,
    priceSum,
    avgPrice: priced ? priceSum / priced : 0,
    lowStock
  }
}
