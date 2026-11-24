// Reloj local en vivo para VITIMAX.
// Una sola fuente de verdad para la interfaz, los prompts de IA, los registros
// y los documentos. Se puede pasar una fecha explícita para verificarlo.

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

function capitalize(value) {
  const text = String(value || '')
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : text
}

function weekNumber(date) {
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
  const isoDay = target.getUTCDay() || 7
  target.setUTCDate(target.getUTCDate() + 4 - isoDay)
  const start = new Date(Date.UTC(target.getUTCFullYear(), 0, 1))
  return Math.ceil(((target.getTime() - start.getTime()) / DAY_MS + 1) / 7)
}

function timezoneOffset(date) {
  const minutes = -date.getTimezoneOffset()
  const sign = minutes < 0 ? '-' : '+'
  const absolute = Math.abs(minutes)
  const hours = Math.floor(absolute / 60)
  const rest = absolute % 60
  return `UTC${sign}${hours}${rest ? `:${String(rest).padStart(2, '0')}` : ''}`
}

export function getTemporalContext(now = new Date()) {
  const date = safeDate(now)
  const locale = 'es-CL'
  const year = date.getFullYear()
  const monthIndex = date.getMonth()
  const day = date.getDate()
  const hour = date.getHours()
  const minute = String(date.getMinutes()).padStart(2, '0')
  const second = String(date.getSeconds()).padStart(2, '0')
  const hora24 = `${String(hour).padStart(2, '0')}:${minute}:${second}`
  const hora24Short = `${String(hour).padStart(2, '0')}:${minute}`

  let hora12 = `${hour % 12 || 12}:${minute}:${second} ${hour >= 12 ? 'p. m.' : 'a. m.'}`
  try {
    hora12 = new Intl.DateTimeFormat(locale, { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true }).format(date)
  } catch { /* el formato manual ya es válido */ }

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
  try { zone = Intl.DateTimeFormat().resolvedOptions().timeZone || zone } catch { /* se conserva la zona por defecto */ }

  const franja = hour >= 5 && hour < 12 ? 'Buenos días' : hour >= 12 && hour < 20 ? 'Buenas tardes' : 'Buenas noches'
  const isoWeekday = date.getDay() === 0 ? 7 : date.getDay()
  const fechaCorta = `${day} de ${month} de ${year}`
  const fechaCompleta = `${capitalize(weekday)}, ${fechaCorta}`

  return {
    iso: date.toISOString(),
    epoch: date.getTime(),
    anio: year,
    mes: capitalize(month),
    mesNumero: monthIndex + 1,
    dia: day,
    diaSemana: capitalize(weekday),
    fechaISO: `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
    fechaCorta,
    fechaCompleta,
    hora24,
    hora24Short,
    hora12,
    hora: hour,
    minuto: date.getMinutes(),
    segundo: date.getSeconds(),
    zonaHoraria: zone,
    offset: timezoneOffset(date),
    numeroSemana: weekNumber(date),
    trimestre: Math.floor(monthIndex / 3) + 1,
    diasRestantesSemana: 7 - isoWeekday,
    diasRestantesMes: new Date(year, monthIndex + 1, 0).getDate() - day,
    franja,
    slot: franja === 'Buenos días' ? 'manana' : franja === 'Buenas tardes' ? 'tarde' : 'noche'
  }
}

export function formatSpanishDate(value = new Date()) {
  return getTemporalContext(value).fechaCompleta
}

export function formatElapsedSince(lastAt, now = new Date()) {
  if (!lastAt) return ''
  const last = safeDate(lastAt)
  if (Number.isNaN(last.getTime())) return ''
  const reference = safeDate(now)
  const seconds = Math.max(0, Math.floor((reference.getTime() - last.getTime()) / 1000))
  const when = getTemporalContext(last)
  const stamp = `${when.fechaCompleta} a las ${when.hora24Short} (${when.zonaHoraria})`

  let elapsed
  if (seconds < 60) elapsed = 'menos de un minuto'
  else if (seconds < 3600) {
    const value = Math.floor(seconds / 60)
    elapsed = `${value} ${value === 1 ? 'minuto' : 'minutos'}`
  } else if (seconds < DAY_MS / 1000) {
    const value = Math.floor(seconds / 3600)
    elapsed = `${value} ${value === 1 ? 'hora' : 'horas'}`
  } else if (seconds < 7 * DAY_MS / 1000) {
    const value = Math.floor(seconds / DAY_MS)
    elapsed = `${value} ${value === 1 ? 'día' : 'días'}`
  } else {
    const value = Math.floor(seconds / (7 * DAY_MS / 1000))
    elapsed = `${value} ${value === 1 ? 'semana' : 'semanas'}`
  }
  return `Última interacción registrada el ${stamp}; han pasado ${elapsed}.`
}

export function lastInteractionAt(history = []) {
  const rows = Array.isArray(history) ? history : []
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const item = rows[index]
    const stamp = item && (item.at || item.timestamp || item.createdAt)
    if (stamp) return stamp
  }
  return ''
}

export function buildTemporalSection(temporal = getTemporalContext(), history = []) {
  const clock = temporal && temporal.fechaCompleta ? temporal : getTemporalContext()
  const elapsed = formatElapsedSince(lastInteractionAt(history), new Date(clock.epoch || Date.now()))
  return `

### RELOJ Y CONTEXTO TEMPORAL EN VIVO: Hoy es ${clock.fechaCompleta}, son las ${clock.hora24} (${clock.franja}). Semana ${clock.numeroSemana} del año. Usa esta referencia exacta para saludar acorde a la hora real, calcular plazos ("hoy", "mañana", "esta semana", "próximo lunes", "fin de mes"), filtrar ventas/asesorías recientes y fechar documentos (Excel, PDF, Word, PowerPoint y Web) sin inventar ni desfasar fechas.
Día de la semana: ${clock.diaSemana}. Mes: ${clock.mes}. Trimestre: ${clock.trimestre}. Hora en 12 horas: ${clock.hora12}. Zona horaria del dispositivo: ${clock.zonaHoraria} (${clock.offset}). Faltan ${clock.diasRestantesSemana} días para terminar la semana y ${clock.diasRestantesMes} días para terminar el mes.
${elapsed || 'Es la primera interacción de esta conversación.'}
Si retomas la charla después de horas o de otro día, adapta el saludo y la continuidad a este paso del tiempo sin mencionarlo de forma robótica. Usa la franja horaria exacta "${clock.franja}" cuando saludes; en respuestas normales empieza por la respuesta, no por un saludo.`
}

export function buildWelcomeMessage(name = '', temporal = getTemporalContext()) {
  const clock = temporal && temporal.fechaCompleta ? temporal : getTemporalContext()
  const who = String(name || '').trim() || 'administrador'
  return `${clock.franja}, ${who}. Hoy es ${clock.fechaCompleta}, son las ${clock.hora24Short} (${clock.zonaHoraria}). Soy VITIMAX IA: puedo ayudarte con código, estudios, cálculos, redacción, negocios y toda la gestión de VITIMAX. También genero Excel, PDF, Word, PowerPoint y páginas web reales. ¿Qué resolvemos?`
}

export function isWelcomeMessage(item) {
  if (!item || typeof item !== 'object') return false
  if (item.welcome === true) return true
  const text = String(item.text || item.content || '')
  return /Soy VITIMAX IA:|Reinicié la conversación para dejar el chat limpio/.test(text)
}
