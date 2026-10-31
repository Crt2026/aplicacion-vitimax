const { onCall, HttpsError } = require('firebase-functions/v2/https')
const { defineSecret } = require('firebase-functions/params')
const { GoogleGenerativeAI } = require('@google/generative-ai')
const admin = require('firebase-admin')

admin.initializeApp()
const geminiApiKey = defineSecret('GEMINI_API_KEY')

const systemInstruction = `Eres el asistente oficial de VITIMAX. Ayudas paso a paso a clientes y administradores con catálogo, inventario, Excel, publicaciones y creación de páginas web.
Reglas obligatorias:
- Usa exclusivamente los datos del catálogo que recibes en la consulta. Nunca inventes productos, precios, stock, características, dosis ni resultados.
- Si falta información, dilo claramente como "pendiente" y pide el dato necesario.
- Para fotografías de plantas, ofrece una orientación preliminar, explica la incertidumbre y recomienda revisión de un especialista de VITIMAX antes de aplicar productos.
- No confirmes diagnósticos ni dosis si no aparecen en una ficha técnica oficial.
- Cuando pidan una página web, guía paso a paso: objetivo, público, producto, contenido confirmado, llamada a la acción y estructura final.
- Responde en español, de forma clara, práctica y accionable.`

exports.vitimaxAssistant = onCall(
  { region: 'us-central1', secrets: [geminiApiKey], enforceAppCheck: false },
  async (request) => {
    if (!request.auth) throw new HttpsError('unauthenticated', 'Debes iniciar sesión para usar la IA.')
    const { message, products = [], attachment } = request.data || {}
    if (typeof message !== 'string' || !message.trim()) {
      throw new HttpsError('invalid-argument', 'La consulta está vacía.')
    }
    if (!Array.isArray(products) || products.length > 200) {
      throw new HttpsError('invalid-argument', 'El catálogo enviado no es válido.')
    }

    const model = new GoogleGenerativeAI(geminiApiKey.value()).getGenerativeModel({
      model: 'gemini-1.5-flash',
      systemInstruction
    })
    const catalog = JSON.stringify(products.map(({ id, name, category, price, stock, format, description, usage, recommendations, features }) => ({
      id, name, category, price, stock, format, description, usage, recommendations, features
    })))
    const parts = [{ text: `CATÁLOGO REAL DE VITIMAX:\n${catalog}\n\nCONSULTA DEL USUARIO:\n${message.trim()}` }]
    if (attachment && typeof attachment.data === 'string' && typeof attachment.mimeType === 'string' && attachment.data.length < 8_000_000) {
      parts.push({ text: `ARCHIVO ADJUNTO: ${String(attachment.name || 'archivo').slice(0, 120)}` })
      parts.push({ inlineData: { mimeType: attachment.mimeType, data: attachment.data } })
    }
    const result = await model.generateContent(parts)
    return { text: result.response.text() }
  }
)
