const fs = require('fs')
const path = require('path')
const sharp = require('sharp')

const root = process.cwd()
const res = path.join(root, 'resources')
const iconSvg = fs.readFileSync(path.join(res, 'icon.svg'))
const markSvg = fs.readFileSync(path.join(res, 'icon-mark.svg'), 'utf8')
const botSvg = fs.readFileSync(path.join(res, 'bot-vitimax.svg'))

const colors = {
  dayStart: '#0A3322',
  dayMid: '#123D2B',
  dayEnd: '#1E6B4A',
  nightStart: '#06130D',
  nightMid: '#0A2418',
  nightEnd: '#103A27',
  lime: '#D8F55F',
  soft: '#C9F87E',
  white: '#F7FFF2',
}

async function writePng(file, buffer) {
  fs.mkdirSync(path.dirname(file), { recursive: true })
  await sharp(buffer).png({ compressionLevel: 9 }).toFile(file)
}

async function writeSplash(file, buffer) {
  fs.mkdirSync(path.dirname(file), { recursive: true })
  await sharp(buffer).webp({ quality: 90, effort: 6 }).toFile(file)
}

function roundIconSvg(size) {
  const grad = `<defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="${colors.dayStart}"/>
      <stop offset=".5" stop-color="${colors.dayMid}"/>
      <stop offset="1" stop-color="${colors.dayEnd}"/>
    </linearGradient>
    <clipPath id="clip"><circle cx="${size / 2}" cy="${size / 2}" r="${size / 2}"/></clipPath>
  </defs>`
  const mark = markSvg
    .replace(/^[\s\S]*?<svg[^>]*>/, '')
    .replace(/<\/svg>\s*$/, '')
  const scale = 1.12 * size / 1024
  const offsetX = size / 2 - 512 * scale
  const offsetY = size / 2 - 548 * scale
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
    ${grad}
    <g clip-path="url(#clip)">
      <rect width="${size}" height="${size}" fill="url(#bg)"/>
      <circle cx="${size / 2}" cy="${size / 2}" r="${size * 0.318}" fill="none" stroke="${colors.lime}" stroke-opacity=".14" stroke-width="${Math.max(1, size * 0.008)}"/>
      <g transform="translate(${offsetX} ${offsetY}) scale(${scale})">${mark}</g>
    </g>
  </svg>`
}

function foregroundSvg(size) {
  const mark = markSvg
    .replace(/^[\s\S]*?<svg[^>]*>/, '')
    .replace(/<\/svg>\s*$/, '')
  const scale = 0.79 * size / 1024
  const offsetX = size / 2 - 512 * scale
  const offsetY = size / 2 - 548 * scale
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
    <g transform="translate(${offsetX} ${offsetY}) scale(${scale})">${mark}</g>
  </svg>`
}

function backgroundSvg(size) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stop-color="${colors.dayStart}"/>
        <stop offset=".5" stop-color="${colors.dayMid}"/>
        <stop offset="1" stop-color="${colors.dayEnd}"/>
      </linearGradient>
      <radialGradient id="glow" cx="0" cy="0" r="1" gradientTransform="translate(${size * 0.3} ${size * 0.2}) rotate(48) scale(${size * 0.62})" gradientUnits="userSpaceOnUse">
        <stop offset="0" stop-color="${colors.lime}" stop-opacity=".24"/>
        <stop offset="1" stop-color="${colors.lime}" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${size}" height="${size}" fill="url(#bg)"/>
    <rect width="${size}" height="${size}" fill="url(#glow)"/>
  </svg>`
}

function splashSvg(width, height, night) {
  const min = Math.min(width, height)
  const landscape = width > height
  const mark = min * (landscape ? 0.33 : 0.3)
  const scale = mark / 1024
  const cx = width / 2
  const markCy = height / 2 - (landscape ? 0 : min * 0.055)
  const offsetX = -512 * scale
  const offsetY = -548 * scale
  const title = Math.max(15, min * 0.082)
  const tag = Math.max(7.5, min * 0.029)
  const titleY = markCy + mark / 2 + title * 1.05
  const tagY = titleY + tag * 2.1
  const start = night ? colors.nightStart : colors.dayStart
  const mid = night ? colors.nightMid : colors.dayMid
  const end = night ? colors.nightEnd : colors.dayEnd
  const markInner = markSvg
    .replace(/^[\s\S]*?<svg[^>]*>/, '')
    .replace(/<\/svg>\s*$/, '')
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <defs>
      <linearGradient id="bg" x1="0" y1="0" x2="${width}" y2="${height}" gradientUnits="userSpaceOnUse">
        <stop offset="0" stop-color="${start}"/>
        <stop offset=".55" stop-color="${mid}"/>
        <stop offset="1" stop-color="${end}"/>
      </linearGradient>
      <radialGradient id="glow" cx="0" cy="0" r="1" gradientTransform="translate(${width * 0.5} ${markCy - mark * 0.15}) rotate(90) scale(${min * 0.62})" gradientUnits="userSpaceOnUse">
        <stop offset="0" stop-color="${colors.lime}" stop-opacity="${night ? '.16' : '.22'}"/>
        <stop offset="1" stop-color="${colors.lime}" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${width}" height="${height}" fill="url(#bg)"/>
    <rect width="${width}" height="${height}" fill="url(#glow)"/>
    <circle cx="${cx}" cy="${markCy}" r="${mark * 0.52}" fill="none" stroke="${colors.lime}" stroke-opacity=".13" stroke-width="${Math.max(1, min * 0.006)}"/>
    <g transform="translate(${cx + offsetX} ${markCy + offsetY}) scale(${scale})">${markInner}</g>
    <text x="${cx}" y="${titleY}" text-anchor="middle" fill="${colors.white}" font-family="Manrope, 'Segoe UI', Arial, sans-serif" font-size="${title}" font-weight="800" letter-spacing="${title * 0.16}">VITIMAX</text>
    <text x="${cx}" y="${tagY}" text-anchor="middle" fill="${colors.soft}" font-family="Manrope, 'Segoe UI', Arial, sans-serif" font-size="${tag}" font-weight="700" letter-spacing="${tag * 0.26}" opacity=".92">CENTRO DE GESTIÓN</text>
  </svg>`
}

const androidRes = path.join(root, 'android', 'app', 'src', 'main', 'res')
const densities = {
  ldpi: 0.75,
  mdpi: 1,
  hdpi: 1.5,
  xhdpi: 2,
  xxhdpi: 3,
  xxxhdpi: 4,
}

const splashTargets = [
  ['drawable/splash.png', 320, 480],
  ['drawable-land-ldpi/splash.png', 320, 240],
  ['drawable-land-mdpi/splash.png', 480, 320],
  ['drawable-land-hdpi/splash.png', 800, 480],
  ['drawable-land-xhdpi/splash.png', 1280, 720],
  ['drawable-land-xxhdpi/splash.png', 1600, 960],
  ['drawable-land-xxxhdpi/splash.png', 1920, 1280],
  ['drawable-port-ldpi/splash.png', 240, 320],
  ['drawable-port-mdpi/splash.png', 320, 480],
  ['drawable-port-hdpi/splash.png', 480, 800],
  ['drawable-port-xhdpi/splash.png', 720, 1280],
  ['drawable-port-xxhdpi/splash.png', 960, 1600],
  ['drawable-port-xxxhdpi/splash.png', 1280, 1920],
  ['drawable-night/splash.png', 320, 480],
  ['drawable-land-night-ldpi/splash.png', 320, 240],
  ['drawable-land-night-mdpi/splash.png', 480, 320],
  ['drawable-land-night-hdpi/splash.png', 800, 480],
  ['drawable-land-night-xhdpi/splash.png', 1280, 720],
  ['drawable-land-night-xxhdpi/splash.png', 1600, 960],
  ['drawable-land-night-xxxhdpi/splash.png', 1920, 1280],
  ['drawable-port-night-ldpi/splash.png', 240, 320],
  ['drawable-port-night-mdpi/splash.png', 320, 480],
  ['drawable-port-night-hdpi/splash.png', 480, 800],
  ['drawable-port-night-xhdpi/splash.png', 720, 1280],
  ['drawable-port-night-xxhdpi/splash.png', 960, 1600],
  ['drawable-port-night-xxxhdpi/splash.png', 1280, 1920],
]

async function main() {
  for (const [name, factor] of Object.entries(densities)) {
    const dir = path.join(androidRes, `mipmap-${name}`)
    const legacy = Math.round(48 * factor)
    const adaptive = Math.round(108 * factor)
    await writePng(path.join(dir, 'ic_launcher.png'), await sharp(iconSvg).resize(legacy, legacy).png().toBuffer())
    await writePng(path.join(dir, 'ic_launcher_round.png'), await sharp(Buffer.from(roundIconSvg(legacy))).png().toBuffer())
    await writePng(path.join(dir, 'ic_launcher_foreground.png'), await sharp(Buffer.from(foregroundSvg(adaptive))).png().toBuffer())
    await writePng(path.join(dir, 'ic_launcher_background.png'), await sharp(Buffer.from(backgroundSvg(adaptive))).png().toBuffer())
  }

  for (const [rel, width, height] of splashTargets) {
    const night = rel.includes('night')
    const png = path.join(androidRes, rel)
    if (fs.existsSync(png)) fs.unlinkSync(png)
    await writeSplash(png.replace(/\.png$/, '.webp'), Buffer.from(splashSvg(width, height, night)))
  }

  const iosIcon = path.join(root, 'ios', 'App', 'App', 'Assets.xcassets', 'AppIcon.appiconset', 'AppIcon-512@2x.png')
  await writePng(iosIcon, await sharp(iconSvg).resize(1024, 1024).flatten({ background: colors.dayMid }).removeAlpha().png().toBuffer())

  const iosSplash = path.join(root, 'ios', 'App', 'App', 'Assets.xcassets', 'Splash.imageset')
  for (const name of ['splash-2732x2732.png', 'splash-2732x2732-1.png', 'splash-2732x2732-2.png']) {
    await writePng(path.join(iosSplash, name), await sharp(Buffer.from(splashSvg(2732, 2732, false))).png().toBuffer())
  }

  const pub = path.join(root, 'public')
  fs.copyFileSync(path.join(res, 'icon.svg'), path.join(pub, 'favicon.svg'))
  await writePng(path.join(pub, 'favicon-16.png'), await sharp(iconSvg).resize(16, 16).png().toBuffer())
  await writePng(path.join(pub, 'favicon-32.png'), await sharp(iconSvg).resize(32, 32).png().toBuffer())
  await writePng(path.join(pub, 'icon-192.png'), await sharp(iconSvg).resize(192, 192).png().toBuffer())
  await writePng(path.join(pub, 'icon-512.png'), await sharp(iconSvg).resize(512, 512).png().toBuffer())
  await writePng(path.join(pub, 'icon-512-maskable.png'), await sharp(Buffer.from(backgroundSvg(512))).composite([{ input: await sharp(Buffer.from(foregroundSvg(512))).png().toBuffer() }]).png().toBuffer())
  await writePng(path.join(pub, 'apple-touch-icon.png'), await sharp(iconSvg).resize(180, 180).flatten({ background: colors.dayMid }).removeAlpha().png().toBuffer())

  await writePng(path.join(pub, 'bot-vitimax.png'), await sharp(botSvg).resize(512, 512).png().toBuffer())
  await writePng(path.join(pub, 'bot-vitimax-192.png'), await sharp(botSvg).resize(192, 192).png().toBuffer())

  console.log('Iconos generados')
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
