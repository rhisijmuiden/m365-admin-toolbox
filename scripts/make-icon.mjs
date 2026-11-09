// Renders build/icon.png (512x512) from the RHIS antenna mark. Run with `npm run icon`.
import sharp from 'sharp'

const s = 0.33
const tx = 256 - 2048.5 * s
const ty = 256 - 560 * s
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512" viewBox="0 0 512 512">
  <rect width="512" height="512" rx="104" fill="#14213D"/>
  <g transform="translate(${tx} ${ty}) scale(${s})">
    <rect x="1941" y="392" width="215" height="688" rx="40" fill="#FF6B57"/>
    <path d="M1942.4 215.9A150 150 0 0 1 2154.6 215.9 M1850.5 124.0A280 280 0 0 1 2246.5 124.0" fill="none" stroke="#FF6B57" stroke-width="75" stroke-linecap="round"/>
  </g>
</svg>`

await sharp(Buffer.from(svg)).png().toFile('build/icon.png')
console.log('build/icon.png written')
