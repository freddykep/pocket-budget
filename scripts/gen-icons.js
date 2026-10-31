// Erzeugt die PNG-Icons aus dem Logo (Muenze mit Plus). Einmalig ausfuehren: node scripts/gen-icons.js
import { mkdir, writeFile } from 'node:fs/promises';
import sharp from 'sharp';

const TEAL = '#0F766E';
const art = `
  <circle cx="232" cy="288" r="140" fill="#F8FAFC"/>
  <circle cx="232" cy="288" r="118" fill="none" stroke="${TEAL}" stroke-width="8" opacity="0.35"/>
  <path d="M286 250A65 65 0 1 0 286 326" fill="none" stroke="${TEAL}" stroke-width="22" stroke-linecap="round"/>
  <path d="M170 273H262M170 303H262" stroke="${TEAL}" stroke-width="16" stroke-linecap="round"/>
  <circle cx="392" cy="128" r="72" fill="#F59E0B"/>
  <path d="M392 92v72M356 128h72" stroke="#fff" stroke-width="22" stroke-linecap="round"/>`;

const svg = (rounded) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><rect width="512" height="512"${rounded ? ' rx="112"' : ''} fill="${TEAL}"/>${art}</svg>`;
// Maskable: Motiv in der sicheren Zone (80 %), Hintergrund bis zum Rand.
const maskable = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><rect width="512" height="512" fill="${TEAL}"/><g transform="translate(51.2 51.2) scale(0.8)">${art}</g></svg>`;

await mkdir('public/icons', { recursive: true });
await writeFile('public/icons/logo.svg', svg(true));
const png = (s, size, file) => sharp(Buffer.from(s)).resize(size, size).png().toFile(`public/icons/${file}`);
await png(svg(true), 192, 'icon-192.png');
await png(svg(true), 512, 'icon-512.png');
await png(maskable, 512, 'icon-maskable-512.png');
await png(svg(false), 180, 'apple-touch-icon.png'); // iOS rundet selbst ab
console.log('Icons erzeugt');
