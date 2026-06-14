const sharp = require('sharp');
const { default: pngToIco } = require('png-to-ico');
const fs = require('fs');
const path = require('path');

async function main() {
  const svgPath = path.join(__dirname, 'public', 'favicon.svg');
  const pngPath = path.join(__dirname, 'public', 'icon-256.png');
  const icoPath = path.join(__dirname, 'public', 'favicon.ico');

  // SVG → 256×256 PNG
  await sharp(svgPath)
    .resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png()
    .toFile(pngPath);
  console.log('Created PNG:', pngPath);

  // PNG → ICO
  const ico = await pngToIco(pngPath);
  fs.writeFileSync(icoPath, ico);
  console.log('Created ICO:', icoPath);
}

main().catch(e => { console.error(e); process.exit(1); });
