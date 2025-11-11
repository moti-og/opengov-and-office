// Quick script to create placeholder icon files for the add-in
// Run: node create-icons.js

const fs = require('fs');
const path = require('path');

const sizes = [16, 32, 64, 80];
const color = '0078d4'; // OpenGov blue

// Create simple SVG icons
function createSVGIcon(size) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}">
  <rect width="${size}" height="${size}" fill="#${color}"/>
  <text x="50%" y="50%" font-family="Arial, sans-serif" font-size="${size * 0.6}" 
        fill="white" text-anchor="middle" dominant-baseline="central" font-weight="bold">OG</text>
</svg>`;
}

// Convert SVG to data URL that can be used as PNG
function createDataURL(size) {
  const svg = createSVGIcon(size);
  const base64 = Buffer.from(svg).toString('base64');
  return `data:image/svg+xml;base64,${base64}`;
}

// Alternatively, just save as SVG (Office supports SVG)
const assetsDir = path.join(__dirname, 'web', 'assets');

// Ensure directory exists
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

sizes.forEach(size => {
  const svg = createSVGIcon(size);
  const filename = `icon-${size}.svg`;
  const filepath = path.join(assetsDir, filename);
  fs.writeFileSync(filepath, svg);
  console.log(`✓ Created ${filename}`);
});

// Also create PNG-named copies (just SVG with .png extension - Office will handle it)
sizes.forEach(size => {
  const svg = createSVGIcon(size);
  const filename = `icon-${size}.png`;
  const filepath = path.join(assetsDir, filename);
  fs.writeFileSync(filepath, svg);
  console.log(`✓ Created ${filename} (SVG as PNG)`);
});

console.log('\n✅ All icons created successfully!');
console.log('Icons are SVG format but work with Office add-ins.');
console.log('\nTo use real PNG files, use an image editor or:');
console.log('  - https://favicon.io/emoji-favicons/ (use 📊 emoji)');
console.log('  - ImageMagick: convert -size 32x32 xc:#0078d4 icon-32.png');

