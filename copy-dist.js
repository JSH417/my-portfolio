const fs = require('fs');
const path = require('path');

const dist = path.join(__dirname, 'dist');
if (!fs.existsSync(dist)) fs.mkdirSync(dist, { recursive: true });

const files = [
  'index.html',
  'manifest.json',
  'sw.js',
  'apple-touch-icon.png',
  'favicon.png',
  'icon-192.png',
  'icon-512.png'
];

files.forEach(f => {
  const src = path.join(__dirname, f);
  if (fs.existsSync(src)) {
    fs.copyFileSync(src, path.join(dist, f));
    console.log(`Copied ${f} to dist/`);
  }
});
console.log('All web assets synced to dist/ successfully!');

