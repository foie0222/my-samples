const sharp = require('sharp');
const path = require('path');

// Slide dimensions: 720pt x 405pt at 96 DPI = 960px x 540px
const WIDTH = 960;
const HEIGHT = 540;

// Create radial gradient background
async function generateBackground() {
  // Create SVG with radial gradient
  const svg = `
    <svg width="${WIDTH}" height="${HEIGHT}" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <radialGradient id="glow" cx="50%" cy="50%" r="70%" fx="50%" fy="50%">
          <stop offset="0%" style="stop-color:rgb(78,205,196);stop-opacity:0.08" />
          <stop offset="100%" style="stop-color:rgb(13,17,23);stop-opacity:0" />
        </radialGradient>
      </defs>
      <rect width="100%" height="100%" fill="#0d1117"/>
      <rect width="100%" height="100%" fill="url(#glow)"/>
    </svg>
  `;

  const outputPath = path.join(__dirname, '..', 'slides', 'bg-dark.png');

  await sharp(Buffer.from(svg))
    .png()
    .toFile(outputPath);

  console.log(`Created: ${outputPath}`);
}

generateBackground().catch(console.error);
