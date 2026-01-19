const PptxGenJS = require('pptxgenjs');
const html2pptx = require('./html2pptx.cjs');
const path = require('path');
const fs = require('fs');

async function main() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = '井上 大貴';
  pptx.title = '競馬AIの歴史 〜なぜ今から参入しても勝てないのか〜';
  pptx.subject = 'ALT 2026 Winter';
  pptx.company = 'Product Div. AT Dept.';

  const slideFiles = [
    'slide01.html',
    'slide02.html',
    'slide03.html',
    'slide04.html',
    'slide05.html',
    'slide06.html',
    'slide07.html',
    'slide08.html',
    'slide09.html',
  ];

  try {
    for (const slideFile of slideFiles) {
      const slidePath = path.resolve(__dirname, '..', 'slides', slideFile);
      console.log(`Processing: ${slideFile}`);

      if (!fs.existsSync(slidePath)) {
        console.error(`File not found: ${slidePath}`);
        continue;
      }

      const { slide, placeholders } = await html2pptx(slidePath, pptx);
      console.log(`  -> Created (${placeholders.length} placeholders)`);
    }

    const outputFile = path.resolve(__dirname, '..', 'output', '03_presentation.pptx');
    await pptx.writeFile({ fileName: outputFile });
    console.log(`\nCreated: ${outputFile}`);

    const stats = fs.statSync(outputFile);
    console.log(`File size: ${Math.round(stats.size / 1024)} KB`);
  } catch (err) {
    console.error('Error:', err.message);
    if (err.stack) console.error(err.stack);
    process.exit(1);
  }
}

main();
