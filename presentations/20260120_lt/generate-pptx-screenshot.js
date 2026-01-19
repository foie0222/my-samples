const { chromium } = require('playwright');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

const SLIDES_DIR = path.join(__dirname, 'slides');
const OUTPUT_DIR = path.join(__dirname, 'output');
const SCREENSHOTS_DIR = path.join(OUTPUT_DIR, 'screenshots');

// スライドファイル一覧
const SLIDE_FILES = [
  'slide01.html',
  'slide02.html',
  'slide03.html',
  'slide04.html',
  'slide05.html',
  'slide06.html',
  'slide07.html',
  'slide08.html',
  'slide09.html',
  'slide10.html',
];

async function main() {
  // スクリーンショット保存ディレクトリを作成
  if (!fs.existsSync(SCREENSHOTS_DIR)) {
    fs.mkdirSync(SCREENSHOTS_DIR, { recursive: true });
  }

  console.log('Starting browser...');
  const browser = await chromium.launch();

  // HTMLのサイズ: 720pt × 405pt
  // ブラウザでは 1pt = 1.333px なので 960px × 540px
  const slideWidth = 960;
  const slideHeight = 540;

  const context = await browser.newContext({
    viewport: { width: slideWidth, height: slideHeight },
    deviceScaleFactor: 2, // 高解像度 (出力は1920×1080px)
  });
  const page = await context.newPage();

  const screenshotPaths = [];

  // 各スライドのスクリーンショットを撮影
  for (let i = 0; i < SLIDE_FILES.length; i++) {
    const slideFile = SLIDE_FILES[i];
    const slidePath = path.join(SLIDES_DIR, slideFile);
    const screenshotPath = path.join(SCREENSHOTS_DIR, `slide${String(i + 1).padStart(2, '0')}.png`);

    console.log(`Capturing ${slideFile}...`);

    await page.goto(`file://${slidePath}`);

    // オーバーフローを防ぐ
    await page.evaluate(() => {
      document.body.style.overflow = 'hidden';
    });

    // bodyのサイズに合わせてスクリーンショット
    await page.screenshot({
      path: screenshotPath,
      clip: { x: 0, y: 0, width: slideWidth, height: slideHeight }
    });

    screenshotPaths.push(screenshotPath);
  }

  await browser.close();
  console.log('Browser closed.');

  // PPTXを作成
  console.log('Creating PPTX...');
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = '井上 大貴';
  pptx.title = '競馬AIの歴史';

  for (const screenshotPath of screenshotPaths) {
    const slide = pptx.addSlide();
    slide.addImage({
      path: screenshotPath,
      x: 0,
      y: 0,
      w: 10,      // スライド幅 (16:9 = 10" × 5.625")
      h: 5.625,   // スライド高さ
    });
  }

  const outputPath = path.join(OUTPUT_DIR, '04_presentation_screenshot.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log(`PPTX generated: ${outputPath}`);
}

main().catch(console.error);
