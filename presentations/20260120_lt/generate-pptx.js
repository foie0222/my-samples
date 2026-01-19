const PptxGenJS = require('pptxgenjs');

// カラーテーマ
const COLORS = {
  bg: '0d1117',
  primary: 'FFFFFF',
  secondary: '8b949e',
  muted: '6b7280',
  accent: '4ECDC4',
  danger: 'DC143C',
  warning: 'FFD700',
  boxBg: '1a1f26',
  boxBorder: '30363d'
};

const FONT = 'Arial';
const IMG_PATH = '/home/inoue-d/dev/my-samples/presentations/20260120_lt/slides/images';
const BG_IMAGE = `${IMG_PATH}/bg-dark.png`;

// pt to inch変換 (72pt = 1inch)
const pt = (val) => val / 72;

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';
pptx.author = '井上 大貴';
pptx.title = '競馬AIの歴史';

// ===== Slide 01: タイトル =====
const slide01 = pptx.addSlide();
slide01.background = { path: BG_IMAGE };
// 中央配置
slide01.addText('ALT 2026 Winter', { x: 0, y: 2.0, w: 10, h: 0.3, fontSize: 11, color: COLORS.accent, fontFace: FONT, align: 'center', charSpacing: 4 });
slide01.addText('競馬AIの歴史', { x: 0, y: 2.35, w: 10, h: 0.8, fontSize: 48, bold: true, color: COLORS.primary, fontFace: FONT, align: 'center' });
slide01.addText('なぜ私は競馬AIをやめたのか', { x: 0, y: 3.1, w: 10, h: 0.4, fontSize: 20, color: COLORS.secondary, fontFace: FONT, align: 'center' });
// フッター (bottom: 24pt = 0.33inch from bottom)
slide01.addText('Presented by', { x: pt(40), y: 4.85, w: 2, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT });
slide01.addText('井上 大貴', { x: pt(40), y: 5.05, w: 2, h: 0.2, fontSize: 10, color: COLORS.primary, fontFace: FONT });
slide01.addText('2026.01.20', { x: 7.5, y: 4.85, w: 2, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'right' });
slide01.addText('ALT', { x: 7.5, y: 5.05, w: 2, h: 0.2, fontSize: 10, color: COLORS.primary, fontFace: FONT, align: 'right' });

// ===== Slide 02: 前回のあらすじ =====
const slide02 = pptx.addSlide();
slide02.background = { path: BG_IMAGE };
const s2x = pt(40), s2y = pt(36);
slide02.addText('PREVIOUS TALK', { x: s2x, y: s2y, w: 9, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide02.addText('前回のあらすじ', { x: s2x, y: s2y + pt(22), w: 9, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });

// タイムライン (中央配置)
const tlY = s2y + pt(70);
const tlItemW = 1.6, tlItemH = 0.85;
const tlGap = 0.15;
const tlStartX = (10 - (tlItemW * 3 + tlGap * 4)) / 2;

// 5万円
slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: tlStartX, y: tlY, w: tlItemW, h: tlItemH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide02.addText('5万円', { x: tlStartX, y: tlY + 0.15, w: tlItemW, h: 0.35, fontSize: 20, bold: true, color: COLORS.primary, fontFace: FONT, align: 'center' });
slide02.addText('軍資金', { x: tlStartX, y: tlY + 0.5, w: tlItemW, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

slide02.addText('→', { x: tlStartX + tlItemW, y: tlY + 0.25, w: tlGap * 2 + 0.3, h: 0.35, fontSize: 20, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// +150万円
const tl2X = tlStartX + tlItemW + tlGap * 2 + 0.3;
slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: tl2X, y: tlY, w: tlItemW, h: tlItemH, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide02.addText('+150万円', { x: tl2X, y: tlY + 0.15, w: tlItemW, h: 0.35, fontSize: 20, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide02.addText('最高到達点', { x: tl2X, y: tlY + 0.5, w: tlItemW, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

slide02.addText('→', { x: tl2X + tlItemW, y: tlY + 0.25, w: tlGap * 2 + 0.3, h: 0.35, fontSize: 20, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 全損
const tl3X = tl2X + tlItemW + tlGap * 2 + 0.3;
slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: tl3X, y: tlY, w: tlItemW, h: tlItemH, fill: { color: '2d1a1f' }, line: { color: COLORS.danger, pt: 1 } });
slide02.addText('全損', { x: tl3X, y: tlY + 0.15, w: tlItemW, h: 0.35, fontSize: 20, bold: true, color: COLORS.danger, fontFace: FONT, align: 'center' });
slide02.addText('結末', { x: tl3X, y: tlY + 0.5, w: tlItemW, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 敗因ラベル
slide02.addText('敗因は2つあった', { x: 0, y: tlY + tlItemH + 0.15, w: 10, h: 0.25, fontSize: 12, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 敗因ボックス
const causeY = tlY + tlItemH + 0.5;
const causeW = 4.2, causeH = 0.85;
const causeGap = pt(16);

slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s2x, y: causeY, w: causeW, h: causeH, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide02.addText('① 期待値が取れなくなった（技術的要因）', { x: s2x + 0.15, y: causeY + 0.12, w: causeW - 0.3, h: 0.25, fontSize: 11, bold: true, color: COLORS.accent, fontFace: FONT });
slide02.addText('オッズの変動、競合の増加、資金力の壁', { x: s2x + 0.15, y: causeY + 0.42, w: causeW - 0.3, h: 0.25, fontSize: 10, color: COLORS.secondary, fontFace: FONT });

slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s2x + causeW + causeGap, y: causeY, w: causeW, h: causeH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide02.addText('② 人間の愚かさ（人的要因）', { x: s2x + causeW + causeGap + 0.15, y: causeY + 0.12, w: causeW - 0.3, h: 0.25, fontSize: 11, bold: true, color: COLORS.secondary, fontFace: FONT });
slide02.addText('裁量介入の失敗、感情によるブレ', { x: s2x + causeW + causeGap + 0.15, y: causeY + 0.42, w: causeW - 0.3, h: 0.25, fontSize: 10, color: COLORS.muted, fontFace: FONT });

// メッセージボックス
const msgY = causeY + causeH + 0.2;
slide02.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: msgY, w: 7, h: 0.55, fill: { color: '1a2f2d' }, line: { color: '3d5c59', pt: 1 } });
slide02.addText('今日は①の構造を深掘りします', { x: 1.5, y: msgY, w: 7, h: 0.55, fontSize: 14, color: COLORS.primary, fontFace: FONT, align: 'center', valign: 'middle' });

// ===== Slide 03: 期待値の構造 =====
const slide03 = pptx.addSlide();
slide03.background = { path: BG_IMAGE };
const s3x = pt(40), s3y = pt(36);
slide03.addText('KEY CONCEPT', { x: s3x, y: s3y, w: 9, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide03.addText('期待値の構造', { x: s3x, y: s3y + pt(22), w: 9, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });

// 数式
slide03.addText('期待値 = 確率 × オッズ', { x: 0, y: 1.4, w: 10, h: 0.6, fontSize: 36, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// コンポーネントボックス
const compY = 2.2, compW = 3.8, compH = 1.1;
slide03.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: compY, w: compW, h: compH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide03.addText('確率', { x: 1.0, y: compY + 0.15, w: compW, h: 0.35, fontSize: 18, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide03.addText('機械学習等で算出\n（強力なモデルを作れる人が有利）', { x: 1.1, y: compY + 0.55, w: compW - 0.2, h: 0.5, fontSize: 12, color: COLORS.secondary, fontFace: FONT, align: 'center' });

slide03.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 5.2, y: compY, w: compW, h: compH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide03.addText('オッズ', { x: 5.2, y: compY + 0.15, w: compW, h: 0.35, fontSize: 18, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide03.addText('全員の投票で決まる\n（変動する）', { x: 5.3, y: compY + 0.55, w: compW - 0.2, h: 0.5, fontSize: 12, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 歪みボックス
const distY = 3.5;
slide03.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: distY, w: 8.8, h: 1.3, fill: { color: '1a2f2d' }, line: { color: '3d5c59', pt: 1 } });
slide03.addText('「歪み」を取る = 期待値 > 1 を狙う', { x: 0.6, y: distY + 0.15, w: 8.8, h: 0.35, fontSize: 14, bold: true, color: COLORS.primary, fontFace: FONT, align: 'center' });
slide03.addText('AIが「勝率15%」と計算、オッズ10倍（市場は10%と見てる）', { x: 0.6, y: distY + 0.55, w: 8.8, h: 0.3, fontSize: 12, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide03.addText('→ 15% × 10倍 = 150%（期待値 > 1）→ 理論上勝てる', { x: 0.6, y: distY + 0.9, w: 8.8, h: 0.3, fontSize: 13, color: COLORS.accent, fontFace: FONT, align: 'center' });

// ===== Slide 04: オッズの構造 =====
const slide04 = pptx.addSlide();
slide04.background = { path: BG_IMAGE };
slide04.addText('KEY CONCEPT', { x: pt(40), y: pt(36), w: 9, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide04.addText('オッズの構造 — パリミュチュエル方式', { x: pt(40), y: pt(36) + pt(22), w: 9, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });
slide04.addImage({ path: `${IMG_PATH}/pari-mutuel-diagram.png`, x: 0.4, y: 1.15, w: 9.2, h: 4.3 });

// ===== Slide 05: 競馬AI黎明期 =====
const slide05 = pptx.addSlide();
slide05.background = { path: BG_IMAGE };
const s5x = pt(40), s5y = pt(36);
// コンテンツ幅: 720pt - 80pt = 640pt = 8.89"
const s5contentW = pt(640);
slide05.addText('HISTORY', { x: s5x, y: s5y, w: 6, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide05.addText('競馬AI黎明期（〜2016年頃）', { x: s5x, y: s5y + pt(18), w: 6, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });
slide05.addText('1995年「馬王」登場（荒井俊也氏）— 統計的競馬予測の先駆け', { x: s5x, y: s5y + pt(60), w: 6.5, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// 馬王画像 (右上、paddingに合わせてy=0.5から開始)
slide05.addImage({ path: `${IMG_PATH}/baou-book.jpg`, x: 7.8, y: s5y, w: 1.6, h: 1.8 });
slide05.addText('出典: Amazon / baoland.com', { x: 7.8, y: s5y + 1.85, w: 1.6, h: 0.15, fontSize: 7, color: COLORS.muted, fontFace: FONT, align: 'center' });

// Stats (ヘッダー後: 約2.5"から開始)
// stat-card幅: (640pt - 16pt gap) / 2 = 312pt = 4.33"
const s5statY = 2.5, s5statW = pt(312), s5statH = 0.7, s5statGap = pt(16);
slide05.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s5x, y: s5statY, w: s5statW, h: s5statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide05.addText('キーワード', { x: s5x, y: s5statY + 0.08, w: s5statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide05.addText('データ格差', { x: s5x, y: s5statY + 0.32, w: s5statW, h: 0.3, fontSize: 16, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

slide05.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s5x + s5statW + s5statGap, y: s5statY, w: s5statW, h: s5statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide05.addText('2013年 話題になった事件', { x: s5x + s5statW + s5statGap, y: s5statY + 0.08, w: s5statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide05.addText('データ分析で1.4億円稼いだ人', { x: s5x + s5statW + s5statGap, y: s5statY + 0.32, w: s5statW, h: 0.3, fontSize: 14, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// 箇条書き (Stats後: 約3.4"から開始)
const s5listY = 3.4, s5lh = 0.35;
slide05.addText('•', { x: s5x, y: s5listY, w: 0.2, h: s5lh, fontSize: 12, color: COLORS.accent, fontFace: FONT });
slide05.addText('当時はデータが高価 → 持っている人だけが勝てた', { x: s5x + 0.22, y: s5listY, w: s5contentW, h: s5lh, fontSize: 12, color: COLORS.secondary, fontFace: FONT });
slide05.addText('•', { x: s5x, y: s5listY + s5lh, w: 0.2, h: s5lh, fontSize: 12, color: COLORS.accent, fontFace: FONT });
slide05.addText('参入者が少ない → オッズ変動も小さい', { x: s5x + 0.22, y: s5listY + s5lh, w: s5contentW, h: s5lh, fontSize: 12, color: COLORS.secondary, fontFace: FONT });

// ハイライトボックス (下部に配置、コンテンツ幅いっぱい)
slide05.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s5x, y: 4.55, w: s5contentW, h: 0.55, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide05.addText('「データを持っている者が勝つ」時代', { x: s5x, y: 4.55, w: s5contentW, h: 0.55, fontSize: 13, color: COLORS.primary, fontFace: FONT, align: 'center', valign: 'middle' });

// ===== Slide 06: AI Mamba =====
const slide06 = pptx.addSlide();
slide06.background = { path: BG_IMAGE };
// 基準値 (すべてpt単位で計算後にpt()でインチ変換)
const s6x = pt(40);  // 左パディング
const s6leftW = pt(436);  // 左側幅: 640-180-24
const s6rightW = pt(180); // 右側幅
const s6gap = pt(24);     // gap
const s6rightX = pt(40 + 436 + 24); // 右側開始X: 500pt

// ヘッダー部分 (厳密なY座標)
slide06.addText('CASE STUDY', { x: s6x, y: pt(36), w: s6leftW, h: pt(10), fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide06.addText('AI Mamba登場（2018年〜）', { x: s6x, y: pt(56), w: s6leftW, h: pt(28), fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });
slide06.addText('ドワンゴのエンジニアチームが開発', { x: s6x, y: pt(90), w: s6leftW, h: pt(12), fontSize: 12, color: COLORS.secondary, fontFace: FONT });

// 右側: Mamba画像 (垂直中央: 画像220pt+source11pt=231pt, content-row251pt, offset=(251-231)/2=10pt)
slide06.addImage({ path: `${IMG_PATH}/mamba-illust.png`, x: s6rightX, y: pt(128), w: s6rightW, h: pt(220) });
slide06.addText('出典: mamba.jinkochinobokin.nicovideo.jp', { x: s6rightX, y: pt(352), w: s6rightW, h: pt(10), fontSize: 7, color: COLORS.muted, fontFace: FONT, align: 'center' });

// 左側: Hero stat (content-row開始: 118pt)
slide06.addText('+730万円', { x: s6x, y: pt(118), w: s6leftW, h: pt(48), fontSize: 48, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide06.addText('ニコ生企画 3ヶ月の成果', { x: s6x, y: pt(170), w: s6leftW, h: pt(11), fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 箇条書き (ul開始: 197pt, li高さ: 11*1.7=18.7pt)
const s6listY = pt(197), s6lh = pt(18.7);
slide06.addText('•', { x: s6x, y: s6listY, w: pt(16), h: s6lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide06.addText('機械学習で約3,000の特徴量を使用（手法は非公開）', { x: s6x + pt(16), y: s6listY, w: s6leftW - pt(16), h: s6lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });
slide06.addText('•', { x: s6x, y: s6listY + s6lh, w: pt(16), h: s6lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide06.addText('メディアで話題に →「AIで競馬に勝てる」という夢が広まる', { x: s6x + pt(16), y: s6listY + s6lh, w: s6leftW - pt(16), h: s6lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });
slide06.addText('•', { x: s6x, y: s6listY + s6lh * 2, w: pt(16), h: s6lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide06.addText('ここから松風、AlphaImpactなど派生プロジェクトが誕生', { x: s6x + pt(16), y: s6listY + s6lh * 2, w: s6leftW - pt(16), h: s6lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// highlight-box (margin-top:auto → 下部固定: 337pt, 高さ: padding10*2+font12=32pt)
slide06.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s6x, y: pt(337), w: s6leftW, h: pt(32), fill: { color: '2d2a1a' }, line: { color: COLORS.warning, pt: 1 } });
slide06.addText('競馬AIブームの火付け役', { x: s6x, y: pt(337), w: s6leftW, h: pt(32), fontSize: 12, color: COLORS.warning, fontFace: FONT, align: 'center', valign: 'middle' });

// ===== Slide 07: 松風 =====
const slide07 = pptx.addSlide();
slide07.background = { path: BG_IMAGE };
const s7x = pt(40), s7y = pt(36);
slide07.addText('CASE STUDY', { x: s7x, y: s7y, w: 6, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide07.addText('松風の衝撃（2017年〜）', { x: s7x, y: s7y + pt(18), w: 6.5, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });
slide07.addText('毎レース締切直前に買い目を無料公開（2017年〜2020年8月に公開終了）', { x: s7x, y: s7y + pt(58), w: 7, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// 松風アイコン
slide07.addImage({ path: `${IMG_PATH}/matsukaze-icon.jpg`, x: 8.0, y: 0.4, w: 1.1, h: 1.1 });
slide07.addText('出典: X (@matsukaze_f)', { x: 7.5, y: 1.55, w: 2, h: 0.15, fontSize: 7, color: COLORS.muted, fontFace: FONT, align: 'center' });

// Stats
const s7statY = 1.8, s7statW = 3.8, s7statH = 0.75;
slide07.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s7x, y: s7statY, w: s7statW, h: s7statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide07.addText('2019年', { x: s7x, y: s7statY + 0.08, w: s7statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide07.addText('+1,739万円', { x: s7x, y: s7statY + 0.32, w: s7statW, h: 0.35, fontSize: 20, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

slide07.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s7x + s7statW + pt(16), y: s7statY, w: s7statW, h: s7statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide07.addText('2020年', { x: s7x + s7statW + pt(16), y: s7statY + 0.08, w: s7statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide07.addText('+2億円', { x: s7x + s7statW + pt(16), y: s7statY + 0.32, w: s7statW, h: 0.35, fontSize: 20, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// Big stat
const s7bigY = s7statY + s7statH + pt(12);
slide07.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s7x, y: s7bigY, w: 8.3, h: 0.7, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide07.addText('2020年10月〜12月（3ヶ月間）', { x: s7x, y: s7bigY + 0.08, w: 8.3, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide07.addText('購入 17.5億円 → 払戻 20.8億円 → 収支 +3.3億円', { x: s7x, y: s7bigY + 0.32, w: 8.3, h: 0.3, fontSize: 14, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// 箇条書き
const s7listY = s7bigY + 0.85, s7lh = 0.32;
slide07.addText('•', { x: s7x, y: s7listY, w: 0.2, h: s7lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide07.addText('AI Mambaから派生、個人開発者がここまでスケール', { x: s7x + 0.22, y: s7listY, w: 7, h: s7lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });
slide07.addText('•', { x: s7x, y: s7listY + s7lh, w: 0.2, h: s7lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide07.addText('2020年8月に買い目公開を終了 → 自分専用に', { x: s7x + 0.22, y: s7listY + s7lh, w: 7, h: s7lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// バッジ
slide07.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 1.2, y: 4.65, w: 7, h: 0.5, fill: { color: '2d2a1a' }, line: { color: COLORS.warning, pt: 1 } });
slide07.addText('「先行者利益で札束勢になった」象徴', { x: 1.2, y: 4.65, w: 7, h: 0.5, fontSize: 12, color: COLORS.warning, fontFace: FONT, align: 'center', valign: 'middle' });

// ===== Slide 08: 競馬AIゆま =====
const slide08 = pptx.addSlide();
slide08.background = { path: BG_IMAGE };
const s8x = pt(40), s8y = pt(36);
slide08.addText('CASE STUDY', { x: s8x, y: s8y, w: 6, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide08.addText('競馬AIゆまの終焉（2022年）', { x: s8x, y: s8y + pt(18), w: 6.5, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });
slide08.addText('AIが算出した勝率を無料公開するサービス（2018年〜2022年）', { x: s8x, y: s8y + pt(58), w: 7, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// ゆまアイコン
slide08.addImage({ path: `${IMG_PATH}/yuma-icon.png`, x: 8.0, y: 0.4, w: 1.1, h: 1.1 });
slide08.addText('出典: はてなブログ', { x: 7.5, y: 1.55, w: 2, h: 0.15, fontSize: 7, color: COLORS.muted, fontFace: FONT, align: 'center' });

// Stats
const s8statY = 1.8, s8statW = 3.8, s8statH = 0.75;
slide08.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s8x, y: s8statY, w: s8statW, h: s8statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide08.addText('通算57,000レース', { x: s8x, y: s8statY + 0.08, w: s8statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide08.addText('回収率 約111%', { x: s8x, y: s8statY + 0.32, w: s8statW, h: 0.35, fontSize: 18, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

slide08.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s8x + s8statW + pt(16), y: s8statY, w: s8statW, h: s8statH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide08.addText('4年間の運営実績', { x: s8x + s8statW + pt(16), y: s8statY + 0.08, w: s8statW, h: 0.2, fontSize: 10, color: COLORS.secondary, fontFace: FONT, align: 'center' });
slide08.addText('5000万PV / 5万フォロワー', { x: s8x + s8statW + pt(16), y: s8statY + 0.32, w: s8statW, h: 0.35, fontSize: 15, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// End box
const s8endY = s8statY + s8statH + pt(12);
slide08.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: s8x, y: s8endY, w: 8.3, h: 0.75, fill: { color: '2d1a1f' }, line: { color: COLORS.danger, pt: 1 } });
slide08.addText('2022年8月にサービス終了', { x: s8x, y: s8endY + 0.1, w: 8.3, h: 0.3, fontSize: 16, bold: true, color: COLORS.danger, fontFace: FONT, align: 'center' });
slide08.addText('終了理由:「利用者増加によるオッズ低下」', { x: s8x, y: s8endY + 0.42, w: 8.3, h: 0.25, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// 箇条書き
const s8listY = s8endY + 0.95, s8lh = 0.32;
slide08.addText('•', { x: s8x, y: s8listY, w: 0.2, h: s8lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide08.addText('「勝てない予想を提供したくない」', { x: s8x + 0.22, y: s8listY, w: 7, h: s8lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });
slide08.addText('•', { x: s8x, y: s8listY + s8lh, w: 0.2, h: s8lh, fontSize: 11, color: COLORS.accent, fontFace: FONT });
slide08.addText('確率を公開すること自体が自滅につながる', { x: s8x + 0.22, y: s8listY + s8lh, w: 7, h: s8lh, fontSize: 11, color: COLORS.secondary, fontFace: FONT });

// ===== Slide 09: なぜ勝てないのか =====
const slide09 = pptx.addSlide();
slide09.background = { path: BG_IMAGE };
const s9x = pt(40), s9y = pt(36);
slide09.addText('CONCLUSION', { x: s9x, y: s9y, w: 9, h: 0.2, fontSize: 10, color: COLORS.danger, fontFace: FONT, charSpacing: 4 });
slide09.addText('なぜ今から参入しても勝てないのか', { x: s9x, y: s9y + pt(22), w: 9, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });

// 箇条書き（マーカーは赤、3番目だけハイライト）
const s9listY = 1.35, s9lh = 0.4;
slide09.addText('•', { x: s9x, y: s9listY, w: 0.2, h: s9lh, fontSize: 12, color: COLORS.danger, fontFace: FONT });
slide09.addText('参入者増加 → 同じ買い目に殺到 → オッズ急落', { x: s9x + 0.25, y: s9listY, w: 8, h: s9lh, fontSize: 12, color: COLORS.secondary, fontFace: FONT });

slide09.addText('•', { x: s9x, y: s9listY + s9lh, w: 0.2, h: s9lh, fontSize: 12, color: COLORS.danger, fontFace: FONT });
slide09.addText('計算時点と締切時点でオッズがズレる → 期待値が取れない', { x: s9x + 0.25, y: s9listY + s9lh, w: 8, h: s9lh, fontSize: 12, color: COLORS.secondary, fontFace: FONT });

slide09.addText('•', { x: s9x, y: s9listY + s9lh * 2, w: 0.2, h: s9lh, fontSize: 12, color: COLORS.danger, fontFace: FONT });
slide09.addText('資金力があれば高精度なモデル開発 & 投票後のオッズ変動も予測可能', { x: s9x + 0.25, y: s9listY + s9lh * 2, w: 8, h: s9lh, fontSize: 12, color: COLORS.accent, fontFace: FONT, bold: true });

slide09.addText('•', { x: s9x, y: s9listY + s9lh * 3, w: 0.2, h: s9lh, fontSize: 12, color: COLORS.danger, fontFace: FONT });
slide09.addText('→ 札束勢は有利、後発組は不利という構造', { x: s9x + 0.25, y: s9listY + s9lh * 3, w: 8, h: s9lh, fontSize: 12, color: COLORS.secondary, fontFace: FONT });

// Result box
slide09.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 3.15, w: 8.4, h: 0.7, fill: { color: '2d1a1f' }, line: { color: COLORS.danger, pt: 1 } });
slide09.addText('今から普通に参入しても、オッズ変動を予測できず期待値が取れない', { x: 0.8, y: 3.15, w: 8.4, h: 0.7, fontSize: 15, bold: true, color: COLORS.danger, fontFace: FONT, align: 'center', valign: 'middle' });

// Note
slide09.addText('競馬AIゆまも「公開したら勝てなくなる」と判断した世界', { x: 0, y: 4.1, w: 10, h: 0.35, fontSize: 12, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// ===== Slide 10: まとめ =====
const slide10 = pptx.addSlide();
slide10.background = { path: BG_IMAGE };
const s10x = pt(40), s10y = pt(36);
slide10.addText('SUMMARY', { x: s10x, y: s10y, w: 9, h: 0.2, fontSize: 10, color: COLORS.accent, fontFace: FONT, charSpacing: 4 });
slide10.addText('まとめ', { x: s10x, y: s10y + pt(22), w: 9, h: 0.5, fontSize: 28, bold: true, color: COLORS.primary, fontFace: FONT });

// Formula
slide10.addText('期待値 = 確率 × オッズ', { x: 0, y: 1.2, w: 10, h: 0.5, fontSize: 28, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });

// ポイントカード
const s10cardY = 1.85, s10cardW = 4.2, s10cardH = 0.95;
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: s10cardY, w: s10cardW, h: s10cardH, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide10.addText('確率', { x: 0.7, y: s10cardY + 0.12, w: s10cardW, h: 0.3, fontSize: 14, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide10.addText('資金力があれば\n高精度なモデルを開発できる', { x: 0.8, y: s10cardY + 0.45, w: s10cardW - 0.2, h: 0.45, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 5.1, y: s10cardY, w: s10cardW, h: s10cardH, fill: { color: '1a2f2d' }, line: { color: COLORS.accent, pt: 1 } });
slide10.addText('オッズ', { x: 5.1, y: s10cardY + 0.12, w: s10cardW, h: 0.3, fontSize: 14, bold: true, color: COLORS.accent, fontFace: FONT, align: 'center' });
slide10.addText('資金力があれば\n投票後の変動を予測できる', { x: 5.2, y: s10cardY + 0.45, w: s10cardW - 0.2, h: 0.45, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center' });

// タイムライン
const s10tlY = 3.0, s10tlW = 1.8, s10tlH = 0.45;
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 1.3, y: s10tlY, w: s10tlW, h: s10tlH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide10.addText('データ格差', { x: 1.3, y: s10tlY, w: s10tlW, h: s10tlH, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center', valign: 'middle' });

slide10.addText('→', { x: 3.15, y: s10tlY, w: 0.5, h: s10tlH, fontSize: 14, color: COLORS.accent, fontFace: FONT, align: 'center', valign: 'middle' });

slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 3.7, y: s10tlY, w: s10tlW, h: s10tlH, fill: { color: COLORS.boxBg }, line: { color: COLORS.boxBorder, pt: 1 } });
slide10.addText('技術格差', { x: 3.7, y: s10tlY, w: s10tlW, h: s10tlH, fontSize: 11, color: COLORS.secondary, fontFace: FONT, align: 'center', valign: 'middle' });

slide10.addText('→', { x: 5.55, y: s10tlY, w: 0.5, h: s10tlH, fontSize: 14, color: COLORS.accent, fontFace: FONT, align: 'center', valign: 'middle' });

slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 6.1, y: s10tlY, w: s10tlW, h: s10tlH, fill: { color: '2d1a1f' }, line: { color: COLORS.danger, pt: 1 } });
slide10.addText('資金力格差', { x: 6.1, y: s10tlY, w: s10tlW, h: s10tlH, fontSize: 11, bold: true, color: COLORS.danger, fontFace: FONT, align: 'center', valign: 'middle' });

// Conclusion
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 1.5, y: 3.65, w: 7, h: 0.55, fill: { color: '2d1a1f' }, line: { color: COLORS.danger, pt: 1 } });
slide10.addText('今から普通に参入しても勝てない', { x: 1.5, y: 3.65, w: 7, h: 0.55, fontSize: 14, bold: true, color: COLORS.danger, fontFace: FONT, align: 'center', valign: 'middle' });

// Footer
slide10.addText('ギャンブルは無理のない範囲で楽しみましょう', { x: 0, y: 4.85, w: 10, h: 0.3, fontSize: 12, color: COLORS.warning, fontFace: FONT, align: 'center' });

// PPTXファイルを保存
pptx.writeFile({ fileName: '/home/inoue-d/dev/my-samples/presentations/20260120_lt/output/03_presentation.pptx' })
  .then(() => console.log('PPTX generated: output/03_presentation.pptx'))
  .catch(err => console.error(err));
