const fs = require('node:fs');
const path = require('node:path');
const pptxgen = require('pptxgenjs');

async function main() {
  const fixtureDir = path.join(__dirname, 'fixtures');
  const outputPath = path.join(fixtureDir, 'sample-deck.pptx');

  fs.mkdirSync(fixtureDir, { recursive: true });

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'OpenCode';
  pptx.subject = 'Regression Test Deck';
  pptx.title = 'Sample Deck';
  pptx.company = 'OpenCode';
  pptx.lang = 'en-US';
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
    lang: 'en-US',
  };

  const slide1 = pptx.addSlide();
  slide1.background = { color: '0F172A' };
  slide1.addText('Quarterly Product Review', {
    x: 0.6,
    y: 0.6,
    w: 11,
    h: 0.8,
    color: 'F8FAFC',
    fontSize: 24,
    bold: true,
  });
  slide1.addText('Regression fixture slide 1', {
    x: 0.6,
    y: 1.5,
    w: 5,
    h: 0.5,
    color: 'CBD5E1',
    fontSize: 14,
  });
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 0.7,
    y: 2.2,
    w: 3.8,
    h: 1.8,
    rectRadius: 0.12,
    fill: { color: '2563EB' },
    line: { color: '60A5FA', pt: 1 },
  });
  slide1.addText('Launch checklist', {
    x: 1,
    y: 2.55,
    w: 3,
    h: 0.4,
    color: 'FFFFFF',
    fontSize: 18,
    bold: true,
  });
  slide1.addText('UI polish complete\nPlayback stable\nRegression tests green', {
    x: 1,
    y: 3.05,
    w: 3,
    h: 0.8,
    color: 'DBEAFE',
    fontSize: 14,
    bullet: { indent: 18 },
    breakLine: false,
  });

  const slide2 = pptx.addSlide();
  slide2.background = { color: '111827' };
  slide2.addText('Regression fixture slide 2', {
    x: 0.6,
    y: 0.6,
    w: 11,
    h: 0.6,
    color: 'F9FAFB',
    fontSize: 24,
    bold: true,
  });
  slide2.addText('Navigation and zoom validation', {
    x: 0.6,
    y: 1.4,
    w: 5,
    h: 0.4,
    color: 'D1D5DB',
    fontSize: 14,
  });
  slide2.addShape(pptx.ShapeType.rect, {
    x: 0.8,
    y: 2.2,
    w: 5.5,
    h: 2.8,
    fill: { color: '7C3AED' },
    line: { color: 'C4B5FD', pt: 1.5 },
  });
  slide2.addText('Slide two content block', {
    x: 1.2,
    y: 3.1,
    w: 4.6,
    h: 0.5,
    color: 'FFFFFF',
    align: 'center',
    fontSize: 20,
    bold: true,
  });

  await pptx.writeFile({ fileName: outputPath });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
