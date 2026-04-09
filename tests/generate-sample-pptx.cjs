const fs = require('node:fs');
const path = require('node:path');
const pptxgen = require('pptxgenjs');

function createDeck({ title, subtitle, outputPath, colors, slide2Title, slide3Title }) {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'OpenCode';
  pptx.subject = 'Regression Test Deck';
  pptx.title = title;
  pptx.company = 'OpenCode';
  pptx.lang = 'en-US';
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
    lang: 'en-US',
  };

  const slide1 = pptx.addSlide();
  slide1.background = { color: colors.background1 };
  slide1.addText(title, {
    x: 0.6,
    y: 0.6,
    w: 11,
    h: 0.8,
    color: 'F8FAFC',
    fontSize: 24,
    bold: true,
  });
  slide1.addText(subtitle, {
    x: 0.6,
    y: 1.5,
    w: 7,
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
    fill: { color: colors.shape1 },
    line: { color: colors.line1, pt: 1 },
  });
  slide1.addText('Regression coverage', {
    x: 1,
    y: 2.55,
    w: 3,
    h: 0.4,
    color: 'FFFFFF',
    fontSize: 18,
    bold: true,
  });
  slide1.addText('Initial render\nIndependent tabs\nStable controls', {
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
  slide2.background = { color: colors.background2 };
  slide2.addText(slide2Title, {
    x: 0.6,
    y: 0.6,
    w: 11,
    h: 0.6,
    color: 'F9FAFB',
    fontSize: 24,
    bold: true,
  });
  slide2.addText('Navigation regression fixture', {
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
    fill: { color: colors.shape2 },
    line: { color: colors.line2, pt: 1.5 },
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

  const slide3 = pptx.addSlide();
  slide3.background = { color: colors.background3 };
  slide3.addText(slide3Title, {
    x: 0.6,
    y: 0.7,
    w: 9,
    h: 0.6,
    color: 'FFFFFF',
    fontSize: 24,
    bold: true,
  });
  slide3.addText('Used to prove active tab state is isolated.', {
    x: 0.6,
    y: 1.5,
    w: 6,
    h: 0.4,
    color: 'E5E7EB',
    fontSize: 14,
  });
  slide3.addShape(pptx.ShapeType.ellipse, {
    x: 1,
    y: 2.2,
    w: 4.2,
    h: 2.4,
    fill: { color: colors.shape3 },
    line: { color: colors.line3, pt: 1.5 },
  });

  return pptx.writeFile({ fileName: outputPath });
}

async function main() {
  const fixtureDir = path.join(__dirname, 'fixtures');
  fs.mkdirSync(fixtureDir, { recursive: true });

  await createDeck({
    title: 'Quarterly Product Review',
    subtitle: 'Regression fixture slide 1',
    outputPath: path.join(fixtureDir, 'sample-deck.pptx'),
    slide2Title: 'Regression fixture slide 2',
    slide3Title: 'Regression fixture slide 3',
    colors: {
      background1: '0F172A',
      background2: '111827',
      background3: '172554',
      shape1: '2563EB',
      shape2: '7C3AED',
      shape3: '0EA5E9',
      line1: '60A5FA',
      line2: 'C4B5FD',
      line3: '7DD3FC',
    },
  });

  await createDeck({
    title: 'Customer Onboarding Plan',
    subtitle: 'Alternate tab fixture slide 1',
    outputPath: path.join(fixtureDir, 'sample-deck-alt.pptx'),
    slide2Title: 'Alternate fixture slide 2',
    slide3Title: 'Alternate fixture slide 3',
    colors: {
      background1: '1F2937',
      background2: '052E16',
      background3: '3F1D2E',
      shape1: '10B981',
      shape2: '84CC16',
      shape3: 'EC4899',
      line1: '6EE7B7',
      line2: 'BEF264',
      line3: 'F9A8D4',
    },
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
