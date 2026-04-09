const path = require('node:path');
const { test, expect, _electron: electron } = require('@playwright/test');

const appDir = path.join(__dirname, '..', '..');
const sampleDeck = path.join(appDir, 'tests', 'fixtures', 'sample-deck.pptx');
const alternateDeck = path.join(appDir, 'tests', 'fixtures', 'sample-deck-alt.pptx');

let electronApp;
let window;

test.beforeEach(async () => {
  electronApp = await electron.launch({
    args: ['.'],
    cwd: appDir,
  });

  window = await electronApp.firstWindow();
  await window.waitForLoadState('domcontentloaded');
  await window.bringToFront();
  await closeAllTabs();
});

test.afterEach(async () => {
  await electronApp?.close();
});

async function getBrowserZoomFactor() {
  return electronApp.evaluate(({ BrowserWindow }) => {
    return BrowserWindow.getAllWindows()[0].webContents.zoomFactor;
  });
}

async function getCanvasViewportSize() {
  return window.locator('#slide-canvas').evaluate((canvas) => {
    const rect = canvas.getBoundingClientRect();

    return {
      width: rect.width,
      height: rect.height,
    };
  });
}

function tabChip(name) {
  return window.locator('.tab-chip').filter({ has: window.locator('.tab-chip-name', { hasText: name }) });
}

async function waitForCanvasToRender() {
  await expect.poll(async () => {
    return window.locator('#slide-canvas').evaluate((canvas) => canvas.width);
  }).toBeGreaterThan(0);

  await expect.poll(async () => {
    return window.locator('#slide-canvas').evaluate((canvas) => canvas.height);
  }).toBeGreaterThan(0);
}

async function openDecks(files) {
  await window.locator('#file-input').setInputFiles(files);
  await expect(window.locator('#loading-state')).toHaveClass(/hidden/);
  await waitForCanvasToRender();
}

async function openSampleDeck() {
  await openDecks(sampleDeck);
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 3');
  await expect(window.locator('#error-state')).toHaveClass(/hidden/);
  await expect(tabChip('sample-deck.pptx')).toHaveCount(1);
  await expect(tabChip('sample-deck.pptx')).toHaveClass(/active/);
}

async function closeAllTabs() {
  while (await window.locator('.tab-chip').count()) {
    const currentCount = await window.locator('.tab-chip').count();
    await window.locator('.tab-close-button').first().click();
    await expect(window.locator('.tab-chip')).toHaveCount(currentCount - 1);
  }

  await expect(window.locator('.tab-chip')).toHaveCount(0);
}

test('opens and renders a real pptx deck in a tab', async () => {
  await openSampleDeck();

  await expect(window.locator('#prev-button')).toBeDisabled();
  await expect(window.locator('#next-button')).toBeEnabled();
  await expect(tabChip('sample-deck.pptx')).toHaveClass(/active/);
});

test('supports slide navigation with buttons and keyboard', async () => {
  await openSampleDeck();
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.locator('#next-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 3');
  await expect(window.locator('#prev-button')).toBeEnabled();
  await expect(window.locator('#next-button')).toBeEnabled();
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.keyboard.press('ArrowRight');
  await expect(window.locator('#slide-status')).toHaveText('Slide 3 / 3');
  await expect(window.locator('#next-button')).toBeDisabled();

  await window.keyboard.press('ArrowLeft');
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 3');
  await expect.poll(getBrowserZoomFactor).toBe(1);
});

test('updates zoom controls and resets with fit slide', async () => {
  await openSampleDeck();
  const initialCanvasSize = await getCanvasViewportSize();

  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await window.locator('#zoom-in-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 110%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect.poll(async () => (await getCanvasViewportSize()).width).toBeGreaterThan(initialCanvasSize.width);

  await window.locator('#zoom-out-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect.poll(async () => Math.abs((await getCanvasViewportSize()).width - initialCanvasSize.width)).toBeLessThan(20);

  await window.locator('#zoom-in-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 110%');
  await window.locator('#fit-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
});

test('supports ctrl/cmd zoom shortcuts', async () => {
  await openSampleDeck();

  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';

  await window.keyboard.press(`${modifier}+=`);
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 110%');
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.keyboard.press(`${modifier}+-`);
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.keyboard.press(`${modifier}+=`);
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 110%');
  await window.keyboard.press(`${modifier}+0`);
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
});

test('keeps chrome controls visible while navigating and zooming', async () => {
  await openSampleDeck();
  const initialCanvasSize = await getCanvasViewportSize();

  await window.locator('#next-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 3');
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.locator('#zoom-out-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 90%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect.poll(async () => (await getCanvasViewportSize()).width).toBeLessThan(initialCanvasSize.width);

  await expect(window.locator('#prev-button')).toBeVisible();
  await expect(window.locator('#next-button')).toBeVisible();
  await expect(window.locator('#zoom-out-button')).toBeVisible();
});

test('keeps zoom controls visible when zooming in repeatedly', async () => {
  await openSampleDeck();

  for (let index = 0; index < 8; index += 1) {
    await window.locator('#zoom-in-button').click();
  }

  await expect(window.locator('#zoom-status')).toHaveText('Zoom 180%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect(window.locator('#zoom-in-button')).toBeVisible();
  await expect(window.locator('#zoom-out-button')).toBeVisible();
  await expect(window.locator('#zoom-in-button')).toBeInViewport();
  await expect(window.locator('#zoom-out-button')).toBeInViewport();
});

test('opens multiple presentations as separate tabs and switches between them', async () => {
  await openDecks([sampleDeck, alternateDeck]);

  await expect(window.locator('.tab-chip')).toHaveCount(2);
  await expect(tabChip('sample-deck.pptx')).toHaveCount(1);
  await expect(tabChip('sample-deck-alt.pptx')).toHaveCount(1);
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 3');
  await expect(tabChip('sample-deck-alt.pptx')).toHaveClass(/active/);

  await tabChip('sample-deck.pptx').locator('.tab-chip-button').click();
  await expect(tabChip('sample-deck.pptx')).toHaveClass(/active/);
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 3');
});

test('preserves slide position and zoom independently per tab', async () => {
  await openDecks([sampleDeck, alternateDeck]);

  await tabChip('sample-deck.pptx').locator('.tab-chip-button').click();
  await window.locator('#next-button').click();
  await window.locator('#next-button').click();
  await window.locator('#zoom-in-button').click();
  await window.locator('#zoom-in-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 3 / 3');
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 120%');

  await tabChip('sample-deck-alt.pptx').locator('.tab-chip-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 3');
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');

  await window.locator('#next-button').click();
  await window.locator('#zoom-out-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 3');
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 90%');

  await tabChip('sample-deck.pptx').locator('.tab-chip-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 3 / 3');
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 120%');

  await tabChip('sample-deck-alt.pptx').locator('.tab-chip-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 3');
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 90%');
});

test('closing tabs keeps the remaining presentation active', async () => {
  await openDecks([sampleDeck, alternateDeck]);

  await tabChip('sample-deck-alt.pptx').locator('.tab-close-button').click();
  await expect(window.locator('.tab-chip')).toHaveCount(1);
  await expect(tabChip('sample-deck.pptx')).toHaveClass(/active/);

  await tabChip('sample-deck.pptx').locator('.tab-close-button').click();
  await expect(window.locator('.tab-chip')).toHaveCount(0);
  await expect(window.locator('#slide-status')).toHaveText('Slide 0 / 0');
  await expect(window.locator('#canvas-stage')).toHaveClass(/hidden/);
});
