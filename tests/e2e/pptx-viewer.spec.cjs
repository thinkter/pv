const path = require('node:path');
const { test, expect, _electron: electron } = require('@playwright/test');

const appDir = path.join(__dirname, '..', '..');
const sampleDeck = path.join(appDir, 'tests', 'fixtures', 'sample-deck.pptx');

let electronApp;
let window;

test.beforeAll(async () => {
  electronApp = await electron.launch({
    args: ['.'],
    cwd: appDir,
  });

  window = await electronApp.firstWindow();
  await window.waitForLoadState('domcontentloaded');
});

test.afterAll(async () => {
  await electronApp?.close();
});

test.beforeEach(async () => {
  await window.bringToFront();
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

async function openSampleDeck() {
  await window.locator('#file-input').setInputFiles(sampleDeck);
  await expect(window.locator('#file-name')).toContainText('sample-deck.pptx');
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 2');
  await expect(window.locator('#loading-state')).toHaveClass(/hidden/);

  await expect
    .poll(async () => {
      return window.locator('#slide-canvas').evaluate((canvas) => ({
        width: canvas.width,
        height: canvas.height,
      }));
    })
    .toMatchObject({
      width: expect.any(Number),
      height: expect.any(Number),
    });

  const canvasSize = await window.locator('#slide-canvas').evaluate((canvas) => ({
    width: canvas.width,
    height: canvas.height,
  }));

  expect(canvasSize.width).toBeGreaterThan(0);
  expect(canvasSize.height).toBeGreaterThan(0);
}

test('opens and renders a real pptx deck', async () => {
  await openSampleDeck();

  await expect(window.locator('#prev-button')).toBeDisabled();
  await expect(window.locator('#next-button')).toBeEnabled();
  await expect(window.locator('#error-state')).toHaveClass(/hidden/);
});

test('supports slide navigation with buttons and keyboard', async () => {
  await openSampleDeck();
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.locator('#next-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 2');
  await expect(window.locator('#prev-button')).toBeEnabled();
  await expect(window.locator('#next-button')).toBeDisabled();
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.keyboard.press('ArrowLeft');
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 2');
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
  await expect
    .poll(async () => (await getCanvasViewportSize()).width)
    .toBeGreaterThan(initialCanvasSize.width);

  await window.locator('#zoom-out-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect
    .poll(async () => (await getCanvasViewportSize()).width)
    .toBeCloseTo(initialCanvasSize.width, 0);

  await window.locator('#zoom-in-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 110%');
  await window.locator('#fit-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 100%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect
    .poll(async () => (await getCanvasViewportSize()).width)
    .toBeCloseTo(initialCanvasSize.width, 0);
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

test('keeps the app chrome visible when navigation and zoom-out controls are clicked', async () => {
  await openSampleDeck();
  const initialCanvasSize = await getCanvasViewportSize();

  await window.locator('#next-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 2 / 2');
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.locator('#prev-button').click();
  await expect(window.locator('#slide-status')).toHaveText('Slide 1 / 2');
  await expect.poll(getBrowserZoomFactor).toBe(1);

  await window.locator('#zoom-out-button').click();
  await expect(window.locator('#zoom-status')).toHaveText('Zoom 90%');
  await expect.poll(getBrowserZoomFactor).toBe(1);
  await expect
    .poll(async () => (await getCanvasViewportSize()).width)
    .toBeLessThan(initialCanvasSize.width);

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
