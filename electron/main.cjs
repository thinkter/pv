const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');

const DEV_SERVER_URL = process.env.ELECTRON_RENDERER_URL;

function resolveLaunchPresentationPath(argv = process.argv) {
  for (const arg of argv.slice(1)) {
    if (!arg || arg.startsWith('-')) {
      continue;
    }

    const candidatePath = path.resolve(arg);
    if (!candidatePath.toLowerCase().endsWith('.pptx')) {
      continue;
    }

    if (fsSync.existsSync(candidatePath)) {
      return candidatePath;
    }
  }

  return null;
}

let launchPresentationPath = resolveLaunchPresentationPath();

function createWindow() {
  const win = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 960,
    minHeight: 640,
    backgroundColor: '#111111',
    title: 'PPTX Viewer',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.removeMenu();

  win.webContents.setZoomFactor(1);
  win.webContents.setVisualZoomLevelLimits(1, 1);
  win.webContents.on('before-input-event', (event, input) => {
    if (!input.control && !input.meta) {
      return;
    }

    if (['+', '=', '-', '_', '0'].includes(input.key)) {
      event.preventDefault();
      win.webContents.setZoomFactor(1);
    }
  });

  if (DEV_SERVER_URL) {
    win.loadURL(DEV_SERVER_URL);
    win.webContents.openDevTools({ mode: 'detach' });
    return win;
  }

  win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  return win;
}

async function readPresentation(filePath) {
  const bytes = await fs.readFile(filePath);

  return {
    name: path.basename(filePath),
    path: filePath,
    data: new Uint8Array(bytes),
  };
}

ipcMain.handle('presentation:open', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      {
        name: 'PowerPoint Presentations',
        extensions: ['pptx'],
      },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return readPresentation(result.filePaths[0]);
});

ipcMain.handle('presentation:read', async (_event, filePath) => {
  return readPresentation(filePath);
});

ipcMain.handle('presentation:get-launch-file', async () => {
  if (!launchPresentationPath) {
    return null;
  }

  return readPresentation(launchPresentationPath);
});

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
