import JSZip from 'jszip';
import Chart from 'chart.js/auto';
import './styles.css';

window.JSZip = JSZip;
window.Chart = Chart;

let PPTXViewerConstructor = null;

const elements = {
  openButton: document.getElementById('open-button'),
  fileInput: document.getElementById('file-input'),
  fitButton: document.getElementById('fit-button'),
  fullscreenButton: document.getElementById('fullscreen-button'),
  prevButton: document.getElementById('prev-button'),
  nextButton: document.getElementById('next-button'),
  zoomInButton: document.getElementById('zoom-in-button'),
  zoomOutButton: document.getElementById('zoom-out-button'),
  fileName: document.getElementById('file-name'),
  slideStatus: document.getElementById('slide-status'),
  slideStatusCompact: document.getElementById('slide-status-compact'),
  slideProgress: document.getElementById('slide-progress'),
  zoomStatus: document.getElementById('zoom-status'),
  viewerShell: document.getElementById('viewer-shell'),
  emptyState: document.getElementById('empty-state'),
  loadingState: document.getElementById('loading-state'),
  errorState: document.getElementById('error-state'),
  canvasStage: document.getElementById('canvas-stage'),
  canvas: document.getElementById('slide-canvas'),
};

const state = {
  viewer: null,
  filePath: null,
  fileName: null,
  slideCount: 0,
  currentSlide: 0,
  zoom: 1,
  baseSlideWidth: 0,
  baseSlideHeight: 0,
};

function showLoading(message = 'Loading presentation...') {
  elements.loadingState.textContent = message;
  elements.loadingState.classList.remove('hidden');
  elements.errorState.classList.add('hidden');
}

function hideLoading() {
  elements.loadingState.classList.add('hidden');
}

function showError(message) {
  elements.errorState.textContent = message;
  elements.errorState.classList.remove('hidden');
}

function clearError() {
  elements.errorState.classList.add('hidden');
  elements.errorState.textContent = '';
}

function updateUi() {
  const hasDeck = Boolean(state.viewer);

  elements.viewerShell.classList.toggle('empty', !hasDeck);
  elements.emptyState.classList.toggle('hidden', hasDeck);
  elements.canvasStage.classList.toggle('hidden', !hasDeck);
  elements.prevButton.disabled = !hasDeck || state.currentSlide <= 0;
  elements.nextButton.disabled = !hasDeck || state.currentSlide >= state.slideCount - 1;
  elements.zoomInButton.disabled = !hasDeck;
  elements.zoomOutButton.disabled = !hasDeck;
  elements.fitButton.disabled = !hasDeck;

  elements.fileName.textContent = state.fileName ?? 'No presentation loaded';
  elements.slideStatus.textContent = `Slide ${hasDeck ? state.currentSlide + 1 : 0} / ${state.slideCount}`;
  elements.slideStatusCompact.textContent = `${hasDeck ? state.currentSlide + 1 : 0} / ${state.slideCount}`;
  elements.zoomStatus.textContent = `Zoom ${Math.round(state.zoom * 100)}%`;
  const progress = hasDeck && state.slideCount > 0 ? ((state.currentSlide + 1) / state.slideCount) * 100 : 0;
  elements.slideProgress.style.width = `${progress}%`;
}

function syncStateFromViewer() {
  if (!state.viewer) {
    state.slideCount = 0;
    state.currentSlide = 0;
    return;
  }

  state.slideCount = state.viewer.getSlideCount();
}

async function renderCurrentSlide() {
  if (!state.viewer || !state.baseSlideWidth || !state.baseSlideHeight) {
    return;
  }

  const availableWidth = Math.max(elements.canvasStage.clientWidth - 40, 320);
  const availableHeight = Math.max(elements.canvasStage.clientHeight - 40, 180);
  const fitScale = Math.min(
    availableWidth / state.baseSlideWidth,
    availableHeight / state.baseSlideHeight,
  );
  const displayWidth = Math.max(1, Math.round(state.baseSlideWidth * fitScale * state.zoom));
  const displayHeight = Math.max(1, Math.round(state.baseSlideHeight * fitScale * state.zoom));

  elements.canvas.style.width = `${displayWidth}px`;
  elements.canvas.style.height = `${displayHeight}px`;
  await state.viewer.renderSlide(state.currentSlide, elements.canvas);
}

async function getPPTXViewerConstructor() {
  if (PPTXViewerConstructor) {
    return PPTXViewerConstructor;
  }

  const module = await import('pptxviewjs');
  PPTXViewerConstructor = module.PPTXViewer;
  return PPTXViewerConstructor;
}

function attachViewerEvents(viewer) {
  viewer.on('loadComplete', ({ slideCount }) => {
    state.slideCount = slideCount;
    state.currentSlide = 0;
    state.zoom = 1;
    updateUi();
  });
}

async function loadPresentationFile(file) {
  if (!file) {
    return;
  }

  showLoading('Opening presentation...');
  clearError();

  try {
    if (state.viewer) {
      state.viewer.destroy();
    }

    const PPTXViewer = await getPPTXViewerConstructor();
    const viewer = new PPTXViewer({ canvas: elements.canvas });
    attachViewerEvents(viewer);

    await viewer.loadFile(file);
    state.currentSlide = 0;
    await viewer.renderSlide(0, elements.canvas, { scale: 1 });

    state.viewer = viewer;
    state.filePath = null;
    state.fileName = file.name;
    syncStateFromViewer();
    state.baseSlideWidth = elements.canvas.width || 1;
    state.baseSlideHeight = elements.canvas.height || 1;
    state.zoom = 1;

    updateUi();
    await renderCurrentSlide();
    hideLoading();
    elements.viewerShell.focus();
  } catch (error) {
    state.viewer = null;
    state.filePath = null;
    state.fileName = null;
    state.slideCount = 0;
    state.currentSlide = 0;
    state.baseSlideWidth = 0;
    state.baseSlideHeight = 0;
    hideLoading();
    updateUi();
    showError(error instanceof Error ? error.message : 'Failed to open presentation.');
  }
}

async function loadLaunchPresentation() {
  if (!window.pptxViewer?.getLaunchPresentation) {
    return;
  }

  const launchPresentation = await window.pptxViewer.getLaunchPresentation();
  if (!launchPresentation) {
    return;
  }

  const serializedBytes = launchPresentation.data;
  const fileBytes = serializedBytes instanceof Uint8Array
    ? serializedBytes
    : new Uint8Array(Object.values(serializedBytes));

  const file = new File([fileBytes], launchPresentation.name, {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  await loadPresentationFile(file);
}

async function openPresentation() {
  elements.fileInput.click();
}

async function goToPreviousSlide() {
  if (!state.viewer || state.currentSlide <= 0) {
    return;
  }

  state.currentSlide = Math.max(0, state.currentSlide - 1);
  await renderCurrentSlide();
  updateUi();
}

async function goToNextSlide() {
  if (!state.viewer || state.currentSlide >= state.slideCount - 1) {
    return;
  }

  state.currentSlide = Math.min(state.slideCount - 1, state.currentSlide + 1);
  await renderCurrentSlide();
  updateUi();
}

function applyZoom(nextZoom) {
  if (!state.viewer) {
    return;
  }

  state.zoom = Math.min(2.5, Math.max(0.4, nextZoom));
  updateUi();
  void renderCurrentSlide();
}

function fitSlide() {
  applyZoom(1);
}

function isZoomInShortcut(event) {
  return (event.ctrlKey || event.metaKey) && (event.key === '+' || event.key === '=' || event.key === 'Add');
}

function isZoomOutShortcut(event) {
  return (event.ctrlKey || event.metaKey) && (event.key === '-' || event.key === '_' || event.key === 'Subtract');
}

function isZoomResetShortcut(event) {
  return (event.ctrlKey || event.metaKey) && event.key === '0';
}

async function toggleFullscreen() {
  if (!document.fullscreenElement) {
    await document.documentElement.requestFullscreen();
    return;
  }

  await document.exitFullscreen();
}

elements.openButton.addEventListener('click', () => {
  void openPresentation();
});

elements.fileInput.addEventListener('change', (event) => {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.files?.[0]) {
    return;
  }

  void loadPresentationFile(input.files[0]);
  input.value = '';
});

elements.prevButton.addEventListener('click', () => {
  void goToPreviousSlide();
});

elements.nextButton.addEventListener('click', () => {
  void goToNextSlide();
});

elements.zoomInButton.addEventListener('click', () => applyZoom(state.zoom + 0.1));
elements.zoomOutButton.addEventListener('click', () => applyZoom(state.zoom - 0.1));
elements.fitButton.addEventListener('click', fitSlide);
elements.fullscreenButton?.addEventListener('click', () => {
  void toggleFullscreen();
});

window.addEventListener('resize', () => {
  if (state.viewer) {
    void renderCurrentSlide();
  }
});

window.addEventListener('keydown', (event) => {
  if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
    return;
  }

  if (isZoomInShortcut(event)) {
    event.preventDefault();
    applyZoom(state.zoom + 0.1);
    return;
  }

  if (isZoomOutShortcut(event)) {
    event.preventDefault();
    applyZoom(state.zoom - 0.1);
    return;
  }

  if (isZoomResetShortcut(event)) {
    event.preventDefault();
    fitSlide();
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
    event.preventDefault();
    void openPresentation();
    return;
  }

  switch (event.key) {
    case 'ArrowLeft':
      event.preventDefault();
      void goToPreviousSlide();
      break;
    case 'ArrowRight':
    case ' ':
      event.preventDefault();
      void goToNextSlide();
      break;
    case '+':
    case '=':
      event.preventDefault();
      applyZoom(state.zoom + 0.1);
      break;
    case '-':
    case '_':
      event.preventDefault();
      applyZoom(state.zoom - 0.1);
      break;
    case '0':
      event.preventDefault();
      fitSlide();
      break;
    case 'f':
    case 'F':
      event.preventDefault();
      void toggleFullscreen();
      break;
    default:
      break;
  }
});

updateUi();
void loadLaunchPresentation();
