import JSZip from 'jszip';
import Chart from 'chart.js/auto';
import './styles.css';

window.JSZip = JSZip;
window.Chart = Chart;

let PPTXViewerConstructor = null;
let nextTabId = 0;

const elements = {
  openButton: document.getElementById('open-button'),
  fileInput: document.getElementById('file-input'),
  fitButton: document.getElementById('fit-button'),
  fullscreenButton: document.getElementById('fullscreen-button'),
  prevButton: document.getElementById('prev-button'),
  nextButton: document.getElementById('next-button'),
  zoomInButton: document.getElementById('zoom-in-button'),
  zoomOutButton: document.getElementById('zoom-out-button'),
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
  tabBar: document.getElementById('tab-bar'),
};

const state = {
  tabs: [],
  activeTabId: null,
  loadingCount: 0,
};

function getActiveTab() {
  return state.tabs.find((tab) => tab.id === state.activeTabId) ?? null;
}

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

function renderTabs() {
  elements.tabBar.replaceChildren();

  for (const tab of state.tabs) {
    const chip = document.createElement('div');
    chip.className = `tab-chip${tab.id === state.activeTabId ? ' active' : ''}`;
    chip.dataset.tabId = String(tab.id);

    const switchButton = document.createElement('button');
    switchButton.type = 'button';
    switchButton.className = 'tab-chip-button';
    switchButton.dataset.tabAction = 'activate';
    switchButton.dataset.tabId = String(tab.id);

    const tabName = document.createElement('div');
    tabName.className = 'tab-chip-name';
    tabName.textContent = tab.fileName;

    switchButton.append(tabName);

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'tab-close-button';
    closeButton.dataset.tabAction = 'close';
    closeButton.dataset.tabId = String(tab.id);
    closeButton.setAttribute('aria-label', `Close ${tab.fileName}`);
    closeButton.textContent = 'x';

    chip.append(switchButton, closeButton);
    elements.tabBar.append(chip);
  }
}

function updateUi() {
  const activeTab = getActiveTab();
  const hasDeck = Boolean(activeTab?.viewer);

  renderTabs();
  elements.viewerShell.classList.toggle('empty', !hasDeck);
  elements.emptyState.classList.toggle('hidden', hasDeck);
  elements.canvasStage.classList.toggle('hidden', !hasDeck);
  elements.prevButton.disabled = !hasDeck || activeTab.currentSlide <= 0;
  elements.nextButton.disabled = !hasDeck || activeTab.currentSlide >= activeTab.slideCount - 1;
  elements.zoomInButton.disabled = !hasDeck;
  elements.zoomOutButton.disabled = !hasDeck;
  elements.fitButton.disabled = !hasDeck;

  elements.slideStatus.textContent = `Slide ${hasDeck ? activeTab.currentSlide + 1 : 0} / ${activeTab?.slideCount ?? 0}`;
  elements.slideStatusCompact.textContent = `${hasDeck ? activeTab.currentSlide + 1 : 0} / ${activeTab?.slideCount ?? 0}`;
  elements.zoomStatus.textContent = `Zoom ${Math.round((activeTab?.zoom ?? 1) * 100)}%`;
  const progress = hasDeck && activeTab.slideCount > 0 ? ((activeTab.currentSlide + 1) / activeTab.slideCount) * 100 : 0;
  elements.slideProgress.style.width = `${progress}%`;
}

async function getPPTXViewerConstructor() {
  if (PPTXViewerConstructor) {
    return PPTXViewerConstructor;
  }

  const module = await import('pptxviewjs');
  PPTXViewerConstructor = module.PPTXViewer;
  return PPTXViewerConstructor;
}

async function renderActiveSlide() {
  const activeTab = getActiveTab();
  if (!activeTab?.viewer || !activeTab.baseSlideWidth || !activeTab.baseSlideHeight) {
    return;
  }

  const availableWidth = Math.max(elements.canvasStage.clientWidth - 40, 320);
  const availableHeight = Math.max(elements.canvasStage.clientHeight - 40, 180);
  const fitScale = Math.min(
    availableWidth / activeTab.baseSlideWidth,
    availableHeight / activeTab.baseSlideHeight,
  );
  const displayWidth = Math.max(1, Math.round(activeTab.baseSlideWidth * fitScale * activeTab.zoom));
  const displayHeight = Math.max(1, Math.round(activeTab.baseSlideHeight * fitScale * activeTab.zoom));

  elements.canvas.style.width = `${displayWidth}px`;
  elements.canvas.style.height = `${displayHeight}px`;
  await activeTab.viewer.renderSlide(activeTab.currentSlide, elements.canvas);
}

function attachViewerEvents(viewer, tab) {
  viewer.on('loadComplete', ({ slideCount }) => {
    tab.slideCount = slideCount;
    tab.currentSlide = 0;
    tab.zoom = 1;
    updateUi();
  });
}

function beginLoading(message) {
  state.loadingCount += 1;
  showLoading(message);
}

function endLoading() {
  state.loadingCount = Math.max(0, state.loadingCount - 1);
  if (state.loadingCount === 0) {
    hideLoading();
  }
}

async function createTabFromFile(file) {
  const PPTXViewer = await getPPTXViewerConstructor();
  const viewer = new PPTXViewer({ canvas: elements.canvas });
  const tab = {
    id: nextTabId += 1,
    viewer,
    fileName: file.name,
    slideCount: 0,
    currentSlide: 0,
    zoom: 1,
    baseSlideWidth: 0,
    baseSlideHeight: 0,
  };

  attachViewerEvents(viewer, tab);
  await viewer.loadFile(file);
  await viewer.renderSlide(0, elements.canvas, { scale: 1 });
  tab.slideCount = viewer.getSlideCount();
  tab.baseSlideWidth = elements.canvas.width || 1;
  tab.baseSlideHeight = elements.canvas.height || 1;

  state.tabs.push(tab);
  state.activeTabId = tab.id;
  updateUi();
  await renderActiveSlide();
  elements.viewerShell.focus();
}

async function loadPresentationFiles(files) {
  if (!files.length) {
    return;
  }

  beginLoading(files.length > 1 ? 'Opening presentations...' : 'Opening presentation...');
  clearError();

  try {
    for (const file of files) {
      await createTabFromFile(file);
    }
  } catch (error) {
    showError(error instanceof Error ? error.message : 'Failed to open presentation.');
  } finally {
    endLoading();
    updateUi();
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

  await loadPresentationFiles([file]);
}

function openPresentation() {
  elements.fileInput.click();
}

async function activateTab(tabId) {
  if (tabId === state.activeTabId || !state.tabs.some((tab) => tab.id === tabId)) {
    return;
  }

  state.activeTabId = tabId;
  clearError();
  updateUi();
  await renderActiveSlide();
  elements.viewerShell.focus();
}

async function closeTab(tabId) {
  const closingIndex = state.tabs.findIndex((tab) => tab.id === tabId);
  if (closingIndex === -1) {
    return;
  }

  const [closingTab] = state.tabs.splice(closingIndex, 1);
  closingTab.viewer?.destroy();

  if (state.activeTabId === tabId) {
    const fallbackTab = state.tabs[Math.max(0, closingIndex - 1)] ?? state.tabs[0] ?? null;
    state.activeTabId = fallbackTab?.id ?? null;
  }

  clearError();
  updateUi();

  if (state.activeTabId) {
    await renderActiveSlide();
    elements.viewerShell.focus();
    return;
  }

  elements.canvas.width = 0;
  elements.canvas.height = 0;
}

async function goToPreviousSlide() {
  const activeTab = getActiveTab();
  if (!activeTab?.viewer || activeTab.currentSlide <= 0) {
    return;
  }

  activeTab.currentSlide = Math.max(0, activeTab.currentSlide - 1);
  await renderActiveSlide();
  updateUi();
}

async function goToNextSlide() {
  const activeTab = getActiveTab();
  if (!activeTab?.viewer || activeTab.currentSlide >= activeTab.slideCount - 1) {
    return;
  }

  activeTab.currentSlide = Math.min(activeTab.slideCount - 1, activeTab.currentSlide + 1);
  await renderActiveSlide();
  updateUi();
}

function applyZoom(nextZoom) {
  const activeTab = getActiveTab();
  if (!activeTab?.viewer) {
    return;
  }

  activeTab.zoom = Math.min(2.5, Math.max(0.4, nextZoom));
  updateUi();
  void renderActiveSlide();
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
  openPresentation();
});

elements.fileInput.addEventListener('change', (event) => {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.files?.length) {
    return;
  }

  void loadPresentationFiles(Array.from(input.files));
  input.value = '';
});

  elements.tabBar.addEventListener('click', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  const actionElement = target.closest('[data-tab-action]');
  if (!(actionElement instanceof HTMLElement)) {
    return;
  }

  const action = actionElement.dataset.tabAction;
  const tabId = Number(actionElement.dataset.tabId);
  if (!action || !Number.isFinite(tabId)) {
    return;
  }

  if (action === 'activate') {
    void activateTab(tabId);
    return;
  }

  if (action === 'close') {
    void closeTab(tabId);
  }
});

elements.prevButton.addEventListener('click', () => {
  void goToPreviousSlide();
});

elements.nextButton.addEventListener('click', () => {
  void goToNextSlide();
});

elements.zoomInButton.addEventListener('click', () => applyZoom((getActiveTab()?.zoom ?? 1) + 0.1));
elements.zoomOutButton.addEventListener('click', () => applyZoom((getActiveTab()?.zoom ?? 1) - 0.1));
elements.fitButton.addEventListener('click', fitSlide);
elements.fullscreenButton?.addEventListener('click', () => {
  void toggleFullscreen();
});

window.addEventListener('resize', () => {
  if (getActiveTab()?.viewer) {
    void renderActiveSlide();
  }
});

window.addEventListener('keydown', (event) => {
  if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
    return;
  }

  if (isZoomInShortcut(event)) {
    event.preventDefault();
    applyZoom((getActiveTab()?.zoom ?? 1) + 0.1);
    return;
  }

  if (isZoomOutShortcut(event)) {
    event.preventDefault();
    applyZoom((getActiveTab()?.zoom ?? 1) - 0.1);
    return;
  }

  if (isZoomResetShortcut(event)) {
    event.preventDefault();
    fitSlide();
    return;
  }

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'o') {
    event.preventDefault();
    openPresentation();
    return;
  }

  switch (event.key) {
    case 'ArrowLeft':
    case 'k':
    case 'K':
      event.preventDefault();
      void goToPreviousSlide();
      break;
    case 'ArrowRight':
    case 'j':
    case 'J':
    case ' ':
      event.preventDefault();
      void goToNextSlide();
      break;
    case '+':
    case '=':
      event.preventDefault();
      applyZoom((getActiveTab()?.zoom ?? 1) + 0.1);
      break;
    case '-':
    case '_':
      event.preventDefault();
      applyZoom((getActiveTab()?.zoom ?? 1) - 0.1);
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
