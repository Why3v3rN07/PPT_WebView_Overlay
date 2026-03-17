const { app, BrowserWindow, Tray, Menu, ipcMain, globalShortcut, screen } = require('electron');
const path = require('path');
const fs   = require('fs');
const PowerPointMonitor = require('./powerpoint-monitor');

let mainWindow   = null;
let tray         = null;
let monitor      = new PowerPointMonitor();
let isMonitoring = false;

// ---------------------------------------------------------------------------
// Settings  (persisted to userData/settings.json)
// ---------------------------------------------------------------------------

const settingsPath = path.join(app.getPath('userData'), 'settings.json');

function loadSettings() {
  try { return JSON.parse(fs.readFileSync(settingsPath, 'utf8')); }
  catch { return { persistByDefault: false, interactiveByDefault: true }; }
}
function saveSettings() {
  fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
}

let settings = loadSettings();

ipcMain.handle('get-setting', (_, key)        => settings[key] ?? null);
ipcMain.handle('set-setting', (_, key, value) => { settings[key] = value; saveSettings(); });

// ---------------------------------------------------------------------------
// Overlay pools
//
// persistPool – Map<key, BrowserWindow>
//   Overlays that survive slide changes. Hidden when their slide isn't
//   showing, re-shown (without reload) when it is.
//   Key: "<url>|<left>|<top>|<width>|<height>"  (slide-point coords, stable)
//
// reloadWindows – BrowserWindow[]
//   Overlays that are destroyed and recreated on every slide change.
// ---------------------------------------------------------------------------

const persistPool  = new Map();   // key → BrowserWindow
let reloadWindows  = [];          // closed on each slide change

function persistKey(shape) {
  // Use raw slide-point coords as the key — they never change between visits.
  return `${shape.url}|${shape.left}|${shape.top}|${shape.width}|${shape.height}`;
}

// ---------------------------------------------------------------------------
// Coordinate conversion: PPT window pts → Electron logical pixels
//
// PPT COM gives window geometry in "points" whose numeric value equals
// (logical_pixels * primaryDPI / 96 / scaleFactor) — effectively opaque.
// We recover logical pixels empirically: try every Electron display, compute
// factor = display.bounds.width / win.widthPts, check the height error.
// The display with the smallest height error is the one the window is on,
// and its factor converts all four coordinates correctly.
// ---------------------------------------------------------------------------

function pptWindowToLogical(win) {
  const displays = screen.getAllDisplays();
  let best = { display: displays[0], factor: 96 / 72, error: Infinity };

  for (const d of displays) {
    const f     = d.bounds.width / win.widthPts;
    const error = Math.abs(win.heightPts * f - d.bounds.height) / d.bounds.height;
    if (error < best.error) best = { display: d, factor: f, error };
  }

  return {
    left:   win.leftPts   * best.factor,
    top:    win.topPts    * best.factor,
    width:  win.widthPts  * best.factor,
    height: win.heightPts * best.factor,
  };
}

// ---------------------------------------------------------------------------
// Overlay creation
// ---------------------------------------------------------------------------

function createOverlayWindow(x, y, width, height, url, interactive) {
  const win = new BrowserWindow({
    x, y, width, height,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    skipTaskbar: true,
    hasShadow: false,
    focusable: interactive,
    webPreferences: { nodeIntegration: false, contextIsolation: true },
  });

  if (!interactive) win.setIgnoreMouseEvents(true);

  win.loadURL(url);
  win.setAlwaysOnTop(true, 'screen-saver');
  console.log(`  Overlay (${x},${y}) ${width}x${height}  interactive=${interactive}  ${url}`);
  return win;
}

// ---------------------------------------------------------------------------
// Show / hide / close helpers
// ---------------------------------------------------------------------------

function closeReloadWindows() {
  reloadWindows.forEach(w => { if (!w.isDestroyed()) w.close(); });
  reloadWindows = [];
}

function hideAllPersisted() {
  persistPool.forEach(w => { if (!w.isDestroyed()) w.hide(); });
}

function closeAllOverlays() {
  closeReloadWindows();
  persistPool.forEach(w => { if (!w.isDestroyed()) w.close(); });
  persistPool.clear();
}

// ---------------------------------------------------------------------------
// Monitoring
// ---------------------------------------------------------------------------

function startMonitoring() {
  if (isMonitoring) return;
  isMonitoring = true;
  console.log('Starting PowerPoint monitoring…');
  monitor.start((slideIndex, state) => handleSlideChange(slideIndex, state));
}

function stopMonitoring() {
  if (!isMonitoring) return;
  isMonitoring = false;
  monitor.stop();
  closeAllOverlays();
}

// ---------------------------------------------------------------------------
// Core layout: place overlays for one slideshow window
// ---------------------------------------------------------------------------

function overlaysForWindow(win, slideSize, shapes) {
  const logical = pptWindowToLogical(win);

  // Letterbox: find the rendered slide rect inside the logical window.
  const slideAspect  = slideSize.width / slideSize.height;
  const windowAspect = logical.width / logical.height;

  let renderLeft = logical.left,  renderTop = logical.top;
  let renderW    = logical.width, renderH   = logical.height;

  if (Math.abs(slideAspect - windowAspect) > 0.001) {
    if (slideAspect > windowAspect) {
      renderH   = logical.width / slideAspect;
      renderTop = logical.top + (logical.height - renderH) / 2;
    } else {
      renderW    = logical.height * slideAspect;
      renderLeft = logical.left + (logical.width - renderW) / 2;
    }
  }

  const scaleX = renderW / slideSize.width;
  const scaleY = renderH / slideSize.height;

  console.log(`  logical (${logical.left.toFixed(0)},${logical.top.toFixed(0)}) ${logical.width.toFixed(0)}x${logical.height.toFixed(0)}`);
  console.log(`  render  (${renderLeft.toFixed(0)},${renderTop.toFixed(0)}) ${renderW.toFixed(0)}x${renderH.toFixed(0)}`);

  // Track which persist keys are active on this slide so we can hide the rest.
  const activeKeys = new Set();

  for (const shape of shapes) {
    const ox = Math.round(renderLeft + shape.left * scaleX);
    const oy = Math.round(renderTop  + shape.top  * scaleY);
    const ow = Math.round(shape.width  * scaleX);
    const oh = Math.round(shape.height * scaleY);

    // Resolve flags: per-shape overrides > global default
    const persist = shape.flagReload  ? false
                  : shape.flagPersist ? true
                  : (settings.persistByDefault ?? false);

    const interactive = shape.flagStatic      ? false
                      : shape.flagInteractive ? true
                      : (settings.interactiveByDefault ?? true);

    console.log(`  shape (${ox},${oy}) ${ow}x${oh}  persist=${persist}  interactive=${interactive}  ${shape.url}`);

    if (persist) {
      const key = persistKey(shape);
      activeKeys.add(key);
      const existing = persistPool.get(key);

      if (existing && !existing.isDestroyed()) {
        // Re-show without reloading; reposition in case DPI changed.
        existing.setBounds({ x: ox, y: oy, width: ow, height: oh });
        existing.show();
        existing.moveTop();
        console.log(`      restored from persist pool`);
      } else {
        // First visit: create and cache.
        const w = createOverlayWindow(ox, oy, ow, oh, shape.url, interactive);
        persistPool.set(key, w);
        w.on('closed', () => persistPool.delete(key));
      }
    } else {
      // Reload overlay: just create it; it will be closed on next slide change.
      reloadWindows.push(createOverlayWindow(ox, oy, ow, oh, shape.url, interactive));
    }
  }

  // Hide persisted overlays that belong to OTHER slides.
  persistPool.forEach((w, key) => {
    if (!activeKeys.has(key) && !w.isDestroyed()) w.hide();
  });
}

// ---------------------------------------------------------------------------
// Slide change handler
// ---------------------------------------------------------------------------

async function handleSlideChange(slideIndex, state) {
  // Always close reload overlays from the previous slide.
  closeReloadWindows();

  if (slideIndex === -1 || !state) { closeAllOverlays(); return; }
  if (state.isEndScreen)           { hideAllPersisted(); return; }

  if (!state.shapes || state.shapes.length === 0) {
    hideAllPersisted();
    console.log(`\n=== SLIDE ${slideIndex} — no [WEBVIEW] shapes ===`);
    return;
  }
  if (!state.windows || state.windows.length === 0) {
    console.error('No window info'); return;
  }

  console.log(`\n=== SLIDE ${slideIndex} ===`);

  for (let i = 0; i < state.windows.length; i++) {
    const label = state.windows[i].isPresenterView ? 'presenter' : 'audience';
    console.log(`  [${label}]`);
    overlaysForWindow(state.windows[i], state.slideSize, state.shapes);
  }

  console.log('===================\n');
}

// ---------------------------------------------------------------------------
// Window / tray
// ---------------------------------------------------------------------------

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 680, height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });
  mainWindow.loadFile('index.html');
  mainWindow.webContents.openDevTools();
  mainWindow.on('closed', () => { mainWindow = null; });
}

function createTray() {
  const contextMenu = Menu.buildFromTemplate([
    { label: 'Show Config',      click: () => { if (!mainWindow) createWindow(); else mainWindow.show(); } },
    { label: 'Start Monitoring', click: () => startMonitoring() },
    { label: 'Stop Monitoring',  click: () => stopMonitoring()  },
    { type: 'separator' },
    { label: 'Quit',             click: () => app.quit() }
  ]);
  // tray.setContextMenu(contextMenu);
}

// ---------------------------------------------------------------------------
// IPC + lifecycle
// ---------------------------------------------------------------------------

ipcMain.on('test-overlay',     (_, d) => reloadWindows.push(createOverlayWindow(d.x, d.y, d.width, d.height, d.url, true)));
ipcMain.on('close-overlays',   ()     => closeAllOverlays());
ipcMain.on('start-monitoring', ()     => startMonitoring());
ipcMain.on('stop-monitoring',  ()     => stopMonitoring());

app.whenReady().then(() => {
  createWindow();
  createTray();
  globalShortcut.register('CommandOrControl+Shift+Q', () => closeAllOverlays());

  screen.on('display-metrics-changed', (_e, _d, changed) => {
    if (changed.includes('scaleFactor') || changed.includes('bounds')) {
      console.log('Display metrics changed — closing overlays');
      closeAllOverlays();
    }
  });

  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on('window-all-closed', () => {});
app.on('will-quit',          () => globalShortcut.unregisterAll());
app.on('before-quit',        () => closeAllOverlays());