const {app, BrowserWindow, Tray, Menu, ipcMain, globalShortcut, screen, protocol, net} = require('electron');
const path = require('path');
const fs = require('fs');
const PowerPointMonitor = require('./powerpoint-monitor');

let mainWindow = null;
let tray = null;
let monitor = new PowerPointMonitor();
let isMonitoring = false;

// ---------------------------------------------------------------------------
// Settings  (persisted to userData/settings.json)
// ---------------------------------------------------------------------------

const settingsPath = path.join(app.getPath('userData'), 'settings.json');

function loadSettings() {
    try {
        return JSON.parse(fs.readFileSync(settingsPath, 'utf8'));
    } catch {
        return {persistByDefault: false, interactiveByDefault: true};
    }
}

function saveSettings() {
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
}

let settings = loadSettings();

ipcMain.handle('get-setting', (_, key) => settings[key] ?? null);
ipcMain.handle('set-setting', (_, key, value) => {
    settings[key] = value;
    saveSettings();
});

// ---------------------------------------------------------------------------
// widget:// custom protocol
//
// Resolves widget://clock  → widgets/clock.html
//            widget://weather → widgets/weather.html
//            widget://date    → widgets/date.html
//
// Query params from the URL are forwarded as-is so the widget HTML can
// read them with new URLSearchParams(location.search).
// ---------------------------------------------------------------------------

app.whenReady().then(() => {
    protocol.handle('widget', (request) => {
        const url = new URL(request.url);
        const name = url.hostname;   // "clock", "weather", "date"
        const filePath = path.join(__dirname, 'widgets', `${name}.html`);
        // Serve the local file; query params are available to the page via location.search
        return net.fetch('file://' + filePath);
    });
});

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

const persistPool = new Map();   // key → BrowserWindow
let reloadWindows = [];          // closed on each slide change

// Pending transition timer — cancelled if the user advances before it fires.
let transitionTimer = null;

function persistKey(shape) {
    return `${shape.url}|${shape.left}|${shape.top}|${shape.width}|${shape.height}`;
}

// ---------------------------------------------------------------------------
// Coordinate conversion: PPT window pts → Electron logical pixels
// ---------------------------------------------------------------------------

function pptWindowToLogical(win) {
    const displays = screen.getAllDisplays();
    let best = {display: displays[0], factor: 96 / 72, error: Infinity};

    for (const d of displays) {
        const f = d.bounds.width / win.widthPts;
        const error = Math.abs(win.heightPts * f - d.bounds.height) / d.bounds.height;
        if (error < best.error) best = {display: d, factor: f, error};
    }

    return {
        left: win.leftPts * best.factor,
        top: win.topPts * best.factor,
        width: win.widthPts * best.factor,
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
        webPreferences: {nodeIntegration: false, contextIsolation: true},
    });

    if (!interactive) win.setIgnoreMouseEvents(true);

    win.loadURL(url);
    win.setAlwaysOnTop(true, 'screen-saver');
    // Suppress the Windows system context menu on frameless windows.
    // Without this, right-clicks are swallowed by the OS before reaching the page.
    win.hookWindowMessage?.(0x0313, () => { /* WM_CONTEXTMENU suppressed */ });
    win.on('system-context-menu', (event) => {
        event.preventDefault();
    });
    console.log(`  Overlay (${x},${y}) ${width}x${height}  interactive=${interactive}  ${url}`);
    return win;
}

// ---------------------------------------------------------------------------
// Widget URL resolution
//
// Appends settings and per-shape overrides as query params so the widget
// HTML can access them without needing IPC.
//
// widget://clock  → widget://clock?tz=America/New_York
// widget://weather → widget://weather?loc=London&units=metric
// widget://date   → widget://date?loc=Jerusalem
// ---------------------------------------------------------------------------

function resolveWidgetUrl(shape) {
    const base = shape.url;
    if (!base.startsWith('widget://')) return base;

    const params = new URLSearchParams();
    const name = base.replace('widget://', '');

    if (name === 'clock') {
        // Timezone: per-shape override → global default → system (omit param)
        const tz = shape.widgetTz || settings.widgetClockTz || '';
        if (tz) params.set('tz', tz);

        // showDate: '1' (default) or '0'
        const showDate = settings.widgetClockShowDate ?? '1';
        if (showDate === '0') params.set('showDate', '0');
    }

    if (name === 'weather') {
        // Location: per-shape → global → omit (widget falls back to IP geolocation)
        const loc = shape.widgetLoc || settings.widgetWeatherLoc || '';
        if (loc) params.set('loc', loc);
        const units = settings.widgetWeatherUnits || 'metric';
        params.set('units', units);
    }

    if (name === 'date') {
        // Date uses its own location setting (widgetDateLoc) separate from weather
        const loc = shape.widgetLoc || settings.widgetDateLoc || '';
        if (loc) params.set('loc', loc);
        // mode param from shape (e.g. [DATE mode=heb]) overrides global setting
        const mode = shape.widgetMode || settings.widgetDateMode || 'both';
        if (mode !== 'both') params.set('mode', mode);
    }

    const qs = params.toString();
    return qs ? `${base}?${qs}` : base;
}

// ---------------------------------------------------------------------------
// Show / hide / close helpers
// ---------------------------------------------------------------------------

function cancelTransitionTimer() {
    if (transitionTimer !== null) {
        clearTimeout(transitionTimer);
        transitionTimer = null;
    }
}

function closeReloadWindows() {
    reloadWindows.forEach(w => {
        if (!w.isDestroyed()) w.close();
    });
    reloadWindows = [];
}

function hideAllPersisted() {
    persistPool.forEach(w => {
        if (!w.isDestroyed()) w.hide();
    });
}

function closeAllOverlays() {
    cancelTransitionTimer();
    closeReloadWindows();
    persistPool.forEach(w => {
        if (!w.isDestroyed()) w.close();
    });
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
    const slideAspect = slideSize.width / slideSize.height;
    const windowAspect = logical.width / logical.height;

    let renderLeft = logical.left, renderTop = logical.top;
    let renderW = logical.width, renderH = logical.height;

    if (Math.abs(slideAspect - windowAspect) > 0.001) {
        if (slideAspect > windowAspect) {
            renderH = logical.width / slideAspect;
            renderTop = logical.top + (logical.height - renderH) / 2;
        } else {
            renderW = logical.height * slideAspect;
            renderLeft = logical.left + (logical.width - renderW) / 2;
        }
    }

    const scaleX = renderW / slideSize.width;
    const scaleY = renderH / slideSize.height;

    console.log(`  logical (${logical.left.toFixed(0)},${logical.top.toFixed(0)}) ${logical.width.toFixed(0)}x${logical.height.toFixed(0)}`);
    console.log(`  render  (${renderLeft.toFixed(0)},${renderTop.toFixed(0)}) ${renderW.toFixed(0)}x${renderH.toFixed(0)}`);

    const activeKeys = new Set();

    for (const shape of shapes) {
        const ox = Math.round(renderLeft + shape.left * scaleX);
        const oy = Math.round(renderTop + shape.top * scaleY);
        const ow = Math.round(shape.width * scaleX);
        const oh = Math.round(shape.height * scaleY);

        const persist = shape.flagReload ? false
            : shape.flagPersist ? true
                : (settings.persistByDefault ?? false);

        const interactive = shape.flagStatic ? false
            : shape.flagInteractive ? true
                : (settings.interactiveByDefault ?? true);

        // Resolve widget:// URLs to include settings as query params
        const url = resolveWidgetUrl(shape);

        console.log(`  shape (${ox},${oy}) ${ow}x${oh}  persist=${persist}  interactive=${interactive}  ${url}`);

        if (persist) {
            const key = persistKey(shape);
            activeKeys.add(key);
            const existing = persistPool.get(key);

            if (existing && !existing.isDestroyed()) {
                existing.setBounds({x: ox, y: oy, width: ow, height: oh});
                existing.show();
                existing.moveTop();
                console.log(`      restored from persist pool`);
            } else {
                const w = createOverlayWindow(ox, oy, ow, oh, url, interactive);
                persistPool.set(key, w);
                w.on('closed', () => persistPool.delete(key));
            }
        } else {
            reloadWindows.push(createOverlayWindow(ox, oy, ow, oh, url, interactive));
        }
    }

    // Hide persisted overlays that belong to other slides.
    persistPool.forEach((w, key) => {
        if (!activeKeys.has(key) && !w.isDestroyed()) w.hide();
    });
}

// ---------------------------------------------------------------------------
// Slide change handler
//
// Transition timing strategy:
//   1. A slide change is detected immediately when currentSlide increments.
//      At that moment the transition animation is just starting.
//   2. We immediately hide/close all current overlays so they don't sit on
//      top of the animation.
//   3. We wait transitionDuration seconds, then place the new overlays.
//   4. If another slide change fires before the timer expires, we cancel
//      the pending placement and start fresh — this handles rapid advances.
//
// A small TRANSITION_BUFFER_MS is added to ensure the animation has truly
// finished before overlays appear (accounts for COM polling jitter).
// ---------------------------------------------------------------------------

const TRANSITION_BUFFER_MS = 150;

async function handleSlideChange(slideIndex, state) {
    // Cancel any pending transition timer from a previous slide change.
    cancelTransitionTimer();

    // Immediately hide / close current overlays.
    closeReloadWindows();
    hideAllPersisted();

    if (slideIndex === -1 || !state) {
        closeAllOverlays();
        return;
    }
    if (state.isEndScreen) {
        return; /* already hidden above */
    }

    if (!state.shapes || state.shapes.length === 0) {
        console.log(`\n=== SLIDE ${slideIndex} — no overlay shapes ===`);
        return;
    }
    if (!state.windows || state.windows.length === 0) {
        console.error('No window info');
        return;
    }

    const delayMs = Math.round((state.transitionDuration || 0) * 1000) + TRANSITION_BUFFER_MS;
    console.log(`\n=== SLIDE ${slideIndex}  (placing overlays in ${delayMs}ms) ===`);

    transitionTimer = setTimeout(() => {
        transitionTimer = null;
        console.log(`  Placing overlays for slide ${slideIndex}`);

        for (let i = 0; i < state.windows.length; i++) {
            const label = state.windows[i].isPresenterView ? 'presenter' : 'audience';
            console.log(`  [${label}]`);
            overlaysForWindow(state.windows[i], state.slideSize, state.shapes);
        }

        console.log('===================\n');
    }, delayMs);
}

// ---------------------------------------------------------------------------
// Window / tray
// ---------------------------------------------------------------------------

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 680, height: 780,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });
    mainWindow.loadFile('index.html');
    mainWindow.webContents.openDevTools();
    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

function createTray() {
    const contextMenu = Menu.buildFromTemplate([
        {
            label: 'Show Config', click: () => {
                if (!mainWindow) createWindow(); else mainWindow.show();
            }
        },
        {label: 'Start Monitoring', click: () => startMonitoring()},
        {label: 'Stop Monitoring', click: () => stopMonitoring()},
        {type: 'separator'},
        {label: 'Quit', click: () => app.quit()}
    ]);
    // tray.setContextMenu(contextMenu);
}

// ---------------------------------------------------------------------------
// IPC + lifecycle
// ---------------------------------------------------------------------------

ipcMain.on('test-overlay', (_, d) => reloadWindows.push(createOverlayWindow(d.x, d.y, d.width, d.height, d.url, true)));
ipcMain.on('close-overlays', () => closeAllOverlays());
ipcMain.on('start-monitoring', () => startMonitoring());
ipcMain.on('stop-monitoring', () => stopMonitoring());

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

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', () => {
});
app.on('will-quit', () => globalShortcut.unregisterAll());
app.on('before-quit', () => closeAllOverlays());