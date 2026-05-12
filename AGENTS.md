# AGENTS.md – ppt-webview Architecture & Development Guide

## Quick Context
**Windows-only Electron app** that overlays web content on PowerPoint presentations via shape annotations. Real-time monitoring via embedded Python helper using COM automation. Three major components: Electron main process, Python COM monitor, and web overlay renderer UIs.

---

## Architecture Overview

### Components & Data Flow
1. **Python Monitor** (`powerpoint-monitor.py`)
   - Spawned as subprocess by `powerpoint-monitor.js`
   - Polls PowerPoint via `win32com.client` every 0.25s, emits JSON lines to stdout
   - Parses shape alt-text, extracts overlay metadata (URL, flags, widget params)
   - Runs indefinitely; killed on `stop()`

2. **Main Process** (`main.js`)
   - Receives JSON from Python, detects slide changes (checks `currentSlide` value)
   - On slide change: hides/closes old overlays, schedules new overlay placement after transition delay (`TRANSITION_BUFFER_MS`)
   - Creates & manages two overlay pools:
     - **persistPool**: BrowserWindows that survive slide changes (hidden/shown, not reloaded)
     - **reloadWindows**: Windows destroyed on each slide change
   - Registers `widget://` custom protocol to serve local widget HTML files
   - Appends settings as query params to widget URLs before loading

3. **Renderer UI** (`index.html`)
   - Settings UI: toggles for overlay behavior, widget config (timezone, location, units)
   - IPC bridge via preload.js (`window.electronAPI.*`)
   - All settings saved to `%APPDATA%/ppt-webviewer/settings.json`

4. **Built-in Widgets** (`widgets/*.html`)
   - Standalone HTML pages that read `location.search` for config
   - No cross-origin requests needed; self-contained with Intl APIs
   - Completely transparent background for blending with slides

### Settings Persistence
- Saved to: `path.join(app.getPath('userData'), 'settings.json')`
- Keys: `persistByDefault`, `interactiveByDefault`, `widgetClockTz`, `widgetWeatherLoc`, `widgetWeatherUnits`, `widgetDateLoc`, `widgetDateMode`, `widgetClockShowDate`
- Loaded on app start; synced to main process via IPC

---

## Critical Patterns & Workflows

### Alt-Text Shape Annotation Format
Parsed by Python regex in `_parse_shape_alt_text()`:
```
[WEBVIEW]https://example.com [persist] [static]
[CLOCK tz=America/New_York] [persist]
[WEATHER loc=London] [interactive]
[DATE loc=Jerusalem mode=heb] [reload]
```
- Flags: `persist`, `reload`, `static`, `interactive` (booleans)
- Widget params: `tz=`, `loc=`, `mode=` (extracted to `widgetTz`, `widgetLoc`, `widgetMode`)
- Returns: JSON object with `url`, `flagPersist`, `flagReload`, `flagStatic`, `flagInteractive`, widget params

### Coordinate System Conversion (Complex)
**Input**: PPT window geometry in Office points (72 DPI) + slide content in slide-space points  
**Output**: Electron logical pixel coordinates matching actual screen position

Process (in `overlaysForWindow()`):
1. Match PPT window size to physical display to guess scaling factor (via `pptWindowToLogical()`)
2. Calculate letterbox rect (pillarbox/letterbox if slideshow window aspect ≠ slide aspect)
3. Scale shape coordinates from slide-space to render space
4. Round & position overlay windows

**Why this matters**: Multi-monitor setups, DPI scaling, and window mode vs presenter view all affect the math.

### Slide Change Transition Handling
1. Python detects new `currentSlide` value, emits state
2. `handleSlideChange()` immediately called with `slideIndex` & `state`
3. **Immediately** hide all persisted + close reload windows (avoid flicker over transition animation)
4. Schedule overlay placement after `(state.transitionDuration + TRANSITION_BUFFER_MS)` ms
5. **If another slide change fires before timer**: cancel pending placement, start fresh
   - Handles rapid advance/reverse without overlay ghosting

### Overlay Window Lifecycle

**Persistent overlays** (survive slide changes):
- Pool key: `"${url}|${left}|${top}|${width}|${height}"` (slide-point coords)
- If key exists & window not destroyed: reuse (move, show, top)
- If new key: create window, add to pool
- If key not in active set for current slide: hide window

**Reload overlays**:
- Array of windows destroyed at each slide change
- No pooling; always recreated

### Widget URL Resolution
`resolveWidgetUrl(shape)`:
- Base URL: `widget://clock`, `widget://weather`, `widget://date`
- Append query params from shape overrides + global settings
- Example: `widget://clock?tz=America/New_York&showDate=0`
- Renderer reads params via `new URLSearchParams(location.search)`

---

## File Inventory & Responsibilities

| File | Purpose |
|------|---------|
| `main.js` | Electron app entry, overlay management, settings, protocol handler, IPC |
| `powerpoint-monitor.js` | Spawns Python, parses JSON stream, detects slide changes |
| `powerpoint-monitor.py` | COM polling loop, shape parsing, state emission |
| `preload.js` | Exposes 6 IPC methods to renderer (`getSetting`, `setSetting`, `testOverlay`, etc.) |
| `index.html` | Settings UI; uses `window.electronAPI.*` |
| `widgets/clock.html`, `weather.html`, `date.html` | Self-contained widget renderers |
| `package.json` | npm metadata; single script: `start` = `electron .` |
| `resources/python/python.exe` | Embedded Python interpreter (must have `pywin32`) |

---

## Key Behaviors & Edge Cases

### When Nothing Works
- **No overlays appear**: Check `inSlideshow: false` in Python debug output → slideshow not running or COM error
- **Python import fails**: `resources/python/python.exe` missing or pywin32 not installed
- **Overlays misaligned**: Display scaling mismatch; multi-monitor setup may need manual adjustment
- **Widget not rendering**: Unknown tag name in alt-text ignored silently; check parsing in `_parse_shape_alt_text()`

### Display Scaling Edge Cases
- Presenter view (second window) detected via `isPresenterView: (i > 1)` in Python
- System DPI scaling found by matching window bounds to known display metrics
- Letterbox calculated separately for each window (supports dual-display with different content)

### Settings Defaults
- `persistByDefault`: false (recreate on each slide change)
- `interactiveByDefault`: true (mouse clicks to overlay)
- Shape flags override globals: `flagPersist` ⟹ persist, `flagReload` ⟹ reload, `flagStatic` ⟹ static, `flagInteractive` ⟹ interactive

### Python Process Management
- Spawned as: `spawn(pythonPath, [scriptPath])`
- Stdout: JSON lines (one per poll cycle, 0.25s interval)
- Stderr: logged as errors (e.g., COM failures)
- Exit: treated as loss of slideshow; calls callback with `slideIndex=-1`
- Killed on `stop()`: `process.kill()` (sends SIGTERM, caught by signal handler in Python)

---

## Common Tasks & How To

### Add a New Widget
1. Create `widgets/mywidget.html` (transparent background, read `location.search` for config)
2. Add entry to `WIDGET_URLS` dict in `powerpoint-monitor.py`
3. Handle in `resolveWidgetUrl()` in `main.js` to append default settings as query params
4. Add settings fields to `index.html` UI if needed
5. Document alt-text format in README & index.html instructions

### Modify Overlay Placement Logic
- Entry point: `overlaysForWindow(win, slideSize, shapes)` in `main.js`
- Shape geometry: `shape.left`, `shape.top`, `shape.width`, `shape.height` (slide-point coords)
- Window bounds: calculated by `pptWindowToLogical()` + letterbox math
- Each overlay created via `createOverlayWindow(x, y, w, h, url, interactive)`

### Debug PowerPoint State
- Python debug output printed to console; enable devtools with `mainWindow.webContents.openDevTools()`
- Each JSON line from Python: `console.log()` in `_handleLine()` shows parsed state
- Main process logs overlay placement with coordinates, URL, flags

### Handle New Alt-Text Parameters
1. Add case to `_KV_RE.findall()` loop in `powerpoint-monitor.py`
2. Store in shape dict (e.g., `shape['widgetX'] = val`)
3. Handle in `resolveWidgetUrl()` in `main.js` to append to query string
4. Read in widget HTML via `URLSearchParams(location.search).get('x')`

---

## Dependencies & Environment

- **Runtime**: Node.js (npm), Electron 42.0.0-alpha.5
- **Python**: Embedded at `resources/python/python.exe` with `pywin32` (for `win32com.client`)
- **OS**: Windows only (COM, paths, shortcuts)
- **No external APIs** (widgets use Intl, no network calls to load widget code itself)

---

## Known Issues & Constraints
1. **Multi-monitor scaling**: Heuristic matching can fail on exotic setups; manual adjustment may be needed.
2. **No test suite**: `npm test` is placeholder.
3. **Animation unsupported**: Shape animations don't trigger overlay moves; only alt-text updates on slide boundaries.
4. **Long transitions problematic**: Overlays from first slide stay visible during long exit transitions (timing detection happens only on slide change detection).

---

## Debugging Tips

- **Always check console**: `mainWindow.webContents.openDevTools()` is called by default on startup
- **Python output**: Stderr goes to console; stdout is parsed as JSON
- **Overlay test feature**: Use "Test Overlay" section in UI to verify window creation before relying on PowerPoint shapes
- **Keyboard shortcut**: Ctrl+Shift+Q closes all overlays immediately (useful for stuck windows)
- **Settings storage**: Edit `%APPDATA%/ppt-webviewer/settings.json` manually to reset or bulk-change config

