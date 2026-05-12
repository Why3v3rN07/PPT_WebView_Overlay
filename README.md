# PowerPoint Web Viewer (Electron)

Overlay live web content on top of a running Microsoft PowerPoint slideshow. 

This Electron app watches the current slide in PowerPoint (via a small Python helper that uses COM automation) and 
creates transparent, always‑on‑top overlay windows exactly where you place shapes in your slide. Shapes are annotated 
via Alternative Text to specify what to render (a URL or a built‑in widget) and optional flags.


## Overview
- Desktop app for Windows built with Electron (Node.js + Chromium).
- Monitors PowerPoint in real time using an embedded Python interpreter and the `win32com` automation API.
- Renders overlay BrowserWindows positioned to match annotated shapes on the current slide.
- Supports built‑in widgets via a custom protocol:
  - `widget://clock`
  - `widget://weather`
  - `widget://date`
- Shapes are configured by writing special tags in the shape’s Alternative Text.


## Stack
- Language/runtime: JavaScript (Node.js), HTML/CSS (renderer)
- Framework: Electron
- Package manager: npm (package.json, package‑lock.json)
- Auxiliary: Python (embedded at `resources/python/python.exe`) with `pywin32` (for `win32com.client`)

Entry points:
- Main process: `main.js` (declared as `main` in package.json)
- Renderer UI: `index.html` (loaded by the main process)
- Preload: `preload.js` (exposes limited IPC helpers)
- PowerPoint monitor (helper): `powerpoint-monitor.py` (spawned by `powerpoint-monitor.js`)


## Requirements
- Windows (PowerShell paths, Win32 APIs, and PowerPoint COM are required)
- Microsoft PowerPoint installed and running your presentation
- Node.js (to run Electron via npm)
- Embedded Python at `resources/python/python.exe`
  - Must include the `pywin32` package so `import win32com.client` works

Note: The app expects the embedded Python interpreter at `resources/python/python.exe` (see `powerpoint-monitor.js`). 
If it’s missing or lacks `pywin32`, monitoring will not work.


## Installation
1. Install Node.js (LTS recommended).
2. From the project directory, install dependencies:
   - `npm install`
3. Ensure the embedded Python exists at `resources/python/python.exe` and can import `win32com.client`.


## Running
- Development run:
  - `npm start`

When running, open your PowerPoint slideshow (Slide Show mode). The app will monitor slide changes and add overlays 
according to annotated shapes.


## Shape Annotations (Alternative Text)
The Python helper parses specific patterns from a shape’s Alternative Text. Recognized forms (from `powerpoint-monitor.py`):

- `[WEBVIEW]<url> [flag] [flag] ...`
- `[CLOCK] [flag] ...`
- `[CLOCK tz=<tz>] [flag] ...`
- `[WEATHER] [flag] ...`
- `[WEATHER loc=<loc>] [flag] ...`
- `[DATE] [flag] ...`
- `[DATE loc=<loc>] [flag] ...`

Returned fields include: `url`, `flagPersist`, `flagReload`, `flagStatic`, `flagInteractive`, 
and widget params: `widgetTz`, `widgetLoc`, `widgetMode`.

Built‑in widgets are resolved by a custom Electron protocol to local HTML files in `widgets/`:
- `widget://clock` → `widgets/clock.html`
- `widget://weather` → `widgets/weather.html`
- `widget://date` → `widgets/date.html`

Flags (as parsed):
- `persist` — keep overlay window alive across slides (hidden/shown without reload)
- `reload` — recreate overlay on each slide change
- `static` — indicates static content (handled by app logic)
- `interactive` — make overlay accept input

Examples:
- Web content in a shape: `[WEBVIEW]https://example.com/news ticker persist`
- Clock widget with timezone: `[CLOCK tz=Europe/Berlin] persist`
- Weather widget with location: `[WEATHER loc=Seattle]`

Notes:
- Query parameters can be added to widget URLs (e.g., `widget://clock?tz=UTC`). The app forwards them as-is to the 
widget HTML (accessible via `location.search`).
- The app also estimates display scaling to convert PowerPoint window points to Electron logical pixels.


## Application Controls and Behavior
- The renderer UI (`index.html`) provides controls and settings. Settings are persisted to 
`%APPDATA%/ppt-webviewer/settings.json` with keys observed in code:
  - `persistByDefault`
  - `interactiveByDefault`
- The main process registers a `widget://` protocol handler on `app.whenReady()`.
- Overlays are created with `BrowserWindow` using `frame: false`, `transparent: true`, and `alwaysOnTop: true`.


## npm Scripts
From `package.json`:
- `npm start` — Start Electron (`electron .`)
- `npm test` — Placeholder (prints an error and exits)

No other npm scripts are defined.


## Environment Variables
No environment variables are used by the code (no `process.env` references found).


## Tests
There are no automated tests in this repository. The `npm test` script is a placeholder.


## Project Structure
Top‑level files and directories:
- `main.js` — Electron main process logic (window creation, monitoring, overlays, custom protocol, settings)
- `index.html` — Renderer UI
- `preload.js` — Exposes limited IPC to the renderer
- `powerpoint-monitor.js` — Spawns and reads JSON lines from the Python helper
- `powerpoint-monitor.py` — Queries PowerPoint via COM, parses shape alt text, emits state as JSON
- `widgets/` — Built‑in widgets (e.g., `clock.html`, `weather.html`, `date.html`)
- `resources/python/` — Expected location of embedded Python (`python.exe`)
- `package.json`, `package-lock.json` — npm metadata and lockfile
- `node_modules/` — Installed dependencies (includes `electron`)
- `LICENSE` — License file (see note below)
- `debug.js`, `debugging.py`, `notes.txt` — Auxiliary materials (not used by core flow)


## License
- The repository contains a `LICENSE` file with the GNU General Public License v3.0 text.


## Troubleshooting
- No overlays appear:
  - Ensure the slideshow is running (Slide Show mode). The helper reports `inSlideshow: false` otherwise.
  - Verify `resources/python/python.exe` exists and can import `win32com.client` (i.e., has `pywin32`).
  - Check the app console logs for Python stderr or JSON parse errors.
- Overlays misaligned:
  - Display scaling is inferred from the slideshow window size; multi‑monitor setups may need adjustment in slide/page 
  sizes and monitor scaling settings.
- Widget not loading:
  - Confirm the tag is recognized and spelled as shown above. Unknown tags are ignored by the parser.


## Current Date
Generated on: 2026-05-03 22:27 (local time)
