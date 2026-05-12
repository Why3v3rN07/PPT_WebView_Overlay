const { spawn } = require('child_process');
const path = require('path');

class PowerPointMonitor {
  constructor(basePath) {
    this.isMonitoring = false;
    this.currentSlideIndex = -1;
    this.onSlideChangeCallback = null;
    this.process = null;
    this.lineBuffer = '';

    // Path to embedded Python — resolved by main.js based on app.isPackaged
    this.basePath = basePath || path.join(__dirname, 'resources');
    this.pythonPath = path.join(this.basePath, 'python', 'python.exe');
    this.scriptPath = path.join(this.basePath, '..', 'powerpoint-monitor.py');
  }

  // Start monitoring PowerPoint
  start(onSlideChange) {
    if (this.isMonitoring) {
      console.log('Already monitoring');
      return;
    }

    this.onSlideChangeCallback = onSlideChange;
    this.isMonitoring = true;
    this.lineBuffer = '';

    console.log('Starting PowerPoint monitor...');
    console.log('Python path:', this.pythonPath);
    console.log('Script path:', this.scriptPath);

    this.process = spawn(this.pythonPath, [this.scriptPath], {
      cwd: path.join(this.basePath, 'python')
    });

    this.process.stdout.on('data', (data) => {
      // Buffer incoming data and process complete lines.
      this.lineBuffer += data.toString();
      const lines = this.lineBuffer.split('\n');
      // The last element is either empty or an incomplete line — keep it.
      this.lineBuffer = lines.pop();
      for (const line of lines) {
        this._handleLine(line.trim());
      }
    });

    this.process.stderr.on('data', (data) => {
      console.error('Python error:', data.toString());
    });

    this.process.on('close', (code) => {
      if (this.isMonitoring) {
        // Unexpected exit — log and treat as loss of slideshow.
        console.error(`Python process exited unexpectedly (code ${code})`);
        this.isMonitoring = false;
        this.currentSlideIndex = -1;
        if (this.onSlideChangeCallback) this.onSlideChangeCallback(-1, null);
      }
    });
  }

  // Stop monitoring
  stop() {
    this.isMonitoring = false;
    this.currentSlideIndex = -1;
    this.lineBuffer = '';

    if (this.process) {
      this.process.kill();
      this.process = null;
    }

    console.log('Stopped PowerPoint monitor');
  }

  // Handle one complete JSON line from the Python process.
  _handleLine(line) {
    if (!line || !this.isMonitoring) return;

    let state;
    try {
      state = JSON.parse(line);
    } catch (err) {
      console.error('Error parsing Python output:', err.message);
      console.error('Line was:', line);
      return;
    }

    if (state.error) {
      console.error('PowerPoint error:', state.error);
      return;
    }

    if (state.inSlideshow) {
      const slideIndex = state.currentSlide;
      if (slideIndex !== this.currentSlideIndex) {
        console.log(`Slide changed: ${this.currentSlideIndex} -> ${slideIndex}`);
        this.currentSlideIndex = slideIndex;
        if (this.onSlideChangeCallback) this.onSlideChangeCallback(slideIndex, state);
      }
    } else {
      if (this.currentSlideIndex !== -1) {
        console.log('Exited slideshow mode');
        this.currentSlideIndex = -1;
        if (this.onSlideChangeCallback) this.onSlideChangeCallback(-1, null);
      }
    }
  }
}

module.exports = PowerPointMonitor;