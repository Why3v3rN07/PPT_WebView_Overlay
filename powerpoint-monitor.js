const { spawn } = require('child_process');
const path = require('path');

class PowerPointMonitor {
  constructor() {
    this.isMonitoring = false;
    this.checkInterval = null;
    this.currentSlideIndex = -1;
    this.onSlideChangeCallback = null;

    // Path to embedded Python
    this.pythonPath = path.join(__dirname, 'resources', 'python', 'python.exe');
    this.scriptPath = path.join(__dirname, 'powerpoint-monitor.py');
  }

  // Start monitoring PowerPoint
  start(onSlideChange) {
    if (this.isMonitoring) {
      console.log('Already monitoring');
      return;
    }

    this.onSlideChangeCallback = onSlideChange;
    this.isMonitoring = true;

    console.log('Starting PowerPoint monitor...');
    console.log('Python path:', this.pythonPath);
    console.log('Script path:', this.scriptPath);

    // Check PowerPoint state every 500ms
    this.checkInterval = setInterval(() => {
      this.checkPowerPoint();
    }, 500);
  }

  // Stop monitoring
  stop() {
    if (this.checkInterval) {
      clearInterval(this.checkInterval);
      this.checkInterval = null;
    }
    this.isMonitoring = false;
    this.currentSlideIndex = -1;
    console.log('Stopped PowerPoint monitor');
  }

  // Check PowerPoint state by calling Python script
  checkPowerPoint() {
    const python = spawn(this.pythonPath, [this.scriptPath]);

    let output = '';
    let errorOutput = '';

    python.stdout.on('data', (data) => {
      output += data.toString();
    });

    python.stderr.on('data', (data) => {
      errorOutput += data.toString();
    });

    python.on('close', (code) => {
      try {
        if (!this.isMonitoring) return;  // stopped while Python was in flight

        if (errorOutput) {
          console.error('Python error:', errorOutput);
        }

        if (!output.trim()) {
          return;
        }

        const state = JSON.parse(output);

        if (state.error) {
          console.error('PowerPoint error:', state.error);
          return;
        }

        if (state.inSlideshow) {
          console.log('PowerPoint is in slideshow mode - Slide', state.currentSlide);

          const slideIndex = state.currentSlide;

          // If slide changed, trigger callback
          if (slideIndex !== this.currentSlideIndex) {
            console.log(`Slide changed: ${this.currentSlideIndex} -> ${slideIndex}`);
            this.currentSlideIndex = slideIndex;

            if (this.onSlideChangeCallback) {
              this.onSlideChangeCallback(slideIndex, state);
            }
          }
        } else {
          console.log('PowerPoint is NOT in slideshow mode');

          // Not in slideshow - reset
          if (this.currentSlideIndex !== -1) {
            console.log('Exited slideshow mode');
            this.currentSlideIndex = -1;

            if (this.onSlideChangeCallback) {
              this.onSlideChangeCallback(-1, null);
            }
          }
        }
      } catch (error) {
        console.error('Error parsing Python output:', error.message);
        console.error('Output was:', output);
      }
    });
  }
}

module.exports = PowerPointMonitor;