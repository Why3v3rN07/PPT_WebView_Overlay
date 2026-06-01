import win32com.client
import json
import sys
import time
import signal
import re

# ---------------------------------------------------------------------------
# PowerPoint state
# ---------------------------------------------------------------------------
#
# Window geometry is emitted as raw Office points (leftPts/topPts/etc.).
# main.js converts to Electron logical pixels by matching the window size
# against the known logical dimensions of each display.
#
# Shape geometry (left/top/width/height) is in slide-space points and is
# used as-is by main.js to compute overlay positions within the slide.
#
# Recognised alt-text prefixes:
#   [WEBVIEW]<url> [flags...]        — embed any URL
#   [CLOCK] [flags...]               — built-in clock widget
#   [CLOCK tz=<tz>] [flags...]       — clock with explicit timezone
#   [WEATHER] [flags...]             — built-in weather widget
#   [WEATHER loc=<loc>] [flags...]
#   [DATE] [flags...]                — built-in Hebrew/Gregorian date widget
#   [DATE loc=<loc>] [flags...]
#
# Alt-text flags ([persist], [reload], [static], [interactive]) are parsed
# here and emitted as individual booleans on each shape object.
#
# transitionDuration: seconds for the *incoming* slide's transition.
# main.js uses this to delay showing new overlays until after the
# transition completes, avoiding the "old overlay over new slide" flicker.
#
# Persistent-process mode: this script is spawned once and loops forever,
# emitting one JSON line per iteration. main.js kills it on stop().
# ---------------------------------------------------------------------------

INTERVAL = 0.25  # seconds between checks

# Matches key=value pairs inside a tag or as standalone tokens.
_KV_RE = re.compile(r'(\w+)=(\S+)')

# Sentinel URLs for built-in widgets — resolved by main.js via custom protocol.
WIDGET_URLS = {
    'CLOCK': 'widget://clock',
    'WEATHER': 'widget://weather',
    'DATE': 'widget://date',
}


def _parse_shape_alt_text(alt_text):
    """
    Parse an alt-text string into a shape dict or None.

    Recognised formats:
      [WEBVIEW]<url> [flag] [flag] ...
      [CLOCK] [flag] ...
      [CLOCK tz=<tz>] [flag] ...
      [WEATHER] [flag] ...
      [WEATHER loc=<loc>] [flag] ...
      [DATE] [flag] ...
      [DATE loc=<loc>] [flag] ...

    Returns a dict with url, flagPersist, flagReload, flagStatic,
    flagInteractive, widgetTz, widgetLoc — or None if not recognised.
    """
    if not alt_text:
        return None

    url = None
    widget_tz = None
    widget_loc = None
    widget_mode = None

    # --- [WEBVIEW]<url> ---
    if alt_text.startswith('[WEBVIEW]'):
        rest = alt_text[9:].strip()
        tokens = rest.split()
        url = tokens[0] if tokens else ''
        if not url:
            return None
        flag_tokens = tokens[1:]

    # --- [CLOCK ...], [WEATHER ...], [DATE ...] ---
    else:
        m = re.match(r'^\[(\w+)([^\]]*)\](.*)', alt_text)
        if not m:
            return None
        tag_name = m.group(1).upper()
        tag_params = m.group(2).strip()
        rest = m.group(3).strip()

        if tag_name not in WIDGET_URLS:
            return None

        url = WIDGET_URLS[tag_name]

        for key, val in _KV_RE.findall(tag_params + ' ' + rest):
            if key.lower() == 'tz':
                widget_tz = val
            elif key.lower() == 'loc':
                widget_loc = val
            elif key.lower() == 'mode':
                widget_mode = val

        flag_tokens = [t for t in rest.split() if '=' not in t]

    flags = set(t.strip('[]').lower() for t in flag_tokens if t.strip('[]'))

    return {
        'url': url,
        'flagPersist': 'persist' in flags,
        'flagReload': 'reload' in flags,
        'flagStatic': 'static' in flags,
        'flagInteractive': 'interactive' in flags,
        'widgetTz': widget_tz,
        'widgetLoc': widget_loc,
        'widgetMode': widget_mode,
    }


def get_powerpoint_state():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        if ppt.SlideShowWindows.Count == 0:
            return {'inSlideshow': False, 'currentSlide': -1}

        slideshow = ppt.SlideShowWindows(1)
        current_slide = slideshow.View.CurrentShowPosition
        presentation = ppt.ActivePresentation
        total_slides = presentation.Slides.Count
        slide_width = float(presentation.PageSetup.SlideWidth)
        slide_height = float(presentation.PageSetup.SlideHeight)

        if current_slide > total_slides:
            return {
                'inSlideshow': True,
                'currentSlide': current_slide,
                'totalSlides': total_slides,
                'isEndScreen': True,
                'shapes': []
            }

        slide = presentation.Slides(current_slide)

        # Transition duration for the incoming slide (seconds).
        # Used by main.js to delay overlay placement until after the animation.
        try:
            transition_duration = float(slide.SlideShowTransition.Duration)
        except Exception:
            transition_duration = 0.0

        windows = []
        for i in range(1, ppt.SlideShowWindows.Count + 1):
            sw = ppt.SlideShowWindows(i)
            windows.append({
                'leftPts': float(sw.Left),
                'topPts': float(sw.Top),
                'widthPts': float(sw.Width),
                'heightPts': float(sw.Height),
                'isPresenterView': (i > 1),
            })

        webview_shapes = []
        for shape in slide.Shapes:
            try:
                alt_text = shape.AlternativeText
                parsed = _parse_shape_alt_text(alt_text)
                if not parsed:
                    continue
                parsed['left'] = float(shape.Left)
                parsed['top'] = float(shape.Top)
                parsed['width'] = float(shape.Width)
                parsed['height'] = float(shape.Height)
                webview_shapes.append(parsed)
            except Exception:
                pass

        return {
            'inSlideshow': True,
            'currentSlide': current_slide,
            'totalSlides': total_slides,
            'isEndScreen': False,
            'slideSize': {'width': slide_width, 'height': slide_height},
            'transitionDuration': transition_duration,
            'windows': windows,
            'shapes': webview_shapes,
        }

    except Exception as e:
        return {'error': str(e), 'inSlideshow': False}


def emit(state):
    print(json.dumps(state), flush=True)


def main():
    signal.signal(signal.SIGTERM, lambda *_: sys.exit(0))
    while True:
        emit(get_powerpoint_state())
        time.sleep(INTERVAL)


if __name__ == '__main__':
    main()
