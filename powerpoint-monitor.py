import win32com.client
import json
import sys


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
# Alt-text flags ([persist], [reload], [static], [interactive]) are parsed
# here and emitted as individual booleans on each shape object.
# ---------------------------------------------------------------------------

def get_powerpoint_state():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        if ppt.SlideShowWindows.Count == 0:
            return {'inSlideshow': False, 'currentSlide': -1}

        slideshow     = ppt.SlideShowWindows(1)
        current_slide = slideshow.View.CurrentShowPosition
        presentation  = ppt.ActivePresentation
        total_slides  = presentation.Slides.Count
        slide_width   = float(presentation.PageSetup.SlideWidth)
        slide_height  = float(presentation.PageSetup.SlideHeight)

        if current_slide > total_slides:
            return {
                'inSlideshow':  True,
                'currentSlide': current_slide,
                'totalSlides':  total_slides,
                'isEndScreen':  True,
                'shapes':       []
            }

        slide = presentation.Slides(current_slide)

        windows = []
        for i in range(1, ppt.SlideShowWindows.Count + 1):
            sw = ppt.SlideShowWindows(i)
            windows.append({
                'leftPts':         float(sw.Left),
                'topPts':          float(sw.Top),
                'widthPts':        float(sw.Width),
                'heightPts':       float(sw.Height),
                'isPresenterView': (i > 1),
            })

        webview_shapes = []
        for shape in slide.Shapes:
            try:
                alt_text = shape.AlternativeText
                if not alt_text or not alt_text.startswith('[WEBVIEW]'):
                    continue
                tokens   = alt_text[9:].strip().split()
                url      = tokens[0] if tokens else ''
                if not url:
                    continue
                flags = set(t.strip('[]').lower() for t in tokens[1:] if t.strip('[]'))
                webview_shapes.append({
                    'url':         url,
                    'left':        float(shape.Left),
                    'top':         float(shape.Top),
                    'width':       float(shape.Width),
                    'height':      float(shape.Height),
                    'flagPersist': 'persist'     in flags,
                    'flagReload':  'reload'      in flags,
                    'flagStatic':  'static'      in flags,
                    'flagInteractive': 'interactive' in flags,
                })
            except Exception:
                pass

        return {
            'inSlideshow':  True,
            'currentSlide': current_slide,
            'totalSlides':  total_slides,
            'isEndScreen':  False,
            'slideSize':    {'width': slide_width, 'height': slide_height},
            'windows':      windows,
            'shapes':       webview_shapes,
        }

    except Exception as e:
        return {'error': str(e), 'inSlideshow': False}


if __name__ == '__main__':
    state = get_powerpoint_state()
    print(json.dumps(state))
    sys.stdout.flush()