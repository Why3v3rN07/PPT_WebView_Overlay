"""
Run this while PowerPoint presenter view is active.
It will dump all visible top-level windows belonging to the PowerPoint process.
"""
import win32gui
import win32process
import json

def main():
    ppt_hwnd = win32gui.FindWindow('PPTFrameClass', None)
    if not ppt_hwnd:
        print(json.dumps({'error': 'PPTFrameClass window not found — is PowerPoint open?'}))
        return

    _, ppt_pid = win32process.GetWindowThreadProcessId(ppt_hwnd)
    print(f'PowerPoint PID: {ppt_pid}', flush=True)

    windows = []

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != ppt_pid:
                return
        except Exception:
            return

        cls   = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        rect  = win32gui.GetWindowRect(hwnd)
        w     = rect[2] - rect[0]
        h     = rect[3] - rect[1]

        windows.append({
            'hwnd':  hwnd,
            'class': cls,
            'title': title,
            'rect':  {'left': rect[0], 'top': rect[1], 'width': w, 'height': h},
        })

    win32gui.EnumWindows(enum_callback, None)

    for win in windows:
        print(json.dumps(win), flush=True)

if __name__ == '__main__':
    main()