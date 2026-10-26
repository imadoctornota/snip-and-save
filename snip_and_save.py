#!/usr/bin/env python3
"""
Snip & Save - a simple screenshot / snipping tool with in-place annotation.

Features
--------
- Choose (and remember) the folder screenshots are saved to.
- Set a base filename; auto-numbering avoids overwriting old captures, saved
  as "N.name.ext" (e.g. 1.screenshot.png, 2.screenshot.png, ...).
  - If you type a number prefix yourself (e.g. "23.screenshot"), numbering
    picks up from there instead of starting over at 1.
  - A small panel lists existing PNG/JPG files that match the current name,
    and shows the next number that will be used.
- HDR-aware on Windows: asks the display config whether the monitor being
  captured has HDR enabled (and what its SDR reference white is), and tone
  maps the capture back to SDR so it isn't blown out. Auto / On / Off, and
  the correction is non-destructive - the raw capture is kept, so toggling
  re-renders instantly without re-capturing.
- "Import Slides..." (bottom right) turns every slide of a .pptx/.ppt (or
  .odp) deck into an image, saved through the same folder / file name /
  auto-numbering / format settings as screenshots. Hidden slides are
  skipped, and it runs in the background so the window stays usable.
  Afterwards the whole deck can be browsed in the preview: Prev / Next
  buttons, the left/right arrow keys (or PageUp/PageDown, Home/End), or the
  mouse wheel. Annotations are kept per slide, and Save writes the slide on
  screen (with its drawing) as a new numbered file. "Done" leaves browsing.
- Default output format: PNG or JPG.
- Capture delay: 0 / 3 / 5 / 10 seconds (countdown happens before the shot).
- Multi-monitor aware: pick "All Screens", a specific monitor, or
  "Other Display (not this window)" for both Full Screen and Region
  capture. That last option automatically figures out which monitor this
  program's window is currently sitting on and captures whichever other
  screen(s) are connected - handy for grabbing a shot of the other
  display without having to remember which one is "Screen 1" vs
  "Screen 2".
- Full-screen capture, or drag-to-select a region (classic "snipping tool"
  mode) with a thick red selection outline. Region selection is
  scale-corrected so it lines up with the actual screenshot on HiDPI/scaled
  displays.
- "Capture Region" sits directly above "Save Screenshot" (capture on top,
  save underneath, "Capture Full Screen" above "Manual Save..."), so the
  two back-to-back clicks share one column; View > Put "Capture Region" and
  "Save" on the left (in the menu bar) moves the pair to the left or right
  column.
- Every capture is automatically copied to the system clipboard right after
  it's taken, and Save re-copies the final version (with any annotations
  baked in) so the clipboard always matches what ends up on disk. The copy
  runs in the background, so the window never waits on it.
- "Copy to Clipboard" (bottom left), Edit > Copy to Clipboard, or Ctrl+C
  copies whatever is in the preview again - drawing included - any time,
  e.g. after something else has replaced the clipboard. (With text selected
  in a field, Ctrl+C copies that text as usual.)
- Options > "Print Screen key opens Snip & Save" (Windows): pressing Print
  Screen anywhere starts a Snip & Save capture instead of Windows' Snipping
  Tool - a region capture by default, or the full screen (also under
  Options). It captures straight away, ignoring Delay, and brings the window
  up with the result; press Esc while the window was minimised and it just
  stays minimised. While the option is on, Windows' own "Use the Print screen
  key to open screen capture" setting is switched off; turning the option off
  puts that setting back the way it was. Snip & Save has to be running for
  this (minimised is fine) - while it's closed, Print Screen does Windows'
  plain copy-the-screen-to-the-clipboard. If another program already holds
  the key, the app says so and keeps retrying in the background, taking the
  key as soon as it's free. Needs no extra installs.
- Fast capture: no timers between the click and the selection overlay. On
  Windows the window's hide/show animation is switched off and the capture
  waits exactly one screen refresh for the window to disappear; with "mss"
  installed only the needed screen area is grabbed. The status line shows
  how long the click-to-overlay took (SHOW_TIMINGS near the top turns it off).
- The captured screenshot loads directly into the main window's preview
  panel, where you can draw freehand lines of variable thickness and color
  right on it. Hit Save whenever you like, keep drawing more, and Save again
  - nothing closes or resets in between.
- Drawing tools above the preview: Pen (freehand), Rectangle (drag to box
  something) and Arrow (drag from the tail to what it points at). All three
  use the same Thickness and Color, Undo / Clear all cover them all, and the
  arrowhead grows with the thickness. Hold Shift while dragging for a perfect
  square or an arrow snapped to 45 degrees. Keyboard: P / R / A switch tools
  and Ctrl+Z undoes (both ignored while you're typing in a field). The chosen
  tool is remembered.
- Quick-capture button: a small round scissors button that floats above every
  other window (top centre of the main screen at first), so a capture is one
  click away without opening the app. Click it to capture a region straight
  away (Delay is ignored, as with Print Screen) and the window comes up with
  the result; drag it anywhere and the spot is remembered; right-click it for
  Capture Full Screen, Open Snip & Save, or Hide this button. It stays put
  while the main window is minimised and hides itself during every capture so
  it never ends up in the shot. View > Show quick-capture button turns it on
  or off, and View > Move quick-capture button back to the top resets it.
  (Round on Windows; a small square tile elsewhere.)
- Window size, position and all settings persist between runs (saved next
  to this script - or next to the exe - as snipper_config.json, falling back
  to your user app-data folder if that location isn't writable). The window
  geometry is saved shortly after every resize/move, not only on close.

Requirements
------------
    pip install pillow

Optional, for choosing a specific monitor (or "Other Display") in a
multi-monitor setup:
    pip install screeninfo
(Without it, only "All Screens" is available, which still spans every
connected display for full-screen capture. "Other Display" also needs
2+ monitors to be detected.)

Optional, for faster captures on Windows (recommended):
    pip install mss
(Grabs only the screen area being captured, straight into memory, instead of
the whole desktop. Without it the app uses Pillow's grab, as before. If you
build the exe with build_snip.py, add "mss" to its PACKAGES list.)

Optional, for clipboard copy support:
    Windows : pip install pywin32
    macOS   : nothing extra needed (uses the built-in osascript)
    Linux   : install xclip or wl-copy, e.g. sudo apt install xclip

On Linux, PIL's screen grab also needs one of these as a fallback on some
setups:
    sudo apt install scrot
    sudo apt install gnome-screenshot

Optional, for "Import Slides..." (either one is enough):
    Windows with PowerPoint installed : pip install pywin32   (best fidelity;
                                        same package as clipboard support)
    No PowerPoint (any OS)            : install LibreOffice (free):
                                        https://www.libreoffice.org/download/
    LibreOffice on its own is enough - no extra pip packages are needed. If
    PyMuPDF (pip install pymupdf) or Poppler's pdftoppm happens to be
    available, the app renders through a PDF instead; otherwise LibreOffice
    exports each slide itself, using the Python that ships with it.
On Windows the app uses PowerPoint when it's installed and switches to
LibreOffice automatically when it isn't (or if PowerPoint fails). LibreOffice
is looked for on PATH, in the usual install folders and in the Windows
registry; if it's somewhere unusual, the app asks you to point to it once and
remembers the location.

Run it with:
    python snip_and_save.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageGrab, ImageDraw, ImageTk
import os
import io
import re
import json
import time
import tempfile
import threading
import subprocess
import math
import platform
import sys
import shutil
import pathlib
import glob
import webbrowser
import queue
import importlib

CONFIG_NAME = "snipper_config.json"

# Append "overlay in N ms" etc. to the status line after each capture, so
# you can see what a capture actually costs on your machine.
SHOW_TIMINGS = True

IS_WINDOWS = platform.system() == "Windows"

# After hiding the window, how long to let the desktop redraw before grabbing
# it. On Windows the app instead waits for exactly one screen refresh
# (DwmFlush), which is both quicker and certain, so this only applies
# elsewhere.
HIDE_SETTLE_MS = 0 if IS_WINDOWS else 40


def _dir_is_writable(path):
    """Really try to create a file there - os.access() lies on Windows."""
    try:
        with tempfile.TemporaryFile(dir=path):
            pass
        return True
    except Exception:
        return False


def _pick_config_path():
    """Where settings live.

    Normally next to the script, or next to the exe when built with
    PyInstaller (a one-file exe's __file__ is a temp folder wiped on every
    launch, so it can't be used). If that folder isn't writable (e.g. the exe
    sits in Program Files), fall back to the per-user app-data folder so
    settings still stick instead of silently vanishing."""
    app_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    beside = os.path.join(app_dir, CONFIG_NAME)
    if _dir_is_writable(app_dir):
        return beside
    system = platform.system()
    if system == "Windows":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif system == "Darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "SnipAndSave")
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        return beside
    return os.path.join(folder, CONFIG_NAME)


CONFIG_FILE = _pick_config_path()

DEFAULT_CONFIG = {
    "save_dir": os.path.join(os.path.expanduser("~"), "Pictures", "Screenshots"),
    "filename": "screenshot",
    "format": "PNG",
    "delay": 0,
    "auto_increment": True,
    "window_geometry": "600x900",
    "region_on_left": True,
    "screen_choice": "All Screens",
    "hdr_mode": "Auto",          # Auto / On / Off
    "hdr_curve": "Washed out",   # "Washed out" (PQ) or "Too bright" (linear scRGB)
    "hdr_strength": 100,         # 0-100%, how much of the correction to apply
    "slides_dir": "",            # last folder a presentation was imported from
    "soffice_path": "",          # LibreOffice's soffice, if you had to point to it by hand
    "printscreen_hotkey": False, # Print Screen starts a Snip & Save capture (Windows)
    "printscreen_mode": "region",  # what Print Screen captures: "region" or "full"
    "printscreen_restore": "",   # Windows' own Print Screen setting before we changed it ("on"/"unset")
    "draw_tool": "pen",          # editor tool: "pen", "rect" or "arrow"
    "quick_button": True,        # show the small always-on-top capture button
    "quick_button_pos": "",      # "x,y" of that button; "" = top centre of the main screen
}

DRAW_TOOLS = ("pen", "rect", "arrow")

LIBREOFFICE_DOWNLOAD_URL = "https://www.libreoffice.org/download/"

# Pixel width each slide is rendered at; height follows the deck's aspect
# ratio (1920 -> 1920x1080 for a 16:9 deck). Raise it for sharper exports of
# image-heavy slides.
SLIDE_EXPORT_WIDTH = 1920

SLIDE_FILETYPES = [
    ("Presentations", "*.pptx *.ppt *.pptm *.ppsx *.pps *.potx *.odp"),
    ("All files", "*.*"),
]

# Special "SCREEN" choice meaning "whichever monitor this window is NOT on
# right now" - handy for grabbing a shot of the other display without
# having to know/remember which one is "Screen 1" vs "Screen 2". Only
# offered when 2+ monitors are detected (requires the "screeninfo" package).
OTHER_DISPLAY_LABEL = "Other Display (not this window)"

# ---------------------------------------------------------------- palette
# Higher-contrast light theme. Text/background pairs are chosen to clear
# WCAG AA (4.5:1) so labels stay crisp on any monitor.
BG = "#e9ecf3"            # window background
CARD = "#ffffff"          # card surfaces
SUNKEN = "#dfe3ec"        # preview well behind the screenshot
TEXT = "#141722"          # primary text
MUTED = "#454b5e"         # secondary text (was #767b8c - too faint)
BORDER = "#b9bfce"        # card + field outlines (was #e3e5f0 - nearly invisible)
BORDER_STRONG = "#8a91a6" # field outlines on hover/idle inputs

ACCENT = "#3730d9"        # Capture Region (primary action)
ACCENT_HOVER = "#2821b0"
SAVE = "#0b7a42"          # Save Screenshot
SAVE_HOVER = "#085f33"
SLATE = "#2b3040"         # Capture Full Screen
SLATE_HOVER = "#1b1f2b"
DANGER = "#c2262d"
DANGER_TINT = "#fbe4e5"
GHOST_HOVER = "#e6e9f2"

FONT_FAMILY = "Segoe UI"  # Tk falls back gracefully on platforms without it


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
            merged = DEFAULT_CONFIG.copy()
            merged.update(cfg)
            return merged
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    """Atomic write (temp file, then replace) so being closed or killed
    mid-save can never leave a half-written file that resets everything.
    Returns None on success, or the error text on failure."""
    tmp_path = CONFIG_FILE + ".tmp"
    try:
        with open(tmp_path, "w") as f:
            json.dump(cfg, f, indent=2)
        os.replace(tmp_path, CONFIG_FILE)
        return None
    except Exception as e:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        return str(e)


def parse_filename(raw):
    """Split a leading number prefix ("23.myshot" -> (23, "myshot")).
    No prefix -> (None, raw)."""
    raw = raw.strip()
    m = re.match(r'^(\d+)\.(.+)$', raw)
    if m:
        return int(m.group(1)), m.group(2)
    return None, raw


def get_monitors():
    """List connected monitors as dicts with label/x/y/width/height.
    Returns [] if the optional 'screeninfo' package isn't installed."""
    try:
        import screeninfo
        mons = screeninfo.get_monitors()
        result = []
        for i, m in enumerate(mons):
            is_primary = bool(getattr(m, "is_primary", False))
            label = f"Screen {i + 1}" + (" (Primary)" if is_primary else "")
            result.append({
                "label": label, "x": m.x, "y": m.y, "width": m.width, "height": m.height,
                "name": getattr(m, "name", None),  # e.g. "\\\\.\\DISPLAY1" on Windows
            })
        return result
    except Exception:
        return []


def monitor_index_containing(monitors, x, y):
    """Index of the monitor whose bounds contain point (x, y). If no
    monitor contains it exactly (e.g. the point is slightly off due to
    window-manager quirks), fall back to whichever monitor's center is
    closest. Returns None if `monitors` is empty."""
    if not monitors:
        return None
    for i, m in enumerate(monitors):
        if m["x"] <= x < m["x"] + m["width"] and m["y"] <= y < m["y"] + m["height"]:
            return i
    best_i, best_d = None, None
    for i, m in enumerate(monitors):
        cx = m["x"] + m["width"] / 2
        cy = m["y"] + m["height"] / 2
        d = (cx - x) ** 2 + (cy - y) ** 2
        if best_d is None or d < best_d:
            best_d = d
            best_i = i
    return best_i


GEOMETRY_RE = re.compile(r"^(\d+)x(\d+)(?:([+-]-?\d+)([+-]-?\d+))?$")
DEFAULT_GEOMETRY = "600x900"


def _geometry_offset(token):
    """Tk writes a negative offset as "+-1500" (and a plain one as "+100"),
    so drop the leading sign character and read the rest as a signed int."""
    number = int(token[1:])
    if token[0] == "-" and not token[1:].startswith("-"):
        number = -number
    return number


def usable_geometry(saved, monitors, screen_w, screen_h):
    """Validate a saved "WxH" or "WxH+X+Y" string before applying it.

    Returns a geometry string that is safe to use: the size always survives
    (unless it's absurdly small), and the position is kept only if the window's
    title bar would land on a connected monitor - so unplugging a screen can't
    strand the window off-screen. Falls back to DEFAULT_GEOMETRY."""
    m = GEOMETRY_RE.match(str(saved or "").strip())
    if not m:
        return DEFAULT_GEOMETRY
    w, h = int(m.group(1)), int(m.group(2))
    if w < 300 or h < 300:
        return DEFAULT_GEOMETRY
    if m.group(3) is None:
        return f"{w}x{h}"
    x, y = _geometry_offset(m.group(3)), _geometry_offset(m.group(4))
    if x <= -30000 or y <= -30000:          # Windows parks minimised windows here
        return f"{w}x{h}"
    probe_x, probe_y = x + min(w, 120) // 2, y + 12   # a point on the title bar
    if monitors:
        on_screen = any(
            mo["x"] <= probe_x < mo["x"] + mo["width"] and mo["y"] <= probe_y < mo["y"] + mo["height"]
            for mo in monitors
        )
    else:
        on_screen = -screen_w < probe_x < 2 * screen_w and -screen_h < probe_y < 2 * screen_h
    return f"{w}x{h}+{x}+{y}" if on_screen else f"{w}x{h}"


# ------------------------------------------------------------ slides -> images
def write_image(img, path, fmt):
    """Save img as PNG or JPG (quality 92, same as screenshots)."""
    if fmt == "JPG":
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.save(path, "JPEG", quality=92)
    else:
        img.save(path, "PNG")


def _no_console_kwargs():
    """Stop a console window flashing up on Windows when we shell out."""
    if platform.system() == "Windows":
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    return {}


def _child_env():
    """Environment for LibreOffice / its Python. A stray PYTHONHOME or
    PYTHONPATH (conda, another Python install, a frozen exe) is a classic way
    to break LibreOffice's bundled Python, so drop them, and undo the library
    path PyInstaller points at its own bundle."""
    env = dict(os.environ)
    for key in ("PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP", "PYTHONNOUSERSITE"):
        env.pop(key, None)
    for key in ("LD_LIBRARY_PATH", "DYLD_LIBRARY_PATH"):
        orig = env.pop(key + "_ORIG", None)
        if orig is not None:
            env[key] = orig
        elif getattr(sys, "frozen", False):
            env.pop(key, None)
    return env


def powerpoint_available():
    """(True, "") when PowerPoint can be driven through COM, otherwise
    (False, reason). Cheap enough to call on the UI thread."""
    if platform.system() != "Windows":
        return False, "only used on Windows"
    try:
        import winreg
    except ImportError:
        return False, "Windows registry not readable"
    installed = False
    for view in (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0)):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID",
                                0, winreg.KEY_READ | view):
                installed = True
                break
        except OSError:
            continue
    if not installed:
        return False, "PowerPoint isn't installed"
    try:
        import win32com.client  # noqa: F401  (part of pywin32)
    except ImportError:
        return False, "PowerPoint is installed, but pywin32 isn't (pip install pywin32)"
    return True, ""


def render_presentation_to_pngs(pres_path, out_dir, width, progress=None, status=None,
                                soffice_hint=None):
    """Render each visible slide of a presentation to a lossless PNG inside
    out_dir. Returns (png_paths_in_slide_order, engine_name).

    progress(done, total) is called after each slide and status(text) when
    the conversion engine changes, from whatever thread this runs on.

    Windows: PowerPoint when it's installed (exact fidelity), otherwise - or
    if PowerPoint fails - LibreOffice. Elsewhere: LibreOffice.
    Raises RuntimeError explaining what was tried and how to fix it."""
    def say(text):
        if status:
            status(text)

    errors = []
    is_windows = platform.system() == "Windows"
    if is_windows:
        ok, why = powerpoint_available()
        if ok:
            say("Converting with PowerPoint...")
            try:
                return _render_with_powerpoint(pres_path, out_dir, width, progress), "PowerPoint"
            except Exception as e:
                errors.append(f"PowerPoint: {e}")
                for leftover in glob.glob(os.path.join(out_dir, "slide_*.png")):
                    try:
                        os.remove(leftover)
                    except OSError:
                        pass
        else:
            errors.append(why)

    soffice = _find_soffice(soffice_hint)
    if soffice:
        say(("PowerPoint not available - converting with LibreOffice..." if is_windows
             else "Converting with LibreOffice...") + " (the first run can take a little while)")
        try:
            return _render_with_libreoffice(pres_path, out_dir, width, progress, soffice), "LibreOffice"
        except Exception as e:
            errors.append(f"LibreOffice ({soffice}): {e}")
    else:
        errors.append("LibreOffice wasn't found on this computer.")

    fixes = [f"Install LibreOffice (free): {LIBREOFFICE_DOWNLOAD_URL}"]
    if is_windows:
        fixes.append("or, if PowerPoint is installed: pip install pywin32")
    if soffice:
        fixes.append("If LibreOffice is installed but failing, pip install pymupdf "
                     "gives it a second way to render the slides.")
    raise RuntimeError(
        "Couldn't convert the presentation.\n\n" + "\n\n".join(errors) +
        "\n\nTo fix:\n  " + "\n  ".join(fixes)
    )


def _render_with_powerpoint(pres_path, out_dir, width, progress):
    import pythoncom          # part of pywin32
    import win32com.client

    pythoncom.CoInitialize()  # COM must be initialised on this worker thread
    app = pres = None
    quit_after = False
    try:
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            raise RuntimeError(f"couldn't start PowerPoint ({e})") from None
        # PowerPoint is single-instance: if it's already open with the user's
        # own decks, attach quietly and leave it running afterwards.
        quit_after = app.Presentations.Count == 0
        # Positional args: FileName, ReadOnly, Untitled, WithWindow
        pres = app.Presentations.Open(os.path.abspath(pres_path), True, False, False)
        slide_w = float(pres.PageSetup.SlideWidth)
        slide_h = float(pres.PageSetup.SlideHeight)
        height = max(1, round(width * slide_h / slide_w))

        # Hidden returns msoTrue (-1) or msoFalse (0). Skip hidden slides so
        # the output matches what the LibreOffice route produces.
        slides = [s for s in pres.Slides if not s.SlideShowTransition.Hidden]
        total = len(slides)
        outputs = []
        for i, slide in enumerate(slides, 1):
            out = os.path.join(out_dir, f"slide_{i:04d}.png")
            slide.Export(out, "PNG", width, height)
            outputs.append(out)
            if progress:
                progress(i, total)
        return outputs
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None and quit_after:
            try:
                app.Quit()
            except Exception:
                pass
        app = pres = None
        pythoncom.CoUninitialize()


def _soffice_from_windows_registry():
    """soffice.exe paths LibreOffice's installer recorded in the registry
    (covers custom install folders and 32-bit/64-bit mismatches)."""
    try:
        import winreg
    except ImportError:
        return []
    found = []
    views = (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0))
    lookups = (
        (r"SOFTWARE\LibreOffice\UNO\InstallPath", True),   # value = the "program" folder
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe", False),  # value = the exe
    )
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for subkey, is_folder in lookups:
            for view in views:
                try:
                    with winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ | view) as key:
                        value, _type = winreg.QueryValueEx(key, "")
                except OSError:
                    continue
                if value:
                    value = os.path.expandvars(str(value).strip('"'))
                    found.append(os.path.join(value, "soffice.exe") if is_folder else value)
    return found


def _find_soffice(hint=None):
    """Full path to LibreOffice's soffice program, or None.

    Checks, in order: the location the user picked by hand (hint), PATH, the
    Windows registry, and the standard install folders for each OS."""
    if hint:
        hint = _normalise_soffice_choice(hint)
        if hint and os.path.isfile(hint):
            return hint
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates = []
    system = platform.system()
    if system == "Windows":
        candidates += _soffice_from_windows_registry()
        bases = []
        # ProgramW6432 is the real 64-bit "Program Files" even when this is a
        # 32-bit Python (where PROGRAMFILES points at "Program Files (x86)").
        for env in ("ProgramW6432", "PROGRAMFILES", "PROGRAMFILES(X86)"):
            base = os.environ.get(env)
            if base and base not in bases:
                bases.append(base)
        local = os.environ.get("LOCALAPPDATA")
        if local:
            bases.append(os.path.join(local, "Programs"))   # per-user installs
        for base in bases:
            candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
            # older installers used versioned folders, e.g. "LibreOffice 5"
            candidates += sorted(glob.glob(os.path.join(base, "LibreOffice*", "program", "soffice.exe")),
                                 reverse=True)
    elif system == "Darwin":
        for apps in ("/Applications", os.path.expanduser("~/Applications")):
            candidates.append(os.path.join(apps, "LibreOffice.app", "Contents", "MacOS", "soffice"))
    else:
        candidates += ["/usr/lib/libreoffice/program/soffice", "/usr/lib64/libreoffice/program/soffice"]
        candidates += sorted(glob.glob("/opt/libreoffice*/program/soffice"), reverse=True)
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def _normalise_soffice_choice(path):
    """Accept what people actually pick in a file dialog: soffice itself, a
    macOS LibreOffice.app bundle, or LibreOffice's install / program folder."""
    if not path:
        return None
    path = os.path.expanduser(str(path).strip().strip('"'))
    if path.lower().endswith(".app"):
        return os.path.join(path, "Contents", "MacOS", "soffice")
    if os.path.isdir(path):
        exe = "soffice.exe" if platform.system() == "Windows" else "soffice"
        for c in (os.path.join(path, exe), os.path.join(path, "program", exe)):
            if os.path.isfile(c):
                return c
    return path


def _render_with_libreoffice(pres_path, out_dir, width, progress, soffice):
    renderer = _pdf_renderer()
    if renderer is None:
        # Nothing to turn a PDF into images with, so have LibreOffice write
        # each slide as a PNG itself. Needs no extra installs.
        return _render_with_libreoffice_direct(pres_path, out_dir, width, progress, soffice)

    # A throwaway profile means this works even if LibreOffice is already
    # open - otherwise a headless run hands off to the open instance and
    # silently converts nothing.
    profile_uri = pathlib.Path(out_dir, "lo_profile").as_uri()
    cmd = [
        soffice, f"-env:UserInstallation={profile_uri}",
        "--headless", "--norestore", "--nologo", "--nolockcheck",
        "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(pres_path),
    ]
    stem = os.path.splitext(os.path.basename(pres_path))[0]
    pdf = os.path.join(out_dir, stem + ".pdf")
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=600, env=_child_env(), **_no_console_kwargs(),
        )
        # On some Windows setups soffice returns a moment before the file is
        # fully written; give it a few seconds to show up.
        deadline = time.time() + 5
        while not os.path.isfile(pdf) and time.time() < deadline:
            time.sleep(0.25)
        if not os.path.isfile(pdf):
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError("LibreOffice didn't produce a PDF." + (f" ({detail[-400:]})" if detail else ""))
        return _pdf_to_pngs(pdf, out_dir, width, progress, renderer)
    finally:
        # The slides now live on as PNGs (kept for browsing), so drop the
        # bulky leftovers.
        shutil.rmtree(os.path.join(out_dir, "lo_profile"), ignore_errors=True)
        try:
            os.remove(pdf)
        except OSError:
            pass


def _pdf_renderer():
    """What's available to rasterise a PDF: ("pymupdf", module),
    ("pdftoppm", path), or None."""
    try:
        import pymupdf
        return "pymupdf", pymupdf
    except ImportError:
        pass
    try:
        import fitz as pymupdf  # older PyMuPDF releases
        pymupdf.open            # guard against the unrelated "fitz" package
        return "pymupdf", pymupdf
    except (ImportError, AttributeError):
        pass
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        return "pdftoppm", pdftoppm
    return None


def _pdf_to_pngs(pdf, out_dir, width, progress, renderer):
    kind, tool = renderer
    if kind == "pymupdf":
        pymupdf = tool
        outputs = []
        doc = pymupdf.open(pdf)
        try:
            total = doc.page_count
            for i, page in enumerate(doc, 1):
                zoom = width / page.rect.width
                pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
                out = os.path.join(out_dir, f"slide_{i:04d}.png")
                pix.save(out)
                outputs.append(out)
                if progress:
                    progress(i, total)
        finally:
            doc.close()
        return outputs

    subprocess.run(
        [tool, "-png", "-scale-to-x", str(width), "-scale-to-y", "-1",
         pdf, os.path.join(out_dir, "page")],
        check=True, capture_output=True, timeout=600, **_no_console_kwargs(),
    )
    pages = []
    for fname in os.listdir(out_dir):
        m = re.match(r"^page-(\d+)\.png$", fname)
        if m:
            pages.append((int(m.group(1)), os.path.join(out_dir, fname)))
    pages.sort()
    if progress and pages:
        progress(len(pages), len(pages))
    return [p for _n, p in pages]


# Runs inside LibreOffice's own Python (which always has the "uno" module):
# connects to a headless soffice over a named pipe, opens the deck and exports
# every visible slide as a PNG at the requested pixel size.
_LO_EXPORT_SCRIPT = r'''
import os, sys, time, traceback
import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main():
    pipe, pres_path, out_dir, width = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    ctx = None
    deadline = time.time() + 120          # a first start with a fresh profile is slow
    while ctx is None:
        try:
            ctx = resolver.resolve(
                "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe)
        except Exception:
            if time.time() > deadline:
                raise SystemExit("Couldn't connect to LibreOffice (it didn't start in time).")
            time.sleep(0.5)
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(pres_path)), "_blank", 0,
            (prop("Hidden", True), prop("ReadOnly", True), prop("MacroExecutionMode", 0)))
        if doc is None or not hasattr(doc, "getDrawPages"):
            raise SystemExit("LibreOffice couldn't open that file as a presentation.")
        try:
            pages = doc.getDrawPages()
            visible = []
            for i in range(pages.getCount()):
                page = pages.getByIndex(i)
                try:
                    if not page.getPropertyValue("Visible"):   # hidden slide
                        continue
                except Exception:
                    pass
                visible.append(page)
            total = len(visible)
            export = smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", ctx)
            for n, page in enumerate(visible, 1):
                height = max(1, int(round(width * float(page.Height) / float(page.Width))))
                out = os.path.join(out_dir, "slide_%04d.png" % n)
                size = uno.Any("[]com.sun.star.beans.PropertyValue",
                               (prop("PixelWidth", width), prop("PixelHeight", height)))
                args = (prop("URL", uno.systemPathToFileUrl(out)),
                        prop("MediaType", "image/png"),
                        prop("FilterData", size))
                export.setSourceDocument(page)
                uno.invoke(export, "filter", (args,))
                print("SLIDE %d %d %s" % (n, total, out), flush=True)
        finally:
            doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass                          # the connection drops as it quits


try:
    main()
except SystemExit:
    raise
except Exception:
    traceback.print_exc()
    sys.exit(1)
'''


def _python_has_uno(python, timeout=20):
    try:
        r = subprocess.run([python, "-c", "import uno"], capture_output=True,
                           timeout=timeout, env=_child_env(), **_no_console_kwargs())
        return r.returncode == 0
    except Exception:
        return False


def _find_lo_python(soffice):
    """A Python interpreter that can 'import uno' for this LibreOffice."""
    program_dir = os.path.dirname(os.path.realpath(soffice))
    if platform.system() == "Windows":
        bundled = [os.path.join(program_dir, "python.exe")]
    else:
        bundled = [
            os.path.join(program_dir, "python"),                                 # libreoffice.org builds
            os.path.normpath(os.path.join(program_dir, "..", "Resources", "python")),  # macOS app bundle
        ]
    for py in bundled:
        if os.path.isfile(py):
            return py
    # Linux distro packages use the system Python plus the python3-uno package.
    others = []
    if not getattr(sys, "frozen", False):
        others.append(sys.executable)
    for name in ("python3", "python"):
        found = shutil.which(name)
        if found and found not in others:
            others.append(found)
    for py in others:
        if _python_has_uno(py):
            return py
    return None


def _render_with_libreoffice_direct(pres_path, out_dir, width, progress, soffice):
    python = _find_lo_python(soffice)
    if not python:
        hint = (" On Linux, install the python3-uno package (e.g. sudo apt install python3-uno)."
                if platform.system() == "Linux" else "")
        raise RuntimeError("couldn't find LibreOffice's Python to export slides with." + hint +
                           " Alternatively, pip install pymupdf.")

    script = os.path.join(out_dir, "lo_export_slides.py")
    with open(script, "w", encoding="utf-8") as f:
        f.write(_LO_EXPORT_SCRIPT)
    pipe = f"snipsave_{os.getpid()}_{int(time.time() * 1000)}"
    profile_uri = pathlib.Path(out_dir, "lo_profile").as_uri()
    env = _child_env()
    office = subprocess.Popen(
        [soffice, f"-env:UserInstallation={profile_uri}", "--headless", "--invisible",
         "--nologo", "--norestore", "--nodefault", "--nolockcheck",
         f"--accept=pipe,name={pipe};urp;StarOffice.ComponentContext"],
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        env=env, **_no_console_kwargs(),
    )
    worker = None
    succeeded = False
    outputs = []
    try:
        worker = subprocess.Popen(
            [python, script, pipe, os.path.abspath(pres_path), out_dir, str(width)],
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", errors="replace", env=env, **_no_console_kwargs(),
        )
        # stderr is drained on its own thread so a chatty error can't fill the
        # pipe and stall the export while stdout is being read.
        err_chunks = []
        err_reader = threading.Thread(target=lambda: err_chunks.append(worker.stderr.read()), daemon=True)
        err_reader.start()
        # Safety net: never let a stuck LibreOffice hang the import forever.
        watchdog = threading.Timer(900, lambda: (worker.kill(), office.kill()))
        watchdog.daemon = True
        watchdog.start()
        # If LibreOffice itself quits (bad install, crash) the export script
        # would sit waiting to connect for two minutes; stop it right away.
        # The grace period lets a normal finish - where the script tells
        # LibreOffice to quit as its last step - complete untouched.
        office_died = []

        def watch_office():
            while worker.poll() is None:
                if office.poll() is not None:
                    try:
                        worker.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        office_died.append(office.returncode)
                        worker.kill()
                    return
                time.sleep(0.5)

        threading.Thread(target=watch_office, daemon=True).start()
        try:
            for line in worker.stdout:
                parts = line.strip().split(" ", 3)
                if len(parts) == 4 and parts[0] == "SLIDE":
                    outputs.append(parts[3])
                    if progress:
                        progress(int(parts[1]), int(parts[2]))
            code = worker.wait()
            err_reader.join(timeout=5)
        finally:
            watchdog.cancel()
        if office_died:
            raise RuntimeError(f"LibreOffice quit before the slides could be exported "
                               f"(exit code {office_died[0]}). Try opening LibreOffice once by "
                               f"hand to check it starts, or reinstall it.")
        if code != 0:
            lines = [l for l in "".join(err_chunks).strip().splitlines() if l.strip()]
            raise RuntimeError(lines[-1] if lines else f"slide export exited with code {code}")
        outputs = [p for p in outputs if os.path.isfile(p)]
        if not outputs:
            raise RuntimeError("LibreOffice exported no slides (is every slide hidden?).")
        succeeded = True
        return outputs
    finally:
        if worker is not None and worker.poll() is None:
            worker.kill()
        # After a clean export the script has already asked LibreOffice to
        # quit, so give it a moment; after a failure, don't make anyone wait.
        try:
            office.wait(timeout=20 if succeeded else 1)
        except subprocess.TimeoutExpired:
            office.kill()
            try:
                office.wait(timeout=5)
            except Exception:
                pass
        shutil.rmtree(os.path.join(out_dir, "lo_profile"), ignore_errors=True)
        try:
            os.remove(script)
        except OSError:
            pass


# ------------------------------------------------------------------ HDR
# Windows composites the desktop in scRGB when HDR is on. A plain GDI /
# ImageGrab capture hands back 8-bit data from that buffer with no tone
# mapping, which is why HDR screenshots come out blown out or milky.
#
# We can't fix that inside ImageGrab, but we can (a) ask Windows whether
# the display being captured actually has HDR enabled and what its SDR
# reference white is, and (b) push the capture back down to an SDR-looking
# image with a transfer curve. See hdr_lut() for the honest limits.

QDC_ONLY_ACTIVE_PATHS = 2
DEVICE_INFO_GET_SOURCE_NAME = 1
DEVICE_INFO_GET_ADVANCED_COLOR_INFO = 9
DEVICE_INFO_GET_SDR_WHITE_LEVEL = 11


def windows_hdr_displays():
    """Map GDI device name -> {"supported", "enabled", "sdr_white_nits"}.

    Uses QueryDisplayConfig + DisplayConfigGetDeviceInfo. Returns {} on
    non-Windows, on older Windows without the advanced-color APIs, or if
    anything at all goes wrong - callers treat {} as "no HDR info".
    """
    if platform.system() != "Windows":
        return {}
    try:
        import ctypes
        from ctypes import wintypes

        class LUID(ctypes.Structure):
            _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]

        class RATIONAL(ctypes.Structure):
            _fields_ = [("Numerator", wintypes.UINT), ("Denominator", wintypes.UINT)]

        class PATH_SOURCE_INFO(ctypes.Structure):
            _fields_ = [("adapterId", LUID), ("id", wintypes.UINT),
                        ("modeInfoIdx", wintypes.UINT), ("statusFlags", wintypes.UINT)]

        class PATH_TARGET_INFO(ctypes.Structure):
            _fields_ = [("adapterId", LUID), ("id", wintypes.UINT), ("modeInfoIdx", wintypes.UINT),
                        ("outputTechnology", wintypes.UINT), ("rotation", wintypes.UINT),
                        ("scaling", wintypes.UINT), ("refreshRate", RATIONAL),
                        ("scanLineOrdering", wintypes.UINT), ("targetAvailable", wintypes.BOOL),
                        ("statusFlags", wintypes.UINT)]

        class PATH_INFO(ctypes.Structure):
            _fields_ = [("sourceInfo", PATH_SOURCE_INFO), ("targetInfo", PATH_TARGET_INFO),
                        ("flags", wintypes.UINT)]

        class MODE_INFO(ctypes.Structure):
            # the trailing union is 48 bytes; we never read it
            _fields_ = [("infoType", wintypes.UINT), ("id", wintypes.UINT),
                        ("adapterId", LUID), ("blob", ctypes.c_ubyte * 48)]

        class INFO_HEADER(ctypes.Structure):
            _fields_ = [("type", wintypes.UINT), ("size", wintypes.UINT),
                        ("adapterId", LUID), ("id", wintypes.UINT)]

        class ADVANCED_COLOR_INFO(ctypes.Structure):
            _fields_ = [("header", INFO_HEADER), ("value", wintypes.UINT),
                        ("colorEncoding", wintypes.UINT), ("bitsPerColorChannel", wintypes.UINT)]

        class SDR_WHITE_LEVEL(ctypes.Structure):
            _fields_ = [("header", INFO_HEADER), ("SDRWhiteLevel", wintypes.UINT)]

        class SOURCE_DEVICE_NAME(ctypes.Structure):
            _fields_ = [("header", INFO_HEADER), ("viewGdiDeviceName", wintypes.WCHAR * 32)]

        user32 = ctypes.windll.user32
        n_paths = wintypes.UINT()
        n_modes = wintypes.UINT()
        if user32.GetDisplayConfigBufferSizes(
            QDC_ONLY_ACTIVE_PATHS, ctypes.byref(n_paths), ctypes.byref(n_modes)
        ) != 0:
            return {}
        paths = (PATH_INFO * n_paths.value)()
        modes = (MODE_INFO * n_modes.value)()
        if user32.QueryDisplayConfig(
            QDC_ONLY_ACTIVE_PATHS, ctypes.byref(n_paths), paths,
            ctypes.byref(n_modes), modes, None,
        ) != 0:
            return {}

        result = {}
        for i in range(n_paths.value):
            path = paths[i]

            name = SOURCE_DEVICE_NAME()
            name.header.type = DEVICE_INFO_GET_SOURCE_NAME
            name.header.size = ctypes.sizeof(SOURCE_DEVICE_NAME)
            name.header.adapterId = path.sourceInfo.adapterId
            name.header.id = path.sourceInfo.id
            if user32.DisplayConfigGetDeviceInfo(ctypes.byref(name)) != 0:
                continue
            gdi_name = name.viewGdiDeviceName

            color = ADVANCED_COLOR_INFO()
            color.header.type = DEVICE_INFO_GET_ADVANCED_COLOR_INFO
            color.header.size = ctypes.sizeof(ADVANCED_COLOR_INFO)
            color.header.adapterId = path.targetInfo.adapterId
            color.header.id = path.targetInfo.id
            if user32.DisplayConfigGetDeviceInfo(ctypes.byref(color)) != 0:
                continue
            supported = bool(color.value & 0x1)         # advancedColorSupported
            enabled = bool(color.value & 0x2)           # advancedColorEnabled

            nits = None
            white = SDR_WHITE_LEVEL()
            white.header.type = DEVICE_INFO_GET_SDR_WHITE_LEVEL
            white.header.size = ctypes.sizeof(SDR_WHITE_LEVEL)
            white.header.adapterId = path.targetInfo.adapterId
            white.header.id = path.targetInfo.id
            if user32.DisplayConfigGetDeviceInfo(ctypes.byref(white)) == 0 and white.SDRWhiteLevel:
                # SDRWhiteLevel is in thousandths of 80-nit scRGB units
                nits = white.SDRWhiteLevel / 1000.0 * 80.0

            result[gdi_name.upper()] = {
                "supported": supported,
                "enabled": enabled,
                "sdr_white_nits": nits,
                "bits_per_channel": color.bitsPerColorChannel,
            }
        return result
    except Exception as e:
        print("HDR detection unavailable:", e)
        return {}


def _srgb_encode(x):
    if x <= 0.0031308:
        return 12.92 * x
    return 1.055 * (x ** (1 / 2.4)) - 0.055


def _srgb_decode(x):
    if x <= 0.04045:
        return x / 12.92
    return ((x + 0.055) / 1.055) ** 2.4


def _pq_eotf(v):
    """SMPTE ST 2084 inverse: code value (0-1) -> absolute luminance in nits."""
    m1, m2 = 0.1593017578125, 78.84375
    c1, c2, c3 = 0.8359375, 18.8515625, 18.6875
    vp = max(v, 0.0) ** (1 / m2)
    num = max(vp - c1, 0.0)
    den = c2 - c3 * vp
    if den <= 0:
        return 10000.0
    return ((num / den) ** (1 / m1)) * 10000.0


def _highlight_knee(y, knee=0.8):
    """Soft-clip toward 1.0 above `knee` instead of hard clipping, so
    near-white detail compresses rather than flattening into a block."""
    if y <= knee:
        return y
    return knee + (1.0 - knee) * (1.0 - math.exp(-(y - knee) / (1.0 - knee)))


def hdr_lut(curve="PQ", sdr_white_nits=200.0, strength=1.0):
    """256-entry LUT mapping an HDR capture back to something SDR-correct.

    Two failure modes, two curves:

      "PQ"     - the capture holds ST 2084 (HDR10) encoded values that are
                 being shown as if they were sRGB. Symptom: flat, milky,
                 grey-looking whites. Decode PQ to luminance, normalise
                 against the display's SDR reference white, re-encode sRGB.
      "linear" - the capture came out of the scRGB compositor already
                 gamma-encoded once too often. Symptom: blown out, too
                 bright, washed highlights. Undo one sRGB encode, which
                 pulls the midtones back down.

    `strength` blends between the raw capture (0.0) and the fully
    corrected version (1.0).

    Limit worth knowing: whatever Windows clipped on its way down to 8
    bits is gone. This restores exposure and contrast, not clipped
    highlight detail.
    """
    ref_nits = max(80.0, float(sdr_white_nits))
    lut = []
    for i in range(256):
        v = i / 255.0
        if curve == "PQ":
            y = _pq_eotf(v) / ref_nits
            out = _srgb_encode(min(max(_highlight_knee(y), 0.0), 1.0))
        else:
            out = _srgb_decode(v)
        out = (1.0 - strength) * v + strength * out
        lut.append(int(round(min(max(out, 0.0), 1.0) * 255)))
    return lut


def apply_hdr_correction(img, curve="PQ", sdr_white_nits=200.0, strength=1.0):
    lut = hdr_lut(curve, sdr_white_nits, strength)
    rgb = img.convert("RGB")
    return rgb.point(lut * 3)


# ------------------------------------------------------------ clipboard
def copy_image_to_clipboard(img):
    system = platform.system()
    try:
        if system == "Windows":
            _copy_clipboard_windows(img)
        elif system == "Darwin":
            _copy_clipboard_mac(img)
        else:
            _copy_clipboard_linux(img)
        return True
    except Exception as e:
        print("Clipboard copy failed:", e)
        return False


def _copy_clipboard_windows(img):
    import win32clipboard  # requires: pip install pywin32
    output = io.BytesIO()
    img.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]  # strip the 14-byte BMP file header to get a DIB
    output.close()
    # Another program can hold the clipboard for a moment; this runs on a
    # background thread now, so a few short retries cost the UI nothing.
    for attempt in range(5):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            if attempt == 4:
                raise
            time.sleep(0.05)
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


def _copy_clipboard_mac(img):
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        img.save(path, "PNG")
        script = f'set the clipboard to (read (POSIX file "{path}") as «class PNGf»)'
        subprocess.run(["osascript", "-e", script], check=True)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _copy_clipboard_linux(img):
    output = io.BytesIO()
    img.save(output, "PNG")
    data = output.getvalue()
    output.close()
    last_err = None
    for cmd in (
        ["xclip", "-selection", "clipboard", "-t", "image/png"],
        ["wl-copy", "--type", "image/png"],
    ):
        try:
            subprocess.run(cmd, input=data, check=True)
            return
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"No clipboard tool available (tried xclip, wl-copy): {last_err}")


class ClipboardWorker:
    """Copies images to the clipboard on one background thread.

    Encoding a big capture for the clipboard (a full two-monitor grab is tens
    of MB as a bitmap) used to hold up the window coming back after every
    capture and every Save. Requests are handled in order on a single thread;
    if a newer one arrives while an older one is still waiting, only the newer
    one is copied, so the clipboard always ends up holding the latest image.
    The image passed in is never modified afterwards (strokes are baked into
    a copy), so it's safe to read from this thread."""

    def __init__(self, call_on_ui):
        self._call_on_ui = call_on_ui
        self._cond = threading.Condition()
        self._job = None
        threading.Thread(target=self._run, daemon=True, name="clipboard").start()

    def copy(self, img, on_done=None):
        """Queue img for the clipboard; on_done(ok) then runs on the UI thread."""
        with self._cond:
            self._job = (img, on_done)
            self._cond.notify()

    def _run(self):
        if IS_WINDOWS:
            # Load pywin32 now rather than on the first capture (its DLLs are
            # one of the slower things to load in a built exe).
            try:
                import win32clipboard  # noqa: F401
            except Exception:
                pass
        while True:
            with self._cond:
                while self._job is None:
                    self._cond.wait()
                img, on_done = self._job
                self._job = None
            ok = copy_image_to_clipboard(img)
            if on_done is not None:
                self._call_on_ui(on_done, ok)


def apply_style(root):
    """Set up a clean, high-contrast look for the whole app."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    root.configure(bg=BG)
    root.option_add("*TCombobox*Listbox.background", CARD)
    root.option_add("*TCombobox*Listbox.foreground", TEXT)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
    root.option_add("*TCombobox*Listbox.font", (FONT_FAMILY, 10))
    # NOTE: a blanket "*Font" option would override every ttk style's font
    # (it did in v6 - headers and field titles never rendered bold), so set
    # fonts only on the widget classes that don't take them from a style.
    for cls in ("Listbox", "TEntry", "TCombobox", "Button"):
        root.option_add(f"*{cls}.Font", (FONT_FAMILY, 10))

    # ---- frames & labels
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("Toolbar.TFrame", background=CARD)

    style.configure("TLabel", background=BG, foreground=TEXT, font=(FONT_FAMILY, 10))
    style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=(FONT_FAMILY, 10))
    style.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=(FONT_FAMILY, 9))
    style.configure("Header.TLabel", background=BG, foreground=TEXT, font=(FONT_FAMILY, 20, "bold"))
    style.configure("FieldTitle.TLabel", background=CARD, foreground=TEXT, font=(FONT_FAMILY, 10, "bold"))
    style.configure("NextNum.TLabel", background=CARD, foreground=ACCENT, font=(FONT_FAMILY, 10, "bold"))
    style.configure("Status.TLabel", background=BG, foreground=MUTED, font=(FONT_FAMILY, 9))

    # ---- inputs: visible borders, accent ring on focus
    style.configure(
        "TEntry", fieldbackground=CARD, foreground=TEXT, insertcolor=TEXT, padding=7,
        bordercolor=BORDER_STRONG, lightcolor=CARD, darkcolor=CARD, borderwidth=1,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", ACCENT)],
        lightcolor=[("focus", ACCENT)],
        darkcolor=[("focus", ACCENT)],
    )
    style.configure(
        "TCombobox", fieldbackground=CARD, foreground=TEXT, padding=6,
        background=CARD, arrowcolor=TEXT, arrowsize=14,
        bordercolor=BORDER_STRONG, lightcolor=CARD, darkcolor=CARD,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", CARD)],
        foreground=[("readonly", TEXT)],
        selectbackground=[("readonly", CARD)],
        selectforeground=[("readonly", TEXT)],
        background=[("active", GHOST_HOVER)],
        bordercolor=[("focus", ACCENT), ("active", ACCENT)],
        lightcolor=[("focus", ACCENT)],
        darkcolor=[("focus", ACCENT)],
    )

    # ---- checkbuttons: filled accent box when on
    for name, bg, fg in (
        ("TCheckbutton", BG, TEXT),
        ("Card.TCheckbutton", CARD, TEXT),
        ("Order.TCheckbutton", BG, MUTED),
    ):
        style.configure(
            name, background=bg, foreground=fg, font=(FONT_FAMILY, 9),
            indicatorbackground=CARD, indicatorforeground="#ffffff",
            upperbordercolor=BORDER_STRONG, lowerbordercolor=BORDER_STRONG,
            indicatormargin=(0, 0, 8, 0), focuscolor=bg,
        )
        style.map(
            name,
            background=[("active", bg)],
            foreground=[("active", TEXT)],
            indicatorbackground=[("selected", ACCENT), ("active", GHOST_HOVER)],
            upperbordercolor=[("selected", ACCENT)],
            lowerbordercolor=[("selected", ACCENT)],
        )

    # ---- slider & scrollbar
    style.configure(
        "Horizontal.TScale", background=ACCENT, troughcolor=SUNKEN,
        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT, sliderthickness=16,
    )
    style.map("Horizontal.TScale", background=[("active", ACCENT_HOVER)])
    style.configure(
        "Vertical.TScrollbar", background=SUNKEN, troughcolor=CARD, arrowcolor=TEXT,
        bordercolor=BORDER, lightcolor=SUNKEN, darkcolor=SUNKEN, gripcount=0,
    )
    style.map("Vertical.TScrollbar", background=[("active", BORDER)])

    style.configure(
        "TSpinbox", fieldbackground=CARD, foreground=TEXT, padding=5, arrowsize=12,
        background=CARD, arrowcolor=TEXT, bordercolor=BORDER_STRONG, lightcolor=CARD, darkcolor=CARD,
    )
    style.map("TSpinbox", bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])

    style.configure("TSeparator", background=BORDER)

    # ---- buttons
    def button_style(name, bg, fg, hover, size=9, pad=(12, 8), border=None, hover_fg=None):
        edge = border or bg
        style.configure(
            name, background=bg, foreground=fg, font=(FONT_FAMILY, size, "bold"),
            padding=pad, borderwidth=1, bordercolor=edge, lightcolor=bg, darkcolor=bg,
            focuscolor=fg, width=0,
        )
        style.map(
            name,
            background=[("pressed", hover), ("active", hover)],
            lightcolor=[("pressed", hover), ("active", hover)],
            darkcolor=[("pressed", hover), ("active", hover)],
            bordercolor=[("active", border or hover)],
            foreground=[("disabled", MUTED), ("active", hover_fg or fg)],
        )

    # Big action grid (bottom of window)
    button_style("Region.TButton", ACCENT, "#ffffff", ACCENT_HOVER, size=11, pad=(16, 11))
    button_style("Save.TButton", SAVE, "#ffffff", SAVE_HOVER, size=11, pad=(16, 11))
    button_style("Full.TButton", SLATE, "#ffffff", SLATE_HOVER, size=11, pad=(16, 11))
    button_style("Manual.TButton", CARD, TEXT, GHOST_HOVER, size=11, pad=(16, 10), border=BORDER_STRONG)

    # Small utility buttons
    button_style("Primary.TButton", ACCENT, "#ffffff", ACCENT_HOVER)
    button_style("Secondary.TButton", CARD, TEXT, GHOST_HOVER, border=BORDER_STRONG)
    button_style("Ghost.TButton", CARD, TEXT, GHOST_HOVER, pad=(10, 6), border=BORDER_STRONG)
    button_style("Danger.TButton", CARD, DANGER, DANGER_TINT, pad=(10, 6), border=DANGER)

    # Drawing-tool toggles (Pen / Rectangle / Arrow): radio buttons drawn
    # like the outlined Ghost buttons, with the chosen one filled in accent.
    style.layout("Tool.TRadiobutton", style.layout("TButton"))
    style.configure(
        "Tool.TRadiobutton", background=CARD, foreground=TEXT, font=(FONT_FAMILY, 9, "bold"),
        padding=(10, 6), borderwidth=1, bordercolor=BORDER_STRONG, lightcolor=CARD,
        darkcolor=CARD, focuscolor=CARD, anchor="center", relief="raised",
    )
    style.map(
        "Tool.TRadiobutton",
        background=[("selected", "active", ACCENT_HOVER), ("selected", ACCENT), ("active", GHOST_HOVER)],
        foreground=[("selected", "#ffffff")],
        bordercolor=[("selected", ACCENT)],
        lightcolor=[("selected", "active", ACCENT_HOVER), ("selected", ACCENT), ("active", GHOST_HOVER)],
        darkcolor=[("selected", "active", ACCENT_HOVER), ("selected", ACCENT), ("active", GHOST_HOVER)],
        focuscolor=[("selected", ACCENT)],
    )


def enable_dpi_awareness():
    """Best-effort: make both Tk and screeninfo/ImageGrab agree on real,
    physical screen pixels on Windows.

    This matters most with multiple monitors running DIFFERENT scaling
    percentages (e.g. laptop panel at 150%, external monitor at 100%).
    With only coarse "system DPI aware" mode (the old behavior here),
    Windows silently rescales non-primary-monitor coordinates for us,
    so the monitor bounds screeninfo reports and the pixels ImageGrab
    actually captures can disagree slightly - which shows up as the
    edge of a non-primary display being cut off or misaligned.
    "Per-Monitor v2" awareness (Windows 10 1703+) is the fix: each
    monitor's coordinates come back in its own real pixels. We fall
    back gracefully on older Windows where it isn't available."""
    if platform.system() != "Windows":
        return
    import ctypes

    # Try newest -> oldest API. Any one succeeding is enough; try the next
    # option only if the current one raises/returns failure.
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4 (Windows 10 1703+)
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2 (Windows 8.1+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        # PROCESS_SYSTEM_DPI_AWARE = 1 (Windows 8.1+, coarser fallback)
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        pass
    try:
        # Vista+ fallback, no per-monitor support
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


# ------------------------------------------------------------ fast capture
def _toplevel_hwnd(win):
    """Windows handle of a Tk toplevel's outer frame, or None."""
    try:
        return int(win.wm_frame(), 16) or None
    except Exception:
        return None


def disable_window_animations(win):
    """Stop Windows animating this window when it hides and reappears.

    Windows can play a short fade/zoom whenever a window is hidden or shown.
    During a capture the window has to be fully gone before the screen is
    grabbed, so any such animation is pure waiting, and a fade caught midway
    would show up in the screenshot. Returns True if Windows accepted it."""
    if not IS_WINDOWS:
        return False
    hwnd = _toplevel_hwnd(win)
    if not hwnd:
        return False
    try:
        import ctypes
        DWMWA_TRANSITIONS_FORCEDISABLED = 3
        on = ctypes.c_int(1)
        result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            ctypes.c_void_p(hwnd), DWMWA_TRANSITIONS_FORCEDISABLED,
            ctypes.byref(on), ctypes.sizeof(on),
        )
        return result == 0
    except Exception:
        return False


def wait_for_screen_refresh():
    """Block until Windows has composed the next frame (DwmFlush), so a window
    that was just hidden is really off the screen. Costs at most one refresh
    (about 17 ms at 60 Hz, 7 ms at 144 Hz). False where it isn't available."""
    if not IS_WINDOWS:
        return False
    try:
        import ctypes
        return ctypes.windll.dwmapi.DwmFlush() == 0
    except Exception:
        return False


class FastGrabber:
    """Screen grabs through the optional "mss" package (Windows).

    Pillow's ImageGrab copies the entire desktop - every monitor - through a
    24-bit conversion and then crops. mss copies only the rectangle asked
    for, as 32-bit pixels straight into a bitmap it reuses between captures.
    Create it and use it on the UI thread only."""

    def __init__(self):
        import mss
        # Pillow's grab never included "layered" windows (the CAPTUREBLT
        # flag); mss does by default, which is slower and makes the mouse
        # pointer flicker. Keep the old behaviour. (The flag lives in
        # mss.windows.gdi from mss 10.2 on, and in mss.windows before that.)
        for name in ("mss.windows.gdi", "mss.windows"):
            try:
                mod = importlib.import_module(name)
            except Exception:
                continue
            if hasattr(mod, "CAPTUREBLT"):
                mod.CAPTUREBLT = 0
        factory = getattr(mss, "MSS", None) or mss.mss   # mss.MSS since 10.2
        self._sct = factory()

    def grab(self, bbox=None):
        if bbox is None:
            area = self._sct.monitors[0]          # the whole virtual desktop
        else:
            left, top, right, bottom = bbox
            area = {"left": left, "top": top, "width": right - left, "height": bottom - top}
        shot = self._sct.grab(area)
        data = getattr(shot, "raw", None)   # BGRA bytes, no extra copy
        if data is None:
            data = shot.bgra
        return Image.frombytes("RGB", tuple(shot.size), data, "raw", "BGRX")

    def close(self):
        try:
            self._sct.close()
        except Exception:
            pass


# ------------------------------------------------------------ Print Screen key
# With "Use the Print screen key to open screen capture" on (Settings >
# Accessibility > Keyboard - the default in current Windows 11), Windows keeps
# Print Screen for its own Snipping Tool. That switch is the registry value
# below. Snip & Save turns it off while its own Print Screen option is on, and
# remembers how it was so switching the option off can put it back.
SNIP_KEY_REG_PATH = r"Control Panel\Keyboard"
SNIP_KEY_REG_VALUE = "PrintScreenKeyForSnippingEnabled"
WINDOWS_KEYBOARD_SETTINGS_URI = "ms-settings:easeofaccess-keyboard"


def windows_snip_key_state():
    """Windows' own Print Screen -> screen capture setting: "on", "off", or
    "unset" (nothing stored, so Windows uses its default). None if it can't
    be read, or not on Windows."""
    if not IS_WINDOWS:
        return None
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, SNIP_KEY_REG_PATH) as key:
            value, _type = winreg.QueryValueEx(key, SNIP_KEY_REG_VALUE)
        return "on" if int(value) else "off"
    except FileNotFoundError:
        return "unset"
    except Exception:
        return None


def set_windows_snip_key_state(state):
    """Set that switch to "on" or "off", or "unset" to remove the value and go
    back to Windows' default. Raises OSError if the registry refuses."""
    import winreg
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, SNIP_KEY_REG_PATH, 0,
                            winreg.KEY_SET_VALUE) as key:
        if state == "unset":
            try:
                winreg.DeleteValue(key, SNIP_KEY_REG_VALUE)
            except FileNotFoundError:
                pass
        else:
            winreg.SetValueEx(key, SNIP_KEY_REG_VALUE, 0, winreg.REG_DWORD,
                              1 if state == "on" else 0)


def open_windows_keyboard_settings():
    """Open Settings > Accessibility > Keyboard, where Windows' Print Screen
    switch lives."""
    try:
        os.startfile(WINDOWS_KEYBOARD_SETTINGS_URI)
    except Exception as e:
        print("Couldn't open Windows Settings:", e)


def force_foreground(win):
    """Best effort: make win the active window so it gets the keyboard (Esc,
    Ctrl+C) even though another program was active a moment ago."""
    try:
        win.lift()
        win.focus_force()
    except tk.TclError:
        return
    hwnd = _toplevel_hwnd(win) if IS_WINDOWS else None
    if hwnd:
        try:
            import ctypes
            ctypes.windll.user32.SetForegroundWindow(ctypes.c_void_p(hwnd))
        except Exception:
            pass


class PrintScreenHotkey:
    """Claims the Print Screen key system-wide (Windows) so pressing it starts
    a Snip & Save capture.

    RegisterHotKey runs on a small thread of its own that waits for WM_HOTKEY;
    Tk's message loop never delivers thread-level hotkey messages. If the key
    is taken (Windows' own switch, or another screenshot tool), the thread
    keeps retrying every couple of seconds and takes the key the moment it's
    free. Both callbacks run on the hotkey thread, so the app routes them to
    the UI thread:
        on_press(perf_counter_time_of_the_key_press)
        on_status(registered, windows_error_code)
    on_status(False, ...) is sent once, after ~1.5 s of failed attempts (Windows
    can take a moment to let go of the key); on_status(True, 0) when it's ours."""

    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012
    VK_SNAPSHOT = 0x2C
    MOD_NOREPEAT = 0x4000          # holding the key down doesn't fire again and again
    HOTKEY_ID = 0xB5A1             # any id in 0x0000-0xBFFF is fine for an application
    ERROR_HOTKEY_ALREADY_REGISTERED = 1409
    QUICK_RETRIES = 6              # 0.25 s apart, before reporting the key as taken
    SLOW_RETRY_S = 2.0

    def __init__(self, on_press, on_status):
        self._on_press = on_press
        self._on_status = on_status
        self._thread = None
        self._thread_id = None
        self._stop = threading.Event()

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if not IS_WINDOWS or self.running:
            return
        self._stop = threading.Event()
        self._thread_id = None
        self._thread = threading.Thread(target=self._run, args=(self._stop,),
                                        daemon=True, name="printscreen-hotkey")
        self._thread.start()

    def stop(self):
        thread, self._thread = self._thread, None
        if thread is None:
            return
        self._stop.set()                       # ends the retry loop
        tid = self._thread_id
        if tid:
            try:
                import ctypes
                # ends GetMessage; the thread then releases the key itself
                ctypes.windll.user32.PostThreadMessageW(tid, self.WM_QUIT, 0, 0)
            except Exception:
                pass
        thread.join(timeout=1.0)

    def _run(self, stop):
        try:
            import ctypes
            from ctypes import wintypes
            # A private WinDLL so these argtypes don't touch ctypes.windll.user32,
            # which the HDR and capture code call without argtypes.
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            kernel32 = ctypes.WinDLL("kernel32")
            user32.RegisterHotKey.argtypes = (wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT)
            user32.RegisterHotKey.restype = wintypes.BOOL
            user32.UnregisterHotKey.argtypes = (wintypes.HWND, ctypes.c_int)
            user32.UnregisterHotKey.restype = wintypes.BOOL
            user32.GetMessageW.argtypes = (ctypes.POINTER(wintypes.MSG), wintypes.HWND,
                                           wintypes.UINT, wintypes.UINT)
            user32.GetMessageW.restype = wintypes.BOOL      # can be -1 on error
            user32.PeekMessageW.argtypes = (ctypes.POINTER(wintypes.MSG), wintypes.HWND,
                                            wintypes.UINT, wintypes.UINT, wintypes.UINT)
            user32.PeekMessageW.restype = wintypes.BOOL
            msg = wintypes.MSG()
            # Create this thread's message queue before publishing its id, so a
            # WM_QUIT posted by stop() can never be lost.
            user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)
            self._thread_id = kernel32.GetCurrentThreadId()
        except Exception as e:
            print("Print Screen hotkey unavailable:", e)
            self._on_status(False, -1)
            return

        registered = False
        failures = 0
        try:
            while not stop.is_set():
                if user32.RegisterHotKey(None, self.HOTKEY_ID, self.MOD_NOREPEAT, self.VK_SNAPSHOT):
                    registered = True
                    break
                error = ctypes.get_last_error()
                failures += 1
                if failures == self.QUICK_RETRIES:
                    self._on_status(False, error)
                stop.wait(0.25 if failures < self.QUICK_RETRIES else self.SLOW_RETRY_S)
            if not registered:
                return
            self._on_status(True, 0)
            while not stop.is_set():
                result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if result == 0 or result == -1:          # WM_QUIT, or an error
                    break
                if msg.message == self.WM_HOTKEY and msg.wParam == self.HOTKEY_ID:
                    self._on_press(time.perf_counter())
        finally:
            if registered:
                user32.UnregisterHotKey(None, self.HOTKEY_ID)


# ----------------------------------------------------------------------
# Main application
# ----------------------------------------------------------------------
# ------------------------------------------------------------ shape annotations
def arrow_shape(x0, y0, x1, y1, width):
    """Geometry of an arrow from (x0, y0) to its tip at (x1, y1) with a shaft
    `width` thick: (shaft_end, [tip, left, right]) or None if it has no length.
    The head grows with the line width, shrinks to fit short arrows, and the
    shaft stops inside the head so its flat end never pokes past the sides.
    Works in any coordinate space, so the preview and the saved image share it."""
    dx, dy = x1 - x0, y1 - y0
    length = math.hypot(dx, dy)
    if length < 1e-6:
        return None
    ux, uy = dx / length, dy / length
    head_len = max(10.0, width * 3.2)
    head_half = max(6.0, width * 1.9)
    if head_len > length * 0.6:            # short arrow: scale the head down
        k = length * 0.6 / head_len
        head_len, head_half = head_len * k, head_half * k
    bx, by = x1 - ux * head_len, y1 - uy * head_len
    px, py = -uy, ux                       # perpendicular
    left = (bx + px * head_half, by + py * head_half)
    right = (bx - px * head_half, by - py * head_half)
    shaft_end = (x1 - ux * head_len * 0.5, y1 - uy * head_len * 0.5)
    return shaft_end, [(x1, y1), left, right]


def constrain_shape(kind, x0, y0, x1, y1):
    """Shift held while dragging: rectangles become squares, arrows snap to
    the nearest 45 degrees."""
    dx, dy = x1 - x0, y1 - y0
    if kind == "rect":
        side = max(abs(dx), abs(dy))
        return x0 + math.copysign(side, dx or 1), y0 + math.copysign(side, dy or 1)
    if kind == "arrow":
        length = math.hypot(dx, dy)
        angle = round(math.atan2(dy, dx) / (math.pi / 4)) * (math.pi / 4)
        return x0 + length * math.cos(angle), y0 + length * math.sin(angle)
    return x1, y1


# ------------------------------------------------------------ quick-capture button
class QuickCaptureButton:
    """A small round button that floats above every other window, so a
    capture is one click away without opening the main window first.

    - Click: capture a region (straight away - Delay is for the main buttons).
    - Drag: move it anywhere; the spot is remembered.
    - Right-click: full screen capture, open Snip & Save, or hide the button
      (View > Show quick-capture button brings it back).

    It stays up while the main window is minimised and hides itself for the
    moment of every capture so it's never in the shot."""

    KEY = "#ff00fe"      # painted where the button isn't; Windows makes it see-through
    DRAG_THRESHOLD = 5   # px the pointer must move before a press becomes a drag

    def __init__(self, app):
        self.app = app
        root = app.root
        try:
            dpi_scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        except tk.TclError:
            dpi_scale = 1.0
        self.size = S = int(round(44 * dpi_scale))

        self.win = win = tk.Toplevel(root)
        win.withdraw()
        win.overrideredirect(True)          # no title bar, no taskbar entry
        win.attributes("-topmost", True)
        win.title("Snip & Save - quick capture")

        round_ok = False
        if IS_WINDOWS:
            try:
                win.configure(bg=self.KEY)
                win.attributes("-transparentcolor", self.KEY)
                round_ok = True
            except tk.TclError:
                pass
        self.canvas = c = tk.Canvas(win, width=S, height=S, highlightthickness=0, bd=0,
                                    bg=self.KEY if round_ok else ACCENT, cursor="hand2")
        c.pack()
        ring = max(2, int(round(2 * dpi_scale)))
        if round_ok:
            self.shape = c.create_oval(ring, ring, S - ring, S - ring,
                                       fill=ACCENT, outline="#ffffff", width=ring)
        else:   # other systems can't cut a window to shape: a square tile
            self.shape = c.create_rectangle(ring // 2, ring // 2, S - ring // 2, S - ring // 2,
                                            fill=ACCENT, outline="#ffffff", width=ring)
        c.create_text(S / 2, S / 2, text="✂", fill="#ffffff", font=(FONT_FAMILY, 16, "bold"))

        self._press_pt = None
        self._win_start = (0, 0)
        self._drag_pos = (0, 0)
        self._dragging = False
        self._tip = None
        self._tip_after = None

        look = dict(tearoff=0, bg=CARD, fg=TEXT, activebackground=ACCENT,
                    activeforeground="#ffffff", activeborderwidth=0, relief="flat", bd=1)
        self.menu = tk.Menu(win, **look)
        self.menu.add_command(label="Capture Region", command=lambda: app.quick_capture("region"))
        self.menu.add_command(label="Capture Full Screen", command=lambda: app.quick_capture("full"))
        self.menu.add_separator()
        self.menu.add_command(label="Open Snip & Save", command=app.show_main_window)
        self.menu.add_command(label="Hide this button", command=app.hide_quick_button)

        c.bind("<ButtonPress-1>", self._on_press)
        c.bind("<B1-Motion>", self._on_motion)
        c.bind("<ButtonRelease-1>", self._on_release)
        c.bind("<Enter>", self._on_enter)
        c.bind("<Leave>", self._on_leave)
        c.bind("<Button-3>", self._on_menu)
        if platform.system() == "Darwin":
            c.bind("<Button-2>", self._on_menu)

    # ---- placement
    def default_position(self, monitors):
        """Top centre of the main screen."""
        if monitors:
            m = next((mo for mo in monitors if "(Primary)" in mo["label"]), monitors[0])
            return m["x"] + (m["width"] - self.size) // 2, m["y"] + 6
        return (self.app.root.winfo_screenwidth() - self.size) // 2, 6

    def place(self, saved, monitors):
        """Move to a saved "x,y" if it's still on a connected screen, else the default."""
        pos = None
        m = re.match(r"^\s*(-?\d+)\s*,\s*(-?\d+)\s*$", str(saved or ""))
        if m:
            x, y = int(m.group(1)), int(m.group(2))
            cx, cy = x + self.size // 2, y + self.size // 2
            if monitors:
                if any(mo["x"] <= cx < mo["x"] + mo["width"] and mo["y"] <= cy < mo["y"] + mo["height"]
                       for mo in monitors):
                    pos = (x, y)
            else:
                sw, sh = self.app.root.winfo_screenwidth(), self.app.root.winfo_screenheight()
                if 0 <= cx < sw and 0 <= cy < sh:
                    pos = (x, y)
        if pos is None:
            pos = self.default_position(monitors)
        self.win.geometry(f"{self.size}x{self.size}+{pos[0]}+{pos[1]}")

    def position(self):
        return self.win.winfo_x(), self.win.winfo_y()

    def show(self):
        self.win.deiconify()
        self.win.lift()
        self.win.attributes("-topmost", True)

    def hide(self):
        self._hide_tip()
        self._press_pt = None
        self._dragging = False
        self.win.withdraw()

    def visible(self):
        try:
            return bool(self.win.winfo_ismapped())
        except tk.TclError:
            return False

    # ---- mouse
    def _set_hover(self, on):
        self.canvas.itemconfigure(self.shape, fill=ACCENT_HOVER if on else ACCENT)

    def _on_press(self, e):
        self._hide_tip()
        self._press_pt = (e.x_root, e.y_root)
        self._win_start = self.position()
        self._dragging = False

    def _on_motion(self, e):
        if self._press_pt is None:
            return
        dx, dy = e.x_root - self._press_pt[0], e.y_root - self._press_pt[1]
        if not self._dragging and abs(dx) + abs(dy) < self.DRAG_THRESHOLD:
            return
        self._dragging = True
        self._drag_pos = (self._win_start[0] + dx, self._win_start[1] + dy)
        self.win.geometry(f"+{self._drag_pos[0]}+{self._drag_pos[1]}")

    def _on_release(self, e):
        if self._press_pt is None:
            return
        self._press_pt = None
        if self._dragging:
            self._dragging = False
            self.app.quick_button_moved(*self._drag_pos)
        elif 0 <= e.x < self.size and 0 <= e.y < self.size:
            self._set_hover(False)
            self.app.quick_capture("region")

    def _on_menu(self, e):
        self._hide_tip()
        try:
            self.menu.tk_popup(e.x_root, e.y_root)
        finally:
            self.menu.grab_release()

    # ---- tooltip
    def _on_enter(self, e):
        self._set_hover(True)
        if self._tip_after is None and self._press_pt is None:
            self._tip_after = self.win.after(650, self._show_tip)

    def _on_leave(self, e):
        self._set_hover(False)
        self._hide_tip()

    def _show_tip(self):
        self._tip_after = None
        if self._tip is not None or not self.visible():
            return
        tip = tk.Toplevel(self.win)
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        tk.Label(tip, text="Click to capture a region\nRight-click for more  •  drag to move",
                 bg=SLATE, fg="#ffffff", font=(FONT_FAMILY, 9), justify="left",
                 padx=10, pady=6).pack()
        tip.update_idletasks()
        x, y = self.position()
        tx = x + self.size // 2 - tip.winfo_reqwidth() // 2
        ty = y + self.size + 6
        if ty + tip.winfo_reqheight() > self.app.root.winfo_screenheight() and y > tip.winfo_reqheight():
            ty = y - tip.winfo_reqheight() - 6       # no room below: show it above
        tip.geometry(f"+{tx}+{ty}")
        self._tip = tip

    def _hide_tip(self):
        if self._tip_after is not None:
            try:
                self.win.after_cancel(self._tip_after)
            except tk.TclError:
                pass
            self._tip_after = None
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None

    def destroy(self):
        self._hide_tip()
        try:
            self.win.destroy()
        except tk.TclError:
            pass


class SnippingApp:
    COPY_LABEL = "\U0001F4CB  Copy to Clipboard"

    def __init__(self, root):
        self.root = root
        self.root.title("Snip & Save")
        self.root.resizable(True, True)
        self.root.minsize(540, 840)
        self.cfg = load_config()
        self.monitors = get_monitors()
        self._last_geometry = usable_geometry(
            self.cfg.get("window_geometry"), self.monitors,
            self.root.winfo_screenwidth(), self.root.winfo_screenheight(),
        )
        self.root.geometry(self._last_geometry)
        os.makedirs(self.cfg["save_dir"], exist_ok=True)
        self._geom_after = None            # pending debounced geometry save
        self._config_warning_shown = False

        # --- editing state ---
        self.raw_img = None           # capture exactly as grabbed, never modified
        self.working_img = None       # raw_img after HDR correction; what gets drawn/saved
        # HDR info is asked for on a background thread (once just after the
        # window appears, then at every capture), never while you wait.
        self.hdr_displays = {}
        self._hdr_checked = not IS_WINDOWS
        self._hdr_probe = None
        self._hdr_note = ""           # what HDR did on the last capture, for the status line
        self._detected_nits = 200.0   # SDR reference white reported by Windows, when known
        # list of {"kind": "pen"|"rect"|"arrow", "points": [(x,y)...], "width": px, "color": "#rrggbb"}
        # (rect / arrow keep two points: where the drag started and ended; no "kind" = pen)
        self.strokes = []
        self._current_stroke = None
        self._last_canvas_pt = None
        self._preview_photo = None    # keep a reference so Tk doesn't garbage-collect it
        self._preview_cache = None    # (image, (w, h), photo): skips re-resizing on undo/clear/redraw
        self._img_scale = 1.0
        self._img_offset = (0, 0)
        self.color = "#ff3b30"
        self._window_rect_before_capture = None  # (x, y, w, h), captured just before hiding

        # --- capture speed ---
        self._grabber = None          # FastGrabber (mss) once ready; None = Pillow
        self._grabber_failed = False
        self._animations_off = False
        self._last_grab_engine = "Pillow"
        self._t_click = None          # perf_counter() at the capture click
        self._timing_note = ""        # " [overlay in N ms ...]" for the status line

        # --- slide import / browsing state ---
        self._slides_dir = self.cfg.get("slides_dir", "")
        self._soffice_path = self.cfg.get("soffice_path", "")   # only set if picked by hand
        self._slide_export_running = False
        self.import_btn = None
        # A browsable slide session: lossless PNGs kept in a temp folder,
        # plus the names they were saved under. Annotations are remembered
        # per slide so flipping away and back doesn't lose your drawing.
        self._slide_paths = []      # temp PNG per slide, in order
        self._slide_names = []      # saved filename per slide (for the status line)
        self._slide_idx = 0
        self._slide_tmp = None      # temp folder holding the PNGs
        self._slide_strokes = {}    # slide index -> that slide's strokes
        self._settings_visible = True   # the folder/name/format + existing-files cards

        # Background threads (capture delay, slide export) hand work to the
        # window through this queue. Calling root.after() from another thread
        # only works on Pythons whose Tcl is built with thread support - on
        # others the call is silently dropped and the app looks frozen.
        self._ui_calls = queue.Queue()
        self._clipboard = ClipboardWorker(self._call_on_ui)

        # --- copy button / Print Screen key ---
        self._copy_flash_after = None        # pending reset of the "Copied" button text
        self._prtsc_restore = self.cfg.get("printscreen_restore", "")
        self._prtsc_user_initiated = False   # just switched on from the menu: pop up if the key is taken
        self._prtsc_waiting = False          # key is taken; the hotkey thread keeps retrying
        self._from_hotkey = False            # the capture under way was started by Print Screen
        self._state_before_capture = "normal"
        self._trigger_word = "click"         # for the timing note: "click" or "key press"
        self._prtsc = PrintScreenHotkey(
            on_press=lambda t: self._call_on_ui(self._on_print_screen, t),
            on_status=lambda ok, err: self._call_on_ui(self._on_print_screen_status, ok, err),
        )

        # --- quick-capture button (small, round, always on top) ---
        self._quick_pos = self.cfg.get("quick_button_pos", "")
        self.quick_btn = None

        self._build_ui()
        self.quick_btn = QuickCaptureButton(self)
        self.quick_btn.place(self._quick_pos, self.monitors)
        if self.quick_btn_var.get():
            self.quick_btn.show()
        self.refresh_file_list()
        self._pump_ui_calls()
        if IS_WINDOWS and self.prtsc_var.get():
            self._enable_print_screen(user_initiated=False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        # Save the window's size/position shortly after every resize or move,
        # not just on a clean close, so it sticks however the app is exited.
        self.root.bind("<Configure>", self._on_root_configure)
        # Once the window is up, get the first capture's one-time costs out
        # of the way (and ask Windows about HDR) without delaying startup.
        self.root.after(150, self._warm_up)

    def _warm_up(self):
        if IS_WINDOWS:
            self._start_hdr_probe(announce=True)
        self._ensure_capture_ready()

    def _ensure_capture_ready(self):
        """One-time capture setup; cheap to call again."""
        if not self._animations_off and self.root.winfo_ismapped():
            disable_window_animations(self.root)
            self._animations_off = True
        if self._grabber is None and IS_WINDOWS and not self._grabber_failed:
            try:
                self._grabber = FastGrabber()
            except Exception as e:   # mss not installed, or it couldn't start
                self._grabber_failed = True
                if not isinstance(e, ImportError):
                    print("mss unavailable, using Pillow's grab:", e)

    # ================================================================= UI
    def _card(self, parent, pady=(12, 0), fill="x", expand=False, padding=(18, 14)):
        """White card with a visible 1px outline; returns the inner body frame."""
        wrap = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        wrap.pack(fill=fill, expand=expand, pady=pady)
        body = ttk.Frame(wrap, style="Card.TFrame", padding=padding)
        body.pack(fill=fill, expand=expand)
        return body

    def _build_menu(self):
        """Menu bar along the top of the window. Layout preferences live here
        so they don't take up room in the window itself."""
        # Windows draws its menus natively and ignores these colours; on Linux
        # Tk draws them itself, so match the app's palette there.
        look = dict(tearoff=0, bg=CARD, fg=TEXT, activebackground=ACCENT,
                    activeforeground="#ffffff", activeborderwidth=0, relief="flat", bd=1)
        menubar = tk.Menu(self.root, **look)

        edit = tk.Menu(menubar, **look)
        edit.add_command(label="Copy to Clipboard", accelerator="Ctrl+C",
                         command=self.copy_current, state="disabled")
        menubar.add_cascade(label="Edit", menu=edit, underline=0)

        view = tk.Menu(menubar, **look)
        view.add_checkbutton(
            label='Put "Capture Region" and "Save" on the left',
            variable=self.region_left_var, onvalue=True, offvalue=False,
            command=self._on_region_side_changed,
        )
        view.add_separator()
        view.add_checkbutton(
            label="Show quick-capture button (always on top)",
            variable=self.quick_btn_var, onvalue=True, offvalue=False,
            command=self._on_quick_button_toggled,
        )
        view.add_command(label="Move quick-capture button back to the top",
                         command=self.reset_quick_button)
        menubar.add_cascade(label="View", menu=view, underline=0)

        if IS_WINDOWS:
            options = tk.Menu(menubar, **look)
            options.add_checkbutton(
                label="Print Screen key opens Snip & Save (instead of Snipping Tool)",
                variable=self.prtsc_var, onvalue=True, offvalue=False,
                command=self._on_print_screen_toggled,
            )
            options.add_separator()
            options.add_radiobutton(label="Print Screen captures a region",
                                    variable=self.prtsc_mode_var, value="region",
                                    command=self._on_print_screen_mode_changed)
            options.add_radiobutton(label="Print Screen captures the full screen",
                                    variable=self.prtsc_mode_var, value="full",
                                    command=self._on_print_screen_mode_changed)
            options.add_separator()
            options.add_command(label="Windows keyboard settings...",
                                command=open_windows_keyboard_settings)
            menubar.add_cascade(label="Options", menu=options, underline=0)

        self.root.config(menu=menubar)
        self.menubar, self.view_menu, self.edit_menu = menubar, view, edit

    def _on_region_side_changed(self):
        self._build_action_buttons()
        self._persist_config()   # remember straight away, not just on close

    def _build_ui(self):
        self.region_left_var = tk.BooleanVar(value=self.cfg.get("region_on_left", True))
        self.prtsc_var = tk.BooleanVar(value=IS_WINDOWS and bool(self.cfg.get("printscreen_hotkey", False)))
        prtsc_mode = self.cfg.get("printscreen_mode", "region")
        self.prtsc_mode_var = tk.StringVar(value=prtsc_mode if prtsc_mode in ("region", "full") else "region")
        self.quick_btn_var = tk.BooleanVar(value=bool(self.cfg.get("quick_button", True)))
        tool = self.cfg.get("draw_tool", "pen")
        self.tool_var = tk.StringVar(value=tool if tool in DRAW_TOOLS else "pen")
        self._build_menu()

        outer = ttk.Frame(self.root, padding=(20, 16, 20, 14))
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Snip & Save", style="Header.TLabel").pack(anchor="w", pady=(0, 10))

        # -------------------- settings card --------------------
        body = self._card(outer, pady=(0, 0))
        self._settings_wrap = body.master     # the card's outer frame, for hide/show while browsing slides

        ttk.Label(body, text="Save folder", style="FieldTitle.TLabel").pack(anchor="w")
        dir_row = ttk.Frame(body, style="Card.TFrame")
        dir_row.pack(fill="x", pady=(4, 10))
        self.dir_var = tk.StringVar(value=self.cfg["save_dir"])
        ttk.Entry(dir_row, textvariable=self.dir_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(dir_row, text="Browse...", style="Secondary.TButton", command=self.browse_dir).pack(
            side="right", fill="y"
        )

        ttk.Label(body, text="File name", style="FieldTitle.TLabel").pack(anchor="w")
        self.name_var = tk.StringVar(value=self.cfg["filename"])
        ttk.Entry(body, textvariable=self.name_var).pack(fill="x", pady=(4, 6))
        self.increment_var = tk.BooleanVar(value=self.cfg.get("auto_increment", True))
        ttk.Checkbutton(
            body, text="Auto-number to avoid overwriting  (1.name, 2.name, ...)",
            variable=self.increment_var, style="Card.TCheckbutton",
            command=self.refresh_file_list,
        ).pack(anchor="w", pady=(0, 10))

        # Format / Delay / Screen share one row
        opts = ttk.Frame(body, style="Card.TFrame")
        opts.pack(fill="x")
        opts.grid_columnconfigure(2, weight=1)

        ttk.Label(opts, text="Format", style="FieldTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(opts, text="Delay (sec)", style="FieldTitle.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Label(opts, text="Screen", style="FieldTitle.TLabel").grid(row=0, column=2, sticky="w", padx=(12, 0))

        self.format_var = tk.StringVar(value=self.cfg["format"])
        format_combo = ttk.Combobox(
            opts, textvariable=self.format_var, values=["PNG", "JPG"], width=6, state="readonly"
        )
        format_combo.grid(row=1, column=0, sticky="w", pady=(4, 0))
        format_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_file_list())

        self.delay_var = tk.IntVar(value=self.cfg["delay"])
        ttk.Combobox(
            opts, textvariable=self.delay_var, values=[0, 3, 5, 10], width=5, state="readonly"
        ).grid(row=1, column=1, sticky="w", pady=(4, 0), padx=(12, 0))

        screen_values = ["All Screens"] + [m["label"] for m in self.monitors]
        if len(self.monitors) >= 2:
            screen_values.append(OTHER_DISPLAY_LABEL)
        saved_screen_choice = self.cfg.get("screen_choice", "All Screens")
        if saved_screen_choice not in screen_values:
            saved_screen_choice = "All Screens"
        self.screen_var = tk.StringVar(value=saved_screen_choice)
        ttk.Combobox(
            opts, textvariable=self.screen_var, values=screen_values, width=18, state="readonly"
        ).grid(row=1, column=2, sticky="ew", pady=(4, 0), padx=(12, 0))
        self.screen_var.trace_add("write", lambda *a: self._refresh_hdr_status())
        if not self.monitors:
            ttk.Label(
                body, text='Install "screeninfo" to list individual monitors',
                style="CardMuted.TLabel",
            ).pack(anchor="w", pady=(6, 0))

        # HDR row: tone-map an HDR capture back to SDR so it isn't blown out
        hdr_row = ttk.Frame(body, style="Card.TFrame")
        hdr_row.pack(fill="x", pady=(10, 0))
        ttk.Label(hdr_row, text="HDR", style="FieldTitle.TLabel").pack(side="left")
        self.hdr_mode_var = tk.StringVar(value=self.cfg.get("hdr_mode", "Auto"))
        ttk.Combobox(
            hdr_row, textvariable=self.hdr_mode_var, values=["Auto", "On", "Off"],
            width=6, state="readonly",
        ).pack(side="left", padx=(8, 0))
        self.hdr_mode_var.trace_add("write", lambda *a: self._reapply_hdr())

        ttk.Label(hdr_row, text="Look", style="FieldTitle.TLabel").pack(side="left", padx=(12, 0))
        self.hdr_curve_var = tk.StringVar(value=self.cfg.get("hdr_curve", "Washed out"))
        ttk.Combobox(
            hdr_row, textvariable=self.hdr_curve_var, values=["Washed out", "Too bright"],
            width=11, state="readonly",
        ).pack(side="left", padx=(8, 0))
        self.hdr_curve_var.trace_add("write", lambda *a: self._reapply_hdr())

        ttk.Label(hdr_row, text="Strength", style="FieldTitle.TLabel").pack(side="left", padx=(12, 0))
        self.hdr_strength_var = tk.StringVar(value=str(self.cfg.get("hdr_strength", 100)))
        ttk.Spinbox(
            hdr_row, textvariable=self.hdr_strength_var, from_=0, to=100, increment=10,
            width=4, command=self._reapply_hdr,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(hdr_row, text="%", style="CardMuted.TLabel").pack(side="left", padx=(3, 0))
        self.hdr_strength_var.trace_add("write", lambda *a: self._reapply_hdr())

        self.hdr_status_var = tk.StringVar(value="")
        ttk.Label(body, textvariable=self.hdr_status_var, style="CardMuted.TLabel").pack(
            anchor="w", pady=(4, 0)
        )
        self._refresh_hdr_status()

        # refresh the existing-files panel whenever folder or filename changes
        self.dir_var.trace_add("write", lambda *a: self.refresh_file_list())
        self.name_var.trace_add("write", lambda *a: self.refresh_file_list())

        # -------------------- existing files panel --------------------
        files_body = self._card(outer, padding=(18, 12))
        self._files_wrap = files_body.master

        files_top = ttk.Frame(files_body, style="Card.TFrame")
        files_top.pack(fill="x")
        ttk.Label(files_top, text="Existing files for this name", style="FieldTitle.TLabel").pack(side="left")
        self.next_number_var = tk.StringVar(value="")
        ttk.Label(files_top, textvariable=self.next_number_var, style="NextNum.TLabel").pack(side="right")

        list_frame = tk.Frame(files_body, bg=CARD)
        list_frame.pack(fill="x", pady=(6, 0))
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox = tk.Listbox(
            list_frame, height=3, bg="#f5f6fa", fg=TEXT, bd=0, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
            selectbackground=ACCENT, selectforeground="#ffffff",
            font=(FONT_FAMILY, 9), activestyle="none", yscrollcommand=list_scroll.set,
        )
        list_scroll.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y", padx=(4, 0))

        # -------------------- drawing toolbar --------------------
        # One row when the window is wide enough:
        #   [Pen][Rectangle][Arrow]  Thickness --o-- 16 px  Color [\u25A0]   [Undo] [Clear all]
        # otherwise two, so nothing gets squeezed off the edge:
        #   [Pen][Rectangle][Arrow]                          [Undo] [Clear all]
        #   Thickness --o-- 16 px  Color [\u25A0]
        toolbar = self._card(outer, padding=(16, 9))
        self._toolbar_wrap = toolbar.master
        self._toolbar_body = toolbar
        self._toolbar_wide = None
        tools_row = self._tool_group = ttk.Frame(toolbar, style="Card.TFrame")
        style_row = self._style_group = ttk.Frame(toolbar, style="Card.TFrame")
        edit_row = self._edit_group = ttk.Frame(toolbar, style="Card.TFrame")

        ttk.Button(edit_row, text="\u21B6 Undo", style="Ghost.TButton", command=self.undo_stroke).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(edit_row, text="Clear all", style="Danger.TButton", command=self.clear_strokes).pack(
            side="left"
        )

        # Drawing tool: P / R / A on the keyboard switch too.
        for i, (value, label) in enumerate((("pen", "\u270E  Pen"), ("rect", "\u25AD  Rectangle"),
                                            ("arrow", "\u2197  Arrow"))):
            ttk.Radiobutton(
                tools_row, text=label, value=value, variable=self.tool_var,
                style="Tool.TRadiobutton", cursor="hand2", takefocus=False,
                command=self._on_tool_changed,
            ).pack(side="left", padx=(0 if i == 0 else 4, 0))

        ttk.Label(style_row, text="Thickness", style="FieldTitle.TLabel").pack(side="left")
        self.thickness_var = tk.IntVar(value=16)
        ttk.Scale(
            style_row, from_=1, to=30, orient="horizontal", variable=self.thickness_var, length=110,
            command=lambda v: self.thickness_var.set(int(float(v))),
        ).pack(side="left", padx=(8, 0))
        self.thickness_label = ttk.Label(style_row, text="16 px", style="Card.TLabel", width=5)
        self.thickness_label.pack(side="left", padx=(6, 0))
        self.thickness_var.trace_add(
            "write", lambda *a: self.thickness_label.configure(text=f"{self.thickness_var.get()} px")
        )

        ttk.Label(style_row, text="Color", style="FieldTitle.TLabel").pack(side="left", padx=(10, 0))
        self.color_btn = tk.Button(
            style_row, bg=self.color, activebackground=self.color, width=2, relief="flat", bd=0,
            cursor="hand2", command=self.pick_color,
            highlightbackground=TEXT, highlightcolor=TEXT, highlightthickness=2,
        )
        self.color_btn.pack(side="left", padx=(8, 0), ipady=1)

        self._layout_toolbar(wide=False)
        toolbar.bind("<Configure>", self._on_toolbar_resize)

        for key, tool in (("p", "pen"), ("r", "rect"), ("a", "arrow")):
            self.root.bind(f"<KeyPress-{key}>", lambda e, t=tool: self._on_tool_key(t))
            self.root.bind(f"<KeyPress-{key.upper()}>", lambda e, t=tool: self._on_tool_key(t))
        self.root.bind("<Control-z>", self._on_undo_key)
        self.root.bind("<Control-Z>", self._on_undo_key)

        # Bottom section is packed (side="bottom", bottom-up) BEFORE the
        # expanding preview, so the preview only ever gets leftover space and
        # the action buttons can never be squeezed off-screen.
        self.status_var = tk.StringVar(value="Ready.")
        self.status_label = ttk.Label(outer, textvariable=self.status_var, style="Status.TLabel", wraplength=460)
        self.status_label.pack(side="bottom", anchor="w", fill="x", pady=(6, 0))
        # keep the status text wrapping to the window's real width
        outer.bind("<Configure>", lambda e: self.status_label.configure(wraplength=max(200, e.width - 44)))

        footer = ttk.Frame(outer)
        footer.pack(side="bottom", fill="x", pady=(10, 0))
        # Lives on the footer row rather than as a grid row of its own, so it
        # costs the preview almost no height.
        self.import_btn = ttk.Button(
            footer, text="\U0001F4D1  Import Slides...", style="Ghost.TButton",
            command=self.import_slides, cursor="hand2",
        )
        self.import_btn.pack(side="right")
        # Re-copy what's in the preview (drawing included) at any time.
        self.copy_btn = ttk.Button(
            footer, text=self.COPY_LABEL, style="Ghost.TButton",
            command=self.copy_current, cursor="hand2",
        )
        self.copy_btn.pack(side="left")
        self.copy_btn.state(["disabled"])        # until there's something to copy
        self.root.bind("<Control-c>", self.copy_current)
        self.root.bind("<Control-C>", self.copy_current)   # Caps Lock on

        # -------------------- action grid --------------------
        # 2 x 2, equal-width columns:
        #     [ Capture Region  ] [ Capture Full Screen ]
        #     [ Save Screenshot ] [ Manual Save...      ]
        # Capture Region always sits directly above Save, and both swap
        # sides together via View > Put "Capture Region" and "Save" on the left.
        self.action_grid = ttk.Frame(outer)
        self.action_grid.pack(side="bottom", fill="x", pady=(12, 0))
        self.action_grid.grid_columnconfigure(0, weight=1, uniform="actions")
        self.action_grid.grid_columnconfigure(1, weight=1, uniform="actions")
        self._build_action_buttons()

        # -------------------- preview / edit canvas --------------------
        preview_wrap = tk.Frame(outer, bg=SUNKEN, highlightbackground=BORDER, highlightthickness=1)
        preview_wrap.pack(fill="both", expand=True, pady=(12, 0))
        self._preview_wrap = preview_wrap

        # Slide navigator: built now, only shown (just above the preview)
        # while an imported deck is being browsed, so it costs no space the
        # rest of the time.
        self.slide_nav = ttk.Frame(outer)
        self.prev_btn = ttk.Button(self.slide_nav, text="\u25C0  Prev", style="Ghost.TButton",
                                   command=lambda: self.step_slide(-1), cursor="hand2")
        self.prev_btn.pack(side="left")
        self.next_btn = ttk.Button(self.slide_nav, text="Next  \u25B6", style="Ghost.TButton",
                                   command=lambda: self.step_slide(1), cursor="hand2")
        self.next_btn.pack(side="right")
        self.done_btn = ttk.Button(self.slide_nav, text="Done", style="Ghost.TButton",
                                   command=self.close_slides, cursor="hand2")
        self.done_btn.pack(side="right", padx=(0, 6))
        self.settings_btn = ttk.Button(self.slide_nav, text="Show settings", style="Ghost.TButton",
                                       command=lambda: self._set_settings_visible(not self._settings_visible),
                                       cursor="hand2")
        self.settings_btn.pack(side="right", padx=(0, 6))
        self.slide_label = ttk.Label(self.slide_nav, text="", style="FieldTitle.TLabel", anchor="center",
                                     background=BG)
        self.slide_label.pack(side="left", fill="x", expand=True)
        self.preview_canvas = tk.Canvas(preview_wrap, bg=SUNKEN, highlightthickness=0, cursor="cross")
        self.preview_canvas.pack(fill="both", expand=True, padx=2, pady=2)
        self.preview_canvas.bind("<Configure>", self._redraw_preview)
        self.preview_canvas.bind("<ButtonPress-1>", self.on_press)
        self.preview_canvas.bind("<B1-Motion>", self.on_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.on_release)
        # Browse slides: mouse wheel over the preview, or arrow / page keys.
        self.preview_canvas.bind("<MouseWheel>", self._on_wheel)                 # Windows / macOS
        self.preview_canvas.bind("<Button-4>", lambda e: self.step_slide(-1))    # Linux wheel up
        self.preview_canvas.bind("<Button-5>", lambda e: self.step_slide(1))     # Linux wheel down
        self.preview_canvas.bind("<ButtonPress-1>", lambda e: self.preview_canvas.focus_set(), add="+")
        for key, delta in (("<Left>", -1), ("<Up>", -1), ("<Prior>", -1),
                           ("<Right>", 1), ("<Down>", 1), ("<Next>", 1)):
            self.root.bind(key, lambda e, d=delta: self._on_nav_key(d))
        self.root.bind("<Home>", lambda e: self._on_nav_key(None, to=0))
        self.root.bind("<End>", lambda e: self._on_nav_key(None, to=-1))


    TOOLBAR_GAP = 16   # px between the toolbar's groups when they share one row

    def _layout_toolbar(self, wide):
        """Arrange the drawing toolbar's three groups in one row or two."""
        if wide == self._toolbar_wide:
            return
        self._toolbar_wide = wide
        body = self._toolbar_body
        for g in (self._tool_group, self._style_group, self._edit_group):
            g.grid_forget()
        for col in range(3):
            body.grid_columnconfigure(col, weight=0)
        if wide:
            self._tool_group.grid(row=0, column=0, sticky="w")
            self._style_group.grid(row=0, column=1, sticky="w", padx=(self.TOOLBAR_GAP, 0))
            self._edit_group.grid(row=0, column=2, sticky="e", padx=(self.TOOLBAR_GAP, 0))
            body.grid_columnconfigure(1, weight=1)
        else:
            self._tool_group.grid(row=0, column=0, sticky="w")
            self._edit_group.grid(row=0, column=1, sticky="e", padx=(self.TOOLBAR_GAP, 0))
            self._style_group.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
            body.grid_columnconfigure(0, weight=1)

    def _on_toolbar_resize(self, event):
        need = (self._tool_group.winfo_reqwidth() + self._style_group.winfo_reqwidth()
                + self._edit_group.winfo_reqwidth() + 2 * self.TOOLBAR_GAP + 32)   # 32 = card padding
        self._layout_toolbar(event.width >= need)

    def _build_action_buttons(self):
        for w in self.action_grid.winfo_children():
            w.destroy()

        region_col = 0 if self.region_left_var.get() else 1
        other_col = 1 - region_col
        gap = 5  # half of the gutter between columns

        def padx(col):
            return (0, gap) if col == 0 else (gap, 0)

        # Top row: the two capture buttons. Bottom row: the two saves.
        ttk.Button(
            self.action_grid, text="\u2702  Capture Region", style="Region.TButton",
            command=lambda: self.start_capture("region"), cursor="hand2",
        ).grid(row=0, column=region_col, sticky="nsew", padx=padx(region_col), pady=(0, 10))

        ttk.Button(
            self.action_grid, text="\U0001F4F7  Capture Full Screen", style="Full.TButton",
            command=lambda: self.start_capture("full"), cursor="hand2",
        ).grid(row=0, column=other_col, sticky="nsew", padx=padx(other_col), pady=(0, 10))

        ttk.Button(
            self.action_grid, text="\U0001F4BE  Save Screenshot", style="Save.TButton",
            command=self.save_current, cursor="hand2",
        ).grid(row=1, column=region_col, sticky="nsew", padx=padx(region_col))

        ttk.Button(
            self.action_grid, text="\U0001F4DD  Manual Save...", style="Manual.TButton",
            command=self.save_current_manual, cursor="hand2",
        ).grid(row=1, column=other_col, sticky="nsew", padx=padx(other_col))

    def browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.dir_var.get() or os.path.expanduser("~"))
        if d:
            self.dir_var.set(d)

    def pick_color(self):
        _, hexcode = colorchooser.askcolor(color=self.color, title="Line color")
        if hexcode:
            self.color = hexcode
            self.color_btn.configure(bg=hexcode, activebackground=hexcode)

    # ================================================================= config
    def current_config(self):
        return {
            "save_dir": self.dir_var.get().strip() or DEFAULT_CONFIG["save_dir"],
            "filename": self.name_var.get().strip() or "screenshot",
            "format": self.format_var.get(),
            "delay": int(self.delay_var.get()),
            "auto_increment": bool(self.increment_var.get()),
            "region_on_left": bool(self.region_left_var.get()),
            "screen_choice": self.screen_var.get(),
            "hdr_mode": self.hdr_mode_var.get(),
            "hdr_curve": self.hdr_curve_var.get(),
            "hdr_strength": self._hdr_strength(),
            "slides_dir": self._slides_dir,
            "soffice_path": self._soffice_path,
            "printscreen_hotkey": bool(self.prtsc_var.get()),
            "printscreen_mode": self.prtsc_mode_var.get(),
            "printscreen_restore": self._prtsc_restore,
            "draw_tool": self.tool_var.get(),
            "quick_button": bool(self.quick_btn_var.get()),
            "quick_button_pos": self._quick_pos,
        }

    def _current_geometry(self):
        """The window's current "WxH+X+Y", or None when it shouldn't be
        recorded (minimised, hidden for a capture, maximised, or still tiny)."""
        try:
            if self.root.state() != "normal":
                return None
            g = self.root.geometry()
        except tk.TclError:
            return None
        m = GEOMETRY_RE.match(g)
        if not m or int(m.group(1)) < 300 or int(m.group(2)) < 300:
            return None
        return g

    def _persist_config(self):
        """Save current settings plus the window's size and position. When the
        window can't be measured right now (minimised etc.) the last good
        geometry is kept instead of overwriting it with junk."""
        cfg = self.current_config()
        geom = self._current_geometry()
        if geom:
            self._last_geometry = geom
        cfg["window_geometry"] = self._last_geometry
        err = save_config(cfg)
        if err and not self._config_warning_shown:
            self._config_warning_shown = True
            self.status_var.set(f"Couldn't save settings to {CONFIG_FILE}: {err}")
            messagebox.showwarning(
                "Settings not saved",
                f"Snip & Save couldn't write its settings file:\n{CONFIG_FILE}\n\n{err}\n\n"
                "Window size and options won't be remembered until this is fixed.",
            )

    def _on_root_configure(self, event):
        if event.widget is not self.root:   # ignore child-widget resizes
            return
        if self._geom_after is not None:
            self.root.after_cancel(self._geom_after)
        self._geom_after = self.root.after(700, self._save_geometry_now)

    def _save_geometry_now(self):
        self._geom_after = None
        self._persist_config()

    def _on_close(self):
        if self._geom_after is not None:
            self.root.after_cancel(self._geom_after)
            self._geom_after = None
        if getattr(self, "_pump_after", None) is not None:
            self.root.after_cancel(self._pump_after)
            self._pump_after = None
        self._persist_config()
        self._end_slide_session()
        if self.quick_btn is not None:
            self.quick_btn.destroy()
        self._prtsc.stop()      # (Windows' own switch stays off while the option is on)
        if self._grabber is not None:
            self._grabber.close()
            self._grabber = None
        self.root.destroy()

    def build_filepath(self, cfg):
        ext = "png" if cfg["format"] == "PNG" else "jpg"
        folder = cfg["save_dir"]
        os.makedirs(folder, exist_ok=True)
        start_n, base = parse_filename(cfg["filename"])
        if not cfg["auto_increment"]:
            if start_n is not None:
                return os.path.join(folder, f"{start_n}.{base}.{ext}")
            return os.path.join(folder, f"{base}.{ext}")
        n = self._next_number_for_ext(folder, base, ext, start_n)
        return os.path.join(folder, f"{n}.{base}.{ext}")

    # ================================================================= existing-files panel
    def _matching_files(self, folder, base):
        """Existing png/jpg files in folder belonging to this base name,
        as (number_or_None, filename), sorted for display."""
        if not base or not os.path.isdir(folder):
            return []
        pattern = re.compile(rf'^(?:(\d+)\.)?{re.escape(base)}\.(png|jpe?g)$', re.IGNORECASE)
        matches = []
        try:
            entries = os.listdir(folder)
        except OSError:
            return []
        for fname in entries:
            m = pattern.match(fname)
            if m:
                num = int(m.group(1)) if m.group(1) else None
                matches.append((num, fname))
        matches.sort(key=lambda t: (t[0] if t[0] is not None else 10**9, t[1].lower()))
        return matches

    def _next_number_for_ext(self, folder, base, ext, start_n):
        used = set()
        if os.path.isdir(folder):
            pattern = re.compile(rf'^(\d+)\.{re.escape(base)}\.{ext}$', re.IGNORECASE)
            try:
                entries = os.listdir(folder)
            except OSError:
                entries = []
            for fname in entries:
                m = pattern.match(fname)
                if m:
                    used.add(int(m.group(1)))
        n = start_n if start_n is not None else 1
        while n in used:
            n += 1
        return n

    def refresh_file_list(self, *args):
        if not hasattr(self, "file_listbox"):
            return
        cfg = self.current_config()
        folder = cfg["save_dir"]
        start_n, base = parse_filename(cfg["filename"])
        ext = "png" if cfg["format"] == "PNG" else "jpg"

        self.file_listbox.delete(0, tk.END)
        matches = self._matching_files(folder, base)
        if not matches:
            self.file_listbox.insert(tk.END, "  (no existing files for this name yet)")
        else:
            for _num, fname in matches:
                self.file_listbox.insert(tk.END, f"  {fname}")

        next_n = self._next_number_for_ext(folder, base, ext, start_n)
        self.next_number_var.set(f"Next number: {next_n}")

    # ================================================================= screen selection
    def _monitor_by_label(self, label):
        for m in self.monitors:
            if m["label"] == label:
                return m
        return None

    def _selected_region_bounds(self):
        """(x, y, width, height) of the chosen screen(s) in desktop
        coordinates, or None if no monitor info is available (screeninfo
        not installed) - in which case capture falls back to the whole
        virtual desktop / primary fullscreen behavior."""
        if not self.monitors:
            return None
        choice = self.screen_var.get()
        if choice == "All Screens":
            xs = [m["x"] for m in self.monitors]
            ys = [m["y"] for m in self.monitors]
            x2s = [m["x"] + m["width"] for m in self.monitors]
            y2s = [m["y"] + m["height"] for m in self.monitors]
            x, y = min(xs), min(ys)
            return (x, y, max(x2s) - x, max(y2s) - y)
        if choice == OTHER_DISPLAY_LABEL:
            return self._other_display_bounds()
        m = self._monitor_by_label(choice)
        if m:
            return (m["x"], m["y"], m["width"], m["height"])
        return None

    def _other_display_bounds(self):
        """Bounds of whichever monitor(s) the app window is NOT currently
        on. Uses the window position captured just before it was hidden
        for the capture (see start_capture); falls back to the window's
        live position if that wasn't recorded for some reason. With more
        than two monitors, this returns the union of every monitor except
        the one the window is on."""
        if not self.monitors:
            return None
        rect = self._window_rect_before_capture
        if rect:
            wx, wy, ww, wh = rect
        else:
            wx, wy = self.root.winfo_x(), self.root.winfo_y()
            ww, wh = self.root.winfo_width(), self.root.winfo_height()
        center_x, center_y = wx + ww / 2, wy + wh / 2
        cur_idx = monitor_index_containing(self.monitors, center_x, center_y)
        others = [m for i, m in enumerate(self.monitors) if i != cur_idx]
        if not others:
            others = self.monitors  # couldn't tell which is "current"; fall back to all
        xs = [m["x"] for m in others]
        ys = [m["y"] for m in others]
        x2s = [m["x"] + m["width"] for m in others]
        y2s = [m["y"] + m["height"] for m in others]
        x, y = min(xs), min(ys)
        return (x, y, max(x2s) - x, max(y2s) - y)

    def _get_capture_bbox(self):
        bounds = self._selected_region_bounds()
        if bounds is None:
            return None
        x, y, w, h = bounds
        return (x, y, x + w, y + h)

    # ================================================================= HDR
    def _hdr_strength(self):
        try:
            return max(0, min(100, int(float(self.hdr_strength_var.get()))))
        except (ValueError, tk.TclError):
            return 100

    def _hdr_for_selection(self):
        """(hdr_enabled, sdr_white_nits_or_None) for the screen(s) about to
        be captured. Falls back to 'any connected display has HDR on' when
        we can't tie the choice to one specific monitor."""
        if not self.hdr_displays:
            return False, None
        choice = self.screen_var.get()
        target_names = []
        if choice not in ("All Screens", OTHER_DISPLAY_LABEL):
            m = self._monitor_by_label(choice)
            if m and m.get("name"):
                target_names = [m["name"].upper()]
        if not target_names:
            target_names = list(self.hdr_displays.keys())
        on = [self.hdr_displays[n] for n in target_names
              if n in self.hdr_displays and self.hdr_displays[n]["enabled"]]
        if not on:
            return False, None
        nits = [d["sdr_white_nits"] for d in on if d["sdr_white_nits"]]
        return True, (max(nits) if nits else None)

    def _refresh_hdr_status(self):
        """One line under the settings card describing the HDR situation."""
        if platform.system() != "Windows":
            self.hdr_status_var.set("HDR detection is Windows-only; leave this Off elsewhere.")
            return
        if not self._hdr_checked:
            self.hdr_status_var.set("Checking the displays for HDR...")
            return
        if not self.hdr_displays:
            self.hdr_status_var.set("No HDR info from Windows (older build, or no HDR-capable display).")
            return
        enabled, nits = self._hdr_for_selection()
        if not enabled:
            self.hdr_status_var.set("HDR is off on the selected display; no correction needed.")
            return
        self._detected_nits = nits or 200.0
        detail = f", SDR white {nits:.0f} nits" if nits else ""
        self.hdr_status_var.set(
            f"HDR is on for this display{detail}. Captures are tone-mapped to SDR."
        )

    def _start_hdr_probe(self, announce=False):
        """Ask Windows which displays have HDR on, on a background thread.
        The answer is picked up by _take_hdr_probe (or, with announce=True,
        shown on the HDR line as soon as it arrives)."""
        if not IS_WINDOWS:
            return
        probe = {"done": threading.Event(), "result": {}}

        def run():
            probe["result"] = windows_hdr_displays()
            probe["done"].set()
            if announce:
                self._call_on_ui(self._take_hdr_probe, 0)

        self._hdr_probe = probe
        threading.Thread(target=run, daemon=True, name="hdr-probe").start()

    def _take_hdr_probe(self, wait=0.0):
        """Use the latest probe's answer if it's in (waiting up to `wait` s)."""
        probe = self._hdr_probe
        if probe is None or not probe["done"].wait(wait):
            return
        self._hdr_probe = None
        self.hdr_displays = probe["result"]
        self._hdr_checked = True
        self._refresh_hdr_status()

    def _hdr_should_apply(self):
        mode = self.hdr_mode_var.get()
        if mode == "Off":
            return False
        if mode == "On":
            return True
        enabled, _ = self._hdr_for_selection()
        return enabled

    def _apply_hdr(self, raw):
        """Return the image to actually work with, plus a note for the status line."""
        if raw is None:
            return None, ""
        if not self._hdr_should_apply():
            return raw, ""
        label = self.hdr_curve_var.get()
        curve = "PQ" if label == "Washed out" else "linear"
        strength = self._hdr_strength()
        if strength == 0:
            return raw, ""
        try:
            corrected = apply_hdr_correction(
                raw, curve=curve, sdr_white_nits=getattr(self, "_detected_nits", 200.0),
                strength=strength / 100.0,
            )
        except Exception as e:
            print("HDR correction failed:", e)
            return raw, " (HDR correction failed; showing the raw capture.)"
        return corrected, f" HDR tone-mapped ({label}, {strength}%) - set HDR to Off if it looks wrong."

    def _reapply_hdr(self, *args):
        """Re-derive the working image from the untouched capture. Strokes
        are stored in image coordinates, so they survive this."""
        self._refresh_hdr_status()
        if self.raw_img is None:
            return
        self.working_img, note = self._apply_hdr(self.raw_img)
        self._hdr_note = note
        self._redraw_preview()

    # ================================================================= capture flow
    def start_capture(self, mode, delay=None, from_hotkey=False, t_pressed=None, trigger=None):
        # Nothing slow happens between the click and the grab any more:
        #  - settings are saved after the capture, not before it;
        #  - no worker thread + 40 ms UI-queue hop: root.after() runs the
        #    capture right here once the delay (if any) is up;
        #  - instead of hoping the window has gone after a fixed wait, the
        #    grab waits exactly one screen refresh (see _before_grab).
        # Print Screen passes delay=0 (you're already looking at what you
        # want) and the moment the key went down, so its timings count from
        # the key press.
        if delay is None:
            try:
                delay = max(0, int(self.delay_var.get()))
            except (ValueError, tk.TclError):
                delay = 0
        self._from_hotkey = from_hotkey
        self._trigger_word = trigger or ("key press" if from_hotkey else "click")
        self._state_before_capture = self.root.state()   # "iconic" when minimised
        self._ensure_capture_ready()   # also switches off the hide/show animation
        self._window_rect_before_capture = self._window_rect_for_capture()
        if self.quick_btn is not None:
            self.quick_btn.hide()   # the floating button mustn't be in the shot either
        self.root.withdraw()  # hide our own window so it isn't in the shot
        self.root.update()
        start = t_pressed if t_pressed is not None else time.perf_counter()
        self._t_click = start + delay   # timings exclude the chosen delay
        target = self._do_full_capture if mode == "full" else self._start_region_select
        self.root.after(delay * 1000 if delay else HIDE_SETTLE_MS, target)

    def _window_rect_for_capture(self):
        """(x, y, w, h) of this window, for "Other Display". While it's
        minimised (a Print Screen capture) its live position is meaningless,
        so use where it last sat on screen."""
        try:
            state = self.root.state()
        except tk.TclError:
            state = "normal"
        if state in ("normal", "zoomed"):
            return (self.root.winfo_x(), self.root.winfo_y(),
                    self.root.winfo_width(), self.root.winfo_height())
        m = GEOMETRY_RE.match(self._last_geometry or "")
        if m and m.group(3) is not None:
            return (_geometry_offset(m.group(3)), _geometry_offset(m.group(4)),
                    int(m.group(1)), int(m.group(2)))
        return None

    def _before_grab(self):
        """Make sure the just-hidden window is really off the screen."""
        wait_for_screen_refresh()

    def _ms_since_click(self):
        if self._t_click is None:
            return 0
        return max(0, round((time.perf_counter() - self._t_click) * 1000))

    def _grab_screen(self, bbox=None):
        """Screenshot of bbox (desktop coordinates), or of everything if None."""
        self._last_grab_engine = "Pillow"
        if self._grabber is not None:
            try:
                img = self._grabber.grab(bbox)
                self._last_grab_engine = "mss"
                return img
            except Exception as e:
                print("mss grab failed; using Pillow from now on:", e)
                self._grabber.close()
                self._grabber = None
                self._grabber_failed = True
        try:
            try:
                return ImageGrab.grab(bbox=bbox, all_screens=True)
            except TypeError:
                return ImageGrab.grab(bbox=bbox)
        except Exception as e:
            if platform.system() == "Linux":
                tmp = "/tmp/_snip_tmp.png"
                for tool in (["scrot", tmp], ["gnome-screenshot", "-f", tmp]):
                    try:
                        subprocess.run(tool, check=True)
                        img = Image.open(tmp)
                        if bbox:
                            img = img.crop(bbox)
                        return img
                    except Exception:
                        continue
            raise e

    def _do_full_capture(self):
        self._before_grab()
        self._start_hdr_probe()        # answers while the screen is being grabbed
        bbox = self._get_capture_bbox()
        t_grab = time.perf_counter()
        try:
            img = self._grab_screen(bbox=bbox)
        except Exception as e:
            self._on_error(str(e))
            self.root.deiconify()
            self._restore_quick_button()
            return
        grab_ms = round((time.perf_counter() - t_grab) * 1000)
        self._timing_note = (f" [captured {self._ms_since_click()} ms after the {self._trigger_word}; "
                             f"screen grab {grab_ms} ms via {self._last_grab_engine}]")
        self._load_into_editor(img)

    def _start_region_select(self):
        self._before_grab()
        bounds = self._selected_region_bounds()
        bbox = None if bounds is None else (bounds[0], bounds[1],
                                            bounds[0] + bounds[2], bounds[1] + bounds[3])
        t_grab = time.perf_counter()
        try:
            full_img = self._grab_screen(bbox=bbox)
        except Exception as e:
            self._on_error(str(e))
            self.root.deiconify()
            self._restore_quick_button()
            return
        grab_ms = round((time.perf_counter() - t_grab) * 1000)
        engine = self._last_grab_engine

        overlay = tk.Toplevel()
        overlay.attributes("-topmost", True)
        overlay.configure(bg="black")

        if bounds:
            x, y, w, h = bounds
            overlay.overrideredirect(True)
            overlay.geometry(f"{w}x{h}+{x}+{y}")
        else:
            overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.35)

        canvas = tk.Canvas(overlay, cursor="cross", bg="black", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        # Built before the overlay is first drawn so it appears complete in
        # one go rather than the hint popping in a moment later.
        hint = tk.Label(
            overlay, text="Drag to select an area  \u2022  Esc to cancel",
            bg=SLATE, fg="#ffffff", font=(FONT_FAMILY, 11, "bold"), padx=14, pady=8,
        )
        hint.place(relx=0.5, y=24, anchor="n")

        # Scale correction: the overlay's realized (logical) size can differ
        # from the screenshot's physical pixel size on HiDPI / scaled
        # displays. Without this, the cropped region drifts from what was
        # dragged.
        overlay.update_idletasks()
        logical_w = overlay.winfo_width() or full_img.width
        logical_h = overlay.winfo_height() or full_img.height
        scale_x = full_img.width / logical_w
        scale_y = full_img.height / logical_h

        self._timing_note = (f" [overlay up {self._ms_since_click()} ms after the {self._trigger_word}; "
                             f"screen grab {grab_ms} ms via {engine}]")
        # Ask about HDR now: it finishes while you drag.
        self._start_hdr_probe()

        coords = {}

        def on_press(event):
            coords["x0"], coords["y0"] = event.x, event.y
            coords["rect"] = canvas.create_rectangle(
                event.x, event.y, event.x, event.y,
                outline="#ff0000", width=16,
            )

        def on_drag(event):
            if "rect" in coords:
                canvas.coords(coords["rect"], coords["x0"], coords["y0"], event.x, event.y)

        def on_release(event):
            if "x0" not in coords:     # a release left over from a click before the overlay
                return
            x0, y0 = coords["x0"], coords["y0"]
            x1, y1 = event.x, event.y
            overlay.destroy()
            left, right = sorted((x0, x1))
            top, bottom = sorted((y0, y1))
            if right - left < 2 or bottom - top < 2:
                self._on_error("Selection too small; capture cancelled.")
                self.root.deiconify()
                self._restore_quick_button()
                return
            left_p = round(left * scale_x)
            right_p = round(right * scale_x)
            top_p = round(top * scale_y)
            bottom_p = round(bottom * scale_y)
            cropped = full_img.crop((left_p, top_p, right_p, bottom_p))
            self._load_into_editor(cropped)

        def on_escape(event):
            overlay.destroy()
            self.status_var.set("Region capture cancelled.")
            self._return_after_cancel()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", on_escape)
        if self._from_hotkey:
            # Another program was active when Print Screen was pressed: take
            # the keyboard as soon as the overlay is on screen so Esc works.
            overlay.bind("<Map>", lambda e: force_foreground(overlay) if e.widget is overlay else None)
        overlay.focus_force()

    def _load_into_editor(self, img):
        self._end_slide_session()
        self.raw_img = img
        # HDR can be toggled between captures; the probe started at the grab
        # has normally answered by now (it's only waited on if it hasn't).
        self._take_hdr_probe(wait=1.0)
        self._refresh_hdr_status()
        self.working_img, self._hdr_note = self._apply_hdr(img)
        self.strokes = []
        # Draw the preview while the window is still hidden, so it comes back
        # already showing the new capture.
        drawn = self._redraw_preview()
        self.root.deiconify()
        self._restore_quick_button()
        if self._from_hotkey:
            self._from_hotkey = False
            self._bring_to_front()
        self._update_copy_state()
        if not drawn:
            self.root.after_idle(self._redraw_preview)
        timing = self._timing_note if SHOW_TIMINGS else ""
        self._timing_note = ""
        hdr_note = self._hdr_note
        self._copy_with_status(
            self.working_img,
            lambda clip: (f"Draw on the screenshot below, then Save Screenshot whenever you're ready. "
                          f"{clip}{hdr_note}{timing}"),
            done="Copied to clipboard.",
            failed="(Clipboard copy unavailable - see script notes.)",
        )
        # Settings used to be written between the click and the grab; now
        # it happens once the capture is on screen.
        self.root.after(300, self._persist_config)

    def _copy_with_status(self, img, compose, done, failed, then=None):
        """Copy img to the clipboard in the background. The status line says
        "Copying to clipboard..." and then the outcome, unless something else
        has replaced the message in the meantime. then(ok), if given, runs
        afterwards on the UI thread."""
        pending = compose("Copying to clipboard...")
        self.status_var.set(pending)

        def finished(ok):
            if self.status_var.get() == pending:
                self.status_var.set(compose(done if ok else failed))
            if then is not None:
                then(ok)

        self._clipboard.copy(img, finished)

    def copy_current(self, event=None):
        """Copy what's in the preview - drawing included - to the clipboard
        again. Also bound to Ctrl+C, which is left alone in text fields."""
        if event is not None and self._text_selected_in_focus():
            return None                  # the field copies its own selected text
        if self.working_img is None:
            self.status_var.set("Nothing to copy yet - capture a screenshot first.")
            return "break"
        marked = bool(self.strokes)
        img = self._bake_strokes(self.working_img, self.strokes) if marked else self.working_img
        what = f"Slide {self._slide_idx + 1}" if self._in_slide_mode() else "Screenshot"
        if marked:
            what += " (with your drawing)"
        self._copy_with_status(
            img, lambda clip: clip,
            done=f"{what} copied to the clipboard.",
            failed="Couldn't copy to the clipboard - see the clipboard notes at the top of the script.",
            then=self._flash_copy_button,
        )
        return "break"

    def _text_selected_in_focus(self):
        """True if the keyboard focus is in a text field with text selected,
        so Ctrl+C belongs to that text. With nothing selected (say you just
        typed a file name), Ctrl+C copies the screenshot instead."""
        try:
            focus = self.root.focus_get()
        except (tk.TclError, KeyError):   # focus is in a native dialog etc.
            return False
        if focus is None:
            return False
        cls = focus.winfo_class()
        try:
            if cls == "Text":
                return bool(focus.tag_ranges("sel"))
            if cls in ("TEntry", "Entry", "TSpinbox", "Spinbox", "TCombobox"):
                return bool(focus.tk.getboolean(focus.tk.call(focus._w, "selection", "present")))
        except tk.TclError:
            return cls in ("TEntry", "Entry", "TSpinbox", "Spinbox", "Text")
        return False

    def _flash_copy_button(self, ok):
        """Brief "Copied" on the button itself - the status line is easy to miss."""
        if not ok:
            return
        if self._copy_flash_after is not None:
            self.root.after_cancel(self._copy_flash_after)

        def restore():
            self._copy_flash_after = None
            self.copy_btn.configure(text=self.COPY_LABEL)

        self.copy_btn.configure(text="\u2713  Copied")
        self._copy_flash_after = self.root.after(1500, restore)

    def _update_copy_state(self):
        """Copy to Clipboard (button and Edit menu) is usable once there's an image."""
        ready = self.working_img is not None
        self.copy_btn.state(["!disabled"] if ready else ["disabled"])
        try:
            self.edit_menu.entryconfigure(0, state="normal" if ready else "disabled")
        except tk.TclError:
            pass

    # ================================================================= preview / drawing
    def _draw_preview_placeholder(self):
        c = self.preview_canvas
        c.delete("all")
        w = c.winfo_width() or 400
        h = c.winfo_height() or 220
        pad = 14
        c.create_rectangle(pad, pad, w - pad, h - pad, outline=BORDER_STRONG, dash=(6, 4), width=2)
        c.create_text(w / 2, h / 2 - 18, text="\U0001F5BC", font=(FONT_FAMILY, 30), fill=MUTED)
        c.create_text(
            w / 2, h / 2 + 18, text="No screenshot yet",
            font=(FONT_FAMILY, 11, "bold"), fill=TEXT,
        )
        c.create_text(
            w / 2, h / 2 + 40, text="Capture one to see a preview here",
            font=(FONT_FAMILY, 9), fill=MUTED,
        )

    def _redraw_preview(self, event=None):
        """Redraw the preview. Returns False if the canvas has no size yet."""
        c = self.preview_canvas
        w = c.winfo_width()
        h = c.winfo_height()
        if w < 10 or h < 10:
            return False
        if self.working_img is None:
            self._draw_preview_placeholder()
            return True

        pad = 16
        avail_w = max(1, w - pad * 2)
        avail_h = max(1, h - pad * 2)
        img = self.working_img
        scale = min(avail_w / img.width, avail_h / img.height)
        disp_w = max(1, int(img.width * scale))
        disp_h = max(1, int(img.height * scale))
        offset_x = (w - disp_w) / 2
        offset_y = (h - disp_h) / 2

        self._img_scale = scale
        self._img_offset = (offset_x, offset_y)

        self._preview_photo = self._preview_image(img, disp_w, disp_h)
        c.delete("all")
        c.create_rectangle(
            offset_x - 1, offset_y - 1, offset_x + disp_w, offset_y + disp_h,
            outline=BORDER_STRONG, width=1,
        )
        c.create_image(offset_x, offset_y, anchor="nw", image=self._preview_photo)

        for stroke in self.strokes:
            self._draw_stroke_on_canvas(stroke)
        return True

    def _draw_stroke_on_canvas(self, stroke, tags=()):
        """Draw one annotation (pen line, rectangle or arrow) on the preview,
        mapped from image pixels to the preview's scale and position."""
        pts = stroke["points"]
        if len(pts) < 2:
            return
        c = self.preview_canvas
        scale = self._img_scale
        offset_x, offset_y = self._img_offset
        sw = max(1, stroke["width"] * scale)
        color = stroke["color"]
        kind = stroke.get("kind", "pen")

        def to_canvas(p):
            return p[0] * scale + offset_x, p[1] * scale + offset_y

        if kind == "rect":
            (x0, y0), (x1, y1) = to_canvas(pts[0]), to_canvas(pts[-1])
            c.create_rectangle(x0, y0, x1, y1, outline=color, width=sw, tags=tags)
        elif kind == "arrow":
            shape = arrow_shape(*pts[0], *pts[-1], stroke["width"])
            if shape is None:
                return
            shaft_end, head = shape
            c.create_line(*to_canvas(pts[0]), *to_canvas(shaft_end), width=sw, fill=color,
                          capstyle="round", tags=tags)
            flat = [v for p in head for v in to_canvas(p)]
            c.create_polygon(*flat, fill=color, outline=color, tags=tags)
        else:
            flat = [v for p in pts for v in to_canvas(p)]
            c.create_line(*flat, width=sw, fill=color, capstyle="round", joinstyle="round",
                          smooth=True, tags=tags)

    def _preview_image(self, img, disp_w, disp_h):
        """The screenshot scaled to the preview, as a Tk image. Scaling a big
        capture is the slow part of a redraw, so the result is reused until
        the image or the preview size changes (undo, clear and every stroke
        redraw then cost almost nothing). reducing_gap shrinks in a fast
        integer step first and finishes with Lanczos - visually the same,
        several times quicker on multi-monitor captures."""
        cache = self._preview_cache
        if cache is not None and cache[0] is img and cache[1] == (disp_w, disp_h):
            return cache[2]
        try:
            resized = img.resize((disp_w, disp_h), Image.LANCZOS, reducing_gap=3.0)
        except TypeError:          # Pillow older than 7.0
            resized = img.resize((disp_w, disp_h), Image.LANCZOS)
        photo = ImageTk.PhotoImage(resized)
        self._preview_cache = (img, (disp_w, disp_h), photo)
        return photo

    def _canvas_to_orig(self, event):
        ox = (event.x - self._img_offset[0]) / self._img_scale
        oy = (event.y - self._img_offset[1]) / self._img_scale
        return ox, oy

    def on_press(self, event):
        if self.working_img is None:
            return
        ox, oy = self._canvas_to_orig(event)
        tool = self.tool_var.get()
        self._current_stroke = {
            "kind": tool,
            "points": [(ox, oy)] if tool == "pen" else [(ox, oy), (ox, oy)],
            "width": self.thickness_var.get() / self._img_scale,
            "color": self.color,
        }
        self._last_canvas_pt = (event.x, event.y)

    def on_drag(self, event):
        stroke = self._current_stroke
        if stroke is None:
            return
        ox, oy = self._canvas_to_orig(event)
        kind = stroke.get("kind", "pen")
        if kind == "pen":
            stroke["points"].append((ox, oy))
            lx, ly = self._last_canvas_pt
            self.preview_canvas.create_line(
                lx, ly, event.x, event.y,
                width=max(1, self.thickness_var.get()),
                fill=self.color, capstyle="round", joinstyle="round", smooth=True,
            )
        else:
            # Rectangle / arrow: rubber-band from the press point to the pointer.
            x0, y0 = stroke["points"][0]
            if event.state & 0x0001:            # Shift: square / 45-degree arrow
                ox, oy = constrain_shape(kind, x0, y0, ox, oy)
            stroke["points"][1] = (ox, oy)
            self.preview_canvas.delete("live_shape")
            self._draw_stroke_on_canvas(stroke, tags=("live_shape",))
        self._last_canvas_pt = (event.x, event.y)

    def on_release(self, event):
        stroke = self._current_stroke
        if stroke:
            if stroke.get("kind", "pen") == "pen":
                if len(stroke["points"]) > 1:
                    self.strokes.append(stroke)
            else:
                (x0, y0), (x1, y1) = stroke["points"]
                # A click without a real drag would leave a dot; ignore it.
                if math.hypot(x1 - x0, y1 - y0) * self._img_scale >= 4:
                    self.strokes.append(stroke)
                else:
                    self.preview_canvas.delete("live_shape")
            self.preview_canvas.dtag("live_shape", "live_shape")
        self._current_stroke = None
        self._last_canvas_pt = None

    TOOL_HINTS = {
        "pen": "Pen: draw freehand on the screenshot.",
        "rect": "Rectangle: drag on the screenshot to box something. Hold Shift for a square.",
        "arrow": "Arrow: drag from where it starts to what it points at. Hold Shift to snap to 45°.",
    }

    def _on_tool_changed(self):
        self.status_var.set(self.TOOL_HINTS.get(self.tool_var.get(), ""))
        self._persist_config()

    def _typing_in_field(self):
        """Keyboard focus is in a text box / dropdown, so letter keys are typing."""
        try:
            focus = self.root.focus_get()
        except (tk.TclError, KeyError):
            return True     # a native dialog has the keyboard
        return focus is not None and focus.winfo_class() in (
            "TEntry", "Entry", "TCombobox", "TSpinbox", "Spinbox", "Text", "Listbox"
        )

    def _on_tool_key(self, tool):
        """P / R / A pick Pen / Rectangle / Arrow (not while typing a file name)."""
        if self._typing_in_field() or self._current_stroke is not None:
            return
        if self.tool_var.get() != tool:
            self.tool_var.set(tool)
            self._on_tool_changed()

    def _on_undo_key(self, event=None):
        if self._typing_in_field():
            return None
        self.undo_stroke()
        return "break"

    def undo_stroke(self):
        if self.strokes:
            self.strokes.pop()
            self._redraw_preview()

    def clear_strokes(self):
        if self.strokes:
            self.strokes = []
            self._redraw_preview()

    def _bake_strokes(self, img, strokes):
        """Return a new full-resolution image with all strokes drawn onto it."""
        result = img.convert("RGB") if img.mode not in ("RGB", "RGBA") else img.copy()
        draw = ImageDraw.Draw(result)
        for stroke in strokes:
            pts = stroke["points"]
            if len(pts) < 2:
                continue
            width_px = max(1, round(stroke["width"]))
            color = stroke["color"]
            r = width_px / 2
            kind = stroke.get("kind", "pen")
            if kind == "rect":
                (x0, y0), (x1, y1) = pts[0], pts[-1]
                x0, x1 = sorted((x0, x1))
                y0, y1 = sorted((y0, y1))
                # Pillow draws the outline inside the box; push it out by half
                # the width so it sits centred on the dragged edge, as in the preview.
                draw.rectangle([round(x0 - r), round(y0 - r), round(x1 + r), round(y1 + r)],
                               outline=color, width=width_px)
            elif kind == "arrow":
                shape = arrow_shape(*pts[0], *pts[-1], stroke["width"])
                if shape is None:
                    continue
                shaft_end, head = shape
                (sx, sy) = pts[0]
                draw.line([pts[0], shaft_end], fill=color, width=width_px)
                draw.ellipse([sx - r, sy - r, sx + r, sy + r], fill=color)   # round tail
                draw.polygon(head, fill=color)
            else:
                draw.line(pts, fill=color, width=width_px, joint="curve")
                for (x, y) in pts:
                    draw.ellipse([x - r, y - r, x + r, y + r], fill=color)
        return result

    # ================================================================= save
    def save_current(self):
        if self.working_img is None:
            messagebox.showinfo("Nothing to save", "Capture a screenshot first.")
            return
        cfg = self.current_config()
        self._persist_config()
        final_img = self._bake_strokes(self.working_img, self.strokes) if self.strokes else self.working_img
        path = self.build_filepath(cfg)
        try:
            if cfg["format"] == "JPG":
                if final_img.mode in ("RGBA", "P"):
                    final_img = final_img.convert("RGB")
                final_img.save(path, "JPEG", quality=92)
            else:
                final_img.save(path, "PNG")
        except Exception as e:
            self._on_error(str(e))
            return
        # Re-copy to clipboard at save time too: if strokes were drawn after
        # the initial capture-time copy, the clipboard would otherwise still
        # hold the un-annotated version instead of what was just saved.
        self._copy_with_status(
            final_img,
            lambda clip: f"Saved: {path}  (keep drawing and Save again anytime). {clip}",
            done="Clipboard updated.",
            failed="(Clipboard copy unavailable.)",
        )
        self.refresh_file_list()

    def save_current_manual(self):
        """Let the user type/pick an exact filename (and location) instead
        of using the auto-numbered name from the settings card. Bypasses
        auto-increment entirely - if the chosen file already exists, the
        standard OS overwrite prompt (built into asksaveasfilename) applies."""
        if self.working_img is None:
            messagebox.showinfo("Nothing to save", "Capture a screenshot first.")
            return
        cfg = self.current_config()
        self._persist_config()
        ext = "png" if cfg["format"] == "PNG" else "jpg"
        _start_n, base = parse_filename(cfg["filename"])
        png_type = ("PNG image", "*.png")
        jpg_type = ("JPEG image", "*.jpg *.jpeg")
        # Whichever format is selected in the settings card goes first, since
        # the dialog's type filter always defaults to the first entry in this
        # list - listing it first is what actually makes it the pre-selected
        # (and thus matching) filter, not just defaultextension on its own.
        ordered_types = [png_type, jpg_type] if cfg["format"] == "PNG" else [jpg_type, png_type]
        path = filedialog.asksaveasfilename(
            title="Manual Save Screenshot",
            initialdir=cfg["save_dir"] if os.path.isdir(cfg["save_dir"]) else os.path.expanduser("~"),
            initialfile=f"{base}.{ext}",
            defaultextension=f".{ext}",
            filetypes=ordered_types + [("All files", "*.*")],
        )
        if not path:
            self.status_var.set("Manual save cancelled.")
            return
        final_img = self._bake_strokes(self.working_img, self.strokes) if self.strokes else self.working_img
        save_ext = os.path.splitext(path)[1].lower()
        try:
            if save_ext in (".jpg", ".jpeg"):
                if final_img.mode in ("RGBA", "P"):
                    final_img = final_img.convert("RGB")
                final_img.save(path, "JPEG", quality=92)
            else:
                final_img.save(path, "PNG")
        except Exception as e:
            self._on_error(str(e))
            return
        self._copy_with_status(
            final_img,
            lambda clip: f"Saved: {path}  (keep drawing and Save again anytime). {clip}",
            done="Clipboard updated.",
            failed="(Clipboard copy unavailable.)",
        )
        self.refresh_file_list()

    # ================================================================= slide import
    def import_slides(self):
        """Pick a presentation and save every visible slide as an image,
        using the same folder / name / numbering / format as screenshots."""
        if self._slide_export_running:
            return
        # Check for a converter up front, so nobody picks a deck only to be
        # told afterwards that nothing can open it.
        if not self._ensure_slide_converter():
            return
        start_dir = self._slides_dir if os.path.isdir(self._slides_dir) else os.path.expanduser("~")
        pres_path = filedialog.askopenfilename(
            title="Choose a presentation to export as images",
            initialdir=start_dir, filetypes=SLIDE_FILETYPES,
        )
        if not pres_path:
            return
        self._slides_dir = os.path.dirname(pres_path)

        # Snapshot settings on the main thread (Tk variables aren't safe to
        # read from the worker). Numbering is forced on: N slides can't all
        # share one filename.
        cfg = self.current_config()
        cfg["auto_increment"] = True
        self._persist_config()

        self._slide_export_running = True
        if self.import_btn is not None:
            self.import_btn.state(["disabled"])
        self.status_var.set(f"Opening {os.path.basename(pres_path)}...")
        threading.Thread(
            target=self._slide_export_thread, args=(pres_path, cfg, self._soffice_path), daemon=True
        ).start()

    def _ensure_slide_converter(self):
        """True when PowerPoint or LibreOffice can be used. If neither is
        found, explain and offer to download LibreOffice or point to it."""
        ppt_ok, ppt_why = powerpoint_available()
        if ppt_ok or _find_soffice(self._soffice_path):
            return True
        choice = self._libreoffice_needed_dialog(ppt_why)
        if choice == "download":
            webbrowser.open(LIBREOFFICE_DOWNLOAD_URL)
            self.status_var.set("Once LibreOffice is installed, click Import Slides again "
                                "(no need to restart this app).")
            return False
        if choice != "locate":
            return False
        if platform.system() == "Windows":
            types = [("LibreOffice (soffice.exe)", "soffice.exe soffice.com"), ("Programs", "*.exe")]
            start = os.environ.get("ProgramW6432") or os.environ.get("PROGRAMFILES") or "C:\\"
        elif platform.system() == "Darwin":
            types = [("LibreOffice.app or soffice", "*")]
            start = "/Applications"
        else:
            types = [("soffice", "soffice*"), ("All files", "*")]
            start = "/usr"
        picked = filedialog.askopenfilename(
            title="Locate LibreOffice's soffice program", initialdir=start, filetypes=types,
        )
        if not picked:
            return False
        exe = _normalise_soffice_choice(picked)
        name = os.path.basename(exe or "").lower()
        if not (exe and os.path.isfile(exe) and name.startswith(("soffice", "libreoffice"))):
            messagebox.showerror(
                "Not LibreOffice",
                "That doesn't look like LibreOffice's soffice program.\n\n"
                "It's normally in LibreOffice's \"program\" folder, e.g.\n"
                "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            )
            return False
        self._soffice_path = exe
        self._persist_config()
        return True

    def _libreoffice_needed_dialog(self, ppt_why):
        """Small modal: returns "download", "locate" or None (cancel)."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Import Slides")
        dlg.configure(bg=CARD)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        body = ttk.Frame(dlg, style="Card.TFrame", padding=(22, 18))
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Importing slides needs PowerPoint or LibreOffice",
                  style="FieldTitle.TLabel").pack(anchor="w")
        found = (f"{ppt_why}.\nLibreOffice wasn't found either.\n\n"
                 if platform.system() == "Windows" else "LibreOffice wasn't found on this computer.\n\n")
        ttk.Label(
            body, style="Card.TLabel", justify="left", wraplength=400,
            text=found + "LibreOffice is free, and it's all this needs - no other installs. "
                         "If it's already installed somewhere unusual, point to it once and "
                         "it'll be remembered.",
        ).pack(anchor="w", pady=(8, 16))

        choice = {"value": None}

        def pick(value):
            choice["value"] = value
            dlg.destroy()

        row = ttk.Frame(body, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Button(row, text="Cancel", style="Ghost.TButton", command=lambda: pick(None),
                   cursor="hand2").pack(side="right")
        ttk.Button(row, text="Locate LibreOffice...", style="Secondary.TButton",
                   command=lambda: pick("locate"), cursor="hand2").pack(side="right", padx=(0, 6))
        download = ttk.Button(row, text="Download LibreOffice", style="Primary.TButton",
                              command=lambda: pick("download"), cursor="hand2")
        download.pack(side="right", padx=(0, 6))
        dlg.bind("<Escape>", lambda e: pick(None))
        dlg.protocol("WM_DELETE_WINDOW", lambda: pick(None))

        # Centre over the main window.
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        download.focus_set()
        try:
            dlg.grab_set()
        except tk.TclError:
            pass
        self.root.wait_window(dlg)
        return choice["value"]

    def _slide_status(self, text):
        """Status-line update that's safe to call from the worker thread."""
        self._call_on_ui(self.status_var.set, text)

    def _slide_export_thread(self, pres_path, cfg, soffice_hint):
        tmp = tempfile.mkdtemp(prefix="snip_slides_")
        saved = []
        keep_tmp = False
        try:
            pngs, engine = render_presentation_to_pngs(
                pres_path, tmp, SLIDE_EXPORT_WIDTH,
                progress=lambda done, total: self._slide_status(
                    f"Rendering slide {done} of {total}..."),
                status=self._slide_status,
                soffice_hint=soffice_hint,
            )
            if not pngs:
                raise RuntimeError("No slides to export (is every slide hidden?).")

            for i, png in enumerate(pngs, 1):
                with Image.open(png) as im:
                    im.load()
                    # build_filepath re-scans the folder each time, so every
                    # slide gets the next free number, gaps and all.
                    path = self.build_filepath(cfg)
                    write_image(im, path, cfg["format"])
                saved.append(path)
                self._slide_status(f"Saving slide {i} of {len(pngs)}...")
            # Hand the temp folder to the UI thread: the lossless PNGs in it
            # are what the preview browses.
            keep_tmp = True
            self._call_on_ui(self._on_slides_done, pres_path, saved, pngs, tmp, engine)
        except Exception as e:
            self._call_on_ui(self._on_slides_failed, str(e), saved)
        finally:
            if not keep_tmp:
                shutil.rmtree(tmp, ignore_errors=True)

    def _finish_slide_export(self):
        self._slide_export_running = False
        if self.import_btn is not None:
            self.import_btn.state(["!disabled"])
        self.refresh_file_list()

    def _on_slides_done(self, pres_path, saved, pngs, tmp, engine=""):
        self._finish_slide_export()
        self._end_slide_session()           # drop a previously browsed deck
        self._slide_paths = list(pngs)
        self._slide_names = [os.path.basename(p) for p in saved]
        self._slide_tmp = tmp
        self._slide_strokes = {}
        self.slide_nav.pack(fill="x", pady=(10, 0), before=self._preview_wrap)
        # Give the slides the room: the settings cards aren't needed while
        # flipping through (Show settings brings them back).
        self._set_settings_visible(False)
        self._show_slide(0)
        n = len(saved)
        first, last = self._slide_names[0], self._slide_names[-1]
        span = first if n == 1 else f"{first} ... {last}"
        self.status_var.set(
            f"Exported {n} slide{'s' if n != 1 else ''} from {os.path.basename(pres_path)}"
            f"{f' with {engine}' if engine else ''}: "
            f"{span} in {os.path.dirname(saved[0])}. Browse with Prev/Next, the arrow keys or the "
            f"mouse wheel; draw on any slide and Save for a marked-up copy."
        )

    def _on_slides_failed(self, msg, saved):
        self._finish_slide_export()
        if saved:
            msg += f"\n\n{len(saved)} slide(s) were saved before the error."
        self.status_var.set("Slide export failed - see popup.")
        messagebox.showerror("Slide export error", msg)

    # ------------------------------------------------------------ slide browsing
    def _in_slide_mode(self):
        return bool(self._slide_paths)

    def _show_slide(self, idx):
        """Load slide idx into the preview, remembering the current slide's drawing."""
        if not self._slide_paths:
            return
        idx = max(0, min(idx, len(self._slide_paths) - 1))
        if self.working_img is not None:
            self._slide_strokes[self._slide_idx] = self.strokes   # keep what was drawn
        try:
            with Image.open(self._slide_paths[idx]) as im:
                im.load()
                img = im.copy()
        except Exception as e:
            self._on_error(f"Couldn't load slide {idx + 1}: {e}")
            return
        self._slide_idx = idx
        # Slides are rendered files, not screen captures, so they're outside
        # the HDR pipeline: clearing raw_img stops an HDR toggle from pulling
        # the last screenshot back over the slide.
        self.raw_img = None
        self._hdr_note = ""
        self.working_img = img
        self.strokes = self._slide_strokes.get(idx, [])
        self._current_stroke = None
        self._redraw_preview()
        self._update_copy_state()
        total = len(self._slide_paths)
        self.slide_label.configure(text=f"Slide {idx + 1} of {total}")
        self.prev_btn.state(["!disabled"] if idx > 0 else ["disabled"])
        self.next_btn.state(["!disabled"] if idx < total - 1 else ["disabled"])
        if self._slide_names:
            self.status_var.set(
                f"Slide {idx + 1} of {total}  (exported as {self._slide_names[idx]}).  "
                f"Draw on it and Save Screenshot for a marked-up copy."
            )

    def step_slide(self, delta):
        if self._in_slide_mode():
            self._show_slide(self._slide_idx + delta)

    def _on_wheel(self, event):
        if not self._in_slide_mode():
            return
        # Windows/macOS report a signed delta; scrolling down = next slide.
        self.step_slide(1 if event.delta < 0 else -1)

    def _on_nav_key(self, delta, to=None):
        if not self._in_slide_mode():
            return
        # Don't hijack the arrow keys while typing in a field or a dropdown.
        focus = self.root.focus_get()
        if focus is not None and focus.winfo_class() in (
            "TEntry", "Entry", "TCombobox", "Listbox", "TScale", "TSpinbox", "Spinbox"
        ):
            return
        if to is None:
            self.step_slide(delta)
        else:
            self._show_slide(len(self._slide_paths) - 1 if to == -1 else to)

    def _set_settings_visible(self, visible):
        """Show/hide the settings and existing-files cards (their widgets and
        values stay alive, so Save keeps using the current settings)."""
        if visible == self._settings_visible:
            return
        if visible:
            self._settings_wrap.pack(fill="x", pady=(0, 0), before=self._toolbar_wrap)
            self._files_wrap.pack(fill="x", pady=(12, 0), before=self._toolbar_wrap)
        else:
            self._settings_wrap.pack_forget()
            self._files_wrap.pack_forget()
        self._settings_visible = visible
        self.settings_btn.configure(text="Hide settings" if visible else "Show settings")

    def close_slides(self):
        """Leave slide browsing. The slide on screen stays loaded as an
        ordinary image, so it can still be annotated and saved."""
        self._end_slide_session()
        self.status_var.set("Closed the slide browser. The current slide is still loaded in the preview.")

    def _end_slide_session(self):
        """Tear down a slide session: hide the navigator and delete the temp PNGs."""
        was_active = bool(self._slide_paths) or self._slide_tmp is not None
        self._slide_paths = []
        self._slide_names = []
        self._slide_strokes = {}
        self._slide_idx = 0
        if self.slide_nav.winfo_ismapped():
            self.slide_nav.pack_forget()
        self._set_settings_visible(True)
        if self._slide_tmp:
            shutil.rmtree(self._slide_tmp, ignore_errors=True)
            self._slide_tmp = None
        return was_active

    # ================================================================= quick-capture button
    def quick_capture(self, mode):
        """The floating button was clicked: capture straight away, ignoring
        Delay (like Print Screen), and bring the window up with the result.
        Esc in the overlay while the window was minimised leaves it minimised."""
        try:
            if self.root.state() == "withdrawn" or self.root.grab_current() is not None:
                return     # a capture is already under way, or one of our dialogs is open
        except tk.TclError:
            return
        if self._native_dialog_open():
            self.show_main_window()
            self.status_var.set("Close the open dialog first, then capture again.")
            return
        self.start_capture(mode, delay=0, from_hotkey=True, trigger="click")

    def show_main_window(self):
        self.root.deiconify()
        self._bring_to_front()

    def _restore_quick_button(self):
        """Put the floating button back after a capture (it hides for the grab)."""
        if self.quick_btn is not None and self.quick_btn_var.get():
            self.quick_btn.show()

    def _on_quick_button_toggled(self):
        if self.quick_btn_var.get():
            self.quick_btn.show()
            self.status_var.set("Quick-capture button is on: click it any time to capture a region "
                                "(right-click for more, drag to move).")
        else:
            self.quick_btn.hide()
            self.status_var.set("Quick-capture button hidden (View menu brings it back).")
        self._persist_config()

    def hide_quick_button(self):
        self.quick_btn_var.set(False)
        self._on_quick_button_toggled()

    def quick_button_moved(self, x, y):
        self._quick_pos = f"{x},{y}"
        self._persist_config()

    def reset_quick_button(self):
        self._quick_pos = ""
        self.quick_btn.place("", self.monitors)
        self.quick_btn_var.set(True)
        self.quick_btn.show()
        self._persist_config()
        self.status_var.set("Quick-capture button is back at the top of the main screen.")

    # ================================================================= Print Screen key
    def _on_print_screen_toggled(self):
        if self.prtsc_var.get():
            self._enable_print_screen(user_initiated=True)
        else:
            self._disable_print_screen()
        self._persist_config()

    def _on_print_screen_mode_changed(self):
        self._persist_config()
        if self.prtsc_var.get():
            what = "a region" if self.prtsc_mode_var.get() == "region" else "the full screen"
            self.status_var.set(f"Print Screen now captures {what}.")

    def _enable_print_screen(self, user_initiated):
        """Switch off Windows' own Print Screen -> screen capture (remembering
        how it was) and claim the key."""
        self._prtsc_user_initiated = user_initiated
        self._prtsc_waiting = False
        state = windows_snip_key_state()
        if state in ("on", "unset"):
            if not self._prtsc_restore:          # keep the original from the first time
                self._prtsc_restore = state
            try:
                set_windows_snip_key_state("off")
            except OSError as e:
                print("Couldn't change Windows' Print Screen setting:", e)
        if user_initiated:
            self.status_var.set("Taking over the Print Screen key...")
        self._prtsc.start()

    def _disable_print_screen(self):
        """Release the key and put Windows' own setting back the way it was."""
        self._prtsc.stop()
        self._prtsc_waiting = False
        self._prtsc_user_initiated = False
        restore, self._prtsc_restore = self._prtsc_restore, ""
        handed_back = False
        if restore in ("on", "unset"):
            try:
                set_windows_snip_key_state(restore)
                handed_back = True
            except OSError as e:
                print("Couldn't restore Windows' Print Screen setting:", e)
        self.status_var.set("Print Screen is handed back to Windows." if handed_back
                            else "Print Screen no longer opens Snip & Save.")

    def _on_print_screen(self, t_pressed):
        """Print Screen was pressed, in any program. Ignored while a capture
        is already under way or one of this app's dialogs is open."""
        if not self.prtsc_var.get():
            return
        try:
            if self.root.state() == "withdrawn" or self.root.grab_current() is not None:
                return
        except tk.TclError:
            return
        if self._native_dialog_open():
            self.status_var.set("Close the open dialog first, then press Print Screen again.")
            return
        mode = "full" if self.prtsc_mode_var.get() == "full" else "region"
        self.start_capture(mode, delay=0, from_hotkey=True, t_pressed=t_pressed)

    def _on_print_screen_status(self, ok, error_code):
        if not self.prtsc_var.get():
            return                                # switched off in the meantime
        what = "a region" if self.prtsc_mode_var.get() == "region" else "the full screen"
        if ok:
            announce = self._prtsc_user_initiated or self._prtsc_waiting
            self._prtsc_user_initiated = self._prtsc_waiting = False
            if announce:
                self.status_var.set(f"Print Screen now opens Snip & Save and captures {what}. "
                                    f"Keep the app running for this - minimised is fine.")
            elif self.status_var.get() == "Ready.":
                self.status_var.set(f"Ready. Press Print Screen to capture {what}.")
            return
        self._prtsc_waiting = True
        self.status_var.set("Print Screen is still taken by Windows or another program; Snip & Save "
                            "will take it over as soon as it's free (Options > Windows keyboard settings...).")
        if self._prtsc_user_initiated:
            self._prtsc_user_initiated = False
            # Shown from its own callback so the background-call pump keeps
            # running (clipboard results etc.) while the dialog is open.
            self.root.after(0, self._print_screen_busy_dialog, error_code)

    def _native_dialog_open(self):
        """A file / colour / message dialog is open: Windows disables the
        main window while one is up (Tk's own dialogs use a grab instead)."""
        hwnd = _toplevel_hwnd(self.root) if IS_WINDOWS else None
        if not hwnd:
            return False
        try:
            import ctypes
            return not ctypes.windll.user32.IsWindowEnabled(ctypes.c_void_p(hwnd))
        except Exception:
            return False

    def _print_screen_busy_dialog(self, error_code):
        detail = ("" if error_code in (0, PrintScreenHotkey.ERROR_HOTKEY_ALREADY_REGISTERED)
                  else f"\n\n(Windows error {error_code}.)")
        choice = self._ask(
            "Print Screen key", "Print Screen is still taken",
            "Windows' own screenshot shortcut, or another program, is still holding the Print "
            "Screen key.\n\n"
            "In Windows Settings > Accessibility > Keyboard, turn off \"Use the Print screen key "
            "to open screen capture\" (older builds say \"screen snipping\"). If it already shows "
            "Off, switch it On and back Off, or sign out and back in.\n\n"
            "Screenshot tools such as OneDrive, Dropbox, Greenshot, ShareX and Lightshot can claim "
            "the key too - turn off their Print Screen option.\n\n"
            "Snip & Save keeps trying in the background and takes the key over as soon as it's "
            "free." + detail,
            [("Open Windows Settings", "settings", "Primary.TButton"), ("OK", None, "Ghost.TButton")],
        )
        if choice == "settings":
            open_windows_keyboard_settings()

    def _ask(self, title, heading, text, buttons):
        """Small modal in the app's style. buttons: [(label, value, style), ...]
        laid out left to right; the first one has the focus (Enter picks it).
        Returns the value of the button clicked, or None (Esc / closed)."""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg=CARD)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        body = ttk.Frame(dlg, style="Card.TFrame", padding=(22, 18))
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=heading, style="FieldTitle.TLabel").pack(anchor="w")
        ttk.Label(body, text=text, style="Card.TLabel", justify="left",
                  wraplength=420).pack(anchor="w", pady=(8, 16))

        choice = {"value": None}

        def pick(value):
            choice["value"] = value
            dlg.destroy()

        row = ttk.Frame(body, style="Card.TFrame")
        row.pack(fill="x")
        first = None
        for label, value, style in reversed(buttons):   # packed from the right
            b = ttk.Button(row, text=label, style=style, cursor="hand2",
                           command=lambda v=value: pick(v))
            b.pack(side="right", padx=(6, 0))
            first = b
        dlg.bind("<Escape>", lambda e: pick(None))
        if buttons:
            dlg.bind("<Return>", lambda e: pick(buttons[0][1]))
        dlg.protocol("WM_DELETE_WINDOW", lambda: pick(None))

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_reqwidth()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_reqheight()) // 3
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        if first is not None:
            first.focus_set()
        try:
            dlg.grab_set()
        except tk.TclError:
            pass
        self.root.wait_window(dlg)
        return choice["value"]

    def _bring_to_front(self):
        """After a Print Screen capture another program was usually active;
        pull the window on top so the capture is right there."""
        root = self.root
        try:
            root.lift()
            root.attributes("-topmost", True)
        except tk.TclError:
            return

        def settle():
            try:
                root.attributes("-topmost", False)
            except tk.TclError:
                pass

        root.after(100, settle)
        force_foreground(root)

    def _return_after_cancel(self):
        """Back to how things were before a capture that was called off: a
        Print Screen pressed while the window was minimised leaves it minimised."""
        if self._from_hotkey and self._state_before_capture == "iconic":
            self.root.iconify()
        else:
            self.root.deiconify()
        self._from_hotkey = False
        self._restore_quick_button()

    def _call_on_ui(self, fn, *args):
        """Run fn(*args) on the Tk thread. Safe to call from any thread."""
        self._ui_calls.put((fn, args))

    def _pump_ui_calls(self):
        try:
            while True:
                fn, args = self._ui_calls.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    self.root.report_callback_exception(*sys.exc_info())
        except queue.Empty:
            pass
        try:
            self._pump_after = self.root.after(40, self._pump_ui_calls)
        except tk.TclError:
            pass   # window is closing

    def _on_error(self, msg):
        self.status_var.set("Error - see popup.")
        messagebox.showerror("Screenshot error", msg)


def main():
    enable_dpi_awareness()
    root = tk.Tk()
    apply_style(root)
    SnippingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
