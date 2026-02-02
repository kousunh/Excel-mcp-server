"""Shared utilities for cross-platform Excel MCP scripts."""

import sys
import json
import os

IS_WINDOWS = sys.platform == 'win32'
IS_MAC = sys.platform == 'darwin'


# ---------------------------------------------------------------------------
# xlwings helpers
# ---------------------------------------------------------------------------

def get_app():
    """Get active Excel application."""
    import xlwings as xw
    try:
        app = xw.apps.active
    except Exception as e:
        return None, f"Cannot connect to Excel: {e}"
    if app is None:
        return None, "Excel is not running"
    return app, None


def get_or_create_app():
    """Get active Excel app or launch a new one."""
    import xlwings as xw
    try:
        app = xw.apps.active
        if app is not None:
            return app, False, None
    except Exception:
        pass
    try:
        app = xw.App(visible=True)
        return app, True, None
    except Exception as e:
        return None, False, f"Cannot launch Excel: {e}"


def get_workbook(app, name=None):
    """Get workbook by name or active workbook."""
    if name:
        for book in app.books:
            if book.name == name or book.fullname == name:
                return book, None
        return None, f"Workbook '{name}' not found"
    wb = app.books.active
    if wb is None:
        return None, "No active workbook"
    return wb, None


def open_path(path):
    """Open a file path in Excel via xlwings.
    If already open, returns that workbook.
    Returns (wb, was_already_open, error)."""
    import xlwings as xw
    abspath = os.path.abspath(path)

    # Check if already open in any Excel instance
    try:
        for app in xw.apps:
            for book in app.books:
                try:
                    if os.path.abspath(book.fullname) == abspath:
                        return book, True, None
                except Exception:
                    continue
    except Exception:
        pass

    # Not open - get or create Excel app and open the file
    if not os.path.exists(abspath):
        return None, False, f"File not found: {abspath}"

    app, _, err = get_or_create_app()
    if err:
        return None, False, err

    try:
        wb = app.books.open(abspath)
        return wb, False, None
    except Exception as e:
        return None, False, f"Cannot open file: {e}"


def get_sheet(wb, name=None):
    """Get sheet by name or active sheet."""
    if name:
        try:
            return wb.sheets[name], None
        except Exception:
            return None, f"Sheet '{name}' not found"
    return wb.sheets.active, None


def hex_to_rgb_int(hex_color):
    """Convert hex color (#RRGGBB) to Excel RGB integer (BGR for Windows COM)."""
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)


def rgb_tuple_to_hex(rgb):
    """Convert RGB tuple (r, g, b) to hex string."""
    if rgb is None:
        return None
    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    return "#{:02x}{:02x}{:02x}".format(r, g, b)


def output_json(result):
    """Print result as JSON with proper encoding."""
    def json_serial(obj):
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        if isinstance(obj, bytes):
            return obj.decode('utf-8', errors='replace')
        raise TypeError(f"Type {type(obj)} not serializable")

    import io
    if IS_WINDOWS:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    output = json.dumps(result, ensure_ascii=False, default=json_serial)
    try:
        print(output)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')


def set_performance_mode(app, enable):
    """Toggle Excel performance mode."""
    if not enable:
        return None
    try:
        if IS_WINDOWS:
            orig_screen = app.api.Application.ScreenUpdating
            orig_calc = app.api.Application.Calculation
            app.api.Application.ScreenUpdating = False
            app.api.Application.Calculation = -4135  # xlCalculationManual
            return (orig_screen, orig_calc)
        else:
            app.screen_updating = False
            return (True,)
    except Exception:
        return None


def restore_performance_mode(app, settings):
    """Restore Excel performance settings."""
    if settings is None:
        return
    try:
        if IS_WINDOWS:
            app.api.Application.ScreenUpdating = settings[0]
            app.api.Application.Calculation = settings[1]
        else:
            app.screen_updating = True
    except Exception:
        pass
