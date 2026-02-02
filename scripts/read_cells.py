"""Read cell values and optionally formatting from an Excel range."""

import argparse
import sys
import os
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet, open_path,
    rgb_tuple_to_hex, output_json, IS_WINDOWS
)


def clean_value(val):
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.isoformat()
    if isinstance(val, float) and val == int(val):
        return int(val)
    if isinstance(val, str):
        return val.replace('\x00', '').strip()
    return val


def _get_wb(workbook=None, path=None):
    """Resolve workbook from either name or path. Returns (wb, was_opened, error)."""
    if path:
        return open_path(path)
    app, err = get_app()
    if err:
        return None, False, err
    wb, err = get_workbook(app, workbook)
    return wb, False, err


def read_cells(workbook=None, path=None, cell_range=None, sheet=None, include_formats=False):
    wb, was_opened, err = _get_wb(workbook, path)
    if err:
        return {"error": err}

    ws, err = get_sheet(wb, sheet)
    if err:
        return {"error": err}

    try:
        rng = ws.range(cell_range)
    except Exception as e:
        return {"error": f"Invalid range '{cell_range}': {e}"}

    # Read values
    raw = rng.value
    if rng.shape[0] == 1 and rng.shape[1] == 1:
        values = [[clean_value(raw)]]
    elif rng.shape[0] == 1:
        values = [[clean_value(v) for v in raw]]
    elif rng.shape[1] == 1:
        values = [[clean_value(v)] for v in raw]
    else:
        values = [[clean_value(v) for v in row] for row in raw]

    result = {
        "workbook": wb.name,
        "sheet": ws.name,
        "range": cell_range,
        "values": values
    }

    if include_formats:
        top_left = rng[0, 0]
        bottom_right = rng[-1, -1] if rng.shape[0] > 1 or rng.shape[1] > 1 else top_left
        result["formats"] = _read_formats(
            ws, top_left.row, top_left.column,
            bottom_right.row, bottom_right.column
        )

    # Close if we opened it via path
    if was_opened:
        try:
            wb.close()
        except Exception:
            pass

    return result


def _read_formats(sheet, r1, c1, r2, c2):
    formats = []
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            try:
                cell = sheet.range(row, col)
                fmt = {}
                bg = rgb_tuple_to_hex(cell.color)
                if bg:
                    fmt["bg"] = bg
                try:
                    if cell.font.bold:
                        fmt["bold"] = True
                    if cell.font.italic:
                        fmt["italic"] = True
                    if cell.font.size:
                        fmt["fontSize"] = cell.font.size
                    if cell.font.name:
                        fmt["fontName"] = cell.font.name
                    fc = rgb_tuple_to_hex(cell.font.color)
                    if fc and fc != "#000000":
                        fmt["fontColor"] = fc
                except Exception:
                    pass
                try:
                    nf = cell.number_format
                    if nf and nf != "General":
                        fmt["numberFormat"] = nf
                except Exception:
                    pass
                try:
                    borders = _read_borders(cell)
                    if borders:
                        fmt["borders"] = borders
                except Exception:
                    pass
                try:
                    align = _read_alignment(cell)
                    if align:
                        fmt.update(align)
                except Exception:
                    pass
                if fmt:
                    fmt["cell"] = cell.address.replace('$', '')
                    formats.append(fmt)
            except Exception:
                continue
    return formats


def _read_borders(cell):
    borders = {}
    names_indices = [("top", 8), ("bottom", 9), ("left", 7), ("right", 10)]
    style_map = {1: "thin", -4138: "medium", 4: "thick", -4119: "double", -4118: "dotted", -4115: "dashed"}
    for name, idx in names_indices:
        try:
            if IS_WINDOWS:
                ls = cell.api.Borders(idx).LineStyle
            else:
                ls = cell.api.borders[idx].line_style()
            if ls is not None and ls != -4142 and ls != 0:
                borders[name] = style_map.get(ls, "thin")
        except Exception:
            continue
    return borders or None


def _read_alignment(cell):
    result = {}
    h_map = {-4131: "left", -4108: "center", -4152: "right"}
    v_map = {-4160: "top", -4108: "middle", -4107: "bottom"}
    try:
        if IS_WINDOWS:
            h, v = cell.api.HorizontalAlignment, cell.api.VerticalAlignment
        else:
            h, v = cell.api.horizontal_alignment(), cell.api.vertical_alignment()
        if h in h_map:
            result["textAlign"] = h_map[h]
        if v in v_map and v_map[v] != "bottom":
            result["verticalAlign"] = v_map[v]
    except Exception:
        pass
    return result or None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', default=None)
    parser.add_argument('--path', default=None)
    parser.add_argument('--range', required=True)
    parser.add_argument('--sheet', default=None)
    parser.add_argument('--formats', action='store_true')
    args = parser.parse_args()

    if not args.workbook and not args.path:
        output_json({"error": "Either --workbook or --path is required"})
        return

    result = read_cells(args.workbook, args.path, args.range, args.sheet, args.formats)
    output_json(result)


if __name__ == "__main__":
    main()
