"""Read cell values and optionally formatting from an Excel range."""

import argparse
import json
import sys
import os
from datetime import datetime, date

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet,
    rgb_tuple_to_hex, output_json, IS_WINDOWS
)


def read_cell_formats(sheet, start_row, start_col, end_row, end_col):
    """Read formatting info for cells with non-default formatting."""
    formats = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            try:
                cell = sheet.range(row, col)
                fmt = {}

                # Background color
                bg = rgb_tuple_to_hex(cell.color)
                if bg:
                    fmt["bg"] = bg

                # Font properties
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

                # Number format
                try:
                    nf = cell.number_format
                    if nf and nf != "General":
                        fmt["numberFormat"] = nf
                except Exception:
                    pass

                # Borders (platform-specific)
                try:
                    borders = _read_borders(cell)
                    if borders:
                        fmt["borders"] = borders
                except Exception:
                    pass

                # Alignment (platform-specific)
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
    """Read border info from a cell."""
    borders = {}
    # border indices: left=7, top=8, bottom=9, right=10
    names_indices = [("top", 8), ("bottom", 9), ("left", 7), ("right", 10)]
    style_map = {
        1: "thin", -4138: "medium", 4: "thick",
        -4119: "double", -4118: "dotted", -4115: "dashed"
    }

    for name, idx in names_indices:
        try:
            if IS_WINDOWS:
                border = cell.api.Borders(idx)
                ls = border.LineStyle
            else:
                border = cell.api.borders[idx]
                ls = border.line_style()

            if ls is not None and ls != -4142 and ls != 0:
                borders[name] = style_map.get(ls, "thin")
        except Exception:
            continue

    return borders if borders else None


def _read_alignment(cell):
    """Read alignment info from a cell."""
    result = {}
    h_map = {-4131: "left", -4108: "center", -4152: "right"}
    v_map = {-4160: "top", -4108: "middle", -4107: "bottom"}

    try:
        if IS_WINDOWS:
            h = cell.api.HorizontalAlignment
            v = cell.api.VerticalAlignment
        else:
            h = cell.api.horizontal_alignment()
            v = cell.api.vertical_alignment()

        if h in h_map:
            result["textAlign"] = h_map[h]
        if v in v_map and v_map[v] != "bottom":  # bottom is default
            result["verticalAlign"] = v_map[v]
    except Exception:
        pass

    return result if result else None


def clean_value(val):
    """Clean cell value for JSON serialization."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.isoformat()
    if isinstance(val, float) and val == int(val):
        return int(val)
    if isinstance(val, str):
        return val.replace('\x00', '').strip()
    return val


def read_cells(workbook, cell_range, sheet=None, include_formats=False):
    app, err = get_app()
    if err:
        return {"error": err}

    wb, err = get_workbook(app, workbook)
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
        # Determine row/col bounds from range
        top_left = rng[0, 0]
        bottom_right = rng[-1, -1] if rng.shape[0] > 1 or rng.shape[1] > 1 else top_left
        result["formats"] = read_cell_formats(
            ws,
            top_left.row, top_left.column,
            bottom_right.row, bottom_right.column
        )

    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', required=True)
    parser.add_argument('--range', required=True)
    parser.add_argument('--sheet', default=None)
    parser.add_argument('--formats', action='store_true')
    args = parser.parse_args()

    result = read_cells(args.workbook, args.range, args.sheet, args.formats)
    output_json(result)


if __name__ == "__main__":
    main()
