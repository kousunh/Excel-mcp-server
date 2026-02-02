"""Write values to cells in an Excel workbook (live or file)."""

import argparse
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet,
    open_file_writable, get_sheet_openpyxl,
    set_performance_mode, restore_performance_mode, output_json
)


# ---------------------------------------------------------------------------
# xlwings (live Excel)
# ---------------------------------------------------------------------------

def _write_live(workbook, cell_range, value, sheet):
    app, err = get_app()
    if err:
        return {"error": err}
    wb, err = get_workbook(app, workbook)
    if err:
        return {"error": err}
    ws, err = get_sheet(wb, sheet)
    if err:
        return {"error": err}

    perf = set_performance_mode(app, True)
    try:
        try:
            rng = ws.range(cell_range)
        except Exception as e:
            return {"error": f"Invalid range '{cell_range}': {e}"}

        rows, cols = rng.shape
        value = _reshape(value, rows, cols)
        rng.value = value

        return {
            "success": True,
            "workbook": wb.name,
            "sheet": ws.name,
            "range": cell_range,
            "size": f"{rows}x{cols}"
        }
    except Exception as e:
        return {"error": f"Failed to write: {e}"}
    finally:
        restore_performance_mode(app, perf)


# ---------------------------------------------------------------------------
# openpyxl (file-based)
# ---------------------------------------------------------------------------

def _write_file(path, cell_range, value, sheet):
    wb, err = open_file_writable(path)
    if err:
        return {"error": err}

    ws, err = get_sheet_openpyxl(wb, sheet)
    if err:
        wb.close()
        return {"error": err}

    try:
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except Exception as e:
        wb.close()
        return {"error": f"Invalid range '{cell_range}': {e}"}

    rows = max_row - min_row + 1
    cols = max_col - min_col + 1

    # Normalize value to 2D list
    if not isinstance(value, list):
        data = [[value]]
    elif value and not isinstance(value[0], list):
        if rows > 1 and cols == 1:
            data = [[v] for v in value]
        else:
            data = [value]
    else:
        data = value

    try:
        for r_idx, row_data in enumerate(data):
            for c_idx, val in enumerate(row_data):
                ws.cell(row=min_row + r_idx, column=min_col + c_idx, value=val)
        wb.save(path)
    except Exception as e:
        wb.close()
        return {"error": f"Failed to write: {e}"}

    wb.close()
    return {
        "success": True,
        "path": path,
        "sheet": ws.title,
        "range": cell_range,
        "size": f"{rows}x{cols}"
    }


# ---------------------------------------------------------------------------
# shared
# ---------------------------------------------------------------------------

def _reshape(value, rows, cols):
    """Adjust value shape for a target range."""
    if not isinstance(value, list):
        return value
    if value and not isinstance(value[0], list):
        if rows > 1 and cols == 1:
            return [[item] for item in value[:rows]]
        elif rows == 1 and cols > 1:
            return value[:cols]
    elif value and isinstance(value[0], list):
        return [row[:cols] for row in value[:rows]]
    return value


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', default=None)
    parser.add_argument('--path', default=None)
    parser.add_argument('--range', required=True)
    parser.add_argument('--value', required=True)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    if not args.workbook and not args.path:
        output_json({"error": "Either --workbook or --path is required"})
        return

    try:
        value = json.loads(args.value)
    except (json.JSONDecodeError, ValueError):
        value = args.value

    if args.path:
        result = _write_file(args.path, args.range, value, args.sheet)
    else:
        result = _write_live(args.workbook, args.range, value, args.sheet)

    output_json(result)


if __name__ == "__main__":
    main()
