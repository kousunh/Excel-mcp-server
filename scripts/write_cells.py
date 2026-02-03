"""Write values to cells in an Excel workbook (live or file)."""

import argparse
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet,
    set_performance_mode, restore_performance_mode, output_json
)


# ---------------------------------------------------------------------------
# xlwings (live Excel / workbook mode)
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

    perf = set_performance_mode(wb.app, True)
    try:
        try:
            rng = ws.range(cell_range)
        except Exception as e:
            return {"error": f"Invalid range '{cell_range}': {e}"}

        rows, cols = rng.shape
        value = _reshape(value, rows, cols)
        rng.value = value

        wb.save()

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
        restore_performance_mode(wb.app, perf)


# ---------------------------------------------------------------------------
# xlsx_io (file-based, pure Python ZIP/XML, no Excel needed)
# ---------------------------------------------------------------------------

def _write_file(path, cell_range, value, sheet):
    from xlsx_io import XlsxFile, parse_range

    if not os.path.exists(path):
        return {"error": f"File not found: {path}"}

    try:
        xf = XlsxFile(path).open()
    except Exception as e:
        return {"error": f"Cannot open file: {e}"}

    try:
        sheet_name = sheet or xf.sheet_names[0]
        if sheet_name not in xf.sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found"}

        c1, r1, c2, r2 = parse_range(cell_range)
        rows = r2 - r1 + 1
        cols = c2 - c1 + 1
        value_2d = _to_2d(value, rows, cols)

        xf.write_values(sheet_name, cell_range, value_2d)
        xf.save()

        return {
            "success": True,
            "path": path,
            "sheet": sheet_name,
            "range": cell_range,
            "size": f"{rows}x{cols}"
        }
    except Exception as e:
        return {"error": f"Failed to write: {e}"}
    finally:
        xf.close()


def _reshape(value, rows, cols):
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


def _to_2d(value, rows, cols):
    """Convert value to a 2D list for xlsx_io."""
    if not isinstance(value, list):
        return [[value] * cols for _ in range(rows)]
    if value and not isinstance(value[0], list):
        # Flat array
        if rows > 1 and cols == 1:
            return [[item] for item in value[:rows]]
        elif rows == 1:
            return [value[:cols]]
        else:
            return [value[:cols] for _ in range(rows)]
    elif value and isinstance(value[0], list):
        return [row[:cols] for row in value[:rows]]
    return [[value]]


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
