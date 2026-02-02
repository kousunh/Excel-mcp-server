"""Write values to cells in an Excel workbook (live or file)."""

import argparse
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet, open_path,
    set_performance_mode, restore_performance_mode, output_json
)


def _get_wb(workbook=None, path=None):
    """Resolve workbook from either name or path. Returns (wb, was_opened, error)."""
    if path:
        return open_path(path)
    app, err = get_app()
    if err:
        return None, False, err
    wb, err = get_workbook(app, workbook)
    return wb, False, err


def write_cells(workbook=None, path=None, cell_range=None, value=None, sheet=None):
    wb, was_opened, err = _get_wb(workbook, path)
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
        if was_opened:
            try:
                wb.close()
            except Exception:
                pass


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

    result = write_cells(args.workbook, args.path, args.range, value, args.sheet)
    output_json(result)


if __name__ == "__main__":
    main()
