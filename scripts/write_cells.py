"""Write values to cells in an open Excel workbook."""

import argparse
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import (
    get_app, get_workbook, get_sheet,
    set_performance_mode, restore_performance_mode, output_json
)


def write_cells(workbook, cell_range, value, sheet=None):
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

        # Adjust value shape for the target range
        rows, cols = rng.shape
        if isinstance(value, list):
            if rows > 1 and cols == 1 and value and not isinstance(value[0], list):
                # Flat list to vertical column
                value = [[item] for item in value[:rows]]
            elif rows == 1 and cols > 1 and value and not isinstance(value[0], list):
                # Flat list to horizontal row
                value = value[:cols]
            elif isinstance(value[0], list):
                # 2D list - trim to fit
                value = [row[:cols] for row in value[:rows]]

        rng.value = value
        shape = rng.shape

        return {
            "success": True,
            "workbook": wb.name,
            "sheet": ws.name,
            "range": cell_range,
            "size": f"{shape[0]}x{shape[1]}"
        }
    except Exception as e:
        return {"error": f"Failed to write: {e}"}
    finally:
        restore_performance_mode(app, perf)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', required=True)
    parser.add_argument('--range', required=True)
    parser.add_argument('--value', required=True)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    # Parse value - JSON for arrays, plain string otherwise
    try:
        value = json.loads(args.value)
    except (json.JSONDecodeError, ValueError):
        value = args.value

    result = write_cells(args.workbook, args.range, value, args.sheet)
    output_json(result)


if __name__ == "__main__":
    main()
