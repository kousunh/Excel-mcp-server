"""Find cells matching a value or regex pattern in an Excel sheet."""

import argparse
import re
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import get_app, get_workbook, get_sheet, output_json, IS_WINDOWS


def _clean_value(val):
    """Clean value for JSON output."""
    if val is None:
        return None
    if hasattr(val, 'isoformat'):
        return val.isoformat()
    if isinstance(val, float) and val == int(val):
        return int(val)
    return val


# ---------------------------------------------------------------------------
# xlwings (live Excel / workbook mode)
# ---------------------------------------------------------------------------

def _find_live(workbook, query, sheet, search_range, is_regex, max_results):
    app, err = get_app()
    if err:
        return {"error": err}
    wb, err = get_workbook(app, workbook)
    if err:
        return {"error": err}
    ws, err = get_sheet(wb, sheet)
    if err:
        return {"error": err}

    matches = []

    if is_regex:
        # Regex: must iterate cells
        pattern = re.compile(query)
        rng = ws.range(search_range) if search_range else ws.used_range
        if rng is None:
            return {"workbook": wb.name, "sheet": ws.name, "query": query,
                    "matches": [], "count": 0, "limited": False}

        rows, cols = rng.shape
        for r in range(rows):
            for c in range(cols):
                cell = rng[r, c]
                val = cell.value
                if val is not None and pattern.search(str(val)):
                    matches.append({
                        'cell': cell.address.replace('$', ''),
                        'value': _clean_value(val)
                    })
                    if len(matches) >= max_results:
                        break
            if len(matches) >= max_results:
                break
    else:
        # Exact match: use Excel's native Find API on Windows
        rng = ws.range(search_range) if search_range else ws.used_range
        if rng is None:
            return {"workbook": wb.name, "sheet": ws.name, "query": query,
                    "matches": [], "count": 0, "limited": False}

        found_via_api = False
        if IS_WINDOWS:
            try:
                found = rng.api.Find(
                    What=query,
                    LookIn=-4163,   # xlValues
                    LookAt=1,       # xlWhole
                    SearchOrder=1   # xlByRows
                )
                if found:
                    found_via_api = True
                    first_addr = found.Address
                    while True:
                        addr = found.Address.replace('$', '')
                        val = found.Value
                        matches.append({'cell': addr, 'value': _clean_value(val)})
                        if len(matches) >= max_results:
                            break
                        found = rng.api.FindNext(found)
                        if found is None or found.Address == first_addr:
                            break
                else:
                    found_via_api = True  # API worked, just no matches
            except Exception:
                pass

        if not found_via_api:
            # Fallback: iterate cells
            rows, cols = rng.shape
            for r in range(rows):
                for c in range(cols):
                    cell = rng[r, c]
                    val = cell.value
                    if val is not None and str(val) == query:
                        matches.append({
                            'cell': cell.address.replace('$', ''),
                            'value': _clean_value(val)
                        })
                        if len(matches) >= max_results:
                            break
                if len(matches) >= max_results:
                    break

    return {
        "workbook": wb.name,
        "sheet": ws.name,
        "query": query,
        "matches": matches,
        "count": len(matches),
        "limited": len(matches) >= max_results
    }


# ---------------------------------------------------------------------------
# xlsx_io (file-based, pure Python ZIP/XML, no Excel needed)
# ---------------------------------------------------------------------------

def _find_file(path, query, sheet, search_range, is_regex, max_results):
    from xlsx_io import XlsxFile

    if not os.path.exists(path):
        return {"error": f"File not found: {path}"}

    try:
        xf = XlsxFile(path).open()
    except Exception as e:
        return {"error": f"Cannot open file: {e}"}

    try:
        sheet_name = sheet or xf.sheet_names[0]
        if sheet_name not in xf.sheet_names:
            return {"error": f"Sheet '{sheet_name}' not found. Available: {xf.sheet_names}"}

        matches = xf.find_cells(sheet_name, query, is_regex=is_regex,
                                search_range=search_range, max_results=max_results)

        return {
            "path": path,
            "sheet": sheet_name,
            "query": query,
            "matches": matches,
            "count": len(matches),
            "limited": len(matches) >= max_results
        }
    except Exception as e:
        return {"error": f"Failed to search: {e}"}
    finally:
        xf.close()


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', default=None)
    parser.add_argument('--path', default=None)
    parser.add_argument('--query', required=True)
    parser.add_argument('--sheet', default=None)
    parser.add_argument('--range', default=None)
    parser.add_argument('--regex', action='store_true')
    parser.add_argument('--max', type=int, default=100)
    args = parser.parse_args()

    if not args.workbook and not args.path:
        output_json({"error": "Either --workbook or --path is required"})
        return

    if args.path:
        result = _find_file(args.path, args.query, args.sheet, args.range,
                            args.regex, args.max)
    else:
        result = _find_live(args.workbook, args.query, args.sheet, args.range,
                            args.regex, args.max)

    output_json(result)


if __name__ == "__main__":
    main()
