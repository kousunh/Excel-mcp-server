"""Scan a sheet's structure: dimensions, merged cells, data extent."""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import get_app, get_workbook, get_sheet, output_json, IS_WINDOWS


# ---------------------------------------------------------------------------
# xlwings (live Excel / workbook mode)
# ---------------------------------------------------------------------------

def _scan_live(workbook, sheet):
    app, err = get_app()
    if err:
        return {"error": err}
    wb, err = get_workbook(app, workbook)
    if err:
        return {"error": err}

    # If no sheet specified, return overview of all sheets
    if not sheet:
        sheets_info = []
        for ws in wb.sheets:
            info = {"name": ws.name}
            try:
                ur = ws.used_range
                if ur:
                    info["usedRange"] = ur.address.replace('$', '')
                    info["rowCount"] = ur.shape[0]
                    info["colCount"] = ur.shape[1]
                else:
                    info["usedRange"] = None
                    info["rowCount"] = 0
                    info["colCount"] = 0
            except Exception:
                info["usedRange"] = None
                info["rowCount"] = 0
                info["colCount"] = 0
            sheets_info.append(info)
        return {
            "workbook": wb.name,
            "sheets": sheets_info,
            "sheetCount": len(sheets_info)
        }

    ws, err = get_sheet(wb, sheet)
    if err:
        return {"error": err}

    result = {"workbook": wb.name, "sheet": ws.name}

    # Used range
    try:
        ur = ws.used_range
        if ur:
            result["usedRange"] = ur.address.replace('$', '')
            result["rowCount"] = ur.shape[0]
            result["colCount"] = ur.shape[1]
        else:
            result["usedRange"] = None
            result["rowCount"] = 0
            result["colCount"] = 0
    except Exception:
        result["usedRange"] = None
        result["rowCount"] = 0
        result["colCount"] = 0

    # Merged cells via API
    merged = []
    try:
        if IS_WINDOWS:
            ur = ws.used_range
            if ur:
                seen = set()
                for r in range(ur.shape[0]):
                    for c in range(ur.shape[1]):
                        cell = ur[r, c]
                        try:
                            ma = cell.api.MergeArea
                            addr = ma.Address.replace('$', '')
                            if ':' in addr and addr not in seen:
                                seen.add(addr)
                                merged.append(addr)
                        except Exception:
                            pass
    except Exception:
        pass
    result["mergedCells"] = merged

    return result


# ---------------------------------------------------------------------------
# xlsx_io (file-based, pure Python ZIP/XML, no Excel needed)
# ---------------------------------------------------------------------------

def _scan_file(path, sheet):
    from xlsx_io import XlsxFile

    if not os.path.exists(path):
        return {"error": f"File not found: {path}"}

    try:
        xf = XlsxFile(path).open()
    except Exception as e:
        return {"error": f"Cannot open file: {e}"}

    try:
        if not sheet:
            all_sheets = []
            for sn in xf.sheet_names:
                info = xf.scan_sheet(sn)
                info['name'] = sn
                all_sheets.append(info)
            return {
                "path": path,
                "sheets": all_sheets,
                "sheetCount": len(all_sheets)
            }

        if sheet not in xf.sheet_names:
            return {"error": f"Sheet '{sheet}' not found. Available: {xf.sheet_names}"}

        info = xf.scan_sheet(sheet)
        info['path'] = path
        info['sheet'] = sheet
        info['sheetNames'] = xf.sheet_names
        return info
    except Exception as e:
        return {"error": f"Failed to scan: {e}"}
    finally:
        xf.close()


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', default=None)
    parser.add_argument('--path', default=None)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    if not args.workbook and not args.path:
        output_json({"error": "Either --workbook or --path is required"})
        return

    if args.path:
        result = _scan_file(args.path, args.sheet)
    else:
        result = _scan_live(args.workbook, args.sheet)

    output_json(result)


if __name__ == "__main__":
    main()
