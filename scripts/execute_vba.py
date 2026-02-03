"""Execute VBA macro code in an open Excel workbook."""

import argparse
import json
import sys
import os
import re
import random

sys.path.insert(0, os.path.dirname(__file__))
from excel_utils import get_app, get_workbook, get_sheet, output_json, IS_WINDOWS


def _clean_temp_modules(wb):
    """Remove leftover temporary modules."""
    try:
        to_remove = [c for c in wb.api.VBProject.VBComponents if c.Name.startswith("TempMCP")]
        for comp in to_remove:
            try:
                wb.api.VBProject.VBComponents.Remove(comp)
            except Exception:
                continue
    except Exception:
        pass


def _detect_procedure(code):
    """Detect if code already has Sub/Function structure."""
    match = re.search(r'\b(Sub|Function)\s+(\w+)', code, re.IGNORECASE)
    if match:
        return match.group(2)
    return None


def _clean_code(code):
    """Remove MsgBox statements and clean up code."""
    lines = []
    for line in code.split('\n'):
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            lines.append(line)
            continue
        # Remove MsgBox to prevent popup alerts
        if re.search(r'\bMsgBox\b', stripped, re.IGNORECASE):
            continue
        lines.append(line)
    return '\n'.join(lines)


def execute_vba(workbook, code, sheet=None):
    app, err = get_app()
    if err:
        return {"error": err}

    wb, err = get_workbook(app, workbook)
    if err:
        return {"error": err}

    if sheet:
        ws, err = get_sheet(wb, sheet)
        if err:
            return {"error": err}
        ws.activate()

    # Clean and prepare code
    code = _clean_code(code)
    proc_name = _detect_procedure(code)
    module_name = f"TempMCP{random.randint(1000, 9999)}"

    # Wrap in Sub if no structure detected
    if proc_name is None:
        proc_name = "Main"
        code = f"Sub {proc_name}()\nOn Error GoTo ErrExit\n{code}\nExit Sub\nErrExit:\nExit Sub\nEnd Sub"
    else:
        # Add error handling if not present
        if 'On Error' not in code:
            code = re.sub(
                r'(Sub\s+\w+\s*\([^)]*\))',
                r'\1\nOn Error GoTo ErrExit',
                code, count=1, flags=re.IGNORECASE
            )
            code = re.sub(
                r'End Sub',
                'Exit Sub\nErrExit:\nExit Sub\nEnd Sub',
                code, count=1, flags=re.IGNORECASE
            )

    vba_module = None
    # Suppress Excel alerts during execution
    try:
        if IS_WINDOWS:
            orig_alerts = app.api.Application.DisplayAlerts
            orig_screen = app.api.Application.ScreenUpdating
            app.api.Application.DisplayAlerts = False
            app.api.Application.ScreenUpdating = False
        else:
            app.display_alerts = False
            app.screen_updating = False
    except Exception:
        pass

    try:
        _clean_temp_modules(wb)

        # Add module and code
        vba_module = wb.api.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
        vba_module.Name = module_name
        vba_module.CodeModule.AddFromString(code)

        # Execute
        full_name = f"{module_name}.{proc_name}"
        if IS_WINDOWS:
            wb.api.Application.Run(full_name)
        else:
            app.macro(full_name)()

        return {
            "success": True,
            "message": f"VBA executed: {proc_name}"
        }

    except Exception as e:
        return {"error": f"VBA execution failed: {e}"}

    finally:
        # Clean up module
        if vba_module:
            try:
                wb.api.VBProject.VBComponents.Remove(vba_module)
            except Exception:
                pass
        # Restore settings
        try:
            if IS_WINDOWS:
                app.api.Application.DisplayAlerts = orig_alerts
                app.api.Application.ScreenUpdating = orig_screen
            else:
                app.display_alerts = True
                app.screen_updating = True
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--workbook', required=True)
    parser.add_argument('--code', required=True)
    parser.add_argument('--sheet', default=None)
    args = parser.parse_args()

    result = execute_vba(args.workbook, args.code, args.sheet)
    output_json(result)


if __name__ == "__main__":
    main()
