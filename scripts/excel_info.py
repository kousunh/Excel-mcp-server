"""Get Excel status, open workbooks, and sheet information."""

import json
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from excel_utils import get_app, output_json


def get_excel_info():
    app, err = get_app()
    if err:
        return {"status": "not_running", "message": err}

    workbooks = []
    active_wb = app.books.active

    for book in app.books:
        wb_info = {
            "name": book.name,
            "path": book.fullname,
            "active": book == active_wb,
            "sheets": [],
            "activeSheet": None
        }
        active_sheet = book.sheets.active
        for sheet in book.sheets:
            wb_info["sheets"].append(sheet.name)
            if sheet == active_sheet:
                wb_info["activeSheet"] = sheet.name
        workbooks.append(wb_info)

    return {
        "status": "ready",
        "workbooks": workbooks,
        "count": len(workbooks)
    }


if __name__ == "__main__":
    output_json(get_excel_info())
