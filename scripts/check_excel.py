import xlwings as xw
import json

def check_excel_status():
    try:
        # Check Excel application with robust connection method
        try:
            app = xw.apps.active
        except:
            print(json.dumps({
                "status": "connection_error",
                "message": "Cannot connect to Excel. Please make sure Excel is installed and accessible.",
                "has_active_app": False,
                "has_active_workbook": False
            }))
            return
        
        if app is None:
            print(json.dumps({
                "status": "not_running",
                "message": "Excel is not running",
                "has_active_app": False,
                "has_active_workbook": False
            }))
            return
        
        # すべてのブックの情報を取得
        workbooks_info = []
        active_wb = app.books.active
        
        for book in app.books:
            workbook_info = {
                "name": book.name,
                "fullname": book.fullname if hasattr(book, 'fullname') else "Unsaved",
                "is_active": book == active_wb,
                "sheets": [sheet.name for sheet in book.sheets],
                "active_sheet": book.sheets.active.name if book.sheets.active else None
            }
            # Check if saved attribute exists
            try:
                workbook_info["saved"] = book.api.Saved if hasattr(book.api, 'Saved') else True
            except:
                workbook_info["saved"] = True
            
            workbooks_info.append(workbook_info)
        
        if active_wb is None:
            print(json.dumps({
                "status": "no_active_workbook",
                "message": "Excel is running but no active workbook",
                "has_active_app": True,
                "has_active_workbook": False,
                "excel_version": str(app.api.Version),
                "open_workbooks": workbooks_info,
                "workbook_count": len(workbooks_info)
            }))
            return
        
        # 詳細情報を取得
        print(json.dumps({
            "status": "ready",
            "message": "Excel is ready for VBA execution",
            "has_active_app": True,
            "has_active_workbook": True,
            "active_workbook_name": active_wb.name,
            "active_workbook_path": active_wb.fullname if hasattr(active_wb, 'fullname') else "Unsaved",
            "excel_version": str(app.api.Version),
            "sheets_count": len(active_wb.sheets),
            "active_sheet": active_wb.sheets.active.name,
            "open_workbooks": workbooks_info,
            "workbook_count": len(workbooks_info)
        }))
        
    except Exception as e:
        print(json.dumps({
            "status": "error",
            "message": f"Error checking Excel status: {str(e)}",
            "has_active_app": False,
            "has_active_workbook": False
        }))

if __name__ == "__main__":
    check_excel_status()