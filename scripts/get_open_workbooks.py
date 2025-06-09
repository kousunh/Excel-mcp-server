import xlwings as xw
import json

def get_open_workbooks():
    try:
        # Get active Excel app
        try:
            app = xw.apps.active
        except:
            return {"error": "Cannot connect to Excel. Please make sure Excel is running."}
        
        # Get all open workbooks
        workbooks = []
        active_workbook_name = None
        
        if app.books.active:
            active_workbook_name = app.books.active.name
        
        for book in app.books:
            workbook_info = {
                "name": book.name,
                "fullname": book.fullname if hasattr(book, 'fullname') else book.name,
                "is_active": book.name == active_workbook_name,
                "sheets": [sheet.name for sheet in book.sheets],
                "active_sheet": book.sheets.active.name if book.sheets.active else None
            }
            
            # Check if saved attribute exists
            try:
                workbook_info["saved"] = book.api.Saved if hasattr(book.api, 'Saved') else True
            except:
                workbook_info["saved"] = True
                
            # Check if ReadOnly attribute exists
            try:
                workbook_info["read_only"] = book.api.ReadOnly if hasattr(book.api, 'ReadOnly') else False
            except:
                workbook_info["read_only"] = False
                
            workbooks.append(workbook_info)
        
        result = {
            "success": True,
            "workbooks": workbooks,
            "active_workbook": active_workbook_name,
            "total_count": len(workbooks)
        }
        
        return result
        
    except Exception as e:
        return {"error": f"Error getting open workbooks: {str(e)}"}

def main():
    result = get_open_workbooks()
    # Ensure UTF-8 output to handle Unicode characters
    output = json.dumps(result, ensure_ascii=False, indent=2)
    try:
        print(output)
    except UnicodeEncodeError:
        # Fallback: encode to UTF-8 and write to stdout buffer
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()