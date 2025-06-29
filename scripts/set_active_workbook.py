import xlwings as xw
import argparse
import json
import sys

def set_active_workbook(workbook_name):
    try:
        # Get active Excel app
        try:
            app = xw.apps.active
        except:
            return {"error": "Cannot connect to Excel. Please make sure Excel is running."}
        
        # Find and activate the workbook
        wb = None
        for book in app.books:
            if book.name == workbook_name or book.fullname == workbook_name:
                wb = book
                break
        
        if not wb:
            return {"error": f"Workbook '{workbook_name}' not found"}
        
        # Activate the workbook
        wb.activate()
        
        result = {
            "success": True,
            "workbook": wb.name,
            "fullname": wb.fullname,
            "active_sheet": wb.sheets.active.name if wb.sheets.active else None,
            "message": f"Successfully activated workbook '{wb.name}'"
        }
        
        return result
        
    except Exception as e:
        return {"error": f"Error activating workbook: {str(e)}"}

def main():
    parser = argparse.ArgumentParser(description='Set active workbook in Excel')
    parser.add_argument('--workbook', required=True, help='Name of the workbook to activate')
    
    args = parser.parse_args()
    
    result = set_active_workbook(args.workbook)
    # Ensure UTF-8 output to handle Unicode characters
    output = json.dumps(result, ensure_ascii=False)
    try:
        print(output)
    except UnicodeEncodeError:
        # Fallback: encode to UTF-8 and write to stdout buffer
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()