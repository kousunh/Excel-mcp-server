import xlwings as xw
import argparse
import json

def navigate_to_sheet(sheet_name, filename=None):
    try:
        # Connect to open Excel with robust method
        try:
            app = xw.apps.active
        except:
            # アクティブなアプリがない場合は新しく起動を試みる
            try:
                app = xw.App(visible=True)
            except:
                print(json.dumps({
                    "status": "error",
                    "message": "Cannot connect to Excel. Please make sure Excel is installed and accessible."
                }))
                return
        
        if app is None:
            print(json.dumps({
                "status": "error",
                "message": "Excel is not running. Please open Excel first."
            }))
            return
        
        # 指定されたファイル名のブックを取得、またはアクティブなブックを取得
        if filename:
            wb = None
            for book in app.books:
                if book.name == filename or book.fullname == filename:
                    wb = book
                    break
            if wb is None:
                print(json.dumps({
                    "status": "error",
                    "message": f"Workbook '{filename}' not found. Please check the filename."
                }))
                return
        else:
            wb = app.books.active
            if wb is None:
                print(json.dumps({
                    "status": "error",
                    "message": "No active workbook found. Please open or create a workbook."
                }))
                return
        
        # Activate the workbook
        wb.activate()
        
        # 指定されたシートが存在するか確認
        if sheet_name not in [sheet.name for sheet in wb.sheets]:
            available_sheets = [sheet.name for sheet in wb.sheets]
            print(json.dumps({
                "status": "error",
                "message": f"Sheet '{sheet_name}' not found in workbook.",
                "available_sheets": available_sheets
            }))
            return
        
        # シートに移動（アクティブにする）
        sheet = wb.sheets[sheet_name]
        sheet.activate()
        
        # 結果を返す
        print(json.dumps({
            "status": "success",
            "workbook": wb.name,
            "activated_sheet": sheet_name,
            "message": f"Successfully navigated to sheet '{sheet_name}' in workbook '{wb.name}'"
        }))
        
    except Exception as e:
        print(json.dumps({
            "status": "error",
            "message": f"Error navigating to sheet: {str(e)}"
        }))

def main():
    parser = argparse.ArgumentParser(description='Navigate to a specific sheet in Excel workbook')
    parser.add_argument('--sheet', required=True, help='Sheet name to navigate to')
    parser.add_argument('--filename', help='Excel filename')
    
    args = parser.parse_args()
    
    navigate_to_sheet(args.sheet, args.filename)

if __name__ == "__main__":
    main()