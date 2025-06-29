import xlwings as xw
import argparse
import json

def get_all_sheet_names(filename=None):
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
        
        # すべてのシート名を取得
        sheet_names = [sheet.name for sheet in wb.sheets]
        
        # 結果を返す
        print(json.dumps({
            "status": "success",
            "workbook": wb.name,
            "sheet_count": len(sheet_names),
            "sheet_names": sheet_names,
            "message": f"Found {len(sheet_names)} sheets in workbook '{wb.name}'"
        }))
        
    except Exception as e:
        print(json.dumps({
            "status": "error",
            "message": f"Error getting sheet names: {str(e)}"
        }))

def main():
    parser = argparse.ArgumentParser(description='Get all sheet names in Excel workbook')
    parser.add_argument('--filename', help='Excel filename')
    
    args = parser.parse_args()
    
    get_all_sheet_names(args.filename)

if __name__ == "__main__":
    main()