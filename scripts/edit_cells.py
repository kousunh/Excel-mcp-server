import xlwings as xw
import argparse
import json
import sys

def edit_cells(cell_range, value, filename=None, sheet_name=None):
    try:
        # Get active Excel app（より堅牢な接続方法）
        try:
            app = xw.apps.active
        except:
            return {"error": "Cannot connect to Excel. Please make sure Excel is running."}

        # Get workbook
        if filename:
            # Find workbook by filename
            wb = None
            for book in app.books:
                if book.name == filename or book.fullname == filename:
                    wb = book
                    break
            if not wb:
                return {"error": f"Workbook '{filename}' not found"}
        else:
            # Use active workbook
            wb = app.books.active
            if not wb:
                return {"error": "No active workbook found"}
        
        # Activate the workbook
        wb.activate()

        # Get target sheet
        if sheet_name:
            try:
                sheet = wb.sheets[sheet_name]
                sheet.activate()
            except Exception as e:
                return {"error": f"Cannot navigate to sheet '{sheet_name}': {str(e)}"}
        else:
            sheet = wb.sheets.active
        
        # Edit the cells
        try:
            # Support both single cell (e.g., "A1") and range (e.g., "A1:B5")
            target_range = sheet.range(cell_range)
            
            # If value is a list and we're targeting a column range (e.g., C9:C16)
            if isinstance(value, list) and target_range.shape:
                rows, cols = target_range.shape
                
                # If it's a single column range and we have a flat list
                if cols == 1 and rows > 1:
                    # Convert flat list to 2D array for vertical placement
                    value = [[item] for item in value[:rows]]  # Limit to range size
                
                # If it's a single row range and we have a flat list
                elif rows == 1 and cols > 1:
                    # Keep as flat list for horizontal placement
                    value = value[:cols]  # Limit to range size
            
            sheet.range(cell_range).value = value
            
            # Get the actual range that was edited
            edited_range = sheet.range(cell_range)
            
            # Prepare result
            result = {
                "success": True,
                "workbook": wb.name,
                "sheet": sheet.name,
                "range": cell_range,
                "rows_affected": edited_range.shape[0] if edited_range.shape else 1,
                "columns_affected": edited_range.shape[1] if edited_range.shape else 1,
                "message": f"Successfully updated cells {cell_range} in sheet '{sheet.name}'"
            }
            
            # If it's a small range, include the new values
            if edited_range.size and edited_range.size <= 100:
                result["new_values"] = edited_range.value
            
            return result
            
        except Exception as e:
            return {"error": f"Failed to edit cells: {str(e)}"}

    except Exception as e:
        return {"error": str(e)}

def main():
    parser = argparse.ArgumentParser(description='Edit cells in Excel')
    parser.add_argument('--range', required=True, help='Cell range to edit (e.g., "A1" or "A1:B5")')
    parser.add_argument('--value', required=True, help='Value to set. For multiple cells, use JSON array format')
    parser.add_argument('--filename', help='Optional Excel filename')
    parser.add_argument('--sheet', help='Optional sheet name to navigate to')
    
    args = parser.parse_args()
    
    # Parse value - it could be a simple value or JSON array for multiple cells
    try:
        # Try to parse as JSON first (for arrays)
        value = json.loads(args.value)
    except:
        # If not JSON, use as string
        value = args.value
    
    result = edit_cells(args.range, value, args.filename, args.sheet)
    
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