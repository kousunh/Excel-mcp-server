import xlwings as xw
import argparse
import json
import sys

def edit_cells(cell_range, value, filename=None, sheet_name=None):
    try:
        # Get active Excel app with robust connection method
        try:
            app = xw.apps.active
        except:
            return {"error": "Cannot connect to Excel. Please make sure Excel is running."}

        # Disable screen updating for performance
        original_screen_updating = app.api.Application.ScreenUpdating
        original_calculation = app.api.Application.Calculation
        app.api.Application.ScreenUpdating = False
        app.api.Application.Calculation = -4135  # xlCalculationManual

        try:
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
            
            # Get target sheet
            if sheet_name:
                try:
                    sheet = wb.sheets[sheet_name]
                except Exception as e:
                    return {"error": f"Cannot navigate to sheet '{sheet_name}': {str(e)}"}
            else:
                sheet = wb.sheets.active
        
            # Edit the cells
            try:
                # Support both single cell (e.g., "A1") and range (e.g., "A1:B5")
                target_range = sheet.range(cell_range)
                
                # Optimize data format for Excel
                if isinstance(value, list) and target_range.shape:
                    rows, cols = target_range.shape
                    
                    # If it's a 2D list, ensure it matches the target range dimensions
                    if isinstance(value[0], list):
                        # Trim to fit target range
                        value = [row[:cols] for row in value[:rows]]
                    else:
                        # If it's a single column range and we have a flat list
                        if cols == 1 and rows > 1:
                            # Convert flat list to 2D array for vertical placement
                            value = [[item] for item in value[:rows]]
                        # If it's a single row range and we have a flat list
                        elif rows == 1 and cols > 1:
                            # Keep as flat list for horizontal placement
                            value = value[:cols]
                
                # Set values efficiently
                sheet.range(cell_range).value = value
                
                # Get info about what was updated
                edited_range = sheet.range(cell_range)
                shape = edited_range.shape if edited_range.shape else (1, 1)
                
                # Prepare result
                result = {
                    "success": True,
                    "workbook": wb.name,
                    "sheet": sheet.name,
                    "range": cell_range,
                    "rows_affected": shape[0],
                    "columns_affected": shape[1],
                    "message": f"Successfully updated {shape[0]}x{shape[1]} cells in range {cell_range}"
                }
                
                return result
                
            except Exception as e:
                return {"error": f"Failed to edit cells: {str(e)}"}
                
        finally:
            # Always restore Excel settings
            try:
                app.api.Application.ScreenUpdating = original_screen_updating
                app.api.Application.Calculation = original_calculation
            except:
                pass

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
    
    # Set stdout encoding to UTF-8 for Windows
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # Custom JSON serializer for datetime objects
    def json_serial(obj):
        if hasattr(obj, 'isoformat'):
            return obj.isoformat()
        elif hasattr(obj, '__str__'):
            return str(obj)
        raise TypeError(f"Type {type(obj)} not serializable")
    
    # Ensure UTF-8 output to handle Unicode characters and datetime objects
    output = json.dumps(result, ensure_ascii=False, default=json_serial)
    try:
        print(output)
    except UnicodeEncodeError:
        # Fallback: encode to UTF-8 and write to stdout buffer
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()