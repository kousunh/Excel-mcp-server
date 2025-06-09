import xlwings as xw
import pandas as pd
import argparse
import json
import sys

def read_sheet_data(start_row=1, end_row=100, filename=None, sheet_name=None):
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

        # Activate the workbook to make it visible
        wb.activate()

        # Get sheet
        if sheet_name:
            if sheet_name not in [s.name for s in wb.sheets]:
                return {"error": f"Sheet '{sheet_name}' not found"}
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.active
        
        # Activate the sheet to make it visible
        sheet.activate()

        # Get the used range to determine the actual data boundaries
        used_range = sheet.used_range
        if not used_range:
            return {
                "workbook": wb.name,
                "sheet": sheet.name,
                "rows_read": 0,
                "columns_read": 0,
                "rows": []
            }

        # Get the last column with data
        last_col = used_range.last_cell.column
        
        # Ensure we don't go beyond the actual data
        actual_last_row = used_range.last_cell.row
        end_row = min(end_row, actual_last_row)
        
        # If start_row is beyond actual data, return empty
        if start_row > actual_last_row:
            return {
                "workbook": wb.name,
                "sheet": sheet.name,
                "rows_read": 0,
                "columns_read": 0,
                "rows": []
            }

        # Read the data range
        data_range = sheet.range((start_row, 1), (end_row, last_col))
        data = data_range.value
        
        # No encoding conversion needed - xlwings returns proper Unicode strings
        # The issue was with stdout encoding, not the data itself
        def clean_string(value):
            if isinstance(value, str):
                # Just remove null bytes and normalize whitespace
                value = value.replace('\x00', '').strip()
            return value

        # Convert to pandas DataFrame
        if data:
            # If data is a single row, wrap it in a list
            if not isinstance(data[0], list):
                data = [data]
            
            # Clean the data before creating DataFrame
            if isinstance(data[0], list):
                # Multiple rows
                cleaned_data = [[clean_string(cell) if cell is not None else None for cell in row] for row in data]
            else:
                # Single row
                cleaned_data = [clean_string(cell) if cell is not None else None for cell in data]
            
            # Create DataFrame with proper column handling
            df = pd.DataFrame(cleaned_data)
            
            # Always use Excel-style column names to avoid conflicts
            excel_columns = []
            for i in range(df.shape[1]):
                col_letter = chr(65 + i) if i < 26 else f"A{chr(65 + i - 26)}"
                excel_columns.append(col_letter)
            
            # Store original first row data if it might be headers
            original_headers = None
            if start_row == 1 and df.shape[0] > 1:
                first_row = df.iloc[0]
                if all(isinstance(val, str) or val is None for val in first_row):
                    original_headers = list(first_row)
                    df = df[1:].reset_index(drop=True)
            
            # Set Excel-style column names (A, B, C, etc.)
            df.columns = excel_columns[:df.shape[1]]
            
            # Preserve ALL rows including completely empty ones for accurate row mapping
            # No row filtering - maintain exact Excel structure
            
            # Convert DataFrame to JSON-friendly format with row numbers
            rows_data = []
            for index, row in df.iterrows():
                excel_row_number = start_row + index + (1 if original_headers else 0)  # Adjust for header removal
                row_data = {
                    "excel_row": excel_row_number,
                    "data": row.to_dict()
                }
                rows_data.append(row_data)
            
            result = {
                "workbook": wb.name,
                "sheet": sheet.name,
                "start_row": start_row,
                "end_row": end_row,
                "rows_read": len(rows_data),
                "columns_read": df.shape[1] if not df.empty else 0,
                "column_names": list(df.columns) if not df.empty else [],  # Excel column letters (A, B, C...)
                "original_headers": original_headers,  # Original header row content if detected
                "rows": rows_data,  # Changed from "data" to "rows" for clarity
                "note": "All rows/columns preserved. Each row shows excel_row number and data with Excel column naming (A,B,C...)."
            }
            
            # Add summary statistics for numeric columns
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                result["numeric_summary"] = {
                    col: {
                        "mean": float(df[col].mean()) if not df[col].isna().all() else None,
                        "min": float(df[col].min()) if not df[col].isna().all() else None,
                        "max": float(df[col].max()) if not df[col].isna().all() else None,
                        "count": int(df[col].count())
                    }
                    for col in numeric_cols
                }
            
            return result
        else:
            return {
                "workbook": wb.name,
                "sheet": sheet.name,
                "rows_read": 0,
                "columns_read": 0,
                "rows": []
            }

    except Exception as e:
        return {"error": str(e)}

def main():
    parser = argparse.ArgumentParser(description='Read sheet data and convert to JSON')
    parser.add_argument('--start-row', type=int, default=1, help='Starting row (default: 1)')
    parser.add_argument('--end-row', type=int, default=100, help='Ending row (default: 100)')
    parser.add_argument('--filename', help='Optional Excel filename')
    parser.add_argument('--sheet', help='Optional sheet name')
    
    args = parser.parse_args()
    
    # Set stdout encoding to UTF-8 for Windows
    import sys
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    result = read_sheet_data(args.start_row, args.end_row, args.filename, args.sheet)
    # Ensure UTF-8 output to handle Unicode characters
    output = json.dumps(result, ensure_ascii=False, default=str)
    try:
        print(output)
    except UnicodeEncodeError:
        # Fallback: encode to UTF-8 and write to stdout buffer
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()