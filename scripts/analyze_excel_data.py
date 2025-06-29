import xlwings as xw
import pandas as pd
import argparse
import json
import sys
import io
import os
from datetime import datetime, date

def analyze_excel_data(file_path=None, workbook_name=None, sheet_name=None, start_row=1, end_row=None, analyze_mode='full'):
    """
    Universal Excel data analyzer with clear structure identification.
    
    Args:
        file_path: Path to Excel file (for closed files)
        workbook_name: Name of open workbook (for open files)
        sheet_name: Specific sheet name (optional)
        start_row: Starting row number (default: 1)
        end_row: Ending row number (optional, defaults to all data)
        analyze_mode: 'full' (detailed analysis), 'quick' (basic info), 'data' (data only)
    """
    try:
        result = {
            "source_type": None,
            "workbook": None,
            "sheet": None,
            "analysis_mode": analyze_mode,
            "sheets_info": [],
            "structure_analysis": {},
            "data": None,
            "recommendations": {}
        }
        
        # Determine data source: open file or closed file
        raw_data = None
        all_sheets_info = []
        actual_last_row = 0
        actual_last_col = 0
        
        if file_path and os.path.exists(file_path):
            # Analyze closed file using pandas
            result["source_type"] = "file"
            result["file_path"] = file_path
            
            try:
                # Get all sheet names first
                excel_file = pd.ExcelFile(file_path)
                all_sheet_names = excel_file.sheet_names
                
                # Analyze each sheet or specific sheet
                sheets_to_analyze = [sheet_name] if sheet_name else all_sheet_names
                
                for sheet in sheets_to_analyze:
                    if sheet in all_sheet_names:
                        # Read sheet data with all rows
                        sheet_df = pd.read_excel(file_path, sheet_name=sheet, header=None)
                        
                        # Basic sheet info
                        sheet_info = {
                            "name": sheet,
                            "total_rows": len(sheet_df),
                            "total_columns": len(sheet_df.columns),
                            "has_data": not sheet_df.empty
                        }
                        all_sheets_info.append(sheet_info)
                        
                        # If specific sheet requested, set as main data
                        if sheet_name == sheet or len(sheets_to_analyze) == 1:
                            raw_data = sheet_df.values.tolist() if not sheet_df.empty else []
                            actual_last_row = len(sheet_df)
                            actual_last_col = len(sheet_df.columns)
                            result["sheet"] = sheet
                
                result["workbook"] = os.path.basename(file_path)
                result["sheets_info"] = all_sheets_info
                
            except Exception as e:
                return {"error": f"Failed to read Excel file: {str(e)}"}
                
        elif workbook_name:
            # Analyze open file using xlwings
            result["source_type"] = "open_workbook"
            
            try:
                # Connect to Excel
                try:
                    app = xw.apps.active
                except:
                    return {"error": "Cannot connect to Excel. Please make sure Excel is running."}
                
                # Find workbook
                wb = None
                for book in app.books:
                    if book.name == workbook_name or book.fullname == workbook_name:
                        wb = book
                        break
                
                if not wb:
                    return {"error": f"Workbook '{workbook_name}' not found"}
                
                wb.activate()
                result["workbook"] = wb.name
                
                # Get all sheets info
                for sheet in wb.sheets:
                    try:
                        used_range = sheet.used_range
                        sheet_info = {
                            "name": sheet.name,
                            "total_rows": used_range.last_cell.row if used_range else 0,
                            "total_columns": used_range.last_cell.column if used_range else 0,
                            "has_data": used_range is not None
                        }
                        all_sheets_info.append(sheet_info)
                    except:
                        sheet_info = {
                            "name": sheet.name,
                            "total_rows": 0,
                            "total_columns": 0,
                            "has_data": False
                        }
                        all_sheets_info.append(sheet_info)
                
                result["sheets_info"] = all_sheets_info
                
                # Get specific sheet data
                if sheet_name:
                    if sheet_name not in [s.name for s in wb.sheets]:
                        return {"error": f"Sheet '{sheet_name}' not found"}
                    sheet = wb.sheets[sheet_name]
                else:
                    sheet = wb.sheets.active
                
                sheet.activate()
                result["sheet"] = sheet.name
                
                # Read data from xlwings
                used_range = sheet.used_range
                if used_range:
                    actual_last_row = used_range.last_cell.row
                    actual_last_col = used_range.last_cell.column
                    
                    # Read all data from row 1 to last row
                    data_range = sheet.range((1, 1), (actual_last_row, actual_last_col))
                    data = data_range.value
                    
                    # Convert to consistent list format
                    if data:
                        if not isinstance(data[0], list):
                            raw_data = [data]
                        else:
                            raw_data = data
                    else:
                        raw_data = []
                
            except Exception as e:
                return {"error": f"Failed to access open workbook: {str(e)}"}
        else:
            return {"error": "Must provide either file_path or workbook_name"}
        
        # Process raw data if we have it
        if raw_data is not None:
            # Clean data function
            def clean_value(value):
                if isinstance(value, str):
                    return value.replace('\x00', '').strip()
                elif isinstance(value, (datetime, date)):
                    return value.isoformat()
                elif pd.isna(value):
                    return None
                return value
            
            # Apply cleaning to all data
            cleaned_data = []
            for row in raw_data:
                if isinstance(row, list):
                    cleaned_row = [clean_value(cell) for cell in row]
                else:
                    cleaned_row = [clean_value(row)]
                cleaned_data.append(cleaned_row)
            
            # Ensure all rows have the same number of columns
            max_cols = max(len(row) for row in cleaned_data) if cleaned_data else 0
            for row in cleaned_data:
                while len(row) < max_cols:
                    row.append(None)
            
            # Set Excel-style column names
            excel_columns = []
            for i in range(max_cols):
                if i < 26:
                    col_letter = chr(65 + i)
                else:
                    col_letter = chr(65 + i // 26 - 1) + chr(65 + i % 26)
                excel_columns.append(col_letter)
            
            # STRUCTURE ANALYSIS - This is the key improvement
            structure_analysis = analyze_data_structure(cleaned_data, excel_columns)
            result["structure_analysis"] = structure_analysis
            
            # Apply start_row and end_row filters for data output
            display_start_idx = max(0, start_row - 1)
            display_end_idx = min(len(cleaned_data), end_row) if end_row else len(cleaned_data)
            display_data = cleaned_data[display_start_idx:display_end_idx]
            
            # Prepare data output with CONSISTENT row numbering
            if analyze_mode in ['data', 'full'] and display_data:
                rows_data = []
                for idx, row_data in enumerate(display_data):
                    actual_excel_row = display_start_idx + idx + 1  # Always 1-based Excel row number
                    row_dict = {}
                    for col_idx, value in enumerate(row_data):
                        if col_idx < len(excel_columns):
                            row_dict[excel_columns[col_idx]] = value
                    
                    rows_data.append({
                        "excel_row": actual_excel_row,
                        "data": row_dict
                    })
                
                result["data"] = {
                    "range_analyzed": f"{start_row}:{display_end_idx}",
                    "rows_shown": len(rows_data),
                    "total_columns": max_cols,
                    "column_names": excel_columns,
                    "rows": rows_data
                }
            
            # Generate recommendations
            recommendations = generate_recommendations(structure_analysis, actual_last_row)
            result["recommendations"] = recommendations
            
            # Column analysis for full mode
            if analyze_mode == 'full' and cleaned_data:
                column_analysis = {}
                for col_idx, col_name in enumerate(excel_columns):
                    col_data = [row[col_idx] if col_idx < len(row) else None for row in cleaned_data]
                    non_null_values = [val for val in col_data if val is not None and val != '']
                    
                    col_analysis = {
                        "total_cells": len(col_data),
                        "non_empty_cells": len(non_null_values),
                        "empty_cells": len(col_data) - len(non_null_values),
                        "data_type": determine_column_type(non_null_values)
                    }
                    
                    # Add statistics for numeric columns
                    if col_analysis["data_type"] == "numeric" and non_null_values:
                        try:
                            numeric_values = [float(val) for val in non_null_values if str(val).replace('.','').replace('-','').isdigit()]
                            if numeric_values:
                                col_analysis["statistics"] = {
                                    "mean": sum(numeric_values) / len(numeric_values),
                                    "min": min(numeric_values),
                                    "max": max(numeric_values),
                                    "count": len(numeric_values)
                                }
                        except:
                            pass
                    
                    column_analysis[col_name] = col_analysis
                
                result["column_analysis"] = column_analysis
        
        return result
        
    except Exception as e:
        return {"error": f"Unexpected error: {str(e)}"}

def analyze_data_structure(data, columns):
    """
    Analyze the structure of Excel data to identify headers, data regions, and empty areas.
    """
    if not data:
        return {
            "status": "empty_sheet",
            "message": "Sheet contains no data",
            "header_row": None,
            "data_start_row": None,
            "data_end_row": None,
            "next_input_row": 1,
            "empty_rows": [],
            "data_rows": []
        }
    
    structure = {
        "total_rows_with_data": len(data),
        "header_row": None,
        "data_start_row": None,
        "data_end_row": None,
        "next_input_row": None,
        "empty_rows": [],
        "data_rows": [],
        "partial_rows": [],
        "structure_type": "unknown"
    }
    
    # Analyze each row
    for row_idx, row in enumerate(data):
        excel_row_num = row_idx + 1
        non_empty_cells = [cell for cell in row if cell is not None and str(cell).strip() != '']
        
        if len(non_empty_cells) == 0:
            structure["empty_rows"].append(excel_row_num)
        elif len(non_empty_cells) < len(row) / 2:  # Less than half filled
            structure["partial_rows"].append({
                "row": excel_row_num,
                "filled_cells": len(non_empty_cells),
                "total_cells": len(row)
            })
        else:
            structure["data_rows"].append(excel_row_num)
    
    # Determine structure type and key positions
    if structure["data_rows"]:
        first_data_row = min(structure["data_rows"])
        last_data_row = max(structure["data_rows"])
        
        structure["data_start_row"] = first_data_row
        structure["data_end_row"] = last_data_row
        
        # Check if first row looks like headers (all text, no numbers)
        if first_data_row == 1:
            first_row = data[0]
            non_empty_first = [cell for cell in first_row if cell is not None and str(cell).strip() != '']
            if non_empty_first and all(isinstance(cell, str) and not str(cell).replace('.','').replace('-','').isdigit() for cell in non_empty_first):
                structure["header_row"] = 1
                structure["data_start_row"] = 2 if len(structure["data_rows"]) > 1 else None
                structure["structure_type"] = "table_with_headers"
            else:
                structure["structure_type"] = "data_table"
        else:
            structure["structure_type"] = "data_with_gaps"
        
        # Determine next input row
        if structure["empty_rows"]:
            # Find first empty row after last data
            next_empty = [row for row in structure["empty_rows"] if row > last_data_row]
            structure["next_input_row"] = next_empty[0] if next_empty else last_data_row + 1
        else:
            structure["next_input_row"] = last_data_row + 1
    else:
        structure["structure_type"] = "empty_or_minimal"
        structure["next_input_row"] = 1
    
    return structure

def determine_column_type(values):
    """Determine the predominant data type in a column."""
    if not values:
        return "empty"
    
    numeric_count = 0
    text_count = 0
    date_count = 0
    
    for val in values:
        if isinstance(val, (int, float)):
            numeric_count += 1
        elif isinstance(val, (datetime, date)):
            date_count += 1
        elif isinstance(val, str):
            if val.replace('.','').replace('-','').isdigit():
                numeric_count += 1
            else:
                text_count += 1
    
    if numeric_count > text_count and numeric_count > date_count:
        return "numeric"
    elif date_count > 0:
        return "datetime"
    else:
        return "text"

def generate_recommendations(structure, total_rows):
    """Generate actionable recommendations based on structure analysis."""
    recommendations = {
        "structure_summary": "",
        "next_action": "",
        "input_guidance": "",
        "warnings": []
    }
    
    if structure.get("structure_type") == "table_with_headers":
        recommendations["structure_summary"] = f"Well-structured table with headers in row {structure['header_row']}"
        recommendations["next_action"] = f"Add new data starting from row {structure['next_input_row']}"
        recommendations["input_guidance"] = f"Headers: Row {structure['header_row']} | Data: Rows {structure['data_start_row']}-{structure['data_end_row']} | Next: Row {structure['next_input_row']}"
    
    elif structure.get("structure_type") == "data_table":
        recommendations["structure_summary"] = f"Data table without clear headers"
        recommendations["next_action"] = f"Add new data starting from row {structure['next_input_row']}"
        recommendations["input_guidance"] = f"Data: Rows {structure['data_start_row']}-{structure['data_end_row']} | Next: Row {structure['next_input_row']}"
    
    elif structure.get("structure_type") == "data_with_gaps":
        recommendations["structure_summary"] = "Data with gaps - inconsistent structure"
        recommendations["next_action"] = f"Consider consolidating data or add to row {structure['next_input_row']}"
        recommendations["input_guidance"] = f"Data scattered across rows, next available: Row {structure['next_input_row']}"
        recommendations["warnings"].append("Data structure has gaps - consider reorganizing")
    
    else:
        recommendations["structure_summary"] = "Empty or minimal data"
        recommendations["next_action"] = "Start entering data from row 1"
        recommendations["input_guidance"] = "Sheet is empty - start from row 1"
    
    if structure.get("empty_rows"):
        empty_count = len(structure["empty_rows"])
        recommendations["warnings"].append(f"{empty_count} empty rows detected")
    
    return recommendations

def main():
    parser = argparse.ArgumentParser(description='Universal Excel data analyzer with structure identification')
    parser.add_argument('--file', help='Path to Excel file (for closed files)')
    parser.add_argument('--workbook', help='Name of open workbook')
    parser.add_argument('--sheet', help='Specific sheet name (optional)')
    parser.add_argument('--start-row', type=int, default=1, help='Starting row (default: 1)')
    parser.add_argument('--end-row', type=int, help='Ending row (optional)')
    parser.add_argument('--mode', choices=['full', 'quick', 'data'], default='full', 
                       help='Analysis mode: full (detailed), quick (basic), data (data only)')
    
    args = parser.parse_args()
    
    # Set stdout encoding to UTF-8 for Windows
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    result = analyze_excel_data(
        file_path=args.file,
        workbook_name=args.workbook,
        sheet_name=args.sheet,
        start_row=args.start_row,
        end_row=args.end_row,
        analyze_mode=args.mode
    )
    
    # Custom JSON serializer
    def json_serial(obj):
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        elif pd.isna(obj):
            return None
        elif hasattr(obj, '__str__'):
            return str(obj)
        raise TypeError(f"Type {type(obj)} not serializable")
    
    # Output result
    output = json.dumps(result, ensure_ascii=False, default=json_serial, indent=2)
    try:
        print(output)
    except UnicodeEncodeError:
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()