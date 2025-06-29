import xlwings as xw
import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime, date

def rgb_to_dict(rgb):
    """Convert RGB tuple to dict with hex"""
    if rgb is None:
        return None
    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    return {
        "r": r,
        "g": g,
        "b": b,
        "hex": "#{:02x}{:02x}{:02x}".format(r, g, b)
    }

def get_border_style(border_style):
    """Get border style as string"""
    if border_style is None or border_style == -4142:  # xlNone
        return "none"
    # Excel border styles
    style_map = {
        1: "continuous",     # xlContinuous
        -4115: "dash",       # xlDash
        4: "dashdot",        # xlDashDot
        5: "dashdotdot",     # xlDashDotDot
        -4118: "dot",        # xlDot
        -4119: "double",     # xlDouble
        13: "slantdashdot",  # xlSlantDashDot
        -4142: "none"        # xlLineStyleNone
    }
    return style_map.get(border_style, "continuous")

def get_border_weight(weight):
    """Get border weight as string"""
    if weight is None:
        return "thin"
    weight_map = {
        1: "hairline",    # xlHairline
        -4138: "medium",  # xlMedium
        4: "thick",       # xlThick
        2: "thin"         # xlThin
    }
    return weight_map.get(weight, "thin")

def get_alignment_string(alignment_code):
    """Convert Excel alignment code to string"""
    h_align_map = {
        -4108: "center",
        -4131: "left",
        -4152: "right",
        -4130: "justify",
        -4117: "distributed"
    }
    v_align_map = {
        -4107: "bottom",
        -4108: "center",
        -4160: "top",
        -4130: "justify",
        -4117: "distributed"
    }
    return h_align_map, v_align_map

def get_cell_formats(start_row=1, start_col=1, end_row=20, end_col=15, filename=None, sheet_name=None):
    try:
        # Get active Excel app
        try:
            app = xw.apps.active
        except:
            return {"error": "Cannot connect to Excel. Please make sure Excel is running."}

        # Get workbook
        if filename:
            wb = None
            for book in app.books:
                if book.name == filename or book.fullname == filename:
                    wb = book
                    break
            if not wb:
                return {"error": f"Workbook '{filename}' not found"}
        else:
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
        
        # Collect cell format data
        cells_data = []
        format_groups = defaultdict(list)  # Group cells by similar formatting
        h_align_map, v_align_map = get_alignment_string(None)
        
        for row in range(start_row, min(end_row + 1, start_row + 20)):  # Max 20 rows
            for col in range(start_col, min(end_col + 1, start_col + 15)):  # Max 15 columns
                try:
                    cell = sheet.cells(row, col)
                    cell_range = cell.address
                    
                    # Get cell value
                    value = cell.value
                    if value is not None:
                        # Convert datetime objects to ISO format string
                        if isinstance(value, (datetime, date)):
                            value = value.isoformat()
                        elif isinstance(value, float) and value.is_integer():
                            value = int(value)
                        elif isinstance(value, str):
                            # No encoding conversion needed - xlwings returns proper Unicode
                            # Just clean up null bytes and normalize whitespace
                            value = value.replace('\x00', '').strip()
                    
                    # Get border information
                    borders = {}
                    for border_name, border_idx in [("top", 3), ("bottom", 4), ("left", 1), ("right", 2)]:
                        border = cell.api.Borders(border_idx)
                        border_style = get_border_style(border.LineStyle)
                        if border_style != "none":
                            borders[border_name] = {
                                "style": border_style,
                                "weight": get_border_weight(border.Weight),
                                "color": rgb_to_dict(border.Color) if hasattr(border, 'Color') and border.Color else {"hex": "#000000"}
                            }
                    
                    # Get background color
                    bg_color = None
                    if cell.color:
                        bg_color = rgb_to_dict(cell.color)
                    
                    # Get font color
                    font_color = rgb_to_dict(cell.font.color) if cell.font.color else {"r": 0, "g": 0, "b": 0, "hex": "#000000"}
                    
                    # Get alignment
                    h_align = h_align_map.get(cell.api.HorizontalAlignment, "general")
                    v_align = v_align_map.get(cell.api.VerticalAlignment, "bottom")
                    
                    # Build cell data
                    cell_data = {
                        "address": cell_range,
                        "value": value,
                        "format": {
                            "background_color": bg_color,
                            "font": {
                                "color": font_color,
                                "size": cell.font.size,
                                "bold": cell.font.bold,
                                "italic": cell.font.italic,
                                "name": cell.font.name
                            },
                            "alignment": {
                                "horizontal": h_align,
                                "vertical": v_align,
                                "wrap_text": cell.api.WrapText
                            },
                            "borders": borders,
                            "number_format": cell.number_format
                        }
                    }
                    
                    # Only include cells with content or formatting
                    if (value is not None or bg_color is not None or borders):
                        cells_data.append(cell_data)
                        
                        # Create format key for grouping
                        format_key = json.dumps({
                            "bg": bg_color["hex"] if bg_color else None,
                            "font_color": font_color["hex"],
                            "font_bold": cell.font.bold,
                            "h_align": h_align,
                            "v_align": v_align,
                            "borders": bool(borders)
                        }, sort_keys=True)
                        format_groups[format_key].append(cell_range)
                        
                except Exception as e:
                    # Skip cells that cause errors
                    continue
        
        # Get merged cells information
        merged_ranges = []
        try:
            for area in sheet.api.UsedRange.MergeArea:
                if area.MergeCells:
                    merged_ranges.append(area.Address)
        except:
            pass  # No merged cells or error accessing them
        
        # Detect format patterns
        format_patterns = {}
        pattern_id = 1
        for format_key, addresses in format_groups.items():
            if len(addresses) > 1:  # Only include patterns with multiple cells
                format_info = json.loads(format_key)
                
                # Try to find continuous ranges
                min_row, max_row = float('inf'), 0
                min_col, max_col = float('inf'), 0
                
                for addr in addresses:
                    # Parse address (e.g., "A1" -> row=1, col=1)
                    col_str = ''.join(c for c in addr if c.isalpha()).replace('$', '')
                    row_num = int(''.join(c for c in addr if c.isdigit()))
                    col_num = ord(col_str) - ord('A') + 1 if len(col_str) == 1 else 0
                    
                    min_row = min(min_row, row_num)
                    max_row = max(max_row, row_num)
                    min_col = min(min_col, col_num)
                    max_col = max(max_col, col_num)
                
                # Create pattern description
                pattern_desc = f"pattern_{pattern_id}"
                if format_info["font_bold"] and format_info["bg"]:
                    pattern_desc = "header_row"
                elif format_info["borders"]:
                    pattern_desc = "data_rows"
                
                col_letter = chr(ord('A') + min_col - 1)
                end_col_letter = chr(ord('A') + max_col - 1)
                
                format_patterns[pattern_desc] = {
                    "range": f"{col_letter}{min_row}:{end_col_letter}{max_row}",
                    "background_color": format_info["bg"],
                    "font_color": format_info["font_color"],
                    "font_bold": format_info["font_bold"],
                    "alignment": format_info["h_align"],
                    "has_borders": format_info["borders"]
                }
                pattern_id += 1
        
        result = {
            "success": True,
            "workbook": wb.name,
            "sheet": sheet.name,
            "range": f"{sheet.cells(start_row, start_col).address}:{sheet.cells(end_row, end_col).address}",
            "cells": cells_data,
            "merged_cells": merged_ranges,
            "format_patterns": format_patterns
        }
        
        return result
        
    except Exception as e:
        return {"error": f"Error getting cell formats: {str(e)}"}

def main():
    parser = argparse.ArgumentParser(description='Get cell formats and layout information from Excel')
    parser.add_argument('--start-row', type=int, default=1, help='Starting row (default: 1)')
    parser.add_argument('--start-col', type=int, default=1, help='Starting column (default: 1)')
    parser.add_argument('--end-row', type=int, default=20, help='Ending row (default: 20, max: 20)')
    parser.add_argument('--end-col', type=int, default=15, help='Ending column (default: 15, max: 15)')
    parser.add_argument('--filename', help='Optional Excel filename')
    parser.add_argument('--sheet', help='Optional sheet name')
    
    args = parser.parse_args()
    
    # Set stdout encoding to UTF-8 for Windows
    import sys
    import io
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # Ensure max limits
    end_row = min(args.end_row, args.start_row + 34)  # Max 35 rows
    end_col = min(args.end_col, args.start_col + 14)  # Max 15 columns
    
    result = get_cell_formats(
        args.start_row, 
        args.start_col, 
        end_row, 
        end_col,
        args.filename, 
        args.sheet
    )
    # Custom JSON encoder for datetime objects
    def json_serial(obj):
        """JSON serializer for objects not serializable by default json code"""
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        raise TypeError(f"Type {type(obj)} not serializable")
    
    # Ensure UTF-8 output to handle Unicode characters
    output = json.dumps(result, ensure_ascii=False, indent=2, default=json_serial)
    try:
        print(output)
    except UnicodeEncodeError:
        # Fallback: encode to UTF-8 and write to stdout buffer
        sys.stdout.buffer.write(output.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')

if __name__ == "__main__":
    main()