import xlwings as xw
import argparse
import json
import sys

def set_cell_borders(cell_range, borders, filename=None, sheet_name=None):
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
        
        wb.activate()
        
        # Navigate to sheet if specified
        if sheet_name:
            try:
                sheet = wb.sheets[sheet_name]
                sheet.activate()
            except:
                return {"error": f"Sheet '{sheet_name}' not found"}
        else:
            sheet = wb.sheets.active
        
        # Select the range
        try:
            range_obj = sheet.range(cell_range)
        except:
            return {"error": f"Invalid range: {cell_range}"}
        
        # Excel border constants
        border_styles = {
            'none': 0,        # xlNone
            'thin': 1,        # xlThin
            'medium': -4138,  # xlMedium
            'thick': 4,       # xlThick
            'double': -4119,  # xlDouble
            'dotted': -4118,  # xlDot
            'dashed': -4115   # xlDash
        }
        
        # Excel border position constants
        border_positions = {
            'left': 7,      # xlEdgeLeft
            'right': 10,    # xlEdgeRight
            'top': 8,       # xlEdgeTop
            'bottom': 9,    # xlEdgeBottom
            'inside_vertical': 11,   # xlInsideVertical
            'inside_horizontal': 12  # xlInsideHorizontal
        }
        
        # Apply borders
        for position, border_config in borders.items():
            if not border_config:
                continue
                
            style = border_config.get('style', 'thin')
            color = border_config.get('color', '#000000')
            
            # Convert hex color to RGB integer
            if color.startswith('#'):
                color = color[1:]
            rgb = int(color[:2], 16) + (int(color[2:4], 16) << 8) + (int(color[4:6], 16) << 16)
            
            if position == 'outside':
                # Apply to all outside borders
                for pos in ['left', 'right', 'top', 'bottom']:
                    if pos in border_positions:
                        border = range_obj.api.Borders(border_positions[pos])
                        border.LineStyle = border_styles.get(style, 1)
                        if style != 'none':
                            border.Color = rgb
            elif position == 'inside':
                # Apply to inside borders
                if len(range_obj.address.split(':')) > 1:  # Only if it's a range
                    for pos in ['inside_vertical', 'inside_horizontal']:
                        if pos in border_positions:
                            border = range_obj.api.Borders(border_positions[pos])
                            border.LineStyle = border_styles.get(style, 1)
                            if style != 'none':
                                border.Color = rgb
            else:
                # Apply to specific border
                if position in border_positions:
                    border = range_obj.api.Borders(border_positions[position])
                    border.LineStyle = border_styles.get(style, 1)
                    if style != 'none':
                        border.Color = rgb
        
        # Ensure the changes are visible
        wb.app.screen_updating = True
        
        return {
            "success": True,
            "message": f"Borders applied to range {cell_range}",
            "range": cell_range,
            "sheet": sheet.name if sheet_name else wb.sheets.active.name
        }
    
    except Exception as e:
        return {"error": f"Failed to set borders: {str(e)}"}

def main():
    parser = argparse.ArgumentParser(description='Set cell borders in Excel')
    parser.add_argument('--range', required=True, help='Cell range (e.g., A1:C3)')
    parser.add_argument('--borders', required=True, help='Border configuration as JSON')
    parser.add_argument('--filename', help='Excel workbook name')
    parser.add_argument('--sheet', help='Sheet name')
    
    args = parser.parse_args()
    
    try:
        borders = json.loads(args.borders)
    except json.JSONDecodeError:
        print(json.dumps({"error": "Invalid JSON format for borders"}))
        sys.exit(1)
    
    result = set_cell_borders(args.range, borders, args.filename, args.sheet)
    print(json.dumps(result, ensure_ascii=False))

if __name__ == "__main__":
    main()