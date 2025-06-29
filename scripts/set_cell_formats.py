import xlwings as xw
import argparse
import json
import sys

def set_cell_formats(cell_range, format_config, filename=None, sheet_name=None):
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
        
        # Apply formatting
        applied_formats = []
        
        # Font color
        if 'fontColor' in format_config:
            color = format_config['fontColor']
            if color.startswith('#'):
                color = color[1:]
            rgb = int(color[:2], 16) + (int(color[2:4], 16) << 8) + (int(color[4:6], 16) << 16)
            range_obj.api.Font.Color = rgb
            applied_formats.append(f"font color: {format_config['fontColor']}")
        
        # Background color (Interior color)
        if 'backgroundColor' in format_config:
            color = format_config['backgroundColor']
            if color.startswith('#'):
                color = color[1:]
            rgb = int(color[:2], 16) + (int(color[2:4], 16) << 8) + (int(color[4:6], 16) << 16)
            range_obj.api.Interior.Color = rgb
            applied_formats.append(f"background color: {format_config['backgroundColor']}")
        
        # Bold
        if 'bold' in format_config:
            range_obj.api.Font.Bold = format_config['bold']
            applied_formats.append(f"bold: {format_config['bold']}")
        
        # Italic
        if 'italic' in format_config:
            range_obj.api.Font.Italic = format_config['italic']
            applied_formats.append(f"italic: {format_config['italic']}")
        
        # Underline
        if 'underline' in format_config:
            # Excel underline constants: 1 = Single, 2 = Double, -4142 = None
            range_obj.api.Font.Underline = 1 if format_config['underline'] else -4142
            applied_formats.append(f"underline: {format_config['underline']}")
        
        # Font size
        if 'fontSize' in format_config:
            range_obj.api.Font.Size = format_config['fontSize']
            applied_formats.append(f"font size: {format_config['fontSize']}")
        
        # Font name
        if 'fontName' in format_config:
            range_obj.api.Font.Name = format_config['fontName']
            applied_formats.append(f"font name: {format_config['fontName']}")
        
        # Text alignment (horizontal)
        if 'textAlign' in format_config:
            align_constants = {
                'left': -4131,    # xlLeft
                'center': -4108,  # xlCenter
                'right': -4152    # xlRight
            }
            align_value = align_constants.get(format_config['textAlign'])
            if align_value:
                range_obj.api.HorizontalAlignment = align_value
                applied_formats.append(f"text align: {format_config['textAlign']}")
        
        # Vertical alignment
        if 'verticalAlign' in format_config:
            valign_constants = {
                'top': -4160,     # xlTop
                'middle': -4108,  # xlCenter
                'bottom': -4107   # xlBottom
            }
            valign_value = valign_constants.get(format_config['verticalAlign'])
            if valign_value:
                range_obj.api.VerticalAlignment = valign_value
                applied_formats.append(f"vertical align: {format_config['verticalAlign']}")
        
        # Ensure the changes are visible
        wb.app.screen_updating = True
        
        return {
            "success": True,
            "message": f"Formatting applied to range {cell_range}",
            "applied_formats": applied_formats,
            "range": cell_range,
            "sheet": sheet.name if sheet_name else wb.sheets.active.name
        }
    
    except Exception as e:
        return {"error": f"Failed to set cell formats: {str(e)}"}

def main():
    parser = argparse.ArgumentParser(description='Set cell formatting in Excel')
    parser.add_argument('--range', required=True, help='Cell range (e.g., A1:C3)')
    parser.add_argument('--format', required=True, help='Format configuration as JSON')
    parser.add_argument('--filename', help='Excel workbook name')
    parser.add_argument('--sheet', help='Sheet name')
    
    args = parser.parse_args()
    
    try:
        format_config = json.loads(args.format)
    except json.JSONDecodeError:
        print(json.dumps({"error": "Invalid JSON format for format configuration"}))
        sys.exit(1)
    
    result = set_cell_formats(args.range, format_config, args.filename, args.sheet)
    print(json.dumps(result, ensure_ascii=False))

if __name__ == "__main__":
    main()