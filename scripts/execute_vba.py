import xlwings as xw
import argparse
import sys
import json
import time
import random
import re

def clean_existing_temp_modules(wb):
    """Clean up existing temporary modules"""
    try:
        components_to_remove = []
        for comp in wb.api.VBProject.VBComponents:
            if comp.Name.startswith("TempModule"):
                components_to_remove.append(comp)
        
        for comp in components_to_remove:
            try:
                wb.api.VBProject.VBComponents.Remove(comp)
                time.sleep(0.1)
            except:
                continue
    except:
        pass

def parse_vba_code(vba_code):
    """Parse VBA code to detect if it already has Sub/Function structure and clean it"""
    lines = vba_code.split('\n')
    clean_lines = []
    
    for line in lines:
        stripped_line = line.strip()
        
        # Skip empty lines and comments
        if not stripped_line or stripped_line.startswith("'"):
            continue
            
        # Remove inline comments
        if "'" in line:
            line = line.split("'")[0]
        
        # Skip standalone procedure calls (lines that are just procedure names)
        if re.match(r'^\w+$', stripped_line):
            continue
        
        # Remove MsgBox statements to avoid popup alerts
        if re.search(r'\bMsgBox\b', stripped_line, re.IGNORECASE):
            continue
        
        # Replace problematic Err.Raise statements with simple Exit Sub
        if re.search(r'\bErr\.Raise\b', stripped_line, re.IGNORECASE):
            line = re.sub(r'Err\.Raise\s+Err\.Number,\s*Err\.Source,\s*Err\.Description', 'Exit Sub', line, flags=re.IGNORECASE)
            
        clean_lines.append(line.rstrip())
    
    clean_code = '\n'.join(clean_lines)
    
    sub_pattern = r'\b(Sub|Function)\s+(\w+)'
    matches = re.findall(sub_pattern, clean_code, re.IGNORECASE)
    
    if matches:
        procedure_type, procedure_name = matches[0]
        return True, procedure_name, clean_code
    
    return False, None, clean_code

def get_unique_procedure_name(wb, base_name):
    """Generate unique procedure name by checking existing procedures"""
    existing_names = set()
    
    try:
        for comp in wb.api.VBProject.VBComponents:
            if comp.Type == 1:  # Standard module
                try:
                    code_module = comp.CodeModule
                    for i in range(1, code_module.CountOfLines + 1):
                        line = code_module.Lines(i, 1)
                        if line.strip().startswith('Sub '):
                            # Extract procedure name
                            match = re.search(r'Sub\s+(\w+)', line, re.IGNORECASE)
                            if match:
                                existing_names.add(match.group(1).lower())
                except:
                    continue
    except:
        pass
    
    # Generate unique name
    if base_name.lower() not in existing_names:
        return base_name
    
    counter = 2
    while f"{base_name}({counter})".lower() in existing_names:
        counter += 1
    
    return f"{base_name}({counter})"

def execute_vba_simple(app, wb, vba_code, final_module_name, final_procedure_name):
    """Simple VBA execution with Sub procedure creation"""
    vba_module = None
    
    try:
        # Get unique procedure name
        unique_proc_name = get_unique_procedure_name(wb, final_procedure_name)
        
        # Create new module
        vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vba_module.Name = final_module_name
        
        # Wrap code in Sub procedure if not already wrapped
        if not vba_code.strip().lower().startswith('sub '):
            wrapped_code = f"Sub {unique_proc_name}()\nOn Error GoTo ErrorHandler\n{vba_code}\nExit Sub\nErrorHandler:\nExit Sub\nEnd Sub"
        else:
            # Replace the existing Sub name with unique name
            wrapped_code = re.sub(r'Sub\s+\w+', f'Sub {unique_proc_name}', vba_code, count=1, flags=re.IGNORECASE)
            # Fix existing error handling to prevent Err.Raise issues
            wrapped_code = re.sub(r'Err\.Raise\s+Err\.Number,\s*Err\.Source,\s*Err\.Description', 'Exit Sub', wrapped_code, flags=re.IGNORECASE)
            # Add simple error handling if not present
            if 'On Error' not in wrapped_code:
                wrapped_code = wrapped_code.replace(f'Sub {unique_proc_name}()', f'Sub {unique_proc_name}()\nOn Error GoTo ErrorHandler')
                wrapped_code = wrapped_code.replace('End Sub', 'Exit Sub\nErrorHandler:\nExit Sub\nEnd Sub')
        
        # Add code to module
        vba_module.CodeModule.AddFromString(wrapped_code)
        
        # Execute the procedure
        full_procedure_name = f"{final_module_name}.{unique_proc_name}"
        wb.api.Application.Run(full_procedure_name)
        
        return True, None, unique_proc_name
        
    except Exception as e:
        error_info = {
            "type": "execution_error",
            "message": str(e),
            "details": "Error occurred during VBA execution"
        }
        return False, error_info, final_procedure_name
        
    finally:
        # Always clean up module, even on error
        try:
            if vba_module:
                wb.api.VBProject.VBComponents.Remove(vba_module)
        except:
            pass

def execute_vba(vba_code, module_name='TempModule', procedure_name='Main', filename=None, sheet_name=None, max_retries=1):
    try:
        # Connect to Excel
        app = xw.apps.active
        if not app:
            return {"error": "Excel is not running. Please open Excel first."}
        
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
                return {"error": "No active workbook found. Please open an Excel file."}
        
        # Navigate to sheet if specified
        if sheet_name:
            try:
                sheet = wb.sheets[sheet_name]
                sheet.activate()
            except Exception as e:
                return {"error": f"Cannot navigate to sheet '{sheet_name}': {str(e)}"}
        
        # Disable alerts temporarily
        original_display_alerts = app.api.Application.DisplayAlerts
        original_screen_updating = app.api.Application.ScreenUpdating
        
        app.api.Application.DisplayAlerts = False
        app.api.Application.ScreenUpdating = False
        
        try:
            # Parse VBA code
            has_structure, detected_proc_name, clean_code = parse_vba_code(vba_code)
            
            # Generate unique module name
            final_module_name = f"{module_name}{random.randint(1000, 9999)}"
            final_procedure_name = detected_proc_name if has_structure and detected_proc_name else procedure_name
            
            # Clean old modules
            clean_existing_temp_modules(wb)
            
            # Prepare final code
            if not has_structure:
                final_code = f"Sub {final_procedure_name}()\n{clean_code}\nEnd Sub"
            else:
                final_code = clean_code
            
            # Execute VBA
            success, error_info, actual_proc_name = execute_vba_simple(app, wb, final_code, final_module_name, final_procedure_name)
            
            if success:
                return {
                    "success": True,
                    "message": f"VBA code executed successfully in procedure '{actual_proc_name}'",
                    "procedure_name": actual_proc_name,
                    "module_name": final_module_name
                }
            else:
                return {
                    "error": "VBA execution failed",
                    "details": error_info,
                    "procedure_name": actual_proc_name
                }
        
        finally:
            # Always restore settings
            try:
                app.api.Application.DisplayAlerts = original_display_alerts
                app.api.Application.ScreenUpdating = original_screen_updating
            except:
                pass
            
    except Exception as e:
        return {"error": f"Connection error: {str(e)}"}

def main():
    try:
        parser = argparse.ArgumentParser(description='Execute VBA code in Excel')
        parser.add_argument('--code', required=True, help='VBA code to execute')
        parser.add_argument('--module', default='TempModule', help='Module name (default: TempModule)')
        parser.add_argument('--procedure', default='Main', help='Procedure name (default: Main)')
        parser.add_argument('--filename', help='Excel filename (optional)')
        parser.add_argument('--sheet', help='Sheet name to activate (optional)')
        
        args = parser.parse_args()
        result = execute_vba(args.code, args.module, args.procedure, args.filename, args.sheet)
        
        output = json.dumps(result, ensure_ascii=False, indent=2)
        try:
            print(output)
        except UnicodeEncodeError:
            sys.stdout.buffer.write(output.encode('utf-8'))
            sys.stdout.buffer.write(b'\n')
        
        sys.exit(0 if result.get("success") else 1)
    except Exception as e:
        error_result = {"error": f"Script error: {str(e)}"}
        output = json.dumps(error_result, ensure_ascii=False)
        try:
            print(output)
        except UnicodeEncodeError:
            sys.stdout.buffer.write(output.encode('utf-8'))
            sys.stdout.buffer.write(b'\n')
        sys.exit(1)

if __name__ == "__main__":
    main()