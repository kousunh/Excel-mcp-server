import xlwings as xw
import argparse
import sys
import json
import time
import random
import re

def check_vba_access(app):
    """Check and enable VBA project access if needed"""
    try:
        # Try to access VBProject to test permissions
        test_wb = app.books.active
        if test_wb:
            _ = test_wb.api.VBProject.Name
        return True
    except Exception as e:
        error_str = str(e)
        if "VBProject" in error_str or "プログラムによるアクセス" in error_str:
            try:
                # Try to enable VBA access programmatically
                app.api.Application.VBE.MainWindow.Visible = False
                # Set trust access to VBA project object model
                app.api.Application.EnableEvents = True
                return True
            except:
                return False
        return False

def check_macro_settings(app):
    """Check macro security settings"""
    try:
        # Check if macros are enabled
        security_level = app.api.Application.AutomationSecurity
        # 1 = msoAutomationSecurityLow, 2 = msoAutomationSecurityByUI, 3 = msoAutomationSecurityForceDisable
        if security_level == 3:
            return False, "Macros are disabled by security settings"
        return True, ""
    except:
        # If we can't check, assume it's okay
        return True, ""

def generate_unique_names(wb, base_module="TempModule", base_procedure="Main"):
    """Generate unique module and procedure names"""
    timestamp = int(time.time() * 1000) % 100000
    random_num = random.randint(1000, 9999)
    
    module_name = f"{base_module}{timestamp}"
    procedure_name = f"{base_procedure}{random_num}"
    
    # Ensure uniqueness by checking existing modules
    existing_modules = []
    try:
        for comp in wb.api.VBProject.VBComponents:
            existing_modules.append(comp.Name)
    except:
        pass
    
    counter = 1
    original_module = module_name
    while module_name in existing_modules:
        module_name = f"{original_module}_{counter}"
        counter += 1
    
    return module_name, procedure_name

def clean_existing_temp_modules(wb):
    """Clean up old temporary modules"""
    try:
        components_to_remove = []
        for comp in wb.api.VBProject.VBComponents:
            if comp.Name.startswith("TempModule"):
                components_to_remove.append(comp)
        
        for comp in components_to_remove:
            try:
                wb.api.VBProject.VBComponents.Remove(comp)
                time.sleep(0.1)  # Small delay between removals
            except:
                continue
    except:
        pass

def parse_vba_code(vba_code):
    """Parse VBA code to detect if it already has Sub/Function structure"""
    # Remove comments and clean up
    lines = vba_code.split('\n')
    clean_lines = []
    for line in lines:
        # Remove comments (starting with ')
        if "'" in line:
            line = line.split("'")[0]
        clean_lines.append(line.strip())
    
    clean_code = '\n'.join(clean_lines)
    
    # Check for Sub or Function declarations
    sub_pattern = r'\b(Sub|Function)\s+(\w+)'
    matches = re.findall(sub_pattern, clean_code, re.IGNORECASE)
    
    if matches:
        # Extract the first procedure name
        procedure_type, procedure_name = matches[0]
        return True, procedure_name, clean_code
    
    return False, None, clean_code

def execute_vba(vba_code, module_name='TempModule', procedure_name='Main', filename=None, sheet_name=None, max_retries=3):
    app = None
    wb = None
    vba_module = None
    final_module_name = module_name
    final_procedure_name = procedure_name
    
    # Retry logic with increasing delays
    for attempt in range(max_retries):
        try:
            # Connect to Excel
            try:
                app = xw.apps.active
                if not app:
                    return {"error": "Excel is not running. Please open Excel first."}
            except Exception as e:
                return {"error": f"Cannot connect to Excel: {str(e)}"}
        
            # Check VBA access permissions
            if not check_vba_access(app):
                return {"error": "VBA access denied. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings)."}
            
            # Check macro security settings
            macro_ok, macro_error = check_macro_settings(app)
            if not macro_ok:
                return {"error": f"Macro security issue: {macro_error}. Please enable macros in Excel Trust Center settings."}
            
            # Get workbook
            wb = None
            if filename:
                for book in app.books:
                    if filename in book.name or filename == book.name:
                        wb = book
                        break
                if not wb:
                    return {"error": f"Workbook '{filename}' not found. Available workbooks: {[book.name for book in app.books]}"}
            else:
                wb = app.books.active
                if not wb:
                    return {"error": "No active workbook found. Please open an Excel file."}
            
            # Activate the workbook
            try:
                wb.activate()
                time.sleep(0.2)  # Wait for activation
            except Exception as e:
                return {"error": f"Cannot activate workbook: {str(e)}"}
            
            # Navigate to sheet if specified
            if sheet_name:
                try:
                    sheet = wb.sheets[sheet_name]
                    sheet.activate()
                    time.sleep(0.1)
                except Exception as e:
                    return {"error": f"Cannot navigate to sheet '{sheet_name}': {str(e)}"}
            
            # Parse VBA code
            has_structure, detected_proc_name, clean_code = parse_vba_code(vba_code)
            
            # Generate unique names for this attempt
            final_module_name, final_procedure_name = generate_unique_names(wb, module_name, procedure_name)
            
            # Use detected procedure name if available
            if has_structure and detected_proc_name:
                final_procedure_name = detected_proc_name
            
            # Clean up old temporary modules (only on first attempt)
            if attempt == 0:
                clean_existing_temp_modules(wb)
                time.sleep(0.5)
            
            # Prepare VBA code
            if not has_structure:
                # Wrap in Sub if needed
                final_code = f"Sub {final_procedure_name}()\n{clean_code}\nEnd Sub"
            else:
                final_code = clean_code
            
            # Create and execute VBA module
            try:
                # Create new module
                vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
                vba_module.Name = final_module_name
                time.sleep(0.2)  # Wait for module creation
                
                # Add code to module
                vba_module.CodeModule.AddFromString(final_code)
                time.sleep(0.2)  # Wait for code addition
                
                # Execute the procedure
                full_procedure_name = f"{final_module_name}.{final_procedure_name}"
                wb.api.Application.Run(full_procedure_name)
                
                # Success - clean up and return
                try:
                    wb.api.VBProject.VBComponents.Remove(vba_module)
                except:
                    pass
                
                return {
                    "success": True, 
                    "message": "VBA code executed successfully",
                    "module_name": final_module_name, 
                    "procedure_name": final_procedure_name,
                    "attempt": attempt + 1
                }
                
            except Exception as e:
                error_msg = str(e)
                
                # Clean up module if it was created
                try:
                    if vba_module:
                        wb.api.VBProject.VBComponents.Remove(vba_module)
                        vba_module = None
                except:
                    pass
                
                # If this is not the last attempt, continue to retry
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 0.5  # Increasing delay
                    time.sleep(wait_time)
                    continue
                
                # Last attempt failed - return detailed error
                if "実行できません" in error_msg or "cannot be run" in error_msg.lower():
                    return {"error": f"VBA execution failed: The macro cannot be run. This may be due to macro security settings or VBA project protection. Original error: {error_msg}"}
                elif "VBProject" in error_msg:
                    return {"error": f"VBA project access error: {error_msg}. Please enable VBA project access in Trust Center settings."}
                elif "マクロ" in error_msg:
                    return {"error": f"Macro error: {error_msg}. Please check macro security settings."}
                else:
                    return {"error": f"VBA execution error after {max_retries} attempts: {error_msg}"}
        
        except Exception as outer_e:
            error_msg = str(outer_e)
            
            # Clean up on any error
            try:
                if vba_module and wb:
                    wb.api.VBProject.VBComponents.Remove(vba_module)
                    vba_module = None
            except:
                pass
            
            # If this is not the last attempt, continue to retry
            if attempt < max_retries - 1:
                wait_time = (attempt + 1) * 0.5
                time.sleep(wait_time)
                continue
            
            # Last attempt failed
            if "Excel not running" in error_msg:
                return {"error": "Excel is not running. Please open Excel first."}
            elif "No workbook" in error_msg:
                return {"error": "No active workbook found. Please open an Excel file."}
            else:
                return {"error": f"Connection error after {max_retries} attempts: {error_msg}"}
    
    # Should not reach here
    return {"error": "Unexpected error: Failed all retry attempts"}

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
        
        # Ensure UTF-8 output to handle Unicode characters
        output = json.dumps(result, ensure_ascii=False)
        try:
            print(output)
        except UnicodeEncodeError:
            # Fallback: encode to UTF-8 and write to stdout buffer
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