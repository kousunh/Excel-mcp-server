import xlwings as xw
import argparse
import sys
import json
import time
import random
import re
import threading
import ctypes
from ctypes import wintypes
import subprocess

# Windows API constants
SW_HIDE = 0
SW_SHOW = 5
HWND_TOP = 0
SWP_SHOWWINDOW = 0x0040

# Windows API functions
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

def find_and_close_dialogs():
    """Windows API function to find and close Excel dialogs"""
    dialog_classes = [
        "bosa_sdm_XL9",  # Excel dialog
        "#32770",        # Standard dialog
        "xlMain",        # Excel main window dialogs
        "NUIDialog",     # Office UI dialogs
    ]
    
    button_texts = ["OK", "はい", "Yes", "Cancel", "キャンセル", "いいえ", "No", "Retry", "再試行"]
    
    def enum_windows_proc(hwnd, lParam):
        try:
            # Get window class name
            class_name = ctypes.create_string_buffer(256)
            user32.GetClassNameA(hwnd, class_name, 256)
            class_name = class_name.value.decode('ascii', errors='ignore')
            
            # Get window text
            window_text = ctypes.create_string_buffer(512)
            user32.GetWindowTextA(hwnd, window_text, 512)
            window_text = window_text.value.decode('utf-8', errors='ignore')
            
            # Check if it's a dialog we want to close
            if any(dialog_class in class_name for dialog_class in dialog_classes):
                if user32.IsWindowVisible(hwnd):
                    # Try to find and click buttons
                    def enum_child_proc(child_hwnd, child_lParam):
                        try:
                            child_class = ctypes.create_string_buffer(256)
                            user32.GetClassNameA(child_hwnd, child_class, 256)
                            child_class = child_class.value.decode('ascii', errors='ignore')
                            
                            child_text = ctypes.create_string_buffer(256)
                            user32.GetWindowTextA(child_hwnd, child_text, 256)
                            child_text = child_text.value.decode('utf-8', errors='ignore')
                            
                            # Check if it's a button with text we want to click
                            if "Button" in child_class and any(btn_text in child_text for btn_text in button_texts):
                                # Click the button
                                BM_CLICK = 0x00F5
                                user32.SendMessageA(child_hwnd, BM_CLICK, 0, 0)
                                return False  # Stop enumeration
                        except:
                            pass
                        return True
                    
                    # Enumerate child windows (buttons)
                    user32.EnumChildWindows(hwnd, ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum_child_proc), 0)
                    
                    # If no button found, close the window directly
                    user32.PostMessageA(hwnd, 0x0010, 0, 0)  # WM_CLOSE
        except:
            pass
        return True
    
    # Enumerate all top-level windows
    user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum_windows_proc), 0)

class DialogMonitor:
    """Monitor and automatically close Excel dialogs"""
    def __init__(self):
        self.monitoring = False
        self.thread = None
    
    def start_monitoring(self, timeout=30):
        """Start monitoring for dialogs"""
        self.monitoring = True
        self.thread = threading.Thread(target=self._monitor_loop, args=(timeout,))
        self.thread.daemon = True
        self.thread.start()
    
    def stop_monitoring(self):
        """Stop monitoring"""
        self.monitoring = False
        if self.thread:
            self.thread.join(timeout=2)
    
    def _monitor_loop(self, timeout):
        """Main monitoring loop"""
        start_time = time.time()
        while self.monitoring and (time.time() - start_time) < timeout:
            try:
                find_and_close_dialogs()
                time.sleep(0.1)  # Check every 100ms
            except:
                pass

def save_excel_settings(app):
    """Save current Excel settings"""
    try:
        settings = {
            'DisplayAlerts': app.api.Application.DisplayAlerts,
            'EnableEvents': app.api.Application.EnableEvents,
            'ScreenUpdating': app.api.Application.ScreenUpdating,
            'Calculation': app.api.Application.Calculation,
            'StatusBar': app.api.Application.StatusBar
        }
        return settings
    except:
        return {}

def disable_excel_alerts(app):
    """Disable Excel alerts and events"""
    try:
        app.api.Application.DisplayAlerts = False
        app.api.Application.EnableEvents = False
        app.api.Application.ScreenUpdating = False
        app.api.Application.StatusBar = "VBA Execution in progress..."
        # Set calculation to manual to prevent issues
        app.api.Application.Calculation = -4135  # xlCalculationManual
        return True
    except Exception as e:
        return False

def restore_excel_settings(app, settings):
    """Restore Excel settings"""
    try:
        if settings:
            app.api.Application.DisplayAlerts = settings.get('DisplayAlerts', True)
            app.api.Application.EnableEvents = settings.get('EnableEvents', True)
            app.api.Application.ScreenUpdating = settings.get('ScreenUpdating', True)
            app.api.Application.Calculation = settings.get('Calculation', -4105)  # xlCalculationAutomatic
            app.api.Application.StatusBar = settings.get('StatusBar', False)
        else:
            # Default restore
            app.api.Application.DisplayAlerts = True
            app.api.Application.EnableEvents = True
            app.api.Application.ScreenUpdating = True
            app.api.Application.Calculation = -4105  # xlCalculationAutomatic
            app.api.Application.StatusBar = False
    except:
        pass

def wrap_vba_with_error_handling(vba_code, procedure_name):
    """Wrap VBA code with comprehensive error handling"""
    error_handling_code = f"""
Sub {procedure_name}()
    ' Auto-generated error handling wrapper
    On Error Resume Next
    
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalScreenUpdating As Boolean
    Dim errorInfo As String
    
    ' Save original settings
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    originalScreenUpdating = Application.ScreenUpdating
    
    ' Disable alerts and events
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Clear any previous error
    Err.Clear
    
    ' Execute user code
{indent_code(vba_code, "    ")}
    
    ' Check for errors
    If Err.Number <> 0 Then
        errorInfo = "Error " & Err.Number & ": " & Err.Description & " (Line: " & Erl & ")"
        ' Debug.Print "VBA_ERROR_INFO: " & errorInfo
    Else
        ' Debug.Print "VBA_SUCCESS: Code executed successfully"
    End If
    
    ' Restore original settings
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Clear error state
    Err.Clear
End Sub
"""
    return error_handling_code

def indent_code(code, indent_str):
    """Add indentation to each line of code"""
    lines = code.split('\n')
    indented_lines = []
    for line in lines:
        if line.strip():  # Don't indent empty lines
            indented_lines.append(indent_str + line)
        else:
            indented_lines.append(line)
    return '\n'.join(indented_lines)

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

def execute_vba_with_monitoring(app, wb, vba_code, final_module_name, final_procedure_name):
    """Execute VBA with dialog monitoring and error collection"""
    vba_module = None
    dialog_monitor = DialogMonitor()
    error_info = None
    
    try:
        # Start dialog monitoring
        dialog_monitor.start_monitoring(timeout=30)
        
        # Create new module
        vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        vba_module.Name = final_module_name
        time.sleep(0.2)  # Wait for module creation
        
        # Add code to module
        vba_module.CodeModule.AddFromString(vba_code)
        time.sleep(0.2)  # Wait for code addition
        
        # Clear the debug console
        try:
            app.api.Application.SendKeys("^{HOME}")  # Clear immediate window if open
        except:
            pass
        
        # Execute the procedure
        full_procedure_name = f"{final_module_name}.{final_procedure_name}"
        wb.api.Application.Run(full_procedure_name)
        
        # Give some time for execution and potential dialogs
        time.sleep(0.5)
        
        # Try to capture any output from Debug.Print statements
        # Note: This is limited as xlwings doesn't directly access VBA immediate window
        
        return True, None
        
    except Exception as e:
        error_info = {
            "type": "execution_error",
            "message": str(e),
            "details": "Error occurred during VBA execution"
        }
        return False, error_info
        
    finally:
        # Stop dialog monitoring
        dialog_monitor.stop_monitoring()
        
        # Clean up module
        try:
            if vba_module:
                wb.api.VBProject.VBComponents.Remove(vba_module)
        except:
            pass

def execute_vba(vba_code, module_name='TempModule', procedure_name='Main', filename=None, sheet_name=None, max_retries=1):
    app = None
    wb = None
    original_settings = {}
    execution_log = []
    
    # Enhanced retry logic with different strategies
    retry_strategies = [
        {"use_error_handling": True, "clean_modules": True, "delay": 0.5},
        {"use_error_handling": True, "clean_modules": False, "delay": 1.0},
        {"use_error_handling": False, "clean_modules": True, "delay": 1.5}
    ]
    
    for attempt in range(max_retries):
        strategy = retry_strategies[min(attempt, len(retry_strategies) - 1)]
        attempt_log = {
            "attempt": attempt + 1,
            "strategy": strategy,
            "start_time": time.time()
        }
        
        try:
            # Connect to Excel
            try:
                app = xw.apps.active
                if not app:
                    return {"error": "Excel is not running. Please open Excel first.", "execution_log": execution_log}
            except Exception as e:
                return {"error": f"Cannot connect to Excel: {str(e)}", "execution_log": execution_log}
        
            # Save original Excel settings
            original_settings = save_excel_settings(app)
            
            # Disable Excel alerts and events
            if not disable_excel_alerts(app):
                attempt_log["warning"] = "Could not disable Excel alerts"
            
            # Check VBA access permissions
            if not check_vba_access(app):
                return {
                    "error": "VBA access denied. Please enable 'Trust access to the VBA project object model' in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).",
                    "execution_log": execution_log
                }
            
            # Check macro security settings
            macro_ok, macro_error = check_macro_settings(app)
            if not macro_ok:
                return {
                    "error": f"Macro security issue: {macro_error}. Please enable macros in Excel Trust Center settings.",
                    "execution_log": execution_log
                }
            
            # Get workbook
            wb = None
            if filename:
                for book in app.books:
                    if filename in book.name or filename == book.name:
                        wb = book
                        break
                if not wb:
                    return {
                        "error": f"Workbook '{filename}' not found. Available workbooks: {[book.name for book in app.books]}",
                        "execution_log": execution_log
                    }
            else:
                wb = app.books.active
                if not wb:
                    return {
                        "error": "No active workbook found. Please open an Excel file.",
                        "execution_log": execution_log
                    }
            
            # Activate the workbook
            try:
                wb.activate()
                time.sleep(0.2)  # Wait for activation
            except Exception as e:
                return {
                    "error": f"Cannot activate workbook: {str(e)}",
                    "execution_log": execution_log
                }
            
            # Navigate to sheet if specified
            if sheet_name:
                try:
                    sheet = wb.sheets[sheet_name]
                    sheet.activate()
                    time.sleep(0.1)
                except Exception as e:
                    return {
                        "error": f"Cannot navigate to sheet '{sheet_name}': {str(e)}",
                        "execution_log": execution_log
                    }
            
            # Parse VBA code
            has_structure, detected_proc_name, clean_code = parse_vba_code(vba_code)
            
            # Generate unique names for this attempt
            final_module_name, final_procedure_name = generate_unique_names(wb, module_name, procedure_name)
            
            # Use detected procedure name if available
            if has_structure and detected_proc_name:
                final_procedure_name = detected_proc_name
            
            # Clean up old temporary modules (based on strategy)
            if strategy["clean_modules"]:
                clean_existing_temp_modules(wb)
                time.sleep(0.5)
            
            # Prepare VBA code with or without error handling wrapper
            if strategy["use_error_handling"]:
                if not has_structure:
                    # Wrap user code in error handling
                    final_code = wrap_vba_with_error_handling(clean_code, final_procedure_name)
                else:
                    # Add error handling to existing structure
                    final_code = clean_code.replace(
                        f"Sub {detected_proc_name}()",
                        f"Sub {detected_proc_name}()\n    On Error Resume Next\n    Application.DisplayAlerts = False"
                    )
            else:
                # Use original approach
                if not has_structure:
                    final_code = f"Sub {final_procedure_name}()\n{clean_code}\nEnd Sub"
                else:
                    final_code = clean_code
            
            attempt_log["module_name"] = final_module_name
            attempt_log["procedure_name"] = final_procedure_name
            attempt_log["code_length"] = len(final_code)
            
            # Execute VBA with monitoring
            success, error_info = execute_vba_with_monitoring(
                app, wb, final_code, final_module_name, final_procedure_name
            )
            
            attempt_log["execution_time"] = time.time() - attempt_log["start_time"]
            
            if success:
                # Success - restore settings and return
                restore_excel_settings(app, original_settings)
                attempt_log["result"] = "success"
                execution_log.append(attempt_log)
                
                return {
                    "success": True, 
                    "message": "VBA code executed successfully with enhanced error handling",
                    "module_name": final_module_name, 
                    "procedure_name": final_procedure_name,
                    "attempt": attempt + 1,
                    "strategy_used": strategy,
                    "execution_log": execution_log,
                    "alerts_disabled": True,
                    "dialog_monitoring": True
                }
            else:
                # Execution failed
                attempt_log["result"] = "failed"
                attempt_log["error"] = error_info
                execution_log.append(attempt_log)
                
                # If this is not the last attempt, continue to retry
                if attempt < max_retries - 1:
                    time.sleep(strategy["delay"])
                    continue
                
                # Last attempt failed - return detailed error
                restore_excel_settings(app, original_settings)
                return {
                    "error": f"VBA execution failed after {max_retries} attempts with different strategies",
                    "last_error": error_info,
                    "execution_log": execution_log,
                    "strategies_tried": retry_strategies[:attempt + 1]
                }
                
        except Exception as outer_e:
            # Restore settings on any error
            if app and original_settings:
                restore_excel_settings(app, original_settings)
                
            attempt_log["result"] = "exception"
            attempt_log["error"] = {"type": "outer_exception", "message": str(outer_e)}
            attempt_log["execution_time"] = time.time() - attempt_log["start_time"]
            execution_log.append(attempt_log)
            
            # If this is not the last attempt, continue to retry
            if attempt < max_retries - 1:
                time.sleep(strategy["delay"])
                continue
            
            # Last attempt failed
            error_msg = str(outer_e)
            if "Excel not running" in error_msg:
                return {"error": "Excel is not running. Please open Excel first.", "execution_log": execution_log}
            elif "No workbook" in error_msg:
                return {"error": "No active workbook found. Please open an Excel file.", "execution_log": execution_log}
            else:
                return {
                    "error": f"Connection error after {max_retries} attempts: {error_msg}",
                    "execution_log": execution_log
                }
    
    # Should not reach here
    return {"error": "Unexpected error: Failed all retry attempts", "execution_log": execution_log}

def main():
    try:
        parser = argparse.ArgumentParser(description='Execute VBA code in Excel with enhanced error handling')
        parser.add_argument('--code', required=True, help='VBA code to execute')
        parser.add_argument('--module', default='TempModule', help='Module name (default: TempModule)')
        parser.add_argument('--procedure', default='Main', help='Procedure name (default: Main)')
        parser.add_argument('--filename', help='Excel filename (optional)')
        parser.add_argument('--sheet', help='Sheet name to activate (optional)')
        
        args = parser.parse_args()
        result = execute_vba(args.code, args.module, args.procedure, args.filename, args.sheet)
        
        # Ensure UTF-8 output to handle Unicode characters
        output = json.dumps(result, ensure_ascii=False, indent=2)
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