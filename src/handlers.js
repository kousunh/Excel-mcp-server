import { spawn } from 'child_process';
import { join } from 'path';
import { schemas } from './tools.js';

export class ToolHandlers {
  constructor(scriptsPath) {
    this.scriptsPath = scriptsPath;
  }

  executePython(scriptName, args = []) {
    return new Promise((resolve) => {
      const scriptPath = join(this.scriptsPath, scriptName);
      const pythonCmd = process.env.EXCEL_MCP_PYTHON || 'python';
      const python = spawn(pythonCmd, [scriptPath, ...args], {
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
      });
      let output = '';
      let error = '';
      
      python.stdout.setEncoding('utf8');
      python.stderr.setEncoding('utf8');
      
      python.stdout.on('data', (data) => output += data);
      python.stderr.on('data', (data) => error += data);
      
      python.on('close', (code) => {
        const result = output.trim() || (code !== 0 ? `{"error": "${error || 'Script failed'}"}` : '{"error": "No output"}');
        resolve({
          content: [{ type: 'text', text: result }]
        });
      });
      
      python.on('error', (err) => {
        resolve({
          content: [{ type: 'text', text: `{"error": "Python process error: ${err.message}"}` }]
        });
      });
    });
  }

  async getOpenWorkbooks() {
    return this.executePython('get_open_workbooks.py');
  }

  async setActiveWorkbook(args) {
    const validated = schemas.setActiveWorkbook.parse(args);
    return this.executePython('set_active_workbook.py', ['--workbook', validated.workbookName]);
  }

  async executeVBA(args) {
    const validated = schemas.executeVBA.parse(args);
    const { vbaCode, workbookName, moduleName = 'TempModule', procedureName = 'Main', sheetName } = validated;
    
    const pythonArgs = [
      '--code', vbaCode,
      '--module', moduleName,
      '--procedure', procedureName,
      '--filename', workbookName
    ];
    
    if (sheetName) pythonArgs.push('--sheet', sheetName);
    
    return this.executePython('execute_vba.py', pythonArgs);
  }

  async getExcelStatus() {
    return this.executePython('check_excel.py');
  }

  async getAllSheetNames(args) {
    const validated = schemas.getAllSheetNames.parse(args);
    return this.executePython('get_sheet_names.py', ['--filename', validated.workbookName]);
  }

  async navigateToSheet(args) {
    const validated = schemas.navigateToSheet.parse(args);
    const pythonArgs = ['--sheet', validated.sheetName, '--filename', validated.workbookName];
    return this.executePython('navigate_to_sheet.py', pythonArgs);
  }

  async editCells(args) {
    const validated = schemas.editCells.parse(args);
    const pythonArgs = ['--range', validated.range];
    const valueStr = typeof validated.value === 'object' ? JSON.stringify(validated.value) : String(validated.value);
    pythonArgs.push('--value', valueStr);
    pythonArgs.push('--filename', validated.workbookName);
    if (validated.sheetName) pythonArgs.push('--sheet', validated.sheetName);
    
    return this.executePython('edit_cells.py', pythonArgs);
  }

  async getCellFormats(args) {
    const validated = schemas.getCellFormats.parse(args);
    const pythonArgs = [];
    if (validated.startRow) pythonArgs.push('--start-row', String(validated.startRow));
    if (validated.startCol) pythonArgs.push('--start-col', String(validated.startCol));
    if (validated.endRow) pythonArgs.push('--end-row', String(validated.endRow));
    if (validated.endCol) pythonArgs.push('--end-col', String(validated.endCol));
    pythonArgs.push('--filename', validated.workbookName);
    if (validated.sheetName) pythonArgs.push('--sheet', validated.sheetName);
    
    return this.executePython('get_cell_formats.py', pythonArgs);
  }

  async analyzeExcelData(args) {
    const validated = schemas.analyzeExcelData.parse(args);
    const pythonArgs = [];
    
    // Add file path or workbook name
    if (validated.filePath) {
      pythonArgs.push('--file', validated.filePath);
    } else if (validated.workbookName) {
      pythonArgs.push('--workbook', validated.workbookName);
    }
    
    // Add optional parameters
    if (validated.sheetName) pythonArgs.push('--sheet', validated.sheetName);
    if (validated.startRow) pythonArgs.push('--start-row', String(validated.startRow));
    if (validated.endRow) pythonArgs.push('--end-row', String(validated.endRow));
    if (validated.mode) pythonArgs.push('--mode', validated.mode);
    
    return this.executePython('analyze_excel_data.py', pythonArgs);
  }
}