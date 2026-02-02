import { spawn } from 'child_process';
import { join } from 'path';
import { schemas } from './tools.js';

export class ToolHandlers {
  constructor(scriptsPath) {
    this.scriptsPath = scriptsPath;
  }

  _run(scriptName, args = [], timeout = 30000) {
    return new Promise((resolve) => {
      const scriptPath = join(this.scriptsPath, scriptName);
      const pythonCmd = process.env.EXCEL_MCP_PYTHON || 'python';
      const python = spawn(pythonCmd, [scriptPath, ...args], {
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
      });

      let output = '';
      let error = '';
      let done = false;

      const timer = setTimeout(() => {
        if (!done) {
          done = true;
          python.kill('SIGTERM');
          resolve({ content: [{ type: 'text', text: '{"error":"Timeout"}' }] });
        }
      }, timeout);

      python.stdout.setEncoding('utf8');
      python.stderr.setEncoding('utf8');
      python.stdout.on('data', (d) => output += d);
      python.stderr.on('data', (d) => error += d);

      python.on('close', (code) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          const text = output.trim() || (code !== 0 ? `{"error":"${error || 'Script failed'}"}` : '{"error":"No output"}');
          resolve({ content: [{ type: 'text', text }] });
        }
      });

      python.on('error', (err) => {
        if (!done) {
          done = true;
          clearTimeout(timer);
          resolve({ content: [{ type: 'text', text: `{"error":"${err.message}"}` }] });
        }
      });
    });
  }

  async getExcelInfo() {
    return this._run('excel_info.py');
  }

  async readCells(args) {
    const v = schemas.readCells.parse(args);
    const a = ['--workbook', v.workbook, '--range', v.range];
    if (v.sheet) a.push('--sheet', v.sheet);
    if (v.formats) a.push('--formats');
    return this._run('read_cells.py', a);
  }

  async writeCells(args) {
    const v = schemas.writeCells.parse(args);
    const valueStr = typeof v.value === 'object' ? JSON.stringify(v.value) : String(v.value);
    const a = ['--workbook', v.workbook, '--range', v.range, '--value', valueStr];
    if (v.sheet) a.push('--sheet', v.sheet);
    return this._run('write_cells.py', a, 60000);
  }

  async formatCells(args) {
    const v = schemas.formatCells.parse(args);
    const a = ['--workbook', v.workbook, '--range', v.range, '--format', JSON.stringify(v.format)];
    if (v.sheet) a.push('--sheet', v.sheet);
    return this._run('format_cells.py', a);
  }

  async executeVba(args) {
    const v = schemas.executeVba.parse(args);
    const a = ['--workbook', v.workbook, '--code', v.code];
    if (v.sheet) a.push('--sheet', v.sheet);
    return this._run('execute_vba.py', a, 60000);
  }
}
