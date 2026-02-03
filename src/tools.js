import { z } from 'zod';

export const schemas = {
  readCells: z.object({
    workbook: z.string().optional(),
    path: z.string().optional(),
    range: z.string(),
    sheet: z.string().optional(),
    formats: z.boolean().optional()
  }),
  writeCells: z.object({
    workbook: z.string().optional(),
    path: z.string().optional(),
    range: z.string(),
    value: z.union([z.string(), z.number(), z.boolean(), z.array(z.any())]),
    sheet: z.string().optional()
  }),
  formatCells: z.object({
    workbook: z.string().optional(),
    path: z.string().optional(),
    range: z.string(),
    format: z.record(z.any()),
    sheet: z.string().optional()
  }),
  executeVba: z.object({
    workbook: z.string(),
    code: z.string(),
    sheet: z.string().optional()
  })
};

export const toolDefinitions = [
  {
    name: 'get_excel_info',
    description: 'Get Excel running status, open workbooks, and their sheets.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: []
    }
  },
  {
    name: 'read_cells',
    description: 'Read cell values from a range. Use "workbook" for an open Excel workbook, or "path" for a .xlsx file on disk (Excel not required). Set formats=true to include formatting details.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: { type: 'string', description: 'Open workbook name (live Excel)' },
        path: { type: 'string', description: 'File path to .xlsx (read-only, Excel not required)' },
        range: { type: 'string', description: 'Cell range (e.g. "A1" or "A1:C10")' },
        sheet: { type: 'string', description: 'Sheet name (default: active sheet)' },
        formats: { type: 'boolean', description: 'Include cell formatting (default: false)' }
      },
      required: ['range']
    }
  },
  {
    name: 'write_cells',
    description: 'Write values to a cell or range. Use "workbook" for live Excel, or "path" for a file on disk (opens in Excel). Accepts a single value, a flat array, or a 2D array.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: { type: 'string', description: 'Open workbook name (live Excel)' },
        path: { type: 'string', description: 'File path to .xlsx (opens in Excel)' },
        range: { type: 'string', description: 'Cell range (e.g. "A1" or "A1:B5")' },
        value: {
          oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'array' }],
          description: 'Value(s) to write'
        },
        sheet: { type: 'string', description: 'Sheet name (default: active sheet)' }
      },
      required: ['range', 'value']
    }
  },
  {
    name: 'format_cells',
    description: 'Apply formatting to cells. Use "workbook" for live Excel, or "path" for a file on disk (opens in Excel). Options: bold, italic, underline, fontSize, fontName, fontColor, backgroundColor, textAlign (left/center/right), verticalAlign (top/middle/bottom), numberFormat, wrapText, borders ({top/bottom/left/right/inside/outside: {style, color}}).',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: { type: 'string', description: 'Open workbook name (live Excel)' },
        path: { type: 'string', description: 'File path to .xlsx (opens in Excel)' },
        range: { type: 'string', description: 'Cell range (e.g. "A1:C3")' },
        format: {
          type: 'object',
          description: 'Formatting options',
          properties: {
            bold: { type: 'boolean' },
            italic: { type: 'boolean' },
            underline: { type: 'boolean' },
            fontSize: { type: 'number' },
            fontName: { type: 'string' },
            fontColor: { type: 'string', description: 'Hex color (e.g. "#FF0000")' },
            backgroundColor: { type: 'string', description: 'Hex color' },
            textAlign: { type: 'string', enum: ['left', 'center', 'right'] },
            verticalAlign: { type: 'string', enum: ['top', 'middle', 'bottom'] },
            numberFormat: { type: 'string' },
            wrapText: { type: 'boolean' },
            borders: {
              type: 'object',
              description: 'Border settings per position',
              properties: {
                top: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } },
                bottom: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } },
                left: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } },
                right: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } },
                inside: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } },
                outside: { type: 'object', properties: { style: { type: 'string' }, color: { type: 'string' } } }
              }
            }
          }
        },
        sheet: { type: 'string', description: 'Sheet name (default: active sheet)' }
      },
      required: ['range', 'format']
    }
  },
  {
    name: 'execute_vba',
    description: 'Execute VBA code in an open workbook. Code is wrapped in a Sub automatically if needed. MsgBox calls are stripped. Temp modules are cleaned up after execution.',
    inputSchema: {
      type: 'object',
      properties: {
        workbook: { type: 'string', description: 'Workbook name' },
        code: { type: 'string', description: 'VBA code to execute' },
        sheet: { type: 'string', description: 'Sheet to activate before execution' }
      },
      required: ['workbook', 'code']
    }
  }
];
