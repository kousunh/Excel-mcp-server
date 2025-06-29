import { z } from 'zod';

export const schemas = {
  executeVBA: z.object({
    vbaCode: z.string().describe('The VBA code to execute'),
    workbookName: z.string().describe('Excel workbook name (required)'),
    moduleName: z.string().optional().describe('Optional module name (default: TempModule)'),
    procedureName: z.string().optional().describe('Optional procedure name (default: Main)'),
    sheetName: z.string().optional().describe('Optional sheet name to navigate to before executing VBA')
  }),
  
  setActiveWorkbook: z.object({
    workbookName: z.string().describe('Name of the workbook to activate')
  }),
  
  getAllSheetNames: z.object({
    workbookName: z.string().describe('Excel workbook name (required)')
  }),
  
  navigateToSheet: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    sheetName: z.string().describe('The name of the sheet to navigate to')
  }),
  
  editCells: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    range: z.string().describe('Cell or range to edit (e.g., "A1" or "A1:B5")'),
    value: z.union([
      z.string(),
      z.number(),
      z.array(z.union([z.string(), z.number()]))
    ]).describe('Value to set. Can be a single value or an array for multiple cells'),
    sheetName: z.string().optional().describe('Optional sheet name to navigate to before editing cells')
  }),
  
  getCellFormats: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    startRow: z.number().optional().describe('Starting row number (default: 1)'),
    startCol: z.number().optional().describe('Starting column number (default: 1)'),
    endRow: z.number().optional().describe('Ending row number (default: 20, max: 35 rows from start)'),
    endCol: z.number().optional().describe('Ending column number (default: 15, max: 15 columns from start)'),
    sheetName: z.string().optional().describe('Optional sheet name. If not specified, uses the active sheet')
  }),

  analyzeExcelData: z.object({
    filePath: z.string().optional().describe('Full path to Excel file (for closed files)'),
    workbookName: z.string().optional().describe('Name of open workbook (for open files)'),
    sheetName: z.string().optional().describe('Specific sheet name (optional, analyzes all sheets if not specified)'),
    startRow: z.number().optional().describe('Starting row number (default: 1)'),
    endRow: z.number().optional().describe('Ending row number (optional, analyzes all rows if not specified)'),
    mode: z.enum(['full', 'quick', 'data']).optional().describe('Analysis mode: full (detailed analysis), quick (basic info), data (data only)')
  })
};

export const toolDefinitions = [
  {
    name: 'get_open_workbooks',
    description: 'Get list of all open Excel workbooks',
    inputSchema: {
      type: 'object',
      properties: {},
      required: []
    },
  },
  {
    name: 'set_active_workbook',
    description: 'Set the active workbook in Excel',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Name of the workbook to activate'
        }
      },
      required: ['workbookName']
    },
  },
  {
    name: 'execute_vba',
    description: 'Execute VBA code with automatic retry (up to 2 attempts). Avoid MsgBox/Alert dialogs - use Debug.Print or cell output. If first attempt fails, automatically retries with different module names. If persistent failures occur, try changing procedure names manually.',
    inputSchema: {
      type: 'object',
      properties: {
        vbaCode: {
          type: 'string',
          description: 'The VBA code to execute'
        },
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        moduleName: {
          type: 'string',
          description: 'Optional module name (default: TempModule)'
        },
        procedureName: {
          type: 'string',
          description: 'Optional procedure name (default: Main)'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name to navigate to before executing VBA'
        }
      },
      required: ['vbaCode', 'workbookName']
    },
  },
  {
    name: 'get_excel_status',
    description: 'Check if Excel is running and has an active workbook',
    inputSchema: {
      type: 'object',
      properties: {},
      required: []
    },
  },
  {
    name: 'get_all_sheet_names',
    description: 'Get all sheet names in the specified workbook',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        }
      },
      required: ['workbookName']
    },
  },
  {
    name: 'navigate_to_sheet',
    description: 'Navigate to a specified sheet in the workbook',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        sheetName: {
          type: 'string',
          description: 'The name of the sheet to navigate to'
        }
      },
      required: ['workbookName', 'sheetName']
    },
  },
  {
    name: 'edit_cells',
    description: 'Edit one or multiple cells in Excel. Supports single cells (e.g., "A1") or ranges (e.g., "A1:B5")',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        range: {
          type: 'string',
          description: 'Cell or range to edit (e.g., "A1" or "A1:B5")'
        },
        value: {
          oneOf: [
            { type: 'string' },
            { type: 'number' },
            { type: 'array', items: { oneOf: [{ type: 'string' }, { type: 'number' }] } }
          ],
          description: 'Value to set. Can be a single value or an array for multiple cells'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name to navigate to before editing cells'
        }
      },
      required: ['workbookName', 'range', 'value']
    },
  },
  {
    name: 'get_cell_formats',
    description: 'Get detailed cell formatting (colors, borders, fonts) with Japanese text support. Now supports up to 35 rows x 15 columns. Use after read_sheet_data to understand visual formatting.',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        startRow: {
          type: 'number',
          description: 'Starting row number (default: 1)'
        },
        startCol: {
          type: 'number',
          description: 'Starting column number (default: 1)'
        },
        endRow: {
          type: 'number',
          description: 'Ending row number (default: 20, max: 35 rows from start)'
        },
        endCol: {
          type: 'number',
          description: 'Ending column number (default: 15, max: 15 columns from start)'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name. If not specified, uses the active sheet'
        }
      },
      required: ['workbookName']
    },
  },
  {
    name: 'analyze_excel_data',
    description: '**UNIVERSAL Excel Data Analyzer** - Works with both OPEN and CLOSED Excel files. Use pandas for closed files or xlwings for open workbooks. Provides comprehensive analysis including sheet structure, data types, statistics, and actual data. Choose analysis modes: "full" (detailed analysis), "quick" (basic info), or "data" (data extraction only). Perfect for understanding Excel structure and content regardless of file state.',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Full path to Excel file (for analyzing closed files). Use this when the Excel file is not currently open.'
        },
        workbookName: {
          type: 'string',
          description: 'Name of open workbook (for analyzing open files). Use this when the Excel file is already open in Excel.'
        },
        sheetName: {
          type: 'string',
          description: 'Specific sheet name to analyze. If not provided, analyzes all sheets and uses the first/active sheet for detailed data.'
        },
        startRow: {
          type: 'number',
          description: 'Starting row number for data extraction (default: 1)'
        },
        endRow: {
          type: 'number',
          description: 'Ending row number for data extraction (optional, analyzes all rows if not specified)'
        },
        mode: {
          type: 'string',
          enum: ['full', 'quick', 'data'],
          description: 'Analysis mode: "full" (detailed analysis with statistics), "quick" (basic sheet info), "data" (extract data only)'
        }
      },
      required: []
    },
  }
];