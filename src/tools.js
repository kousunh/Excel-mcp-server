import { z } from 'zod';

export const schemas = {
  fallbackExecuteVBA: z.object({
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
  
  finalVerifyLayoutFormats: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    startRow: z.number().optional().describe('Starting row number (default: 1)'),
    startCol: z.number().optional().describe('Starting column number (default: 1)'),
    endRow: z.number().optional().describe('Ending row number (default: 20, max: 35 rows from start)'),
    endCol: z.number().optional().describe('Ending column number (default: 15, max: 15 columns from start)'),
    sheetName: z.string().optional().describe('Optional sheet name. If not specified, uses the active sheet')
  }),

  firstAnalyzeExcelData: z.object({
    filePath: z.string().optional().describe('Full path to Excel file (for closed files)'),
    workbookName: z.string().optional().describe('Name of open workbook (for open files)'),
    sheetName: z.string().optional().describe('Specific sheet name (optional, analyzes all sheets if not specified)'),
    startRow: z.number().optional().describe('Starting row number (default: 1)'),
    endRow: z.number().optional().describe('Ending row number (optional, analyzes all rows if not specified)'),
    mode: z.enum(['full', 'quick', 'data']).optional().describe('Analysis mode: full (detailed analysis), quick (basic info), data (data only)')
  }),

  setCellBorders: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    range: z.string().describe('Cell range (e.g., "A1:C3")'),
    borders: z.object({
      top: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional(),
      bottom: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional(),
      left: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional(),
      right: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional(),
      inside: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional(),
      outside: z.object({
        style: z.enum(['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none']).optional(),
        color: z.string().optional().describe('Border color in hex format (e.g., "#000000")')
      }).optional()
    }),
    sheetName: z.string().optional().describe('Optional sheet name')
  }),

  setCellFormats: z.object({
    workbookName: z.string().describe('Excel workbook name (required)'),
    range: z.string().describe('Cell range (e.g., "A1:C3")'),
    format: z.object({
      fontColor: z.string().optional().describe('Font color in hex format (e.g., "#000000")'),
      backgroundColor: z.string().optional().describe('Background color in hex format (e.g., "#FFFFFF")'),
      bold: z.boolean().optional().describe('Bold formatting'),
      italic: z.boolean().optional().describe('Italic formatting'),
      underline: z.boolean().optional().describe('Underline formatting'),
      fontSize: z.number().optional().describe('Font size'),
      fontName: z.string().optional().describe('Font name (e.g., "Arial", "Times New Roman")'),
      textAlign: z.enum(['left', 'center', 'right']).optional().describe('Text alignment'),
      verticalAlign: z.enum(['top', 'middle', 'bottom']).optional().describe('Vertical alignment')
    }),
    sheetName: z.string().optional().describe('Optional sheet name')
  })
};

export const toolDefinitions = [
  {
    name: 'essential_inspect_excel_data',
    description: '🔍 STEP 1 - ALWAYS USE FIRST! Analyze Excel data structure and content before any operations. Essential for understanding current state, sheet structure, data types, and content. Works with both open and closed files. MANDATORY first step before editing, formatting, or other operations. Analysis modes: "full" (detailed analysis), "quick" (basic info), or "data" (content only).',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Full path to Excel file (for closed files). Use when Excel file is not currently open in Excel.'
        },
        workbookName: {
          type: 'string',
          description: 'Name of open workbook (for open files). Use when Excel file is already open in Excel.'
        },
        sheetName: {
          type: 'string',
          description: 'Specific sheet name to analyze. If not provided, analyzes all sheets in the workbook.'
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
          description: 'Analysis mode: "full" (detailed analysis with statistics), "quick" (basic sheet info), "data" (content extraction only)'
        }
      },
      required: []
    },
  },
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
    description: 'Execute custom VBA code in Excel. Creates a temporary Sub procedure, executes it, and automatically cleans up. Supports error handling and unique procedure naming to avoid conflicts. Use for operations that require custom VBA logic beyond the standard Excel tools.',
    inputSchema: {
      type: 'object',
      properties: {
        vbaCode: {
          type: 'string',
          description: 'The VBA code to execute. Will be wrapped in a Sub procedure automatically if not already structured.'
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
    name: 'essential_check_excel_format',
    description: '✅ FINAL STEP - MANDATORY layout and format verification! Validates visual appearance: cell formatting, colors, borders, fonts, and overall layout. ALWAYS use as LAST STEP after any editing, formatting, or border changes. Use multiple times to check different ranges. If layout/format issues found, fix and re-verify. Supports up to 35 rows x 15 columns per check.',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        startRow: {
          type: 'number',
          description: 'Starting row number for validation (default: 1)'
        },
        startCol: {
          type: 'number',
          description: 'Starting column number for validation (default: 1)'
        },
        endRow: {
          type: 'number',
          description: 'Ending row number for validation (default: 20, max: 35 rows from start)'
        },
        endCol: {
          type: 'number',
          description: 'Ending column number for validation (default: 15, max: 15 columns from start)'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name. If not specified, validates the active sheet'
        }
      },
      required: ['workbookName']
    },
  },
  {
    name: 'set_cell_borders',
    description: 'Set borders for multiple cells in a range. Supports various border styles and colors for different border positions (top, bottom, left, right, inside, outside).',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        range: {
          type: 'string',
          description: 'Cell range (e.g., "A1:C3")'
        },
        borders: {
          type: 'object',
          properties: {
            top: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            },
            bottom: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            },
            left: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            },
            right: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            },
            inside: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            },
            outside: {
              type: 'object',
              properties: {
                style: {
                  type: 'string',
                  enum: ['thin', 'thick', 'medium', 'double', 'dotted', 'dashed', 'none'],
                  description: 'Border style'
                },
                color: {
                  type: 'string',
                  description: 'Border color in hex format (e.g., "#000000")'
                }
              }
            }
          },
          description: 'Border settings for different positions'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name'
        }
      },
      required: ['workbookName', 'range', 'borders']
    },
  },
  {
    name: 'set_cell_formats',
    description: 'Set formatting for multiple cells including font color, background color, bold, italic, underline, font size, font name, and text alignment.',
    inputSchema: {
      type: 'object',
      properties: {
        workbookName: {
          type: 'string',
          description: 'Excel workbook name (required)'
        },
        range: {
          type: 'string',
          description: 'Cell range (e.g., "A1:C3")'
        },
        format: {
          type: 'object',
          properties: {
            fontColor: {
              type: 'string',
              description: 'Font color in hex format (e.g., "#000000")'
            },
            backgroundColor: {
              type: 'string',
              description: 'Background color in hex format (e.g., "#FFFFFF")'
            },
            bold: {
              type: 'boolean',
              description: 'Bold formatting'
            },
            italic: {
              type: 'boolean',
              description: 'Italic formatting'
            },
            underline: {
              type: 'boolean',
              description: 'Underline formatting'
            },
            fontSize: {
              type: 'number',
              description: 'Font size'
            },
            fontName: {
              type: 'string',
              description: 'Font name (e.g., "Arial", "Times New Roman")'
            },
            textAlign: {
              type: 'string',
              enum: ['left', 'center', 'right'],
              description: 'Text alignment'
            },
            verticalAlign: {
              type: 'string',
              enum: ['top', 'middle', 'bottom'],
              description: 'Vertical alignment'
            }
          },
          description: 'Formatting options'
        },
        sheetName: {
          type: 'string',
          description: 'Optional sheet name'
        }
      },
      required: ['workbookName', 'range', 'format']
    },
  }
];