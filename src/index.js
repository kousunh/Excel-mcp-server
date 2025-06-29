import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { 
  ListToolsRequestSchema, 
  CallToolRequestSchema 
} from '@modelcontextprotocol/sdk/types.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { ToolHandlers } from './handlers.js';
import { toolDefinitions } from './tools.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

class ExcelVBAServer {
  constructor() {
    this.server = new Server(
      {
        name: 'excel-mcp',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.handlers = new ToolHandlers(join(__dirname, '..', 'scripts'));
    this.setupToolHandlers();
    
    this.server.onerror = (error) => {
      // Log errors to stderr for debugging if needed
      // console.error('[MCP Error]', error);
    };
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: toolDefinitions
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      
      switch (name) {
        case 'get_open_workbooks':
          return await this.handlers.getOpenWorkbooks();
        case 'set_active_workbook':
          return await this.handlers.setActiveWorkbook(args);
        case 'fallback_execute_vba':
          return await this.handlers.executeVBA(args);
        case 'get_excel_status':
          return await this.handlers.getExcelStatus();
        case 'get_all_sheet_names':
          return await this.handlers.getAllSheetNames(args);
        case 'navigate_to_sheet':
          return await this.handlers.navigateToSheet(args);
        case 'edit_cells':
          return await this.handlers.editCells(args);
        case 'zz_final_verify_layout_formats':
          return await this.handlers.getCellFormats(args);
        case '01_first_analyze_excel_data':
          return await this.handlers.analyzeExcelData(args);
        case 'set_cell_borders':
          return await this.handlers.setCellBorders(args);
        case 'set_cell_formats':
          return await this.handlers.setCellFormats(args);
        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    // console.error('Excel VBA MCP server running on stdio');
  }
}

const server = new ExcelVBAServer();
server.run().catch(console.error);