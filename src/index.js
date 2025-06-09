import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { 
  ListToolsRequestSchema, 
  CallToolRequestSchema 
} from '@modelcontextprotocol/sdk/types.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { toolDefinitions } from './tools.js';
import { ToolHandlers } from './handlers.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Helper function to get the scripts directory path
const getScriptsPath = () => {
  return process.env.EXCEL_MCP_BUNDLED ? join(__dirname, 'scripts') : join(__dirname, '..', 'scripts');
};

class ExcelMCPServer {
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

    this.handlers = new ToolHandlers(getScriptsPath());
    this.setupToolHandlers();
    
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: toolDefinitions,
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      
      const handlerMap = {
        'get_open_workbooks': () => this.handlers.getOpenWorkbooks(),
        'set_active_workbook': () => this.handlers.setActiveWorkbook(args),
        'execute_vba': () => this.handlers.executeVBA(args),
        'get_excel_status': () => this.handlers.getExcelStatus(),
        'get_all_sheet_names': () => this.handlers.getAllSheetNames(args),
        'navigate_to_sheet': () => this.handlers.navigateToSheet(args),
        'edit_cells': () => this.handlers.editCells(args),
        'read_sheet_data': () => this.handlers.readSheetData(args),
        'get_cell_formats': () => this.handlers.getCellFormats(args),
      };

      const handler = handlerMap[name];
      if (!handler) {
        throw new Error(`Unknown tool: ${name}`);
      }
      
      return await handler();
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Excel MCP server running on stdio');
  }
}

const server = new ExcelMCPServer();
server.run().catch(console.error);