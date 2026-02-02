import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { ToolHandlers } from './handlers.js';
import { toolDefinitions } from './tools.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const server = new Server(
  { name: 'excel-mcp', version: '3.0.0' },
  { capabilities: { tools: {} } }
);

const handlers = new ToolHandlers(join(__dirname, '..', 'scripts'));

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: toolDefinitions
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  switch (name) {
    case 'get_excel_info':  return handlers.getExcelInfo();
    case 'read_cells':      return handlers.readCells(args);
    case 'write_cells':     return handlers.writeCells(args);
    case 'format_cells':    return handlers.formatCells(args);
    case 'execute_vba':     return handlers.executeVba(args);
    default: throw new Error(`Unknown tool: ${name}`);
  }
});

server.onerror = () => {};
process.on('SIGINT', async () => { await server.close(); process.exit(0); });

const transport = new StdioServerTransport();
server.connect(transport).catch(console.error);
