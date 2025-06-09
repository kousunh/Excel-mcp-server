# Excel MCP Server

*[日本語版 README はこちら](README_ja.md)*

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Excel through various operations including reading data, editing cells, executing VBA code, and managing worksheets.

## Prerequisites

- **Node.js** 18 or higher
- **Python** 3.x
- **xlwings** Python package
- **Microsoft Excel** (Windows or macOS)

*Note: This tool is optimized for Windows environments where Excel is most commonly used.*

## Installation

### Option 1: Quick Install with npx (Recommended)

For Windows PowerShell/Command Prompt:
```cmd
npx github:kousunh/excel-mcp-server
```

For Windows PowerShell (alternative):
```powershell
npx github:kousunh/excel-mcp-server
```

This will automatically:
- Create a Python virtual environment
- Install all required Python packages (xlwings, pandas, numpy)
- Start the MCP server

### Option 2: From Source

1. Clone the repository
2. Install dependencies:
   
   For Windows PowerShell/Command Prompt:
   ```cmd
   npm install
   npm run setup
   ```
   
   The `npm run setup` command creates a virtual environment and installs Python packages.

### Option 3: Manual Installation

1. Download the repository
2. Install Python dependencies:
   
   For Windows Command Prompt:
   ```cmd
   pip install xlwings pandas numpy
   ```
   
   For Windows PowerShell:
   ```powershell
   pip install xlwings pandas numpy
   ```

## Configuration

### Claude Code Setup

#### Option 1: Using Claude Code MCP Add Command (Recommended)

For Claude Code (Windows WSL/Linux/macOS):
```bash
claude mcp add excel-mcp -- npx -y github:kousunh/excel-mcp-server
```

#### Option 2: Using Claude Code MCP Add-JSON Command

For Claude Code (Windows WSL/Linux/macOS):
```bash
claude mcp add-json excel-mcp '{
  "command": "npx",
  "args": [
    "-y",
    "github:kousunh/excel-mcp-server"
  ]
}'
```

#### Option 3: Manual Configuration (.mcp.json)

Create or edit `.mcp.json` file in your project root:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "npx",
      "args": [
        "-y",
        "github:kousunh/excel-mcp-server"
      ],
      "env": {}
    }
  }
}
```

### Cursor Integration

If you're using Cursor IDE, you can also configure the MCP server through Cursor's settings:

1. Open Cursor settings
2. Navigate to MCP configuration
3. Add the Excel MCP server with the configuration above

## Usage

Start Excel and open a workbook before using. The server provides the following tools:

### Core Tools

- **read_sheet_data** - Read worksheet data with full column preservation
- **edit_cells** - Edit single cells or ranges
- **execute_vba** - Run VBA code with automatic retry
- **get_cell_formats** - Get cell formatting details (colors, fonts, borders)

### Workbook Management

- **get_open_workbooks** - List all open Excel workbooks
- **set_active_workbook** - Switch between open workbooks
- **get_excel_status** - Check if Excel is running

### Sheet Navigation

- **get_all_sheet_names** - List all sheets in a workbook
- **navigate_to_sheet** - Switch to a specific sheet

## Examples

### Basic Operations

```
"Read data from the active Excel sheet"
"Write 'Hello World' to cell A1"
"Get all sheet names in the current workbook"
```

### Data Processing

```
"Read data from rows 1-50 and analyze the trends"
"Copy values from column A to column B with formatting"
"Fill cells A1:A10 with sequential numbers"
```

### VBA Automation

```
"Create a VBA macro to sort column A"
"Run VBA code to apply conditional formatting"
"Execute a macro to generate a summary report"
```

## Troubleshooting

1. **Excel not found**: Ensure Excel is running with at least one workbook open
2. **VBA errors**: Check Excel's macro security settings
3. **Permission errors**: Some operations require "Trust access to the VBA project object model" to be enabled in Excel
## Security Notes

- The server only operates on local Excel files
- VBA code is executed in temporary modules that are deleted after execution
## License

MIT