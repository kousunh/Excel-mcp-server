# Excel MCP Server

*[日本語版 README はこちら](README_ja.md)*

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Excel through various operations including reading data, editing cells, executing VBA code, and managing worksheets.

## Prerequisites

- **Windows OS** (Required for win32com usage)
- **Microsoft Excel** installed
- **Node.js** 18 or higher
- **Python** 3.8 or higher

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
```

### 2. Run Setup Script

The setup script will create a Python virtual environment and install all dependencies.

**For Windows:**
```cmd
setup.bat
```

**For Linux/Mac (WSL):**
```bash
./setup.sh
```

The setup script will:
- Create a Python virtual environment (`venv`)
- Install Python dependencies (pywin32)
- Install Node.js dependencies
- Display configuration for Claude Desktop

### 3. Configure Claude Desktop

Add the server to Claude Desktop's configuration file.

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

Add to the `mcpServers` section:

```json
{
  "excel-mcp": {
    "command": "node",
    "args": [
      "C:\\path\\to\\excel-mcp-server\\excel-mcp.js"
    ],
    "env": {},
    "cwd": "C:\\path\\to\\excel-mcp-server"
  }
}
```

Replace `C:\\path\\to\\excel-mcp-server` with your actual clone path.

### 4. Configure Cursor IDE

If you're using Cursor IDE, you can add the MCP server to Cursor's configuration:

**Windows**: `%USERPROFILE%\.cursor\mcp.json`

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": [
        "C:\\path\\to\\excel-mcp-server\\excel-mcp.js"
      ],
      "env": {},
      "cwd": "C:\\path\\to\\excel-mcp-server"
    }
  }
}
```

Replace `C:\\path\\to\\excel-mcp-server` with your actual clone path.

## Manual Setup (Alternative)

If the setup script doesn't work, you can set up manually:

1. Create Python virtual environment:
   ```bash
   python -m venv venv
   ```

2. Activate the virtual environment:
   
   **Windows:**
   ```cmd
   venv\Scripts\activate
   ```
   
   **Linux/Mac:**
   ```bash
   source venv/bin/activate
   ```

3. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Install Node.js dependencies:
   ```bash
   npm install
   ```

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
2. **Python not found**: Ensure Python 3.8+ is installed and in PATH
3. **Import errors**: Re-run setup script or manually install dependencies
4. **VBA errors**: Check Excel's macro security settings
5. **Permission errors**: Some operations require "Trust access to the VBA project object model" to be enabled

## Development

To contribute or make modifications:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Security Notes

- The server only operates on local Excel files
- VBA code is executed in temporary modules that are deleted after execution
- Python virtual environment isolates dependencies

## License

MIT