# Excel MCP Server

*[日本語 README](README_ja.md)*

> Currently working on stabilizing the tool. If you encounter any strange behavior with any of the tools, please let me know via [Issues](https://github.com/kousunh/excel-mcp-server/issues).

MCP (Model Context Protocol) server for Excel operations. Performs various Excel operations including data reading, formatting, and VBA execution via xlwings.

## Requirements

- **Windows OS** (Required for win32com)
- **Microsoft Excel** installed
- **Node.js** 18 or higher
- **Python** 3.8 or higher

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
```

### 2. Install dependencies

```bash
npm install
pip install -r scripts/requirements.txt
```

### 3. Setup (Windows)

Run the setup script:

```bash
setup.bat
```

Or manually install Python dependencies:

```bash
pip install xlwings pandas openpyxl
```

## Usage

### 1. Start MCP Server

```bash
npm start
```

### 2. Configure Claude Desktop

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["C:/path/to/excel-mcp-server/src/index.js"],
      "env": {}
    }
  }
}
```

### 3. Example Usage

Use prompts like these in your AI client:

```
Please perform the following Excel operations:
- Enter "Hello World" in cell A1
- Enter current date/time in cell B1  
- Set yellow background color for range A1:B1
```

## Available Tools

### read_sheet_data (Recommended)
Reads Excel sheet data and returns it in a structured format. **Always use this tool first when checking or analyzing data.**

Parameters:
- `startRow` (optional): Starting row (default: 1)
- `endRow` (optional): Ending row (default: 100)
- `workbookName` (optional): Workbook name
- `sheetName` (optional): Sheet name

Features:
- Automatic Japanese text encoding correction
- Includes data statistics
- Automatic header row detection
- Automatic removal of empty rows/columns

### execute_vba
Executes VBA code.

Parameters:
- `vbaCode` (required): VBA code to execute
- `moduleName` (optional): Module name (default: TempModule)
- `procedureName` (optional): Procedure name (default: Main)

Features:
- Automatic retry functionality (up to 1 attempt)
- Automatic module name change on error

### get_cell_formats
Gets cell formatting information.

Parameters:
- `startRow`, `startCol`: Starting position
- `endRow`, `endCol`: Ending position (max 35 rows × 15 columns)
- `workbookName` (optional): Workbook name
- `sheetName` (optional): Sheet name

### edit_cells
Edits cell values.

Parameters:
- `range`: Cell range (e.g. "A1", "C9:C16")
- `value`: Value to set (arrays automatically detect vertical/horizontal direction)

### get_excel_status
Checks Excel status.

## Notes

- Excel must be running with at least one workbook open
- VBA code is added as temporary modules and automatically deleted after execution
- VBA execution may be restricted depending on security settings