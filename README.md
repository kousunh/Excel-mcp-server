*[日本語 README](README_ja.md)*

A comprehensive Model Context Protocol (MCP) server for Microsoft Excel operations. Enables AI assistants to perform data analysis, cell editing, formatting, border styling, VBA execution, and complete worksheet management through a structured workflow.

## Key Features

### Prioritized Workflow Tools
- **Step 1**: `essential_inspect_excel_data` - MANDATORY first step for understanding data structure
- **Final Step**: `essential_check_excel_format` - MANDATORY final verification of layout and formatting

### Advanced Excel Operations
- **Data Analysis & Reading** - Comprehensive sheet analysis with statistics
- **Cell Editing** - Single cell or range editing with array support
- **Formatting Control** - Font colors, background colors, text styles, alignment
- **Border Management** - Complete border styling with multiple styles and colors
- **VBA Execution** - Simplified and stable VBA code execution
- **Workbook Management** - Multi-workbook handling and navigation

## Requirements

- **Windows OS** (Required for COM integration)
- **Microsoft Excel** installed and running
- **Node.js** 18 or higher
- **Python** 3.8 or higher

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
```

### 2. Run setup script

The setup script will create a Python virtual environment and install all dependencies.

**Windows (Command Prompt):**
```cmd
setup.bat
```

**Windows (PowerShell):**
```powershell
.\setup.bat
```

**Linux/Mac (WSL):**
```bash
./setup.sh
```

### 3. Configure Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["C:/path/to/excel-mcp-server/src/index.js"],
      "env": {},
      "cwd": "C:/path/to/excel-mcp-server"
    }
  }
}
```

### 4. Configure Cursor

Add to your Cursor settings:

```json
{
  "mcpServers": {
    "Excel-mcp": {
      "command": "node",
      "args": [
        "C:\\path\\to\\excel-mcp-server\\src\\index.js"
      ]
    }
  }
}
```

## Available Tools

### Priority Tools (Use in Order)

#### `essential_inspect_excel_data` - STEP 1 - ALWAYS USE FIRST
Analyzes Excel data structure and content before any operations. Essential for understanding current state, sheet structure, data types, and content. Works with both open and closed files.

#### `essential_check_excel_format` - FINAL STEP - MANDATORY VERIFICATION
Validates layout and formatting after any changes. Use multiple times to check different ranges. If layout/format issues found, fix and re-verify.

### Core Operations

#### `edit_cells`
Edit single cells or ranges with optimized performance for large data operations.

#### `set_cell_formats`
Apply comprehensive formatting to cell ranges including font colors, background colors, bold, italic, underline, font size, font name, and text alignment.

#### `set_cell_borders`
Apply detailed border styling to cell ranges. Supports various border styles (thin, thick, medium, double, dotted, dashed) and colors for different positions (top, bottom, left, right, inside, outside).

### Utility Tools

#### `get_open_workbooks`
Lists all currently open Excel workbooks.

#### `set_active_workbook`
Switches between open workbooks.

#### `get_all_sheet_names`
Lists all sheets in a workbook.

#### `navigate_to_sheet`
Switches to a specific sheet.

#### `get_excel_status`
Checks if Excel is running and responsive.

### VBA Execution

#### `execute_vba`
Executes custom VBA code in Excel. Creates a temporary Sub procedure, executes it, and automatically cleans up. Supports error handling and unique procedure naming to avoid conflicts.

## Usage Examples

### Basic Workflow
```
1. "First, analyze the current Excel data structure"
2. "Edit cells A1:C3 with employee data"
3. "Set borders around the data range"
4. "Apply formatting with bold headers and colored backgrounds"
5. "Finally, verify the layout and formatting"
```

### Data Operations
```
"Analyze the sales data in workbook 'Sales2024.xlsx'"
"Edit range A1:D10 with quarterly sales figures"
"Apply borders to create a table structure"
"Format headers with bold font and blue background"
"Verify the final layout looks correct"
```

## Workflow Best Practices

1. **Always start** with `essential_inspect_excel_data` to understand current state
2. **Use dedicated tools** (edit_cells, set_cell_formats, set_cell_borders) instead of VBA when possible
3. **Always end** with `essential_check_excel_format` to confirm changes
4. **Use execute_vba** for custom VBA logic when standard tools are insufficient
5. **Verify multiple ranges** if working with large spreadsheets

## Troubleshooting

1. **Excel not found**: Ensure Excel is running with at least one workbook open
2. **Tool timeouts**: Large operations automatically use extended timeouts (60 seconds)
3. **VBA errors**: Simplified VBA execution reduces hanging and freezing issues
4. **Format verification**: Use verification tool multiple times for different ranges
5. **Permission errors**: Enable "Trust access to the VBA project object model" in Excel Trust Center

## Security

- Server operates only on local Excel files
- VBA code runs in temporary modules that are automatically deleted
- Python virtual environment isolates dependencies
- No network access or external file operations

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly with Excel operations
5. Submit a pull request

---

**Version 2.0 Changes:**
- Added prioritized workflow tools with clear naming
- Implemented comprehensive formatting and border controls
- Optimized performance for large data operations
- Simplified and stabilized VBA execution
- Added mandatory verification step for quality assurance