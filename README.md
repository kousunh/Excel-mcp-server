*[日本語 README](README_ja.md)*

# Excel MCP Server

A cross-platform MCP server for reading, writing, and formatting Excel files. Works with both open Excel workbooks (live) and closed .xlsx files (no Excel needed).

## Two Modes

- **`workbook`** — Operate on an open Excel workbook in real-time via xlwings
- **`path`** — Edit closed .xlsx files directly using pure Python (preserves images, charts, and all embedded content)

## Tools

| Tool | workbook | path | Required |
|------|:--------:|:----:|----------|
| `get_excel_info` | - | - | (none) |
| `read_cells` | OK | OK | range |
| `write_cells` | OK | OK | range, value |
| `format_cells` | OK | OK | range, format |
| `execute_vba` | OK | - | workbook, code |

## Requirements

- **Node.js** 18+
- **Python** 3.8+
- **xlwings** (for live Excel mode — `pip install xlwings`)
- **Microsoft Excel** (only needed for `workbook` mode and `execute_vba`)

Works on **Windows** and **macOS**. The `path` mode also works on Linux.

## Installation

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
npm install
pip install -r scripts/requirements.txt
```

### Configure MCP Client

Add to your MCP client config (Claude Desktop, Cursor, etc.):

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/path/to/excel-mcp-server/src/index.js"]
    }
  }
}
```

To use a specific Python executable, set the `EXCEL_MCP_PYTHON` environment variable:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/path/to/excel-mcp-server/src/index.js"],
      "env": {
        "EXCEL_MCP_PYTHON": "/path/to/python"
      }
    }
  }
}
```

## Usage

### Closed files (path mode)

```
read_cells   path="/data/report.xlsx" range="A1:D20" formats=true
write_cells  path="/data/report.xlsx" range="A1:C3" value=[["Name","Age","City"],["Alice",30,"NYC"],["Bob",25,"LA"]]
format_cells path="/data/report.xlsx" range="A1:C1" format={"bold":true,"backgroundColor":"#4472C4","fontColor":"#FFFFFF"}
```

No Excel installation required. Images, charts, and shapes are preserved.

### Open workbooks (workbook mode)

```
get_excel_info
read_cells   workbook="Sales.xlsx" range="B2:F10"
write_cells  workbook="Sales.xlsx" range="G2" value="=SUM(B2:F2)"
execute_vba  workbook="Sales.xlsx" code="Range(\"A1:G10\").AutoFilter"
```

Requires Excel to be running with the workbook open.

## License

MIT
