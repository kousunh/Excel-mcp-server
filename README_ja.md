*[English README](README.md)*

# Excel MCP Server

Excel ファイルの読み書き・書式設定を行うクロスプラットフォーム MCP サーバー。開いている Excel ブック（ライブ）と、閉じた .xlsx ファイル（Excel 不要）の両方に対応。

## 2つのモード

- **`workbook`** — 開いている Excel ブックを xlwings でリアルタイム操作
- **`path`** — 閉じた .xlsx ファイルを純 Python で直接編集（画像・グラフ・図形をすべて保持）

## ツール

| ツール | workbook | path | 必須パラメータ |
|--------|:--------:|:----:|---------------|
| `get_excel_info` | - | - | なし |
| `read_cells` | OK | OK | range |
| `write_cells` | OK | OK | range, value |
| `format_cells` | OK | OK | range, format |
| `execute_vba` | OK | - | workbook, code |

## 必要な環境

- **Node.js** 18+
- **Python** 3.8+
- **xlwings**（ライブ Excel モード用 — `pip install xlwings`）
- **Microsoft Excel**（`workbook` モードと `execute_vba` のみ必要）

**Windows** と **macOS** で動作。`path` モードは Linux でも動作。

## インストール

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
npm install
pip install -r scripts/requirements.txt
```

### MCP クライアントの設定

MCP クライアント（Claude Desktop、Cursor 等）の設定に追加:

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

特定の Python を使う場合は `EXCEL_MCP_PYTHON` 環境変数を設定:

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

## 使用例

### 閉じたファイル（path モード）

```
read_cells   path="/data/report.xlsx" range="A1:D20" formats=true
write_cells  path="/data/report.xlsx" range="A1:C3" value=[["名前","年齢","都市"],["太郎",30,"東京"],["花子",25,"大阪"]]
format_cells path="/data/report.xlsx" range="A1:C1" format={"bold":true,"backgroundColor":"#4472C4","fontColor":"#FFFFFF"}
```

Excel のインストール不要。画像・グラフ・図形はそのまま保持。

### 開いているブック（workbook モード）

```
get_excel_info
read_cells   workbook="Sales.xlsx" range="B2:F10"
write_cells  workbook="Sales.xlsx" range="G2" value="=SUM(B2:F2)"
execute_vba  workbook="Sales.xlsx" code="Range(\"A1:G10\").AutoFilter"
```

Excel が起動中でブックが開いている必要あり。

## ライセンス

MIT
