# Excel MCPサーバー

*[English README](README.md)*

Microsoft Excelとの連携を可能にするModel Context Protocol (MCP)サーバーです。データ読取、セル編集、VBAコード実行、ワークシート管理など、様々な操作をAIアシスタントから実行できます。

## 必要な環境

- **Node.js** 18以上
- **Python** 3.x
- **xlwings** Pythonパッケージ
- **Microsoft Excel** (WindowsまたはmacOS)

## インストール

### 方法1: npxで簡単インストール（推奨）

Windows PowerShell / コマンドプロンプトの場合:
```cmd
npx github:kousunh/excel-mcp-server
```

Windows PowerShell（代替方法）:
```powershell
npx github:kousunh/excel-mcp-server
```

これにより自動的に:
- Python仮想環境を作成
- 必要なPythonパッケージ（xlwings、pandas、numpy）をインストール
- MCPサーバーを起動

### 方法2: ソースコードから

1. リポジトリをクローン
2. 依存関係をインストール:
   
   Windows PowerShell / コマンドプロンプトの場合:
   ```cmd
   npm install
   npm run setup
   ```
   
   `npm run setup`コマンドは仮想環境を作成しPythonパッケージをインストールします。

### 方法3: 手動インストール

1. リポジトリをダウンロード
2. Python依存関係をインストール:
   
   Windows コマンドプロンプトの場合:
   ```cmd
   pip install xlwings pandas numpy
   ```
   
   Windows PowerShellの場合:
   ```powershell
   pip install xlwings pandas numpy
   ```

## 設定

### Claude Codeの設定

#### 方法1: Claude Code MCP Addコマンドを使用（推奨）

Claude Code (Windows WSL/Linux/macOS)の場合:
```bash
claude mcp add excel-mcp -- npx -y github:kousunh/excel-mcp-server
```

#### 方法2: Claude Code MCP Add-JSONコマンドを使用

Claude Code (Windows WSL/Linux/macOS)の場合:
```bash
claude mcp add-json excel-mcp '{
  "command": "npx",
  "args": [
    "-y",
    "github:kousunh/excel-mcp-server"
  ]
}'
```

#### 方法3: 手動設定（.mcp.json）

プロジェクトルートに`.mcp.json`ファイルを作成または編集:

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

### Cursor IDEとの連携

Cursor IDEを使用している場合は、Cursorの設定からMCPサーバーを設定することもできます：

1. Cursorの設定を開く
2. MCP設定に移動
3. 上記の設定でExcel MCPサーバーを追加

## 使用方法

使用前にExcelを起動し、ワークブックを開いてください。以下のツールが利用可能です:

### 基本ツール

- **read_sheet_data** - ワークシートデータの読み取り（全列保持）
- **edit_cells** - 単一セルまたは範囲の編集
- **execute_vba** - VBAコードの実行（自動リトライ機能付き）
- **get_cell_formats** - セル書式の詳細取得（色、フォント、罫線）

### ワークブック管理

- **get_open_workbooks** - 開いているExcelワークブックの一覧表示
- **set_active_workbook** - ワークブックの切り替え
- **get_excel_status** - Excelの実行状態確認

### シート操作

- **get_all_sheet_names** - ワークブック内の全シート名取得
- **navigate_to_sheet** - 特定のシートへの移動

## 使用例

### 基本操作

```
「アクティブなExcelシートからデータを読み取って」
「A1セルに'Hello World'と入力して」
「現在のワークブックの全シート名を取得して」
```

### データ処理

```
「1行目から50行目のデータを読み取って傾向を分析して」
「A列の値を書式付きでB列にコピーして」
「A1:A10のセルに連番を入力して」
```

### VBA自動化

```
「A列をソートするVBAマクロを作成して」
「条件付き書式を適用するVBAコードを実行して」
「サマリーレポートを生成するマクロを実行して」
```

## トラブルシューティング

1. **Excelが見つからない**: Excelが起動しており、少なくとも1つのワークブックが開いていることを確認
2. **VBAエラー**: Excelのマクロセキュリティ設定を確認
3. **権限エラー**: 一部の操作には「VBAプロジェクトオブジェクトモデルへの信頼できるアクセス」の有効化が必要
## セキュリティ注意事項

- サーバーはローカルのExcelファイルのみを操作
- VBAコードは実行後に削除される一時モジュールで実行
## ライセンス

MIT