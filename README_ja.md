# Excel MCPサーバー

*[English README](README.md)*

> 現在ツールの安定化に取り組んでいます。もしツールの中で挙動がおかしいものがありましたら、[Issues](https://github.com/kousunh/excel-mcp-server/issues)でお知らせください。

Microsoft Excelとの連携を可能にするModel Context Protocol (MCP)サーバーです。データ読取、セル編集、VBAコード実行、ワークシート管理など、様々な操作をAIアシスタントから実行できます。

## 必要な環境

- **Windows OS** (win32com使用のため必須)
- **Microsoft Excel** インストール済み
- **Node.js** 18以上
- **Python** 3.8以上

## インストール

### 1. リポジトリをクローン

```bash
git clone https://github.com/kousunh/excel-mcp-server.git
cd excel-mcp-server
```

### 2. セットアップスクリプトを実行

セットアップスクリプトがPython仮想環境を作成し、すべての依存関係をインストールします。

**Windows (コマンドプロンプト) の場合:**
```cmd
setup.bat
```

**Windows (PowerShell) の場合:**
```powershell
.\setup.bat
```

**Linux/Mac (WSL) の場合:**
```bash
./setup.sh
```

セットアップスクリプトは以下を実行します:
- Python仮想環境（`venv`）の作成
- Python依存関係（pywin32）のインストール
- Node.js依存関係のインストール
- Claude Desktop用の設定を表示

### 3. Claude Desktopの設定

Claude Desktopの設定ファイルにサーバーを追加します。

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

`mcpServers`セクションに以下を追加:

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

`C:\\path\\to\\excel-mcp-server`を実際のクローン先パスに置き換えてください。

### 4. Cursor IDEの設定

Cursor IDEを使用している場合は、MCPサーバーをCursorの設定に追加できます：

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

`C:\\path\\to\\excel-mcp-server`を実際のクローン先パスに置き換えてください。

## 手動セットアップ（代替方法）

セットアップスクリプトが動作しない場合は、手動でセットアップできます:

1. Python仮想環境を作成:
   ```bash
   python -m venv venv
   ```

2. 仮想環境を有効化:
   
   **Windows:**
   ```cmd
   venv\Scripts\activate
   ```
   
   **Linux/Mac:**
   ```bash
   source venv/bin/activate
   ```

3. Python依存関係をインストール:
   ```bash
   pip install -r requirements.txt
   ```

4. Node.js依存関係をインストール:
   ```bash
   npm install
   ```

## 使用方法

使用前にExcelを起動し、ワークブックを開いてください。サーバーは以下のツールを提供します:

### コアツール

- **read_sheet_data** - 完全な列保持でワークシートデータを読取
- **edit_cells** - 単一セルまたは範囲を編集
- **execute_vba** - 自動リトライ付きでVBAコードを実行
- **get_cell_formats** - セル書式の詳細を取得（色、フォント、罫線）

### ワークブック管理

- **get_open_workbooks** - 開いているExcelワークブックを一覧表示
- **set_active_workbook** - 開いているワークブック間を切替
- **get_excel_status** - Excelが実行中かチェック

### シートナビゲーション

- **get_all_sheet_names** - ワークブック内のすべてのシートを一覧表示
- **navigate_to_sheet** - 特定のシートに切替

## 使用例

### 基本操作

```
"アクティブなExcelシートからデータを読み込む"
"セルA1に'Hello World'を書き込む"
"現在のワークブックのすべてのシート名を取得"
```

### データ処理

```
"1-50行のデータを読み込んでトレンドを分析"
"A列からB列に書式付きで値をコピー"
"A1:A10セルに連番を入力"
```

### VBA自動化

```
"A列をソートするVBAマクロを作成"
"条件付き書式を適用するVBAコードを実行"
"サマリーレポートを生成するマクロを実行"
```

## トラブルシューティング

1. **Excelが見つからない**: Excelが実行中で、少なくとも1つのワークブックが開いていることを確認
2. **Pythonが見つからない**: Python 3.8以上がインストールされ、PATHに含まれていることを確認
3. **インポートエラー**: セットアップスクリプトを再実行するか、手動で依存関係をインストール
4. **VBAエラー**: Excelのマクロセキュリティ設定を確認
5. **権限エラー**: 一部の操作には「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」の有効化が必要

## 開発

貢献や修正をする場合:

1. リポジトリをフォーク
2. フィーチャーブランチを作成
3. 変更を実施
4. 十分にテスト
5. プルリクエストを送信

## セキュリティに関する注意

- サーバーはローカルのExcelファイルのみを操作します
- VBAコードは実行後に削除される一時モジュールで実行されます
- Python仮想環境により依存関係が分離されます

## ライセンス

MIT