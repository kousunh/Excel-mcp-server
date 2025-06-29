*[English README](README.md)*

Microsoft Excelとの包括的な連携を可能にするModel Context Protocol (MCP)サーバーです。AIアシスタントが構造化されたワークフローを通じて、データ分析、セル編集、フォーマット、罫線スタイリング、VBA実行、完全なワークシート管理を実行できます。

## 主要機能

### 優先順位付きワークフローツール
- **ステップ1**: `essential_inspect_excel_data` - データ構造理解のための必須最初ステップ
- **最終ステップ**: `essential_check_excel_format` - レイアウトとフォーマットの必須最終確認

### 高度なExcel操作
- **データ分析・読み取り** - 統計情報付き包括的シート分析
- **セル編集** - 配列サポート付き単一セルまたは範囲編集
- **フォーマット制御** - フォント色、背景色、テキストスタイル、配置
- **罫線管理** - 複数スタイルと色での完全な罫線スタイリング
- **VBA実行** - 簡素化され安定したVBAコード実行
- **ワークブック管理** - マルチワークブック処理とナビゲーション

## 必要な環境

- **Windows OS** (COM統合のため必須)
- **Microsoft Excel** インストール・実行中
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

**Windows (コマンドプロンプト):**
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

### 3. Claude Desktopの設定

`claude_desktop_config.json`に追加:

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

### 4. Cursorの設定

Cursorの設定に追加:

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

## 利用可能ツール

### 優先ツール（順序に従って使用）

#### `essential_inspect_excel_data` - ステップ1 - 常に最初に使用
あらゆる操作前にExcelデータ構造と内容を分析します。現在の状態、シート構造、データタイプ、内容を理解するために必須。開いたファイルと閉じたファイルの両方で動作。

#### `essential_check_excel_format` - 最終ステップ - 必須確認
あらゆる変更後にレイアウトとフォーマットを検証します。異なる範囲を確認するために複数回使用。レイアウト・フォーマット問題が見つかった場合は修正して再確認。

### コア操作

#### `edit_cells`
大量データ操作用に最適化されたパフォーマンスでの単一セルまたは範囲編集。

#### `set_cell_formats`
フォント色、背景色、太字、斜体、下線、フォントサイズ、フォント名、テキスト配置を含む包括的フォーマットをセル範囲に適用。

#### `set_cell_borders`
セル範囲への詳細な罫線スタイリング適用。様々な罫線スタイル（細線、太線、中線、二重線、点線、破線）と色を異なる位置（上、下、左、右、内側、外側）でサポート。

### ユーティリティツール

#### `get_open_workbooks`
現在開いているすべてのExcelワークブックを一覧表示。

#### `set_active_workbook`
開いているワークブック間を切り替え。

#### `get_all_sheet_names`
ワークブック内のすべてのシートを一覧表示。

#### `navigate_to_sheet`
特定のシートに切り替え。

#### `get_excel_status`
Excelが実行中で応答しているかチェック。

### VBA実行

#### `execute_vba`
ExcelでカスタムVBAコードを実行。一時的なSubプロシージャを作成・実行し、自動的にクリーンアップ。エラーハンドリングと重複回避の機能付き。

## 使用例

### 基本ワークフロー
```
1. "まず、現在のExcelデータ構造を分析"
2. "セルA1:C3に従業員データを編集"
3. "データ範囲に罫線を設定"
4. "太字ヘッダーと色付き背景でフォーマット適用"
5. "最後に、レイアウトとフォーマットを確認"
```

### データ操作
```
"ワークブック'Sales2024.xlsx'の売上データを分析"
"範囲A1:D10に四半期売上数値を編集"
"テーブル構造を作成するために罫線を適用"
"太字フォントと青い背景でヘッダーをフォーマット"
"最終レイアウトが正しく見えるか確認"
```

## ワークフローベストプラクティス

1. **常に開始** `essential_inspect_excel_data`で現在の状態を理解
2. **専用ツールを使用** VBAではなく（edit_cells, set_cell_formats, set_cell_borders）を可能な限り使用
3. **常に終了** `essential_check_excel_format`で変更を確認
4. **execute_vbaを使用** 標準ツールでは不十分な場合にカスタムVBAロジックで使用
5. **複数範囲を確認** 大きなスプレッドシートで作業する場合

## トラブルシューティング

1. **Excelが見つからない**: Excelが実行中で少なくとも1つのワークブックが開いていることを確認
2. **ツールタイムアウト**: 大規模操作は自動的に延長タイムアウト（60秒）を使用
3. **VBAエラー**: 簡素化されたVBA実行でハング・フリーズ問題を軽減
4. **フォーマット確認**: 異なる範囲に対して確認ツールを複数回使用
5. **権限エラー**: Excelトラストセンターで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効化

## セキュリティ

- サーバーはローカルExcelファイルのみで動作
- VBAコードは自動削除される一時モジュールで実行
- Python仮想環境が依存関係を分離
- ネットワークアクセスや外部ファイル操作なし

## ライセンス

MIT License - 詳細はLICENSEファイルを参照。

## 貢献

1. リポジトリをフォーク
2. フィーチャーブランチを作成
3. 変更を実施
4. Excel操作で十分にテスト
5. プルリクエストを送信

---

**バージョン 2.0 の変更点:**
- 明確な命名による優先順位付きワークフローツールを追加
- 包括的なフォーマットと罫線制御を実装
- 大量データ操作のパフォーマンスを最適化
- VBA実行を簡素化・安定化
- 品質保証のための必須確認ステップを追加