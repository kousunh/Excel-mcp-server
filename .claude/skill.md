# Excel MCP Skill

Excel ファイルの読み書き・書式設定を行うための MCP ツール群。

## 2つのモード

- **`workbook`** — 開いている Excel ブックをリアルタイム操作（xlwings）
- **`path`** — 閉じた .xlsx ファイルを直接編集（Excel 不要、画像・グラフを壊さない）

どちらか一方を指定する。両方指定しない。

## 基本フロー

### 閉じたファイルを扱う場合

1. `read_cells` で現状を確認（path, range, formats=true）
2. `write_cells` で値を書き込む（path, range, value）
3. `format_cells` で書式を整える（path, range, format）

### 開いている Excel を扱う場合

1. `get_excel_info` でブック名・シート名を取得
2. `read_cells` / `write_cells` / `format_cells` を workbook 指定で使う
3. 複雑な操作は `execute_vba` で VBA を実行（開いている時のみ）

## ツール早見表

| ツール | workbook | path | 必須パラメータ |
|--------|:--------:|:----:|---------------|
| get_excel_info | - | - | なし |
| read_cells | OK | OK | range |
| write_cells | OK | OK | range, value |
| format_cells | OK | OK | range, format |
| execute_vba | OK | ✕ | workbook, code |

## value の指定方法

```
単一値:      "hello"  /  42  /  true
1行:         ["A", "B", "C"]
複数行:      [["A1","B1"], ["A2","B2"]]
1列:         range="A1:A3", value=["x","y","z"]
```

range のサイズと value の形状を合わせる。単一値を範囲に書くと全セルに同じ値が入る。

## format の指定方法

```json
{
  "bold": true,
  "fontSize": 14,
  "fontColor": "#FF0000",
  "backgroundColor": "#FFFF00",
  "textAlign": "center",
  "numberFormat": "#,##0",
  "borders": {
    "outside": { "style": "thin", "color": "#000000" }
  }
}
```

使えるプロパティ: bold, italic, underline, fontSize, fontName, fontColor, backgroundColor, textAlign (left/center/right), verticalAlign (top/middle/bottom), numberFormat, wrapText, borders

borders の位置: top, bottom, left, right, inside, outside
borders の style: thin, medium, thick, double, dotted, dashed, none

## 注意点

- `path` モードは .xlsx 形式のみ対応（.xls, .csv は不可）
- `path` モードで編集してもExcelには即反映されない（ファイルを開き直す必要あり）
- `execute_vba` は開いているブック専用。閉じたファイルには使えない
- sheet を省略すると最初のシート（またはアクティブシート）が対象
- 大きな範囲を一度に読むより、必要な範囲だけ指定する方が高速
