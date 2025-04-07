# Spreadsheet MCP

Google Spreadsheet(Google Sheets)を操作するためのMCPサーバー

## 機能

- Google Spreadsheetのデータを取得してMarkdownテーブル形式で表示
- Google Spreadsheetのシートのカラム名（ヘッダー行）の一覧を取得
- Google Spreadsheet内で、指定されたシート名のシートのURLを特定
- Google Spreadsheet内に存在するシート名の一覧を取得
- CSVファイルをGoogle Spreadsheetにアップロード
- Google Spreadsheetのシートの内容を別のシートに転記

表形式のシート(１行目にカラムがあり、カラム名に重複がない)であることを想定しています。

## 認証

https://docs.gspread.org/en/latest/oauth2.html

## 設定

```json
{
  "mcpServers": {
    "spreadsheet": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/yamitzky/spreadsheet-mcp.git",
        "spreadsheet-mcp"
      ]
    }
  }
}
```

## 開発

### 開発環境のセットアップ

```bash
# リポジトリのクローン
git clone https://github.com/yamitzky/spreadsheet-mcp.git
cd spreadsheet-mcp

# 開発用依存関係のインストール
uv sync
```

### テストとリンティング

```bash
# ruffでリンティングを実行
ruff check .

# pyrightで型チェックを実行
pyright
```
