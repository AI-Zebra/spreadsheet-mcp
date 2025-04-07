from pathlib import Path

from mcp.server.fastmcp import FastMCP

from sheet import open_csv, open_spreadsheet, upload_df_to_worksheet, worksheet_to_df

mcp = FastMCP("SpreadsheetMCP")


@mcp.tool()
def load_sheet(spreadsheet_url: str) -> str:
    """指定されたGoogle Spreadsheetのシートからデータを取得し、Markdownテーブル形式で返します。"""
    try:
        spreadsheet, worksheet = open_spreadsheet(spreadsheet_url)
        if not worksheet:
            worksheet = spreadsheet.sheet1
            if not worksheet:
                raise ValueError("シートが見つかりませんでした。")
        df = worksheet_to_df(worksheet)

        # polars から直接 Markdown テーブルを生成
        header = "| " + " | ".join(df.columns) + " |"
        separator = "|-" + "-|".join(["-" for _ in df.columns]) + "-|"
        rows = ["| " + " | ".join(map(str, row)) + " |" for row in df.iter_rows()]
        return "\n".join([header, separator] + rows)
    except Exception as e:
        return f"Error loading sheet: {e}"


@mcp.tool()
def get_column_names(spreadsheet_url: str) -> list[str]:
    """指定されたGoogle Spreadsheetのシートのカラム名（ヘッダー行）の一覧を取得します。"""
    try:
        spreadsheet, worksheet = open_spreadsheet(spreadsheet_url)
        if not worksheet:
            worksheet = spreadsheet.sheet1
            if not worksheet:
                raise ValueError("シートが見つかりませんでした。")

        header = worksheet.row_values(1)

        return header if header else []
    except Exception as e:
        return [f"Error getting column names: {e}"]


@mcp.tool()
def detect_sheet_url(spreadsheet_url: str, sheet_name: str) -> str:
    """指定されたGoogle Spreadsheet内で、指定されたシート名のシートのURLを特定します。"""
    try:
        spreadsheet, _ = open_spreadsheet(spreadsheet_url.split("#")[0])
        worksheet = spreadsheet.worksheet(sheet_name)
        return worksheet.url
    except Exception as e:
        return f"Error detecting sheet URL: {e}"


@mcp.tool()
def get_sheet_names(spreadsheet_url: str) -> list[str]:
    """指定されたGoogle Spreadsheet内に存在するシートの名前の一覧を取得します。"""
    try:
        spreadsheet, _ = open_spreadsheet(spreadsheet_url.split("#")[0])
        return [sheet.title for sheet in spreadsheet.worksheets()]
    except Exception as e:
        return [f"Error getting sheet names: {e}"]


@mcp.tool()
def upload_csv_to_spreadsheet(file_name: str, spreadsheet_url: str) -> str:
    """ローカル環境にあるCSVファイルを指定されたGoogle Spreadsheetのシートにアップロード（上書き）します。"""

    try:
        csv_path = Path(file_name)
        if not csv_path.exists():
            return f"Error: File not found at {csv_path.absolute()}"

        spreadsheet, to_worksheet = open_spreadsheet(spreadsheet_url)
        if not to_worksheet:
            to_worksheet = spreadsheet.sheet1
            if not to_worksheet:
                raise ValueError("アップロード先のシートが見つかりませんでした。")

        from_df = open_csv(csv_path)
        result = upload_df_to_worksheet(from_df, to_worksheet)

        return f"""シートの内容を更新しました。
- 行数変化: {result['row_diff']}
- 更新されたカラム: {', '.join(result['updated_columns'])}
- スキップされたカラム: {', '.join(result['skipped_columns']) if result['skipped_columns'] else 'なし'}"""
    except Exception as e:
        return f"Error uploading CSV: {e}"


@mcp.tool()
def upload_spreadsheet_to_spreadsheet(
    from_spreadsheet_url: str, to_spreadsheet_url: str
) -> str:
    """あるGoogle Spreadsheetのシートの内容を、別の（または同じ）Google Spreadsheetのシートに転記（上書き）します。"""
    try:
        from_spreadsheet, from_worksheet = open_spreadsheet(from_spreadsheet_url)
        if not from_worksheet:
            from_worksheet = from_spreadsheet.sheet1
            if not from_worksheet:
                raise ValueError("転記元のシートが見つかりませんでした。")

        to_spreadsheet, to_worksheet = open_spreadsheet(to_spreadsheet_url)
        if not to_worksheet:
            to_worksheet = to_spreadsheet.sheet1
            if not to_worksheet:
                raise ValueError("転記先のシートが見つかりませんでした。")

        from_df = worksheet_to_df(from_worksheet)
        result = upload_df_to_worksheet(from_df, to_worksheet)

        return f"""シートの内容を更新しました。
- 行数変化: {result['row_diff']}
- 更新されたカラム: {', '.join(result['updated_columns'])}
- スキップされたカラム: {', '.join(result['skipped_columns']) if result['skipped_columns'] else 'なし'}"""
    except Exception as e:
        return f"Error uploading spreadsheet: {e}"
