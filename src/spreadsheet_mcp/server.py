from pathlib import Path

from mcp.server.fastmcp import FastMCP

from spreadsheet_mcp.sheet import (
    open_csv,
    open_spreadsheet,
    upload_df_to_worksheet,
    worksheet_to_df,
)

mcp = FastMCP("SpreadsheetMCP")


@mcp.tool()
def load_sheet(spreadsheet_url: str) -> str:
    """Retrieves data from the specified Google Spreadsheet sheet and returns it in Markdown table format."""
    try:
        spreadsheet, worksheet = open_spreadsheet(spreadsheet_url)
        if not worksheet:
            worksheet = spreadsheet.sheet1
            if not worksheet:
                raise ValueError("Sheet not found.")
        df = worksheet_to_df(worksheet)

        # Generate Markdown table directly from polars
        header = "| " + " | ".join(df.columns) + " |"
        separator = "|:-" + "-|".join(["-" for _ in df.columns]) + "-|"
        rows = ["| " + " | ".join(map(str, row)) + " |" for row in df.iter_rows()]
        return "\n".join([header, separator] + rows)
    except Exception as e:
        return f"Error loading sheet: {e}"


@mcp.tool()
def get_column_names(spreadsheet_url: str) -> list[str]:
    """Retrieves the list of column names (header row) from the specified Google Spreadsheet sheet."""
    try:
        spreadsheet, worksheet = open_spreadsheet(spreadsheet_url)
        if not worksheet:
            worksheet = spreadsheet.sheet1
            if not worksheet:
                raise ValueError("Sheet not found.")

        header = worksheet.row_values(1)

        return header if header else []
    except Exception as e:
        return [f"Error getting column names: {e}"]


@mcp.tool()
def detect_sheet_url(spreadsheet_url: str, sheet_name: str) -> str:
    """Identifies the URL of the sheet with the specified name within the given Google Spreadsheet."""
    try:
        spreadsheet, _ = open_spreadsheet(spreadsheet_url.split("#")[0])
        worksheet = spreadsheet.worksheet(sheet_name)
        return worksheet.url
    except Exception as e:
        return f"Error detecting sheet URL: {e}"


@mcp.tool()
def get_sheet_names(spreadsheet_url: str) -> list[str]:
    """Retrieves the list of sheet names present in the specified Google Spreadsheet."""
    try:
        spreadsheet, _ = open_spreadsheet(spreadsheet_url.split("#")[0])
        return [sheet.title for sheet in spreadsheet.worksheets()]
    except Exception as e:
        return [f"Error getting sheet names: {e}"]


@mcp.tool()
def upload_csv_to_spreadsheet(file_name: str, spreadsheet_url: str) -> str:
    """Uploads (overwrites) a CSV file from the local environment to the specified Google Spreadsheet sheet."""

    try:
        csv_path = Path(file_name)
        if not csv_path.exists():
            return f"Error: File not found at {csv_path.absolute()}"

        spreadsheet, to_worksheet = open_spreadsheet(spreadsheet_url)
        if not to_worksheet:
            to_worksheet = spreadsheet.sheet1
            if not to_worksheet:
                raise ValueError("Destination sheet for upload not found.")

        from_df = open_csv(csv_path)
        result = upload_df_to_worksheet(from_df, to_worksheet)

        return f"""Sheet content updated.
- Row count change: {result["row_diff"]}
- Updated columns: {", ".join(result["updated_columns"])}
- Skipped columns: {", ".join(result["skipped_columns"]) if result["skipped_columns"] else "None"}"""
    except Exception as e:
        return f"Error uploading CSV: {e}"


@mcp.tool()
def upload_spreadsheet_to_spreadsheet(
    from_spreadsheet_url: str, to_spreadsheet_url: str
) -> str:
    """Copies (overwrites) the content of one Google Spreadsheet sheet to another (or the same) Google Spreadsheet sheet."""
    try:
        from_spreadsheet, from_worksheet = open_spreadsheet(from_spreadsheet_url)
        if not from_worksheet:
            from_worksheet = from_spreadsheet.sheet1
            if not from_worksheet:
                raise ValueError("Source sheet for copying not found.")

        to_spreadsheet, to_worksheet = open_spreadsheet(to_spreadsheet_url)
        if not to_worksheet:
            to_worksheet = to_spreadsheet.sheet1
            if not to_worksheet:
                raise ValueError("Destination sheet for copying not found.")

        from_df = worksheet_to_df(from_worksheet)
        result = upload_df_to_worksheet(from_df, to_worksheet)

        return f"""Sheet content updated.
- Row count change: {result["row_diff"]}
- Updated columns: {", ".join(result["updated_columns"])}
- Skipped columns: {", ".join(result["skipped_columns"]) if result["skipped_columns"] else "None"}"""
    except Exception as e:
        return f"Error uploading spreadsheet: {e}"
