from pathlib import Path

import gspread
import polars as pl
from gspread.utils import ValueRenderOption, rowcol_to_a1
from gspread.worksheet import Worksheet


def open_spreadsheet(url: str):
    gc = gspread.oauth()  # type: ignore

    spreadsheet = gc.open_by_url(url)
    worksheet: Worksheet | None = None
    if "#gid=" in url:
        gid = url.split("#gid=")[1]
        worksheet = [ws for ws in spreadsheet.worksheets() if str(ws.id) == gid][0]
    elif "?gid=" in url:
        gid = url.split("?gid=")[1]
        worksheet = [ws for ws in spreadsheet.worksheets() if str(ws.id) == gid][0]

    return spreadsheet, worksheet


def worksheet_to_df(worksheet: Worksheet) -> pl.DataFrame:
    data = worksheet.get_all_records(
        value_render_option=ValueRenderOption.unformatted,
        default_blank=None,  # type: ignore
    )
    return pl.DataFrame(data, infer_schema_length=1000)


def upload_df_to_worksheet(from_df: pl.DataFrame, to_worksheet: Worksheet) -> dict:
    """
    DataFrameのデータをワークシートにアップロードします。

    Returns:
        dict: 以下の情報を含む辞書
            - row_diff (int): 変更があった差分行数
            - updated_columns (list): 更新されたカラム名のリスト
            - skipped_columns (list): アップロードされなかったカラム名のリスト
    """
    to_df = worksheet_to_df(to_worksheet)

    # 共通カラムと、アップロードされないカラムを特定
    common_columns = set(from_df.columns) & set(to_df.columns)
    skipped_columns = list(set(from_df.columns) - common_columns)

    # 行数の差分を計算
    row_diff = from_df.shape[0] - to_worksheet.row_count
    if row_diff > 0:
        to_worksheet.add_rows(row_diff)
    elif row_diff < 0:
        to_worksheet.delete_rows(from_df.shape[0] + 2, to_worksheet.row_count + 1)

    # カラムごとに、データを更新
    batch_update: list[dict] = []
    updated_columns = []

    for col in common_columns:
        col_index = to_df.columns.index(col)
        range_name = f"{rowcol_to_a1(2, col_index + 1)}:{rowcol_to_a1(from_df.shape[0] + 1, col_index + 1)}"
        values = [
            # None will be skipped, so set the string value to an empty string.
            # https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets.values#ValueRange
            [val if val is not None else ""]
            for val in from_df[col].to_list()
        ]
        batch_update.append({"range": range_name, "values": values})
        updated_columns.append(col)

    if batch_update:
        to_worksheet.batch_update(batch_update)

    return {
        "row_diff": row_diff,
        "updated_columns": updated_columns,
        "skipped_columns": skipped_columns,
    }


def open_csv(path: Path | str):
    try:
        return pl.read_csv(path, encoding="utf8")
    except Exception:
        pass
    try:
        return pl.read_csv(path, encoding="shift-jis")
    except Exception:
        pass
    return pl.read_csv(path, encoding="cp932")
