"""
xlsx_utils.py
Utility functions to open and read .xlsx files into pandas DataFrames.
Read-only — no write or modify operations.
Requires: pandas, openpyxl
"""

from typing import Optional
import pandas as pd


# ---------------------------------------------------------------------------
# Inspection
# ---------------------------------------------------------------------------

def get_sheet_names(path: str) -> list[str]:
    """Return the list of sheet names in the workbook."""
    xl = pd.ExcelFile(path, engine="openpyxl")
    return xl.sheet_names


def get_file_info(path: str) -> dict:
    """
    Return basic metadata about the workbook:
    number of sheets, sheet names, and row/column counts per sheet.
    """
    xl = pd.ExcelFile(path, engine="openpyxl")
    info = {"path": path, "sheets": {}}
    for name in xl.sheet_names:
        df = xl.parse(name, header=None)
        info["sheets"][name] = {"rows": len(df), "cols": len(df.columns)}
    return info


# ---------------------------------------------------------------------------
# Reading — single sheet
# ---------------------------------------------------------------------------

def read_sheet(
    path: str,
    sheet: int | str = 0,
    header_row: int = 0,
    skip_rows: Optional[int | list[int]] = None,
    use_cols: Optional[str | list] = None,
    dtype: Optional[dict] = None,
) -> pd.DataFrame:
    """
    Read a single sheet into a DataFrame.

    Args:
        path:       Path to the .xlsx file.
        sheet:      Sheet name or 0-based index (default: first sheet).
        header_row: Row index to use as column names (default: 0).
        skip_rows:  Row(s) to skip after the header.
        use_cols:   Columns to read — accepts an Excel range like "A:D",
                    a list of column names, or a list of column indices.
        dtype:      Dict mapping column names to dtypes, e.g. {"id": str}.

    Returns:
        pd.DataFrame
    """
    return pd.read_excel(
        path,
        sheet_name=sheet,
        header=header_row,
        skiprows=skip_rows,
        usecols=use_cols,
        dtype=dtype,
        engine="openpyxl",
    )


def read_sheet_no_header(path: str, sheet: int | str = 0) -> pd.DataFrame:
    """Read a sheet without treating any row as a header. Columns are integers."""
    return pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")


def read_range(
    path: str,
    sheet: int | str = 0,
    col_range: str = None,
    start_row: int = 0,
    end_row: Optional[int] = None,
) -> pd.DataFrame:
    """
    Read a specific rectangular range from a sheet.

    Args:
        col_range:  Excel-style column range, e.g. "B:E" or "A:Z".
        start_row:  First data row index (0-based, after header).
        end_row:    Last data row index (inclusive). None = all rows.

    Returns:
        pd.DataFrame with the selected slice.
    """
    df = pd.read_excel(
        path,
        sheet_name=sheet,
        usecols=col_range,
        engine="openpyxl",
    )
    return df.iloc[start_row:end_row]


# ---------------------------------------------------------------------------
# Reading — multiple sheets
# ---------------------------------------------------------------------------

def read_all_sheets(path: str, header_row: int = 0) -> dict[str, pd.DataFrame]:
    """
    Read every sheet in the workbook.
    Returns a dict mapping sheet name -> DataFrame.
    """
    return pd.read_excel(
        path,
        sheet_name=None,
        header=header_row,
        engine="openpyxl",
    )


def read_sheets(
    path: str,
    sheets: list[str | int],
    header_row: int = 0,
) -> dict[str, pd.DataFrame]:
    """
    Read a specific subset of sheets by name or index.
    Returns a dict mapping sheet name -> DataFrame.
    """
    return pd.read_excel(
        path,
        sheet_name=sheets,
        header=header_row,
        engine="openpyxl",
    )


# ---------------------------------------------------------------------------
# Filtering & exploration helpers
# ---------------------------------------------------------------------------

def preview(df: pd.DataFrame, n: int = 5) -> pd.DataFrame:
    """Return the first n rows of a DataFrame."""
    return df.head(n)


def get_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a summary DataFrame with column name, dtype, non-null count,
    and percentage of null values for each column.
    """
    total = len(df)
    schema = pd.DataFrame({
        "column": df.columns,
        "dtype": df.dtypes.values,
        "non_null": df.notna().sum().values,
        "null_pct": (df.isna().sum().values / total * 100).round(2),
    })
    return schema.reset_index(drop=True)


def filter_rows(df: pd.DataFrame, column: str, value) -> pd.DataFrame:
    """Return rows where df[column] == value."""
    return df[df[column] == value]


def filter_rows_query(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """
    Filter using a pandas query string.
    Example: filter_rows_query(df, "age > 30 and city == 'Madrid'")
    """
    return df.query(query)


def find_in_column(
    df: pd.DataFrame,
    column: str,
    search_text: str,
    case_sensitive: bool = False,
) -> pd.DataFrame:
    """Return rows where df[column] contains search_text (substring match)."""
    return df[df[column].astype(str).str.contains(search_text, case=case_sensitive, na=False)]


def get_unique_values(df: pd.DataFrame, column: str) -> list:
    """Return the list of unique values in a column."""
    return df[column].dropna().unique().tolist()


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the DataFrame with fully empty rows removed."""
    return df.dropna(how="all").reset_index(drop=True)


def drop_empty_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the DataFrame with fully empty columns removed."""
    return df.dropna(axis=1, how="all")


def rename_columns(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    """Return a copy of the DataFrame with columns renamed according to mapping."""
    return df.rename(columns=mapping)


def select_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Return a DataFrame with only the specified columns."""
    return df[columns]
