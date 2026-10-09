"""
Python Engine — Core computation module for the Excel Intelligence System.

This module contains all data processing, analysis, and visualization logic.
It uses Pandas, OpenPyXL, NumPy, SciPy, Plotly, and Matplotlib.

NO LLM or Agent logic belongs here. Functions receive structured inputs
and return structured (JSON-serializable) outputs.
"""

import os
import json
import traceback
from pathlib import Path
from typing import Any, Optional

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import matplotlib
matplotlib.use("Agg")  # Non-interactive backend
import matplotlib.pyplot as plt
import scipy.stats as stats

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

# Ensure directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _resolve_input_path(file_path: str) -> Path:
    """Resolve a file path — if relative, look in data/input/."""
    p = Path(file_path)
    if p.is_absolute():
        return p
    return INPUT_DIR / p


def _resolve_output_path(file_path: str) -> Path:
    """Resolve an output file path — if relative, put in data/output/."""
    p = Path(file_path)
    if p.is_absolute():
        return p
    return OUTPUT_DIR / p


def _df_to_serializable(df: pd.DataFrame, max_rows: int = 500) -> dict:
    """Convert a DataFrame to a JSON-serializable dict with truncation."""
    truncated = len(df) > max_rows
    sample = df.head(max_rows)
    # Replace NaN/Inf with None for JSON compatibility
    records = json.loads(sample.to_json(orient="records", date_format="iso"))
    return {
        "columns": list(df.columns),
        "row_count": len(df),
        "truncated": truncated,
        "rows_shown": len(sample),
        "data": records,
    }


# ===================================================================
# 1. READ EXCEL
# ===================================================================
def read_excel(
    file_path: str,
    sheet_name: Optional[str] = None,
) -> dict:
    """
    Open an Excel workbook and return structured metadata + data.

    Parameters
    ----------
    file_path : str
        Path to the Excel file (absolute, or relative to data/input/).
    sheet_name : str, optional
        Specific sheet to read. If None, reads the first sheet.

    Returns
    -------
    dict with keys: success, file_path, sheet_names, active_sheet, metadata, data
    """
    try:
        resolved = _resolve_input_path(file_path)
        if not resolved.exists():
            return {"success": False, "error": f"File not found: {resolved}"}

        # Get sheet names via openpyxl (lightweight)
        wb = load_workbook(str(resolved), read_only=True, data_only=True)
        all_sheets = wb.sheetnames
        wb.close()

        # Determine which sheet to read
        target_sheet = sheet_name if sheet_name else all_sheets[0]
        if target_sheet not in all_sheets:
            return {
                "success": False,
                "error": f"Sheet '{target_sheet}' not found. Available: {all_sheets}",
            }

        # Read into pandas
        df = pd.read_excel(str(resolved), sheet_name=target_sheet, engine="openpyxl")

        # Build metadata
        metadata = {
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
        }

        return {
            "success": True,
            "file_path": str(resolved),
            "sheet_names": all_sheets,
            "active_sheet": target_sheet,
            "metadata": metadata,
            "data": _df_to_serializable(df),
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


# ===================================================================
# 2. CREATE EXCEL
# ===================================================================
def create_excel(
    file_path: str,
    sheets_data: dict[str, list[dict]],
    formatting: Optional[dict] = None,
) -> dict:
    """
    Create a new Excel workbook with the given data.

    Parameters
    ----------
    file_path : str
        Output file name/path (relative to data/output/ if not absolute).
    sheets_data : dict
        Mapping of sheet_name -> list of row dicts.
        Example: {"Sales": [{"Product": "A", "Revenue": 100}, ...]}
    formatting : dict, optional
        Formatting options: {"header_color": "4472C4", "bold_header": True}

    Returns
    -------
    dict with keys: success, file_path, sheets_created, rows_per_sheet
    """
    try:
        resolved = _resolve_output_path(file_path)
        resolved.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        rows_per_sheet = {}

        for sheet_name, rows in sheets_data.items():
            ws = wb.create_sheet(title=sheet_name)
            if not rows:
                rows_per_sheet[sheet_name] = 0
                continue

            df = pd.DataFrame(rows)
            rows_per_sheet[sheet_name] = len(df)

            # Write headers + data
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)

                    # Format header row
                    if r_idx == 1:
                        fmt = formatting or {}
                        if fmt.get("bold_header", True):
                            cell.font = Font(bold=True, color="FFFFFF")
                        header_color = fmt.get("header_color", "4472C4")
                        cell.fill = PatternFill(
                            start_color=header_color,
                            end_color=header_color,
                            fill_type="solid",
                        )
                        cell.alignment = Alignment(horizontal="center")

            # Auto-width columns
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    try:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

        wb.save(str(resolved))
        wb.close()

        return {
            "success": True,
            "file_path": str(resolved),
            "sheets_created": list(sheets_data.keys()),
            "rows_per_sheet": rows_per_sheet,
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


# ===================================================================
# 3. EDIT EXCEL
# ===================================================================
def edit_excel(
    file_path: str,
    operations: list[dict],
    output_path: Optional[str] = None,
) -> dict:
    """
    Edit an existing Excel workbook.

    Parameters
    ----------
    file_path : str
        Path to the source Excel file.
    operations : list[dict]
        List of edit operations. Each operation is a dict with:
        - "type": one of "add_column", "update_cells", "add_formula", "delete_column", "rename_column"
        - "sheet": sheet name (optional, defaults to first sheet)
        - Additional keys depending on the type.

        Examples:
        {"type": "add_column", "sheet": "Sheet1", "column": "Profit",
         "values": [10, 20, 30]}

        {"type": "add_formula", "sheet": "Sheet1", "column": "Margin",
         "formula": "=D{row}/B{row}", "start_row": 2}

        {"type": "update_cells", "sheet": "Sheet1",
         "updates": [{"row": 2, "column": "B", "value": 999}]}

        {"type": "delete_column", "sheet": "Sheet1", "column": "TempCol"}

        {"type": "rename_column", "sheet": "Sheet1",
         "old_name": "Rev", "new_name": "Revenue"}

    output_path : str, optional
        Where to save the result. Defaults to data/output/<original_name>.

    Returns
    -------
    dict with keys: success, file_path, operations_applied
    """
    try:
        resolved_input = _resolve_input_path(file_path)
        if not resolved_input.exists():
            # Also check output dir (might be editing a previously created file)
            resolved_input = _resolve_output_path(file_path)
            if not resolved_input.exists():
                return {"success": False, "error": f"File not found: {file_path}"}

        # Determine output path
        if output_path:
            resolved_output = _resolve_output_path(output_path)
        else:
            resolved_output = OUTPUT_DIR / Path(file_path).name

        resolved_output.parent.mkdir(parents=True, exist_ok=True)

        # Load workbook (preserve formatting)
        wb = load_workbook(str(resolved_input))
        ops_applied = []

        for op in operations:
            op_type = op.get("type")
            sheet_name = op.get("sheet", wb.sheetnames[0])

            if sheet_name not in wb.sheetnames:
                ops_applied.append(
                    {"type": op_type, "status": "error", "error": f"Sheet '{sheet_name}' not found"}
                )
                continue

            ws = wb[sheet_name]

            try:
                if op_type == "add_column":
                    _op_add_column(ws, op)
                    ops_applied.append({"type": "add_column", "column": op["column"], "status": "ok"})

                elif op_type == "add_formula":
                    _op_add_formula(ws, op)
                    ops_applied.append({"type": "add_formula", "column": op["column"], "status": "ok"})

                elif op_type == "update_cells":
                    count = _op_update_cells(ws, op)
                    ops_applied.append({"type": "update_cells", "cells_updated": count, "status": "ok"})

                elif op_type == "delete_column":
                    _op_delete_column(ws, op)
                    ops_applied.append({"type": "delete_column", "column": op["column"], "status": "ok"})

                elif op_type == "rename_column":
                    _op_rename_column(ws, op)
                    ops_applied.append(
                        {"type": "rename_column", "old": op["old_name"], "new": op["new_name"], "status": "ok"}
                    )
                else:
                    ops_applied.append({"type": op_type, "status": "error", "error": f"Unknown operation type"})

            except Exception as ex:
                ops_applied.append({"type": op_type, "status": "error", "error": str(ex)})

        wb.save(str(resolved_output))
        wb.close()

        return {
            "success": True,
            "file_path": str(resolved_output),
            "operations_applied": ops_applied,
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


def _find_col_index(ws, col_name: str) -> Optional[int]:
    """Find the 1-based column index by header name in row 1."""
    for cell in ws[1]:
        if cell.value and str(cell.value).strip() == str(col_name).strip():
            return cell.column
    return None


def _op_add_column(ws, op: dict):
    """Add a new column with values."""
    col_name = op["column"]
    values = op.get("values", [])
    # Find next empty column
    next_col = ws.max_column + 1
    ws.cell(row=1, column=next_col, value=col_name)
    ws.cell(row=1, column=next_col).font = Font(bold=True)
    for i, val in enumerate(values, start=2):
        ws.cell(row=i, column=next_col, value=val)


def _op_add_formula(ws, op: dict):
    """Add a column with Excel formulas."""
    col_name = op["column"]
    formula_template = op["formula"]
    start_row = op.get("start_row", 2)
    next_col = ws.max_column + 1
    ws.cell(row=1, column=next_col, value=col_name)
    ws.cell(row=1, column=next_col).font = Font(bold=True)
    for row in range(start_row, ws.max_row + 1):
        formula = formula_template.replace("{row}", str(row))
        ws.cell(row=row, column=next_col, value=formula)


def _op_update_cells(ws, op: dict) -> int:
    """Update specific cells."""
    updates = op.get("updates", [])
    count = 0
    for u in updates:
        row = u["row"]
        col = u["column"]
        val = u["value"]
        # col can be a letter ("B") or a number
        if isinstance(col, str) and col.isalpha():
            ws[f"{col}{row}"] = val
        else:
            ws.cell(row=row, column=int(col), value=val)
        count += 1
    return count


def _op_delete_column(ws, op: dict):
    """Delete a column by name."""
    col_idx = _find_col_index(ws, op["column"])
    if col_idx is None:
        raise ValueError(f"Column '{op['column']}' not found")
    ws.delete_cols(col_idx)


def _op_rename_column(ws, op: dict):
    """Rename a column header."""
    col_idx = _find_col_index(ws, op["old_name"])
    if col_idx is None:
        raise ValueError(f"Column '{op['old_name']}' not found")
    ws.cell(row=1, column=col_idx, value=op["new_name"])


# ===================================================================
# 4. PROFILE DATA
# ===================================================================
def profile_data(
    file_path: str,
    sheet_name: Optional[str] = None,
) -> dict:
    """
    Return a comprehensive data profile for an Excel sheet.

    Returns
    -------
    dict with: rows, columns, column_names, dtypes, missing_values,
    duplicate_rows, basic_statistics, numerical_columns, categorical_columns,
    date_columns.
    """
    try:
        resolved = _resolve_input_path(file_path)
        if not resolved.exists():
            resolved = _resolve_output_path(file_path)
            if not resolved.exists():
                return {"success": False, "error": f"File not found: {file_path}"}

        target_sheet = sheet_name if sheet_name is not None else 0
        df = pd.read_excel(str(resolved), sheet_name=target_sheet, engine="openpyxl")

        # Classify columns
        numerical_cols = df.select_dtypes(include=["number"]).columns.tolist()
        categorical_cols = df.select_dtypes(include=["object", "category", "str"]).columns.tolist()
        date_cols = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()

        # Missing values
        missing = df.isnull().sum()
        missing_info = {col: int(count) for col, count in missing.items() if count > 0}

        # Basic statistics for numerical columns
        basic_stats = {}
        if numerical_cols:
            desc = df[numerical_cols].describe()
            basic_stats = json.loads(desc.to_json())

        # Unique value counts for categorical columns (top 10)
        categorical_info = {}
        for col in categorical_cols:
            vc = df[col].value_counts().head(10)
            categorical_info[col] = {
                "unique_count": int(df[col].nunique()),
                "top_values": {str(k): int(v) for k, v in vc.items()},
            }

        return {
            "success": True,
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "missing_values": missing_info,
            "total_missing": int(missing.sum()),
            "duplicate_rows": int(df.duplicated().sum()),
            "basic_statistics": basic_stats,
            "numerical_columns": numerical_cols,
            "categorical_columns": categorical_cols,
            "date_columns": date_cols,
            "categorical_summary": categorical_info,
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


# ===================================================================
# 5. ANALYZE DATA
# ===================================================================
def analyze_data(
    file_path: str,
    analysis_type: str,
    params: Optional[dict] = None,
    sheet_name: Optional[str] = None,
) -> dict:
    """
    Perform a specific analysis on Excel data.

    Parameters
    ----------
    file_path : str
    analysis_type : str
        One of: "aggregation", "group_by", "ranking", "trend", "comparison",
        "correlation", "distribution", "outlier_detection", "missing_analysis",
        "summary_statistics", "custom_query"
    params : dict
        Analysis-specific parameters. See each handler for details.
    sheet_name : str, optional

    Returns
    -------
    dict with: success, analysis_type, metric, finding, supporting_data
    """
    try:
        resolved = _resolve_input_path(file_path)
        if not resolved.exists():
            resolved = _resolve_output_path(file_path)
            if not resolved.exists():
                return {"success": False, "error": f"File not found: {file_path}"}

        target_sheet = sheet_name if sheet_name is not None else 0
        df = pd.read_excel(str(resolved), sheet_name=target_sheet, engine="openpyxl")
        params = params or {}

        handlers = {
            "aggregation": _analyze_aggregation,
            "group_by": _analyze_group_by,
            "ranking": _analyze_ranking,
            "trend": _analyze_trend,
            "comparison": _analyze_comparison,
            "correlation": _analyze_correlation,
            "distribution": _analyze_distribution,
            "outlier_detection": _analyze_outlier_detection,
            "missing_analysis": _analyze_missing,
            "summary_statistics": _analyze_summary_statistics,
            "custom_query": _analyze_custom_query,
        }

        handler = handlers.get(analysis_type)
        if handler is None:
            return {
                "success": False,
                "error": f"Unknown analysis_type: '{analysis_type}'. Available: {list(handlers.keys())}",
            }

        result = handler(df, params)
        result["success"] = True
        result["analysis_type"] = analysis_type
        return result

    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


def _analyze_aggregation(df: pd.DataFrame, params: dict) -> dict:
    """
    Aggregate one or more columns.
    params: {"columns": ["Revenue"], "functions": ["sum", "mean", "max", "min"]}
    """
    columns = params.get("columns", df.select_dtypes(include="number").columns.tolist())
    functions = params.get("functions", ["sum", "mean", "max", "min", "count"])

    results = {}
    for col in columns:
        if col not in df.columns:
            results[col] = {"error": f"Column not found"}
            continue
        col_results = {}
        for func in functions:
            try:
                val = getattr(df[col], func)()
                col_results[func] = round(float(val), 4) if isinstance(val, (int, float, np.floating, np.integer)) else str(val)
            except Exception as ex:
                col_results[func] = f"error: {ex}"
        results[col] = col_results

    return {
        "metric": "aggregation",
        "finding": f"Aggregated {len(columns)} column(s) with {len(functions)} function(s).",
        "supporting_data": results,
    }


def _analyze_group_by(df: pd.DataFrame, params: dict) -> dict:
    """
    Group-by analysis.
    params: {"group_by": "Region", "agg_column": "Revenue", "agg_func": "sum"}
    """
    group_col = params.get("group_by")
    agg_col = params.get("agg_column")
    agg_func = params.get("agg_func", "sum")

    if not group_col or not agg_col:
        return {"error": "params must include 'group_by' and 'agg_column'"}

    grouped = df.groupby(group_col)[agg_col].agg(agg_func).reset_index()
    grouped.columns = [group_col, f"{agg_col}_{agg_func}"]
    grouped = grouped.sort_values(by=grouped.columns[1], ascending=False)

    return {
        "metric": f"{agg_col}_{agg_func}_by_{group_col}",
        "finding": f"Grouped '{agg_col}' by '{group_col}' using '{agg_func}'.",
        "supporting_data": _df_to_serializable(grouped),
    }


def _analyze_ranking(df: pd.DataFrame, params: dict) -> dict:
    """
    Rank rows by a column.
    params: {"column": "Revenue", "top_n": 5, "ascending": False}
    """
    col = params.get("column")
    top_n = params.get("top_n", 10)
    ascending = params.get("ascending", False)

    if not col:
        return {"error": "params must include 'column'"}

    ranked = df.sort_values(by=col, ascending=ascending).head(top_n)

    return {
        "metric": f"top_{top_n}_{col}",
        "finding": f"Top {top_n} rows ranked by '{col}' ({'ascending' if ascending else 'descending'}).",
        "supporting_data": _df_to_serializable(ranked),
    }


def _analyze_trend(df: pd.DataFrame, params: dict) -> dict:
    """
    Analyze time-series trend.
    params: {"date_column": "Date", "value_column": "Revenue", "period": "M"}
    """
    date_col = params.get("date_column")
    value_col = params.get("value_column")
    period = params.get("period", "ME")  # ME=monthly, QE=quarterly, YE=yearly

    # Pandas 3.0 renamed frequency aliases: M→ME, Q→QE, Y→YE
    freq_map = {"M": "ME", "Q": "QE", "Y": "YE"}
    period = freq_map.get(period, period)

    if not date_col or not value_col:
        return {"error": "params must include 'date_column' and 'value_column'"}

    df_copy = df.copy()
    df_copy[date_col] = pd.to_datetime(df_copy[date_col], errors="coerce")
    df_copy = df_copy.dropna(subset=[date_col])
    df_copy = df_copy.set_index(date_col)

    trend = df_copy[value_col].resample(period).sum().reset_index()
    trend.columns = [date_col, value_col]

    # Calculate overall change
    if len(trend) >= 2:
        first_val = trend[value_col].iloc[0]
        last_val = trend[value_col].iloc[-1]
        if first_val != 0:
            pct_change = round(((last_val - first_val) / first_val) * 100, 2)
            direction = "increased" if pct_change > 0 else "decreased"
            finding = f"{value_col} {direction} by {abs(pct_change)}% over the period."
        else:
            finding = f"{value_col} trend computed (first value was 0)."
    else:
        finding = f"Insufficient data for trend analysis."

    # Convert dates to strings for JSON serialization
    trend[date_col] = trend[date_col].dt.strftime("%Y-%m-%d")

    return {
        "metric": f"{value_col}_trend",
        "finding": finding,
        "supporting_data": _df_to_serializable(trend),
    }


def _analyze_comparison(df: pd.DataFrame, params: dict) -> dict:
    """
    Compare values across categories.
    params: {"category_column": "Region", "value_column": "Revenue", "compare_func": "mean"}
    """
    cat_col = params.get("category_column")
    val_col = params.get("value_column")
    func = params.get("compare_func", "mean")

    if not cat_col or not val_col:
        return {"error": "params must include 'category_column' and 'value_column'"}

    comparison = df.groupby(cat_col)[val_col].agg(func).reset_index()
    comparison.columns = [cat_col, f"{val_col}_{func}"]
    comparison = comparison.sort_values(by=comparison.columns[1], ascending=False)

    top = comparison.iloc[0]
    bottom = comparison.iloc[-1]
    finding = (
        f"Highest {func} {val_col}: {top[cat_col]} ({round(float(top.iloc[1]), 2)}). "
        f"Lowest: {bottom[cat_col]} ({round(float(bottom.iloc[1]), 2)})."
    )

    return {
        "metric": f"{val_col}_comparison_by_{cat_col}",
        "finding": finding,
        "supporting_data": _df_to_serializable(comparison),
    }


def _analyze_correlation(df: pd.DataFrame, params: dict) -> dict:
    """
    Compute correlation matrix.
    params: {"columns": ["Revenue", "Cost", "Quantity"]}  # optional subset
    """
    columns = params.get("columns")
    num_df = df.select_dtypes(include="number")
    if columns:
        num_df = num_df[[c for c in columns if c in num_df.columns]]

    if num_df.shape[1] < 2:
        return {"metric": "correlation", "finding": "Need at least 2 numerical columns.", "supporting_data": {}}

    corr = num_df.corr().round(4)

    # Find strongest correlations (excluding self-correlation)
    strong = []
    for i in range(len(corr.columns)):
        for j in range(i + 1, len(corr.columns)):
            val = corr.iloc[i, j]
            if abs(val) >= 0.5:
                strong.append({
                    "column_1": corr.columns[i],
                    "column_2": corr.columns[j],
                    "correlation": float(val),
                })

    strong.sort(key=lambda x: abs(x["correlation"]), reverse=True)

    return {
        "metric": "correlation",
        "finding": f"Found {len(strong)} strong correlation(s) (|r| >= 0.5).",
        "supporting_data": {
            "correlation_matrix": json.loads(corr.to_json()),
            "strong_correlations": strong,
        },
    }


def _analyze_distribution(df: pd.DataFrame, params: dict) -> dict:
    """
    Analyze the distribution of a column.
    params: {"column": "Revenue"}
    """
    col = params.get("column")
    if not col or col not in df.columns:
        return {"error": f"Column '{col}' not found."}

    series = df[col].dropna()
    if not pd.api.types.is_numeric_dtype(series):
        # Categorical distribution
        vc = series.value_counts()
        return {
            "metric": f"{col}_distribution",
            "finding": f"'{col}' has {len(vc)} unique values.",
            "supporting_data": {
                "type": "categorical",
                "unique_count": int(len(vc)),
                "top_values": {str(k): int(v) for k, v in vc.head(20).items()},
            },
        }

    # Numerical distribution
    desc = series.describe()
    skewness = round(float(series.skew()), 4)
    kurtosis = round(float(series.kurtosis()), 4)

    # Normality test (if enough data)
    normality = None
    if len(series) >= 8:
        try:
            stat, p_val = stats.shapiro(series.sample(min(5000, len(series)), random_state=42))
            normality = {"statistic": round(float(stat), 4), "p_value": round(float(p_val), 4), "is_normal": p_val > 0.05}
        except Exception:
            pass

    return {
        "metric": f"{col}_distribution",
        "finding": f"'{col}': mean={round(float(desc['mean']), 2)}, std={round(float(desc['std']), 2)}, skewness={skewness}.",
        "supporting_data": {
            "type": "numerical",
            "count": int(desc["count"]),
            "mean": round(float(desc["mean"]), 4),
            "std": round(float(desc["std"]), 4),
            "min": round(float(desc["min"]), 4),
            "q25": round(float(desc["25%"]), 4),
            "median": round(float(desc["50%"]), 4),
            "q75": round(float(desc["75%"]), 4),
            "max": round(float(desc["max"]), 4),
            "skewness": skewness,
            "kurtosis": kurtosis,
            "normality_test": normality,
        },
    }


def _analyze_outlier_detection(df: pd.DataFrame, params: dict) -> dict:
    """
    Detect outliers using IQR method.
    params: {"column": "Revenue", "method": "iqr", "threshold": 1.5}
    """
    col = params.get("column")
    threshold = params.get("threshold", 1.5)

    if not col or col not in df.columns:
        return {"error": f"Column '{col}' not found."}

    series = df[col].dropna()
    if not pd.api.types.is_numeric_dtype(series):
        return {"error": f"Column '{col}' is not numerical."}

    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower = q1 - threshold * iqr
    upper = q3 + threshold * iqr

    outlier_mask = (series < lower) | (series > upper)
    outlier_count = int(outlier_mask.sum())
    outlier_rows = df[outlier_mask].head(50)

    return {
        "metric": f"{col}_outliers",
        "finding": f"Found {outlier_count} outlier(s) in '{col}' (IQR method, threshold={threshold}).",
        "supporting_data": {
            "outlier_count": outlier_count,
            "total_rows": len(series),
            "outlier_percentage": round(outlier_count / len(series) * 100, 2) if len(series) > 0 else 0,
            "lower_bound": round(float(lower), 4),
            "upper_bound": round(float(upper), 4),
            "q1": round(float(q1), 4),
            "q3": round(float(q3), 4),
            "iqr": round(float(iqr), 4),
            "outlier_rows": _df_to_serializable(outlier_rows),
        },
    }


def _analyze_missing(df: pd.DataFrame, params: dict) -> dict:
    """Analyze missing values in detail."""
    missing = df.isnull().sum()
    total = len(df)
    missing_info = {}
    for col in df.columns:
        count = int(missing[col])
        if count > 0:
            missing_info[col] = {
                "count": count,
                "percentage": round(count / total * 100, 2),
            }

    total_missing = int(missing.sum())
    total_cells = int(df.shape[0] * df.shape[1])

    return {
        "metric": "missing_values",
        "finding": f"{total_missing} missing value(s) across {len(missing_info)} column(s) ({round(total_missing / total_cells * 100, 2) if total_cells else 0}% of all cells).",
        "supporting_data": {
            "total_missing": total_missing,
            "total_cells": total_cells,
            "columns_with_missing": missing_info,
            "complete_rows": int((~df.isnull().any(axis=1)).sum()),
            "complete_rows_percentage": round(float((~df.isnull().any(axis=1)).sum()) / total * 100, 2) if total > 0 else 0,
        },
    }


def _analyze_summary_statistics(df: pd.DataFrame, params: dict) -> dict:
    """Extended summary statistics."""
    num_df = df.select_dtypes(include="number")
    if num_df.empty:
        return {"metric": "summary", "finding": "No numerical columns found.", "supporting_data": {}}

    desc = num_df.describe().round(4)
    summary = json.loads(desc.to_json())

    return {
        "metric": "summary_statistics",
        "finding": f"Summary statistics for {len(num_df.columns)} numerical column(s).",
        "supporting_data": summary,
    }


def _analyze_custom_query(df: pd.DataFrame, params: dict) -> dict:
    """
    Flexible query analysis using Pandas query syntax.
    params: {"query": "Revenue > 1000", "columns": ["Product", "Revenue"]}
    """
    query_str = params.get("query")
    columns = params.get("columns")

    if not query_str:
        return {"error": "params must include 'query'"}

    try:
        filtered = df.query(query_str)
    except Exception as e:
        return {"error": f"Query failed: {e}"}

    if columns:
        filtered = filtered[[c for c in columns if c in filtered.columns]]

    return {
        "metric": "custom_query",
        "finding": f"Query '{query_str}' returned {len(filtered)} row(s) out of {len(df)}.",
        "supporting_data": _df_to_serializable(filtered),
    }


# ===================================================================
# 6. CREATE VISUALIZATION
# ===================================================================
def create_visualization(
    file_path: str,
    chart_type: str,
    params: dict,
    sheet_name: Optional[str] = None,
) -> dict:
    """
    Generate a chart from Excel data and save to data/output/.

    Parameters
    ----------
    file_path : str
    chart_type : str
        One of: "line", "bar", "scatter", "histogram", "box", "pie", "heatmap"
    params : dict
        Chart-specific parameters:
        - "x": column for x-axis
        - "y": column for y-axis (or list of columns)
        - "color": column for color grouping (optional)
        - "title": chart title (optional)
        - "output_filename": name for the saved image (optional)
        - Additional type-specific params.
    sheet_name : str, optional

    Returns
    -------
    dict with: success, chart_path, chart_type
    """
    try:
        resolved = _resolve_input_path(file_path)
        if not resolved.exists():
            resolved = _resolve_output_path(file_path)
            if not resolved.exists():
                return {"success": False, "error": f"File not found: {file_path}"}

        target_sheet = sheet_name if sheet_name is not None else 0
        df = pd.read_excel(str(resolved), sheet_name=target_sheet, engine="openpyxl")

        title = params.get("title", f"{chart_type.title()} Chart")
        output_name = params.get("output_filename", f"chart_{chart_type}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}")
        if not output_name.endswith(".html") and not output_name.endswith(".png"):
            output_name += ".png"

        if Path(output_name).is_absolute():
            output_path = Path(output_name)
        else:
            output_path = OUTPUT_DIR / output_name
        output_path.parent.mkdir(parents=True, exist_ok=True)

        chart_handlers = {
            "line": _chart_line,
            "bar": _chart_bar,
            "scatter": _chart_scatter,
            "histogram": _chart_histogram,
            "box": _chart_box,
            "pie": _chart_pie,
            "heatmap": _chart_heatmap,
        }

        handler = chart_handlers.get(chart_type)
        if handler is None:
            return {
                "success": False,
                "error": f"Unknown chart_type: '{chart_type}'. Available: {list(chart_handlers.keys())}",
            }

        fig = handler(df, params, title)

        # Save as PNG (static)
        if str(output_path).endswith(".html"):
            fig.write_html(str(output_path))
        else:
            fig.write_image(str(output_path), width=1200, height=700, scale=2)

        return {
            "success": True,
            "chart_path": str(output_path),
            "chart_type": chart_type,
            "title": title,
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


def _chart_line(df: pd.DataFrame, params: dict, title: str):
    x = params.get("x")
    y = params.get("y")
    color = params.get("color")
    fig = px.line(df, x=x, y=y, color=color, title=title, markers=True)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_bar(df: pd.DataFrame, params: dict, title: str):
    x = params.get("x")
    y = params.get("y")
    color = params.get("color")
    orientation = params.get("orientation", "v")
    fig = px.bar(df, x=x, y=y, color=color, title=title, barmode=params.get("barmode", "group"),
                 orientation=orientation)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_scatter(df: pd.DataFrame, params: dict, title: str):
    x = params.get("x")
    y = params.get("y")
    color = params.get("color")
    size = params.get("size")
    trendline = params.get("trendline")  # "ols" for linear
    fig = px.scatter(df, x=x, y=y, color=color, size=size, title=title, trendline=trendline)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_histogram(df: pd.DataFrame, params: dict, title: str):
    x = params.get("x")
    nbins = params.get("nbins", 30)
    color = params.get("color")
    fig = px.histogram(df, x=x, nbins=nbins, color=color, title=title)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_box(df: pd.DataFrame, params: dict, title: str):
    x = params.get("x")
    y = params.get("y")
    color = params.get("color")
    fig = px.box(df, x=x, y=y, color=color, title=title)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_pie(df: pd.DataFrame, params: dict, title: str):
    names = params.get("names")
    values = params.get("values")
    fig = px.pie(df, names=names, values=values, title=title)
    fig.update_layout(template="plotly_white")
    return fig


def _chart_heatmap(df: pd.DataFrame, params: dict, title: str):
    columns = params.get("columns")
    num_df = df.select_dtypes(include="number")
    if columns:
        num_df = num_df[[c for c in columns if c in num_df.columns]]
    corr = num_df.corr().round(4)

    fig = go.Figure(data=go.Heatmap(
        z=corr.values,
        x=corr.columns.tolist(),
        y=corr.columns.tolist(),
        colorscale="RdBu_r",
        zmin=-1, zmax=1,
        text=corr.values.round(2),
        texttemplate="%{text}",
    ))
    fig.update_layout(title=title, template="plotly_white")
    return fig


# ===================================================================
# 7. VALIDATE RESULT
# ===================================================================
def validate_result(
    validation_config: dict,
) -> dict:
    """
    Validate that an operation produced correct results.

    Parameters
    ----------
    validation_config : dict
        - "file_path": path to check
        - "expected_sheets": list of sheet names (optional)
        - "expected_columns": dict of {sheet: [columns]} (optional)
        - "non_empty": bool — check file has data (optional)
        - "chart_path": path to chart file to verify (optional)
        - "check_calculations": dict — verify specific cell values (optional)
            e.g. {"sheet": "Sheet1", "checks": [{"cell": "E2", "expected_type": "number"}]}

    Returns
    -------
    dict: {"valid": bool, "warnings": [], "errors": []}
    """
    warnings = []
    errors = []

    file_path = validation_config.get("file_path")
    chart_path = validation_config.get("chart_path")

    # --- Validate Excel file ---
    if file_path:
        resolved = _resolve_output_path(file_path)
        if not resolved.exists():
            resolved = _resolve_input_path(file_path)

        if not resolved.exists():
            errors.append(f"File not found: {file_path}")
        else:
            try:
                wb = load_workbook(str(resolved), read_only=True)
                sheets = wb.sheetnames

                # Check expected sheets
                expected_sheets = validation_config.get("expected_sheets")
                if expected_sheets:
                    for s in expected_sheets:
                        if s not in sheets:
                            errors.append(f"Expected sheet '{s}' not found. Available: {sheets}")

                # Check expected columns
                expected_columns = validation_config.get("expected_columns")
                if expected_columns:
                    for sheet, cols in expected_columns.items():
                        if sheet in sheets:
                            ws = wb[sheet]
                            headers = [cell.value for cell in ws[1] if cell.value]
                            for col in cols:
                                if col not in headers:
                                    errors.append(f"Expected column '{col}' not found in sheet '{sheet}'. Found: {headers}")
                        else:
                            errors.append(f"Cannot check columns: sheet '{sheet}' not found.")

                # Check non-empty
                if validation_config.get("non_empty", False):
                    for sheet in sheets:
                        ws = wb[sheet]
                        if ws.max_row is None or ws.max_row <= 1:
                            warnings.append(f"Sheet '{sheet}' appears empty or has only headers.")

                wb.close()
            except Exception as e:
                errors.append(f"Cannot open Excel file: {e}")

    # --- Validate chart file ---
    if chart_path:
        chart_resolved = _resolve_output_path(chart_path)
        if not chart_resolved.exists():
            chart_resolved = Path(chart_path)
        if not chart_resolved.exists():
            errors.append(f"Chart file not found: {chart_path}")
        else:
            if chart_resolved.stat().st_size == 0:
                errors.append(f"Chart file is empty: {chart_path}")

    valid = len(errors) == 0

    return {
        "valid": valid,
        "warnings": warnings,
        "errors": errors,
    }
