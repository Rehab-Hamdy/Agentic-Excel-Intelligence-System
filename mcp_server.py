"""
MCP Server — Exposes Python Engine capabilities as MCP tools.

This is a thin wrapper layer. Each MCP tool delegates to the corresponding
function in python_engine.py. No business logic belongs here.

Transport: stdio (the Agent spawns this server as a subprocess).
"""

import json
from typing import Any, Optional, Union

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    from mcp.server import MCPServer as FastMCP

import python_engine

# Initialize the MCP Server
mcp = FastMCP("ExcelIntelligenceServer")


def _parse_json_or_obj(val: Any) -> Any:
    """Parse JSON string if string, or return object as-is if already dict/list."""
    if val is None:
        return None
    if isinstance(val, (dict, list)):
        return val
    if isinstance(val, str):
        val_str = val.strip()
        if (val_str.startswith("{") and val_str.endswith("}")) or (val_str.startswith("[") and val_str.endswith("]")):
            try:
                return json.loads(val_str)
            except Exception:
                pass
    return val


# ===================================================================
# Excel Tools
# ===================================================================

@mcp.tool()
def read_excel(file_path: str, sheet_name: Optional[str] = None) -> str:
    """
    Read an Excel file and return its metadata and data.

    Opens the workbook, lists all available sheet names, reads the
    specified sheet (or the first sheet by default) into a structured
    format, and returns column names, data types, row count, and the
    actual data rows.

    Args:
        file_path: Path to the Excel file. Can be an absolute path, or a
                   filename relative to the data/input/ directory.
        sheet_name: Optional specific sheet to read. If not provided,
                    reads the first sheet.

    Returns:
        JSON string with: success, file_path, sheet_names, active_sheet,
        metadata (rows, columns, column_names, dtypes), and data.
    """
    result = python_engine.read_excel(file_path, sheet_name)
    return json.dumps(result, default=str)


@mcp.tool()
def create_excel(
    file_path: str,
    sheets_data: Union[dict, list, str],
    formatting: Optional[Union[dict, str]] = None,
) -> str:
    """
    Create a new Excel workbook with structured data.

    Creates a workbook with one or more sheets, writes the provided data,
    applies formatting, and saves to the data/output/ directory.

    Args:
        file_path: Output file name (relative to data/output/) or absolute path.
        sheets_data: Mapping of sheet names to lists of row dicts (as dict or JSON string).
                     Example: {"Sales": [{"Product": "A", "Revenue": 100}]}
        formatting: Optional formatting options (as dict or JSON string).
                    Example: {"header_color": "4472C4", "bold_header": true}

    Returns:
        JSON string with: success, file_path, sheets_created, rows_per_sheet.
    """
    parsed_data = _parse_json_or_obj(sheets_data)
    parsed_fmt = _parse_json_or_obj(formatting)
    result = python_engine.create_excel(file_path, parsed_data, parsed_fmt)
    return json.dumps(result, default=str)


@mcp.tool()
def edit_excel(
    file_path: str,
    operations: Union[list, str],
    output_path: Optional[str] = None,
) -> str:
    """
    Edit an existing Excel workbook.

    Opens the workbook, applies the specified operations (add columns,
    formulas, update cells, delete/rename columns), preserves existing
    content, and saves to a new file in data/output/.

    Args:
        file_path: Path to the source Excel file.
        operations: List of operation dicts (or JSON string).
                    Supported types: "add_column", "add_formula",
                    "update_cells", "delete_column", "rename_column".

                    Examples:
                    - Add column: {"type": "add_column", "sheet": "Sales",
                      "column": "Profit", "values": [10, 20, 30]}
                    - Add formula: {"type": "add_formula", "sheet": "Sales",
                      "column": "Margin", "formula": "=I{row}/G{row}", "start_row": 2}
                    - Update cells: {"type": "update_cells", "sheet": "Sales",
                      "updates": [{"row": 2, "column": "B", "value": 999}]}
        output_path: Optional output path. Defaults to data/output/<original_name>.

    Returns:
        JSON string with: success, file_path, operations_applied.
    """
    parsed_ops = _parse_json_or_obj(operations)
    result = python_engine.edit_excel(file_path, parsed_ops, output_path)
    return json.dumps(result, default=str)


# ===================================================================
# Analysis Tools
# ===================================================================

@mcp.tool()
def profile_data(file_path: str, sheet_name: Optional[str] = None) -> str:
    """
    Generate a comprehensive data profile for an Excel sheet.

    Returns structured information including: row count, column count,
    column names, data types, missing value counts, duplicate row count,
    descriptive statistics for numerical columns, unique value summaries
    for categorical columns, and column type classifications (numerical,
    categorical, date).

    Args:
        file_path: Path to the Excel file.
        sheet_name: Optional specific sheet to profile.

    Returns:
        JSON string with complete profile data.
    """
    result = python_engine.profile_data(file_path, sheet_name)
    return json.dumps(result, default=str)


@mcp.tool()
def analyze_data(
    file_path: str,
    analysis_type: str,
    params: Optional[Union[dict, str]] = None,
    sheet_name: Optional[str] = None,
) -> str:
    """
    Perform a specific data analysis on an Excel file.

    Supports multiple analysis types, each returning structured results
    with a metric name, a human-readable finding, and supporting data.

    Args:
        file_path: Path to the Excel file.
        analysis_type: Type of analysis. One of:
            - "aggregation": Aggregate columns with sum/mean/max/min/count.
              params: {"columns": ["Revenue"], "functions": ["sum", "mean"]}
            - "group_by": Group and aggregate data.
              params: {"group_by": "Region", "agg_column": "Revenue", "agg_func": "sum"}
            - "ranking": Rank rows by a column.
              params: {"column": "Revenue", "top_n": 5, "ascending": false}
            - "trend": Time-series trend analysis.
              params: {"date_column": "Date", "value_column": "Revenue", "period": "ME"}
            - "comparison": Compare values across categories.
              params: {"category_column": "Region", "value_column": "Revenue"}
            - "correlation": Compute correlation matrix.
              params: {"columns": ["Revenue", "Cost", "Profit"]}
            - "distribution": Analyze column distribution.
              params: {"column": "Revenue"}
            - "outlier_detection": Detect outliers using IQR.
              params: {"column": "Revenue", "threshold": 1.5}
            - "missing_analysis": Analyze missing values.
              params: {}
            - "summary_statistics": Extended descriptive statistics.
              params: {}
            - "custom_query": Filter using Pandas query syntax.
              params: {"query": "Revenue > 1000", "columns": ["Product", "Revenue"]}
        params: Optional dict (or JSON string) with analysis-specific parameters.
        sheet_name: Optional specific sheet.

    Returns:
        JSON string with: success, analysis_type, metric, finding, supporting_data.
    """
    parsed_params = _parse_json_or_obj(params)
    result = python_engine.analyze_data(file_path, analysis_type, parsed_params, sheet_name)
    return json.dumps(result, default=str)


# ===================================================================
# Visualization Tool
# ===================================================================

@mcp.tool()
def create_visualization(
    file_path: str,
    chart_type: str,
    params: Union[dict, str],
    sheet_name: Optional[str] = None,
) -> str:
    """
    Generate a chart/visualization from Excel data and save it as an image.

    Creates publication-quality charts using Plotly. The chart is saved
    as a PNG file in the data/output/ directory.

    Args:
        file_path: Path to the Excel file.
        chart_type: Type of chart. One of:
            - "line": Line chart (for time-series, trends).
            - "bar": Bar chart (for category comparisons).
            - "scatter": Scatter plot (for relationships between variables).
            - "histogram": Histogram (for distributions).
            - "box": Box plot (for distributions and outliers by group).
            - "pie": Pie chart (for proportions/shares).
            - "heatmap": Heatmap (for correlation matrices).
        params: Parameters dict (or JSON string):
            - "x": Column name for x-axis.
            - "y": Column name for y-axis.
            - "color": Optional column for color grouping.
            - "title": Optional chart title.
            - "output_filename": Optional output filename.
            - For pie: use "names" and "values" instead of x/y.
            - For heatmap: use "columns" for the correlation subset.
        sheet_name: Optional specific sheet.

    Returns:
        JSON string with: success, chart_path, chart_type, title.
    """
    parsed_params = _parse_json_or_obj(params)
    result = python_engine.create_visualization(file_path, chart_type, parsed_params, sheet_name)
    return json.dumps(result, default=str)


# ===================================================================
# Validation Tool
# ===================================================================

@mcp.tool()
def validate_result(validation_config: Union[dict, str]) -> str:
    """
    Validate that an operation produced correct results.

    Checks file existence, Excel readability, expected sheets/columns,
    data presence, and chart file existence.

    Args:
        validation_config: Configuration dict (or JSON string):
            - "file_path": Path to Excel file to validate.
            - "expected_sheets": Optional list of expected sheet names.
            - "expected_columns": Optional dict of {sheet: [column_names]}.
            - "non_empty": Optional boolean to check for data.
            - "chart_path": Optional path to chart file to verify.

    Returns:
        JSON string with: valid (boolean), warnings (list), errors (list).
    """
    parsed_config = _parse_json_or_obj(validation_config)
    result = python_engine.validate_result(parsed_config)
    return json.dumps(result, default=str)


# ===================================================================
# Entry point — run the MCP server via stdio transport
# ===================================================================
if __name__ == "__main__":
    mcp.run(transport="stdio")
