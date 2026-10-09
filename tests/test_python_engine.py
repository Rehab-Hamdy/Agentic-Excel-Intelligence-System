"""
Comprehensive tests for python_engine.py — all 7 functions.
"""
import sys
import json
from pathlib import Path

import pytest
import pandas as pd
import numpy as np

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import python_engine as engine


# ===================================================================
# Helpers
# ===================================================================
@pytest.fixture
def sales_path(sample_excel_path):
    """Return the sample sales.xlsx absolute path as a string."""
    return str(sample_excel_path)


@pytest.fixture
def output_path(output_dir):
    """Return a convenient output directory string."""
    return str(output_dir)


# ===================================================================
# 1. TEST READ EXCEL
# ===================================================================
class TestReadExcel:
    def test_read_existing_file(self, sales_path):
        result = engine.read_excel(sales_path)
        assert result["success"] is True
        assert "Sales" in result["sheet_names"]
        assert result["active_sheet"] == "Sales"
        assert result["metadata"]["rows"] == 100
        assert result["metadata"]["columns"] == 9
        assert len(result["data"]["data"]) > 0

    def test_read_specific_sheet(self, sales_path):
        result = engine.read_excel(sales_path, sheet_name="Sales")
        assert result["success"] is True
        assert result["active_sheet"] == "Sales"

    def test_read_nonexistent_file(self):
        result = engine.read_excel("nonexistent_file.xlsx")
        assert result["success"] is False
        assert "not found" in result["error"].lower()

    def test_read_nonexistent_sheet(self, sales_path):
        result = engine.read_excel(sales_path, sheet_name="NonExistent")
        assert result["success"] is False
        assert "not found" in result["error"].lower()

    def test_metadata_columns(self, sales_path):
        result = engine.read_excel(sales_path)
        expected_cols = ["Date", "Product", "Category", "Region",
                         "Unit_Price", "Quantity", "Revenue", "Cost", "Profit"]
        assert result["metadata"]["column_names"] == expected_cols

    def test_data_serializable(self, sales_path):
        result = engine.read_excel(sales_path)
        # Should be JSON-serializable
        json_str = json.dumps(result, default=str)
        assert len(json_str) > 0


# ===================================================================
# 2. TEST CREATE EXCEL
# ===================================================================
class TestCreateExcel:
    def test_create_simple_workbook(self, output_path):
        path = str(Path(output_path) / "test_create.xlsx")
        result = engine.create_excel(
            file_path=path,
            sheets_data={
                "Products": [
                    {"Name": "A", "Price": 10},
                    {"Name": "B", "Price": 20},
                ],
            },
        )
        assert result["success"] is True
        assert Path(result["file_path"]).exists()
        assert result["sheets_created"] == ["Products"]
        assert result["rows_per_sheet"]["Products"] == 2

    def test_create_multiple_sheets(self, output_path):
        path = str(Path(output_path) / "test_multi_sheet.xlsx")
        result = engine.create_excel(
            file_path=path,
            sheets_data={
                "Sheet1": [{"A": 1}],
                "Sheet2": [{"B": 2}, {"B": 3}],
            },
        )
        assert result["success"] is True
        assert len(result["sheets_created"]) == 2

    def test_create_with_formatting(self, output_path):
        path = str(Path(output_path) / "test_formatted.xlsx")
        result = engine.create_excel(
            file_path=path,
            sheets_data={"Data": [{"X": 1, "Y": 2}]},
            formatting={"header_color": "FF5733", "bold_header": True},
        )
        assert result["success"] is True

    def test_created_file_readable(self, output_path):
        path = str(Path(output_path) / "test_readable.xlsx")
        engine.create_excel(
            file_path=path,
            sheets_data={"Test": [{"Col1": "hello", "Col2": 42}]},
        )
        result = engine.read_excel(path)
        assert result["success"] is True
        assert result["metadata"]["rows"] == 1


# ===================================================================
# 3. TEST EDIT EXCEL
# ===================================================================
class TestEditExcel:
    def test_add_column(self, sales_path, output_path):
        out = str(Path(output_path) / "test_edit_add_col.xlsx")
        result = engine.edit_excel(
            file_path=sales_path,
            operations=[
                {"type": "add_column", "sheet": "Sales", "column": "Test_Col",
                 "values": list(range(100))},
            ],
            output_path=out,
        )
        assert result["success"] is True
        assert result["operations_applied"][0]["status"] == "ok"

        # Verify column exists
        read_result = engine.read_excel(out)
        assert "Test_Col" in read_result["metadata"]["column_names"]

    def test_add_formula(self, sales_path, output_path):
        out = str(Path(output_path) / "test_edit_formula.xlsx")
        result = engine.edit_excel(
            file_path=sales_path,
            operations=[
                {"type": "add_formula", "sheet": "Sales", "column": "Margin",
                 "formula": "=I{row}/G{row}", "start_row": 2},
            ],
            output_path=out,
        )
        assert result["success"] is True
        assert result["operations_applied"][0]["status"] == "ok"

    def test_update_cells(self, sales_path, output_path):
        out = str(Path(output_path) / "test_edit_update.xlsx")
        result = engine.edit_excel(
            file_path=sales_path,
            operations=[
                {"type": "update_cells", "sheet": "Sales",
                 "updates": [{"row": 2, "column": "B", "value": "Updated_Product"}]},
            ],
            output_path=out,
        )
        assert result["success"] is True
        assert result["operations_applied"][0]["cells_updated"] == 1

    def test_nonexistent_input(self, output_path):
        result = engine.edit_excel(
            file_path="no_such_file.xlsx",
            operations=[],
        )
        assert result["success"] is False

    def test_invalid_sheet(self, sales_path, output_path):
        out = str(Path(output_path) / "test_edit_bad_sheet.xlsx")
        result = engine.edit_excel(
            file_path=sales_path,
            operations=[
                {"type": "add_column", "sheet": "NoSheet", "column": "X", "values": [1]},
            ],
            output_path=out,
        )
        assert result["success"] is True  # File still saves
        assert result["operations_applied"][0]["status"] == "error"


# ===================================================================
# 4. TEST PROFILE DATA
# ===================================================================
class TestProfileData:
    def test_basic_profile(self, sales_path):
        result = engine.profile_data(sales_path)
        assert result["success"] is True
        assert result["rows"] == 100
        assert result["columns"] == 9
        assert "Revenue" in result["numerical_columns"]
        assert "Product" in result["categorical_columns"]

    def test_missing_values_detected(self, sales_path):
        result = engine.profile_data(sales_path)
        assert result["total_missing"] > 0
        assert "Unit_Price" in result["missing_values"]

    def test_statistics_present(self, sales_path):
        result = engine.profile_data(sales_path)
        assert len(result["basic_statistics"]) > 0

    def test_categorical_summary(self, sales_path):
        result = engine.profile_data(sales_path)
        assert "Product" in result["categorical_summary"]
        assert result["categorical_summary"]["Product"]["unique_count"] > 0

    def test_nonexistent_file(self):
        result = engine.profile_data("nope.xlsx")
        assert result["success"] is False


# ===================================================================
# 5. TEST ANALYZE DATA
# ===================================================================
class TestAnalyzeData:
    def test_aggregation(self, sales_path):
        result = engine.analyze_data(
            sales_path, "aggregation",
            {"columns": ["Revenue", "Profit"], "functions": ["sum", "mean"]},
        )
        assert result["success"] is True
        assert "Revenue" in result["supporting_data"]
        assert "sum" in result["supporting_data"]["Revenue"]

    def test_group_by(self, sales_path):
        result = engine.analyze_data(
            sales_path, "group_by",
            {"group_by": "Region", "agg_column": "Revenue", "agg_func": "sum"},
        )
        assert result["success"] is True
        assert result["supporting_data"]["row_count"] == 4  # 4 regions

    def test_ranking(self, sales_path):
        result = engine.analyze_data(
            sales_path, "ranking",
            {"column": "Revenue", "top_n": 5},
        )
        assert result["success"] is True
        assert result["supporting_data"]["rows_shown"] == 5

    def test_trend(self, sales_path):
        result = engine.analyze_data(
            sales_path, "trend",
            {"date_column": "Date", "value_column": "Revenue", "period": "M"},
        )
        assert result["success"] is True
        assert "trend" in result["metric"]

    def test_comparison(self, sales_path):
        result = engine.analyze_data(
            sales_path, "comparison",
            {"category_column": "Product", "value_column": "Revenue", "compare_func": "mean"},
        )
        assert result["success"] is True
        assert "Highest" in result["finding"]

    def test_correlation(self, sales_path):
        result = engine.analyze_data(
            sales_path, "correlation",
            {"columns": ["Revenue", "Cost", "Profit", "Quantity"]},
        )
        assert result["success"] is True
        assert "correlation_matrix" in result["supporting_data"]

    def test_distribution(self, sales_path):
        result = engine.analyze_data(
            sales_path, "distribution",
            {"column": "Revenue"},
        )
        assert result["success"] is True
        assert result["supporting_data"]["type"] == "numerical"

    def test_outlier_detection(self, sales_path):
        result = engine.analyze_data(
            sales_path, "outlier_detection",
            {"column": "Revenue"},
        )
        assert result["success"] is True
        assert "outlier_count" in result["supporting_data"]

    def test_missing_analysis(self, sales_path):
        result = engine.analyze_data(sales_path, "missing_analysis")
        assert result["success"] is True
        assert result["supporting_data"]["total_missing"] > 0

    def test_summary_statistics(self, sales_path):
        result = engine.analyze_data(sales_path, "summary_statistics")
        assert result["success"] is True

    def test_custom_query(self, sales_path):
        result = engine.analyze_data(
            sales_path, "custom_query",
            {"query": "Revenue > 5000", "columns": ["Product", "Revenue"]},
        )
        assert result["success"] is True

    def test_unknown_analysis(self, sales_path):
        result = engine.analyze_data(sales_path, "unknown_type")
        assert result["success"] is False


# ===================================================================
# 6. TEST CREATE VISUALIZATION
# ===================================================================
class TestCreateVisualization:
    def test_bar_chart(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "bar",
            {"x": "Product", "y": "Revenue", "title": "Revenue by Product",
             "output_filename": str(Path(output_path) / "test_bar.png")},
        )
        assert result["success"] is True
        assert Path(result["chart_path"]).exists()

    def test_line_chart(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "line",
            {"x": "Date", "y": "Revenue", "title": "Revenue Over Time",
             "output_filename": str(Path(output_path) / "test_line.png")},
        )
        assert result["success"] is True

    def test_scatter_chart(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "scatter",
            {"x": "Revenue", "y": "Profit", "title": "Revenue vs Profit",
             "output_filename": str(Path(output_path) / "test_scatter.png")},
        )
        assert result["success"] is True

    def test_histogram(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "histogram",
            {"x": "Revenue", "nbins": 20, "title": "Revenue Distribution",
             "output_filename": str(Path(output_path) / "test_hist.png")},
        )
        assert result["success"] is True

    def test_box_chart(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "box",
            {"x": "Region", "y": "Revenue", "title": "Revenue by Region",
             "output_filename": str(Path(output_path) / "test_box.png")},
        )
        assert result["success"] is True

    def test_pie_chart(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "pie",
            {"names": "Product", "values": "Revenue", "title": "Revenue Share",
             "output_filename": str(Path(output_path) / "test_pie.png")},
        )
        assert result["success"] is True

    def test_heatmap(self, sales_path, output_path):
        result = engine.create_visualization(
            sales_path, "heatmap",
            {"columns": ["Revenue", "Cost", "Profit", "Quantity"],
             "title": "Correlation Heatmap",
             "output_filename": str(Path(output_path) / "test_heatmap.png")},
        )
        assert result["success"] is True

    def test_unknown_chart_type(self, sales_path):
        result = engine.create_visualization(sales_path, "unknown_type", {})
        assert result["success"] is False


# ===================================================================
# 7. TEST VALIDATE RESULT
# ===================================================================
class TestValidateResult:
    def test_validate_existing_file(self, sales_path):
        result = engine.validate_result({"file_path": sales_path})
        assert result["valid"] is True
        assert len(result["errors"]) == 0

    def test_validate_expected_sheets(self, sales_path):
        result = engine.validate_result({
            "file_path": sales_path,
            "expected_sheets": ["Sales"],
        })
        assert result["valid"] is True

    def test_validate_missing_sheet(self, sales_path):
        result = engine.validate_result({
            "file_path": sales_path,
            "expected_sheets": ["NonExistent"],
        })
        assert result["valid"] is False
        assert any("NonExistent" in e for e in result["errors"])

    def test_validate_expected_columns(self, sales_path):
        result = engine.validate_result({
            "file_path": sales_path,
            "expected_columns": {"Sales": ["Revenue", "Product"]},
        })
        assert result["valid"] is True

    def test_validate_missing_column(self, sales_path):
        result = engine.validate_result({
            "file_path": sales_path,
            "expected_columns": {"Sales": ["NonExistentCol"]},
        })
        assert result["valid"] is False

    def test_validate_nonexistent_file(self):
        result = engine.validate_result({"file_path": "no_file.xlsx"})
        assert result["valid"] is False

    def test_validate_chart_file(self, output_path):
        # Create a dummy chart file
        chart_path = Path(output_path) / "dummy_chart.png"
        chart_path.write_bytes(b"fake png data")

        result = engine.validate_result({"chart_path": str(chart_path)})
        assert result["valid"] is True

    def test_validate_missing_chart(self):
        result = engine.validate_result({"chart_path": "no_chart.png"})
        assert result["valid"] is False
