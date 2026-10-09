"""
Tests for mcp_server.py — verify each MCP tool independently.

These tests import the python_engine directly and verify the MCP
tool wrappers correctly delegate and return JSON strings.
"""
import sys
import json
from pathlib import Path

import pytest

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import mcp_server


# ===================================================================
# Helpers
# ===================================================================
@pytest.fixture
def sales_path(sample_excel_path):
    """Return absolute path string to the sample sales.xlsx."""
    return str(sample_excel_path)


@pytest.fixture
def output_path(output_dir):
    """Return output directory string."""
    return str(output_dir)


def parse_result(json_str: str) -> dict:
    """Parse a JSON string result from an MCP tool."""
    return json.loads(json_str)


# ===================================================================
# TEST MCP TOOL: read_excel
# ===================================================================
class TestMCPReadExcel:
    def test_read_returns_json(self, sales_path):
        result_str = mcp_server.read_excel(sales_path)
        result = parse_result(result_str)
        assert result["success"] is True
        assert isinstance(result["sheet_names"], list)

    def test_read_with_sheet(self, sales_path):
        result = parse_result(mcp_server.read_excel(sales_path, "Sales"))
        assert result["active_sheet"] == "Sales"

    def test_read_nonexistent(self):
        result = parse_result(mcp_server.read_excel("fake.xlsx"))
        assert result["success"] is False


# ===================================================================
# TEST MCP TOOL: create_excel
# ===================================================================
class TestMCPCreateExcel:
    def test_create_returns_json(self, output_path):
        path = str(Path(output_path) / "mcp_test_create.xlsx")
        sheets = json.dumps({"Test": [{"A": 1, "B": 2}]})
        result = parse_result(mcp_server.create_excel(path, sheets))
        assert result["success"] is True
        assert Path(result["file_path"]).exists()

    def test_create_with_formatting(self, output_path):
        path = str(Path(output_path) / "mcp_test_fmt.xlsx")
        sheets = json.dumps({"Data": [{"X": 10}]})
        fmt = json.dumps({"header_color": "FF0000", "bold_header": True})
        result = parse_result(mcp_server.create_excel(path, sheets, fmt))
        assert result["success"] is True


# ===================================================================
# TEST MCP TOOL: edit_excel
# ===================================================================
class TestMCPEditExcel:
    def test_edit_add_column(self, sales_path, output_path):
        out = str(Path(output_path) / "mcp_test_edit.xlsx")
        ops = json.dumps([
            {"type": "add_column", "sheet": "Sales", "column": "MCP_Test",
             "values": list(range(100))}
        ])
        result = parse_result(mcp_server.edit_excel(sales_path, ops, out))
        assert result["success"] is True
        assert result["operations_applied"][0]["status"] == "ok"


# ===================================================================
# TEST MCP TOOL: profile_data
# ===================================================================
class TestMCPProfileData:
    def test_profile_returns_json(self, sales_path):
        result = parse_result(mcp_server.profile_data(sales_path))
        assert result["success"] is True
        assert result["rows"] == 100
        assert "numerical_columns" in result


# ===================================================================
# TEST MCP TOOL: analyze_data
# ===================================================================
class TestMCPAnalyzeData:
    def test_aggregation(self, sales_path):
        params = json.dumps({"columns": ["Revenue"], "functions": ["sum", "mean"]})
        result = parse_result(mcp_server.analyze_data(sales_path, "aggregation", params))
        assert result["success"] is True
        assert "Revenue" in result["supporting_data"]

    def test_group_by(self, sales_path):
        params = json.dumps({"group_by": "Region", "agg_column": "Revenue", "agg_func": "sum"})
        result = parse_result(mcp_server.analyze_data(sales_path, "group_by", params))
        assert result["success"] is True

    def test_correlation(self, sales_path):
        params = json.dumps({"columns": ["Revenue", "Cost", "Profit"]})
        result = parse_result(mcp_server.analyze_data(sales_path, "correlation", params))
        assert result["success"] is True


# ===================================================================
# TEST MCP TOOL: create_visualization
# ===================================================================
class TestMCPCreateVisualization:
    def test_bar_chart(self, sales_path, output_path):
        params = json.dumps({
            "x": "Product", "y": "Revenue",
            "title": "MCP Test Bar",
            "output_filename": str(Path(output_path) / "mcp_test_bar.png"),
        })
        result = parse_result(mcp_server.create_visualization(sales_path, "bar", params))
        assert result["success"] is True
        assert Path(result["chart_path"]).exists()


# ===================================================================
# TEST MCP TOOL: validate_result
# ===================================================================
class TestMCPValidateResult:
    def test_validate_existing(self, sales_path):
        config = json.dumps({"file_path": sales_path, "expected_sheets": ["Sales"]})
        result = parse_result(mcp_server.validate_result(config))
        assert result["valid"] is True

    def test_validate_missing(self):
        config = json.dumps({"file_path": "nonexistent.xlsx"})
        result = parse_result(mcp_server.validate_result(config))
        assert result["valid"] is False
