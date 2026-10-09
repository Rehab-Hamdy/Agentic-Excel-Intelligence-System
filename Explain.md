# Excel Agent File Guide

This project is an Excel assistant that lets a user work with Excel files using natural language.

The main request flow is:

```text
User request
    -> agent.py
    -> mcp_server.py
    -> python_engine.py
    -> Excel file, analysis result, or chart
    -> response back to the user
```

## 1. `agent.py`

### Main responsibility

`agent.py` is the AI orchestration layer. It connects Google Gemini to the available Excel tools and uses LangGraph to decide which tools to call and in what order.

It does not directly read or edit Excel files. Instead, it asks the MCP server to perform those operations.

### Important configuration

- Loads environment variables from `.env` using `load_dotenv()`.
- Reads `GOOGLE_API_KEY` for Gemini authentication.
- Reads `MODEL_NAME`, defaulting to `gemini-2.5-flash`.
- Finds `mcp_server.py` next to the agent file.
- Selects the project's virtual-environment Python executable when it exists.

### System prompt

`SYSTEM_PROMPT` tells the language model:

- What Excel tools are available.
- That it must inspect data before making assumptions.
- That it must use real tool results instead of inventing numbers.
- That it should validate important changes.
- Which chart type is appropriate for different tasks.
- Where input and output files are located.

### `create_agent()`

This asynchronous function builds the complete agent:

1. Starts `mcp_server.py` as a subprocess using stdio communication.
2. Opens an MCP client session.
3. Initializes the session.
4. Discovers the tools exposed by the MCP server.
5. Creates a Google Gemini chat model.
6. Binds the discovered MCP tools to the model.
7. Builds a LangGraph `StateGraph` with two nodes:
   - `agent`: sends the conversation to Gemini.
   - `tools`: executes requested MCP tools.
8. Repeats the agent/tool cycle until Gemini has no more tool calls.
9. Returns the compiled graph and an asynchronous cleanup function.

The agent is therefore dynamic. There is no hard-coded sequence such as "always read, then analyze". Gemini chooses the next step based on the user's request and the previous tool results.

### `run_agent_query(graph, query, conversation_history=None)`

Runs one natural-language question through the compiled graph.

It:

- Adds the user's query as a `HumanMessage`.
- Adds the system prompt if it is not already present.
- Calls the graph asynchronously.
- Returns the final text response and the full message history.

The message history can be supplied again to support follow-up questions in the same conversation.

### Cleanup

The cleanup function returned by `create_agent()` closes the MCP session. It should be called when the application exits.

---

## 2. `mcp_server.py`

### Main responsibility

`mcp_server.py` is the integration layer between the AI agent and the Python computation code.

It exposes normal Python functions as MCP tools. Each tool receives simple arguments, converts JSON strings into Python objects when needed, calls the matching function in `python_engine.py`, and converts the result back to a JSON string.

This file should stay thin. The actual Excel and analysis logic belongs in `python_engine.py`.

### Server setup

```python
mcp = MCPServer("ExcelIntelligenceServer")
```

The `@mcp.tool()` decorator registers a function as an MCP tool that the agent can discover and call.

### Excel tools

- `read_excel(file_path, sheet_name=None)`
  - Reads workbook metadata and rows.
  - Returns sheet names, the selected sheet, column names, data types, and data.

- `create_excel(file_path, sheets_data, formatting=None)`
  - Parses the supplied JSON data and formatting options.
  - Creates a new workbook through the Python engine.

- `edit_excel(file_path, operations, output_path=None)`
  - Parses a JSON list of edit operations.
  - Supports adding columns, adding formulas, updating cells, deleting columns, and renaming columns.

### Analysis tools

- `profile_data(file_path, sheet_name=None)`
  - Produces a complete data profile, including types, missing values, duplicates, statistics, and column classifications.

- `analyze_data(file_path, analysis_type, params=None, sheet_name=None)`
  - Runs one selected analysis mode.
  - Supported modes include aggregation, grouping, ranking, trends, comparison, correlation, distribution, outlier detection, missing-value analysis, summary statistics, and custom queries.

### Visualization tool

- `create_visualization(file_path, chart_type, params, sheet_name=None)`
  - Creates line, bar, scatter, histogram, box, pie, or heatmap charts.
  - Saves the chart under `data/output/`.

### Validation tool

- `validate_result(validation_config)`
  - Checks whether an output file exists and can be opened.
  - Can check expected sheets, expected columns, non-empty sheets, and chart files.

### Entry point

When run directly, the file starts the MCP server with stdio transport:

```python
python mcp_server.py
```

Normally, `agent.py` starts it automatically, so it is not necessary to run it manually.

---

## 3. `python_engine.py`

### Main responsibility

`python_engine.py` contains the actual Excel, data-processing, analysis, charting, and validation logic.

It is independent of the language model and MCP protocol. Functions receive structured Python arguments and return dictionaries that can be serialized to JSON.

### Libraries used

- `pandas`: reads Excel data and performs data manipulation.
- `openpyxl`: creates, edits, and validates Excel workbooks.
- `numpy`: supports numerical values and serialization checks.
- `scipy.stats`: performs the Shapiro normality test.
- `plotly`: creates charts.
- `matplotlib`: configured with a non-interactive backend for server-side use.

### File locations

- Relative input paths are resolved under `data/input/`.
- Relative output paths are resolved under `data/output/`.
- The two directories are created automatically when the module is imported.
- Some operations can also read a file from `data/output/` if it is not found in `data/input/`.

### Helper functions

- `_resolve_input_path()` converts a relative input filename into a path under `data/input/`.
- `_resolve_output_path()` converts a relative output filename into a path under `data/output/`.
- `_df_to_serializable()` converts a Pandas DataFrame into JSON-compatible data and limits the displayed rows to 500.

### Excel operations

#### `read_excel()`

Opens a workbook, lists its sheets, selects the requested sheet or the first sheet, reads it with Pandas, and returns:

- File path.
- Sheet names.
- Active sheet.
- Row and column counts.
- Column names and data types.
- Up to 500 data rows.

#### `create_excel()`

Creates a workbook from a mapping of sheet names to lists of row dictionaries.

It also:

- Removes the default empty worksheet.
- Writes headers and data.
- Applies header color and bold formatting.
- Auto-sizes columns.
- Saves the workbook under `data/output/`.

#### `edit_excel()`

Loads an existing workbook and applies a list of operations. It saves the result to a new output file by default.

Supported operation types:

- `add_column`: adds a column and writes supplied values.
- `add_formula`: adds a column and fills it with row-based Excel formulas.
- `update_cells`: updates individual cells by letter or numeric column index.
- `delete_column`: removes a column by its header name.
- `rename_column`: changes a column header.

Internal helpers implement these individual operations:

- `_find_col_index()`
- `_op_add_column()`
- `_op_add_formula()`
- `_op_update_cells()`
- `_op_delete_column()`
- `_op_rename_column()`

### Data profiling

#### `profile_data()`

Reads a worksheet and reports:

- Number of rows and columns.
- Column names and Pandas data types.
- Missing values by column.
- Total missing values.
- Duplicate row count.
- Descriptive statistics for numerical columns.
- Common values and unique counts for categorical columns.
- Numerical, categorical, and date column classifications.

### Data analysis

#### `analyze_data()`

Loads the selected worksheet and dispatches the request to a handler based on `analysis_type`.

Available handlers are:

- `_analyze_aggregation()`: calculates functions such as sum, mean, maximum, minimum, and count.
- `_analyze_group_by()`: groups by one column and aggregates another column.
- `_analyze_ranking()`: sorts rows and returns the top or bottom rows.
- `_analyze_trend()`: converts a date column, resamples data by month, quarter, or year, and calculates overall change.
- `_analyze_comparison()`: compares category values using an aggregation function such as mean.
- `_analyze_correlation()`: creates a numerical correlation matrix and lists strong correlations.
- `_analyze_distribution()`: describes categorical or numerical distributions, including skewness and kurtosis.
- `_analyze_outlier_detection()`: finds numerical outliers with the IQR method.
- `_analyze_missing()`: counts missing cells and complete rows.
- `_analyze_summary_statistics()`: returns extended descriptive statistics for numerical columns.
- `_analyze_custom_query()`: filters rows using Pandas query syntax.

Each handler returns a structured result containing a metric, a human-readable finding, and supporting data where appropriate.

### Visualization

#### `create_visualization()`

Reads Excel data, selects a chart handler, and saves the result to `data/output/`.

Supported chart handlers:

- `_chart_line()`: line chart.
- `_chart_bar()`: bar chart.
- `_chart_scatter()`: scatter plot, optionally with a trendline.
- `_chart_histogram()`: distribution histogram.
- `_chart_box()`: box plot.
- `_chart_pie()`: proportions by names and values.
- `_chart_heatmap()`: correlation heatmap for numerical columns.

Charts are saved as PNG files by default. If the requested output filename ends in `.html`, Plotly saves an interactive HTML chart instead.

### Validation

#### `validate_result()`

Checks the result of an Excel or chart operation.

For Excel files it can verify:

- The file exists.
- The workbook opens successfully.
- Expected sheet names exist.
- Expected column headers exist.
- Sheets are not empty when `non_empty` is enabled.

For charts it checks that the chart file exists and is not empty.

It returns:

```python
{
    "valid": True or False,
    "warnings": [...],
    "errors": [...]
}
```

---

## How the three files work together

Example request: "Find the top five products by revenue in sales.xlsx."

1. `agent.py` receives the natural-language request and sends it to Gemini.
2. Gemini decides that the `analyze_data` tool is appropriate.
3. `mcp_server.py` receives the tool call and parses its JSON parameters.
4. `mcp_server.py` calls `python_engine.analyze_data()`.
5. `python_engine.py` selects `_analyze_ranking()` and calculates the result with Pandas.
6. The result travels back through `mcp_server.py` as JSON.
7. `agent.py` gives the tool result back to Gemini.
8. Gemini writes a natural-language answer for the user.

In short:

- `agent.py` decides what should happen.
- `mcp_server.py` exposes the available actions.
- `python_engine.py` performs the actual work.
