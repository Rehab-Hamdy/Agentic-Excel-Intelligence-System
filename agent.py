"""
Agent — LangGraph-based agentic workflow with MCP tool integration.

This module implements a ReAct-style agent using LangGraph's StateGraph.
The Agent connects to the MCP Server via stdio, discovers available tools,
and dynamically decides which tools to call based on the user's request.

NO fixed workflow. The LLM reasoning drives tool selection and iteration.
"""

import os
import sys
import asyncio
from pathlib import Path
from typing import Literal

from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage, SystemMessage
from langgraph.graph import StateGraph, MessagesState, START, END
from langgraph.prebuilt import ToolNode
from langchain_mcp_adapters.client import MultiServerMCPClient

# Load environment variables
load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
LLM_PROVIDER = os.getenv("LLM_PROVIDER", "").lower()
MODEL_NAME = os.getenv("MODEL_NAME", "")

# Path to the MCP server script
MCP_SERVER_PATH = str(Path(__file__).resolve().parent / "mcp_server.py")

# Python executable (use venv if available)
VENV_PYTHON = str(Path(__file__).resolve().parent / "venv" / "Scripts" / "python.exe")
if not Path(VENV_PYTHON).exists():
    VENV_PYTHON = sys.executable

# ---------------------------------------------------------------------------
# System Prompt
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """You are an Excel Intelligence Assistant. You help users interact with Excel files using natural language.

You have access to the following tools through the MCP server:

**Excel Operations:**
- `read_excel`: Read an Excel file to inspect its sheets, columns, and data.
- `create_excel`: Create a new Excel workbook with structured data.
- `edit_excel`: Edit an existing workbook (add columns, formulas, update cells).

**Data Analysis:**
- `profile_data`: Get a comprehensive profile of the data (types, missing values, statistics).
- `analyze_data`: Perform specific analyses (aggregation, group_by, ranking, trend, correlation, outlier_detection, etc.).

**Visualization:**
- `create_visualization`: Generate charts (line, bar, scatter, histogram, box, pie, heatmap).

**Validation:**
- `validate_result`: Verify that output files and results are correct.

## YOUR RULES:

1. **Never invent numerical results.** Always use tools to get actual data.
2. **Inspect data before assuming.** Always read/profile the file first to understand what columns exist.
3. **Use the right tool for the job.** Don't over-call tools or under-call them.
4. **Validate important modifications.** After creating or editing files, validate the results.
5. **Don't claim success if validation fails.** Report errors clearly.
6. **Recover from errors when possible.** If a tool fails, try to understand why and retry with corrected parameters.
7. **Be specific in your responses.** Include actual numbers, file paths, and findings.
8. **Choose appropriate visualizations:**
   - Time-series trends → line chart
   - Category comparisons → bar chart
   - Two numerical variables → scatter plot
   - Distribution → histogram
   - Outliers/spread by group → box plot
   - Proportions → pie chart
   - Correlation matrix → heatmap

## WORKFLOW:
- You decide the sequence of tool calls based on the user's request.
- You may call multiple tools sequentially.
- Use the result of one tool to determine the next action.
- Stop when you have enough information to answer the user's question.

## FILE PATHS:
- Input files are in the `data/input/` directory. Use just the filename (e.g., "sales.xlsx").
- Output files go to `data/output/`. They will be saved there automatically.
- Never overwrite the original input file unless explicitly asked.

## RESPONSE FORMAT:
When presenting results:
- Start with a brief summary of what you found.
- Present key insights with actual numbers.
- Mention any generated files (Excel, charts) with their paths.
- Flag any warnings or issues.
"""


# ---------------------------------------------------------------------------
# OpenRouter Free Models Failover Pool
# ---------------------------------------------------------------------------
class OpenRouterFailoverPool:
    """
    Manages automatic failover across OpenRouter free models when a model
    exhausts its tokens, hits rate limits (429/402), or encounters provider errors.
    """

    DEFAULT_FREE_MODELS = [
        "meta-llama/llama-3.3-70b-instruct:free",
        "mistralai/mistral-small-24b-instruct-2501:free",
        "google/gemini-2.0-flash-exp:free",
        "qwen/qwen-2.5-72b-instruct:free",
        "meta-llama/llama-3.2-3b-instruct:free",
        "deepseek/deepseek-r1:free",
    ]

    def __init__(self, api_key: str, preferred_models: list[str] | None, tools: list):
        self.api_key = api_key
        self.tools = tools

        # Build model sequence: user preferences first, followed by default free models
        models = []
        if preferred_models:
            for m in preferred_models:
                m = m.strip()
                if m and m not in models:
                    models.append(m)

        for m in self.DEFAULT_FREE_MODELS:
            if m not in models:
                models.append(m)

        self.models = models
        self.current_idx = 0
        self.bound_llms = {}

    def get_current_model_name(self) -> str:
        return self.models[self.current_idx]

    def get_current_llm(self):
        model_name = self.models[self.current_idx]
        if model_name not in self.bound_llms:
            from langchain_openai import ChatOpenAI
            base_llm = ChatOpenAI(
                model=model_name,
                api_key=self.api_key,
                base_url="https://openrouter.ai/api/v1",
                temperature=0,
                default_headers={
                    "HTTP-Referer": "https://github.com/excel-agent",
                    "X-Title": "Excel Agent",
                },
            )
            self.bound_llms[model_name] = base_llm.bind_tools(self.tools)
        return self.bound_llms[model_name]

    def switch_to_next(self, reason: str = ""):
        old_model = self.models[self.current_idx]
        self.current_idx = (self.current_idx + 1) % len(self.models)
        new_model = self.models[self.current_idx]
        print(f"\n  [FAILOVER] Model '{old_model}' finished tokens / rate-limited ({reason}).")
        print(f"  [FAILOVER] Automatically switched to next free model: '{new_model}'\n")

    async def ainvoke(self, messages):
        attempts = 0
        max_attempts = len(self.models)
        last_exception = None

        while attempts < max_attempts:
            current_model = self.get_current_model_name()
            llm = self.get_current_llm()
            try:
                response = await llm.ainvoke(messages)

                # Check if the text output explicitly reports a quota or rate error
                content_str = str(getattr(response, "content", "")).lower()
                if any(kw in content_str for kw in ["rate limit", "credits exhausted", "quota exceeded", "no available providers"]):
                    self.switch_to_next("limit detected in response")
                    attempts += 1
                    continue

                return response

            except Exception as e:
                err_msg = str(e).lower()
                self.switch_to_next(f"{type(e).__name__}: {str(e)[:100]}")
                last_exception = e
                attempts += 1

        raise last_exception or RuntimeError("All free models in the OpenRouter pool exhausted their quotas.")


# ---------------------------------------------------------------------------
# Gemini API Key Failover Pool
# ---------------------------------------------------------------------------
class GeminiFailoverPool:
    """
    Manages automatic failover across multiple Gemini API keys when a key
    exhausts its quota or hits rate limits.
    """

    def __init__(self, model_name: str, keys: list[str], tools: list):
        self.model_name = model_name
        self.keys = keys
        self.tools = tools
        self.current_idx = 0
        self.bound_llms = {}

    def get_current_key(self) -> str:
        return self.keys[self.current_idx]

    def get_current_llm(self):
        key = self.get_current_key()
        if key not in self.bound_llms:
            from langchain_google_genai import ChatGoogleGenerativeAI
            base_llm = ChatGoogleGenerativeAI(
                model=self.model_name,
                google_api_key=key,
                temperature=0,
            )
            self.bound_llms[key] = base_llm.bind_tools(self.tools)
        return self.bound_llms[key]

    def switch_to_next(self, reason: str = ""):
        old_key = self.keys[self.current_idx]
        self.current_idx = (self.current_idx + 1) % len(self.keys)
        new_key = self.keys[self.current_idx]
        # Mask the keys for logging
        old_masked = f"{old_key[:4]}...{old_key[-4:]}" if len(old_key) > 8 else "***"
        new_masked = f"{new_key[:4]}...{new_key[-4:]}" if len(new_key) > 8 else "***"
        print(f"\n  [FAILOVER] Gemini key '{old_masked}' failed ({reason}).")
        print(f"  [FAILOVER] Automatically switched to next Gemini key: '{new_masked}'\n")

    async def ainvoke(self, messages):
        attempts = 0
        max_attempts = len(self.keys)
        last_exception = None

        while attempts < max_attempts:
            llm = self.get_current_llm()
            try:
                response = await llm.ainvoke(messages)
                return response
            except Exception as e:
                self.switch_to_next(f"{type(e).__name__}: {str(e)[:100]}")
                last_exception = e
                attempts += 1

        raise last_exception or RuntimeError("All Gemini API keys in the failover pool have been exhausted.")


# ---------------------------------------------------------------------------
# Agent Builder
# ---------------------------------------------------------------------------
async def create_agent():
    """
    Create and return the LangGraph agent connected to the MCP server.

    Returns a tuple of (graph, cleanup_function).
    The cleanup function should be called when done to close the MCP connection.
    """
    # Connect to the MCP server using MultiServerMCPClient
    client = MultiServerMCPClient({
        "excel": {
            "command": VENV_PYTHON,
            "args": [MCP_SERVER_PATH],
            "transport": "stdio",
            "env": {
                **os.environ,
                "PYTHONPATH": str(Path(__file__).resolve().parent),
            },
        }
    })

    # Load MCP tools as LangChain tools
    tools = await client.get_tools()

    print(f"  [OK] Connected to MCP server with {len(tools)} tools:")
    for tool in tools:
        print(f"    - {tool.name}")

    # Read Gemini keys from file if it exists
    gemini_keys_path = Path(__file__).resolve().parent / "gemini_keys.txt"
    gemini_keys = []
    if gemini_keys_path.exists():
        with open(gemini_keys_path, "r", encoding="utf-8") as f:
            gemini_keys = [line.strip() for line in f if line.strip()]
            
    if GOOGLE_API_KEY and GOOGLE_API_KEY not in gemini_keys:
        gemini_keys.insert(0, GOOGLE_API_KEY)

    # Initialize the LLM or Failover Pool
    if OPENROUTER_API_KEY and (LLM_PROVIDER in ("openrouter", "") or (not GROQ_API_KEY and not gemini_keys)):
        pref_models = [m.strip() for m in MODEL_NAME.split(",") if m.strip()] if MODEL_NAME else []
        pool = OpenRouterFailoverPool(
            api_key=OPENROUTER_API_KEY,
            preferred_models=pref_models,
            tools=tools,
        )
        print(f"  [LLM] Using OpenRouter Failover Pool (Primary: {pool.get_current_model_name()})")
        print(f"        Available free fallbacks: {len(pool.models)} models")
        llm_invoker = pool

    elif GROQ_API_KEY and (LLM_PROVIDER in ("groq", "") or not gemini_keys):
        from langchain_groq import ChatGroq
        model = MODEL_NAME or "openai/gpt-oss-120b"
        print(f"  [LLM] Using Groq with model: {model}")
        llm = ChatGroq(
            model=model,
            groq_api_key=GROQ_API_KEY,
            temperature=0,
        )
        llm_invoker = llm.bind_tools(tools)

    elif gemini_keys:
        model = MODEL_NAME or "gemini-3.8-flash"
        if len(gemini_keys) > 1:
            print(f"  [LLM] Using Gemini Failover Pool with {len(gemini_keys)} keys (model: {model})")
            llm_invoker = GeminiFailoverPool(model_name=model, keys=gemini_keys, tools=tools)
        else:
            print(f"  [LLM] Using Google Gemini with model: {model}")
            llm = ChatGoogleGenerativeAI(
                model=model,
                google_api_key=gemini_keys[0],
                temperature=0,
            )
            llm_invoker = llm.bind_tools(tools)

    else:
        raise ValueError("No API key configured in .env or gemini_keys.txt.")

    # ----- Define the StateGraph -----

    async def agent_node(state: MessagesState):
        """The agent node -- invokes the LLM with conversation + tools."""
        messages = state["messages"]

        # Ensure system prompt is at the beginning
        if not messages or not isinstance(messages[0], SystemMessage):
            messages = [SystemMessage(content=SYSTEM_PROMPT)] + messages

        response = await llm_invoker.ainvoke(messages)
        return {"messages": [response]}

    def should_continue(state: MessagesState) -> Literal["tools", "__end__"]:
        """Decide whether to call tools or end the conversation."""
        last_message = state["messages"][-1]
        # If the LLM made tool calls, route to the tool node
        if hasattr(last_message, "tool_calls") and last_message.tool_calls:
            return "tools"
        # Otherwise, we're done
        return "__end__"

    # Build the graph
    graph_builder = StateGraph(MessagesState)

    # Add nodes
    graph_builder.add_node("agent", agent_node)
    graph_builder.add_node("tools", ToolNode(tools))

    # Add edges
    graph_builder.add_edge(START, "agent")
    graph_builder.add_conditional_edges("agent", should_continue)
    graph_builder.add_edge("tools", "agent")

    # Compile the graph
    graph = graph_builder.compile()

    # Return graph and a cleanup context
    async def cleanup():
        """Clean up MCP connection."""
        pass

    return graph, cleanup


async def run_agent_query(graph, query: str, conversation_history: list = None) -> str:
    """
    Run a single query through the agent.

    Args:
        graph: The compiled LangGraph agent.
        query: The user's natural language query.
        conversation_history: Optional list of previous messages for context.

    Returns:
        The agent's final text response.
    """
    messages = conversation_history or []
    messages.append(HumanMessage(content=query))

    # Add system prompt if not present
    if not messages or not isinstance(messages[0], SystemMessage):
        messages = [SystemMessage(content=SYSTEM_PROMPT)] + messages

    result = await graph.ainvoke({"messages": messages})

    # Extract the final response text cleanly
    final_message = result["messages"][-1]
    raw_content = final_message.content
    if isinstance(raw_content, list):
        text_parts = []
        for part in raw_content:
            if isinstance(part, dict) and "text" in part:
                text_parts.append(part["text"])
            elif isinstance(part, str):
                text_parts.append(part)
        content_text = "\n".join(text_parts) if text_parts else str(raw_content)
    else:
        content_text = str(raw_content)

    return content_text, result["messages"]
