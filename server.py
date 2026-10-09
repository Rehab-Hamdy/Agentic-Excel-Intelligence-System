import os
from pathlib import Path
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Dict, Any, Optional
import json

from agent import create_agent, run_agent_query

# Configure directories
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
STATIC_DIR = BASE_DIR / "static"

# Ensure directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
STATIC_DIR.mkdir(parents=True, exist_ok=True)

# Global variables for agent
app_state = {
    "graph": None,
    "cleanup": None,
    "history": [] # Store message history in memory for simplicity
}

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("[Server] Initializing Agent...")
    try:
        graph, cleanup = await create_agent()
        app_state["graph"] = graph
        app_state["cleanup"] = cleanup
        print("[Server] Agent initialized successfully.")
    except Exception as e:
        print(f"[Server] Error initializing agent: {e}")
    yield
    print("[Server] Shutting down agent...")
    if app_state["cleanup"]:
        await app_state["cleanup"]()

app = FastAPI(lifespan=lifespan, title="Excel Agent API")

# Serve static files
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Expose output files so they can be loaded in the browser (e.g., images)
app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")


class ChatRequest(BaseModel):
    message: str

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(STATIC_DIR / "index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.get("/api/files")
async def list_files():
    """List available input and output files."""
    try:
        input_files = [f.name for f in INPUT_DIR.iterdir() if f.is_file() and not f.name.startswith(".")]
        output_files = [f.name for f in OUTPUT_DIR.iterdir() if f.is_file() and not f.name.startswith(".")]
        return {"input_files": input_files, "output_files": output_files}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.post("/api/chat")
async def chat_endpoint(request: ChatRequest):
    """Handle chat messages with the agent."""
    if not app_state["graph"]:
        return JSONResponse(
            status_code=500, 
            content={"error": "Agent is not initialized. Check server logs for details."}
        )
    
    user_message = request.message
    if not user_message.strip():
        return {"response": "Please enter a message."}
        
    if user_message.lower().strip() == "clear":
        app_state["history"] = []
        return {"response": "Conversation history cleared."}

    try:
        print(f"[Server] Routing message to agent: {user_message}")
        response_text, new_history = await run_agent_query(
            app_state["graph"], 
            user_message, 
            app_state["history"]
        )
        # Update history (excluding system prompt to save memory, or just keep as is)
        app_state["history"] = new_history
        
        return {"response": response_text}
    except Exception as e:
        print(f"[Server] Error during agent query: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=8000, reload=True)
