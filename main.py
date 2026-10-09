"""
Main Entry Point — Agentic Excel Intelligence System.

Initializes configuration, connects to the MCP server, creates the agent,
and runs an interactive chat loop.
"""

import asyncio
import sys
import os
from pathlib import Path

from dotenv import load_dotenv

# Load environment variables first
load_dotenv(Path(__file__).resolve().parent / ".env")


def print_banner():
    """Print the application banner."""
    print("=" * 60)
    print("   [Excel Agent] Agentic Excel Intelligence System")
    print("=" * 60)
    print()
    print("  Ask me anything about your Excel files!")
    print("  Place Excel files in: data/input/")
    print("  Generated files go to: data/output/")
    print()
    print("  Commands:")
    print("    exit / quit  -- Exit the application")
    print("    clear        -- Clear conversation history")
    print("    files        -- List available input files")
    print()
    print("-" * 60)


def list_input_files():
    """List available Excel files in data/input/."""
    input_dir = Path(__file__).resolve().parent / "data" / "input"
    files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls"))
    if files:
        print("\n  Available input files:")
        for f in files:
            size_kb = f.stat().st_size / 1024
            print(f"    * {f.name} ({size_kb:.1f} KB)")
    else:
        print("\n  No Excel files found in data/input/")
        print("     Place your .xlsx files there to get started.")
    print()


async def main():
    """Main application loop."""
    # Validate API key (OpenRouter, Groq, or Google)
    openrouter_key = os.getenv("OPENROUTER_API_KEY", "")
    groq_key = os.getenv("GROQ_API_KEY", "")
    google_key = os.getenv("GOOGLE_API_KEY", "")
    if not openrouter_key and not groq_key and (not google_key or google_key == "your-api-key-here"):
        print("[ERROR] Please set an API key in the .env file:")
        print("   OpenRouter: https://openrouter.ai/keys  (for free models)")
        print("   Groq:       https://console.groq.com/keys")
        print("   Google:     https://aistudio.google.com/apikey")
        sys.exit(1)

    print_banner()

    print("  Initializing agent and MCP server...")

    try:
        from agent import create_agent, run_agent_query

        graph, cleanup = await create_agent()
        print("  [OK] Agent ready!\n")
    except Exception as e:
        print(f"\n  [ERROR] Failed to initialize: {e}")
        print(f"     Make sure all dependencies are installed: pip install -r requirements.txt")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    conversation_history = []

    try:
        while True:
            try:
                user_input = input("\n  You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\n\n  Goodbye!")
                break

            if not user_input:
                continue

            # Handle commands
            if user_input.lower() in ("exit", "quit", "q"):
                print("\n  Goodbye!")
                break
            elif user_input.lower() == "clear":
                conversation_history = []
                print("  [INFO] Conversation cleared.")
                continue
            elif user_input.lower() == "files":
                list_input_files()
                continue

            # Run the agent
            print("\n  Thinking...\n")

            try:
                response, updated_history = await run_agent_query(
                    graph, user_input, conversation_history
                )
                conversation_history = updated_history
                print(f"  Agent: {response}")
            except Exception as e:
                print(f"\n  [ERROR] {e}")
                import traceback
                traceback.print_exc()

    finally:
        # Clean up MCP connection
        print("\n  Shutting down...")
        try:
            await cleanup()
        except Exception:
            pass
        print("  [OK] Done.")


if __name__ == "__main__":
    asyncio.run(main())
