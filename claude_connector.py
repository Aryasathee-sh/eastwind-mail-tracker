"""
claude_connector.py
-------------------
Connects any Excel workbook to Claude for intelligent Q&A.

Usage:
    python claude_connector.py

Requirements:
    pip install anthropic

Set your API key:
    Windows:  set ANTHROPIC_API_KEY=sk-ant-...
    or create a .env file with: ANTHROPIC_API_KEY=sk-ant-...
"""

import os
import sys
import json
from excel_loader import load_workbook, get_summary

# ── Try to load .env if present ───────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv is optional

# ── Anthropic client ──────────────────────────────────────────────────────────
try:
    import anthropic
except ImportError:
    raise SystemExit(
        "anthropic package not found.\n"
        "Install it with:  pip install anthropic"
    )

# ── Config ────────────────────────────────────────────────────────────────────
MODEL       = "claude-sonnet-4-5"
MAX_TOKENS  = 4096
RECENT_ROWS = 50  # number of recent rows sent to Claude per session


def build_system_prompt(data: dict) -> str:
    """Build the system prompt with workbook context."""
    summary = get_summary(data)

    # Pull the most recent RECENT_ROWS rows from Email Tracker
    rows = data.get("Email Tracker", [])[-RECENT_ROWS:]
    rows_json = json.dumps(rows, ensure_ascii=False, indent=2)

    # Pull Settings lookup lists
    settings = data.get("Settings", [])
    settings_json = json.dumps(settings[:30], ensure_ascii=False, indent=2)

    return f"""You are an intelligent assistant for the Eastwind Proposals team.
You have access to their Mail Tracker Excel workbook with the following sheets:
  • Email Tracker  – the main log of all proposal emails
  • Dashboard      – aggregated KPIs
  • Settings       – dropdown lists (statuses, priorities, assignees, categories)
  • MT_Users       – team members
  • MT_Lists       – reference lists
  • MT_Settings    – system settings
  • MT_Audit       – change log

=== WORKBOOK SUMMARY ===
{summary}

=== SETTINGS / LOOKUP VALUES ===
{settings_json}

=== MOST RECENT {RECENT_ROWS} EMAIL TRACKER ROWS ===
{rows_json}

Answer questions about the data above accurately and concisely.
When asked to filter or analyse, use the data you have been given.
If a question requires data not in the snippet, say so and suggest the user narrow the filter.
"""


def chat(data: dict):
    """Interactive chat loop with Claude."""
    client  = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))
    system  = build_system_prompt(data)
    history = []

    print("\n🤖  Eastwind Mail Tracker — Claude Assistant")
    print("    Type your question and press Enter. Type 'exit' to quit.\n")

    while True:
        try:
            user_input = input("You: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nGoodbye!")
            break

        if not user_input:
            continue
        if user_input.lower() in {"exit", "quit", "bye"}:
            print("Goodbye!")
            break

        history.append({"role": "user", "content": user_input})

        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system,
            messages=history,
        )

        reply = response.content[0].text
        history.append({"role": "assistant", "content": reply})

        print(f"\nClaude: {reply}\n")


def main():
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print(
            "⚠  ANTHROPIC_API_KEY is not set.\n"
            "   Set it with:  set ANTHROPIC_API_KEY=sk-ant-...\n"
            "   or add it to a .env file in this folder.\n"
        )

    # Script always asks the user for the Excel file path at runtime
    xlsx_path = input("📂  Enter the full path to your Excel file: ").strip().strip('"')

    print(f"📂  Loading: {xlsx_path}")
    data = load_workbook(xlsx_path)
    total = len(data.get("Email Tracker", []))
    print(f"✅  Loaded {total} Email Tracker rows across {len(data)} sheets.")

    chat(data)


if __name__ == "__main__":
    main()
