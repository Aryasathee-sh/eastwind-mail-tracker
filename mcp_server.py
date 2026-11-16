"""
mcp_server.py
-------------
Comprehensive MCP Server for the Eastwind Proposals Excel Workbook.
Provides Claude with access to ALL sheets and tabs:
- Email Tracker (queries, filtering, searching, overdue tracking)
- Dashboard (Service Dashboard KPIs, daily employee performance, 15-day averages)
- MT_Users (team users, roles, assignable/approver permissions)
- MT_Lists (master categories, statuses, priorities, tones)
- Settings (dropdown configs, product groups, team notification emails)
- MT_Audit (audit trails, field changes, administrative actions)
"""

import os
import sys
import json
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from starlette.responses import HTMLResponse
from excel_loader import load_workbook, get_summary

# ── Init MCP Server ───────────────────────────────────────────────────────────
server = MCPServer(name="Eastwind Mail Tracker")

# ── Locate Excel file path ────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = SCRIPT_DIR / "Eastwind_Proposals_-_Mail_Tracker.xlsx"

EXCEL_PATH = os.environ.get("EXCEL_FILE_PATH", str(DEFAULT_EXCEL))
_DATA_CACHE = None
_LAST_MODIFIED_TIME = 0


def get_data():
    """
    Returns the workbook data.
    Automatically detects if the Excel file was edited/saved on disk
    and reloads the latest records in real-time.
    """
    global _DATA_CACHE, _LAST_MODIFIED_TIME

    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel workbook not found at: {EXCEL_PATH}")

    current_mtime = os.path.getmtime(EXCEL_PATH)

    if _DATA_CACHE is None or current_mtime > _LAST_MODIFIED_TIME:
        print("[AUTO-RELOAD] Detected updated Excel file. Loading newest changes...")
        _DATA_CACHE = load_workbook(EXCEL_PATH)
        _LAST_MODIFIED_TIME = current_mtime
        print("[AUTO-RELOAD] Successfully refreshed with latest data!")

    return _DATA_CACHE


# ═════════════════════════════════════════════════════════════════════════════
# 0. ROOT WELCOME PAGE & HEALTH CHECK (For browsers & Render)
# ═════════════════════════════════════════════════════════════════════════════
@server.custom_route("/", methods=["GET", "HEAD"])
async def root_page(request):
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Eastwind Proposals MCP Service</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #0f172a; color: #f8fafc; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 2.5rem; border-radius: 1rem; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5); max-width: 600px; text-align: center; border: 1px solid #334155; }
        h1 { color: #38bdf8; margin-top: 0; font-size: 1.8rem; }
        .badge { background: #10b981; color: white; padding: 0.25rem 0.75rem; border-radius: 9999px; font-size: 0.875rem; font-weight: 600; display: inline-block; margin-bottom: 1.5rem; }
        p { color: #94a3b8; line-height: 1.6; }
        .code-box { background: #0f172a; padding: 1rem; border-radius: 0.5rem; text-align: left; font-family: monospace; font-size: 0.9rem; color: #e2e8f0; border: 1px solid #334155; overflow-x: auto; margin: 1.5rem 0; }
        .url { color: #a855f7; font-weight: bold; }
    </style>
</head>
<body>
    <div class="card">
        <span class="badge">● Service Live & Operational</span>
        <h1>Eastwind Proposals Mail Tracker</h1>
        <p>This is a <strong>Model Context Protocol (MCP) Cloud Server</strong> designed for <strong>Claude Desktop</strong> integration.</p>
        
        <p>To connect Claude Desktop to this service, use the <strong>/sse</strong> endpoint:</p>
        <div class="code-box">
{<br>
&nbsp;&nbsp;"mcpServers": {<br>
&nbsp;&nbsp;&nbsp;&nbsp;"eastwind-proposals": {<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"url": "<span class="url">https://eastwind-mail-tracker.onrender.com/sse</span>"<br>
&nbsp;&nbsp;&nbsp;&nbsp;}<br>
&nbsp;&nbsp;}<br>
}
        </div>
        <p style="font-size: 0.85rem; color: #64748b;">Powered by FastMCP v2 & Anthropic Model Context Protocol</p>
    </div>
</body>
</html>"""
    return HTMLResponse(content=html_content)


# ═════════════════════════════════════════════════════════════════════════════
# 1. OVERALL WORKBOOK SUMMARY
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_workbook_summary",
    description="Get an overview across ALL sheets in the workbook (Email Tracker, Dashboard KPIs, Users, Master Lists, Audit logs).",
)
def get_workbook_summary() -> str:
    data = get_data()
    return get_summary(data)


# ═════════════════════════════════════════════════════════════════════════════
# 2. DASHBOARD SHEET (KPIs & Performance)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_dashboard_sheet",
    description="Get all KPIs and performance metrics from the 'Dashboard' sheet, including employee daily performance, 15-day averages, overdue counts, and total emails.",
)
def get_dashboard_sheet() -> str:
    data = get_data()
    dash = data.get("Dashboard", {})
    return json.dumps(dash, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 3. EMAIL TRACKER: SEARCH & FILTER
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="search_email_tracker",
    description="Search and filter records in 'Email Tracker' by status, assignee, priority, category, sender, date, or keyword in subject/request.",
)
def search_email_tracker(
    status: str = "",
    assignee: str = "",
    priority: str = "",
    category: str = "",
    sender: str = "",
    date: str = "",
    keyword: str = "",
    limit: int = 50,
) -> str:
    data = get_data()
    rows = data.get("Email Tracker", [])
    results = []

    for row in rows:
        if status and status.lower() not in row.get("Status", "").lower():
            continue
        if assignee and assignee.lower() not in row.get("Assigned to", "").lower():
            continue
        row_priority = (row.get("Priority ", "") or row.get("Priority", "")).strip()
        if priority and priority.lower() not in row_priority.lower():
            continue
        if category and category.lower() not in row.get("Category", "").lower():
            continue
        if sender and sender.lower() not in row.get("From", "").lower():
            continue
        if date and not str(row.get("Date", "")).startswith(date):
            continue
        if keyword:
            subject = row.get("Subject", "").lower()
            request = row.get("Customer Request", "").lower()
            kw = keyword.lower()
            if kw not in subject and kw not in request:
                continue
        results.append(row)

    if not results:
        return "No email records found matching the specified criteria."

    total = len(results)
    limited = results[:limit]
    return f"Found {total} records (showing first {len(limited)}):\n\n" + json.dumps(limited, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 4. EMAIL TRACKER: RECENT EMAILS
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_recent_email_tracker",
    description="Get the latest N records from the 'Email Tracker' sheet (default: 25).",
)
def get_recent_email_tracker(limit: int = 25) -> str:
    data = get_data()
    limit = min(max(1, limit), 100)
    rows = data.get("Email Tracker", [])[-limit:]
    if not rows:
        return "No email tracker data found."
    return json.dumps(rows, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 5. EMAIL TRACKER: OVERDUE PROPOSALS
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_overdue_proposals",
    description="List all emails from the 'Email Tracker' where the 'Over Due' column is marked.",
)
def get_overdue_proposals() -> str:
    data = get_data()
    rows = data.get("Email Tracker", [])
    overdue = [r for r in rows if r.get("Over Due", "").strip()]
    if not overdue:
        return "No overdue proposals found."
    return f"Found {len(overdue)} overdue items:\n\n" + json.dumps(overdue, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 6. MT_USERS SHEET (Team & Permissions)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_users_sheet",
    description="Get all user accounts from 'MT_Users' sheet, including Display Name, Email, Role, Active status, Assignable, and Approver flags.",
)
def get_users_sheet() -> str:
    data = get_data()
    users = data.get("MT_Users", [])
    if not users:
        return "No user records found in MT_Users."
    return json.dumps(users, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 7. MT_LISTS SHEET (Master Categories, Statuses, Priorities, Tones)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_master_lists_sheet",
    description="Get all master reference lists from 'MT_Lists' sheet (All valid Categories, Statuses, Priorities, and Tones with Sort Order and Closed flags).",
)
def get_master_lists_sheet(filter_type: str = "") -> str:
    data = get_data()
    lists = data.get("MT_Lists", [])
    if filter_type:
        lists = [item for item in lists if item.get("Type", "").lower() == filter_type.lower()]
    return json.dumps(lists, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 8. SETTINGS SHEET (Dropdowns, Product Groups, Notification Emails)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_settings_sheet",
    description="Get configuration from the 'Settings' sheet: Product Groups, Team Notification Mail IDs, Priority settings, and Assignees.",
)
def get_settings_sheet() -> str:
    data = get_data()
    settings = data.get("Settings", [])
    return json.dumps(settings, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 9. MT_AUDIT SHEET (Audit Trail / Activity History)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="get_audit_log_sheet",
    description="Query user audit history and change logs from 'MT_Audit' sheet (User, Action, Field, Old Value, New Value, Timestamp).",
)
def get_audit_log_sheet(user: str = "", action: str = "", limit: int = 50) -> str:
    data = get_data()
    audits = data.get("MT_Audit", [])
    results = []

    for a in audits:
        if user and user.lower() not in a.get("User", "").lower():
            continue
        if action and action.lower() not in a.get("Action", "").lower():
            continue
        results.append(a)

    if not results:
        return "No audit logs found matching criteria."

    total = len(results)
    limited = results[-limit:]
    return f"Total matching audit events: {total} (showing latest {len(limited)}):\n\n" + json.dumps(limited, ensure_ascii=False, indent=2)


# ═════════════════════════════════════════════════════════════════════════════
# 10. GENERIC SHEET READER (Any Sheet by Name)
# ═════════════════════════════════════════════════════════════════════════════
@server.tool(
    name="read_any_sheet",
    description="Read data from any sheet in the workbook by its exact name: 'Email Tracker', 'Dashboard', 'Settings', 'MT_Users', 'MT_Lists', 'MT_Settings', or 'MT_Audit'.",
)
def read_any_sheet(sheet_name: str, limit: int = 50) -> str:
    data = get_data()
    if sheet_name not in data:
        available = list(data.keys())
        return f"Sheet '{sheet_name}' not found. Available sheets: {available}"

    sheet_data = data[sheet_name]
    if isinstance(sheet_data, list):
        total = len(sheet_data)
        limited = sheet_data[:limit]
        return f"Sheet '{sheet_name}' has {total} rows. Showing first {len(limited)}:\n\n" + json.dumps(limited, ensure_ascii=False, indent=2)
    elif isinstance(sheet_data, dict):
        return f"Sheet '{sheet_name}' (Structured Object):\n\n" + json.dumps(sheet_data, ensure_ascii=False, indent=2)
    return str(sheet_data)


# ═════════════════════════════════════════════════════════════════════════════
# Execution Entry Point
# ═════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    server.run(transport="stdio")
