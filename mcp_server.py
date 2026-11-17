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
import time
import tempfile
import urllib.request
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from starlette.responses import HTMLResponse, JSONResponse
from excel_loader import load_workbook, get_summary
from query_engine import process_query

# ── Init MCP Server ───────────────────────────────────────────────────────────
server = MCPServer(name="Eastwind Mail Tracker")

# ── Locate Excel file path ────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = SCRIPT_DIR / "Eastwind_Proposals_-_Mail_Tracker.xlsx"

EXCEL_PATH = os.environ.get("EXCEL_FILE_PATH", str(DEFAULT_EXCEL))

# ── Live SharePoint URL (fetch on every query) ────────────────────────────────
SHAREPOINT_DOWNLOAD_URL = (
    "https://gravitybp365.sharepoint.com/:x:/s/GravityProposalManagement/"
    "IQB2a4F6JCnBRZfwL60RGcQsAaaUrJk1YE_ByL4obrNDnpw?e=VD5aEf&download=1"
)
SHAREPOINT_CACHE_TTL = 300  # re-fetch from SharePoint every 5 minutes
_DATA_CACHE = None
_LAST_MODIFIED_TIME = 0
_SHAREPOINT_FETCH_TIME = 0


def fetch_from_sharepoint():
    """Download latest Excel from SharePoint sharing link and reload data."""
    global _DATA_CACHE, _LAST_MODIFIED_TIME, _SHAREPOINT_FETCH_TIME
    try:
        print("[SHAREPOINT] Fetching latest Excel from SharePoint...")
        req = urllib.request.Request(
            SHAREPOINT_DOWNLOAD_URL,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            content = resp.read()
        # Validate it's a real xlsx (zip) file — first bytes must be PK
        if not content[:2] == b'PK':
            print("[SHAREPOINT] ⚠️ Response is not a valid xlsx file (got HTML/redirect). Skipping update.")
            _SHAREPOINT_FETCH_TIME = time.time()  # don't hammer — wait 5 min before retry
            return
        # Write to temp file then overwrite local Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(content)
            tmp_path = tmp.name
        import shutil
        shutil.move(tmp_path, EXCEL_PATH)
        _DATA_CACHE = load_workbook(EXCEL_PATH)
        _LAST_MODIFIED_TIME = os.path.getmtime(EXCEL_PATH)
        _SHAREPOINT_FETCH_TIME = time.time()
        print(f"[SHAREPOINT] ✅ Loaded {len(_DATA_CACHE.get('Email Tracker', []))} rows from SharePoint.")
    except Exception as e:
        print(f"[SHAREPOINT] ⚠️ Could not fetch from SharePoint: {e}. Using cached data.")
        _SHAREPOINT_FETCH_TIME = time.time()  # wait before retry


def get_data():
    """
    Returns the workbook data.
    Fetches from SharePoint every 5 minutes for live data.
    Falls back to local file if SharePoint is unreachable.
    """
    global _DATA_CACHE, _LAST_MODIFIED_TIME

    # Try SharePoint fetch if cache is expired
    if time.time() - _SHAREPOINT_FETCH_TIME > SHAREPOINT_CACHE_TTL:
        fetch_from_sharepoint()

    # Fallback: load local file if cache is still empty
    if _DATA_CACHE is None:
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"Excel workbook not found at: {EXCEL_PATH}")
        current_mtime = os.path.getmtime(EXCEL_PATH)
        if _DATA_CACHE is None or current_mtime > _LAST_MODIFIED_TIME:
            _DATA_CACHE = load_workbook(EXCEL_PATH)
            _LAST_MODIFIED_TIME = current_mtime

    return _DATA_CACHE


# ═════════════════════════════════════════════════════════════════════════════
# 0. ROOT CHAT PORTAL & API ENDPOINTS
# ═════════════════════════════════════════════════════════════════════════════
@server.custom_route("/", methods=["GET", "HEAD"])
@server.custom_route("/chat", methods=["GET", "HEAD"])
async def root_chat_portal(request):
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Eastwind Proposals AI Assistant</title>
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <style>
        :root {
            --bg-dark: #0f172a;
            --card-dark: #1e293b;
            --border: #334155;
            --accent: #38bdf8;
            --accent-hover: #0284c7;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --user-bubble: #0369a1;
            --bot-bubble: #1e293b;
            --badge: #10b981;
        }
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: var(--bg-dark);
            color: var(--text-main);
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            height: 100vh;
        }
        header {
            background: var(--card-dark);
            border-bottom: 1px solid var(--border);
            padding: 0.9rem 1.5rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .header-left { display: flex; align-items: center; gap: 0.75rem; }
        .logo-icon {
            width: 38px; height: 38px; border-radius: 8px;
            background: linear-gradient(135deg, #0284c7, #38bdf8);
            display: flex; align-items: center; justify-content: center;
            font-size: 1.2rem; font-weight: bold; color: white;
        }
        .header-title h1 { margin: 0; font-size: 1.15rem; color: var(--text-main); }
        .header-title p { margin: 0; font-size: 0.78rem; color: var(--text-muted); }
        .badge {
            background: rgba(16, 185, 129, 0.15);
            color: var(--badge);
            border: 1px solid rgba(16, 185, 129, 0.3);
            padding: 0.3rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 600;
            display: flex; align-items: center; gap: 0.4rem;
        }
        .badge-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--badge); }

        #chat-container {
            flex: 1;
            overflow-y: auto;
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
            max-width: 950px;
            width: 100%;
            margin: 0 auto;
        }
        .message-row { display: flex; gap: 0.75rem; width: 100%; }
        .message-row.user { justify-content: flex-end; }
        .avatar {
            width: 32px; height: 32px; border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            font-size: 0.85rem; font-weight: bold; flex-shrink: 0;
        }
        .avatar.bot { background: #0284c7; color: white; }
        .avatar.user { background: #64748b; color: white; }
        .bubble {
            max-width: 82%;
            padding: 1rem 1.25rem;
            border-radius: 1rem;
            font-size: 0.92rem;
            line-height: 1.6;
        }
        .bubble.user {
            background: var(--user-bubble);
            color: white;
            border-bottom-right-radius: 0.2rem;
        }
        .bubble.bot {
            background: var(--bot-bubble);
            color: var(--text-main);
            border: 1px solid var(--border);
            border-bottom-left-radius: 0.2rem;
        }
        .bubble table {
            width: 100%; border-collapse: collapse; margin: 0.75rem 0; font-size: 0.82rem;
        }
        .bubble th, .bubble td {
            border: 1px solid var(--border); padding: 0.45rem 0.6rem; text-align: left;
        }
        .bubble th { background: rgba(255,255,255,0.06); color: var(--accent); }
        .bubble tr:nth-child(even) { background: rgba(255,255,255,0.02); }
        .bubble code { background: #0f172a; padding: 0.15rem 0.35rem; border-radius: 4px; font-family: monospace; font-size: 0.85rem; color: #f472b6; }
        .bubble pre code { display: block; padding: 0.75rem; overflow-x: auto; color: #e2e8f0; }

        .suggestions {
            display: flex; flex-wrap: wrap; gap: 0.5rem; margin-top: 0.75rem;
        }
        .chip {
            background: #0f172a; border: 1px solid var(--border);
            color: #cbd5e1; padding: 0.35rem 0.75rem; border-radius: 9999px;
            font-size: 0.78rem; cursor: pointer; transition: all 0.2s;
        }
        .chip:hover { border-color: var(--accent); color: var(--accent); background: rgba(56, 189, 248, 0.08); }

        #input-panel {
            background: var(--card-dark);
            border-top: 1px solid var(--border);
            padding: 1rem 1.5rem;
        }
        .input-wrapper {
            max-width: 950px; margin: 0 auto; display: flex; gap: 0.75rem;
        }
        input[type="text"] {
            flex: 1; background: #0f172a; border: 1px solid var(--border);
            color: var(--text-main); padding: 0.85rem 1.1rem; border-radius: 0.65rem;
            font-size: 0.95rem; outline: none; transition: border-color 0.2s;
        }
        input[type="text"]:focus { border-color: var(--accent); }
        button.send-btn {
            background: linear-gradient(135deg, #0284c7, #38bdf8);
            border: none; color: white; padding: 0.85rem 1.4rem;
            border-radius: 0.65rem; font-weight: 600; cursor: pointer;
            transition: opacity 0.2s; display: flex; align-items: center; gap: 0.4rem;
        }
        button.send-btn:hover { opacity: 0.9; }
        .upload-btn {
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid var(--border);
            color: var(--accent);
            padding: 0.4rem 0.85rem;
            border-radius: 0.5rem;
            font-size: 0.82rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            display: flex; align-items: center; gap: 0.4rem;
        }
        .upload-btn:hover {
            background: rgba(56, 189, 248, 0.15);
            border-color: var(--accent);
        }
        .loading-dots span {
            animation: blink 1.4s infinite both; font-size: 1.4rem; line-height: 0;
        }
        .loading-dots span:nth-child(2) { animation-delay: 0.2s; }
        .loading-dots span:nth-child(3) { animation-delay: 0.4s; }
        @keyframes blink { 0%, 80%, 100% { opacity: 0; } 40% { opacity: 1; } }
    </style>
</head>
<body>
    <header>
        <div class="header-left">
            <div class="logo-icon">EW</div>
            <div class="header-title">
                <h1>Eastwind Proposals Assistant</h1>
                <p>Live Mail Tracker Intelligence Portal</p>
            </div>
        </div>
        <div style="display: flex; align-items: center; gap: 0.75rem;">
            <label for="excelFile" class="upload-btn" title="Upload updated Excel workbook">
                📁 Upload New Excel
            </label>
            <input type="file" id="excelFile" accept=".xlsx" style="display: none;" onchange="handleFileUpload(event)">
            <div class="badge">
                <div class="badge-dot"></div>
                <span id="recordCountText">3,016 Records Live</span>
            </div>
        </div>
    </header>

    <div id="chat-container">
        <div class="message-row">
            <div class="avatar bot">AI</div>
            <div class="bubble bot">
                <strong>Hello! Welcome to the Eastwind Proposals Portal.</strong><br>
                I have full access to all 7 sheets of the master Excel workbook. You can query live proposals, KPIs, assignee workloads, and master lists directly from your browser.
                
                <div class="suggestions">
                    <span class="chip" onclick="sendSuggestion('Show me all overdue proposals')">⚠️ Overdue Proposals</span>
                    <span class="chip" onclick="sendSuggestion('What are the KPIs from the Dashboard?')">📊 Dashboard KPIs</span>
                    <span class="chip" onclick="sendSuggestion('Show proposals assigned to Midlaj')">👤 Midlaj Proposals</span>
                    <span class="chip" onclick="sendSuggestion('Show proposals assigned to Anargha')">👤 Anargha Proposals</span>
                    <span class="chip" onclick="sendSuggestion('List all team members in MT_Users')">👥 Team Structure</span>
                    <span class="chip" onclick="sendSuggestion('Show master categories and statuses')">📋 Master Lists</span>
                    <span class="chip" onclick="sendSuggestion('Show recent audit logs')">📜 Recent Audit Logs</span>
                </div>
            </div>
        </div>
    </div>

    <div id="input-panel">
        <form class="input-wrapper" onsubmit="handleSend(event)">
            <input type="text" id="userInput" placeholder="Ask anything about proposals, KPIs, assignees, or overdue items..." autocomplete="off">
            <button type="submit" class="send-btn">Send ➔</button>
        </form>
    </div>

    <script>
        const chatContainer = document.getElementById('chat-container');
        const userInput = document.getElementById('userInput');

        function appendMessage(sender, text, isMarkdown = true) {
            const row = document.createElement('div');
            row.className = 'message-row ' + sender;
            
            const avatar = document.createElement('div');
            avatar.className = 'avatar ' + sender;
            avatar.innerText = sender === 'user' ? 'You' : 'AI';
            
            const bubble = document.createElement('div');
            bubble.className = 'bubble ' + sender;
            if (sender === 'bot' && isMarkdown && window.marked) {
                bubble.innerHTML = marked.parse(text);
            } else {
                bubble.innerText = text;
            }

            if (sender === 'user') {
                row.appendChild(bubble);
                row.appendChild(avatar);
            } else {
                row.appendChild(avatar);
                row.appendChild(bubble);
            }

            chatContainer.appendChild(row);
            chatContainer.scrollTop = chatContainer.scrollHeight;
            return bubble;
        }

        async function handleSend(e) {
            if (e) e.preventDefault();
            const text = userInput.value.trim();
            if (!text) return;
            
            userInput.value = '';
            appendMessage('user', text, false);

            const loadingBubble = appendMessage('bot', '<span class="loading-dots">Searching Excel records<span>.</span><span>.</span><span>.</span></span>', false);

            try {
                const response = await fetch('/api/chat', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ query: text })
                });

                const data = await response.json();
                if (window.marked && data.answer) {
                    loadingBubble.innerHTML = marked.parse(data.answer);
                } else {
                    loadingBubble.innerText = data.answer || "No response received.";
                }
            } catch (err) {
                loadingBubble.innerText = "Error querying live records: " + err.message;
            }
            chatContainer.scrollTop = chatContainer.scrollHeight;
        }

        function sendSuggestion(text) {
            userInput.value = text;
            handleSend();
        }

        async function handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            if (!file.name.endsWith('.xlsx')) {
                alert('Please select an Excel (.xlsx) file.');
                return;
            }

            const loadingBubble = appendMessage('bot', `⏳ Uploading and processing **${file.name}**...`, true);

            try {
                const response = await fetch('/api/upload', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/octet-stream',
                        'X-Filename': encodeURIComponent(file.name)
                    },
                    body: file
                });

                const data = await response.json();
                if (data.success) {
                    loadingBubble.innerHTML = marked.parse(`✅ **Workbook Updated Successfully!**\n\n- ${data.message}\n- New records are live and ready to query immediately.`);
                    if (data.total_records) {
                        document.getElementById('recordCountText').innerText = `${data.total_records.toLocaleString()} Records Live`;
                    }
                } else {
                    loadingBubble.innerHTML = `❌ Upload failed: ${data.error || 'Unknown error'}`;
                }
            } catch (err) {
                loadingBubble.innerHTML = `❌ Error uploading file: ${err.message}`;
            }
            event.target.value = '';
            chatContainer.scrollTop = chatContainer.scrollHeight;
        }
    </script>
</body>
</html>"""
    return HTMLResponse(content=html_content)


@server.custom_route("/api/chat", methods=["POST"])
async def api_chat(request):
    try:
        body = await request.json()
        query = body.get("query", "")
        data = get_data()
        result = process_query(query, data)
        return JSONResponse(result)
    except Exception as e:
        return JSONResponse({"answer": f"Error processing query: {str(e)}", "type": "error"}, status_code=500)


@server.custom_route("/api/upload", methods=["POST"])
async def api_upload(request):
    global _DATA_CACHE, _LAST_MODIFIED_TIME
    try:
        contents = await request.body()
        if not contents or len(contents) < 500:
            return JSONResponse({"success": False, "error": "Invalid or empty file."}, status_code=400)

        # Write uploaded file directly to EXCEL_PATH
        with open(EXCEL_PATH, "wb") as f:
            f.write(contents)

        # Force cache reload
        _DATA_CACHE = None
        _LAST_MODIFIED_TIME = 0
        fresh_data = get_data()
        count = len(fresh_data.get("Email Tracker", []))

        return JSONResponse({
            "success": True,
            "message": f"Successfully updated workbook with {count:,} proposal records across all 7 sheets!",
            "total_records": count
        })
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, status_code=500)


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
