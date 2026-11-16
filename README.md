# Eastwind Mail Tracker — Claude Desktop MCP Connector

This project turns your **Eastwind Proposals Mail Tracker** Excel workbook into an **MCP Server** that directly connects to **Claude Desktop**. 

When your client opens Claude Desktop, Claude automatically gets tools to read, filter, query, and analyze the live Excel tracker.

---

## 🛠️ Tools Exposed to Claude

Claude will automatically have access to these tools inside the chat:

1. **`get_summary_stats`**: Overview of total records, status counts (Replied, Submitted, etc.), priority breakdowns, and assignee counts.
2. **`search_emails`**: Search by status, assignee, priority, category, or keywords across Subjects, Requests, and Senders.
3. **`get_recent_emails`**: Review the latest emails added to the tracker.
4. **`get_emails_by_date`**: Filter emails by specific date (e.g., `2026-11-10`).
5. **`get_overdue_emails`**: Check for overdue proposal items.
6. **`get_team_members`**: View team users, active status, roles, and approvers from `MT_Users`.
7. **`get_settings`**: View dropdown lists, product groups, categories, and audit references.
8. **`get_dashboard_metrics`**: Access KPIs from the Dashboard sheet.

---

## 🚀 Setup on Client's PC (Step-by-Step)

### 1. Requirements
Ensure Python 3.10+ is installed on the machine.

### 2. Install Dependencies
Open Command Prompt / Terminal and run:
```bash
pip install mcp
```

### 3. Place the Project Folder & Workbook
Ensure the `eastwind-claude-connector` folder is placed on the client's PC (e.g. `C:\Users\Client\Documents\eastwind-claude-connector`). Place their `Eastwind_Proposals_-_Mail_Tracker.xlsx` inside this folder (or keep it wherever they store it).

### 4. Configure Claude Desktop
Open `%APPDATA%\Claude\claude_desktop_config.json` (on Windows: `C:\Users\<Username>\AppData\Roaming\Claude\claude_desktop_config.json`).

Add the `eastwind-mail-tracker` server entry under `mcpServers`:

```json
{
  "mcpServers": {
    "eastwind-mail-tracker": {
      "command": "python",
      "args": [
        "C:\\Users\\GravityBP\\Documents\\Training\\eastwind-claude-connector\\mcp_server.py"
      ],
      "env": {
        "EXCEL_FILE_PATH": "C:\\Users\\GravityBP\\Documents\\Training\\eastwind-claude-connector\\Eastwind_Proposals_-_Mail_Tracker.xlsx"
      }
    }
  }
}
```
*(Adjust the file paths if running on another machine).*

### 5. Restart Claude Desktop
Close Claude Desktop completely and reopen it.
You will see the 🔨 hammer icon in Claude Desktop showing the new **Eastwind Mail Tracker** tools ready to use.

---

## 💬 Sample Prompts for the Client

Your client can simply ask in natural language:

- *"How many proposals were submitted by Midlaj recently?"*
- *"Give me a summary of all High priority emails received this week."*
- *"Are there any overdue proposal items?"*
- *"Show me emails from Dräger regarding Gas Detectors."*
- *"What is our total email count and current breakdown by status?"*
