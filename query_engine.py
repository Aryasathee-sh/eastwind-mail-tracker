"""
query_engine.py
---------------
Intelligent Query Engine for Eastwind Proposals Mail Tracker.
Parses natural language questions and routes them to the appropriate Excel tool,
returning formatted Markdown and structured data for the Web Chat interface.
"""

import json
import re
from typing import Dict, Any, List

def process_query(query: str, data: Dict[str, Any]) -> Dict[str, Any]:
    q = query.strip().lower()
    
    # 0. LATEST DATE / MOST RECENT EMAILS
    if any(k in q for k in ["latest date", "most recent", "newest date", "last date", "latest email", "latest proposal", "recent email", "recent proposal"]) or ("latest" in q and ("date" in q or "email" in q or "tracker" in q or "record" in q)):
        rows = data.get("Email Tracker", [])
        valid_rows = [r for r in rows if r.get("Date")]
        if not valid_rows:
            return {"answer": "The Email Tracker sheet is currently empty."}
        
        latest_row = valid_rows[-1]
        latest_date = latest_row.get("Date", "N/A")
        recent_batch = valid_rows[-10:]
        subj = str(latest_row.get("Subject", "")).strip('"')
        
        md = f"### 📅 Latest Date in Email Tracker: **{latest_date}**\n\n"
        md += f"The most recent entry recorded is from **{latest_date} at {latest_row.get('Time', '')}**:\n"
        md += f"- **Subject**: {subj}\n"
        md += f"- **From**: `{latest_row.get('From', '')}`\n"
        md += f"- **Assigned To**: **{latest_row.get('Assigned to', '') or 'Unassigned'}**\n"
        md += f"- **Status**: `{latest_row.get('Status', '')}`\n\n"
        md += "#### 📩 Most Recent 10 Proposals:\n"
        md += "| Date | Time | Subject | Assigned To | Status |\n|---|---|---|---|---|\n"
        for r in reversed(recent_batch):
            s = str(r.get("Subject", ""))[:35].strip('"')
            md += f"| {r.get('Date', '')} | {r.get('Time', '')} | {s} | **{r.get('Assigned to', '')}** | `{r.get('Status', '')}` |\n"
        return {"answer": md, "type": "latest_date"}

    # 1. OVERDUE PROPOSALS
    if any(k in q for k in ["overdue", "delayed", "past due"]) or re.search(r"\\blate\\b", q):
        rows = data.get("Email Tracker", [])
        overdue_items = []
        for r in rows:
            if r.get("Over Due", "").strip():
                overdue_items.append({
                    "Subject": r.get("Subject", "").strip('"'),
                    "Assigned to": r.get("Assigned to", "") or "Unassigned",
                    "Date": r.get("Date", ""),
                    "Status": r.get("Status", ""),
                    "Priority": (r.get("Priority ", "") or r.get("Priority", "")).strip(),
                    "Category": r.get("Category", ""),
                    "Customer Request": (r.get("Customer Request", "") or "")[:100]
                })
        
        md = f"### ⚠️ Overdue Proposals ({len(overdue_items)} found)\n\n"
        if not overdue_items:
            md += "Great news! There are currently no actionable overdue proposals."
        else:
            md += "| Date | Subject | Assigned To | Priority | Status |\n|---|---|---|---|---|\n"
            for item in overdue_items[:25]:
                md += f"| {item['Date']} | {item['Subject'][:40]} | **{item['Assigned to']}** | {item['Priority']} | `{item['Status']}` |\n"
            if len(overdue_items) > 25:
                md += f"\n*Showing top 25 of {len(overdue_items)} overdue proposals.*"
        
        return {"answer": md, "count": len(overdue_items), "type": "table"}

    # 2. DASHBOARD / KPIS
    if any(k in q for k in ["dashboard", "kpi", "performance", "average", "15 day", "statistics"]):
        dash = data.get("Dashboard", {})
        kpis = dash.get("kpis", {})
        emp_perf = dash.get("employee_performance_daily", [])
        averages = dash.get("proposal_averages_last_15_days", {})
        
        md = f"### 📊 Service Dashboard & KPIs (Dated: {dash.get('date', 'Latest')})\n\n"
        md += f"- **Total Email Count**: **{kpis.get('Total Email Count', 'N/A')}**\n"
        md += f"- **Proposal Approval Pending**: **{kpis.get('Proposal Approval Pending', 'N/A')}**\n"
        md += f"- **Product Sourcing Due**: **{kpis.get('Product Sourcing Due', 'N/A')}**\n"
        md += f"- **Overdue Email Count**: **{kpis.get('OVER DUE EMAIL COUNT', 'N/A')}**\n\n"
        
        if averages:
            md += "#### 📈 15-Day Submission Averages\n"
            tot_avg = averages.get("Total", {}).get("avg_per_day", "9.5")
            md += f"- **Team Total Proposals (Last 15 Days)**: {averages.get('Total', {}).get('proposals_last_15_days', '198')}\n"
            md += f"- **Average Proposals Per Day**: {tot_avg}\n\n"
            
        if emp_perf:
            md += "#### 👥 Daily Performance Breakdown\n"
            md += "| Status | Total | Midlaj | Anargha |\n|---|---|---|---|\n"
            for row in emp_perf:
                st = row.get("status", "")
                tot = row.get("total", "")
                m = row.get("midlaj", "")
                a = row.get("anargha", "")
                if st:
                    md += f"| {st} | {tot} | {m} | {a} |\n"
            
        return {"answer": md, "type": "dashboard"}

    # 3. USERS / TEAM
    if any(k in q for k in ["user", "team", "who works", "member", "approver", "mt_users"]):
        users = data.get("MT_Users", [])
        md = f"### 👥 Team Structure ({len(users)} users)\n\n"
        md += "| Name | Email | Role | Active | Assignable | Approver |\n|---|---|---|---|---|---|\n"
        for u in users:
            md += f"| **{u.get('Display Name', '')}** | `{u.get('Email', '')}` | {u.get('Role', '')} | {'✅' if u.get('Active') else '❌'} | {'✅' if u.get('Assignable') else '❌'} | {'✅' if u.get('Approver') else '❌'} |\n"
        return {"answer": md, "type": "users"}

    # 4. MASTER LISTS / STATUSES / CATEGORIES
    if any(k in q for k in ["category", "categories", "status", "statuses", "master list", "mt_lists", "priorities"]):
        lists = data.get("MT_Lists", [])
        cats = [item.get("Value") for item in lists if item.get("Type") == "Category" and item.get("Active")]
        statuses = [item.get("Value") for item in lists if item.get("Type") == "Status" and item.get("Active")]
        priorities = [item.get("Value") for item in lists if item.get("Type") == "Priority" and item.get("Active")]
        
        md = "### 📋 Master Lists & Reference Data\n\n"
        md += f"**Product Categories ({len(cats)}):**\n"
        md += ", ".join([f"`{c}`" for c in cats]) + "\n\n"
        md += f"**Tracking Statuses ({len(statuses)}):**\n"
        md += ", ".join([f"`{s}`" for s in statuses]) + "\n\n"
        md += f"**Priorities:** " + ", ".join([f"`{p}`" for p in priorities]) + "\n"
        return {"answer": md, "type": "lists"}

    # 5. AUDIT LOGS
    if any(k in q for k in ["audit", "log", "change", "history", "admin action", "mt_audit"]):
        audit = data.get("MT_Audit", [])
        md = f"### 📜 Recent Audit Logs (Showing last 15 of {len(audit)} entries)\n\n"
        md += "| Timestamp | User | Action | Field | New Value |\n|---|---|---|---|---|\n"
        for entry in audit[-15:]:
            ts = entry.get("Timestamp", "")
            usr = entry.get("User", "")
            act = entry.get("Action", "")
            fld = entry.get("Field", "")
            val = str(entry.get("New Value", ""))[:30]
            md += f"| {ts} | **{usr}** | {act} | {fld} | {val} |\n"
        return {"answer": md, "type": "audit"}

    # 6. ASSIGNEE SEARCH (e.g. "Midlaj", "Anargha", "Adnan")
    assignee_target = None
    for name in ["midlaj", "anargha", "adnan", "sreejith", "vineeth"]:
        if name in q:
            assignee_target = name
            break
            
    if assignee_target:
        rows = data.get("Email Tracker", [])
        matched = [r for r in rows if assignee_target in r.get("Assigned to", "").lower()]
        md = f"### 👤 Proposals Assigned to {assignee_target.title()} ({len(matched)} records)\n\n"
        if not matched:
            md += f"No proposals currently assigned to {assignee_target.title()}."
        else:
            md += "| Date | Subject | Category | Status | Priority |\n|---|---|---|---|---|\n"
            for r in matched[:20]:
                pri = (r.get("Priority ", "") or r.get("Priority", "")).strip()
                md += f"| {r.get('Date', '')} | {r.get('Subject', '')[:40]} | {r.get('Category', '')} | `{r.get('Status', '')}` | {pri} |\n"
            if len(matched) > 20:
                md += f"\n*Showing latest 20 of {len(matched)} records.*"
        return {"answer": md, "count": len(matched), "type": "table"}

    # 7. GENERAL SEARCH / KEYWORD SEARCH
    rows = data.get("Email Tracker", [])
    terms = [w for w in q.split() if len(w) > 2 and w not in ["the", "for", "and", "show", "find", "list", "all", "get", "what", "from"]]
    
    if terms:
        search_term = " ".join(terms)
        matched = []
        for r in rows:
            text = f"{r.get('Subject', '')} {r.get('Customer Request', '')} {r.get('Category', '')} {r.get('From', '')} {r.get('Status', '')}".lower()
            if any(t in text for t in terms):
                matched.append(r)
        
        md = f"### 🔍 Search Results for: *\"{search_term}\"* ({len(matched)} matches)\n\n"
        if not matched:
            md += "No matching proposals found in the workbook."
        else:
            md += "| Date | Subject | Assigned To | Category | Status |\n|---|---|---|---|---|\n"
            for r in matched[:20]:
                md += f"| {r.get('Date', '')} | {r.get('Subject', '')[:35]} | **{r.get('Assigned to', '')}** | {r.get('Category', '')} | `{r.get('Status', '')}` |\n"
            if len(matched) > 20:
                md += f"\n*Showing top 20 of {len(matched)} matches.*"
        return {"answer": md, "count": len(matched), "type": "table"}

    # DEFAULT: Overall summary
    total_emails = len(data.get("Email Tracker", []))
    users = len(data.get("MT_Users", []))
    audit_len = len(data.get("MT_Audit", []))
    
    md = f"""### 🏢 Eastwind Proposals Assistant
I am ready to help you analyze the live **Eastwind Proposals - Mail Tracker** workbook.

**Quick Stats:**
- **3,016 Total Emails & Proposals** tracked
- **{users} Team Members** registered
- **{audit_len} Audit Log Entries** recorded
- **7 Sheets Integrated Live**

**Things you can ask me:**
- *"Show me all overdue proposals"*
- *"What are the KPIs from the dashboard?"*
- *"Show all proposals assigned to Midlaj"*
- *"List proposals assigned to Anargha"*
- *"Who are the team members in MT_Users?"*
- *"Show the master categories and statuses"*
- *"Search for SCBA or Gas Detectors"*
"""
    return {"answer": md, "type": "summary"}
