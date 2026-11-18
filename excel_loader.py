"""
excel_loader.py
---------------
Reads all sheets from the Eastwind Proposals Mail Tracker Excel workbook:
- Email Tracker: main tracking sheet (3000+ entries)
- Dashboard: KPIs, performance metrics, Daily dashboard
- Settings: configuration, notification mail IDs, product groups
- MT_Users: user accounts, roles, permissions, approver flags
- MT_Lists: reference lists (categories, statuses, priorities, tones)
- MT_Settings: system-level configurations
- MT_Audit: full history audit log of changes made by users
"""

import zipfile
import xml.etree.ElementTree as ET
import re
from datetime import datetime, timedelta
from pathlib import Path

NS = {"ss": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_EPOCH = datetime(1899, 12, 30)


def _excel_date(serial: str) -> str:
    try:
        return (EXCEL_EPOCH + timedelta(days=float(serial))).strftime("%Y-%m-%d")
    except Exception:
        return serial


def _excel_time(serial: str) -> str:
    try:
        total_seconds = int(round(float(serial) * 24 * 3600))
        hours = (total_seconds // 3600) % 24
        minutes = (total_seconds % 3600) // 60
        period = "AM" if hours < 12 else "PM"
        display_hour = hours if 1 <= hours <= 12 else (12 if hours == 0 else hours - 12)
        return f"{display_hour}:{minutes:02d} {period}"
    except Exception:
        return serial


def _load_shared_strings(zf: zipfile.ZipFile) -> list:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    result = []
    for si in root.findall("ss:si", NS):
        parts = si.findall(".//ss:t", NS)
        result.append("".join(p.text or "" for p in parts))
    return result


def _get_sheet_map(zf: zipfile.ZipFile) -> dict:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_target = {
        r.attrib["Id"]: r.attrib["Target"]
        for r in rels.findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
    }

    sheet_map = {}
    for s in wb.findall(".//ss:sheet", NS):
        name = s.attrib["name"]
        rid = s.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        target = rid_to_target.get(rid, "")
        if target:
            target_norm = target.lstrip("/")
            if target_norm.startswith("xl/"):
                path = target_norm
            else:
                path = f"xl/{target_norm}"
            sheet_map[name] = path
    return sheet_map


def _read_sheet_grid(zf: zipfile.ZipFile, path: str, strings: list) -> list:
    """Reads a sheet into a dense list of lists, preserving columns."""
    root = ET.fromstring(zf.read(path))

    def col_index(ref: str) -> int:
        letters = re.sub(r"\d", "", ref)
        idx = 0
        for ch in letters:
            idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
        return idx - 1

    rows_data = []
    for row_el in root.findall(".//ss:row", NS):
        cells = row_el.findall("ss:c", NS)
        if not cells:
            continue

        max_col = max(col_index(c.attrib["r"]) for c in cells)
        row = [""] * (max_col + 1)

        for c in cells:
            ci = col_index(c.attrib["r"])
            t = c.attrib.get("t", "")
            v_el = c.find("ss:v", NS)
            if v_el is None:
                continue
            v = v_el.text or ""

            if t == "s":
                row[ci] = strings[int(v)]
            elif t == "b":
                row[ci] = "TRUE" if v == "1" else "FALSE"
            else:
                row[ci] = v

        rows_data.append(row)
    return rows_data


def _rows_to_dicts(rows: list) -> list:
    if not rows:
        return []
    headers = [str(h).strip() if str(h).strip() else f"Col_{i}" for i, h in enumerate(rows[0])]
    result = []
    for row in rows[1:]:
        padded = row + [""] * (len(headers) - len(row))
        result.append(dict(zip(headers, padded[:len(headers)])))
    return result


def _parse_dashboard_sheet(raw_rows: list) -> dict:
    """Parse the special non-tabular layout of the Dashboard sheet into a clear dictionary."""
    metrics = {
        "title": "Service Dashboard",
        "date": "",
        "kpis": {},
        "proposal_averages_last_15_days": {},
        "employee_performance_daily": [],
        "product_sourcing_counts": {}
    }
    
    # Process row by row
    for r in raw_rows:
        row_str = " | ".join(str(cell) for cell in r if str(cell).strip())
        if not row_str:
            continue
        
        # Look for Date
        if len(r) > 2 and "Date" in str(r[1]):
            metrics["date"] = _excel_date(str(r[2]))
        
        # Look for KPIs
        if len(r) > 2 and "Total Email Count" in str(r[1]):
            metrics["kpis"]["Total Email Count"] = r[2]
            if len(r) > 5 and "Product sourcing Due" in str(r[4]):
                metrics["kpis"]["Product sourcing Due"] = r[5]
        if len(r) > 2 and "Proposal Approval Pending" in str(r[1]):
            metrics["kpis"]["Proposal Approval Pending"] = r[2]
            if len(r) > 5 and "OVER DUE EMAIL COUNT" in str(r[4]):
                metrics["kpis"]["OVER DUE EMAIL COUNT"] = r[5]
                
        # Performance rows (Midlaj, Anargha)
        if len(r) >= 6 and str(r[1]) in ["Midlaj", "Anargha", "Total"]:
            person = str(r[1])
            metrics["proposal_averages_last_15_days"][person] = {
                "proposals_last_15_days": r[2] if len(r) > 2 else "",
                "avg_per_day": r[3] if len(r) > 3 else "",
                "proposal_processing_count": r[5] if len(r) > 5 else ""
            }
            
        # Daily breakdown table (Status | Total Count | Midlaj | Anargha)
        if len(r) >= 6 and str(r[1]) not in ["Status", "Midlaj", "Anargha", "Total", "Date", "Total Email Count", "Proposal Approval Pending"]:
            status_name = str(r[1]).strip()
            if status_name and (len(r) > 2 and str(r[2]).isdigit()):
                metrics["employee_performance_daily"].append({
                    "status": status_name,
                    "total": r[2],
                    "midlaj": r[4] if len(r) > 4 else "",
                    "anargha": r[5] if len(r) > 5 else ""
                })
    return metrics


def load_workbook(xlsx_path: str) -> dict:
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    with zipfile.ZipFile(path) as zf:
        strings = _load_shared_strings(zf)
        sheet_map = _get_sheet_map(zf)

        result = {}
        for sheet_name, xml_path in sheet_map.items():
            try:
                raw_grid = _read_sheet_grid(zf, xml_path, strings)
                
                if sheet_name == "Email Tracker":
                    records = _rows_to_dicts(raw_grid)
                    for rec in records:
                        if rec.get("Date"): rec["Date"] = _excel_date(rec["Date"])
                        if rec.get("Time"): rec["Time"] = _excel_time(rec["Time"])
                        if rec.get("Replied Date"): rec["Replied Date"] = _excel_date(rec["Replied Date"])
                    result[sheet_name] = records
                elif sheet_name == "Dashboard":
                    result["Dashboard_Raw"] = raw_grid
                    result["Dashboard"] = _parse_dashboard_sheet(raw_grid)
                else:
                    # Tabular sheets: Settings, MT_Users, MT_Lists, MT_Settings, MT_Audit
                    result[sheet_name] = _rows_to_dicts(raw_grid)
            except Exception as e:
                result[sheet_name] = [{"_error": str(e)}]

    return result


def get_summary(data: dict) -> str:
    rows = data.get("Email Tracker", [])
    dash = data.get("Dashboard", {})
    users = data.get("MT_Users", [])
    lists = data.get("MT_Lists", [])
    audits = data.get("MT_Audit", [])
    
    total = len(rows)
    statuses, priorities, assignees = {}, {}, {}

    for r in rows:
        s = r.get("Status", "").strip()
        p = (r.get("Priority ", "") or r.get("Priority", "")).strip()
        a = r.get("Assigned to", "").strip()
        statuses[s] = statuses.get(s, 0) + 1
        priorities[p] = priorities.get(p, 0) + 1
        assignees[a] = assignees.get(a, 0) + 1

    lines = [
        f"=== WORKBOOK OVERVIEW ===",
        f"• Email Tracker: {total} total email records",
        f"• MT_Users: {len(users)} registered team members",
        f"• MT_Lists: {len(lists)} configured categories and statuses",
        f"• MT_Audit: {len(audits)} audit log entries",
        f"• Dashboard: As of date {dash.get('date', 'N/A')}, Total Count={dash.get('kpis', {}).get('Total Email Count', 'N/A')}",
        "",
        "=== EMAIL TRACKER STATUS BREAKDOWN ===",
        *[f"  {k or '(blank)'}: {v}" for k, v in sorted(statuses.items(), key=lambda x: -x[1])],
        "",
        "=== EMAIL TRACKER ASSIGNEE BREAKDOWN ===",
        *[f"  {k or '(blank)'}: {v}" for k, v in sorted(assignees.items(), key=lambda x: -x[1])],
    ]
    return "\n".join(lines)
