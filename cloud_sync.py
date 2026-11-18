"""
cloud_sync.py
-------------
Runs on GitHub Actions or local daemon.
Connects to Microsoft Graph API, compares SharePoint table with local/cloud,
rebuilds clean OpenXML workbook and pushes to Render cloud.
Handles row deletions, additions, and modifications.
"""

import os
import sys
import json
import zipfile
import shutil
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET

SITE_ID = "gravitybp365.sharepoint.com,a3fe8b15-b97f-435a-9877-109104614301,a21796fe-5c45-4cf3-b793-2113eeef9840"
LIST_ID = "23371290-0548-44d5-857a-c56caf79da85"
ITEM_ID = "2079"
SHEET_NAME = urllib.parse.quote("Email Tracker")
RENDER_UPLOAD_URL = "https://eastwind-mail-tracker.onrender.com/api/upload"
SRC_XLSX = os.environ.get("EXCEL_PATH", "Eastwind_Proposals_-_Mail_Tracker.xlsx")

NS_VAL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS_VAL)


def sync():
    token = os.environ.get("MS_GRAPH_TOKEN")
    if not token:
        import subprocess
        try:
            token = subprocess.check_output(
                "az account get-access-token --resource-type ms-graph --query accessToken -o tsv",
                shell=True
            ).decode().strip()
        except Exception as e:
            print("Failed to get token via az CLI:", e)
            return

    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items/{ITEM_ID}/driveItem/workbook/worksheets/{SHEET_NAME}/tables/Received/rows"
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            all_sp_rows = json.loads(resp.read().decode()).get("value", [])
    except Exception as e:
        print("SharePoint fetch failed:", e)
        return

    print(f"Fetched {len(all_sp_rows)} live rows from SharePoint.")
    if not all_sp_rows:
        return

    work_dir = SRC_XLSX + "_tmp_cloud"
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)

    with zipfile.ZipFile(SRC_XLSX, "r") as zin:
        zin.extractall(work_dir)

    ss_path = os.path.join(work_dir, "xl", "sharedStrings.xml")
    ss_tree = ET.parse(ss_path)
    ss_root = ss_tree.getroot()

    def add_shared_string(text):
        si = ET.SubElement(ss_root, f"{{{NS_VAL}}}si")
        t = ET.SubElement(si, f"{{{NS_VAL}}}t")
        t.text = str(text) if text is not None else ""
        return len(ss_root) - 1

    sheet1_path = os.path.join(work_dir, "xl", "worksheets", "sheet1.xml")
    s1_tree = ET.parse(sheet1_path)
    s1_root = s1_tree.getroot()
    sheetData = s1_root.find(f"{{{NS_VAL}}}sheetData")

    header_row = None
    for row in sheetData.findall(f"{{{NS_VAL}}}row"):
        if row.attrib.get("r") == "1":
            header_row = row
            break

    sheetData.clear()
    if header_row is not None:
        sheetData.append(header_row)

    cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
    curr_row = 2

    for r in all_sp_rows:
        vals = r.get("values", [[]])[0]
        row_el = ET.SubElement(sheetData, f"{{{NS_VAL}}}row", {
            "r": str(curr_row),
            "spans": "1:16",
            "ht": "18",
            "customHeight": "1"
        })
        for c_i, c_letter in enumerate(cols):
            if c_i < len(vals) and vals[c_i] is not None and str(vals[c_i]).strip() != "":
                val = vals[c_i]
                cell_ref = f"{c_letter}{curr_row}"
                if c_i in (1, 2):
                    c_el = ET.SubElement(row_el, f"{{{NS_VAL}}}c", {"r": cell_ref})
                    v_el = ET.SubElement(c_el, f"{{{NS_VAL}}}v")
                    v_el.text = str(val)
                else:
                    s_idx = add_shared_string(val)
                    c_el = ET.SubElement(row_el, f"{{{NS_VAL}}}c", {"r": cell_ref, "t": "s"})
                    v_el = ET.SubElement(c_el, f"{{{NS_VAL}}}v")
                    v_el.text = str(s_idx)
        curr_row += 1

    ss_root.attrib["count"] = str(len(ss_root))
    ss_root.attrib["uniqueCount"] = str(len(ss_root))
    ss_tree.write(ss_path, xml_declaration=True, encoding="utf-8")

    dim = s1_root.find(f"{{{NS_VAL}}}dimension")
    if dim is not None:
        dim.attrib["ref"] = f"A1:P{curr_row - 1}"
    s1_tree.write(sheet1_path, xml_declaration=True, encoding="utf-8")

    with zipfile.ZipFile(SRC_XLSX, "w", zipfile.ZIP_DEFLATED) as zout:
        for foldername, subfolders, filenames in os.walk(work_dir):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, work_dir)
                zout.write(filepath, arcname)

    shutil.rmtree(work_dir)

    with open(SRC_XLSX, "rb") as f:
        data = f.read()

    up_req = urllib.request.Request(
        RENDER_UPLOAD_URL,
        data=data,
        headers={"Content-Type": "application/octet-stream"},
        method="POST"
    )
    with urllib.request.urlopen(up_req, timeout=30) as resp:
        print("Pushed to Render:", resp.read().decode())


if __name__ == "__main__":
    sync()
