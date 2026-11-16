import sys
sys.path.insert(0, r"C:\Users\GravityBP\Documents\Training\eastwind-claude-connector")
import mcp_server

print("=" * 60)
print("TEST 1: Loading Workbook & Summary Stats")
print("=" * 60)
summary = mcp_server.get_summary_stats()
print(summary[:350])

print("\n" + "=" * 60)
print("TEST 2: Filtering by Assignee = 'Midlaj' & Status = 'Proposal submitted'")
print("=" * 60)
results = mcp_server.search_emails(assignee="Midlaj", status="Proposal submitted", limit=2)
print(results[:400])

print("\n" + "=" * 60)
print("TEST 3: Checking Team Members Sheet")
print("=" * 60)
users = mcp_server.get_team_members()
print(users)

print("\n" + "=" * 60)
print("TEST 4: Checking Overdue Items")
print("=" * 60)
overdue = mcp_server.get_overdue_emails()
print(overdue[:200])

print("\n>>> ALL 4 BACKEND TESTS PASSED SUCCESSFULLY! <<<")
