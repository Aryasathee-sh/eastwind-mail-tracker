"""
server_cloud.py
---------------
Cloud-ready Remote MCP Server for Eastwind Proposals Mail Tracker.
Runs using Server-Sent Events (SSE) / Streamable-HTTP so clients
can connect directly from Claude Desktop anywhere via URL.
"""

import os
import sys
from pathlib import Path
from mcp_server import server

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"🚀 Starting Eastwind Mail Tracker Cloud MCP on {host}:{port}/sse ...")
    server.run(transport="sse", host=host, port=port)
