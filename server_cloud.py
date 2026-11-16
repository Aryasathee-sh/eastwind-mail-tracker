"""
server_cloud.py
---------------
Cloud-ready Remote MCP Server for Eastwind Proposals Mail Tracker.
Exposes SSE endpoint at /sse and message handling at /messages/
for remote Claude clients over HTTPS.
"""

import os
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from mcp_server import server

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "0.0.0.0")

    print(f"🚀 Starting Eastwind Mail Tracker Cloud MCP on {host}:{port}...")
    print(f"   SSE Endpoint: http://{host}:{port}/sse")
    print(f"   Messages:     http://{host}:{port}/messages/")

    server.run(transport="sse", host=host, port=port)
