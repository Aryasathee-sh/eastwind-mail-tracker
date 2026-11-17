"""
mcp_bridge.py
--------------
Lightweight bridge for Claude Desktop.
Claude Desktop connects to this script via standard IO (stdio),
and this script forwards all tool calls, requests, and responses
to your live Cloud MCP Server on Render (SSE) in real-time.
"""

import sys
import os
import asyncio
from mcp.client.session import ClientSession
from mcp.client.sse import sse_client
from mcp.server.mcpserver import MCPServer
from mcp import types

CLOUD_URL = os.environ.get(
    "EASTWIND_CLOUD_URL", "https://eastwind-mail-tracker.onrender.com/sse"
)

bridge = MCPServer(name="Eastwind Mail Tracker Cloud Bridge")


async def register_remote_tools():
    async with sse_client(CLOUD_URL) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            remote_tools = await session.list_tools()

            for t in remote_tools.tools:
                # Dynamic tool wrapper
                def make_tool_fn(tool_name):
                    async def dynamic_tool_handler(**kwargs):
                        async with sse_client(CLOUD_URL) as (r, w):
                            async with ClientSession(r, w) as sub_session:
                                await sub_session.initialize()
                                res = await sub_session.call_tool(tool_name, kwargs)
                                return "".join(
                                    c.text if hasattr(c, "text") else str(c)
                                    for c in res.content
                                )
                    return dynamic_tool_handler

                bridge.tool(
                    name=t.name,
                    description=t.description or f"Eastwind tool: {t.name}",
                )(make_tool_fn(t.name))

            # Run bridge over stdio
            await bridge.run_stdio_async()


if __name__ == "__main__":
    try:
        asyncio.run(register_remote_tools())
    except Exception as e:
        sys.stderr.write(f"Bridge error: {e}\n")
        sys.exit(1)
