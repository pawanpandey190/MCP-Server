"""SharePoint site tools — thin delegator to sub-modules.

NOTE: Only read tools are registered. Write and provisioning tools have been
disabled because the application only has read-only permissions:
  - Sites.Read.All
  - Files.Read.All
  - User.Read.All
"""

from mcp.server.fastmcp import FastMCP

from tools.read_tools import register_read_tools


def register_site_tools(mcp: FastMCP):
    """Register read-only SharePoint tools with the MCP server."""
    register_read_tools(mcp)
