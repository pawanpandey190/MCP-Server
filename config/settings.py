"""Configuration settings for the SharePoint MCP Server."""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Basic settings
APP_NAME = "SharePoint MCP"
DEBUG = os.getenv("DEBUG", "False").lower() in ("true", "1", "t")

# SharePoint connection settings
SHAREPOINT_CONFIG = {
    "tenant_id": os.getenv("TENANT_ID", ""),
    "client_id": os.getenv("CLIENT_ID", ""),
    "client_secret": os.getenv("CLIENT_SECRET", ""),
    # Legacy single-site key — kept for backward compatibility.
    # Prefer SITE_1_URL / SITE_2_URL / ... for multi-site setups.
    "site_url": os.getenv("SITE_URL", ""),
    "username": os.getenv("USERNAME", ""),
    "password": os.getenv("PASSWORD", ""),
    "scope": [
        "https://graph.microsoft.com/.default",
        # The application must have these permissions:
        # - Sites.Read.All (for reading site content)
        # - Sites.ReadWrite.All (for modifying site content)
        # - Sites.Manage.All (for creating sites)
        # - Files.ReadWrite.All (for document operations)
    ],
}

# ---------------------------------------------------------------------------
# Multi-site registry
# ---------------------------------------------------------------------------
# Reads SITE_1_URL / SITE_1_NAME, SITE_2_URL / SITE_2_NAME, … from the env.
# Falls back to the legacy SITE_URL if no numbered entries are found so that
# existing single-site deployments continue to work without any .env changes.
#
# Example .env:
#   SITE_1_URL=https://contoso.sharepoint.com/sites/home
#   SITE_1_NAME=Home
#   SITE_2_URL=https://contoso.sharepoint.com/sites/HR
#   SITE_2_NAME=HR Portal
# ---------------------------------------------------------------------------
SITES: dict[str, str] = {}  # {display_name: site_url}

_idx = 1
while True:
    _url = os.getenv(f"SITE_{_idx}_URL", "").strip()
    if not _url:
        break
    _name = os.getenv(f"SITE_{_idx}_NAME", f"Site {_idx}").strip()
    SITES[_name] = _url
    _idx += 1

# Backward-compat fallback: honour the plain SITE_URL if no numbered sites
if not SITES and SHAREPOINT_CONFIG["site_url"]:
    SITES["Default"] = SHAREPOINT_CONFIG["site_url"]

# Keep site_url pointing at the first registered site so legacy code works
if SITES and not SHAREPOINT_CONFIG["site_url"]:
    SHAREPOINT_CONFIG["site_url"] = next(iter(SITES.values()))


# Microsoft Graph API settings
GRAPH_API_VERSION = "v1.0"
GRAPH_BASE_URL = f"https://graph.microsoft.com/{GRAPH_API_VERSION}"

# Document processing settings
DOCUMENT_PROCESSING = {
    "max_text_preview_length": 50000,  # Maximum characters for text preview
    "max_rows_preview": 500,           # Maximum rows for CSV/Excel preview
    "supported_extensions": [
        "csv",
        "xlsx",
        "xls",
        "xlsb",
        "docx",
        "pdf",
        "txt",
        "md",
        "html",
        "htm",
    ],
}

# Content generation settings
CONTENT_GENERATION = {
    "default_audience": "general",
    "default_purpose": "general",
    "enable_rich_layout": True,
}
