"""Site and document library operations mixin for GraphClient."""

import logging
from typing import Dict, Any

logger = logging.getLogger("graph_client")


class _GraphSiteOpsMixin:
    """Site and top-level library operations for the Microsoft Graph API."""

    async def get_site_info(self, domain: str, site_name: str) -> Dict[str, Any]:
        """Get SharePoint site information."""
        if site_name == "root" or not site_name:
            endpoint = f"sites/{domain}:"
        else:
            endpoint = f"sites/{domain}:/sites/{site_name}"
        logger.info(f"Getting site info for domain: {domain}, site: {site_name}")
        return await self.get(endpoint)

    async def list_document_libraries(
        self, domain: str, site_name: str
    ) -> Dict[str, Any]:
        """List only visible (non-hidden) document libraries in the site.

        Uses /lists?$expand=drive instead of /drives because the Lists
        endpoint exposes the 'hidden' property, which is the official
        SharePoint signal for system/internal libraries.  Drives endpoint
        has no such filter and always returns everything.
        """
        site_info = await self.get_site_info(domain, site_name)
        site_id = site_info.get("id")

        if not site_id:
            raise Exception(
                f"Failed to get site ID for domain: {domain}, site: {site_name}"
            )

        # $expand=drive fetches the associated drive ID in one call.
        # We select only what we need to keep the payload small.
        endpoint = (
            f"sites/{site_id}/lists"
            f"?$select=id,displayName,description,webUrl,list"
            f"&$expand=drive($select=id,driveType)"
        )
        logger.info(f"Listing document libraries (via lists) for site ID: {site_id}")
        result = await self.get(endpoint)

        # ---------------------------------------------------------------
        # Filter rules:
        #   1. list.hidden == false  →  excludes all system/internal libs
        #   2. list.template == "documentLibrary"  →  excludes plain lists
        # ---------------------------------------------------------------
        visible = [
            item for item in result.get("value", [])
            if not item.get("list", {}).get("hidden", False)
            and item.get("list", {}).get("template") == "documentLibrary"
        ]

        logger.info(
            f"Total lists: {len(result.get('value', []))}, "
            f"visible document libraries: {len(visible)}"
        )

        # Re-shape to the same dict structure read_tools.py expects
        # (identical keys to what /drives used to return).
        drives = []
        for item in visible:
            drive_info = item.get("drive") or {}
            drives.append({
                "id":          drive_info.get("id", ""),        # drive_id for folder ops
                "name":        item.get("displayName", "Unknown"),
                "description": item.get("description", "") or "No description",
                "webUrl":      item.get("webUrl", "Unknown"),
                "driveType":   drive_info.get("driveType", "documentLibrary"),
            })

        return {"value": drives}

    async def create_site(
        self, display_name: str, alias: str, description: str = ""
    ) -> Dict[str, Any]:
        """Create a new SharePoint site."""
        endpoint = "sites/root/sites"
        data = {"displayName": display_name, "alias": alias, "description": description}
        logger.info(f"Creating new site with name: {display_name}, alias: {alias}")
        return await self.post(endpoint, data)
