"""Microsoft Graph Search API operations mixin."""

import logging
from typing import Dict, Any, List, Optional

logger = logging.getLogger("graph_client")


class _GraphSearchOpsMixin:
    """Methods for interacting with Microsoft Graph Search API."""

    async def search_content(
        self, query_string: str, site_url: Optional[str] = None
    ) -> Dict[str, Any]:
        """Perform a full-text search across SharePoint using Microsoft Graph Search API.
        
        Args:
            query_string: The search keyword or KQL query.
            site_url: Optional specific site URL to restrict the search.
            
        Returns:
            The search response JSON from Graph API.
        """
        # If site_url is provided, restrict search to that specific site path
        if site_url:
            query_string = f'{query_string} Path:"{site_url}"'

        endpoint = "search/query"
        payload = {
            "requests": [
                {
                    "entityTypes": ["driveItem", "listItem"],
                    "query": {
                        "queryString": query_string
                    }
                }
            ]
        }

        logger.info(f"Executing search API with query: {query_string}")
        return await self.post(endpoint, payload)
