"""Microsoft Graph Search API operations mixin."""

import logging
from typing import Dict, Any, List, Optional

logger = logging.getLogger("graph_client")


class _GraphSearchOpsMixin:
    """Methods for interacting with Microsoft Graph Search API."""

    async def search_content(
        self,
        query_string: str,
        site_url: Optional[str] = None,
        size: int = 25,
        from_offset: int = 0,
    ) -> Dict[str, Any]:
        """Perform a full-text search across SharePoint using Microsoft Graph Search API.
        
        Args:
            query_string: The search keyword or KQL query.
            site_url: Optional specific site URL to restrict the search.
            size: Number of search results to return.
            from_offset: The offset start point for pagination.
            
        Returns:
            The search response JSON from Graph API.
        """
        # Clean the input query from quotes that might break KQL syntax
        clean_query = query_string.replace('"', '').replace("'", "").strip()
        
        # Build optimized KQL query: match terms in any order (high recall),
        # but boost exact phrase matches (high precision), and exclude system filetypes
        kql_query = f'({clean_query}) XRANK(cb=100.0) ("{clean_query}" OR Title:"{clean_query}" OR filename:"{clean_query}")'
        kql_query += ' NOT FileType:lnk NOT FileType:aspx'
        
        # If site_url is provided, restrict search to that specific site path
        if site_url:
            kql_query = f'({kql_query}) AND Path:"{site_url}"'

        query_string = kql_query

        endpoint = "search/query"
        payload = {
            "requests": [
                {
                    "entityTypes": ["driveItem", "listItem"],
                    "query": {
                        "queryString": query_string
                    },
                    "from": from_offset,
                    "size": size,
                    "trimDuplicates": True,
                }
            ]
        }

        from config.settings import SEARCH_REGION
        if SEARCH_REGION:
            payload["requests"][0]["region"] = SEARCH_REGION

        logger.info(f"Executing search API with query: {query_string} (region: {SEARCH_REGION}, size: {size}, offset: {from_offset})")
        return await self.post(endpoint, payload)

