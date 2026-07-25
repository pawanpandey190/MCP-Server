"""SharePoint read-only tools."""

import base64
import json
import logging

from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from config.settings import SHAREPOINT_CONFIG, SITES
from tools._tool_helpers import _check_auth
from utils.document_processor import DocumentProcessor
from utils.graph_client import GraphClient

logger = logging.getLogger("sharepoint_tools")


def register_read_tools(mcp: FastMCP):
    """Register read-only SharePoint tools with the MCP server."""

    # ------------------------------------------------------------------
    # Helper: resolve site URL from an optional site_name argument
    # ------------------------------------------------------------------
    def _resolve_site_url(site_name: str | None) -> str:
        """Return the URL for *site_name*, or the default site if None."""
        if not site_name:
            return SHAREPOINT_CONFIG["site_url"]
        url = SITES.get(site_name)
        if not url:
            available = ", ".join(SITES.keys()) or "none configured"
            raise ValueError(
                f"Unknown site name '{site_name}'. "
                f"Available sites: {available}. "
                f"Call list_available_sites to see all options."
            )
        return url

    @mcp.tool()
    async def list_available_sites(ctx: Context) -> str:
        """List all SharePoint sites configured in this MCP server.

        Returns each site's display name and URL.
        Use the 'name' field as the site_name argument in other tools
        (get_site_info, list_document_libraries, search_sharepoint) to
        target a specific site.
        """
        logger.info("Tool called: list_available_sites")
        result = [
            {"name": name, "url": url} for name, url in SITES.items()
        ]
        logger.info(f"Returning {len(result)} configured sites")
        return json.dumps(result, indent=2)

    @mcp.tool()
    async def get_site_info(ctx: Context, site_name: str = None) -> str:
        """Get basic information about a SharePoint site.

        Args:
            site_name: Display name of the site to query (e.g. "Home",
                "HR Portal"). Leave empty to use the default site.
                Call list_available_sites to see all configured options.
        """
        logger.info(f"Tool called: get_site_info site_name={site_name!r}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            site_url = _resolve_site_url(site_name)
            site_parts = site_url.replace("https://", "").split("/")
            domain = site_parts[0]
            site_name_path = site_parts[2] if len(site_parts) > 2 else "root"
            logger.info(f"Getting info for site: {site_name_path} in domain: {domain}")

            site_info = await graph_client.get_site_info(domain, site_name_path)
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", site_url),
                "id": site_info.get("id", "Unknown"),
            }
            logger.info(f"Successfully retrieved site info for: {result['name']}")
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            raise

    @mcp.tool()
    async def list_document_libraries(ctx: Context, site_name: str = None) -> str:
        """List all document libraries in a SharePoint site.

        Args:
            site_name: Display name of the site to query (e.g. "Home",
                "HR Portal"). Leave empty to use the default site.
                Call list_available_sites to see all configured options.
        """
        logger.info(f"Tool called: list_document_libraries site_name={site_name!r}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            site_url = _resolve_site_url(site_name)
            site_parts = site_url.replace("https://", "").split("/")
            domain = site_parts[0]
            site_name_path = site_parts[2] if len(site_parts) > 2 else "root"
            logger.info(
                f"Listing document libraries for site: {site_name_path} in domain: {domain}"
            )

            result = await graph_client.list_document_libraries(domain, site_name_path)
            drives = result.get("value", [])

            # SharePoint returns ALL drives including hidden system libraries.
            # We filter those out so Claude only sees real, user-facing libraries.
            # System libraries have well-known names and driveType != "documentLibrary".
            SYSTEM_LIBRARY_NAMES = {
                "site assets",
                "style library",
                "form templates",
                "site pages",
                "preservation hold library",
                "content and structure reports",
                "reusable content",
                "solution gallery",
                "theme gallery",
                "master page gallery",
                "list template gallery",
                "web part gallery",
                "site collection images",
                "sharepoint lists",
            }

            formatted_drives = [
                {
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown"),
                }
                for drive in drives
                # Keep only drives whose names are NOT in the system library blocklist
                if drive.get("name", "").lower() not in SYSTEM_LIBRARY_NAMES
            ]
            logger.info(
                f"Successfully retrieved {len(formatted_drives)} document libraries"
            )
            return json.dumps(formatted_drives, indent=2)
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            raise

    @mcp.tool()
    async def search_sharepoint(
        ctx: Context,
        query: str,
        site_name: str = None,
        limit: int = 25,
        offset: int = 0,
    ) -> str:
        """Search content in a SharePoint site.

        Args:
            query: Search query string
            site_name: Display name of the site to search (e.g. "Home",
                "HR Portal"). Leave empty to search the default site.
                Call list_available_sites to see all configured options.
            limit: Maximum number of search results to return.
            offset: The offset start point for pagination.
        """
        logger.info(
            f"Tool called: search_sharepoint query={query!r} site_name={site_name!r} limit={limit} offset={offset}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            if site_name and site_name.lower() in ("global", "all"):
                site_url = None
                logger.info(f"Searching for '{query}' globally across all sites")
            else:
                site_url = _resolve_site_url(site_name)
                logger.info(f"Searching for '{query}' in site URL: {site_url}")

            # Use the global Graph Search API which correctly supports Application permissions
            # and searches across the specific site path.
            search_results = await graph_client.search_content(
                query, site_url, size=limit, from_offset=offset
            )

            formatted_results = []
            if "value" in search_results and len(search_results["value"]) > 0:
                for result in search_results["value"][0].get("hitsContainers", []):
                    for hit in result.get("hits", []):
                        resource = hit.get("resource", {})
                        parent_ref = resource.get("parentReference", {})
                        formatted_results.append(
                            {
                                "title": resource.get("name", "Unknown"),
                                "url": resource.get("webUrl", "Unknown"),
                                "type": resource.get("@odata.type", "Unknown"),
                                "item_id": resource.get("id", "Unknown"),
                                "site_id": parent_ref.get("siteId", "Unknown"),
                                "drive_id": parent_ref.get("driveId", "Unknown"),
                                "summary": hit.get("summary", "No summary available"),
                            }
                        )
            logger.info(f"Search returned {len(formatted_results)} results")
            return json.dumps(formatted_results, indent=2)
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            raise

    @mcp.tool()
    async def get_document_content(
        ctx: Context,
        site_id: str,
        drive_id: str,
        item_id: str,
        filename: str,
        start_page: int = 1,
        end_page: int = None,
    ) -> str:
        """Get and process content from a SharePoint document.

        Supports PDF, DOCX, XLSX, CSV, TXT, MD, HTML.
        Use start_page and end_page to read any range of pages (PDF),
        paragraphs (DOCX), or lines (TXT). The response always includes
        total_pages so you know how many pages/paragraphs the file has.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document
            filename: Name of the file (for content type detection)
            start_page: First page/paragraph/line to extract (1-indexed, default 1)
            end_page: Last page/paragraph/line to extract, inclusive (default: start+9
                for PDF, start+19 for DOCX, start+29 for TXT). Pass the total page
                count to read all remaining pages from start_page.
        """
        logger.info(
            f"Tool called: get_document_content for file: {filename} "
            f"pages {start_page}-{end_page or 'auto'}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content, tmp_path = await graph_client.get_document_content(
                site_id, drive_id, item_id
            )
            processed_content = DocumentProcessor.process_document(
                content, filename,
                start_page=start_page, end_page=end_page,
                file_path=tmp_path,
            )
            logger.info(f"Successfully processed document content for: {filename}")
            
            result_str = json.dumps(processed_content, indent=2)
            # Enforce MCP 1MB limit gracefully
            if len(result_str) > 750_000:
                logger.warning(f"Response too large ({len(result_str)} bytes). Returning warning to client.")
                return json.dumps({
                    "error": "Extracted content exceeds the 1MB MCP limit. Please request a smaller page/row range (e.g. 50 rows instead of 500).",
                    "original_size_bytes": len(result_str),
                    "type": processed_content.get("type", "unknown"),
                    "total_rows": processed_content.get("total_rows") or processed_content.get("total_pages")
                }, indent=2)
                
            return result_str
        except Exception as e:
            logger.error(f"Error in get_document_content: {str(e)}")
            raise

    @mcp.tool()
    async def list_folder_contents(
        ctx: Context, site_id: str, drive_id: str, folder_path: str = ""
    ) -> str:
        """List files and folders at a given path in a SharePoint document library.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            folder_path: Folder path relative to drive root (e.g. "General" or
                "Docs/2026"). Leave empty to list the root of the drive.
        """
        logger.info(f"Tool called: list_folder_contents path='{folder_path or '/'}'")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.list_folder_contents(
                site_id, drive_id, folder_path
            )
            items = result.get("value", [])
            formatted = [
                {
                    "name": item.get("name", "Unknown"),
                    "type": "folder" if "folder" in item else "file",
                    "size": item.get("size", 0),
                    "id": item.get("id", "Unknown"),
                    "web_url": item.get("webUrl", "Unknown"),
                    "last_modified": item.get("lastModifiedDateTime", "Unknown"),
                }
                for item in items
            ]
            logger.info(
                f"Successfully listed {len(formatted)} items at path '{folder_path or '/'}'"
            )
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in list_folder_contents: {str(e)}")
            raise

    @mcp.tool()
    async def get_document_by_path(
        ctx: Context,
        site_id: str,
        drive_id: str,
        file_path: str,
        filename: str,
        start_page: int = 1,
        end_page: int = None,
    ) -> str:
        """Get and process the content of a SharePoint document by its path.

        Supports PDF, DOCX, XLSX, CSV, TXT, MD, HTML.
        Use start_page and end_page to read any range of pages (PDF),
        paragraphs (DOCX), or lines (TXT). The response always includes
        total_pages so you know how many pages/paragraphs the file has.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            file_path: File path relative to drive root (e.g. "General/report.docx")
            filename: File name used to detect the document type (e.g. "report.docx")
            start_page: First page/paragraph/line to extract (1-indexed, default 1)
            end_page: Last page/paragraph/line to extract, inclusive (default: start+9
                for PDF, start+19 for DOCX, start+29 for TXT). Pass the total page
                count to read all remaining pages from start_page.
        """
        logger.info(
            f"Tool called: get_document_by_path path='{file_path}' "
            f"pages {start_page}-{end_page or 'auto'}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content, tmp_path = await graph_client.get_document_content_by_path(
                site_id, drive_id, file_path
            )
            processed_content = DocumentProcessor.process_document(
                content, filename,
                start_page=start_page, end_page=end_page,
                file_path=tmp_path,
            )
            logger.info(
                f"Successfully processed document content for path: '{file_path}'"
            )
            
            result_str = json.dumps(processed_content, indent=2)
            if len(result_str) > 750_000:
                logger.warning(f"Response too large ({len(result_str)} bytes). Returning warning to client.")
                return json.dumps({
                    "error": "Extracted content exceeds the 1MB MCP limit. Please request a smaller page/row range.",
                    "original_size_bytes": len(result_str)
                }, indent=2)
                
            return result_str
        except Exception as e:
            logger.error(f"Error in get_document_by_path: {str(e)}")
            raise

    @mcp.tool()
    async def get_item_metadata(
        ctx: Context, site_id: str, drive_id: str, item_path: str
    ) -> str:
        """Get metadata of a file or folder by its path in a SharePoint document library.

        Returns the item's id, name, size, web URL, and timestamps.
        Use the returned id with get_document_content to retrieve file content.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library (drive)
            item_path: Item path relative to drive root (e.g. "General/report.docx"
                or "General")
        """
        logger.info(f"Tool called: get_item_metadata path='{item_path}'")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            item = await graph_client.get_item_metadata_by_path(
                site_id, drive_id, item_path
            )
            result = {
                "id": item.get("id", "Unknown"),
                "name": item.get("name", "Unknown"),
                "size": item.get("size", 0),
                "web_url": item.get("webUrl", "Unknown"),
                "created_by": item.get("createdBy", {})
                .get("user", {})
                .get("displayName", "Unknown"),
                "created_datetime": item.get("createdDateTime", "Unknown"),
                "last_modified_datetime": item.get("lastModifiedDateTime", "Unknown"),
            }
            if "folder" in item:
                result["type"] = "folder"
                result["child_count"] = item["folder"].get("childCount", 0)
            elif "file" in item:
                result["type"] = "file"
                result["mime_type"] = item["file"].get("mimeType", "Unknown")

            logger.info(f"Successfully retrieved metadata for path: '{item_path}'")
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in get_item_metadata: {str(e)}")
            raise

    @mcp.tool()
    async def download_file(
        ctx: Context, site_id: str, drive_id: str, item_id: str, filename: str
    ) -> str:
        """Download a file from SharePoint and return its content as base64.

        Use this to retrieve binary files (docx, xlsx, pdf, etc.) so they can
        be edited and re-uploaded. Pair with upload_document to complete edits.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the file
            filename: Name of the file (used for logging)
        """
        logger.info(f"Tool called: download_file for file: {filename}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            content, tmp_path = await graph_client.get_document_content(
                site_id, drive_id, item_id
            )

            # Large files are streamed to disk — read back for base64 encoding
            if tmp_path:
                import os
                with open(tmp_path, "rb") as fh:
                    content = fh.read()
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

            encoded = base64.b64encode(content).decode("utf-8")
            logger.info(
                f"Successfully downloaded file: {filename} ({len(content)} bytes)"
            )
            
            result_str = json.dumps(
                {
                    "filename": filename,
                    "size_bytes": len(content),
                    "content_base64": encoded,
                },
                indent=2,
            )
            
            if len(result_str) > 750_000:
                logger.warning(f"Download response too large ({len(result_str)} bytes).")
                return json.dumps({
                    "error": f"File is too large to download directly as base64 (size: {len(content)} bytes). The MCP limit is 1MB.",
                    "suggestion": "Please use 'get_document_content' to read and analyze the file contents instead of 'download_file'.",
                    "filename": filename,
                    "size_bytes": len(content)
                }, indent=2)
                
            return result_str
        except Exception as e:
            logger.error(f"Error in download_file: {str(e)}")
            raise

    @mcp.tool()
    async def query_document_data(
        ctx: Context, site_id: str, drive_id: str, item_id: str, filename: str, sheet_name: str, query: str
    ) -> str:
        """Execute a Pandas DataFrame query on an Excel or CSV file.

        This is highly recommended for large datasets (e.g. >10,000 rows) where
        get_document_content would timeout or exceed the 1MB limit. The file is
        loaded into a Pandas DataFrame `df` on the server, and your Python `query`
        is executed against it.

        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the file
            filename: Name of the file (must be .csv, .xlsx, .xls, or .xlsb)
            sheet_name: Name of the sheet to load (or "0" for the first sheet)
            query: A valid Pandas Python expression using `df`. 
                   Examples:
                   - "df['Sales'].max()"
                   - "df.groupby('Segment')['Revenue'].sum().to_dict()"
                   - "df[df['Status'] == 'Active'].shape[0]"
        """
        logger.info(f"Tool called: query_document_data for file: {filename}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            # Stream large file to disk
            content, tmp_path = await graph_client.get_document_content(
                site_id, drive_id, item_id
            )

            # If it's a small file and wasn't streamed to disk, write it manually
            if not tmp_path:
                import tempfile
                import os
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".tmp")
                tmp.write(content)
                tmp.flush()
                tmp_path = tmp.name

            try:
                # Execute the query
                result = DocumentProcessor.query_dataframe(
                    file_path=tmp_path,
                    filename=filename,
                    sheet_name=sheet_name if not sheet_name.isdigit() else int(sheet_name),
                    query=query
                )
                
                result_str = json.dumps(result, indent=2)
                if len(result_str) > 750_000:
                    return json.dumps({
                        "error": "Query result is too large. Please aggregate the data further."
                    })
                return result_str
            finally:
                # Always cleanup the temp file
                import os
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

        except Exception as e:
            logger.error(f"Error in query_document_data: {str(e)}")
            raise

    @mcp.tool()
    async def get_lists(ctx: Context, site_id: str) -> str:
        """List all SharePoint lists (and document libraries) in a site.

        Returns list names, IDs, and templates so you can query items.
        For subsites, use the compound site ID format: siteCollectionId,webId

        Args:
            site_id: The site ID. For subsites use compound format
                (e.g. "f474e3b9-...,0b580fc8-...")
        """
        logger.info(f"Tool called: get_lists for site: {site_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.get_lists(site_id)
            lists = result.get("value", [])
            formatted = [
                {
                    "id": lst.get("id", "Unknown"),
                    "name": lst.get("displayName", "Unknown"),
                    "description": lst.get("description", ""),
                    "web_url": lst.get("webUrl", "Unknown"),
                    "template": lst.get("list", {}).get("template", "Unknown"),
                    "created": lst.get("createdDateTime", "Unknown"),
                    "last_modified": lst.get("lastModifiedDateTime", "Unknown"),
                }
                for lst in lists
            ]
            logger.info(f"Successfully retrieved {len(formatted)} lists")
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in get_lists: {str(e)}")
            raise

    @mcp.tool()
    async def get_list_items(
        ctx: Context,
        site_id: str,
        list_id: str,
        top: int = 100,
        filter_query: str = "",
    ) -> str:
        """Get items from a SharePoint list with all their column/field values.

        This queries the Graph API /sites/{siteId}/lists/{listId}/items endpoint
        with $expand=fields to return all column values including Status, Title, etc.

        Args:
            site_id: The site ID. For subsites use compound format
                (e.g. "f474e3b9-...,0b580fc8-...")
            list_id: The list ID (GUID) or display name
            top: Maximum items to return (default 100, max 5000)
            filter_query: Optional OData filter (e.g. "fields/Status eq 'Active'")
        """
        logger.info(f"Tool called: get_list_items for list: {list_id}")
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            result = await graph_client.get_list_items(
                site_id, list_id, top=top, filter_query=filter_query
            )
            items = result.get("value", [])
            formatted = [
                {
                    "id": item.get("id", "Unknown"),
                    "web_url": item.get("webUrl", ""),
                    "created": item.get("createdDateTime", ""),
                    "last_modified": item.get("lastModifiedDateTime", ""),
                    "fields": item.get("fields", {}),
                }
                for item in items
            ]
            logger.info(f"Successfully retrieved {len(formatted)} list items")
            return json.dumps(formatted, indent=2)
        except Exception as e:
            logger.error(f"Error in get_list_items: {str(e)}")
            raise

    @mcp.tool()
    async def search_site_content(
        ctx: Context,
        query: str,
        site_name: str = None,
        limit: int = 25,
        offset: int = 0,
    ) -> str:
        """Search the deep content of all documents and lists in a SharePoint site.
        
        Unlike search_sharepoint which only matches filenames, this uses Microsoft Search
        to perform full-text search inside the actual contents of PDFs, Word docs, etc.

        Args:
            query: The search keywords to look for inside files.
            site_name: Display name of the site to query (e.g. "Finance"). 
                Leave empty to use the default site. Call list_available_sites 
                to see all configured options.
            limit: Maximum number of search results to return.
            offset: The offset start point for pagination.
        """
        logger.info(
            f"Tool called: search_site_content query={query!r} site_name={site_name!r} limit={limit} offset={offset}"
        )
        try:
            sp_ctx = ctx.request_context.lifespan_context
            _check_auth(sp_ctx)
            await refresh_token_if_needed(sp_ctx)
            graph_client = GraphClient(sp_ctx)

            if site_name and site_name.lower() in ("global", "all"):
                site_url = None
                logger.info(f"Searching content globally across all sites")
            else:
                site_url = _resolve_site_url(site_name)
                logger.info(f"Searching content in site URL: {site_url}")

            # Reuse the Graph Search client helper to support dynamic region & pagination parameters
            result = await graph_client.search_content(
                query, site_url, size=limit, from_offset=offset
            )

            # Parse the nested hitsContainers structure
            items = []
            if "value" in result and len(result["value"]) > 0:
                hits_containers = result["value"][0].get("hitsContainers", [])
                for container in hits_containers:
                    for hit in container.get("hits", []):
                        resource = hit.get("resource", {})
                        parent_ref = resource.get("parentReference", {})
                        items.append({
                            "name": resource.get("name", "Unknown"),
                            "summary": hit.get("summary", ""),
                            "web_url": resource.get("webUrl", ""),
                            "item_id": resource.get("id", "Unknown"),
                            "site_id": parent_ref.get("siteId", "Unknown"),
                            "drive_id": parent_ref.get("driveId", "Unknown"),
                            "last_modified": resource.get("lastModifiedDateTime", ""),
                            "created_by": resource.get("createdBy", {}).get("user", {}).get("displayName", "Unknown")
                        })

            logger.info(f"Successfully retrieved {len(items)} content search results")
            return json.dumps(items, indent=2)
        except Exception as e:
            logger.error(f"Error in search_site_content: {str(e)}")
            raise

