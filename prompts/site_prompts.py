"""SharePoint site prompts."""

from mcp.server.fastmcp import FastMCP


def register_site_prompts(mcp: FastMCP):
    """Register SharePoint prompts with the MCP server."""

    @mcp.prompt()
    def search_and_analyze(query: str, site_name: str = "") -> str:
        """Search a SharePoint site for a topic, read relevant documents, and summarize the findings.

        Args:
            query: The search query (e.g., "quarterly revenue", "HR policies").
            site_name: The display name of the SharePoint site (e.g., "Finance", "HR").
        """
        site_context = f" on the site '{site_name}'" if site_name else ""
        return (
            f"You are a SharePoint analyst assistant. Your task is to investigate '{query}'{site_context}.\n\n"
            f"Please follow these steps:\n"
            f"1. Use `search_site_content` to find documents related to '{query}'{site_context}.\n"
            f"2. For the top 2-3 most relevant documents (e.g., PDF, DOCX, XLSX), read their text contents using `get_document_content` (or `get_document_by_path`).\n"
            f"3. Generate a structured summary of the findings, including file names, URLs, and key takeaways."
        )

    @mcp.prompt()
    def site_audit(site_name: str) -> str:
        """Audit a specific SharePoint site by checking document libraries and listing contents.

        Args:
            site_name: The display name of the SharePoint site to audit (e.g., "Finance", "HR").
        """
        return (
            f"You are performing an audit of the SharePoint site '{site_name}'.\n\n"
            f"Please complete these steps:\n"
            f"1. Use `list_available_sites` to verify the site exists and check its details.\n"
            f"2. Call `list_document_libraries` with `site_name='{site_name}'` to retrieve all document libraries.\n"
            f"3. For the primary libraries, browse their root contents using `list_folder_contents`.\n"
            f"4. Provide a report of the site structure, including libraries found and count/types of files in the root folders."
        )
