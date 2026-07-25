import asyncio
import logging
import sys

# Add the app directory to path
sys.path.append("/app")

from auth.sharepoint_auth import get_auth_context
from utils.graph_client import GraphClient
from config.settings import SITES, SEARCH_REGION

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")

async def test_container_tools():
    print("=== DOCKER MCP APPLICATION DIAGNOSTIC ===")
    print(f"SEARCH_REGION configured as: {SEARCH_REGION!r}")
    print(f"Sites configured in this container: {list(SITES.keys())}\n")
    
    try:
        print("[1] Authenticating with Microsoft Graph...")
        ctx = await get_auth_context()
        client = GraphClient(ctx)
        print("✅ Authentication successful. Token acquired.\n")
        
        for site_name, site_url in SITES.items():
            print(f"--- Testing for Site: {site_name} ({site_url}) ---")
            
            # Test A: Resolve site
            print("  [A] Fetching site metadata...")
            try:
                # Resolve site path
                from utils.graph_client import parse_sharepoint_url
                hostname, site_path = parse_sharepoint_url(site_url)
                site_metadata = await client.get(f"sites/{hostname}:/{site_path}")
                site_id = site_metadata.get("id")
                print(f"  ✅ Site resolved successfully. Site ID: {site_id}")
            except Exception as e:
                print(f"  ❌ Failed to resolve site: {e}")
                continue

            # Test B: List libraries (drives)
            print("  [B] Fetching document libraries...")
            try:
                drives = await client.get_drives(site_id)
                print(f"  ✅ Found {len(drives)} document libraries.")
            except Exception as e:
                print(f"  ❌ Failed to fetch libraries: {e}")

            # Test C: Search Content (using dynamic region)
            print("  [C] Testing Search API...")
            try:
                results = await client.search_content(query="test", site_url=site_url, size=3, from_offset=0)
                # Parse hits
                hits = []
                if "value" in results and len(results["value"]) > 0:
                    hits = results["value"][0].get("hitsContainers", [{}])[0].get("hits", [])
                print(f"  ✅ Search API succeeded. Found {len(hits)} results for query 'test'.")
            except Exception as e:
                print(f"  ❌ Search API failed: {e}")
                
            print()
            
    except Exception as e:
        print(f"❌ Critical Error: {e}")

if __name__ == "__main__":
    asyncio.run(test_container_tools())
