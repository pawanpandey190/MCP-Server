import asyncio
import logging
import sys
import json

# Add the app directory to path
sys.path.append("/Users/pawanpandey/Documents/SP-MCP/sharepoint-mcp")

from auth.sharepoint_auth import get_auth_context
from utils.graph_client import GraphClient

logging.basicConfig(level=logging.INFO, format="%(levelname)s - %(message)s")

async def test_highlight_search():
    try:
        ctx = await get_auth_context()
        c = GraphClient(ctx)
        
        print("Testing Search with HitHighlightedSummary field...")
        
        # Construct payload with custom fields requesting HitHighlightedSummary
        payload = {
            "requests": [
                {
                    "entityTypes": ["driveItem"],
                    "query": {
                        "queryString": "shyam sunder"
                    },
                    "fields": [
                        "id",
                        "name",
                        "webUrl",
                        "HitHighlightedSummary"
                    ],
                    "size": 3
                }
            ]
        }
        
        res = await c.post("search/query", payload)
        hits = res.get("value", [{}])[0].get("hitsContainers", [{}])[0].get("hits", [])
        for idx, hit in enumerate(hits, 1):
            resource = hit.get("resource", {})
            print(f"\n--- HIT {idx} ---")
            print("Name:", resource.get("name"))
            print("WebUrl:", resource.get("webUrl"))
            print("HitHighlightedSummary:", resource.get("HitHighlightedSummary"))
            print("Raw Resource Keys:", list(resource.keys()))
            print("Hit summary property:", hit.get("summary"))
    except Exception as e:
        print(f"Error: {e}")

asyncio.run(test_highlight_search())
