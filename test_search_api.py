import asyncio
import os
import json
from auth.sharepoint_auth import SharePointContext
from utils.graph_client import GraphClient

async def test():
    # Setup context (assuming env vars are set or hardcoded test token)
    context = SharePointContext(
        client_id=os.environ.get("SHAREPOINT_CLIENT_ID"),
        client_secret=os.environ.get("SHAREPOINT_CLIENT_SECRET"),
        tenant_id=os.environ.get("SHAREPOINT_TENANT_ID"),
        site_url="https://nbcbearingsrj.sharepoint.com/sites/RD"
    )
    from auth.sharepoint_auth import refresh_token_if_needed
    await refresh_token_if_needed(context)
    
    client = GraphClient(context)
    try:
        res = await client.search_content("SHYAM SUNDER")
        print(json.dumps(res, indent=2))
    except Exception as e:
        print("ERROR:", str(e))

asyncio.run(test())
