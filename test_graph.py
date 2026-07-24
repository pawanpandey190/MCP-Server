import requests
import json

# Read env file
env = {}
with open("/Users/pawanpandey/Documents/SP-MCP/sharepoint-mcp/envs/mcp-rd.env") as f:
    for line in f:
        if "=" in line:
            k, v = line.strip().split("=", 1)
            env[k] = v.strip("'\"")

token_url = f"https://login.microsoftonline.com/{env['SHAREPOINT_TENANT_ID']}/oauth2/v2.0/token"
data = {
    "client_id": env["SHAREPOINT_CLIENT_ID"],
    "scope": "https://graph.microsoft.com/.default",
    "client_secret": env["SHAREPOINT_CLIENT_SECRET"],
    "grant_type": "client_credentials"
}
res = requests.post(token_url, data=data)
token = res.json()["access_token"]

payload = {
    "requests": [
        {
            "entityTypes": ["driveItem", "listItem"],
            "query": {"queryString": "SHYAM SUNDER"}
        }
    ]
}
res = requests.post(
    "https://graph.microsoft.com/v1.0/search/query",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json=payload
)
print("WITHOUT REGION:", res.status_code, res.text)

payload["requests"][0]["region"] = "IND"
res = requests.post(
    "https://graph.microsoft.com/v1.0/search/query",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json=payload
)
print("WITH REGION (IND):", res.status_code, res.text)

payload["requests"][0]["region"] = "NAM"
res = requests.post(
    "https://graph.microsoft.com/v1.0/search/query",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json=payload
)
print("WITH REGION (NAM):", res.status_code, res.text)
