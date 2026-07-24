#!/bin/bash

# Parse env file for secrets
TENANT_ID=$(grep SHAREPOINT_TENANT_ID /Users/pawanpandey/Documents/SP-MCP/sharepoint-mcp/envs/mcp-rd.env | cut -d'=' -f2 | tr -d '"'\')
CLIENT_ID=$(grep SHAREPOINT_CLIENT_ID /Users/pawanpandey/Documents/SP-MCP/sharepoint-mcp/envs/mcp-rd.env | cut -d'=' -f2 | tr -d '"'\')
CLIENT_SECRET=$(grep SHAREPOINT_CLIENT_SECRET /Users/pawanpandey/Documents/SP-MCP/sharepoint-mcp/envs/mcp-rd.env | cut -d'=' -f2 | tr -d '"'\')

# Get token
TOKEN_RES=$(curl -s -X POST "https://login.microsoftonline.com/$TENANT_ID/oauth2/v2.0/token" \
  -d "client_id=$CLIENT_ID" \
  -d "scope=https://graph.microsoft.com/.default" \
  -d "client_secret=$CLIENT_SECRET" \
  -d "grant_type=client_credentials")

TOKEN=$(echo $TOKEN_RES | grep -o '"access_token":"[^"]*' | cut -d'"' -f4)

# Test search without region
echo "=== NO REGION ==="
curl -s -X POST "https://graph.microsoft.com/v1.0/search/query" \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"requests":[{"entityTypes":["driveItem","listItem"],"query":{"queryString":"SHYAM SUNDER"}}]}'

echo -e "\n=== WITH REGION (IND) ==="
curl -s -X POST "https://graph.microsoft.com/v1.0/search/query" \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"requests":[{"entityTypes":["driveItem","listItem"],"query":{"queryString":"SHYAM SUNDER"},"region":"IND"}]}'

echo -e "\n=== WITH REGION (NAM) ==="
curl -s -X POST "https://graph.microsoft.com/v1.0/search/query" \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"requests":[{"entityTypes":["driveItem","listItem"],"query":{"queryString":"SHYAM SUNDER"},"region":"NAM"}]}'
