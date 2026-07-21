#!/bin/bash
# ==============================================================================
# SharePoint MCP - Azure VM Deployment Script
# Run this script (sudo ./deploy.sh) to fetch the latest code and start the containers.
# ==============================================================================

set -e # Exit immediately if a command exits with a non-zero status

# 1. Check if running as root
if [ "$EUID" -ne 0 ]; then
  echo "❌ Please run as root (use sudo ./deploy.sh)"
  exit 1
fi

REPO_DIR="/opt/sharepoint-mcp"
REPO_URL="https://github.com/pawanpandey190/MCP-Server.git"

echo "🚀 Starting SharePoint MCP Deployment..."

# 2. Update system and install basic prerequisites (git)
echo "📦 Installing prerequisites..."
apt-get update -y
apt-get install -y ca-certificates curl gnupg git

# 3. Clone or update the repository
if [ -d "$REPO_DIR" ]; then
    echo "📥 Repository exists. Pulling latest changes..."
    cd $REPO_DIR
    git pull origin main
else
    echo "📥 Cloning repository to $REPO_DIR..."
    git clone $REPO_URL $REPO_DIR
    cd $REPO_DIR
fi

# 4. Check for .env files
echo "🔐 Checking for .env files..."
mkdir -p envs
if [ ! -f "envs/mcp-finance.env" ]; then
    echo "⚠️  WARNING: Secrets are missing!"
    echo "You must manually create the .env files in $REPO_DIR/envs/"
    echo "Example: nano $REPO_DIR/envs/mcp-finance.env"
    echo "After creating the secrets, run: cd $REPO_DIR && docker compose up -d --build"
    exit 1
fi

# 5. Start the Docker Containers
echo "🐳 Building and starting Docker containers..."
docker compose up -d --build

echo "========================================================================="
echo "🎉 Deployment Complete! All 11 MCP containers are now running locally."
echo "Your custom application can now access them at http://localhost:8001 through 8011."
echo "========================================================================="
