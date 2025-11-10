#!/bin/bash

# OpenDocx NuGet Package Publishing Script
# This script helps publish the OpenDocx package to NuGet.org

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}OpenDocx NuGet Package Publisher${NC}"
echo "=================================="

# Check if API key is provided
if [ -z "$1" ]; then
    echo -e "${RED}Error: NuGet API key is required${NC}"
    echo "Usage: $0 <nuget-api-key>"
    echo ""
    echo "To get your API key:"
    echo "1. Go to https://www.nuget.org"
    echo "2. Sign in to your account"
    echo "3. Go to Account Settings > API Keys"
    echo "4. Create a new API key with 'Push new packages and package versions' scope"
    exit 1
fi

API_KEY=$1
PACKAGE_PATH="./bin/Release/OpenDocx.NET.1.0.2.nupkg"

# Check if package exists
if [ ! -f "$PACKAGE_PATH" ]; then
    echo -e "${RED}Error: Package file not found at $PACKAGE_PATH${NC}"
    echo "Run 'dotnet pack -c Release' first to create the package."
    exit 1
fi

echo -e "${YELLOW}Package found: $PACKAGE_PATH${NC}"
echo -e "${YELLOW}Publishing to NuGet.org...${NC}"

# Publish the package
dotnet nuget push "$PACKAGE_PATH" --api-key "$API_KEY" --source https://api.nuget.org/v3/index.json

if [ $? -eq 0 ]; then
    echo -e "${GREEN}✓ Package published successfully!${NC}"
    echo ""
    echo "Your package should be available at:"
    echo "https://www.nuget.org/packages/OpenDocx/"
    echo ""
    echo "Note: It may take a few minutes for the package to appear in search results."
else
    echo -e "${RED}✗ Failed to publish package${NC}"
    exit 1
fi