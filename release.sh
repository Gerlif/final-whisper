#!/bin/bash
set -e

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  Final Whisper Release Creator${NC}"
echo -e "${BLUE}========================================${NC}\n"

# Check if gh CLI is installed
if ! command -v gh &> /dev/null; then
    echo -e "${RED}Error: GitHub CLI (gh) is not installed${NC}"
    echo "Install it with: sudo apt install gh"
    echo "Then authenticate with: gh auth login"
    exit 1
fi

# Check if authenticated
if ! gh auth status &> /dev/null; then
    echo -e "${RED}Error: Not authenticated with GitHub${NC}"
    echo "Run: gh auth login"
    exit 1
fi

# Read version from version.txt
if [ ! -f "version.txt" ]; then
    echo -e "${RED}Error: version.txt not found${NC}"
    exit 1
fi

VERSION=$(cat version.txt | tr -d '[:space:]')
TAG="v$VERSION"
TITLE="Final Whisper v$VERSION"

echo -e "${GREEN}Current version:${NC} $VERSION"
echo -e "${GREEN}Release tag:${NC} $TAG"
echo -e "${GREEN}Release title:${NC} $TITLE"
echo ""

# Check if tag already exists
if gh release view "$TAG" &> /dev/null; then
    echo -e "${RED}Error: Release $TAG already exists${NC}"
    echo "Please increment version.txt first or delete the existing release"
    exit 1
fi

# Prompt for release notes
echo -e "${YELLOW}Enter release notes (or press Enter for default):${NC}"
read -p "> " RELEASE_NOTES

if [ -z "$RELEASE_NOTES" ]; then
    RELEASE_NOTES="Release of Final Whisper v$VERSION

## Features
- AI-powered transcription using OpenAI Whisper
- Smart subtitle formatting with sentence boundary detection
- Optional AI proofreading with Claude
- GPU acceleration support
- Multi-language support (optimized for Danish)

## Installation
Download the EXE and run it. You'll need to install Whisper:
\`\`\`bash
pip install openai-whisper
\`\`\`

For GPU support:
\`\`\`bash
pip install torch --index-url https://download.pytorch.org/whl/cu118
\`\`\`"
fi

echo ""
echo -e "${YELLOW}Creating release...${NC}"

# Create the release
gh release create "$TAG" \
    --title "$TITLE" \
    --notes "$RELEASE_NOTES"

echo ""
echo -e "${GREEN}✓ Release created successfully!${NC}"
echo -e "${BLUE}The build workflow has been triggered automatically.${NC}"
echo -e "${BLUE}The EXE will be attached to the release when the build completes.${NC}"
echo ""
echo -e "View release: ${GREEN}https://github.com/$GITHUB_REPO/releases/tag/$TAG${NC}"
echo -e "Monitor build: ${GREEN}https://github.com/$GITHUB_REPO/actions${NC}"
echo ""
