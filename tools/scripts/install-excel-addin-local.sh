#!/bin/bash

echo "========================================"
echo " OpenGov Office Sync - LOCAL TESTING"
echo "========================================"
echo ""
echo "Installing Excel add-in for local development..."
echo ""

# Close Excel if running
echo "Closing Excel if running..."
osascript -e 'quit app "Microsoft Excel"' 2>/dev/null
sleep 2

# Create wef directory if it doesn't exist
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$WEF_DIR"

# Copy local manifest
echo "Setting up local manifest..."
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
MANIFEST_PATH="$WEF_DIR/opengov-office-sync-local.xml"

cp "$SCRIPT_DIR/manifest-local.xml" "$MANIFEST_PATH"

if [ $? -ne 0 ]; then
  echo ""
  echo "ERROR: Failed to copy manifest"
  exit 1
fi

echo "✓ Manifest installed"
echo ""
echo "========================================"
echo " Installation Complete!"
echo "========================================"
echo ""
echo "IMPORTANT: Make sure your local servers are running!"
echo "  - Backend: http://localhost:3001"
echo "  - Frontend: http://localhost:3000"
echo ""
echo "Run this command in another terminal:"
echo "  npm run server"
echo ""
echo "Opening Excel..."
open -a "Microsoft Excel"

echo ""
echo "========================================"
echo " Next Steps: Activate the Add-in"
echo "========================================"
echo ""
echo "TO ACTIVATE THE ADD-IN (first time only):"
echo "  1. In Excel, click the 'Insert' tab"
echo "  2. Click 'Add-ins' → 'My Add-ins'"
echo "  3. Under 'Developer Add-ins', click 'OpenGov Office Sync (Local)'"
echo ""
echo "The add-in panel will appear on the right side."
echo ""
echo "You can now open: http://localhost:3000"
echo "to see real-time sync between Excel and the web!"
echo ""

exit 0

