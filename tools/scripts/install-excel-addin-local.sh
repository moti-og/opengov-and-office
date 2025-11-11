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

# Remove production version if present
echo "Checking for production version..."
rm -f "$WEF_DIR/opengov-office-sync.xml" 2>/dev/null
echo "   > Production version removed (if present)"
echo ""

# Download local manifest from server
echo "Downloading local manifest from localhost:3001..."
MANIFEST_URL="http://localhost:3001/manifest-local.xml"
MANIFEST_PATH="$WEF_DIR/opengov-office-sync-local.xml"

curl -L -o "$MANIFEST_PATH" "$MANIFEST_URL"

if [ $? -ne 0 ]; then
  echo ""
  echo "ERROR: Failed to download manifest from server"
  echo "Make sure the backend server is running:"
  echo "  npm run server"
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

