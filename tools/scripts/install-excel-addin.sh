#!/bin/bash

echo "========================================"
echo " OpenGov Office Sync - Excel Add-in"
echo "========================================"
echo ""
echo "Installing Excel add-in for Mac..."
echo ""

# Create wef directory if it doesn't exist
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$WEF_DIR"

# Remove local version if present
echo "Checking for local version..."
rm -f "$WEF_DIR/manifest-local.xml" 2>/dev/null
echo "   > Local version removed (if present)"
echo ""

# Download manifest
echo "Downloading manifest..."
MANIFEST_URL="https://excelftw.onrender.com/manifest.xml"
MANIFEST_PATH="$WEF_DIR/opengov-office-sync.xml"

curl -L -o "$MANIFEST_PATH" "$MANIFEST_URL"

if [ $? -ne 0 ]; then
  echo ""
  echo "ERROR: Failed to download manifest"
  echo "Please check your internet connection"
  exit 1
fi

echo "✓ Manifest installed"
echo ""
echo "========================================"
echo " Installation Complete!"
echo "========================================"
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
echo "  3. Under 'Developer Add-ins', click 'OpenGov Office Sync'"
echo ""
echo "The add-in panel will appear on the right side."
echo ""
echo "You can now open: https://excelftw.onrender.com"
echo "to see real-time sync between Excel and the web!"
echo ""

exit 0

