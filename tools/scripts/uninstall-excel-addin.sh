#!/bin/bash

echo "========================================"
echo " OpenGov Office Sync - Uninstaller"
echo "========================================"
echo ""

# Close Excel if running
echo "[1/3] Closing Excel..."
osascript -e 'quit app "Microsoft Excel"' 2>/dev/null
sleep 2

# Remove EXCEL-ONLY OpenGov manifest files
echo "[2/3] Removing OpenGov EXCEL add-ins..."
WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
if [ -d "$WEF_DIR" ]; then
    find "$WEF_DIR" -type f -name "*.xml" 2>/dev/null | while read file; do
        if grep -qi "opengov" "$file" 2>/dev/null; then
            # Check if it's an Excel add-in (Host Name="Workbook")
            if grep -q 'Host.*Name="Workbook"' "$file" 2>/dev/null; then
                echo "   Removing Excel add-in: $(basename "$file")"
                rm -f "$file"
            else
                echo "   Skipping non-Excel add-in: $(basename "$file")"
            fi
        fi
    done
fi

# Clear Office cache
echo "[3/3] Clearing cache..."
if [ -d "$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Caches" ]; then
    rm -rf "$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Caches"/* 2>/dev/null
fi

echo ""
echo "========================================"
echo " Done!"
echo "========================================"
echo ""
echo "All OpenGov EXCEL add-ins have been removed."
echo "Word add-ins were left untouched."
echo ""
echo "Open Excel to verify it's gone:"
echo "  Insert > Add-ins > My Add-ins > Developer Add-ins"
echo ""
echo "If it's still there, restart Excel or reboot."
echo ""

read -p "Press Enter to close..."
exit 0

