# Local Testing Guide

## 🚀 Quick Start

### 1. Start the Servers

```bash
# Start backend (port 3001)
npm run server

# In another terminal, start frontend (port 3000)
npm start
```

Wait for both servers to start completely (~30 seconds).

---

## 📥 Install the Add-in (Local)

### Windows

1. **Run the local installer**:
   ```
   tools\scripts\install-excel-addin-local.bat
   ```
   - Right-click → "Run as Administrator" (recommended)
   - Or just double-click

2. **Excel will open automatically**

3. **Activate the add-in** (first time only):
   - Click **Insert** tab
   - Click **Get Add-ins** or **My Add-ins**
   - Click **Developer Add-ins** at the top
   - Click **"OpenGov Office Sync (Local)"**

4. The add-in panel appears on the right! ✨

### Mac

1. **Make script executable**:
   ```bash
   chmod +x tools/scripts/install-excel-addin-local.sh
   ```

2. **Run the local installer**:
   ```bash
   ./tools/scripts/install-excel-addin-local.sh
   ```

3. **Excel will open automatically**

4. **Activate the add-in** (first time only):
   - Click **Insert** tab
   - Click **Add-ins** → **My Add-ins**
   - Under **Developer Add-ins**, click **"OpenGov Office Sync (Local)"**

5. The add-in panel appears on the right! ✨

---

## ✅ Test the Sync

### In Excel:
1. Type some data in cells (A1, A2, B1, etc.)
2. Watch the status indicator: "✓ Synced"

### In Web Browser:
1. Open http://localhost:3000
2. You should see the data from Excel appear!

### In Web:
1. Click cells and type data
2. Watch Excel update in real-time!

### Test Budget Book Feature:
1. Click "Update my budget book" in the sidepane
2. Success modal should appear

---

## 🗑️ Uninstall

### Windows

```
tools\scripts\uninstall-excel-addin.bat
```

### Mac

```bash
chmod +x tools/scripts/uninstall-excel-addin.sh
./tools/scripts/uninstall-excel-addin.sh
```

**Important**: Always close Excel before installing/uninstalling!

**Note**: The uninstall script automatically detects and removes both local and production versions.

---

## 🐛 Troubleshooting

### Add-in doesn't appear in Excel

**1. Check servers are running**:
- Backend: http://localhost:3001/api/health (should return JSON)
- Frontend: http://localhost:3000 (should load webpage)

**2. Clear Office cache**:

**Windows**:
```
tools\scripts\uninstall-excel-addin.bat
```
Then reinstall.

**Mac**:
```bash
tools\scripts\uninstall-excel-addin.sh
```
Then reinstall.

**3. Check registry (Windows)**:
```
reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
```
Should show `opengov-excel-addin-local` entry.

**4. Verify manifest path**:

**Windows**: `%TEMP%\opengov-excel-addin\manifest.xml`  
**Mac**: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/opengov-office-sync-local.xml`

File should exist and contain localhost URLs.

### Sync not working

**1. Check browser console** (F12):
- Should see "Loaded data from server"
- Should see "Ready for syncing"
- No errors about connections

**2. Check Excel console** (if visible):
- Should see "Ready"
- Should see "Syncing to server" when editing

**3. Check network**:
- Open: http://localhost:3001/api/stream
- Should see SSE connection (don't close, just verify)

**4. Check MongoDB**:
- Make sure MongoDB is running
- Default: `mongodb://localhost:27017/opengov-office`

### Icons not loading

Run the icon generator:
```bash
node tools/scripts/create-icons.js
```

Check files exist:
- `web/assets/icon-16.png`
- `web/assets/icon-32.png`
- `web/assets/icon-64.png`
- `web/assets/icon-80.png`

### "Cannot connect to server"

1. Verify server is running on port 3001:
   ```bash
   # Windows
   netstat -ano | findstr 3001
   
   # Mac
   lsof -i :3001
   ```

2. Check firewall isn't blocking localhost

3. Try accessing directly:
   - http://localhost:3001/api/health

---

## 🔄 Reinstall from Scratch

```bash
# 1. Uninstall (removes any version automatically)
./tools/scripts/uninstall-excel-addin.bat  # Windows
./tools/scripts/uninstall-excel-addin.sh   # Mac

# 2. Close Excel completely
# (check Task Manager / Activity Monitor)

# 3. Kill servers
# Ctrl+C in terminal windows

# 4. Clear data (optional - resets database)
rm -rf data/  # Mac/Linux
rmdir /s data  # Windows

# 5. Start servers
npm run server
npm start

# 6. Reinstall (choose local or production)
./tools/scripts/install-excel-addin-local.bat  # Windows Local
./tools/scripts/install-excel-addin-local.sh   # Mac Local
./tools/scripts/install-excel-addin.bat        # Windows Production
./tools/scripts/install-excel-addin.sh         # Mac Production
```

---

## 📊 What Should Happen

### Successful Installation:
```
✓ Excel opens
✓ Insert → My Add-ins → Developer Add-ins shows "OpenGov Office Sync (Local)"
✓ Click it → Panel appears on right
✓ Status shows "✓ Live sync" with green dot
✓ Manual Sync button visible
✓ Budget Book section with book icon
```

### Successful Sync:
```
Excel → Type "Hello" in A1
  ↓
Status: "✓ Synced"
  ↓
Web (localhost:3000) shows "Hello" in first cell
  ↓
Web → Type "World" in A2
  ↓
Excel cell A2 updates to "World"
```

---

## 🎯 Next Steps

Once local testing works:
1. Update `manifest.xml` with production URLs
2. Deploy to Render
3. Test with `install-excel-addin.bat` (production version)

---

## 📝 File Reference

**Manifests**:
- `manifest-local.xml` - Manifest pointing to localhost
- `manifest.xml` - Manifest pointing to Render (production)

**Installers**:
- `tools/scripts/install-excel-addin-local.bat` - Windows installer (localhost)
- `tools/scripts/install-excel-addin-local.sh` - Mac installer (localhost)
- `tools/scripts/install-excel-addin.bat` - Windows installer (production)
- `tools/scripts/install-excel-addin.sh` - Mac installer (production)

**Uninstallers** (work for both local and production):
- `tools/scripts/uninstall-excel-addin.bat` - Windows uninstaller
- `tools/scripts/uninstall-excel-addin.sh` - Mac uninstaller

---

Need help? Check the console logs in both Excel and browser (F12). 🕵️

