# Deployment Guide - Excel Add-in on Render

## 📋 Pre-Deployment Checklist

### 1. Update manifest.xml URLs

Replace `opengov-office.onrender.com` with your actual Render domain:

```xml
<!-- In manifest.xml, update all URLs -->
<SourceLocation DefaultValue="https://YOUR-APP.onrender.com/addin/excel/taskpane/taskpane.html"/>
<IconUrl DefaultValue="https://YOUR-APP.onrender.com/assets/icon-32.png"/>
```

### 2. Update installer scripts

**install-excel-addin.bat** (line 18):
```batch
powershell -Command "Invoke-WebRequest -Uri 'https://YOUR-APP.onrender.com/manifest.xml' -OutFile '%TEMP_DIR%\manifest.xml'"
```

**install-excel-addin.sh** (line 16):
```bash
MANIFEST_URL="https://YOUR-APP.onrender.com/manifest.xml"
```

### 3. Create icon files

Create a folder `web/assets/` with these icon files:
- `icon-16.png` (16x16px)
- `icon-32.png` (32x32px)
- `icon-64.png` (64x64px)
- `icon-80.png` (80x80px)

Simple OpenGov logo or "OG" text works fine for prototype.

---

## 🚀 Deploy to Render

### Option A: Connect to GitHub

1. Push your code to GitHub
2. Go to [Render Dashboard](https://dashboard.render.com/)
3. Click "New +" → "Web Service"
4. Connect your GitHub repo
5. Configure:
   - **Name**: `opengov-office`
   - **Environment**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm run server`
   - **Port**: `3001` (set in Environment Variables)

### Option B: Deploy from CLI

```bash
# Install Render CLI
npm install -g @render-cli/cli

# Login
render login

# Deploy
render deploy
```

---

## 🔧 Environment Variables on Render

Add these in Render Dashboard → Environment:

```
MONGODB_URI=mongodb+srv://your-atlas-uri
SERVER_PORT=3001
NODE_ENV=production
```

**For MongoDB Atlas:**
1. Go to [MongoDB Atlas](https://www.mongodb.com/cloud/atlas)
2. Create free cluster
3. Get connection string
4. Add to `MONGODB_URI` env var

---

## ✅ Post-Deployment

### 1. Test the manifest

Visit: `https://YOUR-APP.onrender.com/manifest.xml`

Should see XML with correct URLs.

### 2. Test the web interface

Visit: `https://YOUR-APP.onrender.com`

Should see the spreadsheet with "Install Excel Add-in" button.

### 3. Test the installer

Click "Install Excel Add-in" → Download for your platform → Run it.

Excel should show the add-in in Insert → My Add-ins → Developer Add-ins.

---

## 🐛 Troubleshooting

### Manifest not loading

- Check CORS headers in `server/index.js`
- Verify all URLs use HTTPS on Render
- Check Render logs for errors

### Add-in doesn't appear in Excel

1. Clear Office cache:
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`

2. Re-run installer as Administrator (Windows)

3. Check Developer registry key:
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer
   ```

### SSE connection fails

- Render may require WebSocket upgrade
- Check if `/api/stream` endpoint is accessible
- Verify no proxy issues

---

## 📖 User Instructions

Once deployed, share this URL with users:

```
https://YOUR-APP.onrender.com
```

They can:
1. Click "Install Excel Add-in" button
2. Download installer for their platform
3. Run installer (as Administrator on Windows)
4. Open Excel → Insert → My Add-ins → Developer Add-ins
5. Activate "OpenGov Office Sync"

---

## 🔒 Security Notes

- This uses **Developer mode sideloading** (fine for internal/demo)
- For production, consider:
  - AppSource submission (Microsoft store)
  - Centralized deployment via Microsoft 365 admin
  - Code signing for installers

---

## 📝 Files to Commit

Make sure these are in your repo:

```
manifest.xml
install-excel-addin.bat
install-excel-addin.sh
web/assets/icon-*.png
DEPLOYMENT.md
```

---

Good luck! 🚀

