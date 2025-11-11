# Scripts Directory

All installer and utility scripts for the OpenGov Office Sync project.

## Excel Add-in Installers

### Local Testing
- **install-excel-addin-local.bat** - Windows installer (points to localhost:3001)
- **install-excel-addin-local.sh** - Mac installer (points to localhost:3001)

### Production
- **install-excel-addin.bat** - Windows installer (points to Render deployment)
- **install-excel-addin.sh** - Mac installer (points to Render deployment)

### Uninstallers (Smart - work for both local and production)
- **uninstall-excel-addin.bat** - Windows uninstaller (automatically detects & removes any version)
- **uninstall-excel-addin.sh** - Mac uninstaller (automatically detects & removes any version)

## Utility Scripts

- **create-icons.js** - Generates placeholder icon files in `web/assets/`
- **SAFE-START.bat** - Starts servers with safe rate limiting settings
- **START-SIMPLE.bat** - Simple server startup script
- **clear-and-restart.bat** - Kills all node processes and restarts servers
- **run-opengov-office-local.bat** - Comprehensive local development startup

## Usage

See [LOCAL-TESTING.md](../../LOCAL-TESTING.md) for complete testing instructions.

### Quick Start

**Uninstall any existing add-in:**
```bash
./tools/scripts/uninstall-excel-addin.bat  # Windows
./tools/scripts/uninstall-excel-addin.sh   # Mac
```
(Automatically detects and removes both local and production versions)

**Install for local testing:**
```bash
./tools/scripts/install-excel-addin-local.bat  # Windows
./tools/scripts/install-excel-addin-local.sh   # Mac
```

**Install for production:**
```bash
./tools/scripts/install-excel-addin.bat  # Windows
./tools/scripts/install-excel-addin.sh   # Mac
```

**Generate icons:**
```bash
node tools/scripts/create-icons.js
```

