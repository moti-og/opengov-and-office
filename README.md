# OpenGov Office Add-in

**Real-time bidirectional sync between Office apps and MongoDB**

🔗 Built with: Office.js • Express • MongoDB • Server-Sent Events

---

## 🚀 Quick Start

### Prerequisites

- **Node.js 16+**
- **MongoDB** (local or Atlas)
- **Excel Desktop** (Windows/Mac)

### Installation

```bash
npm install
```

### Start the Application

**Windows:**

```bash
tools\scripts\start.bat
```

This will:
- Start backend server (port 3001)
- Start Excel add-in (port 3000)
- Auto-sideload into Excel
- Open Excel with add-in loaded

**Manual start:**

```bash
# Terminal 1: Backend
npm run server

# Terminal 2: Add-in (opens Excel)
npm start
```

---

## 📁 Project Structure

```
opengov-and-office/
├── 📁 addin/               # Microsoft Office add-ins
│   └── excel/              # Excel add-in
│       ├── taskpane/       # Task pane UI (HTML, CSS, JS)
│       ├── commands/       # Ribbon commands
│       └── manifest.xml    # Office manifest
├── 📁 server/              # Backend API (Node.js + Express)
│   ├── index.js            # Server entry point
│   ├── models/             # MongoDB schemas
│   └── routes/             # REST API endpoints
├── 📁 web/                 # Web interface (future)
├── 📁 shared-ui/           # Shared React components (future)
├── 📁 data/                # Runtime data storage
├── 📁 tools/               # Build and deployment scripts
│   └── scripts/            # Windows .bat scripts
├── 📁 docs/                # Documentation
│   ├── spec.md             # Project specification
│   └── Project-Summary.md  # Architecture overview
└── 📁 assets/              # Icons and images
```

---

## 🏗️ Architecture

```
┌─────────────────┐         ┌─────────────────┐         ┌─────────────────┐
│  Excel Add-in   │◄───────►│  Express Server │◄───────►│    MongoDB      │
│   (port 3000)   │   REST  │   (port 3001)   │         │                 │
└─────────────────┘         └─────────────────┘         └─────────────────┘
         ▲                           │
         └───────────────────────────┘
              SSE (Real-time)
```

### Key Features

✅ **Bidirectional Sync** - Excel ↔ MongoDB  
✅ **Real-time Updates** - Server-Sent Events  
✅ **Auto-sideloading** - Yeoman tooling  
✅ **HTTPS** - Trusted localhost certificates  
✅ **Versioning** - MongoDB document tracking  

---

## 🛠️ Development

### Available Scripts

| Command | Description |
|---------|-------------|
| `npm start` | Start add-in (auto-opens Excel) |
| `npm stop` | Stop debugging |
| `npm run server` | Start backend only |
| `npm run build` | Production build |
| `tools\scripts\start.bat` | Start all (Windows) |
| `tools\scripts\stop.bat` | Stop all (Windows) |

### Configuration

Create `.env` file:

```env
MONGODB_URI=mongodb://localhost:27017/opengov-office
SERVER_PORT=3001
```

### Ports

- **3000** - Add-in dev server (HTTPS)
- **3001** - Backend API (HTTP)

---

## 📡 API Reference

Base: `http://localhost:3001`

### Endpoints

```
GET  /api/health              # Health check
GET  /api/stream              # SSE connection
GET  /api/documents           # List documents
GET  /api/documents/:id       # Get document
POST /api/documents/:id/update # Create/update
```

### SSE Events

- `connected` - Initial connection
- `data-update` - Document changed
- `document-created` - New document

---

## 🐛 Troubleshooting

### Add-in doesn't load

```bash
# Check server status
tools\scripts\servers.bat status

# Clear Office cache
# Close Excel, delete: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

### SSL errors

```bash
npx office-addin-dev-certs install
```

### Changes not showing

- Webpack auto-rebuilds on save
- Refresh task pane in Excel
- If stuck: `npm stop` then `npm start`

---

## 🔜 Roadmap

### Phase 1: Excel + Web ✅ (In Progress)

- [x] Excel add-in with Office.js
- [x] MongoDB storage
- [x] REST API
- [x] SSE real-time updates
- [ ] Web interface

### Phase 2: Multi-Platform

- [ ] Word add-in
- [ ] PowerPoint add-in
- [ ] Shared React components

### Phase 3: Collaboration

- [ ] Multi-user editing
- [ ] Conflict resolution
- [ ] Version history

---

## 📚 Documentation

- 📖 [Project Specification](docs/spec.md)
- 📖 [Architecture Overview](docs/Project-Summary.md)
- 📖 [Office.js Docs](https://learn.microsoft.com/office/dev/add-ins/)

---

## 📄 License

MIT

---

**Built for OpenGov** 🚀
