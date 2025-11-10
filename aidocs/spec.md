# OpenGov Office - Specification

## Goals

Demonstrate bidirectional real-time sync between Office add-ins and web using MongoDB.

- Start with Excel + Web (tables only) ✅
- Extend to Word and PowerPoint
- Prove multi-platform architecture works

## Architecture

Excel Add-in (Office.js) <-> Express Server (HTTP/SSE) <-> MongoDB <-> Web Interface (Luckysheet)

**Key Patterns:**
- Server as single source of truth
- SSE (Server-Sent Events) for real-time push updates
- Event-driven sync (no polling)
- Smart debouncing and deduplication

## Tech Stack

**Backend:** 
- Node.js, Express, MongoDB, Mongoose
- SSE for real-time updates
- CORS enabled

**Frontend - Excel Add-in:**
- Office.js API
- Webpack bundler
- HTTPS with trusted certs (office-addin-dev-certs)
- Auto-sideload with office-addin-debugging

**Frontend - Web:**
- **Luckysheet** - Full-featured spreadsheet library with:
  - Excel-like UI and UX
  - Formula support
  - Formatting support
  - Conditional formatting
  - Charts and graphs
- Vanilla JavaScript (no React yet)
- Real-time sync with Excel

## Current Status

✅ **Phase 1: Excel + Web Sync - COMPLETE**
- Excel add-in with onChanged listener (1s debounce)
- Web interface using Luckysheet library
- MongoDB storage
- Bidirectional real-time sync via SSE
- Deduplication to prevent infinite loops
- Green/red sync indicators
- Manual sync override

## Next Steps

1. ✅ Set up MongoDB
2. ✅ Create project folders
3. ✅ Build Phase 1 (Excel + Web) with Luckysheet
4. 🔄 Test formula sync between Excel and Luckysheet
5. 🔄 Test formatting sync (colors, fonts, etc.)
6. 📋 Extend to Word
7. 📋 Extend to PowerPoint
8. 📋 Demo to stakeholders

## How to Run

1. Start MongoDB (local or Atlas)
2. Run: `.\tools\scripts\run-local.bat start`
3. Excel will open with add-in loaded
4. Open web: `http://localhost:3001`
5. Edit in either place - watch real-time sync!

## Why Luckysheet?

Instead of building a spreadsheet from scratch, Luckysheet gives us:
- Professional Excel-like interface
- Built-in formula engine
- Cell formatting, merging, conditional formatting
- Copy/paste, undo/redo
- Charts and pivot tables
- Active development and documentation

This lets us focus on the **sync architecture** rather than rebuilding Excel.
