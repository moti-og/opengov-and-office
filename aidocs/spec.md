# OpenGov Office - Specification

## Goals

Demonstrate bidirectional real-time sync between Office add-ins and web using MongoDB.

- Start with Excel + Web (tables only)
- Extend to Word and PowerPoint
- Prove multi-platform architecture works

## Architecture

Excel/Word/PPT Add-ins + Web Interface â†’ Office.js APIs (REST + SSE) â†’ Express Server (Node.js) â†’ MongoDB

Key Pattern: Server as single source of truth, SSE for real-time updates

## Tech Stack

Backend: Node.js, Express, MongoDB, SSE
Frontend: React, Office.js, vanilla CSS
Tools: npm, MongoDB Atlas or local

## Next Steps

1. Set up MongoDB
2. Create project folders
3. Build Phase 1 (Excel + Web)
4. Test and iterate
5. Demo to stakeholders

Start with web/Excel - that is the proof of concept!
