# OpenGov Office - Multi-Platform Add-in Project - Planning Questions

## Project Scope & Architecture

### Question 1: What is the core vision and scope?
**Status:** ✅ ANSWERED

**Answer:** Unified Office Suite Add-in System (Word + Excel + PowerPoint + Web)

**Project Name:** OpenGov Office (formerly excel_ftw)

**Core Concept:**
- **Shared XML Data Model** - All Office documents (Word, Excel, PowerPoint) use XML as the underlying format (Office Open XML)
- **Multi-Platform Sync** - Same data accessible and editable from:
  - Word add-in
  - Excel add-in
  - PowerPoint add-in
  - Web interface
- **Bidirectional Updates** - Change in one platform updates all others
- **Purpose:** Technical feasibility prototype + UX demonstration

**Incremental Implementation Priority:**
1. Present Excel view in website (tables, charts, graphs)
2. Bidirectional updates (change in one place updates both)
3. Add-in can affect both web and Excel spreadsheet
4. Extend to Word and PowerPoint

**Key Technical Insight:**
- Office Open XML is the common denominator
- Each platform (Word/Excel/PPT) presents different views of the same underlying data
- Platform abstraction layer needed to handle Word tables vs Excel ranges vs PPT tables

---

### Question 2: MVP Platform Scope - Which platforms first?
**Status:** ✅ ANSWERED

**Context:** You have 4 platforms to build (Word, Excel, PPT, Web). Start with all 4 or build incrementally?

**Options:**
- a) Excel + Web only (establish the sync pattern first)
- b) Excel + Word + Web (two Office apps to validate cross-app sync)
- c) All 4 platforms (Excel + Word + PPT + Web) from the start
- d) Web + platform abstraction layer first, then add Office apps one by one

**Answer:** Option (a) - Excel + Web only

**Notes:**
- contract-doc-v3 proves the pattern works with Word + Web
- Start with Excel + Web to establish and validate the bidirectional sync pattern
- Excel is the natural first choice since your project started as "excel_ftw"
- Once sync pattern is solid, extend to Word (reuse contract-doc-v3 knowledge), then PowerPoint
- Incremental approach reduces complexity and allows for course correction

---

### Question 3: Data Model Architecture - Where does the XML live?
**Status:** ✅ ANSWERED

**Context:** Need to decide where the canonical data source lives and how platforms sync.

**Options:**
- a) Server-side XML store - Server holds the truth, platforms push/pull changes
- b) Peer-to-peer - Each Office document contains its own XML, syncs via server
- c) Hybrid - Server maintains XML, but platforms can work offline and sync when connected
- d) Web-first - Web app is the source of truth, Office add-ins are views

**Answer:** Option (a) - Server-side store with MongoDB

**Notes:**
- Server maintains canonical XML/JSON data in MongoDB as the source of truth
- MongoDB provides persistent, reliable storage with query capabilities
- Platforms push changes via REST API and receive updates via SSE
- Simpler mental model: one source of truth, all platforms are views/editors
- Collections structure:
  * documents - Store Office Open XML data and metadata
  * sessions - Track active user sessions and platform connections
  * changes - Optional: Change history/audit log
- Office documents can be generated from MongoDB data on-demand
- Scalable for multi-user scenarios
- Future: Can add option (c) hybrid approach for offline support later
- Note: contract-doc-v3 uses file-based storage, but MongoDB is better for production use

---

### Question 4: Platform Abstraction Strategy
**Status:** ✅ ANSWERED

**Context:** Each Office app has different APIs and data structures (Word paragraphs, Excel ranges, PPT slides).

**Options:**
- a) Build abstraction layer from scratch that normalizes all platforms
- b) Use existing library (e.g., Office.js unified APIs where possible)
- c) Start simple - just handle tables/data that exist across all platforms
- d) Platform-specific implementations initially, abstract later

**Answer:** Option (c) + (d) hybrid - Start simple with tables, platform-specific initially

**Notes:**
- contract-doc-v3 uses thin platform adapters (Office.js wrappers) for Word
- Start with tables/data structures that exist across Excel, Word, and PowerPoint
- Build platform-specific adapters first (excel-adapter.js, word-adapter.js, etc.)
- Abstract commonalities as patterns emerge - don't over-engineer upfront
- Office.js Common API is limited - most functionality requires host-specific APIs
- Focus on: tables, basic formatting, data ranges as MVP scope
- Example: Excel range → XML → Word table → PowerPoint table
- Create abstraction layer once you understand the actual requirements (2-3 platforms working)

---

### Question 5: Relationship to contract-doc-v3 Word add-in
**Status:** ✅ ANSWERED

**Context:** You have a working Word add-in with server, React UI, SSE, state matrix. How does this relate?

**Options:**
- a) Extend contract-doc-v3 to add Excel/PPT/Web support (same codebase)
- b) New standalone project, but reuse patterns/architecture from contract-doc-v3
- c) New project, reference contract-doc-v3 for scaffolding guidance only
- d) Completely separate - different architecture approach

**Answer:** Option (b) - New standalone project, reuse patterns/architecture

**Notes:**
- Keep projects separate (different business domains: contracts vs. OpenGov data)
- Mirror the proven architecture patterns from contract-doc-v3:
  * State matrix for UI state management
  * Express.js server with SSE for real-time updates
  * Shared React components (shared-ui/)
  * File-based storage (data/ folder structure)
  * Office.js manifest and add-in structure
  * Development scripts (tools/scripts/)
- Reuse architectural patterns, NOT code (different data models)
- Can copy folder structure, build configuration, and development workflow
- contract-doc-v3 serves as the reference implementation for "how to do it right"

---

### Question 6: Real-time Sync Mechanism
**Status:** ✅ ANSWERED

**Context:** How do changes propagate across platforms in real-time?

**Options:**
- a) Server-Sent Events (SSE) like contract-doc-v3 uses
- b) WebSockets for true bidirectional communication
- c) Polling (simpler, but less real-time)
- d) Operational Transform / CRDT (like Yjs in contract-doc-v3's Hocuspocus)

**Answer:** Option (a) - Server-Sent Events (SSE), with option (d) for future collaborative editing

**Notes:**
- contract-doc-v3 successfully uses SSE for real-time updates
- SSE is perfect for server → client updates (state matrix pattern)
- Simpler than WebSockets, built-in browser reconnection
- Architecture: 
  * Clients POST changes to REST API
  * Server updates state and broadcasts via SSE to all connected clients
  * Unidirectional flow matches state matrix pattern
- Implementation: GET /api/v1/events endpoint with text/event-stream
- For multi-user collaborative editing (future), add Yjs/Hocuspocus like contract-doc-v3
- SSE handles: document updates, state changes, approval flows, LLM responses
- Works great for both Office.js add-ins and web clients

---

## Additional Context

### Key Information from contract-doc-v3 (Word add-in reference):
- Uses Office.js for Word integration
- React-based shared UI components (shared-ui folder)
- Express.js server with SSE (Server-Sent Events) for real-time updates
- State matrix architecture for UI state management
- File-based storage (no database) - **Note: OpenGov Office will use MongoDB instead**
- Supports sideloading via manifest.xml
- Hocuspocus/Yjs for collaboration features

### Technical Research Questions to Explore:
- [ ] Office Open XML format structure (WordprocessingML, SpreadsheetML, PresentationML)
- [ ] How do Word tables, Excel ranges, and PPT tables map to shared data model?
- [ ] Office.js Common API surface across Word/Excel/PowerPoint
- [ ] Manifest configuration for multi-host add-ins (single manifest for all 3 apps?)
- [ ] Chart/graph representation across platforms
- [ ] Performance implications of real-time XML sync
- [ ] Offline support and conflict resolution

### Architecture Considerations:
1. **Data Model Layer** - Abstract XML representation that works across all platforms
2. **Platform Adapters** - Word/Excel/PPT-specific implementation of data model
3. **Sync Engine** - Real-time bidirectional sync mechanism
4. **Web Renderer** - Display Excel tables, Word content, PPT slides in browser
5. **State Management** - Track which platform has lock, handle concurrent edits

### MVP Success Criteria:
- [ ] Simple data (e.g., table with 3 columns, 5 rows) works in all platforms
- [ ] Edit in Excel → updates in Web in real-time
- [ ] Edit in Web → updates in Excel in real-time
- [ ] Demonstrates technical feasibility
- [ ] Good UX (smooth, not janky)

---

**Last Updated:** November 10, 2025
