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
**Status:** Pending

**Context:** You have 4 platforms to build (Word, Excel, PPT, Web). Start with all 4 or build incrementally?

**Options:**
- a) Excel + Web only (establish the sync pattern first)
- b) Excel + Word + Web (two Office apps to validate cross-app sync)
- c) All 4 platforms (Excel + Word + PPT + Web) from the start
- d) Web + platform abstraction layer first, then add Office apps one by one

**Answer:**

**Notes:**

---

### Question 3: Data Model Architecture - Where does the XML live?
**Status:** Pending

**Context:** Need to decide where the canonical data source lives and how platforms sync.

**Options:**
- a) Server-side XML store - Server holds the truth, platforms push/pull changes
- b) Peer-to-peer - Each Office document contains its own XML, syncs via server
- c) Hybrid - Server maintains XML, but platforms can work offline and sync when connected
- d) Web-first - Web app is the source of truth, Office add-ins are views

**Answer:**

**Notes:**

---

### Question 4: Platform Abstraction Strategy
**Status:** Pending

**Context:** Each Office app has different APIs and data structures (Word paragraphs, Excel ranges, PPT slides).

**Options:**
- a) Build abstraction layer from scratch that normalizes all platforms
- b) Use existing library (e.g., Office.js unified APIs where possible)
- c) Start simple - just handle tables/data that exist across all platforms
- d) Platform-specific implementations initially, abstract later

**Answer:**

**Notes:**

---

### Question 5: Relationship to contract-doc-v3 Word add-in
**Status:** Pending

**Context:** You have a working Word add-in with server, React UI, SSE, state matrix. How does this relate?

**Options:**
- a) Extend contract-doc-v3 to add Excel/PPT/Web support (same codebase)
- b) New standalone project, but reuse patterns/architecture from contract-doc-v3
- c) New project, reference contract-doc-v3 for scaffolding guidance only
- d) Completely separate - different architecture approach

**Answer:**

**Notes:**

---

### Question 6: Real-time Sync Mechanism
**Status:** Pending

**Context:** How do changes propagate across platforms in real-time?

**Options:**
- a) Server-Sent Events (SSE) like contract-doc-v3 uses
- b) WebSockets for true bidirectional communication
- c) Polling (simpler, but less real-time)
- d) Operational Transform / CRDT (like Yjs in contract-doc-v3's Hocuspocus)

**Answer:**

**Notes:**

---

## Additional Context

### Key Information from contract-doc-v3 (Word add-in reference):
- Uses Office.js for Word integration
- React-based shared UI components (shared-ui folder)
- Express.js server with SSE (Server-Sent Events) for real-time updates
- State matrix architecture for UI state management
- File-based storage (no database)
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

