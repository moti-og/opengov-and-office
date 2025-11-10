# OpenGov Office - Project Summary

## Vision

Create a multi-platform Office add-in system that enables real-time bidirectional synchronization between Office applications (Excel, Word, PowerPoint) and a web interface, backed by MongoDB.

## Goals

### Phase 1: Proof of Concept ✅
- [x] Excel add-in with bidirectional sync
- [x] MongoDB document storage
- [x] REST API for document operations
- [x] Server-Sent Events for real-time updates
- [ ] Web interface for browser-based editing

### Phase 2: Multi-Platform
- [ ] Extend to Word add-in
- [ ] Extend to PowerPoint add-in
- [ ] Shared React components across platforms
- [ ] Unified data model

### Phase 3: Advanced Features
- [ ] Multi-user collaboration
- [ ] Conflict resolution
- [ ] Version history
- [ ] User permissions
- [ ] Document templates

## Architecture Principles

### 1. Server as Single Source of Truth
- All data stored in MongoDB
- Server computes state and broadcasts to clients
- No client-side conflict resolution needed

### 2. Real-Time Communication
- **REST API** for CRUD operations
- **Server-Sent Events (SSE)** for push updates
- Simpler than WebSockets for one-way server→client communication

### 3. Platform Isolation
- Each Office app gets its own add-in
- Shared backend logic
- Consistent experience across platforms

### 4. Simple Data Model
- Documents stored as 2D arrays (table data)
- Metadata tracks versions and timestamps
- Platform-agnostic structure

## Technical Stack

### Frontend
- **Office.js** - Microsoft Office JavaScript API
- **Vanilla JavaScript** - Add-in UI (current)
- **React** - Web interface (planned)
- **Webpack** - Build and bundling
- **Babel** - ES6+ transpilation

### Backend
- **Node.js** - Runtime environment
- **Express** - Web framework
- **MongoDB** - Document database
- **Mongoose** - ODM for MongoDB
- **SSE** - Real-time event streaming

### Development Tools
- **Yeoman** - Office add-in scaffolding
- **office-addin-dev-certs** - SSL certificates
- **office-addin-debugging** - Auto-sideloading
- **Webpack Dev Server** - Hot reload

## Data Flow

```
┌─────────────────────────────────────────────────────────────┐
│  Excel Add-in                                                │
│  ┌─────────────┐    Office.js    ┌──────────────────┐       │
│  │  Task Pane  │ ────────────────►│  Excel Workbook  │       │
│  └─────────────┘    read/write    └──────────────────┘       │
└──────────┬──────────────────────────────────────────────────┘
           │
           │ HTTPS (REST)
           ▼
┌─────────────────────────────────────────────────────────────┐
│  Express Server (port 3001)                                  │
│  ┌──────────────┐    ┌────────────┐    ┌────────────┐       │
│  │  REST API    │────│  Business  │────│  MongoDB   │       │
│  │  /api/docs   │    │   Logic    │    │  Models    │       │
│  └──────────────┘    └────────────┘    └────────────┘       │
│         │                                                     │
│         │ Broadcast via SSE                                  │
│         ▼                                                     │
│  ┌──────────────────────────────────────────────┐           │
│  │  SSE Stream (/api/stream)                    │           │
│  │  • data-update                               │           │
│  │  • document-created                          │           │
│  └──────────────────────────────────────────────┘           │
└──────────┬──────────────────────────────────────────────────┘
           │
           │ SSE Events
           ▼
┌─────────────────────────────────────────────────────────────┐
│  All Connected Clients                                       │
│  • Other Excel instances                                     │
│  • Web interface (future)                                    │
│  • Mobile apps (future)                                      │
└─────────────────────────────────────────────────────────────┘
```

## Key Decisions

### Why Office.js?
- Official Microsoft API for Office add-ins
- Cross-platform (Windows, Mac, Office Online)
- Sandboxed security model
- No need to modify Office installation

### Why Server-Sent Events?
- Simpler than WebSockets for unidirectional updates
- Built-in reconnection logic
- HTTP-based (easier to deploy)
- Perfect for server→client push notifications

### Why MongoDB?
- Flexible schema for different document types
- JSON-like storage matches JavaScript objects
- Easy to start (no schema migration needed)
- Can scale later if needed

### Why Yeoman?
- Official scaffolding tool from Microsoft
- Handles HTTPS certificates automatically
- Auto-sideloading into Office apps
- Production-ready webpack configuration

## Document Lifecycle

```
1. CREATE/OPEN
   • User opens Excel with add-in
   • Add-in initializes and connects to server
   • Fetches document data via REST API
   
2. EDIT
   • User modifies cells in Excel
   • Changes remain local until sync
   
3. SYNC TO SERVER
   • User clicks "Sync to Server"
   • Office.js reads Excel data
   • POST to /api/documents/:id/update
   • Server saves to MongoDB
   • Server broadcasts update via SSE
   
4. RECEIVE UPDATES
   • Other clients listen to SSE stream
   • On "data-update" event, fetch new data
   • Office.js writes data to Excel
   • UI shows "Updated from server"
   
5. CONFLICT HANDLING
   • Last write wins (for now)
   • Future: operational transform or CRDTs
```

## Security Considerations

### Current (Development)
- localhost-only access
- Trusted SSL certificates for development
- No authentication (single-user prototype)

### Future (Production)
- User authentication (OAuth 2.0)
- API key validation
- Rate limiting
- CORS whitelist
- Data encryption at rest
- Audit logging

## Performance Considerations

### Current Limitations
- Single server (no horizontal scaling)
- No caching layer
- Synchronous database operations
- Full document sync (no deltas)

### Optimization Opportunities
- Redis for caching and pub/sub
- Delta sync (only changed cells)
- Pagination for large documents
- CDN for static assets
- Load balancing with multiple servers

## Testing Strategy

### Manual Testing
- Open Excel, make changes, sync
- Open second Excel window, verify updates
- Check MongoDB for saved data
- Verify SSE connection in browser console

### Automated Testing (Future)
- Unit tests for server routes
- Integration tests for API endpoints
- E2E tests with Playwright
- Excel automation with Office Scripts

## Deployment Strategy

### Development
- Local MongoDB
- npm scripts to start servers
- Manual testing in desktop Excel

### Staging (Future)
- MongoDB Atlas cluster
- Heroku/Render for Express server
- GitHub Actions for CI/CD
- Test with Office Online

### Production (Future)
- Production MongoDB cluster
- Cloud hosting (Azure/AWS)
- CDN for add-in files
- AppSource distribution
- Monitoring and logging

## Success Metrics

### Phase 1 (Current)
- ✅ Excel add-in loads successfully
- ✅ Data syncs to MongoDB
- ✅ Real-time updates work between instances
- [ ] Web interface functional

### Phase 2
- [ ] Word add-in integrated
- [ ] PowerPoint add-in integrated
- [ ] Shared component library working

### Phase 3
- [ ] Multi-user collaboration tested
- [ ] Performance acceptable (< 2s sync time)
- [ ] Error handling comprehensive

## Known Limitations

1. **Single Document Model**: Only supports table/spreadsheet data
2. **No Offline Support**: Requires active internet connection
3. **Last Write Wins**: No sophisticated conflict resolution
4. **Limited File Size**: Large Excel files may cause performance issues
5. **Desktop Only**: Requires Excel desktop app (not fully tested on Office Online)

## Next Steps

1. **Complete Phase 1**
   - Build web interface
   - Test bidirectional sync Excel ↔ Web
   - Document APIs thoroughly

2. **Start Phase 2**
   - Research Word JavaScript API
   - Design document structure for Word
   - Implement Word add-in

3. **Prepare for Production**
   - Add authentication
   - Implement error handling
   - Set up monitoring
   - Write deployment guide

## Resources

- [Office.js Documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [Excel JavaScript API Reference](https://learn.microsoft.com/javascript/api/excel)
- [Yeoman Office Generator](https://github.com/OfficeDev/generator-office)
- [Server-Sent Events Spec](https://html.spec.whatwg.org/multipage/server-sent-events.html)
- [MongoDB Node.js Driver](https://www.mongodb.com/docs/drivers/node/current/)

---

**Last Updated**: November 2025  
**Status**: Phase 1 in progress

