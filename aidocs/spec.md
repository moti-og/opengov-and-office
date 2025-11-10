# OpenGov Office - Prototype Specification

**Version:** 1.0 (Simplified Prototype)  
**Date:** November 10, 2025  
**Purpose:** Architectural demonstration and technical feasibility proof

---

## 1. Prototype Goals

**Primary Goal:** Demonstrate that data can sync in real-time between Excel and a web interface.

**What This Proves:**
- ✅ Office.js can read/write Excel data
- ✅ SSE enables real-time updates
- ✅ Shared architecture works across platforms
- ✅ Technical feasibility for larger system

**What This is NOT:**
- ❌ Production-ready system
- ❌ Multi-user collaborative editing
- ❌ Secure or authenticated
- ❌ Handling complex Excel features

---

## 2. Simplified Architecture

```
┌─────────────┐                    ┌─────────────┐
│   Excel     │◄─────REST API─────►│             │
│   Add-in    │                    │   Express   │
│ (Office.js) │◄──SSE (updates)────│   Server    │
└─────────────┘                    │  (Node.js)  │
                                   │             │
┌─────────────┐                    │ In-Memory   │
│     Web     │◄─────REST API─────►│   Store     │
│  Interface  │                    │  (simple    │
│   (React)   │◄──SSE (updates)────│   object)   │
└─────────────┘                    └─────────────┘
```

### Key Simplifications

| Area | Production Approach | Prototype Approach |
|------|---------------------|-------------------|
| **Storage** | MongoDB with schemas | In-memory JavaScript object |
| **Authentication** | User accounts, sessions | None - single shared data |
| **State Management** | State matrix pattern | Direct broadcast on change |
| **Conflict Resolution** | Checkout/checkin | Last write wins |
| **Error Handling** | Comprehensive recovery | Basic try/catch |
| **Platform Support** | Excel + Word + PPT + Web | Excel + Web only |
| **Data Scope** | Full Office XML | Simple 2D table array |

---

## 3. Minimal Technical Stack

### Backend
- **Node.js + Express** - Server (< 200 lines of code)
- **No database** - In-memory storage
- **SSE** - Real-time updates

### Frontend
- **React** - Simple UI (< 300 lines)
- **Office.js** - Excel integration
- **Vanilla CSS** - No fancy styling

### Development
- **npm** - Package manager
- **nodemon** - Auto-restart server
- **No build step** - Keep it simple

---

## 4. Data Model

**Single in-memory data structure:**

javascript
// server/store.js
const store = {
  // Simple 2D array - that's it!
  data: [
    ['Name', 'Age', 'City'],
    ['Alice', '30', 'NYC'],
    ['Bob', '25', 'LA'],
    ['Charlie', '35', 'SF']
  ],
  
  // Track SSE connections for broadcast
  clients: []
};


**That's the entire data model.** No MongoDB, no schemas, no complexity.

---

## 5. API Design (Minimal)

### Endpoints (Only 3 needed)

javascript
// 1. Get current data
GET /api/data
Response: { data: [[...], [...]] }

// 2. Update data (from any platform)
POST /api/data
Body: { data: [[...], [...]] }
Response: { success: true }

// 3. SSE stream (real-time updates)
GET /api/stream
Response: text/event-stream


### SSE Events (Only 1 type)

javascript
event: data-update
data: { data: [[...], [...]] }


**That's it.** When anyone updates data, server broadcasts to all connected clients.

---

## 6. Implementation

### 6.1 Server (server/index.js)

javascript
const express = require('express');
const cors = require('cors');
const app = express();

// In-memory store
const store = {
  data: [
    ['Name', 'Age', 'City'],
    ['Alice', '30', 'NYC'],
    ['Bob', '25', 'LA']
  ],
  clients: []
};

app.use(cors());
app.use(express.json());

// Get data
app.get('/api/data', (req, res) => {
  res.json({ data: store.data });
});

// Update data
app.post('/api/data', (req, res) => {
  store.data = req.body.data;
  
  // Broadcast to all connected clients
  const message = `data: ${JSON.stringify({ data: store.data })}\n\n`;
  store.clients.forEach(client => client.write(message));
  
  res.json({ success: true });
});

// SSE stream
app.get('/api/stream', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  // Add this client to broadcast list
  store.clients.push(res);
  
  // Remove on disconnect
  req.on('close', () => {
    store.clients = store.clients.filter(client => client !== res);
  });
  
  // Send initial data
  res.write(`data: ${JSON.stringify({ data: store.data })}\n\n`);
});

app.listen(3000, () => console.log('Server on http://localhost:3000'));


**Total: ~50 lines of server code**

### 6.2 Excel Add-in (addin/excel/excel.js)

javascript
// Read data from Excel
async function readExcelData() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load('values');
    await context.sync();
    return range.values;
  });
}

// Write data to Excel
async function writeExcelData(data) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange().clear();
    const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
    range.values = data;
    await context.sync();
  });
}

// Send to server when user clicks "Sync to Web"
document.getElementById('syncBtn').onclick = async () => {
  const data = await readExcelData();
  await fetch('http://localhost:3000/api/data', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data })
  });
};

// Listen for updates from server
const eventSource = new EventSource('http://localhost:3000/api/stream');
eventSource.onmessage = (event) => {
  const { data } = JSON.parse(event.data);
  writeExcelData(data);
};


**Total: ~40 lines of add-in code**

### 6.3 Web Interface (web/app.js)

javascript
function App() {
  const [data, setData] = React.useState([]);
  
  // Connect to SSE stream
  React.useEffect(() => {
    const eventSource = new EventSource('http://localhost:3000/api/stream');
    eventSource.onmessage = (event) => {
      const { data } = JSON.parse(event.data);
      setData(data);
    };
    return () => eventSource.close();
  }, []);
  
  // Update cell
  const handleCellChange = (rowIdx, colIdx, value) => {
    const newData = [...data];
    newData[rowIdx][colIdx] = value;
    setData(newData);
    
    // Send to server
    fetch('http://localhost:3000/api/data', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: newData })
    });
  };
  
  return (
    <table>
      {data.map((row, rowIdx) => (
        <tr key={rowIdx}>
          {row.map((cell, colIdx) => (
            <td key={colIdx}>
              <input
                value={cell}
                onChange={(e) => handleCellChange(rowIdx, colIdx, e.target.value)}
              />
            </td>
          ))}
        </tr>
      ))}
    </table>
  );
}


**Total: ~40 lines of web code**

---

## 7. Project Structure (Minimal)

opengov-office/
├── server/
│   └── index.js           # Express server (50 lines)
├── addin/
│   └── excel/
│       ├── manifest.xml   # Office add-in config
│       ├── taskpane.html  # Add-in UI
│       └── excel.js       # Office.js code (40 lines)
├── web/
│   ├── index.html         # Web UI
│   └── app.js             # React app (40 lines)
├── package.json
└── README.md


**Total code: ~150 lines** (excluding HTML/config)

---

## 8. Demo Flow

### Setup (5 minutes)
bash
npm install express cors
node server/index.js
# Server running on http://localhost:3000


### Demonstrate (2 minutes)

1. **Open Excel** → Sideload add-in
2. **Enter data** in Excel cells:
   
   Name    | Age | City
   Alice   | 30  | NYC
   Bob     | 25  | LA
   
3. **Click "Sync to Web"** button in add-in
4. **Open browser** → http://localhost:3000
5. **See data** appear in web table instantly
6. **Edit cell** in web interface
7. **Watch Excel** update in real-time

**Demo complete!** You've proven the architecture works.

---

## 9. What This Demonstrates

✅ **Technical Feasibility**
- Office.js APIs work as expected
- SSE provides real-time sync
- Cross-platform data flow is possible

✅ **Architecture Validation**
- Server as single source of truth
- Platform-agnostic approach works
- Simple REST + SSE pattern scales

✅ **UX Proof**
- Bidirectional updates feel natural
- Real-time sync is fast enough
- Multi-platform editing is intuitive

---

## 10. What's NOT Included (Intentionally)

❌ MongoDB or any database  
❌ User authentication  
❌ Multiple documents  
❌ Conflict resolution  
❌ Offline support  
❌ Error recovery  
❌ Testing suite  
❌ Production deployment  
❌ Security measures  
❌ State matrix pattern  
❌ Session management  

**These can be added later when building the real system.**

---

## 11. Implementation Timeline

### Day 1: Server & API
- [ ] Set up Express server
- [ ] Create in-memory store
- [ ] Implement 3 API endpoints
- [ ] Test with curl/Postman

### Day 2: Web Interface
- [ ] Create simple React table UI
- [ ] Connect to SSE stream
- [ ] Implement cell editing
- [ ] Test bidirectional updates

### Day 3: Excel Add-in
- [ ] Create manifest.xml
- [ ] Build taskpane HTML
- [ ] Implement Office.js read/write
- [ ] Connect to server API

### Day 4: Integration & Demo
- [ ] Test Excel ↔ Web sync
- [ ] Polish UI (minimal styling)
- [ ] Prepare demo script
- [ ] Document setup instructions

**Total: 4 days** instead of 10 weeks

---

## 12. Success Criteria

The prototype is successful if:

✅ You can enter data in Excel  
✅ Data appears in web browser within 1 second  
✅ You can edit data in web browser  
✅ Excel updates within 1 second  
✅ Works with 3x5 table (15 cells)  
✅ Code is under 200 lines total  
✅ Setup takes under 10 minutes  
✅ Demo takes under 5 minutes  

**That's it!** Simple, focused, achievable.

---

## 13. Migration Path to Production

When ready to build the real system:

1. **Add MongoDB** → Replace in-memory store
2. **Add authentication** → User accounts and sessions
3. **Add state matrix** → Better conflict resolution
4. **Add Word/PPT** → Extend platform adapters
5. **Add error handling** → Comprehensive recovery
6. **Add testing** → Automated test suite
7. **Add features** → Charts, formatting, formulas

**The prototype proves the architecture works.** Building it out is just engineering.

---

## 14. Key Learnings from Simplification

**Before:** 10 weeks, MongoDB, state matrix, 3 platforms, authentication  
**After:** 4 days, in-memory, direct sync, 2 platforms, no auth

**Prototype Philosophy:**
- Prove the concept, not production-ready
- Minimal code, maximum learning
- Fast iteration over comprehensive features
- Demonstrate architecture, not scalability

---

## 15. Next Steps

1. **Review this simplified spec** ✓
2. **Create project structure** (folders, files)
3. **Build server in 1 day** (50 lines of code)
4. **Build web UI in 1 day** (40 lines of code)
5. **Build Excel add-in in 1 day** (40 lines of code)
6. **Integrate and demo in 1 day**
7. **Present to stakeholders** (show it works!)

---

**Status:** Ready to Build  
**Estimated Time:** 4 days  
**Code Complexity:** ~150 lines  
**Learning Value:** Maximum  

Let's build a prototype, not a product! 🚀
