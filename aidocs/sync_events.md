# Sync Events & APIs Documentation

This document explains what events we track, what data we sync, and the APIs we use for bidirectional synchronization between Excel and the web interface.

## Data Flow Overview

```
Excel ←→ MongoDB (Server) ←→ Web (Luckysheet)
```

- **Excel Add-in**: Uses Office.js API to read/write data and layout
- **Server**: Express.js with MongoDB storage and Server-Sent Events (SSE)
- **Web Interface**: Luckysheet spreadsheet library

## What We Sync

### 1. Cell Data
- **Type**: 2D array of strings `[[String]]`
- **Direction**: Bidirectional (Excel ↔ Web)
- **Storage**: `Document.data` in MongoDB

### 2. Layout (Column/Row Sizing)
- **Column Widths**: Array of numbers (pixels) `[Number]`
- **Row Heights**: Array of numbers (pixels) `[Number]`
- **Direction**: Bidirectional (Excel ↔ Web)
- **Storage**: `Document.layout.columnWidths` and `Document.layout.rowHeights` in MongoDB

## Excel Add-in Events & APIs

### Events We Listen For

#### 1. Cell Changed Event
**API**: `sheet.onChanged.add()`
**Trigger**: When any cell value is modified
**Debounce**: 1000ms (1 second)
**Action**: Reads all data + layout, sends to server

```javascript
sheet.onChanged.add(async (event) => {
  // Debounced - triggers syncExcelToServer() after 1s of no changes
});
```

#### 2. Manual Sync Button
**Trigger**: User clicks "Manual Sync" button in task pane
**Action**: Immediately reads data + layout, sends to server

### APIs We Use (Office.js)

#### Reading Data
```javascript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const range = sheet.getUsedRange();
range.load('values, rowCount, columnCount');
await context.sync();
const data = range.values; // 2D array
```

#### Reading Column Widths
```javascript
for (let c = 0; c < colCount; c++) {
  const col = sheet.getRangeByIndexes(0, c, 1, 1);
  col.load('format/columnWidth');
  // Returns width in pixels
}
```

#### Reading Row Heights
```javascript
for (let r = 0; r < rowCount; r++) {
  const row = sheet.getRangeByIndexes(r, 0, 1, 1);
  row.load('format/rowHeight');
  // Returns height in pixels
}
```

#### Writing Data
```javascript
const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
range.values = data;
```

#### Writing Column Widths
```javascript
const col = sheet.getRangeByIndexes(0, colIndex, 1, 1);
col.format.columnWidth = width; // in pixels
```

#### Writing Row Heights
```javascript
const row = sheet.getRangeByIndexes(rowIndex, 0, 1, 1);
row.format.rowHeight = height; // in pixels
```

## Server (MongoDB + Express)

### Data Model
```javascript
{
  documentId: String,
  title: String,
  type: 'excel' | 'word' | 'powerpoint' | 'web',
  data: [[String]],           // 2D array of cell values
  layout: {
    columnWidths: [Number],   // Array of column widths in pixels
    rowHeights: [Number]      // Array of row heights in pixels
  },
  metadata: {
    createdAt: Date,
    updatedAt: Date,
    version: Number
  }
}
```

### API Endpoints

#### POST `/api/documents/:id/update`
**Purpose**: Update document data and/or layout
**Request Body**:
```javascript
{
  data: [[String]],      // Optional - cell values
  layout: {              // Optional - sizing info
    columnWidths: [Number],
    rowHeights: [Number]
  },
  title: String,
  type: String
}
```

**Response**: Updated document
**Side Effect**: Broadcasts `data-update` event via SSE to all connected clients

#### GET `/api/stream`
**Purpose**: Server-Sent Events (SSE) stream for real-time updates
**Events Sent**:
- `connected`: Client successfully connected
- `data-update`: Document data/layout changed
  ```javascript
  {
    type: 'data-update',
    documentId: String,
    data: [[String]],
    layout: { columnWidths: [Number], rowHeights: [Number] }
  }
  ```

## Web Interface (Luckysheet) Events & APIs

### Events We Listen For

#### 1. Cell Edit
**Hook**: `cellEditAfter(r, c, oldValue, newValue)`
**Trigger**: After user finishes editing a cell
**Debounce**: 500ms
**Action**: Extracts data, sends to server

#### 2. Cell Update
**Hook**: `cellUpdated(r, c, oldValue, newValue, isRefresh)`
**Trigger**: When cell value changes (including programmatic)
**Filter**: Only if `!isInitializing`
**Debounce**: 500ms
**Action**: Extracts data, sends to server

#### 3. Range Edit
**Hook**: `rangeEditAfter(range, data)`
**Trigger**: After editing multiple cells at once
**Debounce**: 500ms
**Action**: Extracts data, sends to server

#### 4. Column Width Changed
**Hook**: `columnWidthChangeAfter(colIndex, colWidth)`
**Trigger**: After user resizes a column
**Debounce**: 500ms
**Action**: Extracts layout, sends to server

#### 5. Row Height Changed
**Hook**: `rowHeightChangeAfter(rowIndex, rowHeight)`
**Trigger**: After user resizes a row
**Debounce**: 500ms
**Action**: Extracts layout, sends to server

### APIs We Use (Luckysheet)

#### Initializing with Layout
```javascript
luckysheet.create({
  data: [{
    name: "Sheet1",
    data: celldata,      // Array of {r, c, v} objects
    config: {
      columnlen: {       // Column widths
        0: 100,          // Column 0 = 100px
        1: 150,          // Column 1 = 150px
        // ...
      },
      rowlen: {          // Row heights
        0: 25,           // Row 0 = 25px
        1: 30,           // Row 1 = 30px
        // ...
      }
    }
  }]
});
```

#### Reading Data
```javascript
const sheetData = luckysheet.getSheetData();
// Returns 2D array of cell objects
```

#### Reading Layout
```javascript
const config = luckysheet.getConfig();
const columnWidths = config.columnlen; // Object: {colIndex: width}
const rowHeights = config.rowlen;      // Object: {rowIndex: height}
```

#### Writing Column Width
```javascript
luckysheet.setColumnWidth(columnIndex, width);
```

#### Writing Row Height
```javascript
luckysheet.setRowHeight(rowIndex, height);
```

#### Writing Cell Value
```javascript
luckysheet.setCellValue(row, col, {
  v: value,          // Actual value
  m: value           // Display value
});
```

## Sync Flow Examples

### Example 1: User Edits Cell in Excel

1. **Excel**: User types "Hello" in cell A1
2. **Excel**: `sheet.onChanged` event fires
3. **Excel**: After 1s debounce, `syncExcelToServer()` called
4. **Excel**: Reads data + layout via Office.js API
5. **Excel**: POSTs to `/api/documents/:id/update`
6. **Server**: Updates MongoDB document
7. **Server**: Broadcasts SSE event to all clients
8. **Web**: Receives SSE event
9. **Web**: Calls `loadDataIntoLuckysheet()` with new data
10. **Web**: Luckysheet displays "Hello" in A1

### Example 2: User Resizes Column in Web

1. **Web**: User drags column A to 150px wide
2. **Web**: `columnWidthChangeAfter` hook fires
3. **Web**: After 500ms debounce, `syncLayoutToServer()` called
4. **Web**: Extracts layout via `luckysheet.getConfig()`
5. **Web**: POSTs to `/api/documents/:id/update` with `layout` only
6. **Server**: Updates `Document.layout.columnWidths` in MongoDB
7. **Server**: Broadcasts SSE event with new layout
8. **Excel**: Receives SSE event
9. **Excel**: Calls `writeExcelData()` with new layout
10. **Excel**: Applies column width via Office.js API
11. **Excel**: Column A is now 150px wide

## Debouncing Strategy

### Why Debounce?
- Prevents excessive server requests during rapid typing/editing
- Reduces database writes
- Improves performance

### Debounce Timings
- **Excel cell changes**: 1000ms (1 second)
  - Reason: Users may type multiple characters quickly
- **Web cell changes**: 500ms
  - Reason: Faster feedback, less typing delay
- **Web layout changes**: 500ms
  - Reason: Column/row resizing is usually a single action

## Deduplication

### Purpose
Prevent infinite sync loops where:
1. Excel updates → Web updates → triggers Web hook → sends to server
2. Server broadcasts → Excel receives → triggers Excel event → infinite loop

### Implementation
Both Excel and Web track `lastSyncedData`:
```javascript
// Before syncing
if (JSON.stringify(newData) === JSON.stringify(lastSyncedData)) {
  return; // Skip sync - data unchanged
}

// After successful sync
lastSyncedData = newData;
```

### SSE Handling
When receiving SSE updates:
```javascript
if (JSON.stringify(payload.data) === JSON.stringify(lastSyncedData)) {
  return; // Skip - we sent this update
}
```

## Initialization Flags

### Purpose
Prevent hooks from firing during initial data load, which would cause unnecessary sync attempts.

### Implementation
```javascript
let isInitializing = true;

// In all hooks
if (isInitializing) {
  return; // Skip sync
}

// After data loaded
setTimeout(() => {
  isInitializing = false;
}, 1000);
```

## Current Limitations

### What We DON'T Sync (Yet)
- ❌ Cell formatting (colors, fonts, bold, italic)
- ❌ Formulas (only calculated values)
- ❌ Merged cells
- ❌ Conditional formatting
- ❌ Charts
- ❌ Images
- ❌ Cell comments
- ❌ Data validation rules
- ❌ Multiple sheets

### Planned Features
1. **Formula Sync**: Store formula + calculated value
2. **Formatting Sync**: Extend schema to include cell styles
3. **Multi-sheet Support**: Track active sheet
4. **Conflict Resolution**: Handle simultaneous edits

## Performance Considerations

### Optimization Strategies
1. **Batch Updates**: Send layout separately from data
2. **Partial Updates**: Only send changed cells (future enhancement)
3. **Compression**: Consider compressing large datasets
4. **Connection Pooling**: MongoDB connection reuse
5. **SSE Connection Management**: Auto-reconnect on failure

### Current Performance
- **Small datasets** (< 100 cells): < 100ms sync time
- **Medium datasets** (100-1000 cells): < 500ms sync time
- **Large datasets** (> 1000 cells): 1-2s sync time

## Troubleshooting

### Common Issues

**Issue**: Changes in Excel don't appear in web
- **Check**: Excel console for errors
- **Check**: Server logs for POST requests
- **Check**: Web console for SSE messages

**Issue**: Changes in web don't appear in Excel
- **Check**: Web console for "Proceeding with sync!"
- **Check**: `isInitializing` flag is false
- **Check**: Excel is not in edit mode (cell-editing error)

**Issue**: Layout changes don't sync
- **Check**: Column/row resize hooks are firing
- **Check**: `layout` object in MongoDB is populated
- **Check**: Office.js has permission to modify formats

### Debug Mode
Enable verbose logging:
```javascript
// All hooks and sync operations log to console
// Look for:
// - "cellUpdated fired:"
// - "Syncing to server"
// - "Received update from Excel/Web"
```

## References

- [Office.js API Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Luckysheet Documentation](https://mengshukeji.github.io/LuckysheetDocs/guide/)
- [Luckysheet Hooks API](https://mengshukeji.github.io/LuckysheetDocs/guide/config.html#hook)
- [Server-Sent Events (SSE) Spec](https://developer.mozilla.org/en-US/docs/Web/API/Server-sent_events)
- [MongoDB Schema Design](https://www.mongodb.com/docs/manual/core/data-modeling-introduction/)

