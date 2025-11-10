# Luckysheet Integration

## What Changed

We've replaced our custom vanilla JS spreadsheet grid with **Luckysheet**, a professional open-source Excel-like spreadsheet library.

## Why Luckysheet?

Instead of syncing each Excel feature one-by-one (formulas, formatting, merging, etc.), Luckysheet provides:

✅ **Built-in Excel Features:**
- Formula engine (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Cell formatting (colors, fonts, borders, alignment)
- Conditional formatting
- Cell merging
- Number formatting
- Multiple sheets
- Charts and pivot tables
- Copy/paste, undo/redo
- Excel import/export

✅ **Professional UI:**
- Looks and feels like Excel
- Keyboard shortcuts (Ctrl+C, Ctrl+V, arrow keys)
- Context menus
- Drag to fill
- Column/row resize

✅ **Developer-Friendly:**
- Hooks for onChange events
- Easy data import/export
- Active community and documentation

## Current Implementation

### Data Flow

```
Excel Cell Values → MongoDB (2D Array) → Luckysheet Display
```

### What Syncs Now
- ✅ Cell values (text, numbers)
- ✅ Grid structure (rows x columns)
- ✅ Real-time bidirectional updates

### What's Coming
- 🔄 Formulas (Excel =SUM(A1:A5) → Luckysheet calculation)
- 🔄 Cell formatting (colors, fonts, bold, italic)
- 🔄 Column widths and row heights
- 🔄 Merged cells
- 🔄 Number formats ($, %, dates)

## Code Structure

### `web/app.js`

**Key Functions:**
- `arrayToLuckysheet(arr)` - Converts 2D array from MongoDB to Luckysheet format
- `luckysheetToArray()` - Converts Luckysheet data back to 2D array for storage
- `loadDataIntoLuckysheet(data)` - Updates the spreadsheet with new data
- `syncToServer()` - Sends changes to MongoDB

**Luckysheet Format:**
```javascript
// Each cell is an object with position and value
[
  {
    r: 0,      // row index
    c: 0,      // column index
    v: {
      v: "Product",     // actual value
      m: "Product",     // displayed value
      ct: { fa: "General", t: "g" }  // cell type
    }
  }
]
```

### `web/index.html`

Includes Luckysheet CSS and JS from `node_modules`:
```html
<link rel="stylesheet" href="../node_modules/luckysheet/dist/css/luckysheet.css" />
<script src="../node_modules/luckysheet/dist/luckysheet.umd.js"></script>
```

### `server/index.js`

Added static file serving for Luckysheet assets:
```javascript
app.use('/node_modules', express.static(path.join(__dirname, '..', 'node_modules')));
```

## Testing

1. **Start the servers:**
   ```bash
   .\tools\scripts\run-local.bat start
   ```

2. **Test basic sync:**
   - Edit a cell in Excel → Watch it appear in web (http://localhost:3001)
   - Edit a cell in web → Watch it update in Excel

3. **Test formulas (in Excel):**
   - Enter `=SUM(A1:A5)` in Excel
   - Currently: Just the value syncs (not the formula)
   - Future: Formula will be preserved and calculated in Luckysheet

## Next Steps

### Phase 2: Formula Sync
Instead of just syncing calculated values, sync the actual formulas:
- Excel: `=SUM(A1:A5)` → MongoDB: `{ formula: "=SUM(A1:A5)", value: 150 }`
- Luckysheet can then calculate and display the formula

### Phase 3: Formatting Sync
Extend the MongoDB schema to store:
```javascript
{
  r: 0,
  c: 0,
  v: "Product",
  bg: "#FF0000",      // background color
  fc: "#FFFFFF",      // font color
  fs: "14px",         // font size
  bl: 1,              // bold
  // etc.
}
```

### Phase 4: Advanced Features
- Multiple sheets
- Merged cells
- Conditional formatting
- Charts

## Resources

- [Luckysheet Documentation](https://mengshukeji.github.io/LuckysheetDocs/guide/)
- [Luckysheet GitHub](https://github.com/mengshukeji/Luckysheet)
- [API Reference](https://mengshukeji.github.io/LuckysheetDocs/guide/api.html)

