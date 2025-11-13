# Budget Book Page - SIMPLIFIED Spec

## What We're Building
A single static web page that looks EXACTLY like the Birmingham screenshot + one button in Excel to push the table data into it.

## The Simplest Approach
1. **Static web page** (`/budget-book.html`) - exact copy of Birmingham layout (header, logo, text, etc.)
2. **One button** in Excel add-in: "Update Budget Book"  
3. **Click button** → Excel table data → Web page table updates
4. **Data reconstruction** (not screenshots) - we pull the data and style it to match Birmingham table

## Files Needed
```
web/budget-book.html    - Birmingham layout (static)
web/budget-book.css     - Birmingham styling
web/budget-book.js      - Simple table updater
```

## Excel Add-in Change
Add ONE button: **"Update Budget Book"**
- In the "OpenGov Budget Book" section
- Next to existing button
- On click: Send table data to server → Web page updates

## How It Works

### 1. Excel Side
```javascript
// User clicks "Update Budget Book" button
// Reads current table/selection in Excel
// Sends to: POST /api/budget-book/update
// Payload: { data: [[row1], [row2], ...] }
```

### 2. Server Side  
```javascript
// Stores in MongoDB: budgetBookData collection
// NO broadcasting, NO SSE (it's manual push only)
```

### 3. Web Page
```javascript
// On load: GET /api/budget-book
// Renders Birmingham-style table
// That's it - it's static until next "Update" button click
```

## Implementation Steps

1. **Create branch**: `feature/budget-book-page`
2. **Build static HTML** with Birmingham styling
3. **Add button** to Excel add-in
4. **Create API endpoint** for storing table data
5. **Wire it up** - button sends, web page displays
6. **Test locally**
7. **Merge & deploy**

## Done!
Simple. No real-time sync. No complex transformations. Just:
- Beautiful static page
- One button click
- Table appears

