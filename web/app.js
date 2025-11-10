const SERVER_URL = 'http://localhost:3001';
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isConnected = false;
let lastSyncedData = null;
let isInitializing = true;
let luckysheetInstance = null;

// Initialize on load
document.addEventListener('DOMContentLoaded', () => {
    // Fetch data first, then initialize Luckysheet with it
    initializeWithData();
});

function updateStatus(text, connected = null) {
    const dot = document.getElementById('syncDot');
    const label = document.getElementById('syncLabel');
    label.textContent = text;

    if (connected !== null) {
        isConnected = connected;
        dot.className = 'dot ' + (connected ? 'connected' : 'disconnected');
    }
}

function initializeLuckysheet(initialData = [], initialLayout = null) {
    console.log('=== initializeLuckysheet called ===');
    console.log('initialData length:', initialData.length);
    console.log('initialData:', initialData);
    console.log('initialLayout:', initialLayout);
    
    const config = {};
    
    // Apply column widths if provided
    if (initialLayout && initialLayout.columnWidths && initialLayout.columnWidths.length > 0) {
        config.columnlen = {};
        initialLayout.columnWidths.forEach((width, index) => {
            config.columnlen[index] = width;
        });
    }
    
    // Apply row heights if provided
    if (initialLayout && initialLayout.rowHeights && initialLayout.rowHeights.length > 0) {
        config.rowlen = {};
        initialLayout.rowHeights.forEach((height, index) => {
            config.rowlen[index] = height;
        });
    }
    
    const options = {
        container: 'luckysheet',
        title: 'OpenGov Office Sync',
        lang: 'en',
        showinfobar: false,
        showsheetbar: false,
        showstatisticBar: false,
        showConfigWindowResize: false,
        enableAddRow: true,
        enableAddCol: true,
        userInfo: false,
        myFolderUrl: false,
        data: [{
            name: "Sheet1",
            color: "",
            status: "1",
            order: "0",
            celldata: initialData,  // Use celldata for the array format [{r, c, v}]
            config: config,
            index: 0
        }],
        hook: {
            cellEditAfter: function(r, c, oldValue, newValue) {
                if (isInitializing) {
                    console.log('Skipping sync - initializing');
                    return;
                }
                
                console.log('cellEditAfter fired:', r, c, 'old:', oldValue, 'new:', newValue);
                // Debounce sync
                clearTimeout(window.luckysheetSyncTimeout);
                window.luckysheetSyncTimeout = setTimeout(() => {
                    console.log('Syncing to server after cell edit');
                    syncToServer();
                }, 500);
            },
            cellUpdated: function(r, c, oldValue, newValue, isRefresh) {
                console.log('cellUpdated fired:', {
                    row: r, 
                    col: c, 
                    oldValue: oldValue, 
                    newValue: newValue, 
                    isRefresh: isRefresh, 
                    isInitializing: isInitializing
                });
                
                // Only skip if initializing, ignore isRefresh
                if (isInitializing) {
                    console.log('Skipping sync - still initializing');
                    return;
                }
                
                console.log('Proceeding with sync!');
                // Debounce sync
                clearTimeout(window.luckysheetSyncTimeout);
                window.luckysheetSyncTimeout = setTimeout(() => {
                    console.log('Syncing to server after cell update');
                    syncToServer();
                }, 500);
            },
            rangeEditAfter: function(range, data) {
                if (isInitializing) {
                    console.log('Skipping sync - initializing');
                    return;
                }
                
                console.log('rangeEditAfter fired:', range, data);
                // Debounce sync
                clearTimeout(window.luckysheetSyncTimeout);
                window.luckysheetSyncTimeout = setTimeout(() => {
                    console.log('Syncing to server after range edit');
                    syncToServer();
                }, 500);
            },
            columnWidthChangeAfter: function(colIndex, colWidth) {
                if (isInitializing) {
                    console.log('Skipping sync - initializing');
                    return;
                }
                
                console.log('Column width changed:', colIndex, colWidth);
                // Debounce sync
                clearTimeout(window.luckysheetLayoutSyncTimeout);
                window.luckysheetLayoutSyncTimeout = setTimeout(() => {
                    console.log('Syncing layout to server after column resize');
                    syncLayoutToServer();
                }, 500);
            },
            rowHeightChangeAfter: function(rowIndex, rowHeight) {
                if (isInitializing) {
                    console.log('Skipping sync - initializing');
                    return;
                }
                
                console.log('Row height changed:', rowIndex, rowHeight);
                // Debounce sync
                clearTimeout(window.luckysheetLayoutSyncTimeout);
                window.luckysheetLayoutSyncTimeout = setTimeout(() => {
                    console.log('Syncing layout to server after row resize');
                    syncLayoutToServer();
                }, 500);
            }
        }
    };

    luckysheet.create(options);
    luckysheetInstance = luckysheet;
}

// Convert simple 2D array to Luckysheet celldata format
function arrayToLuckysheet(arr) {
    if (!arr || !arr.length) return [];
    
    const celldata = [];
    for (let r = 0; r < arr.length; r++) {
        for (let c = 0; c < (arr[r] ? arr[r].length : 0); c++) {
            const cellValue = arr[r][c];
            if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                celldata.push({
                    r: r,
                    c: c,
                    v: String(cellValue) // Just pass the string value directly
                });
            }
        }
    }
    return celldata;
}

// Convert Luckysheet format back to simple 2D array
function luckysheetToArray() {
    const sheetData = luckysheet.getSheetData();
    if (!sheetData || !sheetData.length) return [];
    
    // Find max row and col with actual data
    let maxRow = -1;
    let maxCol = -1;
    
    sheetData.forEach((row, rowIndex) => {
        if (row) {
            row.forEach((cell, colIndex) => {
                if (cell !== null && cell !== undefined) {
                    // Check if cell has a value
                    const value = typeof cell === 'object' ? cell.v : cell;
                    if (value !== null && value !== undefined && value !== '') {
                        maxRow = Math.max(maxRow, rowIndex);
                        maxCol = Math.max(maxCol, colIndex);
                    }
                }
            });
        }
    });
    
    // If no data, return empty array
    if (maxRow === -1 || maxCol === -1) return [];
    
    // Build 2D array
    const result = [];
    for (let r = 0; r <= maxRow; r++) {
        const row = [];
        for (let c = 0; c <= maxCol; c++) {
            const cell = sheetData[r] && sheetData[r][c];
            let value = '';
            
            if (cell !== null && cell !== undefined) {
                if (typeof cell === 'object' && cell.v !== undefined) {
                    value = String(cell.v);
                } else if (typeof cell === 'string' || typeof cell === 'number') {
                    value = String(cell);
                }
            }
            
            row.push(value);
        }
        result.push(row);
    }
    
    return result;
}

function loadDataIntoLuckysheet(data, skipInitFlag = false) {
    if (!data || !data.length) {
        updateStatus('✓ Connected (no data)', true);
        return;
    }

    // ALWAYS block syncs while loading data to prevent infinite loops
    isInitializing = true;
    console.log('Loading data, setting isInitializing = true');

    // Set cell values directly (no clearSheet - it doesn't exist)
    let cellsSet = 0;
    for (let r = 0; r < data.length; r++) {
        const row = data[r];
        if (!row || !Array.isArray(row)) continue; // Skip invalid rows
        
        for (let c = 0; c < row.length; c++) {
            const value = row[c];
            // Skip null, undefined, and empty string values
            if (value === null || value === undefined || value === '') {
                continue;
            }
            
            try {
                luckysheet.setCellValue(r, c, {
                    v: String(value), // Ensure it's a string
                    m: String(value)
                });
                cellsSet++;
            } catch (e) {
                console.warn(`Skipping cell (${r}, ${c}):`, e.message);
            }
        }
    }
    
    console.log(`Set ${cellsSet} cells from ${data.length} rows`);
    
    // Force a refresh after setting data
    try {
        luckysheet.refresh();
    } catch (e) {
        console.warn('Refresh error:', e);
    }
    
    // Re-enable syncs after a short delay (let all cellUpdated hooks finish firing)
    setTimeout(() => {
        isInitializing = false;
        console.log('Data loading complete, isInitializing = false');
    }, 1000);
}

function extractLayout() {
    const config = luckysheet.getConfig();
    const columnWidths = [];
    const rowHeights = [];
    
    // Extract column widths
    if (config && config.columnlen) {
        const maxCol = Math.max(...Object.keys(config.columnlen).map(Number));
        for (let i = 0; i <= maxCol; i++) {
            columnWidths.push(config.columnlen[i] || 73); // Default width
        }
    }
    
    // Extract row heights
    if (config && config.rowlen) {
        const maxRow = Math.max(...Object.keys(config.rowlen).map(Number));
        for (let i = 0; i <= maxRow; i++) {
            rowHeights.push(config.rowlen[i] || 19); // Default height
        }
    }
    
    return { columnWidths, rowHeights };
}

function applyLayoutToLuckysheet(layout) {
    if (!layout) return;
    
    console.log('Applying layout to Luckysheet:', layout);
    
    // Get current sheet dimensions to avoid applying layout to non-existent rows/cols
    const sheetData = luckysheet.getSheetData();
    if (!sheetData || !sheetData.length) {
        console.log('Sheet is empty, skipping layout application');
        return;
    }
    
    const maxRow = sheetData.length;
    const maxCol = sheetData[0] ? sheetData[0].length : 0;
    
    console.log(`Sheet has ${maxRow} rows, ${maxCol} cols`);
    
    // Apply column widths (only up to maxCol)
    if (layout.columnWidths && layout.columnWidths.length > 0) {
        for (let index = 0; index < Math.min(layout.columnWidths.length, maxCol); index++) {
            const width = layout.columnWidths[index];
            if (width && width > 0) {
                try {
                    luckysheet.setColumnWidth(index, width);
                } catch (e) {
                    // Silently ignore
                }
            }
        }
    }
    
    // Apply row heights (only up to maxRow)
    if (layout.rowHeights && layout.rowHeights.length > 0) {
        for (let index = 0; index < Math.min(layout.rowHeights.length, maxRow); index++) {
            const height = layout.rowHeights[index];
            if (height && height > 0) {
                try {
                    luckysheet.setRowHeight(index, height);
                } catch (e) {
                    // Silently ignore
                }
            }
        }
    }
    
    console.log(`Applied layout to ${maxRow} rows, ${maxCol} cols`);
}

async function syncToServer() {
    try {
        const gridData = luckysheetToArray();
        
        // Don't sync if data hasn't changed
        if (JSON.stringify(gridData) === JSON.stringify(lastSyncedData)) {
            return;
        }

        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                data: gridData,
                title: 'Excel Demo',
                type: 'web'
            })
        });

        if (response.ok) {
            lastSyncedData = JSON.parse(JSON.stringify(gridData));
            updateStatus('✓ Synced', true);
        }
    } catch (error) {
        console.error('Sync error:', error);
        updateStatus('Sync failed', false);
    }
}

async function syncLayoutToServer() {
    try {
        const layout = extractLayout();
        
        console.log('Syncing layout:', layout);

        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                layout: layout,
                title: 'Excel Demo',
                type: 'web'
            })
        });

        if (response.ok) {
            updateStatus('✓ Layout synced', true);
        }
    } catch (error) {
        console.error('Layout sync error:', error);
        updateStatus('Layout sync failed', false);
    }
}

function setupSSE() {
    if (eventSource) {
        eventSource.close();
    }

    eventSource = new EventSource(`${SERVER_URL}/api/stream`);

    eventSource.addEventListener('message', (event) => {
        try {
            const payload = JSON.parse(event.data);
            if (payload.type === 'data-update' && payload.documentId === DOCUMENT_ID) {
                // Don't update if data is same
                if (JSON.stringify(payload.data) === JSON.stringify(lastSyncedData)) {
                    return;
                }
                
                console.log('Received update from Excel');
                loadDataIntoLuckysheet(payload.data, true); // Skip init flag for SSE updates
                
                // Apply layout if provided
                if (payload.layout) {
                    applyLayoutToLuckysheet(payload.layout);
                }
                
                lastSyncedData = payload.data;
                updateStatus('✓ Updated from Excel', true);
            }
        } catch (error) {
            console.error('SSE error:', error);
        }
    });

    eventSource.onopen = () => {
        console.log('SSE connected');
        updateStatus('✓ Live sync active', true);
    };

    eventSource.onerror = (error) => {
        console.error('SSE connection error:', error);
        updateStatus('Reconnecting...', false);
        eventSource.close();
        setTimeout(setupSSE, 5000);
    };
}
async function initializeWithData() {
    updateStatus('Connecting...', false);
    
    try {
        // FIRST fetch data and layout
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
        
        let initialData = [];
        let initialLayout = null;
        
        console.log('Fetch response status:', response.status);
        
        if (response.ok) {
            const doc = await response.json();
            console.log('Fetched document:', doc);
            if (doc && doc.data && doc.data.length) {
                initialData = arrayToLuckysheet(doc.data);
                initialLayout = doc.layout;
                lastSyncedData = doc.data;
                console.log('Converted to Luckysheet format:', initialData.length, 'cells');
            } else {
                console.log('No data in document, initializing empty sheet');
            }
        } else {
            console.log('Document not found, initializing empty sheet');
        }
        
        console.log('Initializing Luckysheet with', initialData.length, 'cells');
        // Initialize Luckysheet with data AND layout
        initializeLuckysheet(initialData, initialLayout);
        
        // Wait for Luckysheet to fully render
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Setup SSE after everything is loaded
        isInitializing = false;
        console.log('Initialization complete - isInitializing now false');
        setupSSE();
        updateStatus('✓ Live sync active', true);
        
    } catch (error) {
        console.error('Initialization error:', error);
        updateStatus('Connection failed', false);
        // Initialize empty on error
        initializeLuckysheet([], null);
    }
}

