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

function initializeLuckysheet(initialData = []) {
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
            data: initialData,
            config: {},
            index: 0
        }],
        hook: {
            cellUpdated: function(r, c, oldValue, newValue, isRefresh) {
                if (isInitializing || isRefresh) return;
                
                console.log('Cell updated:', r, c, newValue);
                // Debounce sync
                clearTimeout(window.luckysheetSyncTimeout);
                window.luckysheetSyncTimeout = setTimeout(() => {
                    syncToServer();
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

function loadDataIntoLuckysheet(data) {
    if (!data || !data.length) {
        updateStatus('✓ Connected (no data)', true);
        return;
    }

    isInitializing = true;
    
    // Set cell values directly without clearing (to avoid issues)
    for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < (data[r] ? data[r].length : 0); c++) {
            const value = data[r][c];
            if (value !== null && value !== undefined && value !== '') {
                // Use setCellValue with all required params to avoid undefined issues
                luckysheet.setCellValue(r, c, {
                    v: value,
                    m: value
                });
            }
        }
    }
    
    setTimeout(() => {
        isInitializing = false;
    }, 1000);
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
                loadDataIntoLuckysheet(payload.data);
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
        // Fetch data from server
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
        
        let initialData = [];
        if (response.ok) {
            const doc = await response.json();
            if (doc && doc.data && doc.data.length) {
                initialData = arrayToLuckysheet(doc.data);
                lastSyncedData = doc.data;
            }
        }
        
        // Initialize Luckysheet with the data
        initializeLuckysheet(initialData);
        
        // Wait for Luckysheet to fully render
        setTimeout(() => {
            isInitializing = false;
            setupSSE();
            updateStatus('✓ Live sync active', true);
        }, 1500);
        
    } catch (error) {
        console.error('Initialization error:', error);
        updateStatus('Connection failed', false);
        // Initialize empty on error
        initializeLuckysheet([]);
    }
}

