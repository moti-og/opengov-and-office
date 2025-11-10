const SERVER_URL = 'http://localhost:3001';
const DOCUMENT_ID = 'excel-demo-doc';

const COLS = 26; // A-Z columns
let ROWS = 100; // Will be calculated based on viewport

let eventSource = null;
let isConnected = false;
let gridData = [];
let syncTimeout = null;
let lastSyncedData = null;

// Calculate rows based on viewport height
function calculateRows() {
    const headerHeight = 60; // Header bar height
    const rowHeight = 32; // Each row height
    const colHeaderHeight = 32; // Column header height
    const availableHeight = window.innerHeight - headerHeight - colHeaderHeight;
    ROWS = Math.floor(availableHeight / rowHeight) + 10; // Add extra rows for scrolling
    
    // Initialize empty grid
    gridData = [];
    for (let i = 0; i < ROWS; i++) {
        gridData[i] = new Array(COLS).fill('');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    calculateRows();
    initializeGrid();
    initializeSync();
});

// Recalculate on window resize
window.addEventListener('resize', () => {
    const oldRows = ROWS;
    calculateRows();
    if (oldRows !== ROWS) {
        document.getElementById('spreadsheet').innerHTML = '';
        initializeGrid();
    }
});

function getColumnLabel(index) {
    let label = '';
    while (index >= 0) {
        label = String.fromCharCode(65 + (index % 26)) + label;
        index = Math.floor(index / 26) - 1;
    }
    return label;
}

function initializeGrid() {
    const spreadsheet = document.getElementById('spreadsheet');
    spreadsheet.style.setProperty('--rows', ROWS);
    spreadsheet.style.setProperty('--cols', COLS);

    // Corner cell
    const corner = document.createElement('div');
    corner.className = 'cell corner';
    spreadsheet.appendChild(corner);

    // Column headers (A, B, C, ...)
    for (let col = 0; col < COLS; col++) {
        const header = document.createElement('div');
        header.className = 'cell header';
        header.textContent = getColumnLabel(col);
        spreadsheet.appendChild(header);
    }

    // Rows with row headers
    for (let row = 0; row < ROWS; row++) {
        // Row header (1, 2, 3, ...)
        const rowHeader = document.createElement('div');
        rowHeader.className = 'cell row-header';
        rowHeader.textContent = row + 1;
        spreadsheet.appendChild(rowHeader);

        // Data cells
        for (let col = 0; col < COLS; col++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            
            const input = document.createElement('input');
            input.type = 'text';
            input.dataset.row = row;
            input.dataset.col = col;
            input.value = '';
            
            input.addEventListener('input', handleCellChange);
            input.addEventListener('keydown', handleKeyDown);
            
            cell.appendChild(input);
            spreadsheet.appendChild(cell);
        }
    }
}

function handleKeyDown(e) {
    const input = e.target;
    const row = parseInt(input.dataset.row);
    const col = parseInt(input.dataset.col);

    if (e.key === 'Enter') {
        e.preventDefault();
        moveFocus(row + 1, col);
    } else if (e.key === 'Tab') {
        e.preventDefault();
        moveFocus(row, col + 1);
    } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        moveFocus(row + 1, col);
    } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        moveFocus(row - 1, col);
    } else if (e.key === 'ArrowLeft' && input.selectionStart === 0) {
        e.preventDefault();
        moveFocus(row, col - 1);
    } else if (e.key === 'ArrowRight' && input.selectionStart === input.value.length) {
        e.preventDefault();
        moveFocus(row, col + 1);
    }
}

function moveFocus(row, col) {
    if (row < 0 || row >= ROWS || col < 0 || col >= COLS) return;
    
    const input = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
    if (input) {
        input.focus();
        input.select();
    }
}

function handleCellChange(e) {
    const input = e.target;
    const row = parseInt(input.dataset.row);
    const col = parseInt(input.dataset.col);
    const value = input.value;

    gridData[row][col] = value;

    // Debounce sync
    clearTimeout(syncTimeout);
    syncTimeout = setTimeout(() => {
        syncToServer();
    }, 500);
}

function updateStatus(text, connected = null) {
    const dot = document.getElementById('syncDot');
    const label = document.getElementById('syncLabel');
    label.textContent = text;

    if (connected !== null) {
        isConnected = connected;
        dot.className = 'dot ' + (connected ? 'connected' : 'disconnected');
    }
}

function renderData(data) {
    if (!data || data.length === 0) {
        updateStatus('✓ Connected (no data)', true);
        return;
    }

    // Update gridData with server data
    for (let row = 0; row < Math.min(data.length, ROWS); row++) {
        for (let col = 0; col < Math.min(data[row].length, COLS); col++) {
            gridData[row][col] = data[row][col] || '';
            
            const input = document.querySelector(`input[data-row="${row}"][data-col="${col}"]`);
            if (input && document.activeElement !== input) {
                input.value = gridData[row][col];
            }
        }
    }
}

async function syncToServer() {
    try {
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

async function fetchData() {
    try {
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
        
        if (response.ok) {
            const doc = await response.json();
            if (doc && doc.data && doc.data.length) {
                renderData(doc.data);
                lastSyncedData = doc.data;
                updateStatus('✓ Live sync active', true);
            } else {
                updateStatus('✓ Connected (no data)', true);
            }
        } else {
            updateStatus('Document not found', false);
        }
    } catch (error) {
        console.error('Fetch error:', error);
        updateStatus('Connection failed', false);
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
                renderData(payload.data);
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

async function initializeSync() {
    updateStatus('Connecting...', false);
    await fetchData();
    setupSSE();
}
