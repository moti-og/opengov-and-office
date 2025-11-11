// Auto-detect server URL based on where the add-in is loaded from
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : 'https://opengov-and-office.onrender.com';
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isUpdating = false;
let syncQueue = [];
let syncInProgress = false;

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        await init();
        setupChangeListener();
        setupModalHandlers();
        
        // Manual sync button
        const btn = document.getElementById('manualSyncBtn');
        if (btn) {
            btn.onclick = async () => {
                console.log('Manual sync triggered');
                await sync();
            };
        }
    }
});

function setupModalHandlers() {
    const modal = document.getElementById('modal');
    const updateBudgetBtn = document.getElementById('updateBudgetBtn');
    const closeBtn = document.querySelector('.close');
    
    // Open modal when button clicked
    if (updateBudgetBtn) {
        updateBudgetBtn.onclick = () => {
            modal.style.display = 'block';
            setTimeout(() => {
                modal.style.display = 'none';
            }, 3000); // Auto-close after 3 seconds
        };
    }
    
    // Close modal when X clicked
    if (closeBtn) {
        closeBtn.onclick = () => {
            modal.style.display = 'none';
        };
    }
    
    // Close modal when clicking outside
    window.onclick = (event) => {
        if (event.target === modal) {
            modal.style.display = 'none';
        }
    };
}

function updateStatus(text, connected) {
    document.getElementById('syncLabel').textContent = text;
    document.getElementById('syncDot').className = 'dot ' + (connected ? 'connected' : 'disconnected');
}

async function readData() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load('values');
        await context.sync();
        return usedRange.values.map(row => row.map(cell => cell ? String(cell) : ''));
    });
}

async function writeData(data) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Clear existing data first
        const usedRange = sheet.getUsedRange(true);
        if (usedRange) {
            usedRange.clear();
        }
        
        // Write new data
        if (data && data.length > 0) {
            const maxCols = Math.max(...data.map(r => r.length));
            const range = sheet.getRangeByIndexes(0, 0, data.length, maxCols);
            range.values = data;
        }
        
        await context.sync();
    });
}

async function queueSync() {
    try {
        const data = await readData();
        syncQueue.push(data);
        
        // Debounce the queue processing, not individual edits
        clearTimeout(window.queueTimeout);
        window.queueTimeout = setTimeout(processQueue, 500);
    } catch (err) {
        if (err.message?.includes('cell-editing mode')) {
            console.log('Cell editing mode - will retry');
        }
    }
}

async function processQueue() {
    if (syncInProgress || syncQueue.length === 0) return;
    
    syncInProgress = true;
    const data = syncQueue[syncQueue.length - 1]; // Take latest
    syncQueue = []; // Clear queue (we're sending the latest state)
    
    try {
        console.log('Syncing to server:', data.length, 'rows');
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data, title: 'Excel Demo', type: 'excel' })
        });
        
        if (!response.ok) {
            console.error('Sync failed:', response.status);
            updateStatus('Sync failed', false);
            syncInProgress = false;
            return;
        }
        
        updateStatus('✓ Synced', true);
    } catch (err) {
        console.error('Sync error:', err);
        updateStatus('Sync failed', false);
    }
    
    syncInProgress = false;
    
    // Process next item if queue filled up while we were syncing
    if (syncQueue.length > 0) {
        setTimeout(processQueue, 100);
    }
}

// Legacy sync function for manual sync button
async function sync() {
    await queueSync();
}

async function applyUpdate(data) {
    if (isUpdating) {
        console.log('Already updating');
        return;
    }
    
    console.log('Applying update from web:', data.length, 'rows');
    isUpdating = true;
    
    try {
        await writeData(data);
        updateStatus('✓ Updated from web', true);
    } catch (err) {
        if (err.message?.includes('cell-editing mode')) {
            console.log('Cell editing mode - skipping update');
            isUpdating = false;
            return;
        }
        console.error('Apply error:', err);
        updateStatus('Update failed', false);
    }
    
    setTimeout(() => { isUpdating = false; }, 2000);
}

function setupChangeListener() {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onChanged.add(async () => {
            if (isUpdating) return;
            queueSync();
        });
        await context.sync();
    });
}

let reconnectAttempts = 0;

function setupSSE() {
    if (eventSource) {
        try { eventSource.close(); } catch(e) {}
    }
    
    eventSource = new EventSource(`${SERVER_URL}/api/stream`);
    
    eventSource.addEventListener('message', async (e) => {
        const msg = JSON.parse(e.data);
        if (msg.type === 'data-update' && msg.documentId === DOCUMENT_ID) {
            await applyUpdate(msg.data);
        }
    });
    
    eventSource.onopen = () => {
        reconnectAttempts = 0;
        updateStatus('✓ Live sync', true);
    };
    
    eventSource.onerror = () => {
        eventSource.close();
        reconnectAttempts++;
        
        if (reconnectAttempts > 5) {
            updateStatus('Server offline', false);
            return;
        }
        
        updateStatus(`Reconnecting (${reconnectAttempts}/5)...`, false);
        setTimeout(setupSSE, 5000 * reconnectAttempts);
    };
}

async function init() {
    updateStatus('Connecting...', false);
    isUpdating = true;
    
    try {
        const healthRes = await fetch(`${SERVER_URL}/api/health`);
        if (!healthRes.ok) {
            updateStatus('Server not responding', false);
            console.error('Server health check failed');
            isUpdating = false;
            return;
        }
        
        const excelData = await readData();
        const hasExcel = excelData.length > 0 && excelData.some(row => row.some(cell => cell));
        console.log('Excel has data:', hasExcel, excelData.length, 'rows');
        
        const docRes = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`, {
            headers: { 'Cache-Control': 'no-cache' }
        });
        
        if (docRes.ok) {
            const doc = await docRes.json();
            const hasServer = doc?.data?.length > 0;
            console.log('Server has data:', hasServer, doc?.data?.length || 0, 'rows');
            
            if (hasExcel) {
                // Excel has data - push to server
                console.log('Syncing Excel data to server');
                await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ data: excelData, title: 'Excel Demo', type: 'excel' })
                });
            } else if (hasServer) {
                // Server has data, Excel empty - pull from server
                console.log('Loading server data into Excel');
                await writeData(doc.data);
            }
        }
        
        updateStatus('✓ Connected', true);
        console.log('Init complete, setting up SSE');
        setupSSE();
        
    } catch (error) {
        console.error('Init error:', error);
        updateStatus('Connection failed', false);
    }
    
    await new Promise(resolve => setTimeout(resolve, 1000));
    isUpdating = false;
    console.log('Ready');
}
