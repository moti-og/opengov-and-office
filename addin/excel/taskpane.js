const SERVER_URL = 'https://localhost:3000';
const DOCUMENT_ID = 'excel-prototype-doc';

let eventSource = null;
let isConnected = false;

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('syncBtn').onclick = syncExcelToWeb;
        document.getElementById('refreshBtn').onclick = refreshFromWeb;
        await initializeExcelSync();
    }
});

function updateStatus(text, connected = null) {
    const statusEl = document.getElementById('status');
    const statusText = document.getElementById('statusText');
    statusText.textContent = text;
    
    if (connected !== null) {
        isConnected = connected;
        statusEl.className = 'status ' + (connected ? 'connected' : 'disconnected');
        document.getElementById('syncBtn').disabled = !connected;
        document.getElementById('refreshBtn').disabled = !connected;
    }
}

async function readExcelData() {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load('values');
        await context.sync();
        return range.values;
    });
}

async function writeExcelData(data) {
    if (!data || !data.length) return;
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
        range.values = data;
        await context.sync();
    });
}

async function syncExcelToWeb() {
    try {
        updateStatus('Syncing to web...', true);
        const data = await readExcelData();
        
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data })
        });
        
        if (response.ok) {
            updateStatus('âœ“ Synced to web', true);
            updateDataPreview(data);
        } else {
            updateStatus('Error syncing', false);
        }
    } catch (error) {
        console.error('Error syncing Excel data:', error);
        updateStatus('Error syncing', false);
    }
}

async function refreshFromWeb() {
    try {
        updateStatus('Refreshing from web...', true);
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
        const doc = await response.json();
        
        if (doc && doc.data) {
            await writeExcelData(doc.data);
            updateStatus('âœ“ Refreshed from web', true);
            updateDataPreview(doc.data);
        }
    } catch (error) {
        console.error('Error refreshing data:', error);
        updateStatus('Error refreshing', false);
    }
}

async function initializeExcelSync() {
    try {
        updateStatus('Connecting to server...', false);
        
        // Fetch initial data
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
        
        if (response.ok) {
            const doc = await response.json();
            if (doc && doc.data && doc.data.length > 0) {
                await writeExcelData(doc.data);
                updateDataPreview(doc.data);
            } else {
                // Create default data
                const defaultData = [
                    ['Header 1', 'Header 2', 'Header 3'],
                    ['Row 1 A', 'Row 1 B', 'Row 1 C'],
                    ['Row 2 A', 'Row 2 B', 'Row 2 C']
                ];
                
                await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        data: defaultData,
                        title: 'Excel Prototype',
                        type: 'excel'
                    })
                });
                
                await writeExcelData(defaultData);
                updateDataPreview(defaultData);
            }
            
            updateStatus('âœ“ Connected', true);
        } else {
            updateStatus('Server not responding', false);
        }
        
        // Set up SSE
        setupSSE();
        
    } catch (error) {
        console.error('Error initializing:', error);
        updateStatus('Connection failed', false);
    }
}

function setupSSE() {
    if (eventSource) {
        eventSource.close();
    }
    
    eventSource = new EventSource(`${SERVER_URL}/api/stream`);
    
    eventSource.addEventListener('data-update', async (event) => {
        try {
            const { documentId, data } = JSON.parse(event.data);
            if (documentId === DOCUMENT_ID) {
                console.log('Received update from server');
                await writeExcelData(data);
                updateDataPreview(data);
                updateStatus('âœ“ Updated from server', true);
            }
        } catch (error) {
            console.error('Error handling SSE update:', error);
        }
    });
    
    eventSource.onopen = () => {
        console.log('SSE connection opened');
        updateStatus('âœ“ Live sync active', true);
    };
    
    eventSource.onerror = (error) => {
        console.error('SSE error:', error);
        updateStatus('Connection lost, reconnecting...', false);
        eventSource.close();
        // Retry after 5 seconds
        setTimeout(setupSSE, 5000);
    };
}

function updateDataPreview(data) {
    const preview = document.getElementById('dataPreview');
    if (!data || !data.length) {
        preview.innerHTML = '<em>No data</em>';
        return;
    }
    
    const rows = data.slice(0, 3); // Show first 3 rows
    const html = rows.map(row => 
        row.slice(0, 3).join(' | ') // Show first 3 columns
    ).join('<br>');
    
    preview.innerHTML = `<strong>Preview:</strong><br>${html}${data.length > 3 ? '<br>...' : ''}`;
}
