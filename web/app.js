const SERVER_URL = 'http://localhost:3001';
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isUpdating = false;
let currentData = [];
let syncQueue = [];
let syncInProgress = false;

document.addEventListener('DOMContentLoaded', init);

function updateStatus(text, connected) {
    document.getElementById('syncLabel').textContent = text;
    document.getElementById('syncDot').className = 'dot ' + (connected ? 'connected' : 'disconnected');
}

function setupModalHandlers() {
    // Budget Book Modal
    const modal = document.getElementById('modal');
    const updateBudgetBtn = document.getElementById('updateBudgetBtn');
    const closeBtn = document.querySelector('.close');
    
    // Install Add-in Modal
    const installModal = document.getElementById('installModal');
    const installAddinBtn = document.getElementById('installAddinBtn');
    const closeInstallBtn = document.querySelector('.close-install');
    
    // Detect if running locally and update download links
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const windowsDownload = document.getElementById('windowsDownload');
    const macDownload = document.getElementById('macDownload');
    
    if (windowsDownload && macDownload) {
        if (isLocal) {
            windowsDownload.href = '/install-excel-addin-local.bat';
            macDownload.href = '/install-excel-addin-local.sh';
            console.log('🏠 Local development detected - using local installer files');
        } else {
            windowsDownload.href = '/install-excel-addin.bat';
            macDownload.href = '/install-excel-addin.sh';
            console.log('🌐 Production environment detected - using production installer files');
        }
    }
    
    // Open budget modal when button clicked
    if (updateBudgetBtn) {
        updateBudgetBtn.onclick = () => {
            modal.style.display = 'block';
            setTimeout(() => {
                modal.style.display = 'none';
            }, 3000); // Auto-close after 3 seconds
        };
    }
    
    // Close budget modal when X clicked
    if (closeBtn) {
        closeBtn.onclick = () => {
            modal.style.display = 'none';
        };
    }
    
    // Open install modal when button clicked
    if (installAddinBtn) {
        installAddinBtn.onclick = () => {
            installModal.style.display = 'block';
        };
    }
    
    // Close install modal when X clicked
    if (closeInstallBtn) {
        closeInstallBtn.onclick = () => {
            installModal.style.display = 'none';
        };
    }
    
    // Close modals when clicking outside
    window.onclick = (event) => {
        if (event.target === modal) {
            modal.style.display = 'none';
        }
        if (event.target === installModal) {
            installModal.style.display = 'none';
        }
    };
}

function renderTable(data) {
    const container = document.getElementById('spreadsheet');
    
    // Always show at least a 20x10 grid
    const minRows = 20;
    const minCols = 10;
    
    // Determine grid size
    const dataRows = data?.length || 0;
    const dataCols = dataRows > 0 ? Math.max(...data.map(row => row.length || 0)) : 0;
    const numRows = Math.max(minRows, dataRows);
    const numCols = Math.max(minCols, dataCols);
    
    let html = '<table class="data-table"><thead><tr>';
    
    // Empty corner cell
    html += '<th class="row-header"></th>';
    
    // Column headers (A, B, C, etc.)
    for (let c = 0; c < numCols; c++) {
        html += `<th>${String.fromCharCode(65 + c)}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Data rows
    for (let r = 0; r < numRows; r++) {
        html += '<tr>';
        // Row number
        html += `<th class="row-header">${r + 1}</th>`;
        // Data cells
        for (let c = 0; c < numCols; c++) {
            const value = data?.[r]?.[c] || '';
            html += `<td contenteditable="true" data-row="${r}" data-col="${c}" tabindex="0">${escapeHtml(value)}</td>`;
        }
        html += '</tr>';
    }
    
    html += '</tbody></table>';
    container.innerHTML = html;
    
    // Add event listeners to cells
    container.querySelectorAll('td[contenteditable]').forEach(cell => {
        cell.addEventListener('blur', handleCellEdit);
        cell.addEventListener('keydown', handleCellKeydown);
    });
}

function handleCellKeydown(e) {
    const cell = e.target;
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.col);
    
    // Enter key - move down or blur
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        const nextCell = document.querySelector(`td[data-row="${row + 1}"][data-col="${col}"]`);
        if (nextCell) {
            nextCell.focus();
        } else {
            cell.blur();
        }
        return;
    }
    
    // Arrow key navigation
    let targetRow = row;
    let targetCol = col;
    
    switch(e.key) {
        case 'ArrowUp':
            e.preventDefault();
            targetRow = Math.max(0, row - 1);
            break;
        case 'ArrowDown':
            e.preventDefault();
            targetRow = row + 1;
            break;
        case 'ArrowLeft':
            // Only navigate if cursor is at the start
            if (window.getSelection().toString().length === 0 && 
                cell.selectionStart === 0) {
                e.preventDefault();
                targetCol = Math.max(0, col - 1);
            } else {
                return;
            }
            break;
        case 'ArrowRight':
            // Only navigate if cursor is at the end
            if (window.getSelection().toString().length === 0 && 
                cell.selectionStart === cell.textContent.length) {
                e.preventDefault();
                targetCol = col + 1;
            } else {
                return;
            }
            break;
        case 'Tab':
            e.preventDefault();
            if (e.shiftKey) {
                targetCol = Math.max(0, col - 1);
            } else {
                targetCol = col + 1;
            }
            break;
        default:
            return;
    }
    
    // Move to target cell
    if (targetRow !== row || targetCol !== col) {
        const targetCell = document.querySelector(`td[data-row="${targetRow}"][data-col="${targetCol}"]`);
        if (targetCell) {
            targetCell.focus();
            // Select all content in the cell
            const range = document.createRange();
            range.selectNodeContents(targetCell);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    }
}

function handleCellEdit(e) {
    if (isUpdating) return;
    
    const row = parseInt(e.target.dataset.row);
    const col = parseInt(e.target.dataset.col);
    const newValue = e.target.textContent.trim();
    
    // Update current data immediately
    if (!currentData[row]) currentData[row] = [];
    currentData[row][col] = newValue;
    
    console.log(`Cell [${row},${col}] changed to: "${newValue}"`);
    
    // Queue sync immediately - queue handles batching
    queueSync();
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function readDataFromTable() {
    const data = [];
    let maxRow = -1, maxCol = -1;
    
    // Find actual data bounds
    for (let r = 0; r < currentData.length; r++) {
        if (currentData[r]) {
            for (let c = 0; c < currentData[r].length; c++) {
                if (currentData[r][c]) {
                    maxRow = Math.max(maxRow, r);
                    maxCol = Math.max(maxCol, c);
                }
            }
        }
    }
    
    if (maxRow === -1) return [];
    
    // Build clean data array
    for (let r = 0; r <= maxRow; r++) {
        const row = [];
        for (let c = 0; c <= maxCol; c++) {
            row.push(currentData[r]?.[c] || '');
        }
        data.push(row);
    }
    
    return data;
}

function queueSync() {
    const data = readDataFromTable();
    syncQueue.push(data);
    
    // Debounce the queue processing, not individual edits
    clearTimeout(window.queueTimeout);
    window.queueTimeout = setTimeout(processQueue, 400);
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
            body: JSON.stringify({ data, title: 'Excel Demo', type: 'web' })
        });
        
        if (!response.ok) {
            console.error('Sync failed:', response.status);
            updateStatus('Sync failed', false);
            syncInProgress = false;
            return;
        }
        
        updateStatus('✓ Synced', true);
    } catch (error) {
        console.error('Sync error:', error);
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
    queueSync();
}

function applyUpdate(data) {
    if (isUpdating) return;
    
    console.log('Applying update from Excel:', data.length, 'rows');
    isUpdating = true;
    
    currentData = JSON.parse(JSON.stringify(data)); // Deep clone
    renderTable(currentData);
    
    updateStatus('✓ Updated from Excel', true);
    
    setTimeout(() => { isUpdating = false; }, 2000);
}

let reconnectAttempts = 0;

function setupSSE() {
    if (eventSource) {
        try { eventSource.close(); } catch(e) {}
    }
    
    eventSource = new EventSource(`${SERVER_URL}/api/stream`);
    
    eventSource.addEventListener('message', (e) => {
        const msg = JSON.parse(e.data);
        if (msg.type === 'data-update' && msg.documentId === DOCUMENT_ID) {
            applyUpdate(msg.data);
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
    
    setupModalHandlers();
    
    // Manual sync button
    const btn = document.getElementById('manualSyncBtn');
    if (btn) {
        btn.onclick = async () => {
            console.log('Manual sync triggered');
            await sync();
        };
    }
    
    try {
        const res = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`, {
            headers: { 'Cache-Control': 'no-cache' }
        });
        
        if (res.ok) {
            const doc = await res.json();
            if (doc?.data?.length) {
                currentData = JSON.parse(JSON.stringify(doc.data));
                console.log('Loaded data from server:', currentData.length, 'rows');
            } else {
                console.log('No data on server');
                currentData = [];
            }
        }
    } catch (e) {
        console.error('Failed to load initial data:', e);
        updateStatus('Server offline', false);
        currentData = [];
    }
    
    renderTable(currentData);
    
    console.log('Setting up SSE');
    await new Promise(r => setTimeout(r, 1000));
    
    setupSSE();
    isUpdating = false;
    console.log('Ready');
}
