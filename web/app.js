// Auto-detect server URL (production vs local)
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : window.location.origin;
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isUpdating = false;
let currentRanges = [];  // Array of {address, data}
let currentCharts = [];
let syncQueue = [];
let syncInProgress = false;

document.addEventListener('DOMContentLoaded', init);

function updateStatus(text, connected) {
    // Update header sync status
    const syncLabel = document.getElementById('syncLabel');
    const syncDot = document.getElementById('syncDot');
    if (syncLabel) syncLabel.textContent = text;
    if (syncDot) syncDot.className = 'sync-dot ' + (connected ? 'connected' : 'disconnected');
    
    // Update sidepane sync status
    const sidepaneSyncLabel = document.getElementById('sidepaneSyncLabel');
    const sidepaneSyncDot = document.getElementById('sidepaneSyncDot');
    if (sidepaneSyncLabel) sidepaneSyncLabel.textContent = text;
    if (sidepaneSyncDot) sidepaneSyncDot.className = 'dot ' + (connected ? 'connected' : 'disconnected');
}

// ========== WHAT IS THIS MODAL ==========

async function loadWhatIsThisModal() {
    try {
        const response = await fetch(`${SERVER_URL}/budget-book-info.json`);
        if (!response.ok) {
            console.error('Failed to load modal content');
            return;
        }
        
        const config = await response.json();
        
        // Setup modal handlers
        const modal = document.getElementById('whatIsThisModal');
        const btn = document.getElementById('whatIsThisBtn');
        const closeBtn = document.querySelector('.info-modal-close');
        const gotItBtn = document.getElementById('infoModalBtn');
        
        if (!modal || !btn) return;
        
        // Populate modal content
        document.getElementById('infoModalTitle').textContent = config.title;
        document.getElementById('infoModalBtn').textContent = config.buttonText;
        
        const modalBody = document.getElementById('infoModalBody');
        modalBody.innerHTML = config.items.map(item => {
            let text = item.text;
            
            // If there's a link, wrap the linkText in an anchor
            if (item.link && item.linkText) {
                text = text.replace(item.linkText, `<a href="${item.link}" target="_blank" class="info-item-link">${item.linkText}</a>`);
            }
            
            return `
                <div class="info-item">
                    <span class="info-item-arrow">→</span>
                    <div class="info-item-content">
                        <span class="info-item-label">${item.label}</span>
                        <span class="info-item-text"> ${text}</span>
                    </div>
                </div>
            `;
        }).join('');
        
        // Open modal
        btn.onclick = () => {
            modal.style.display = 'block';
        };
        
        // Close modal handlers
        if (closeBtn) {
            closeBtn.onclick = () => {
                modal.style.display = 'none';
            };
        }
        
        if (gotItBtn) {
            gotItBtn.onclick = () => {
                modal.style.display = 'none';
            };
        }
        
        // Close when clicking outside
        window.addEventListener('click', (event) => {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        });
        
    } catch (error) {
        console.error('Error loading modal content:', error);
    }
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
    
    // Charts Modal
    const chartsModal = document.getElementById('chartsModal');
    const viewChartsBtn = document.getElementById('viewChartsBtn');
    const closeChartsBtn = document.querySelector('.close-charts');
    
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
            }, 3000);
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
    
    // Open charts modal when button clicked
    if (viewChartsBtn) {
        viewChartsBtn.onclick = () => {
            chartsModal.style.display = 'block';
        };
    }
    
    // Close charts modal when X clicked
    if (closeChartsBtn) {
        closeChartsBtn.onclick = () => {
            chartsModal.style.display = 'none';
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
        if (event.target === chartsModal) {
            chartsModal.style.display = 'none';
        }
    };
}

// ========== MULTIPLE RANGES RENDERING ==========

function renderRanges(ranges) {
    const container = document.getElementById('spreadsheet');
    
    // Always show a persistent 100x100 grid
    const numRows = 100;
    const numCols = 100;
    
    // Flatten ranges into a single data structure
    let flatData = [];
    if (ranges && ranges.length > 0) {
        // For now, just use the first range's data
        // (Multi-range display can be added later if needed)
        flatData = ranges[0]?.data || [];
    }
    
    let html = '<table class="data-table"><thead><tr>';
    
    // Empty corner cell
    html += '<th class="row-header"></th>';
    
    // Column headers (A, B, C, etc.)
    for (let c = 0; c < numCols; c++) {
        const colLabel = c < 26 
            ? String.fromCharCode(65 + c) 
            : String.fromCharCode(64 + Math.floor(c / 26)) + String.fromCharCode(65 + (c % 26));
        html += `<th>${colLabel}</th>`;
    }
    html += '</tr></thead><tbody>';
    
    // Data rows
    for (let r = 0; r < numRows; r++) {
        html += '<tr>';
        // Row number
        html += `<th class="row-header">${r + 1}</th>`;
        // Data cells
        for (let c = 0; c < numCols; c++) {
            const value = flatData[r]?.[c] || '';
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
            if (window.getSelection().toString().length === 0 && 
                cell.selectionStart === 0) {
                e.preventDefault();
                targetCol = Math.max(0, col - 1);
            } else {
                return;
            }
            break;
        case 'ArrowRight':
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
    
    // Update current ranges data immediately
    if (!currentRanges[0]) {
        currentRanges[0] = { address: 'A1:CV100', data: [] };
    }
    if (!currentRanges[0].data[row]) {
        currentRanges[0].data[row] = [];
    }
    currentRanges[0].data[row][col] = newValue;
    
    console.log(`Cell [${row},${col}] = "${newValue}"`);
    
    // Queue sync
    queueSync();
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function renderCharts(charts) {
    const container = document.getElementById('charts-container');
    const viewChartsBtn = document.getElementById('viewChartsBtn');
    const chartCount = document.getElementById('chartCount');
    
    if (!charts || charts.length === 0) {
        container.innerHTML = '<p style="text-align: center; color: #666;">No charts to display</p>';
        viewChartsBtn.style.display = 'none';
        return;
    }
    
    // Show button and update count
    viewChartsBtn.style.display = 'block';
    chartCount.textContent = charts.length;
    
    // Render charts in modal
    container.innerHTML = charts.map(chart => `
        <div class="chart">
            <h3>${escapeHtml(chart.name)}</h3>
            <img src="${chart.image}" alt="${escapeHtml(chart.name)}" />
        </div>
    `).join('');
}

// ========== SYNC ==========

function queueSync() {
    // Clone current ranges for queue
    const rangesSnapshot = JSON.parse(JSON.stringify(currentRanges));
    syncQueue.push(rangesSnapshot);
    
    // Debounce
    clearTimeout(window.queueTimeout);
    window.queueTimeout = setTimeout(processQueue, 400);
}

async function processQueue() {
    if (syncInProgress || syncQueue.length === 0) return;
    
    syncInProgress = true;
    const ranges = syncQueue[syncQueue.length - 1]; // Take latest
    syncQueue = []; // Clear queue
    
    try {
        console.log('Syncing to server:', ranges.length, 'ranges');
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ranges, title: 'Excel Demo', type: 'web' })
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
    
    // Process next item if queue filled up
    if (syncQueue.length > 0) {
        setTimeout(processQueue, 100);
    }
}

function applyUpdate(ranges, charts, sourceType) {
    if (isUpdating) return;
    
    // Ignore updates from web (prevent echo)
    if (sourceType === 'web') {
        console.log('Ignoring update - came from web (preventing echo)');
        return;
    }
    
    console.log('Applying update from Excel:', ranges?.length || 0, 'ranges,', charts?.length || 0, 'charts');
    isUpdating = true;
    
    currentRanges = JSON.parse(JSON.stringify(ranges || [])); // Deep clone
    if (charts) {
        currentCharts = JSON.parse(JSON.stringify(charts));
    }
    
    // Don't re-render if user is actively editing a cell (prevent cursor loss)
    const activeElement = document.activeElement;
    const isEditingCell = activeElement && activeElement.tagName === 'TD' && activeElement.isContentEditable;
    
    if (isEditingCell) {
        console.log('User is editing a cell - updating cells in place without re-render');
        const editingRange = parseInt(activeElement.dataset.range);
        const editingRow = parseInt(activeElement.dataset.row);
        const editingCol = parseInt(activeElement.dataset.col);
        
        // Update all cells except the one being edited
        document.querySelectorAll('td[contenteditable]').forEach(cell => {
            const rangeIndex = parseInt(cell.dataset.range);
            const row = parseInt(cell.dataset.row);
            const col = parseInt(cell.dataset.col);
            
            // Skip the cell being edited
            if (rangeIndex === editingRange && row === editingRow && col === editingCol) {
                return;
            }
            
            // Update the cell value
            const value = currentRanges[rangeIndex]?.data[row]?.[col] || '';
            if (cell.textContent !== value) {
                cell.textContent = value;
            }
        });
    } else {
        // No active editing, safe to re-render
        renderRanges(currentRanges);
    }
    
    // Always render charts
    renderCharts(currentCharts);
    
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
            // Use ranges if available, otherwise fall back to legacy data
            const ranges = msg.ranges || (msg.data ? [{ address: 'Legacy', data: msg.data }] : []);
            applyUpdate(ranges, msg.charts, msg.sourceType);
        }
    });
    
    eventSource.onopen = () => {
        reconnectAttempts = 0;
        updateStatus('Sync', true);
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
    await loadWhatIsThisModal();
    
    try {
        const res = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`, {
            headers: { 'Cache-Control': 'no-cache' }
        });
        
        if (res.ok) {
            const doc = await res.json();
            
            // Use ranges if available, otherwise fall back to legacy data
            if (doc?.ranges?.length) {
                currentRanges = JSON.parse(JSON.stringify(doc.ranges));
                console.log('Loaded ranges from server:', currentRanges.length, 'ranges');
            } else if (doc?.data?.length) {
                currentRanges = [{ address: 'Legacy', data: JSON.parse(JSON.stringify(doc.data)) }];
                console.log('Loaded legacy data from server');
            } else {
                console.log('No data on server');
                currentRanges = [];
            }
            
            if (doc?.charts?.length) {
                currentCharts = JSON.parse(JSON.stringify(doc.charts));
                console.log('Loaded charts from server:', currentCharts.length, 'charts');
            } else {
                currentCharts = [];
            }
        }
    } catch (e) {
        console.error('Failed to load initial data:', e);
        updateStatus('Server offline', false);
        currentRanges = [];
        currentCharts = [];
    }
    
    renderRanges(currentRanges);
    renderCharts(currentCharts);
    
    console.log('Setting up SSE');
    await new Promise(r => setTimeout(r, 1000));
    
    setupSSE();
    isUpdating = false;
    console.log('Ready');
}
