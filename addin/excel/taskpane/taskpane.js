// Auto-detect server URL based on where the add-in is loaded from
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : 'https://excelftw.onrender.com';
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
        setupDragDrop();
        setupRangeManagement();
    }
});

function setupModalHandlers() {
    const modal = document.getElementById('modal');
    const closeBtn = document.querySelector('.close');
    
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

// ============ RANGE MANAGEMENT ============

function setupRangeManagement() {
    // Add spreadsheet range
    document.getElementById('addSpreadsheetRange').onclick = () => {
        addRangeToList('spreadsheetRanges');
    };
    
    // Add budget range
    document.getElementById('addBudgetRange').onclick = () => {
        addRangeToList('budgetRanges');
    };
    
    // Sync spreadsheet button
    document.getElementById('syncSpreadsheetBtn').onclick = async () => {
        await syncSpreadsheet();
    };
    
    // Update budget book button
    document.getElementById('updateBudgetBtn').onclick = async () => {
        await updateBudgetBook();
    };
    
    // Setup remove buttons
    setupRemoveButtons();
}

function addRangeToList(listId) {
    const list = document.getElementById(listId);
    const index = list.children.length;
    
    const rangeItem = document.createElement('div');
    rangeItem.className = 'range-item';
    rangeItem.setAttribute('data-index', index);
    rangeItem.draggable = true;
    
    rangeItem.innerHTML = `
        <span class="drag-handle">≡</span>
        <input type="text" class="range-input" placeholder="e.g. A1:F10" value="" />
        <button class="remove-range-btn" title="Remove range">×</button>
    `;
    
    list.appendChild(rangeItem);
    
    // Setup drag/drop for new item
    setupDragDrop();
    setupRemoveButtons();
}

function setupRemoveButtons() {
    document.querySelectorAll('.remove-range-btn').forEach(btn => {
        btn.onclick = (e) => {
            const rangeItem = e.target.closest('.range-item');
            const list = rangeItem.parentElement;
            
            // Don't remove if it's the last one
            if (list.children.length > 1) {
                rangeItem.remove();
                reindexList(list);
            }
        };
    });
}

function reindexList(list) {
    Array.from(list.children).forEach((item, index) => {
        item.setAttribute('data-index', index);
    });
}

function getRangesFromList(listId) {
    const list = document.getElementById(listId);
    const ranges = [];
    
    Array.from(list.children).forEach(item => {
        const input = item.querySelector('.range-input');
        const value = input.value.trim().toUpperCase();
        if (value) {
            ranges.push(value);
        }
    });
    
    return ranges;
}

// ============ DRAG & DROP ============

let draggedElement = null;

function setupDragDrop() {
    document.querySelectorAll('.range-item').forEach(item => {
        item.draggable = true;
        
        item.ondragstart = (e) => {
            draggedElement = item;
            item.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        };
        
        item.ondragend = (e) => {
            item.classList.remove('dragging');
            draggedElement = null;
        };
        
        item.ondragover = (e) => {
            e.preventDefault();
            const list = item.parentElement;
            const afterElement = getDragAfterElement(list, e.clientY);
            
            if (afterElement == null) {
                list.appendChild(draggedElement);
            } else {
                list.insertBefore(draggedElement, afterElement);
            }
            
            reindexList(list);
        };
    });
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.range-item:not(.dragging)')];
    
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

// ============ SPREADSHEET SYNC ============

async function syncSpreadsheet() {
    const ranges = getRangesFromList('spreadsheetRanges');
    
    if (ranges.length === 0) {
        showModal('⚠️', 'Please add at least one range to sync', 'warning');
        return;
    }
    
    console.log('Syncing spreadsheet ranges:', ranges);
    showModal('⏳', 'Syncing ranges to spreadsheet...', 'info');
    
    try {
        // Read data from all specified ranges
        const rangeData = await readMultipleRanges(ranges);
        
        // Get charts (keep existing behavior)
        const charts = await getCharts();
        
        // Send to server
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                ranges: rangeData,
                charts,
                title: 'Excel Demo', 
                type: 'excel' 
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        
        showModal('✅', 'Successfully synced to spreadsheet!', 'success');
        updateStatus('✓ Synced', true);
        
    } catch (error) {
        console.error('Sync failed:', error);
        showModal('❌', 'Failed to sync. Please try again.', 'error');
        updateStatus('Sync failed', false);
    }
}

async function readMultipleRanges(rangeAddresses) {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const results = [];
        
        for (const address of rangeAddresses) {
            try {
                const range = sheet.getRange(address);
                range.load('values, address');
                await context.sync();
                
                results.push({
                    address: address,
                    data: range.values.map(row => row.map(cell => cell ? String(cell) : ''))
                });
            } catch (err) {
                console.error(`Failed to read range ${address}:`, err);
                // Skip invalid ranges
            }
        }
        
        return results;
    });
}

async function writeMultipleRanges(rangeData) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        for (const item of rangeData) {
            try {
                const range = sheet.getRange(item.address);
                range.values = item.data;
            } catch (err) {
                console.error(`Failed to write range ${item.address}:`, err);
            }
        }
        
        await context.sync();
    });
}

// ============ BUDGET BOOK ============

async function updateBudgetBook() {
    const ranges = getRangesFromList('budgetRanges');
    
    if (ranges.length === 0) {
        showModal('⚠️', 'Please add at least one range to capture', 'warning');
        return;
    }
    
    console.log('Capturing budget book ranges:', ranges);
    showModal('⏳', 'Capturing screenshots...', 'info');
    
    try {
        const screenshots = [];
        
        // Capture each range as a screenshot
        for (const rangeAddress of ranges) {
            const image = await captureRangeScreenshot(rangeAddress);
            if (image) {
                screenshots.push({
                    address: rangeAddress,
                    image: image
                });
            }
        }
        
        if (screenshots.length === 0) {
            showModal('⚠️', 'Failed to capture any screenshots', 'warning');
            return;
        }
        
        console.log(`Captured ${screenshots.length} screenshots`);
        
        // Send to budget book API
        const response = await fetch(`${SERVER_URL}/api/budget-book/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ screenshots })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        
        showModal('✅', `Successfully updated budget book with ${screenshots.length} section(s)!`, 'success');
        
    } catch (error) {
        console.error('Failed to update budget book:', error);
        showModal('❌', 'Failed to update budget book. Please try again.', 'error');
    }
}

async function captureRangeScreenshot(rangeAddress) {
    try {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(rangeAddress);
            
            // Capture as image (width, height, fittingMode)
            const rangeImage = range.getImage(1200, 800, Excel.ImageFittingMode.fit);
            await context.sync();
            
            // Add data URI prefix if not present
            let imageData = rangeImage.value;
            if (imageData && !imageData.startsWith('data:')) {
                imageData = 'data:image/png;base64,' + imageData;
            }
            
            return imageData;
        });
    } catch (error) {
        console.error(`Failed to capture ${rangeAddress}:`, error);
        return null;
    }
}

// ============ MODAL HELPERS ============

function showModal(icon, message, type) {
    const modal = document.getElementById('modal');
    const modalIcon = modal.querySelector('.modal-icon');
    const modalText = modal.querySelector('p');
    const modalTitle = modal.querySelector('h3');
    
    modalIcon.textContent = icon;
    modalText.textContent = message;
    
    // Update title based on type
    const titles = {
        success: 'Success!',
        error: 'Error',
        warning: 'Warning',
        info: 'Processing...'
    };
    modalTitle.textContent = titles[type] || 'Notification';
    
    // Update title color
    const colors = {
        success: '#28a745',
        error: '#dc3545',
        warning: '#ffc107',
        info: '#0078d4'
    };
    modalTitle.style.color = colors[type] || '#333';
    
    modal.style.display = 'block';
    
    // Auto-close after 3 seconds unless it's an error
    if (type !== 'info') {
        setTimeout(() => {
            modal.style.display = 'none';
        }, 3000);
    }
}

// ============ LEGACY FUNCTIONS (for auto-sync) ============

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

async function getCharts() {
    try {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const charts = sheet.charts;
            charts.load('items/name, items/id');
            await context.sync();
            
            console.log('Found', charts.items.length, 'charts on active sheet');
            
            const chartImages = [];
            for (let chart of charts.items) {
                const image = chart.getImage(600, 400, Excel.ImageFittingMode.fit);
                await context.sync();
                
                // Add data URI prefix if not present
                let imageData = image.value;
                if (imageData && !imageData.startsWith('data:')) {
                    imageData = 'data:image/png;base64,' + imageData;
                }
                
                console.log('Captured chart:', chart.name, 'image length:', imageData?.length);
                
                chartImages.push({
                    name: chart.name,
                    image: imageData
                });
            }
            return chartImages;
        });
    } catch (err) {
        console.error('Error getting charts:', err);
        return [];
    }
}

async function queueSync() {
    try {
        const ranges = getRangesFromList('spreadsheetRanges');
        
        // Only sync if ranges are specified
        if (ranges.length === 0) {
            console.log('No ranges specified - skipping auto-sync');
            return;
        }
        
        const data = await readMultipleRanges(ranges);
        syncQueue.push(data);
        
        // Debounce the queue processing
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
    const rangeData = syncQueue[syncQueue.length - 1]; // Take latest
    syncQueue = []; // Clear queue
    
    try {
        const charts = await getCharts();
        console.log('Auto-syncing to server:', rangeData.length, 'ranges,', charts.length, 'charts');
        
        const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ranges: rangeData, charts, title: 'Excel Demo', type: 'excel' })
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
    
    // Process next item if queue filled up
    if (syncQueue.length > 0) {
        setTimeout(processQueue, 100);
    }
}

async function applyUpdate(rangeData, sourceType) {
    if (isUpdating) {
        console.log('Already updating');
        return;
    }
    
    // Ignore updates that came from Excel (prevent echo)
    if (sourceType === 'excel') {
        console.log('Ignoring update - came from Excel (preventing echo)');
        return;
    }
    
    console.log('Applying update from web:', rangeData?.length, 'ranges');
    isUpdating = true;
    
    try {
        await writeMultipleRanges(rangeData);
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
            await applyUpdate(msg.ranges, msg.sourceType);
        }
    });
    
    eventSource.onopen = () => {
        reconnectAttempts = 0;
        updateStatus('✓ Sync health', true);
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
