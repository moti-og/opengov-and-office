const SERVER_URL = 'http://localhost:3001';
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isConnected = false;

let lastSyncedData = null;
let isWritingFromSSE = false; // Flag to prevent sync loops when receiving SSE updates

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('manualSyncBtn').onclick = syncExcelToServer;
    initializeSync();
    setupChangeListener();
  }
});

function setupChangeListener() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Listen for cell changes
    sheet.onChanged.add(async (event) => {
      // Ignore changes while writing from SSE to prevent sync loops
      if (isWritingFromSSE) {
        console.log('Ignoring change - writing from SSE');
        return;
      }
      
      console.log('Cell changed, auto-syncing...');
      // Debounce: only sync if no changes for 1 second
      clearTimeout(window.autoSyncTimeout);
      window.autoSyncTimeout = setTimeout(async () => {
        await syncExcelToServer();
      }, 1000);
    });
    
    await context.sync();
  }).catch(error => {
    console.error('Error setting up change listener:', error);
  });
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

async function readExcelData() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load('values, rowCount, columnCount');
    await context.sync();
    
    const data = range.values;
    const rowCount = range.rowCount;
    const colCount = range.columnCount;
    
    // Read column widths
    const columnWidths = [];
    for (let c = 0; c < colCount; c++) {
      const col = sheet.getRangeByIndexes(0, c, 1, 1);
      col.load('format/columnWidth');
      columnWidths.push(col);
    }
    
    // Read row heights
    const rowHeights = [];
    for (let r = 0; r < rowCount; r++) {
      const row = sheet.getRangeByIndexes(r, 0, 1, 1);
      row.load('format/rowHeight');
      rowHeights.push(row);
    }
    
    await context.sync();
    
    return {
      data: data,
      layout: {
        columnWidths: columnWidths.map(col => col.format.columnWidth),
        rowHeights: rowHeights.map(row => row.format.rowHeight)
      }
    };
  });
}

async function writeExcelData(data, layout = null) {
  if (!data || !data.length) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
    range.values = data;
    
    // Apply column widths if provided
    if (layout && layout.columnWidths && layout.columnWidths.length > 0) {
      for (let c = 0; c < Math.min(layout.columnWidths.length, data[0].length); c++) {
        const col = sheet.getRangeByIndexes(0, c, 1, 1);
        col.format.columnWidth = layout.columnWidths[c];
      }
    }
    
    // Apply row heights if provided
    if (layout && layout.rowHeights && layout.rowHeights.length > 0) {
      for (let r = 0; r < Math.min(layout.rowHeights.length, data.length); r++) {
        const row = sheet.getRangeByIndexes(r, 0, 1, 1);
        row.format.rowHeight = layout.rowHeights[r];
      }
    }
    
    await context.sync();
  });
}

async function syncExcelToServer() {
  try {
    const result = await readExcelData();
    
    // Don't sync if data hasn't changed
    if (JSON.stringify(result) === JSON.stringify(lastSyncedData)) {
      return;
    }

    updateStatus('Syncing...', true);

    const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        data: result.data, 
        layout: result.layout,
        title: 'Excel Demo', 
        type: 'excel' 
      })
    });

    if (response.ok) {
      lastSyncedData = result;
      updateStatus('✓ Synced', true);
      updateDataPreview(result.data);
    } else {
      updateStatus('Error syncing', false);
    }
  } catch (error) {
    console.error('Error syncing:', error);
    updateStatus('Error syncing', false);
  }
}

async function refreshFromServer() {
  try {
    updateStatus('Refreshing from server...', true);
    const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);
    const doc = await response.json();

    if (doc && doc.data && doc.data.length) {
      await writeExcelData(doc.data);
      updateStatus('✓ Refreshed from server', true);
      updateDataPreview(doc.data);
    } else {
      updateStatus('No data found', true);
    }
  } catch (error) {
    console.error('Error refreshing:', error);
    updateStatus('Error refreshing', false);
  }
}

async function initializeSync() {
  try {
    updateStatus('Connecting to server...', false);

    // Test server connection
    const healthResponse = await fetch(`${SERVER_URL}/api/health`);
    if (!healthResponse.ok) {
      updateStatus('Server not responding', false);
      return;
    }

    // Read current Excel data first (don't overwrite user's work!)
    const currentExcelData = await readExcelData();
    console.log('Current Excel data on startup:', currentExcelData);

    // Check if Excel has data
    const hasExcelData = currentExcelData.data && currentExcelData.data.length > 0 && 
                         currentExcelData.data.some(row => row.some(cell => cell !== '' && cell !== null));

    // Try to get existing document from server
    const docResponse = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);

    if (docResponse.ok) {
      const doc = await docResponse.json();
      const hasServerData = doc && doc.data && doc.data.length > 0;

      if (hasExcelData) {
        // Excel has data - sync it to server (preserve user's work!)
        console.log('Excel has data, syncing to server');
        await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            data: currentExcelData.data,
            layout: currentExcelData.layout,
            title: 'Excel Demo',
            type: 'excel'
          })
        });
        lastSyncedData = currentExcelData;
        updateDataPreview(currentExcelData.data);
      } else if (hasServerData) {
        // Excel is empty, but server has data - load from server
        console.log('Excel empty, loading from server');
        await writeExcelData(doc.data, doc.layout);
        lastSyncedData = { data: doc.data, layout: doc.layout };
        updateDataPreview(doc.data);
      } else {
        // Both empty - do nothing, user can start fresh
        console.log('Both Excel and server empty, starting fresh');
        lastSyncedData = { data: [], layout: { columnWidths: [], rowHeights: [] } };
      }
    }

    updateStatus('✓ Connected', true);
    setupSSE();
  } catch (error) {
    console.error('Error initializing:', error);
    updateStatus('Connection failed - Make sure server is running on port 3001', false);
  }
}

function setupSSE() {
  if (eventSource) {
    eventSource.close();
  }

  eventSource = new EventSource(`${SERVER_URL}/api/stream`);

  eventSource.addEventListener('message', async (event) => {
    try {
      const payload = JSON.parse(event.data);
      if (payload.type === 'data-update' && payload.documentId === DOCUMENT_ID) {
        // Don't update if data is same
        const newData = { data: payload.data, layout: payload.layout };
        if (JSON.stringify(newData) === JSON.stringify(lastSyncedData)) {
          return;
        }
        
        console.log('Received update from web, disabling onChanged temporarily');
        
        // Disable onChanged handler to prevent sync loop
        isWritingFromSSE = true;
        
        try {
          await writeExcelData(payload.data, payload.layout);
          lastSyncedData = newData;
          updateDataPreview(payload.data);
          updateStatus('✓ Updated from web', true);
        } finally {
          // Re-enable after a delay (let pending events finish)
          setTimeout(() => {
            isWritingFromSSE = false;
            console.log('Re-enabled onChanged handler');
          }, 1500);
        }
      }
    } catch (error) {
      // Silently ignore cell-editing errors
      if (error.message && error.message.includes('cell-editing mode')) {
        console.log('Skipping update - Excel is in edit mode');
        return;
      }
      console.error('Error handling SSE:', error);
    }
  });

  eventSource.onopen = () => {
    console.log('SSE connection opened');
    updateStatus('✓ Live sync active', true);
  };

  eventSource.onerror = (error) => {
    console.error('SSE error:', error);
    updateStatus('Connection lost, reconnecting...', false);
    eventSource.close();
    setTimeout(setupSSE, 5000);
  };
}

function updateDataPreview(data) {
  const preview = document.getElementById('dataPreview');
  if (!data || !data.length) {
    preview.innerHTML = '<em>No data</em>';
    preview.className = 'loading';
    return;
  }

  const rows = data.slice(0, 5);
  const html = rows
    .map(row => row.slice(0, 4).join(' | '))
    .join('<br>');

  preview.innerHTML = `<strong>Preview (${data.length} rows):</strong><br>${html}${
    data.length > 5 ? '<br>...' : ''
  }`;
  preview.className = '';
}
