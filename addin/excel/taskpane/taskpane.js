const SERVER_URL = 'http://localhost:3001';
const DOCUMENT_ID = 'excel-demo-doc';

let eventSource = null;
let isConnected = false;

let lastSyncedData = null;

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
    range.load('values');
    await context.sync();
    return range.values;
  });
}

async function writeExcelData(data) {
  if (!data || !data.length) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange().clear();
    const range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
    range.values = data;
    await context.sync();
  });
}

async function syncExcelToServer() {
  try {
    const data = await readExcelData();
    
    // Don't sync if data hasn't changed
    if (JSON.stringify(data) === JSON.stringify(lastSyncedData)) {
      return;
    }

    updateStatus('Syncing...', true);

    const response = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data, title: 'Excel Demo', type: 'excel' })
    });

    if (response.ok) {
      lastSyncedData = data;
      updateStatus('✓ Synced', true);
      updateDataPreview(data);
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

    // Try to get existing document
    const docResponse = await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}`);

    if (docResponse.ok) {
      const doc = await docResponse.json();
      if (doc && doc.data && doc.data.length > 0) {
        await writeExcelData(doc.data);
        lastSyncedData = doc.data;
        updateDataPreview(doc.data);
      } else {
        // Create default data
        const defaultData = [
          ['Product', 'Q1 Sales', 'Q2 Sales', 'Q3 Sales'],
          ['Widget A', '1500', '2300', '2100'],
          ['Widget B', '2800', '3200', '3500'],
          ['Widget C', '1200', '1400', '1600']
        ];

        await fetch(`${SERVER_URL}/api/documents/${DOCUMENT_ID}/update`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            data: defaultData,
            title: 'Excel Demo',
            type: 'excel'
          })
        });

        await writeExcelData(defaultData);
        lastSyncedData = defaultData;
        updateDataPreview(defaultData);
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
        if (JSON.stringify(payload.data) === JSON.stringify(lastSyncedData)) {
          return;
        }
        
        console.log('Received update from web');
        await writeExcelData(payload.data);
        lastSyncedData = payload.data;
        updateDataPreview(payload.data);
        updateStatus('✓ Updated from web', true);
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
