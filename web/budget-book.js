// Auto-detect server URL (production vs local)
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : window.location.origin;

document.addEventListener('DOMContentLoaded', init);

async function init() {
    console.log('Budget Book page loading...');
    await loadBudgetTable();
}

async function loadBudgetTable() {
    const container = document.getElementById('budget-table');
    const timestampEl = document.getElementById('update-timestamp');
    
    try {
        container.innerHTML = '<div class="loading">Loading budget data...</div>';
        
        const response = await fetch(`${SERVER_URL}/api/budget-book`);
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        
        const result = await response.json();
        
        if (!result.data || result.data.length === 0) {
            container.innerHTML = '<div class="loading">No budget data available. Use the Excel add-in to update the budget book.</div>';
            return;
        }
        
        // Render the table
        renderTable(result.data);
        
        // Update timestamp
        if (result.updatedAt) {
            const date = new Date(result.updatedAt);
            timestampEl.textContent = date.toLocaleString();
        }
        
        console.log('Budget table loaded successfully');
        
    } catch (error) {
        console.error('Failed to load budget data:', error);
        container.innerHTML = '<div class="loading">Failed to load budget data. Please try again later.</div>';
    }
}

function renderTable(data) {
    const container = document.getElementById('budget-table');
    
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="loading">No data to display</div>';
        return;
    }
    
    let html = '<table>';
    
    // Render each row
    data.forEach((row, rowIndex) => {
        if (rowIndex === 0) {
            // Header row
            html += '<thead><tr>';
            row.forEach(cell => {
                html += `<th>${escapeHtml(cell || '')}</th>`;
            });
            html += '</tr></thead><tbody>';
        } else {
            // Data rows
            html += '<tr>';
            row.forEach((cell, colIndex) => {
                const value = cell || '';
                // Format numeric values
                const formatted = colIndex === 0 ? value : formatValue(value);
                html += `<td>${escapeHtml(formatted)}</td>`;
            });
            html += '</tr>';
        }
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

function formatValue(value) {
    if (!value || value === '') return '';
    
    // Try to parse as number
    const cleanValue = String(value).replace(/[,$()]/g, '').trim();
    const num = parseFloat(cleanValue);
    
    if (isNaN(num)) return value; // Not a number, return as-is
    
    // Check if original had parentheses (negative)
    const isNegative = String(value).includes('(') || num < 0;
    
    // Format as currency
    const formatted = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(Math.abs(num));
    
    // Return with parentheses if negative
    return isNegative ? `(${formatted})` : formatted;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

