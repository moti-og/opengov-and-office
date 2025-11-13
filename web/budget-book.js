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
        
        if (!result.image) {
            container.innerHTML = '<div class="loading">No budget data available. Use the Excel add-in to update the budget book.</div>';
            return;
        }
        
        // Display the screenshot
        container.innerHTML = `<img src="${result.image}" alt="Budget Table" class="budget-screenshot" />`;
        
        // Update timestamp
        if (result.updatedAt) {
            const date = new Date(result.updatedAt);
            timestampEl.textContent = date.toLocaleString();
        }
        
        console.log('Budget screenshot loaded successfully');
        
    } catch (error) {
        console.error('Failed to load budget data:', error);
        container.innerHTML = '<div class="loading">Failed to load budget data. Please try again later.</div>';
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

