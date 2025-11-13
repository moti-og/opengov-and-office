// Auto-detect server URL (production vs local)
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : window.location.origin;

let lastUpdatedAt = null;

document.addEventListener('DOMContentLoaded', init);

async function init() {
    console.log('Budget Book page loading...');
    await loadBudgetTable();
    
    // Poll for updates every 2 seconds
    setInterval(checkForUpdates, 2000);
}

async function checkForUpdates() {
    try {
        const response = await fetch(`${SERVER_URL}/api/budget-book`);
        if (!response.ok) return;
        
        const result = await response.json();
        
        // Only reload if we have a previous timestamp AND it changed
        if (lastUpdatedAt !== null && result.updatedAt && result.updatedAt !== lastUpdatedAt) {
            console.log('New budget data detected, reloading...');
            await loadBudgetTable();
        }
    } catch (error) {
        // Silently fail - don't spam console
    }
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
        
        // Check for new format (screenshots) or legacy (single image)
        const screenshots = result.screenshots || [];
        const legacyImage = result.image;
        
        if (screenshots.length === 0 && !legacyImage) {
            container.innerHTML = '<div class="loading">No budget data available. Use the Excel add-in to update the budget book.</div>';
            return;
        }
        
        // Display multiple screenshots or single legacy image
        let html = '';
        
        if (screenshots.length > 0) {
            // New format: multiple screenshots
            html = screenshots.map((screenshot, index) => `
                <div class="screenshot-section">
                    <h3 class="screenshot-title">${escapeHtml(screenshot.address)}</h3>
                    <img src="${screenshot.image}" alt="${escapeHtml(screenshot.address)}" class="budget-screenshot" />
                </div>
            `).join('');
        } else if (legacyImage) {
            // Legacy format: single image
            html = `<img src="${legacyImage}" alt="Budget Table" class="budget-screenshot" />`;
        }
        
        container.innerHTML = html;
        
        // Update timestamp and track it
        if (result.updatedAt) {
            const date = new Date(result.updatedAt);
            timestampEl.textContent = date.toLocaleString();
            lastUpdatedAt = result.updatedAt;
        }
        
        const count = screenshots.length || (legacyImage ? 1 : 0);
        console.log('Budget book loaded:', count, 'screenshot(s)');
        
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

