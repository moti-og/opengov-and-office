// Auto-detect server URL (production vs local)
const SERVER_URL = window.location.hostname === 'localhost' 
    ? 'http://localhost:3001' 
    : window.location.origin;

let lastUpdatedAt = null;

document.addEventListener('DOMContentLoaded', init);

async function init() {
    console.log('Budget Book page loading...');
    await loadBudgetTable();
    await loadWhatIsThisModal();
    setupInstallModal();
    setupLogoVideo();
    
    // Poll for updates every 2 seconds
    setInterval(checkForUpdates, 2000);
}

// ========== INSTALL MODAL ==========

function setupInstallModal() {
    const modal = document.getElementById('modal');
    const installBtn = document.getElementById('installAddinBtn');
    const closeBtn = document.querySelector('.close-install');
    
    if (installBtn) {
        installBtn.onclick = () => {
            modal.style.display = 'block';
        };
    }
    
    if (closeBtn) {
        closeBtn.onclick = () => {
            modal.style.display = 'none';
        };
    }
    
    window.addEventListener('click', (event) => {
        if (event.target === modal) {
            modal.style.display = 'none';
        }
    });
    
    // Detect local vs production
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    
    const windowsDownload = document.getElementById('windowsDownload');
    const macDownload = document.getElementById('macDownload');
    
    if (isLocal) {
        windowsDownload.href = '/install-excel-addin-local.bat';
        macDownload.href = '/install-excel-addin-local.sh';
    } else {
        windowsDownload.href = '/install-excel-addin.bat';
        macDownload.href = '/install-excel-addin.sh';
    }
}

// ========== LOGO VIDEO CONTROL ==========

function setupLogoVideo() {
    const video = document.getElementById('cityLogo');
    if (!video) return;
    
    function playAndScheduleNext() {
        // Play the video
        video.play().catch(err => console.log('Video play prevented:', err));
        
        // When video ends, wait random time before playing again
        video.onended = () => {
            // Random wait time between 30-59 seconds
            const waitTime = (Math.random() * 29 + 30) * 1000; // 30-59 seconds in ms
            console.log(`Logo video will replay in ${Math.round(waitTime / 1000)} seconds`);
            
            setTimeout(() => {
                playAndScheduleNext();
            }, waitTime);
        };
    }
    
    // Wait 10-20 seconds before first play
    const initialWaitTime = (Math.random() * 10 + 10) * 1000; // 10-20 seconds in ms
    console.log(`Logo video will start in ${Math.round(initialWaitTime / 1000)} seconds`);
    setTimeout(() => {
        playAndScheduleNext();
    }, initialWaitTime);
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
        closeBtn.onclick = () => {
            modal.style.display = 'none';
        };
        
        gotItBtn.onclick = () => {
            modal.style.display = 'none';
        };
        
        // Close when clicking outside
        window.onclick = (event) => {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        };
        
    } catch (error) {
        console.error('Error loading modal content:', error);
    }
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

