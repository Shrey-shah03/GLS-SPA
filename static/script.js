// Global workspace state
let workspaceItems = [];
let uploadFileInfo = {};

// Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const metaCard = document.getElementById('meta-card');
const statsArea = document.getElementById('stats-area');
const workspaceArea = document.getElementById('workspace-area');
const tableBody = document.getElementById('table-body');

// Form inputs
const clientNameInput = document.getElementById('client-name');
const iwoNoInput = document.getElementById('iwo-number');
const iwoDateInput = document.getElementById('iwo-date');
const deliveryDateInput = document.getElementById('delivery-date');
const salesNameInput = document.getElementById('sales-name');

// Setup today's date in form
const today = new Date();
const dd = String(today.getDate()).padStart(2, '0');
const mm = String(today.getMonth() + 1).padStart(2, '0');
const yy = String(today.getFullYear()).slice(-2);
iwoDateInput.value = `${dd}/${mm}/${yy}`;

// Set delivery date to 20 days from now as default
const deliveryDay = new Date(today.getTime() + 20 * 24 * 60 * 60 * 1000);
const d_dd = String(deliveryDay.getDate()).padStart(2, '0');
const d_mm = String(deliveryDay.getMonth() + 1).padStart(2, '0');
const d_yy = String(deliveryDay.getFullYear()).slice(-2);
deliveryDateInput.value = `${d_dd}/${d_mm}/${d_yy}`;

// Click event listener to open file explorer
dropZone.addEventListener('click', (e) => {
    if (e.target !== fileInput) {
        fileInput.click();
    }
});

// Drag and drop event listeners
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileUpload(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (fileInput.files.length > 0) {
        handleFileUpload(fileInput.files[0]);
    }
});

function handleFileUpload(file) {
    if (!file.name.endsWith('.xlsx')) {
        alert('Please upload a valid Excel spreadsheet (.xlsx)');
        return;
    }
    
    // Show uploading visual state
    dropZone.innerHTML = `
        <div class="upload-icon"><i class="fa-solid fa-spinner fa-spin"></i></div>
        <h3>Parsing BOQ Spreadsheet...</h3>
        <p>Running matching heuristics against catalogue database...</p>
    `;
    
    const formData = new FormData();
    formData.append('file', file);
    
    fetch('/api/upload-boq', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.error) {
            alert(data.error);
            resetDropZone();
            return;
        }
        
        workspaceItems = data.items;
        uploadFileInfo = {
            filename: data.filename,
            projectName: data.project_name
        };
        
        // Prefill metadata
        clientNameInput.value = data.project_name || 'Food Square';
        // Auto-generate standard IWO code e.g. P176
        const randomNum = Math.floor(100 + Math.random() * 900);
        iwoNoInput.value = `P${randomNum}`;
        
        // Update dashboard UI
        metaCard.classList.remove('disabled');
        const invoiceCard = document.getElementById('invoice-card');
        if (invoiceCard) {
            invoiceCard.classList.remove('disabled');
            document.getElementById('invoice-date').value = `${dd}/${mm}/${yy}`;
            document.getElementById('invoice-no').value = `SOR/26-27/00${randomNum}`;
        }
        statsArea.classList.remove('hidden');
        workspaceArea.classList.remove('hidden');
        
        // Render stats and table
        updateStats();
        renderTable();
        
        // Restore drop zone to mini state
        dropZone.innerHTML = `
            <div class="upload-icon" style="font-size: 2rem; margin-bottom: 0.5rem;"><i class="fa-solid fa-circle-check" style="color: var(--color-success);"></i></div>
            <h3>File Uploaded Successfully</h3>
            <p>${file.name} (${workspaceItems.length} rows parsed)</p>
            <div class="supported-formats" onclick="document.getElementById('file-input').click()">Upload Different File</div>
        `;
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred: ' + (error.message || error));
        resetDropZone();
    });
}

function resetDropZone() {
    dropZone.innerHTML = `
        <input type="file" id="file-input" accept=".xlsx" style="display: none;">
        <div class="upload-icon"><i class="fa-solid fa-file-excel"></i></div>
        <h3>Drag & Drop BOQ Excel here</h3>
        <p>or <span class="highlight-link" onclick="document.getElementById('file-input').click()">browse your files</span></p>
        <div class="supported-formats">Supports .xlsx spreadsheets</div>
    `;
    // Rebind input listener since elements are replaced
    const newInput = document.getElementById('file-input');
    newInput.addEventListener('change', (e) => {
        if (newInput.files.length > 0) {
            handleFileUpload(newInput.files[0]);
        }
    });
}

function updateStats() {
    const totalItems = workspaceItems.length;
    const exactMatches = workspaceItems.filter(item => item.matched_by === 'exact_catalog_specs').length;
    const searchMatches = workspaceItems.filter(item => item.matched_by.startsWith('catalog_text_search')).length;
    
    let totalQty = 0;
    workspaceItems.forEach(item => {
        const qty = parseInt(item.boq_qty) || 0;
        totalQty += qty;
    });
    
    document.getElementById('stat-total-items').innerText = totalItems;
    document.getElementById('stat-exact-matches').innerText = exactMatches;
    document.getElementById('stat-search-matches').innerText = searchMatches;
    document.getElementById('stat-total-qty').innerText = totalQty;
}

function renderTable() {
    tableBody.innerHTML = '';
    
    workspaceItems.forEach(item => {
        const tr = document.createElement('tr');
        tr.id = `row-${item.id}`;
        
        // Build status badge
        let statusBadge = '';
        if (item.matched_by === 'exact_catalog_specs') {
            statusBadge = `<span class="status-badge status-exact" title="Exact match from catalog database"><i class="fa-solid fa-circle-check"></i> pg ${item.page}</span>`;
        } else if (item.matched_by.startsWith('catalog_text_search')) {
            statusBadge = `<span class="status-badge status-fuzzy" title="Page matched using text indexing"><i class="fa-solid fa-circle-notch"></i> pg ${item.page}</span>`;
        } else {
            statusBadge = `<span class="status-badge status-fallback" title="Generic specs fallback"><i class="fa-solid fa-circle-question"></i> fb</span>`;
        }
        
        // Build actions
        let calculatorBtn = '';
        if (item.unit.toLowerCase() === 'mtr') {
            calculatorBtn = `
                <button class="btn btn-secondary btn-action" onclick="openDriverCalculator(${item.id})" title="Calculate Drivers for Linear Strip">
                    <i class="fa-solid fa-calculator"></i>
                </button>
            `;
        }
        
        tr.innerHTML = `
            <td>${item.id}</td>
            <td style="font-size: 0.72rem; white-space: nowrap; vertical-align: middle;">
                <div style="font-family: monospace; font-weight: 600; color: #ffffff;">${item.boq_description}</div>
                <div style="margin-top: 0.25rem;">${statusBadge}</div>
            </td>
            <td>
                <input type="text" class="table-input" value="${item.gls_code || ''}" onchange="updateItemField(${item.id}, 'gls_code', this.value)">
            </td>
            <td>
                <textarea class="table-input" rows="2" style="resize:vertical; min-width:220px;" onchange="updateItemField(${item.id}, 'product_description', this.value)">${item.product_description || ''}</textarea>
            </td>
            <td>
                <input type="text" class="table-input" list="colors-list" value="${item.body_color || ''}" onchange="updateItemField(${item.id}, 'body_color', this.value)">
            </td>
            <td style="width: 100px;">
                <input type="number" class="table-input" value="${item.boq_qty || 0}" onchange="updateItemField(${item.id}, 'boq_qty', this.value); updateStats();">
            </td>
            <td style="width: 110px;">
                <select class="table-select" onchange="updateItemField(${item.id}, 'unit', this.value); renderTable();">
                    <option value="Nos" ${item.unit === 'Nos' ? 'selected' : ''}>Nos</option>
                    <option value="Mtr" ${item.unit === 'Mtr' ? 'selected' : ''}>Mtr</option>
                    <option value="Set" ${item.unit === 'Set' ? 'selected' : ''}>Set</option>
                </select>
            </td>
            <td style="width: 100px;">
                <input type="number" class="table-input" value="${item.rate || 1200}" onchange="updateItemField(${item.id}, 'rate', this.value)">
            </td>
            <td>
                <input type="text" class="table-input" value="${item.driver_details || ''}" onchange="updateItemField(${item.id}, 'driver_details', this.value)">
            </td>
            <td style="width: 90px;">
                <input type="number" class="table-input" value="${item.driver_qty || 0}" onchange="updateItemField(${item.id}, 'driver_qty', this.value)">
            </td>
            <td>
                <input type="text" class="table-input" value="${item.led_details || ''}" onchange="updateItemField(${item.id}, 'led_details', this.value)">
            </td>
            <td>
                <input type="text" class="table-input" value="${item.accessories || ''}" onchange="updateItemField(${item.id}, 'accessories', this.value)">
            </td>
            <td>
                <div class="action-group">
                    <button class="btn btn-secondary btn-action" onclick="openRemapModal(${item.id})" title="Search Catalog to Remap Specs">
                        <i class="fa-solid fa-search"></i>
                    </button>
                    ${calculatorBtn}
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

function updateItemField(itemId, field, value) {
    const item = workspaceItems.find(i => i.id === itemId);
    if (item) {
        item[field] = value;
    }
}

// Remap Modal logic
let currentEditingRowId = null;

function openRemapModal(rowId) {
    currentEditingRowId = rowId;
    document.getElementById('catalog-search-input').value = '';
    document.getElementById('search-results-list').innerHTML = `
        <div class="empty-state">Enter a term and click search to view matching catalog pages</div>
    `;
    document.getElementById('remap-modal').classList.remove('hidden');
}

function searchCatalog() {
    const query = document.getElementById('catalog-search-input').value.trim();
    if (!query) return;
    
    const resultsContainer = document.getElementById('search-results-list');
    resultsContainer.innerHTML = `<div class="empty-state"><i class="fa-solid fa-spinner fa-spin"></i> Searching database...</div>`;
    
    fetch(`/api/search-catalog?q=${encodeURIComponent(query)}`)
    .then(r => r.json())
    .then(results => {
        resultsContainer.innerHTML = '';
        if (results.length === 0) {
            resultsContainer.innerHTML = `<div class="empty-state">No matching pages found in catalog database.</div>`;
            return;
        }
        
        results.forEach(res => {
            const div = document.createElement('div');
            div.className = 'search-item';
            div.onclick = () => selectCatalogPage(res.page);
            div.innerHTML = `
                <div class="search-item-header">
                    <span>Catalog pg ${res.page}</span>
                    <span><i class="fa-solid fa-link"></i> Bind pg</span>
                </div>
                <div class="search-item-snippet">${res.snippet}</div>
            `;
            resultsContainer.appendChild(div);
        });
    })
    .catch(err => {
        console.error(err);
        resultsContainer.innerHTML = `<div class="empty-state">Error searching catalog.</div>`;
    });
}

function selectCatalogPage(pageNum) {
    // Generate text matching page bound and update workspaceItems
    fetch(`/api/search-catalog?q=Page%20${pageNum}`) // retrieve page context
    
    // Close modal
    closeModal('remap-modal');
    
    // Simulate catalog update for this item
    const item = workspaceItems.find(i => i.id === currentEditingRowId);
    if (item) {
        item.page = pageNum;
        item.matched_by = `catalog_text_search_page_${pageNum}`;
        
        // Generate mock description update from page
        item.product_description = `Matched with specifications on Catalog pg ${pageNum}. Verify dimensions on pg.`;
        
        renderTable();
        updateStats();
    }
}

// Driver calculation helper
function openDriverCalculator(rowId) {
    currentEditingRowId = rowId;
    const item = workspaceItems.find(i => i.id === rowId);
    if (!item) return;
    
    document.getElementById('driver-helper-row-id').value = rowId;
    document.getElementById('driver-strip-length').value = parseFloat(item.boq_qty) || 20;
    
    calculateDrivers();
    
    // Bind calculator input listeners
    document.getElementById('driver-strip-length').oninput = calculateDrivers;
    document.getElementById('driver-strip-wattage').oninput = calculateDrivers;
    document.getElementById('driver-split-wattage').onchange = calculateDrivers;
    
    document.getElementById('driver-modal').classList.remove('hidden');
}

function calculateDrivers() {
    const length = parseFloat(document.getElementById('driver-strip-length').value) || 0;
    const wattPerMeter = parseFloat(document.getElementById('driver-strip-wattage').value) || 0;
    const driverCap = parseFloat(document.getElementById('driver-split-wattage').value) || 150;
    
    const totalWatts = length * wattPerMeter;
    // 80% safety load recommend
    const recommendedLoad = driverCap * 0.8;
    const driverQty = Math.ceil(totalWatts / recommendedLoad) || 1;
    
    const resultContainer = document.getElementById('driver-calc-result');
    resultContainer.innerHTML = `
        <div class="calc-row">
            <span>Total Load:</span>
            <span>${totalWatts.toFixed(1)} Watts</span>
        </div>
        <div class="calc-row">
            <span>Recommended Max Load per Driver (80%):</span>
            <span>${recommendedLoad.toFixed(1)} Watts</span>
        </div>
        <div class="calc-row calc-highlight">
            <span>Calculated Drivers Count:</span>
            <span>${driverQty} Nos of ${driverCap}W Drivers</span>
        </div>
    `;
    
    // Save calculations temporarily on calculator element attributes
    resultContainer.setAttribute('data-qty', driverQty);
    resultContainer.setAttribute('data-wattage', `24V/${driverCap}W`);
}

function applyDriverCalculation() {
    const rowId = parseInt(document.getElementById('driver-helper-row-id').value);
    const resultContainer = document.getElementById('driver-calc-result');
    const qty = parseInt(resultContainer.getAttribute('data-qty')) || 1;
    const wattage = resultContainer.getAttribute('data-wattage') || '24V/150W';
    
    const item = workspaceItems.find(i => i.id === rowId);
    if (item) {
        item.driver_details = `Constant Voltage 24V - ${wattage}`;
        item.driver_qty = qty;
        
        renderTable();
        closeModal('driver-modal');
    }
}

function closeModal(id) {
    document.getElementById(id).classList.add('hidden');
}

function resetWorkspace() {
    if (confirm('Are you sure you want to clear current edits and reload the table?')) {
        renderTable();
        updateStats();
    }
}

function generateIWO() {
    if (workspaceItems.length === 0) {
        alert('Workspace is empty. Please upload a BOQ sheet.');
        return;
    }
    
    const payload = {
        client_name: clientNameInput.value.trim(),
        iwo_no: iwoNoInput.value.trim(),
        iwo_date: iwoDateInput.value.trim(),
        delivery_date: deliveryDateInput.value.trim(),
        sales_name: salesNameInput.value.trim(),
        items: workspaceItems
    };
    
    // Change button to spinner
    const exportBtn = document.querySelector('.workspace-actions .btn-primary');
    const originalText = exportBtn.innerHTML;
    exportBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Generating Excel...`;
    exportBtn.disabled = true;
    
    fetch('/api/generate-iwo', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
    })
    .then(r => r.json())
    .then(data => {
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        
        if (data.error) {
            alert(data.error);
            return;
        }
        
        // Trigger download
        window.location.href = data.download_url;
    })
    .catch(err => {
        console.error(err);
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        alert('Failed to generate IWO workbook.');
    });
}

function addNewBOQRow() {
    const nextId = workspaceItems.length > 0 ? Math.max(...workspaceItems.map(i => i.id)) + 1 : 1;
    const newItem = {
        id: nextId,
        boq_description: "Manual Entry Row",
        gls_code: "GS-",
        product_description: "GLS-SPA Luminaire, IP20.",
        body_color: "Black",
        boq_qty: 1,
        unit: "Nos",
        rate: 1200,
        driver_details: "Fulham - 10W",
        driver_qty: 1,
        led_details: "Bridgelux - 4000K",
        accessories: "Standard",
        page: 0,
        matched_by: "manual_entry"
    };
    workspaceItems.push(newItem);
    renderTable();
    updateStats();
}

function generateInvoice() {
    if (workspaceItems.length === 0) {
        alert('Workspace is empty. Please upload a BOQ sheet.');
        return;
    }
    
    const payload = {
        buyer_name: document.getElementById('buyer-name').value.trim(),
        buyer_address: document.getElementById('buyer-address').value.trim(),
        buyer_gstin: document.getElementById('buyer-gstin').value.trim(),
        buyer_contact: document.getElementById('buyer-contact').value.trim(),
        invoice_no: document.getElementById('invoice-no').value.trim(),
        invoice_date: document.getElementById('invoice-date').value.trim(),
        payment_terms: document.getElementById('payment-terms').value.trim(),
        validity: document.getElementById('validity-period').value.trim(),
        destination: document.getElementById('destination').value.trim(),
        items: workspaceItems
    };
    
    // Change button to spinner
    const exportBtn = document.querySelector('.workspace-actions .btn-success');
    const originalText = exportBtn.innerHTML;
    exportBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Generating Invoice...`;
    exportBtn.disabled = true;
    
    fetch('/api/generate-invoice', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
    })
    .then(r => r.json())
    .then(data => {
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        
        if (data.error) {
            alert(data.error);
            return;
        }
        
        // Trigger download
        window.location.href = data.download_url;
    })
    .catch(err => {
        console.error(err);
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        alert('Failed to generate Proforma Invoice.');
    });
}
