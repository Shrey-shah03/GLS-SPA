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

// Click event listener to open file explorer for BOQ to IWO
dropZone.addEventListener('click', (e) => {
    if (e.target !== fileInput) {
        fileInput.click();
    }
});

// Drag and drop event listeners for BOQ to IWO
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
        handleFileUpload(files[0], 'boq');
    }
});

fileInput.addEventListener('change', (e) => {
    if (fileInput.files.length > 0) {
        handleFileUpload(fileInput.files[0], 'boq');
    }
});

// Click and drag-drop event listeners for Invoice
const invoiceDropZone = document.getElementById('invoice-drop-zone');
const invoiceFileInput = document.getElementById('invoice-file-input');

if (invoiceDropZone && invoiceFileInput) {
    invoiceDropZone.addEventListener('click', (e) => {
        if (e.target !== invoiceFileInput) {
            invoiceFileInput.click();
        }
    });
    
    invoiceFileInput.addEventListener('change', (e) => {
        if (invoiceFileInput.files.length > 0) {
            handleFileUpload(invoiceFileInput.files[0], 'invoice');
        }
    });
    
    invoiceDropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        invoiceDropZone.classList.add('dragover');
    });
    invoiceDropZone.addEventListener('dragleave', () => {
        invoiceDropZone.classList.remove('dragover');
    });
    invoiceDropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        invoiceDropZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileUpload(files[0], 'invoice');
        }
    });
}

function handleFileUpload(file, view = 'boq') {
    if (!file.name.endsWith('.xlsx')) {
        alert('Please upload a valid Excel spreadsheet (.xlsx)');
        return;
    }
    
    const activeDropZone = view === 'boq' ? dropZone : invoiceDropZone;
    
    // Show uploading visual state
    activeDropZone.innerHTML = `
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
            resetDropZone(view);
            return;
        }
        
        workspaceItems = data.items;
        uploadFileInfo = {
            filename: data.filename,
            projectName: data.project_name
        };
        
        const randomNum = Math.floor(100 + Math.random() * 900);
        
        // Sync/populate metadata on both layouts
        clientNameInput.value = data.project_name || 'Food Square';
        iwoNoInput.value = `P${randomNum}`;
        
        // Invoice Details Form Fields in View 2
        const invNoField = document.getElementById('inv-invoice-no');
        const invDateField = document.getElementById('inv-invoice-date');
        const invDestField = document.getElementById('inv-destination');
        if (invNoField) invNoField.value = `SOR/26-27/00${randomNum}`;
        if (invDateField) invDateField.value = `${dd}/${mm}/${yy}`;
        if (invDestField) invDestField.value = data.project_name || '';
        
        // Update UI panels based on view
        metaCard.classList.remove('disabled');
        statsArea.classList.remove('hidden');
        workspaceArea.classList.remove('hidden');
        
        const buyerCard = document.getElementById('buyer-profile-card');
        const invMetaCard = document.getElementById('invoice-meta-card');
        const invItemsArea = document.getElementById('invoice-items-area');
        if (buyerCard) buyerCard.classList.remove('disabled');
        if (invMetaCard) invMetaCard.classList.remove('disabled');
        if (invItemsArea) invItemsArea.classList.remove('hidden');
        
        // Render stats & grids for both layouts
        updateStats();
        renderTable();
        renderInvoiceTable();
        updateInvoiceTotals();
        
        // Restore drop zone to mini state
        activeDropZone.innerHTML = `
            <div class="upload-icon" style="font-size: 2rem; margin-bottom: 0.5rem;"><i class="fa-solid fa-circle-check" style="color: var(--color-success);"></i></div>
            <h3>File Uploaded Successfully</h3>
            <p>${file.name} (${workspaceItems.length} rows parsed)</p>
            <div class="supported-formats" onclick="triggerFileInputClick('${view}')">Upload Different File</div>
        `;
    })
    .catch(error => {
        console.error('Error:', error);
        alert('An error occurred: ' + (error.message || error));
        resetDropZone(view);
    });
}

function triggerFileInputClick(view) {
    if (view === 'boq') {
        fileInput.click();
    } else {
        invoiceFileInput.click();
    }
}

function resetDropZone(view = 'boq') {
    if (view === 'boq') {
        dropZone.innerHTML = `
            <input type="file" id="file-input" accept=".xlsx" style="display: none;">
            <div class="upload-icon"><i class="fa-solid fa-file-excel"></i></div>
            <h3>Drag & Drop BOQ Excel here</h3>
            <p>or <span class="highlight-link">browse your files</span></p>
            <div class="supported-formats">Supports .xlsx spreadsheets</div>
        `;
        const newInput = document.getElementById('file-input');
        newInput.addEventListener('change', (e) => {
            if (newInput.files.length > 0) {
                handleFileUpload(newInput.files[0], 'boq');
            }
        });
    } else {
        invoiceDropZone.innerHTML = `
            <input type="file" id="invoice-file-input" accept=".xlsx" style="display: none;">
            <div class="upload-icon"><i class="fa-solid fa-file-excel"></i></div>
            <h3>Drag & Drop BOQ Excel here</h3>
            <p>or <span class="highlight-link">browse your files</span></p>
            <div class="supported-formats">Supports .xlsx spreadsheets</div>
        `;
        const newInput = document.getElementById('invoice-file-input');
        newInput.addEventListener('change', (e) => {
            if (newInput.files.length > 0) {
                handleFileUpload(newInput.files[0], 'invoice');
            }
        });
    }
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

// Invoice view navigation & management
let savedBuyersList = [];

function switchView(view) {
    const navBoq = document.getElementById('nav-boq');
    const navInvoice = document.getElementById('nav-invoice');
    const viewBoq = document.getElementById('boq-generator-view');
    const viewInvoice = document.getElementById('invoice-generator-view');
    const titleHeader = document.getElementById('view-title-header');
    const descHeader = document.getElementById('view-desc-header');
    
    if (view === 'boq') {
        navBoq.classList.add('active');
        navInvoice.classList.remove('active');
        viewBoq.classList.remove('hidden');
        viewInvoice.classList.add('hidden');
        titleHeader.innerText = 'COMMAND CENTER';
        descHeader.innerText = 'Architecture & Engineering Work Order Dashboard';
    } else {
        navBoq.classList.remove('active');
        navInvoice.classList.add('active');
        viewBoq.classList.add('hidden');
        viewInvoice.classList.remove('hidden');
        titleHeader.innerText = 'PROFORMA INVOICE GENERATOR';
        descHeader.innerText = 'Live Billing Editor & GST Quotation Workspace';
        
        loadSavedBuyersDropdown();
        renderInvoiceTable();
        updateInvoiceTotals();
    }
}

function loadSavedBuyersDropdown() {
    fetch('/api/buyers')
    .then(r => r.json())
    .then(data => {
        savedBuyersList = data;
        const select = document.getElementById('saved-buyer-select');
        select.innerHTML = '<option value="">-- Select Saved Buyer --</option>';
        data.forEach(buyer => {
            select.innerHTML += `<option value="${buyer.name}">${buyer.name}</option>`;
        });
    })
    .catch(err => console.error('Failed to load buyers:', err));
}

function loadSavedBuyer(buyerName) {
    if (!buyerName) {
        document.getElementById('inv-buyer-name').value = '';
        document.getElementById('inv-buyer-address').value = '';
        document.getElementById('inv-buyer-gstin').value = '';
        document.getElementById('inv-buyer-contact').value = '';
        return;
    }
    const buyer = savedBuyersList.find(b => b.name === buyerName);
    if (buyer) {
        document.getElementById('inv-buyer-name').value = buyer.name;
        document.getElementById('inv-buyer-address').value = buyer.address;
        document.getElementById('inv-buyer-gstin').value = buyer.gstin;
        document.getElementById('inv-buyer-contact').value = buyer.contact;
    }
}

function saveBuyerProfile() {
    const payload = {
        name: document.getElementById('inv-buyer-name').value.trim(),
        address: document.getElementById('inv-buyer-address').value.trim(),
        gstin: document.getElementById('inv-buyer-gstin').value.trim(),
        contact: document.getElementById('inv-buyer-contact').value.trim()
    };
    
    if (!payload.name) {
        alert('Please fill in the Buyer Name to save.');
        return;
    }
    
    fetch('/api/buyers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    })
    .then(r => r.json())
    .then(data => {
        alert('Buyer profile saved successfully!');
        loadSavedBuyersDropdown();
    })
    .catch(err => {
        console.error(err);
        alert('Failed to save buyer profile.');
    });
}

function renderInvoiceTable() {
    const invTableBody = document.getElementById('invoice-table-body');
    invTableBody.innerHTML = '';
    
    if (workspaceItems.length === 0) {
        invTableBody.innerHTML = `<tr><td colspan="11" style="text-align: center; color: var(--color-text-muted); padding: 24px;">No items in workspace. Upload a BOQ sheet first.</td></tr>`;
        return;
    }
    
    workspaceItems.forEach((item, idx) => {
        const rate = parseFloat(item.rate) || 1200;
        const qty = parseInt(item.boq_qty) || 0;
        const hsn = item.hsn_code || '9405';
        
        const lineAmount = qty * rate;
        const lineGst = lineAmount * 0.18;
        const lineGrand = lineAmount + lineGst;
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td>
                <input type="text" class="table-input" style="width: 130px; font-family: monospace; font-weight: 600;" value="${item.gls_code || ''}" onchange="updateInvoiceItemField(${item.id}, 'gls_code', this.value)">
            </td>
            <td>
                <textarea class="table-input" style="width: 100%; height: 36px; font-size: 0.76rem; resize: none; text-align: left;" onchange="updateInvoiceItemField(${item.id}, 'product_description', this.value)">${item.product_description || ''}</textarea>
            </td>
            <td>
                <input type="text" class="table-input" style="width: 70px;" value="${hsn}" onchange="updateInvoiceItemField(${item.id}, 'hsn_code', this.value)">
            </td>
            <td>
                <input type="number" class="table-input" style="width: 60px;" value="${qty}" onchange="updateInvoiceItemField(${item.id}, 'boq_qty', this.value)">
            </td>
            <td>
                <select class="table-select" style="width: 70px;" onchange="updateInvoiceItemField(${item.id}, 'unit', this.value)">
                    <option value="Nos" ${item.unit === 'Nos' ? 'selected' : ''}>Nos</option>
                    <option value="Mtr" ${item.unit === 'Mtr' ? 'selected' : ''}>Mtr</option>
                    <option value="Set" ${item.unit === 'Set' ? 'selected' : ''}>Set</option>
                </select>
            </td>
            <td>
                <input type="number" class="table-input" style="width: 90px;" value="${rate}" onchange="updateInvoiceItemField(${item.id}, 'rate', this.value)">
            </td>
            <td style="font-weight: 500;">₹${lineAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            <td style="color: var(--color-text-muted);">₹${lineGst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            <td style="font-weight: 600; color: var(--color-primary);">₹${lineGrand.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
            <td>
                <button class="btn btn-secondary btn-action" onclick="deleteInvoiceItem(${item.id})" title="Delete Item">
                    <i class="fa-solid fa-trash-can" style="color: #ff4a4a;"></i>
                </button>
            </td>
        `;
        invTableBody.appendChild(tr);
    });
}

function updateInvoiceItemField(itemId, field, value) {
    updateItemField(itemId, field, value);
    renderInvoiceTable();
    updateInvoiceTotals();
}

function deleteInvoiceItem(itemId) {
    if (confirm('Remove this item from invoice?')) {
        workspaceItems = workspaceItems.filter(i => i.id !== itemId);
        renderInvoiceTable();
        updateInvoiceTotals();
        renderTable();
        updateStats();
    }
}

function numberToWordsIndian(num) {
    if (num === 0) return 'Zero Rupees Only';
    const single = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const double = ['', 'Ten', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    
    function convertLessThanThousand(n) {
        if (n === 0) return '';
        let str = '';
        if (n >= 100) {
            str += single[Math.floor(n / 100)] + ' Hundred ';
            n %= 100;
        }
        if (n >= 20) {
            str += double[Math.floor(n / 10)] + ' ';
            n %= 10;
        }
        if (n > 0) {
            str += single[n] + ' ';
        }
        return str;
    }
    
    let res = '';
    let crore = Math.floor(num / 10000000);
    num %= 10000000;
    let lakh = Math.floor(num / 100000);
    num %= 100000;
    let thousand = Math.floor(num / 1000);
    num %= 1000;
    
    if (crore > 0) {
        res += convertLessThanThousand(crore) + 'Crore ';
    }
    if (lakh > 0) {
        res += convertLessThanThousand(lakh) + 'Lakh ';
    }
    if (thousand > 0) {
        res += convertLessThanThousand(thousand) + 'Thousand ';
    }
    if (num > 0) {
        res += convertLessThanThousand(num);
    }
    return res.trim() + ' Rupees Only';
}

function updateInvoiceTotals() {
    let totalAmount = 0;
    let totalGst = 0;
    let totalGrand = 0;
    
    workspaceItems.forEach(item => {
        const rate = parseFloat(item.rate) || 1200;
        const qty = parseInt(item.boq_qty) || 0;
        const amount = qty * rate;
        const gst = amount * 0.18;
        const grand = amount + gst;
        
        totalAmount += amount;
        totalGst += gst;
        totalGrand += grand;
    });
    
    document.getElementById('inv-total-amount').innerText = `₹${totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('inv-total-gst').innerText = `₹${totalGst.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('inv-total-grand').innerText = `₹${totalGrand.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    
    const wordsEl = document.getElementById('inv-chargeable-words');
    if (wordsEl) {
        wordsEl.innerText = numberToWordsIndian(Math.floor(totalGrand));
    }
}

function generateInvoicePDF() {
    if (workspaceItems.length === 0) {
        alert('Workspace is empty. Please upload a BOQ sheet.');
        return;
    }
    
    const payload = {
        buyer_name: document.getElementById('inv-buyer-name').value.trim(),
        buyer_address: document.getElementById('inv-buyer-address').value.trim(),
        buyer_gstin: document.getElementById('inv-buyer-gstin').value.trim(),
        buyer_contact: document.getElementById('inv-buyer-contact').value.trim(),
        invoice_no: document.getElementById('inv-invoice-no').value.trim(),
        invoice_date: document.getElementById('inv-invoice-date').value.trim(),
        payment_terms: document.getElementById('inv-payment-terms').value.trim(),
        validity: document.getElementById('inv-validity-period').value.trim(),
        destination: document.getElementById('inv-destination').value.trim(),
        items: workspaceItems
    };
    
    const exportBtn = document.querySelector('#invoice-items-area .btn-primary');
    const originalText = exportBtn.innerHTML;
    exportBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Generating PDF...`;
    exportBtn.disabled = true;
    
    fetch('/api/generate-invoice-pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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
        window.location.href = data.download_url;
    })
    .catch(err => {
        console.error(err);
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        alert('Failed to generate PDF Invoice.');
    });
}

function generateInvoiceExcelLegacy() {
    if (workspaceItems.length === 0) {
        alert('Workspace is empty. Please upload a BOQ sheet.');
        return;
    }
    
    const payload = {
        buyer_name: document.getElementById('inv-buyer-name').value.trim(),
        buyer_address: document.getElementById('inv-buyer-address').value.trim(),
        buyer_gstin: document.getElementById('inv-buyer-gstin').value.trim(),
        buyer_contact: document.getElementById('inv-buyer-contact').value.trim(),
        invoice_no: document.getElementById('inv-invoice-no').value.trim(),
        invoice_date: document.getElementById('inv-invoice-date').value.trim(),
        payment_terms: document.getElementById('inv-payment-terms').value.trim(),
        validity: document.getElementById('inv-validity-period').value.trim(),
        destination: document.getElementById('inv-destination').value.trim(),
        items: workspaceItems
    };
    
    const exportBtn = document.querySelector('#invoice-items-area .btn-success');
    const originalText = exportBtn.innerHTML;
    exportBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Generating Excel...`;
    exportBtn.disabled = true;
    
    fetch('/api/generate-invoice', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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
        window.location.href = data.download_url;
    })
    .catch(err => {
        console.error(err);
        exportBtn.innerHTML = originalText;
        exportBtn.disabled = false;
        alert('Failed to generate Excel Invoice.');
    });
}

// Prefill initial buyer fields on load
document.addEventListener('DOMContentLoaded', () => {
    loadSavedBuyersDropdown();
});

