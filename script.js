    // Global variables
    let formData = {};

    // ===== Currency Conversion Rates (Base = INR) =====
const currencyRates = {
    INR: 1,
    USD: 0.012,
    EUR: 0.011,
    GBP: 0.0095,
    BDT: 1.32,
    CNY: 0.086,
    JPY: 1.75,
    AUD: 0.018,
    CAD: 0.016
};

// Currency symbols for display
const currencySymbols = {
    INR: '₹',
    USD: '$',
    EUR: '€',
    GBP: '£',
    BDT: '৳',
    CNY: '¥',
    JPY: '¥',
    AUD: '$',
    CAD: '$'
};


    function formatRs(value) {
        const num = parseFloat(String(value).replace(/[^0-9.]/g, ''));
        return `Rs ${isNaN(num) ? '0.00' : num.toFixed(2)}/-`;
    }

    function safeNumber(value) {
        const num = parseFloat(String(value).replace(/[^0-9.]/g, ''));
        return isNaN(num) ? 0 : num;
    }

    function formatCurrency(inrValue) {
    const currency = document.getElementById('currency')?.value || 'INR';
    const rate = currencyRates[currency] || 1;

    const base = parseFloat(String(inrValue).replace(/[^0-9.]/g, '')) || 0;
    const converted = base * rate;

        const symbol = currencySymbols[currency] || currency;
        // For INR previously used Rs .../-, keep concise symbol-based formatting
        return `${symbol}${converted.toFixed(2)}`;
}

// Update all visible currency symbol spans to match selected currency
function updateCurrencySymbols() {
    const currency = document.getElementById('currency')?.value || 'INR';
    const symbol = currencySymbols[currency] || currency;
    document.querySelectorAll('.currency-symbol').forEach(el => {
        el.textContent = symbol;
    });
}

    // DOM Content Loaded
    document.addEventListener('DOMContentLoaded', function() {
        initializeForm();
        setupEventListeners();
        updateCurrencySymbols();
        setupCalculationListeners();
        setupTheme();
        setupScrollAnimations();
        // setupProgressBar(); // Disabled to avoid visual issues
        setupScrollIndicator();
        setupScrollToTop();
    });

    // Initialize form
    function initializeForm() {
        // No default values set - let users enter their own data
        // All form fields will start empty
    }

    // Setup event listeners
    function setupEventListeners() {
        console.log('Setting up event listeners...');
        
        // Export and action buttons
        const excelBtn = document.getElementById('exportExcelBtn');
        const clearBtn = document.getElementById('clearBtn');

            // Currency change → recalculate totals
    const currencySelect = document.getElementById('currency');
    if (currencySelect) {
        currencySelect.addEventListener('change', function() {
            updateCurrencySymbols();
            calculateAllTotals();
        });
    }

        
        if (excelBtn) {
            excelBtn.addEventListener('click', function(e) {
                e.preventDefault();
                console.log('Excel button clicked!');
                exportToExcel();
            });
            console.log('Excel button listener added');
        } else {
            console.error('Excel button not found!');
        }
        
        
        
        if (clearBtn) {
            clearBtn.addEventListener('click', clearForm);
            console.log('Clear button listener added');
        }
        
        // Form submission
        const form = document.getElementById('bomForm');
        if (form) {
            form.addEventListener('submit', function(e) {
            e.preventDefault();
            calculateAllTotals();
        });
        }
    }

    // Setup calculation listeners for real-time updates
    function setupCalculationListeners() {
        const calculationFields = [
            'fabricAmount', 'fabricPerGarment', 'fabricQty',
            'collarYarnAmount', 'collarYarnPerGarment', 'collarYarnQty',
            'tippingYarnAmount', 'tippingYarnPerGarment', 'tippingYarnQty',
            'collarPerGarment', 'collarQty',
            'cuffPerGarment', 'cuffQty',
            'buttonAmount', 'buttonPerGarment', 'buttonQty',
            'velvetTapeMeter', 'velvetTapePerGarment', 'velvetTapeQty',
            'washCareAmount', 'washCarePerGarment', 'washCareQty',
            'wovenSizeAmount', 'wovenSizePerGarment', 'wovenSizeQty',
            'polybagPerGarment', 'polybagQty'
        ];
        
        calculationFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (field) {
                field.addEventListener('input', calculateAllTotals);
                field.addEventListener('change', calculateAllTotals);
            }
        });
    }

    // Theme setup and utilities
    function setupTheme() {
        try {
            const savedTheme = localStorage.getItem('twinklub_theme');
            const osPrefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
            const initialTheme = savedTheme || (osPrefersDark ? 'dark' : 'light');
            applyTheme(initialTheme);

            // Toggle button support if present
            const toggleButton = document.getElementById('themeToggle');
            if (toggleButton) {
                toggleButton.addEventListener('click', function() {
                    const current = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
                    applyTheme(current);
                });
            }

            // React to OS theme changes dynamically
            if (window.matchMedia) {
                const mql = window.matchMedia('(prefers-color-scheme: dark)');
                if (typeof mql.addEventListener === 'function') {
                    mql.addEventListener('change', function(e) {
                        const userSet = localStorage.getItem('twinklub_theme');
                        if (!userSet) {
                            applyTheme(e.matches ? 'dark' : 'light');
                        }
                    });
                } else if (typeof mql.addListener === 'function') {
                    // Older browsers
                    mql.addListener(function(e) {
                        const userSet = localStorage.getItem('twinklub_theme');
                        if (!userSet) {
                            applyTheme(e.matches ? 'dark' : 'light');
                        }
                    });
                }
            }
        } catch (err) {
            console.error('Error setting up theme:', err);
        }

        document.addEventListener('input', function (e) {
    if (
        e.target.classList.contains('collarPerGarment') ||
        e.target.classList.contains('collarQty') ||
        e.target.classList.contains('cuffPerGarment') ||
        e.target.classList.contains('cuffQty')
    ) {
        calculateAllTotals();
    }
});

    }

    function applyTheme(theme) {
        if (theme !== 'light' && theme !== 'dark') return;
        document.documentElement.setAttribute('data-theme', theme);
        try {
            localStorage.setItem('twinklub_theme', theme);
        } catch (_) {
            // ignore storage errors
        }
    }

    // Calculate individual section totals
    function calculateSectionTotal(amountId, perGarmentId, qtyId, totalId) {
        const amountField = amountId ? document.getElementById(amountId) : null;
        const amount = amountField ? parseFloat(amountField.value || 0) : 0;
        const perGarment = parseFloat(document.getElementById(perGarmentId)?.value || 0);
        const qty = parseFloat(document.getElementById(qtyId)?.value || 0);
        
        const hasAmount = Boolean(amountField);
        const total = hasAmount ? (amount * perGarment * qty) : (perGarment * qty);
        
        const totalField = document.getElementById(totalId);
        if (totalField) {
            // store raw INR value for exports and calculations
            totalField.dataset.inr = String(Number((total || 0).toFixed(2)));
            // display converted value according to selected currency
            totalField.value = formatCurrency(total);
        }
        return total;
    }

    function calculateMultipleCollars() {
    let total = 0;

    document.querySelectorAll('.collar-block').forEach(block => {
        const perGarment = parseFloat(block.querySelector('.collarPerGarment')?.value || 0);
        const qty = parseFloat(block.querySelector('.collarQty')?.value || 0);

        const amount = perGarment * qty;

        const totalField = block.querySelector('.collarTotal');
        if (totalField) {
            totalField.dataset.inr = String(Number((amount || 0).toFixed(2)));
            totalField.value = formatCurrency(amount);
        }

        total += amount;
    });

    return total;
}

function calculateMultipleCuffs() {
    let total = 0;

    document.querySelectorAll('.cuff-block').forEach(block => {
        const perGarment = parseFloat(block.querySelector('.cuffPerGarment')?.value || 0);
        const qty = parseFloat(block.querySelector('.cuffQty')?.value || 0);

        const amount = perGarment * qty;

        const totalField = block.querySelector('.cuffTotal');
        if (totalField) {
            totalField.dataset.inr = String(Number((amount || 0).toFixed(2)));
            totalField.value = formatCurrency(amount);
        }

        total += amount;
    });

    return total;
}



    // Calculate all totals
    function calculateAllTotals() {
        let grandTotal = 0;
        
        // Fabric 1
        grandTotal += calculateSectionTotal('fabricAmount', 'fabricPerGarment', 'fabricQty', 'fabricTotal');
        
        // Fabric 2 removed (no second fabric section in input)
        
        // Collar Yarn
        grandTotal += calculateSectionTotal('collarYarnAmount', 'collarYarnPerGarment', 'collarYarnQty', 'collarYarnTotal');
        
        // Tipping Yarn
        grandTotal += calculateSectionTotal('tippingYarnAmount', 'tippingYarnPerGarment', 'tippingYarnQty', 'tippingYarnTotal');
        
        // Collar (only per garment * qty)
        // Multiple Collar Measurements
            grandTotal += calculateMultipleCollars();


        
        // Multiple Cuff Measurements
        grandTotal += calculateMultipleCuffs();
        
        // Buttons: compute as perGarment * qty only (ignore amount)
        grandTotal += calculateSectionTotal('', 'buttonPerGarment', 'buttonQty', 'buttonTotal');
        
        // Velvet Tapes
        grandTotal += calculateSectionTotal('velvetTapeMeter', 'velvetTapePerGarment', 'velvetTapeQty', 'velvetTapeTotal');
        
        // Wash Care Labels
        grandTotal += calculateSectionTotal('washCareAmount', 'washCarePerGarment', 'washCareQty', 'washCareTotal');
        
        // Woven Size Labels
        grandTotal += calculateSectionTotal('wovenSizeAmount', 'wovenSizePerGarment', 'wovenSizeQty', 'wovenSizeTotal');
        
        // Polybags (only per garment * qty)
        grandTotal += calculateSectionTotal('', 'polybagPerGarment', 'polybagQty', 'polybagTotal');
        
        // Update grand total display with explicit Unicode rupee to avoid encoding issues
        document.getElementById('grandTotal').textContent = formatCurrency(grandTotal);
    
        return grandTotal;
    }
    function addCollarBlock() {
    const container = document.getElementById('collarBlocks');
    if (!container) return;

    const firstBlock = container.querySelector('.collar-block');
    if (!firstBlock) return;

    const clone = firstBlock.cloneNode(true);

    clone.querySelectorAll('input, select').forEach(el => {
        if (el.tagName === 'SELECT') el.selectedIndex = 0;
        else el.value = '';
    });

    container.appendChild(clone);
}

function addCuffBlock() {
    const container = document.getElementById('cuffBlocks');
    if (!container) return;

    const firstBlock = container.querySelector('.cuff-block');
    if (!firstBlock) return;

    const clone = firstBlock.cloneNode(true);

    clone.querySelectorAll('input, select').forEach(el => {
        if (el.tagName === 'SELECT') el.selectedIndex = 0;
        else el.value = '';
    });

    container.appendChild(clone);
}


    // Collect form data
    function collectFormData() {
        const form = document.getElementById('bomForm');
        const formData = new FormData(form);
        const data = {};
        
        // ===== COLLECT MULTIPLE COLLARS =====
        data.collars = [];

    document.querySelectorAll('.collar-block').forEach(block => {
        const totalEl = block.querySelector('.collarTotal');
        const collar = {
            size: block.querySelector('.collarSize')?.value || '',
            dimensions: block.querySelector('.collarDimensions')?.value || '',
            unit: block.querySelector('.collarUnit')?.value || '',
            perGarment: block.querySelector('.collarPerGarment')?.value || '',
            supplier: block.querySelector('.collarSupplier')?.value || '',
            qty: block.querySelector('.collarQty')?.value || '',
            // raw INR value
            total: totalEl && totalEl.dataset && totalEl.dataset.inr ? parseFloat(totalEl.dataset.inr) : 0
        };

        const anyFilled = [collar.size, collar.dimensions, collar.unit, collar.perGarment, collar.supplier, collar.qty, collar.total].some(v => String(v).trim() !== '');
        if (anyFilled) data.collars.push(collar);
    });

        // ===== COLLECT MULTIPLE CUFFS =====
data.cuffs = [];

document.querySelectorAll('.cuff-block').forEach(block => {
    const totalEl = block.querySelector('.cuffTotal');
    const cuff = {
        size: block.querySelector('.cuffSize')?.value || '',
        dimensions: block.querySelector('.cuffDimensions')?.value || '',
        unit: block.querySelector('.cuffUnit')?.value || '',
        perGarment: block.querySelector('.cuffPerGarment')?.value || '',
        supplier: block.querySelector('.cuffSupplier')?.value || '',
        qty: block.querySelector('.cuffQty')?.value || '',
        total: totalEl && totalEl.dataset && totalEl.dataset.inr ? parseFloat(totalEl.dataset.inr) : 0
    };

    const anyFilled = [cuff.size, cuff.dimensions, cuff.unit, cuff.perGarment, cuff.supplier, cuff.qty, cuff.total].some(v => String(v).trim() !== '');
    if (anyFilled) data.cuffs.push(cuff);
});

// If there are no dynamic cuff-blocks, fall back to single cuff fields (existing HTML)
if (data.cuffs.length === 0) {
    const cuffTotalEl = document.getElementById('cuffTotal');
    const singleCuff = {
        size: getVal('cuffSize'),
        dimensions: getVal('cuffDimensions'),
        unit: getVal('cuffUnit'),
        perGarment: getVal('cuffPerGarment'),
        supplier: getVal('cuffSupplier'),
        qty: getVal('cuffQty'),
        total: cuffTotalEl && cuffTotalEl.dataset && cuffTotalEl.dataset.inr ? parseFloat(cuffTotalEl.dataset.inr) : (parseFloat(String(getVal('cuffTotal')||'').replace(/[^0-9.]/g,'')) || 0)
    };
    const anyFilled = [singleCuff.size, singleCuff.dimensions, singleCuff.unit, singleCuff.perGarment, singleCuff.supplier, singleCuff.qty, singleCuff.total].some(v => String(v).trim() !== '');
    if (anyFilled) data.cuffs.push(singleCuff);
}

        
        // Safe getter for optional inputs
        const getVal = (id) => {
            const el = document.getElementById(id);
            return el ? el.value : '';
        };

        // Order Information
        data.orderDate = getVal('orderDate');
        data.expectedDeliveryDate = getVal('expectedDeliveryDate');
        data.orderNumber = getVal('orderNumber');
        data.orderType = getVal('orderType');
        data.currency = getVal('currency') || 'INR';

        // Fabric Section
        data.fabricType = getVal('fabricType');
        data.fabricAmount = getVal('fabricAmount');
        data.fabricUnit = getVal('fabricUnit');
        data.fabricPerGarment = getVal('fabricPerGarment');
        data.fabricSupplier = getVal('fabricSupplier');
        data.fabricQty = getVal('fabricQty');

        // Collar Yarn Section
        data.collarYarnType = getVal('collarYarnType');
        data.collarYarnAmount = getVal('collarYarnAmount');
        data.collarYarnUnit = getVal('collarYarnUnit');
        data.collarYarnPerGarment = getVal('collarYarnPerGarment');
        data.collarYarnSupplier = getVal('collarYarnSupplier');
        data.collarYarnQty = getVal('collarYarnQty');

        // Tipping Yarn Section
        data.tippingYarnType = getVal('tippingYarnType');
        data.tippingYarnAmount = getVal('tippingYarnAmount');
        data.tippingYarnUnit = getVal('tippingYarnUnit');
        data.tippingYarnPerGarment = getVal('tippingYarnPerGarment');
        data.tippingYarnSupplier = getVal('tippingYarnSupplier');
        data.tippingYarnQty = getVal('tippingYarnQty');

        // Buttons Section
        data.buttonMaterial = getVal('buttonMaterial');
        data.buttonAmount = getVal('buttonAmount');
        data.buttonUnit = getVal('buttonUnit');
        data.buttonPerGarment = getVal('buttonPerGarment');
        data.buttonSupplier = getVal('buttonSupplier');
        data.buttonQty = getVal('buttonQty');

        // Velvet Tapes Section
        data.velvetTapeQuality = getVal('velvetTapeQuality');
        data.velvetTapeMeter = getVal('velvetTapeMeter');
        data.velvetTapeUnit = getVal('velvetTapeUnit');
        data.velvetTapePerGarment = getVal('velvetTapePerGarment');
        data.velvetTapeSupplier = getVal('velvetTapeSupplier');
        data.velvetTapeQty = getVal('velvetTapeQty');

        // Wash Care Labels Section
        data.washCareQuality = getVal('washCareQuality');
        data.washCareAmount = getVal('washCareAmount');
        data.washCareUnit = getVal('washCareUnit');
        data.washCarePerGarment = getVal('washCarePerGarment');
        data.washCareSupplier = getVal('washCareSupplier');
        data.washCareQty = getVal('washCareQty');

        // Woven Size Labels Section
        data.wovenSizeLabel = getVal('wovenSizeLabel');
        data.wovenSizeAmount = getVal('wovenSizeAmount');
        data.wovenSizeUnit = getVal('wovenSizeUnit');
        data.wovenSizePerGarment = getVal('wovenSizePerGarment');
        data.wovenSizeSupplier = getVal('wovenSizeSupplier');
        data.wovenSizeQty = getVal('wovenSizeQty');

        // Polybags Section
        data.polybagQuality = getVal('polybagQuality');
        data.polybagSize = getVal('polybagSize');
        data.polybagUnit = getVal('polybagUnit');
        data.polybagPerGarment = getVal('polybagPerGarment');
        data.polybagSupplier = getVal('polybagSupplier');
        data.polybagQty = getVal('polybagQty');
        
        // Add calculated totals (raw INR values read from data-inr when available)
        const getNumericTotal = (id) => {
            const el = document.getElementById(id);
            if (!el) return 0;
            if (el.dataset && el.dataset.inr) return parseFloat(el.dataset.inr) || 0;
            const raw = parseFloat(String(el.value || el.textContent || '').replace(/[^0-9.]/g,''));
            return isNaN(raw) ? 0 : raw;
        };

        data.fabricTotal = getNumericTotal('fabricTotal');
        data.collarYarnTotal = getNumericTotal('collarYarnTotal');
        data.tippingYarnTotal = getNumericTotal('tippingYarnTotal');
        data.collarTotal = getNumericTotal('collarTotal');
        data.cuffTotal = getNumericTotal('cuffTotal');
        data.buttonTotal = getNumericTotal('buttonTotal');
        data.velvetTapeTotal = getNumericTotal('velvetTapeTotal');
        data.washCareTotal = getNumericTotal('washCareTotal');
        data.wovenSizeTotal = getNumericTotal('wovenSizeTotal');
        data.polybagTotal = getNumericTotal('polybagTotal');
        data.grandTotal = calculateAllTotals();
        
        return data;
    }


    // Show export card with specific file type
    function showExportCard(fileType, message, details) {
        const card = document.createElement('div');
        card.className = 'export-card';
        card.id = 'exportCard';
        
        const icons = {
            excel: 'fas fa-file-excel',
            pdf: 'fas fa-file-pdf', 
            word: 'fas fa-file-alt'
        };
        
        const titles = {
            excel: 'Excel Export',
            pdf: 'PDF Export',
            word: 'Word Export'
        };
        
        card.innerHTML = `
            <div class="export-card-header">
                <div class="export-card-icon ${fileType}">
                    <i class="${icons[fileType]}"></i>
                </div>
                <div>
                    <h3 class="export-card-title">${titles[fileType]}</h3>
                    <p class="export-card-subtitle">${fileType.toUpperCase()} Document</p>
                </div>
            </div>
            <div class="export-card-spinner ${fileType}"></div>
            <div class="export-card-message">${message}</div>
            <div class="export-card-details">${details}</div>
        `;
        
        document.body.appendChild(card);
        
        // Trigger animation
        setTimeout(() => {
            card.classList.add('show');
        }, 100);
    }

    // Hide export card
    function hideExportCard() {
        const card = document.getElementById('exportCard');
        if (card) {
            card.classList.remove('show');
            setTimeout(() => {
                if (card.parentNode) {
                    card.remove();
                }
            }, 300);
        }
    }

    // Show loading overlay (fallback)
    function showLoadingOverlay(message = 'Processing...', subtext = 'Please wait while we prepare your export') {
        const overlay = document.createElement('div');
        overlay.className = 'loading-overlay';
        overlay.id = 'loadingOverlay';
        
        overlay.innerHTML = `
            <div class="loading-content">
                <div class="loading-spinner"></div>
                <div class="loading-text">${message}</div>
                <div class="loading-subtext">${subtext}</div>
            </div>
        `;
        
        document.body.appendChild(overlay);
        
        // Trigger animation
        setTimeout(() => {
            overlay.classList.add('show');
        }, 100);
    }

    // Hide loading overlay
    function hideLoadingOverlay() {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.classList.remove('show');
            setTimeout(() => {
                if (overlay.parentNode) {
                    overlay.remove();
                }
            }, 300);
        }
    }

    // Create CSV content
    function createCSVContent(data) {
                const rows = [];
                rows.push(['Twinklub BOM - Bill of Materials']);
                rows.push(['']);
        rows.push(['Order Number:', data.orderNumber || '']);
        rows.push(['Order Type:', data.orderType || '']);
                rows.push(['']);
        rows.push(['Section', 'Details', 'Amount', 'Unit', 'Per Garment', 'Supplier', 'Q.QTY', 'Total']);
        rows.push(['Fabric 1', data.fabricType || '', data.fabricAmount || '', data.fabricUnit || '', data.fabricPerGarment || '', data.fabricSupplier || '', data.fabricQty || '', formatCurrency(data.fabricTotal || 0)]);
        rows.push(['Collar Yarn', data.collarYarnType || '', data.collarYarnAmount || '', data.collarYarnUnit || '', data.collarYarnPerGarment || '', data.collarYarnSupplier || '', data.collarYarnQty || '', formatCurrency(data.collarYarnTotal || 0)]);
        rows.push(['Tipping Yarn', data.tippingYarnType || '', data.tippingYarnAmount || '', data.tippingYarnUnit || '', data.tippingYarnPerGarment || '', data.tippingYarnSupplier || '', data.tippingYarnQty || '', formatCurrency(data.tippingYarnTotal || 0)]);

        // Add collar measurements (multiple) with numbering
        if (Array.isArray(data.collars) && data.collars.length > 0) {
            data.collars.forEach((c, idx) => {
                const details = `${c.size || ''}${c.dimensions ? ' - ' + c.dimensions : ''}`;
                rows.push([`Collar Measurements-${idx + 1}`, details, '', c.unit || '', c.perGarment || '', c.supplier || '', c.qty || '', formatCurrency(c.total || 0)]);
            });
        }

        // Add cuff measurements (multiple) with numbering
        if (Array.isArray(data.cuffs) && data.cuffs.length > 0) {
            data.cuffs.forEach((c, idx) => {
                const details = `${c.size || ''}${c.dimensions ? ' - ' + c.dimensions : ''}`;
                rows.push([`Cuff Measurements-${idx + 1}`, details, '', c.unit || '', c.perGarment || '', c.supplier || '', c.qty || '', formatCurrency(c.total || 0)]);
            });
        }
        
        rows.push(['Buttons', data.buttonMaterial || '', data.buttonAmount || '', data.buttonUnit || '', data.buttonPerGarment || '', data.buttonSupplier || '', data.buttonQty || '', formatCurrency(data.buttonTotal || 0)]);
        rows.push(['Velvet Tapes', data.velvetTapeQuality || '', data.velvetTapeMeter || '', data.velvetTapeUnit || '', data.velvetTapePerGarment || '', data.velvetTapeSupplier || '', data.velvetTapeQty || '', formatCurrency(data.velvetTapeTotal || 0)]);
        rows.push(['Wash Care Labels', data.washCareQuality || '', data.washCareAmount || '', data.washCareUnit || '', data.washCarePerGarment || '', data.washCareSupplier || '', data.washCareQty || '', formatCurrency(data.washCareTotal || 0)]);
        rows.push(['Polybags', data.polybagQuality || '', data.polybagSize || '', data.polybagUnit || '', data.polybagPerGarment || '', data.polybagSupplier || '', data.polybagQty || '', formatCurrency(data.polybagTotal || 0)]);
        rows.push(['']);
        // Ensure grand total uses current currency symbol
        rows.push(['GRAND TOTAL:', formatCurrency(parseFloat(String(data.grandTotal).replace(/[^0-9.]/g, '')) || calculateAllTotals())]);
        
        return rows.map(r => r.map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',')).join('\n');
    }

    // Download file helper
    function downloadFile(content, type, filename) {
        try {
            const blob = new Blob([content], { type: type === 'csv' ? 'text/csv;charset=utf-8;' : 'text/plain' });
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = filename;
            document.body.appendChild(link);
                link.click();
            document.body.removeChild(link);
                URL.revokeObjectURL(link.href);
            console.log('File downloaded:', filename);
        } catch (error) {
            console.error('Error downloading file:', error);
        }
    }

    // Create text content
    function createTextContent(data) {
        return `
    TWINKLUB BOM - BILL OF MATERIALS
    ================================

    Order Information:
    ------------------
    Order Number: ${data.orderNumber || ''}
    Order Type: ${data.orderType || ''}

    BOM Data:
    ---------
    Fabric 1: ${data.fabricType || ''}
    Amount: ${data.fabricAmount || ''} ${data.fabricUnit || ''}
    Per Garment: ${data.fabricPerGarment || ''}
    Supplier: ${data.fabricSupplier || ''}
    Q.QTY: ${data.fabricQty || ''}
    Total: ${formatCurrency(data.fabricTotal || 0)}

    Collar Yarn: ${data.collarYarnType || ''}
    Amount: ${data.collarYarnAmount || ''} ${data.collarYarnUnit || ''}
    Per Garment: ${data.collarYarnPerGarment || ''}
    Supplier: ${data.collarYarnSupplier || ''}
    Q.QTY: ${data.collarYarnQty || ''}
    Total: ${formatCurrency(data.collarYarnTotal || 0)}

    Tipping Yarn: ${data.tippingYarnType || ''}
    Amount: ${data.tippingYarnAmount || ''} ${data.tippingYarnUnit || ''}
    Per Garment: ${data.tippingYarnPerGarment || ''}
    Supplier: ${data.tippingYarnSupplier || ''}
    Q.QTY: ${data.tippingYarnQty || ''}
    Total: ${formatCurrency(data.tippingYarnTotal || 0)}

    Buttons: ${data.buttonMaterial || ''}
    Amount: ${data.buttonAmount || ''} ${data.buttonUnit || ''}
    Per Garment: ${data.buttonPerGarment || ''}
    Supplier: ${data.buttonSupplier || ''}
    Q.QTY: ${data.buttonQty || ''}
    Total: ${formatCurrency(data.buttonTotal || 0)}

    Velvet Tapes: ${data.velvetTapeQuality || ''}
    Amount: ${data.velvetTapeMeter || ''} ${data.velvetTapeUnit || ''}
    Per Garment: ${data.velvetTapePerGarment || ''}
    Supplier: ${data.velvetTapeSupplier || ''}
    Q.QTY: ${data.velvetTapeQty || ''}
    Total: ${formatCurrency(data.velvetTapeTotal || 0)}

    Wash Care Labels: ${data.washCareQuality || ''}
    Amount: ${data.washCareAmount || ''} ${data.washCareUnit || ''}
    Per Garment: ${data.washCarePerGarment || ''}
    Supplier: ${data.washCareSupplier || ''}
    Q.QTY: ${data.washCareQty || ''}
    Total: ${formatCurrency(data.washCareTotal || 0)}

    Polybags: ${data.polybagQuality || ''}
    Size: ${data.polybagSize || ''} ${data.polybagUnit || ''}
    Per Garment: ${data.polybagPerGarment || ''}
    Supplier: ${data.polybagSupplier || ''}
    Q.QTY: ${data.polybagQty || ''}
    Total: ${formatCurrency(data.polybagTotal || 0)}

    ================================
    COLLAR MEASUREMENTS:
    ---------------------
    ${Array.isArray(data.collars) && data.collars.length > 0 ? data.collars.map((c, i) => `Collar MSMTS-${i + 1}: Size: ${c.size || ''} | Dim: ${c.dimensions || ''} | Unit: ${c.unit || ''} | Per Garment: ${c.perGarment || ''} | Supplier: ${c.supplier || ''} | Q.QTY: ${c.qty || ''} | Total: ${formatCurrency(c.total || 0)}`).join('\n    ') : 'None'}

    CUFF MEASUREMENTS:
    -------------------
    ${Array.isArray(data.cuffs) && data.cuffs.length > 0 ? data.cuffs.map((c, i) => `Cuff MSMTS-${i + 1}: Size: ${c.size || ''} | Dim: ${c.dimensions || ''} | Unit: ${c.unit || ''} | Per Garment: ${c.perGarment || ''} | Supplier: ${c.supplier || ''} | Q.QTY: ${c.qty || ''} | Total: ${formatCurrency(c.total || 0)}`).join('\n    ') : 'None'}

    ================================
    GRAND TOTAL: ${formatCurrency(parseFloat(String(data.grandTotal).replace(/[^0-9.]/g, '')) || calculateAllTotals())}
    ================================
        `;
    }

    // Export to Excel
    function exportToExcel() {
        console.log('Excel export function called!');
        try {
            // Show loading overlay
            showLoadingOverlay('Exporting to Excel...', 'Generating spreadsheet with your BOM data');
            
            // Ensure totals are up to date
            calculateAllTotals();

            const data = collectFormData();
            console.log('Form data collected:', data);
            
            // Check if XLSX is available
            if (typeof XLSX === 'undefined' || !XLSX.utils || !XLSX.writeFile) {
                console.warn('XLSX not available, using CSV fallback');
                // Create simple CSV export
                const csvContent = createCSVContent(data);
                downloadFile(csvContent, 'csv', 'Twinklub_BOM_Export.csv');
                hideLoadingOverlay();
                showAnimatedPopup('CSV file exported successfully!', 'success');
                return;
            }

            const workbook = XLSX.utils.book_new();
            
            // Create worksheet data. Only include selected sizes when corresponding QTY > 0.
            const worksheetData = [
                ['Twinklub BOM - Bill of Materials'],
                [''],
                ['Garment Type:', data.orderType || ''],
                ['Order Number:', data.orderNumber || ''],
                ['Order Date:', data.orderDate || ''],
                ['Expected Delivery Date:', data.expectedDeliveryDate || ''],
                ['Currency:', data.currency || 'INR'],

                [''],
                ['QUALITY', 'DETAILS', 'UNIT OF MSMTS', 'PER GARMENT', 'SUPPLIER', 'O.QTY', 'TOTAL AMOUNT'],
                [''],
                // Each row below must have exactly 7 columns to match the header
                ['FABRIC', data.fabricType || '', data.fabricUnit || '', data.fabricPerGarment || '', data.fabricSupplier || '', data.fabricQty || '', formatCurrency(data.fabricTotal || 0)],
                [''],
                ['COLLAR YARN', data.collarYarnType || '', data.collarYarnUnit || '', data.collarYarnPerGarment || '', data.collarYarnSupplier || '', data.collarYarnQty || '', formatCurrency(data.collarYarnTotal || 0)],
                [''],
                ['TIPPING YARN', data.tippingYarnType || '', data.tippingYarnUnit || '', data.tippingYarnPerGarment || '', data.tippingYarnSupplier || '', data.tippingYarnQty || '', formatCurrency(data.tippingYarnTotal || 0)],
                [''],
                // Collar measurements - only show if quantity > 0
                // ===== MULTIPLE COLLAR ROWS =====
...(data.collars || []).flatMap((c, index) => [
    [
        `COLLAR MSMTS-${index + 1}`,
        `${c.size} - ${c.dimensions}`,
        c.unit,
        c.perGarment,
        c.supplier,
        c.qty,
        formatCurrency(c.total)
    ],
    ['']
]),

                // ===== MULTIPLE CUFF ROWS =====
...(data.cuffs || []).flatMap((c, index) => [
    [
        `CUFF MSMTS-${index + 1}`,
        `${c.size} - ${c.dimensions}`,
        c.unit,
        c.perGarment,
        c.supplier,
        c.qty,
        formatCurrency(c.total)
    ],
    ['']
]),
                ['BUTTONS', data.buttonMaterial || '', data.buttonUnit || '', data.buttonPerGarment || '', data.buttonSupplier || '', data.buttonQty || '', formatCurrency(data.buttonTotal || 0)],
                [''],
                ['VELVET TAPES', data.velvetTapeQuality || '', data.velvetTapeUnit || '', data.velvetTapePerGarment || '', data.velvetTapeSupplier || '', data.velvetTapeQty || '', formatCurrency(data.velvetTapeTotal || 0)],
                [''],
                ['WASH CARE LABELS', data.washCareQuality || '', data.washCareUnit || '', data.washCarePerGarment || '', data.washCareSupplier || '', data.washCareQty || '', formatCurrency(data.washCareTotal || 0)],
                [''],
                // Woven size labels - only show if quantity > 0
                ...((() => {
                    const q = parseFloat(data.wovenSizeQty || '');
                    if (!isNaN(q) && q > 0) {
                        return [['SIZE LABELS (WVN SIZE LBL)', `WVN SIZE LBL - ${data.wovenSizeLabel || ''}`, data.wovenSizeUnit || '', data.wovenSizePerGarment || '', data.wovenSizeSupplier || '', data.wovenSizeQty || '', formatCurrency(data.wovenSizeTotal || 0)], ['']];
                    }
                    return [];
                })()),
                ['POLYBAGS', (data.polybagQuality || '') + (data.polybagSize ? ` - ${data.polybagSize}` : ''), data.polybagUnit || '', data.polybagPerGarment || '', data.polybagSupplier || '', data.polybagQty || '', formatCurrency(data.polybagTotal || 0)],
                [''],
               ['GRAND TOTAL:', formatCurrency(calculateAllTotals())]

            ];
            
            const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
            
            // Set column widths
            worksheet['!cols'] = [
                { width: 20 },
                { width: 30 }
            ];
            
            XLSX.utils.book_append_sheet(workbook, worksheet, 'BOM Data');
            
            // Generate filename
            const orderNumber = data.orderNumber || 'BOM';
            const filename = `Twinklub_BOM_${orderNumber}_${new Date().toISOString().split('T')[0]}.xlsx`;
            
            // Simulate processing time for better UX
            setTimeout(() => {
            XLSX.writeFile(workbook, filename);
                hideLoadingOverlay();
                removeButtonLoadingState('exportExcelBtn');
            showAnimatedPopup('Excel file exported successfully!', 'success');
            }, 1500);
            
        } catch (error) {
            console.error('Error exporting to Excel:', error);
            hideLoadingOverlay();
            removeButtonLoadingState('exportExcelBtn');
            showAnimatedPopup('Error exporting to Excel. Please try again.', 'error');
        }
    }
    // Clear form
    function clearForm() {
        if (confirm('Are you sure you want to clear all form data?')) {
            document.getElementById('bomForm').reset();
            
            // Reset calculated fields
            const totalFields = [
                'fabricTotal', 'collarYarnTotal', 'tippingYarnTotal', 'collarTotal',
                'cuffTotal', 'buttonTotal', 'velvetTapeTotal', 'washCareTotal',
                'wovenSizeTotal', 'polybagTotal'
            ];
            
            totalFields.forEach(fieldId => {
                const field = document.getElementById(fieldId);
                if (field) field.value = '';
            });
            // clear stored INR values as well
            totalFields.forEach(fieldId => {
                const field = document.getElementById(fieldId);
                if (field && field.dataset) field.dataset.inr = '';
            });
            document.getElementById('grandTotal').textContent = formatCurrency(0);

            
            // No default values to reset - form will be completely empty
            
            showAnimatedPopup('Form cleared successfully!', 'success');
        }
    }

    // Show message
    function showMessage(message, type = 'success') {
        // Remove existing messages
        const existingMessages = document.querySelectorAll('.message');
        existingMessages.forEach(msg => msg.remove());
        
        // Create new message
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}`;
        messageDiv.textContent = message;
        
        // Insert at the top of the form
        const form = document.getElementById('bomForm');
        form.insertBefore(messageDiv, form.firstChild);
        
        // Auto remove after 5 seconds
        setTimeout(() => {
            if (messageDiv.parentNode) {
                messageDiv.remove();
            }
        }, 5000);
    }

    // Show animated popup notification
    function showAnimatedPopup(message, type = 'success') {
        // Remove existing popups
        const existingPopups = document.querySelectorAll('.animated-popup');
        existingPopups.forEach(popup => popup.remove());
        
        // Create popup container
        const popup = document.createElement('div');
        popup.className = `animated-popup ${type}`;
        
        // Create popup content
        const popupContent = document.createElement('div');
        popupContent.className = 'popup-content';
        
        // Create icon
        const icon = document.createElement('i');
        icon.className = type === 'success' ? 'fas fa-check-circle' : 'fas fa-exclamation-circle';
        
        // Create message text
        const messageText = document.createElement('span');
        messageText.textContent = message;
        
        // Assemble popup
        popupContent.appendChild(icon);
        popupContent.appendChild(messageText);
        popup.appendChild(popupContent);
        
        // Add to body
        document.body.appendChild(popup);
        
        // Trigger animation
        setTimeout(() => {
            popup.classList.add('show');
        }, 100);
        
        // Auto remove after 2 seconds
        setTimeout(() => {
            popup.classList.remove('show');
            setTimeout(() => {
                if (popup.parentNode) {
                    popup.remove();
                }
            }, 300);
        }, 2000);
    }

    // Auto-calculate on page load
    window.addEventListener('load', function() {
        calculateAllTotals();
    }); 

    // Test function to verify form data collection
    function testFormDataCollection() {
        try {
            console.log('Testing form data collection...');
            const data = collectFormData();
            if (data) {
                console.log('Form data collected successfully:', data);
                showAnimatedPopup('Form data collection test passed!', 'success');
                return true;
            } else {
                console.error('Form data collection returned null');
                showAnimatedPopup('Form data collection test failed!', 'error');
                return false;
            }
        } catch (error) {
            console.error('Error in form data collection:', error);
            showAnimatedPopup('Form data collection test failed with error!', 'error');
            return false;
        }
    }

    // Add test button to the page for debugging
    function addTestButton() {
        // Intentionally disabled in production UI
    }

    // Initialize test button when page loads
    // Test button injection disabled

    // Scroll-triggered animations
    function setupScrollAnimations() {
        const sections = document.querySelectorAll('.section');
        
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('animate-in');
                }
            });
        }, {
            threshold: 0.1,
            rootMargin: '0px 0px -50px 0px'
        });
        
        sections.forEach(section => {
            observer.observe(section);
        });
    }

    // Progress bar for form completion
    function setupProgressBar() {
        // Create progress bar
        const progressContainer = document.createElement('div');
        progressContainer.className = 'progress-container';
        progressContainer.innerHTML = '<div class="progress-bar" id="progressBar"></div>';
        document.body.appendChild(progressContainer);
        
        // Update progress on form changes
        function updateProgress() {
            const form = document.getElementById('bomForm');
            const inputs = form.querySelectorAll('input, select');
            const filledInputs = Array.from(inputs).filter(input => 
                input.value && input.value.trim() !== '' && !input.hasAttribute('readonly')
            );
            
            const progress = (filledInputs.length / inputs.length) * 100;
            const progressBar = document.getElementById('progressBar');
            if (progressBar) {
                progressBar.style.width = `${Math.min(progress, 100)}%`;
            }
        }
        
        // Update progress on input changes
        const form = document.getElementById('bomForm');
        form.addEventListener('input', updateProgress);
        form.addEventListener('change', updateProgress);
        
        // Initial progress update
        updateProgress();
    }

    // Scroll indicator
    function setupScrollIndicator() {
        const sections = document.querySelectorAll('.section');
        if (sections.length === 0) return;
        
        // Create scroll indicator
        const indicator = document.createElement('div');
        indicator.className = 'scroll-indicator';
        
        sections.forEach((section, index) => {
            const dot = document.createElement('div');
            dot.className = 'scroll-dot';
            dot.dataset.section = index;
            dot.addEventListener('click', () => {
                section.scrollIntoView({ behavior: 'smooth' });
            });
            indicator.appendChild(dot);
        });
        
        document.body.appendChild(indicator);
        
        // Update active dot based on scroll position
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    const sectionIndex = Array.from(sections).indexOf(entry.target);
                    const dots = indicator.querySelectorAll('.scroll-dot');
                    dots.forEach((dot, index) => {
                        dot.classList.toggle('active', index === sectionIndex);
                    });
                }
            });
        }, {
            threshold: 0.5
        });
        
        sections.forEach(section => {
            observer.observe(section);
        });
    }

    // Scroll to top button
    function setupScrollToTop() {
        const scrollToTopBtn = document.createElement('button');
        scrollToTopBtn.className = 'scroll-to-top';
        scrollToTopBtn.innerHTML = '<i class="fas fa-arrow-up"></i>';
        scrollToTopBtn.setAttribute('aria-label', 'Scroll to top');
        document.body.appendChild(scrollToTopBtn);
        
        // Show/hide button based on scroll position
        window.addEventListener('scroll', () => {
            if (window.pageYOffset > 300) {
                scrollToTopBtn.classList.add('show');
            } else {
                scrollToTopBtn.classList.remove('show');
            }
        });
        
        // Scroll to top functionality
        scrollToTopBtn.addEventListener('click', () => {
            window.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        });
    }

    // Enhanced button loading states - Simplified to avoid conflicts
    function addButtonLoadingStates() {
        // This function is now simplified to avoid conflicts with main event listeners
        // The main event listeners in setupEventListeners() handle the button clicks
    }

    // Remove button loading state
    function removeButtonLoadingState(buttonId) {
        const button = document.getElementById(buttonId);
        if (button) {
            button.classList.remove('loading', 'excel', 'pdf', 'word');
            button.disabled = false;
        }
    }

    // Parallax effect for background
    function setupParallaxEffect() {
        window.addEventListener('scroll', () => {
            const scrolled = window.pageYOffset;
            const parallax = document.querySelector('body::before');
            if (parallax) {
                const speed = scrolled * 0.5;
                parallax.style.transform = `translateY(${speed}px)`;
            }
        });
    }

    // Initialize enhanced features
    document.addEventListener('DOMContentLoaded', function() {
        addButtonLoadingStates();
        setupParallaxEffect();
    });
