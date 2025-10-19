// Global variables
let formData = {};

// DOM Content Loaded
document.addEventListener('DOMContentLoaded', function() {
    initializeForm();
    setupEventListeners();
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
    const pdfBtn = document.getElementById('exportPDFBtn');
    const wordBtn = document.getElementById('exportWordBtn');
    const clearBtn = document.getElementById('clearBtn');
    
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
    
    if (pdfBtn) {
        pdfBtn.addEventListener('click', function(e) {
            e.preventDefault();
            console.log('PDF button clicked!');
            exportToPDF();
        });
        console.log('PDF button listener added');
    } else {
        console.error('PDF button not found!');
    }
    
    if (wordBtn) {
        wordBtn.addEventListener('click', function(e) {
            e.preventDefault();
            console.log('Word button clicked!');
            exportToWord();
        });
        console.log('Word button listener added');
    } else {
        console.error('Word button not found!');
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
		totalField.value = total.toFixed(2);
	}
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
    grandTotal += calculateSectionTotal('', 'collarPerGarment', 'collarQty', 'collarTotal');
    
    // Cuff (only per garment * qty)
    grandTotal += calculateSectionTotal('', 'cuffPerGarment', 'cuffQty', 'cuffTotal');
    
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
    document.getElementById('grandTotal').textContent = '\u20B9' + grandTotal.toFixed(2);
    
    return grandTotal;
}

// Collect form data
function collectFormData() {
    const form = document.getElementById('bomForm');
    const formData = new FormData(form);
    const data = {};
    
    for (let [key, value] of formData.entries()) {
        data[key] = value;
    }
    
    // Safe getter for optional inputs
    const getVal = (id) => {
        const el = document.getElementById(id);
        return el ? el.value : '';
    };
    
    // Add calculated totals (guard against missing fields like Fabric 2)
    data.fabricTotal = getVal('fabricTotal');
    data.collarYarnTotal = getVal('collarYarnTotal');
    data.tippingYarnTotal = getVal('tippingYarnTotal');
    data.collarTotal = getVal('collarTotal');
    data.cuffTotal = getVal('cuffTotal');
    data.buttonTotal = getVal('buttonTotal');
    data.velvetTapeTotal = getVal('velvetTapeTotal');
    data.washCareTotal = getVal('washCareTotal');
    data.wovenSizeTotal = getVal('wovenSizeTotal');
    data.polybagTotal = getVal('polybagTotal');
    data.grandTotal = document.getElementById('grandTotal').textContent;
    
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
    rows.push(['Fabric 1', data.fabricType || '', data.fabricAmount || '', data.fabricUnit || '', data.fabricPerGarment || '', data.fabricSupplier || '', data.fabricQty || '', data.fabricTotal || '']);
    rows.push(['Collar Yarn', data.collarYarnType || '', data.collarYarnAmount || '', data.collarYarnUnit || '', data.collarYarnPerGarment || '', data.collarYarnSupplier || '', data.collarYarnQty || '', data.collarYarnTotal || '']);
    rows.push(['Tipping Yarn', data.tippingYarnType || '', data.tippingYarnAmount || '', data.tippingYarnUnit || '', data.tippingYarnPerGarment || '', data.tippingYarnSupplier || '', data.tippingYarnQty || '', data.tippingYarnTotal || '']);
    
    // Add collar measurements if quantity > 0
    const collarQty = parseFloat(data.collarQty || '');
    if (!isNaN(collarQty) && collarQty > 0) {
        const collarDetails = `${data.collarSize || ''}${data.collarDimensions ? ' - ' + data.collarDimensions : ''}`;
        rows.push(['Collar Measurements', collarDetails, '', data.collarUnit || '', data.collarPerGarment || '', data.collarSupplier || '', data.collarQty || '', data.collarTotal || '']);
    }
    
    // Add cuff measurements if quantity > 0
    const cuffQty = parseFloat(data.cuffQty || '');
    if (!isNaN(cuffQty) && cuffQty > 0) {
        const cuffDetails = `${data.cuffSize || ''}${data.cuffDimensions ? ' - ' + data.cuffDimensions : ''}`;
        rows.push(['Cuff Measurements', cuffDetails, '', data.cuffUnit || '', data.cuffPerGarment || '', data.cuffSupplier || '', data.cuffQty || '', data.cuffTotal || '']);
    }
    
    rows.push(['Buttons', data.buttonMaterial || '', data.buttonAmount || '', data.buttonUnit || '', data.buttonPerGarment || '', data.buttonSupplier || '', data.buttonQty || '', data.buttonTotal || '']);
    rows.push(['Velvet Tapes', data.velvetTapeQuality || '', data.velvetTapeMeter || '', data.velvetTapeUnit || '', data.velvetTapePerGarment || '', data.velvetTapeSupplier || '', data.velvetTapeQty || '', data.velvetTapeTotal || '']);
    rows.push(['Wash Care Labels', data.washCareQuality || '', data.washCareAmount || '', data.washCareUnit || '', data.washCarePerGarment || '', data.washCareSupplier || '', data.washCareQty || '', data.washCareTotal || '']);
    rows.push(['Polybags', data.polybagQuality || '', data.polybagSize || '', data.polybagUnit || '', data.polybagPerGarment || '', data.polybagSupplier || '', data.polybagQty || '', data.polybagTotal || '']);
    rows.push(['']);
    rows.push(['GRAND TOTAL:', data.grandTotal || '₹0.00']);
    
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
Total: ${data.fabricTotal || '₹0.00'}

Collar Yarn: ${data.collarYarnType || ''}
Amount: ${data.collarYarnAmount || ''} ${data.collarYarnUnit || ''}
Per Garment: ${data.collarYarnPerGarment || ''}
Supplier: ${data.collarYarnSupplier || ''}
Q.QTY: ${data.collarYarnQty || ''}
Total: ${data.collarYarnTotal || '₹0.00'}

Tipping Yarn: ${data.tippingYarnType || ''}
Amount: ${data.tippingYarnAmount || ''} ${data.tippingYarnUnit || ''}
Per Garment: ${data.tippingYarnPerGarment || ''}
Supplier: ${data.tippingYarnSupplier || ''}
Q.QTY: ${data.tippingYarnQty || ''}
Total: ${data.tippingYarnTotal || '₹0.00'}

Buttons: ${data.buttonMaterial || ''}
Amount: ${data.buttonAmount || ''} ${data.buttonUnit || ''}
Per Garment: ${data.buttonPerGarment || ''}
Supplier: ${data.buttonSupplier || ''}
Q.QTY: ${data.buttonQty || ''}
Total: ${data.buttonTotal || '₹0.00'}

Velvet Tapes: ${data.velvetTapeQuality || ''}
Amount: ${data.velvetTapeMeter || ''} ${data.velvetTapeUnit || ''}
Per Garment: ${data.velvetTapePerGarment || ''}
Supplier: ${data.velvetTapeSupplier || ''}
Q.QTY: ${data.velvetTapeQty || ''}
Total: ${data.velvetTapeTotal || '₹0.00'}

Wash Care Labels: ${data.washCareQuality || ''}
Amount: ${data.washCareAmount || ''} ${data.washCareUnit || ''}
Per Garment: ${data.washCarePerGarment || ''}
Supplier: ${data.washCareSupplier || ''}
Q.QTY: ${data.washCareQty || ''}
Total: ${data.washCareTotal || '₹0.00'}

Polybags: ${data.polybagQuality || ''}
Size: ${data.polybagSize || ''} ${data.polybagUnit || ''}
Per Garment: ${data.polybagPerGarment || ''}
Supplier: ${data.polybagSupplier || ''}
Q.QTY: ${data.polybagQty || ''}
Total: ${data.polybagTotal || '₹0.00'}

================================
GRAND TOTAL: ${data.grandTotal || '₹0.00'}
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
            [''],
            ['QUALITY', 'DETAILS', 'UNIT OF MSMTS', 'PER GARMENT', 'SUPPLIER', 'O.QTY', 'TOTAL AMOUNT'],
            [''],
            // Each row below must have exactly 7 columns to match the header
            ['FABRIC 1', data.fabricType || '', data.fabricUnit || '', data.fabricPerGarment || '', data.fabricSupplier || '', data.fabricQty || '', data.fabricTotal || ''],
            [''],
            ['COLLAR YARN', data.collarYarnType || '', data.collarYarnUnit || '', data.collarYarnPerGarment || '', data.collarYarnSupplier || '', data.collarYarnQty || '', data.collarYarnTotal || ''],
            [''],
            ['TIPPING YARN', data.tippingYarnType || '', data.tippingYarnUnit || '', data.tippingYarnPerGarment || '', data.tippingYarnSupplier || '', data.tippingYarnQty || '', data.tippingYarnTotal || ''],
            [''],
            // Collar measurements - only show if quantity > 0
            ...((() => {
                const q = parseFloat(data.collarQty || '');
                if (!isNaN(q) && q > 0) {
                    const collarDetails = `${data.collarSize || ''}${data.collarDimensions ? ' - ' + data.collarDimensions : ''}`;
                    return [['COLLAR MSMTS', collarDetails, data.collarUnit || '', data.collarPerGarment || '', data.collarSupplier || '', data.collarQty || '', data.collarTotal || ''], ['']];
                }
                return [];
            })()),
            // Cuff measurements - only show if quantity > 0
            ...((() => {
                const q = parseFloat(data.cuffQty || '');
                if (!isNaN(q) && q > 0) {
                    const cuffDetails = `${data.cuffSize || ''}${data.cuffDimensions ? ' - ' + data.cuffDimensions : ''}`;
                    return [['CUFF MSMTS', cuffDetails, data.cuffUnit || '', data.cuffPerGarment || '', data.cuffSupplier || '', data.cuffQty || '', data.cuffTotal || ''], ['']];
                }
                return [];
            })()),
            ['BUTTONS', data.buttonMaterial || '', data.buttonUnit || '', data.buttonPerGarment || '', data.buttonSupplier || '', data.buttonQty || '', data.buttonTotal || ''],
            [''],
            ['VELVET TAPES', data.velvetTapeQuality || '', data.velvetTapeUnit || '', data.velvetTapePerGarment || '', data.velvetTapeSupplier || '', data.velvetTapeQty || '', data.velvetTapeTotal || ''],
            [''],
            ['WASH CARE LABELS', data.washCareQuality || '', data.washCareUnit || '', data.washCarePerGarment || '', data.washCareSupplier || '', data.washCareQty || '', data.washCareTotal || ''],
            [''],
            // Woven size labels - only show if quantity > 0
            ...((() => {
                const q = parseFloat(data.wovenSizeQty || '');
                if (!isNaN(q) && q > 0) {
                    return [['SIZE LABELS (WVN SIZE LBL)', `WVN SIZE LBL - ${data.wovenSizeLabel || ''}`, data.wovenSizeUnit || '', data.wovenSizePerGarment || '', data.wovenSizeSupplier || '', data.wovenSizeQty || '', data.wovenSizeTotal || ''], ['']];
                }
                return [];
            })()),
            ['POLYBAGS', (data.polybagQuality || '') + (data.polybagSize ? ` - ${data.polybagSize}` : ''), data.polybagUnit || '', data.polybagPerGarment || '', data.polybagSupplier || '', data.polybagQty || '', data.polybagTotal || ''],
            [''],
            ['GRAND TOTAL:', (data.grandTotal && String(data.grandTotal).trim()) ? String(data.grandTotal) : ('\u20B9' + (grandTotal ? grandTotal.toFixed(2) : '0.00'))]
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

// Export to PDF
function exportToPDF() {
    console.log('PDF export function called!');
    try {
        // Show loading overlay
        showLoadingOverlay('Exporting to PDF...', 'Creating professional PDF document');
        
        // Ensure totals are up to date
        calculateAllTotals();

        const data = collectFormData();
        console.log('Form data collected for PDF:', data);
        
        // Check jsPDF availability
        if (!window.jspdf || !window.jspdf.jsPDF) {
            console.warn('jsPDF not available, using TXT fallback');
            const textContent = createTextContent(data);
            downloadFile(textContent, 'txt', 'Twinklub_BOM_Export.txt');
            hideLoadingOverlay();
            showAnimatedPopup('Text file exported successfully!', 'success');
            return;
        }
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        
        // Add title
        doc.setFontSize(20);
        doc.setFont('helvetica', 'bold');
        doc.text('Twinklub BOM - Bill of Materials', 20, 20);
        
        // Add order information
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('Order Information:', 20, 40);
        doc.setFont('helvetica', 'normal');
        doc.text(`Order Number: ${data.orderNumber || ''}`, 20, 50);
        doc.text(`Order Type: ${data.orderType || ''}`, 20, 60);
        
        // Create table data (only visualize sizes when user provided input)
        const tableData = [];
        tableData.push(['Section', 'Details', 'Amount', 'Unit', 'Per Garment', 'Supplier', 'Q.QTY', 'Total (₹)']);
        tableData.push(['Fabric 1', data.fabricType || '', data.fabricAmount || '', data.fabricUnit || '', data.fabricPerGarment || '', data.fabricSupplier || '', data.fabricQty || '', data.fabricTotal || '']);
        // Removed Fabric 2 row since there's no Fabric 2 section in input
        tableData.push(['Collar Yarn', data.collarYarnType || '', data.collarYarnAmount || '', data.collarYarnUnit || '', data.collarYarnPerGarment || '', data.collarYarnSupplier || '', data.collarYarnQty || '', data.collarYarnTotal || '']);
        tableData.push(['Tipping Yarn', data.tippingYarnType || '', data.tippingYarnAmount || '', data.tippingYarnUnit || '', data.tippingYarnPerGarment || '', data.tippingYarnSupplier || '', data.tippingYarnQty || '', data.tippingYarnTotal || '']);
        const collarQtyNum = parseFloat(data.collarQty || '');
        if (!isNaN(collarQtyNum) && collarQtyNum > 0) {
            tableData.push(['Collar', `${data.collarSize || ''}${data.collarDimensions ? ' - ' + data.collarDimensions : ''}`, '', data.collarUnit || '', data.collarPerGarment || '', data.collarSupplier || '', data.collarQty || '', data.collarTotal || '']);
        }
        const cuffQtyNum = parseFloat(data.cuffQty || '');
        if (!isNaN(cuffQtyNum) && cuffQtyNum > 0) {
            tableData.push(['Cuff', `${data.cuffSize || ''}${data.cuffDimensions ? ' - ' + data.cuffDimensions : ''}`, '', data.cuffUnit || '', data.cuffPerGarment || '', data.cuffSupplier || '', data.cuffQty || '', data.cuffTotal || '']);
        }
        tableData.push(['Buttons', data.buttonMaterial || '', data.buttonAmount || '', data.buttonUnit || '', data.buttonPerGarment || '', data.buttonSupplier || '', data.buttonQty || '', data.buttonTotal || '']);
        tableData.push(['Velvet Tapes', data.velvetTapeQuality || '', data.velvetTapeMeter || '', data.velvetTapeUnit || '', data.velvetTapePerGarment || '', data.velvetTapeSupplier || '', data.velvetTapeQty || '', data.velvetTapeTotal || '']);
        tableData.push(['Wash Care Labels', data.washCareQuality || '', data.washCareAmount || '', data.washCareUnit || '', data.washCarePerGarment || '', data.washCareSupplier || '', data.washCareQty || '', data.washCareTotal || '']);
        const wvnQtyNum = parseFloat(data.wovenSizeQty || '');
        if (!isNaN(wvnQtyNum) && wvnQtyNum > 0) {
            tableData.push(['Woven Size Labels', data.wovenSizeLabel || '', data.wovenSizeAmount || '', data.wovenSizeUnit || '', data.wovenSizePerGarment || '', data.wovenSizeSupplier || '', data.wovenSizeQty || '', data.wovenSizeTotal || '']);
        }
        tableData.push(['Polybags', data.polybagQuality || '', data.polybagSize || '', data.polybagUnit || '', data.polybagPerGarment || '', data.polybagSupplier || '', data.polybagQty || '', data.polybagTotal || '']);
        
        // Add table
        doc.autoTable({
            startY: 80,
            head: [tableData[0]],
            body: tableData.slice(1),
            theme: 'grid',
            headStyles: { fillColor: [52, 152, 219] },
            styles: { fontSize: 8 }
        });
        
        // Add grand total
        const finalY = (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : 80) + 10;
        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        const gt = (data.grandTotal && String(data.grandTotal).trim()) ? String(data.grandTotal).replace(/[^0-9.]/g, '') : (typeof calculateAllTotals === 'function' ? calculateAllTotals().toFixed(2) : '0.00');
        // Always use 'Rs ' and append '/-' per requested format
        doc.text(`GRAND TOTAL: Rs ${gt}/-`, 20, finalY);
        
        // Generate filename
        const orderNumber = data.orderNumber || 'BOM';
        const filename = `Twinklub_BOM_${orderNumber}_${new Date().toISOString().split('T')[0]}.pdf`;
        
        // Simulate processing time for better UX
        setTimeout(() => {
        doc.save(filename);
            hideLoadingOverlay();
            removeButtonLoadingState('exportPDFBtn');
        showAnimatedPopup('PDF file exported successfully!', 'success');
        }, 2000);
        
    } catch (error) {
        console.error('Error exporting to PDF:', error);
        hideLoadingOverlay();
        removeButtonLoadingState('exportPDFBtn');
        showAnimatedPopup('Error exporting to PDF. Using fallback export.', 'error');
        try { createTextFileFallback(collectFormData()); } catch (_) {}
    }
}

// Export to Word
function exportToWord() {
    console.log('Word export function called!');
    try {
        // Show loading overlay
        showLoadingOverlay('Exporting to Word...', 'Generating Word document with your BOM data');
        
        // Ensure totals are up to date
        calculateAllTotals();
        const data = collectFormData();
        console.log('Form data collected for Word:', data);
        
        // Check if docx library is loaded
        if (typeof window.docx === 'undefined') {
            console.warn('docx library not loaded, using text file fallback');
            const textContent = createTextContent(data);
            downloadFile(textContent, 'txt', 'Twinklub_BOM_Export.txt');
            hideLoadingOverlay();
            showAnimatedPopup('Text file exported successfully!', 'success');
            return;
        }
        
        console.log('docx library available:', window.docx);
        
        // Try to use the library
        try {
            const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType } = window.docx;
            console.log('docx library components loaded');
            
            // Create a simple document first to test
            const simpleDoc = new Document({
                sections: [{
                    properties: {},
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Twinklub BOM - Bill of Materials',
                                    bold: true,
                                    size: 32
                                })
                            ],
                            alignment: AlignmentType.CENTER
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: '', break: 1 })]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Order Number: ${data.orderNumber || ''}`,
                                    size: 20
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Order Type: ${data.orderType || ''}`,
                                    size: 20
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: '', break: 1 })]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'BOM Data:',
                                    bold: true,
                                    size: 24
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Fabric 1: ${data.fabricType || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Collar Yarn: ${data.collarYarnType || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Tipping Yarn: ${data.tippingYarnType || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Buttons: ${data.buttonMaterial || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Velvet Tapes: ${data.velvetTapeQuality || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Wash Care Labels: ${data.washCareQuality || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Polybags: ${data.polybagQuality || ''}`,
                                    size: 16
                                })
                            ]
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: '', break: 1 })]
                        }),
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: `GRAND TOTAL: ${data.grandTotal || '₹0.00'}`,
                                    bold: true,
                                    size: 28
                                })
                            ],
                            alignment: AlignmentType.CENTER
                        })
                    ]
                }]
            });
            
            console.log('Simple document created, generating blob...');
            
            // Generate and download file
            Packer.toBlob(simpleDoc).then(blob => {
                console.log('Blob generated successfully');
                const orderNumber = data.orderNumber || 'BOM';
                const filename = `Twinklub_BOM_${orderNumber}_${new Date().toISOString().split('T')[0]}.docx`;
                
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                link.click();
                window.URL.revokeObjectURL(url);
                
                console.log('Word file exported successfully');
                hideLoadingOverlay();
                removeButtonLoadingState('exportWordBtn');
                showAnimatedPopup('Word file exported successfully!', 'success');
            }).catch(error => {
                console.error('Error in Packer.toBlob:', error);
                hideLoadingOverlay();
                removeButtonLoadingState('exportWordBtn');
                showAnimatedPopup('Error generating Word file. Using fallback method.', 'error');
                createTextFileFallback(data);
            });
            
        } catch (libraryError) {
            console.error('Error with docx library:', libraryError);
            hideLoadingOverlay();
            removeButtonLoadingState('exportWordBtn');
            showAnimatedPopup('Error with Word library. Using fallback method.', 'error');
            createTextFileFallback(data);
        }
        
    } catch (error) {
        console.error('Error exporting to Word:', error);
        hideLoadingOverlay();
        removeButtonLoadingState('exportWordBtn');
        showAnimatedPopup('Error exporting to Word. Using fallback method.', 'error');
        createTextFileFallback(data);
    }
}

// Fallback function to create a text file with BOM data
function createTextFileFallback(data) {
    try {
        console.log('Creating fallback text file...');
        
        // Create formatted text content
        const content = `
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
Total: ${data.fabricTotal || '₹0.00'}

Collar Yarn: ${data.collarYarnType || ''}
Amount: ${data.collarYarnAmount || ''} ${data.collarYarnUnit || ''}
Per Garment: ${data.collarYarnPerGarment || ''}
Supplier: ${data.collarYarnSupplier || ''}
Q.QTY: ${data.collarYarnQty || ''}
Total: ${data.collarYarnTotal || '₹0.00'}

Tipping Yarn: ${data.tippingYarnType || ''}
Amount: ${data.tippingYarnAmount || ''} ${data.tippingYarnUnit || ''}
Per Garment: ${data.tippingYarnPerGarment || ''}
Supplier: ${data.tippingYarnSupplier || ''}
Q.QTY: ${data.tippingYarnQty || ''}
Total: ${data.tippingYarnTotal || '₹0.00'}

Collar Measurements: ${data.collarSize || ''} - ${data.collarDimensions || ''}
Unit: ${data.collarUnit || ''}
Per Garment: ${data.collarPerGarment || ''}
Supplier: ${data.collarSupplier || ''}
Q.QTY: ${data.collarQty || ''}
Total: ${data.collarTotal || '₹0.00'}

Cuff Measurements: ${data.cuffSize || ''} - ${data.cuffDimensions || ''}
Unit: ${data.cuffUnit || ''}
Per Garment: ${data.cuffPerGarment || ''}
Supplier: ${data.cuffSupplier || ''}
Q.QTY: ${data.cuffQty || ''}
Total: ${data.cuffTotal || '₹0.00'}

Buttons: ${data.buttonMaterial || ''}
Amount: ${data.buttonAmount || ''} ${data.buttonUnit || ''}
Per Garment: ${data.buttonPerGarment || ''}
Supplier: ${data.buttonSupplier || ''}
Q.QTY: ${data.buttonQty || ''}
Total: ${data.buttonTotal || '₹0.00'}

Velvet Tapes: ${data.velvetTapeQuality || ''}
Amount: ${data.velvetTapeMeter || ''} ${data.velvetTapeUnit || ''}
Per Garment: ${data.velvetTapePerGarment || ''}
Supplier: ${data.velvetTapeSupplier || ''}
Q.QTY: ${data.velvetTapeQty || ''}
Total: ${data.velvetTapeTotal || '₹0.00'}

Wash Care Labels: ${data.washCareQuality || ''}
Amount: ${data.washCareAmount || ''} ${data.washCareUnit || ''}
Per Garment: ${data.washCarePerGarment || ''}
Supplier: ${data.washCareSupplier || ''}
Q.QTY: ${data.washCareQty || ''}
Total: ${data.washCareTotal || '₹0.00'}

Woven Size Labels: ${data.wovenSizeLabel || ''}
Amount: ${data.wovenSizeAmount || ''} ${data.wovenSizeUnit || ''}
Per Garment: ${data.wovenSizePerGarment || ''}
Supplier: ${data.wovenSizeSupplier || ''}
Q.QTY: ${data.wovenSizeQty || ''}
Total: ${data.wovenSizeTotal || '₹0.00'}

Polybags: ${data.polybagQuality || ''}
Size: ${data.polybagSize || ''} ${data.polybagUnit || ''}
Per Garment: ${data.polybagPerGarment || ''}
Supplier: ${data.polybagSupplier || ''}
Q.QTY: ${data.polybagQty || ''}
Total: ${data.polybagTotal || '₹0.00'}

================================
GRAND TOTAL: ${data.grandTotal || '₹0.00'}
================================
        `;
        
        // Create and download text file
        const orderNumber = data.orderNumber || 'BOM';
        const filename = `Twinklub_BOM_${orderNumber}_${new Date().toISOString().split('T')[0]}.txt`;
        
        const blob = new Blob([content], { type: 'text/plain' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);
        
        console.log('Fallback text file created successfully');
        showAnimatedPopup('Text file exported successfully! (Word library unavailable)', 'success');
        
    } catch (error) {
        console.error('Error creating fallback file:', error);
        showAnimatedPopup('Error creating export file. Please try again.', 'error');
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
        
        document.getElementById('grandTotal').textContent = '₹0.00';
        
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