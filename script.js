// Global variables
let formData = {};

// DOM Content Loaded
document.addEventListener('DOMContentLoaded', function () {
    initializeForm();
    setupEventListeners();
    setupCalculationListeners();
});

// Initialize form
function initializeForm() {}

// Setup calculation listeners
function setupCalculationListeners() {
    const fields = [
        'fabricConsumption', 'fabricAmount', 'fabricPerGarment', 'fabricQty',
        'collarYarnAmount', 'collarYarnPerGarment', 'collarYarnQty',
        'tippingYarnAmount', 'tippingYarnPerGarment', 'tippingYarnQty',
        'collarPerGarment', 'collarQty',
        'cuffPerGarment', 'cuffQty',
        'buttonPerGarment', 'buttonQty',
        'velvetTapeMeter', 'velvetTapePerGarment', 'velvetTapeQty',
        'washCareAmount', 'washCarePerGarment', 'washCareQty',
        'wovenSizeAmount', 'wovenSizePerGarment', 'wovenSizeQty',
        'polybagPerGarment', 'polybagQty'
    ];

    fields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', calculateAllTotals);
            el.addEventListener('change', calculateAllTotals);
        }
    });
}

// Generic calculator
function calculateSectionTotal(amountId, perGarmentId, qtyId, totalId) {
    const amount = parseFloat(document.getElementById(amountId)?.value || 0);
    const perGarment = parseFloat(document.getElementById(perGarmentId)?.value || 0);
    const qty = parseFloat(document.getElementById(qtyId)?.value || 0);
    const total = amount * perGarment * qty;

    const totalField = document.getElementById(totalId);
    if (totalField) totalField.value = total.toFixed(2);
    return total;
}

// Main total calculation
function calculateAllTotals() {
    let grandTotal = 0;

    // Fabric (Consumption preferred)
    const fabricConsumption =
        parseFloat(document.getElementById('fabricConsumption')?.value ||
            document.getElementById('fabricAmount')?.value || 0);

    const fabricQty =
        parseFloat(document.getElementById('fabricQty')?.value || 0);

    const fabricTotal = fabricConsumption * fabricQty;
    if (document.getElementById('fabricTotal')) {
        document.getElementById('fabricTotal').value = fabricTotal.toFixed(2);
    }
    grandTotal += fabricTotal;

    grandTotal += calculateSectionTotal('collarYarnAmount', 'collarYarnPerGarment', 'collarYarnQty', 'collarYarnTotal');
    grandTotal += calculateSectionTotal('tippingYarnAmount', 'tippingYarnPerGarment', 'tippingYarnQty', 'tippingYarnTotal');
    grandTotal += calculateSectionTotal('', 'collarPerGarment', 'collarQty', 'collarTotal');
    grandTotal += calculateSectionTotal('', 'cuffPerGarment', 'cuffQty', 'cuffTotal');
    grandTotal += calculateSectionTotal('', 'buttonPerGarment', 'buttonQty', 'buttonTotal');
    grandTotal += calculateSectionTotal('velvetTapeMeter', 'velvetTapePerGarment', 'velvetTapeQty', 'velvetTapeTotal');
    grandTotal += calculateSectionTotal('washCareAmount', 'washCarePerGarment', 'washCareQty', 'washCareTotal');
    grandTotal += calculateSectionTotal('wovenSizeAmount', 'wovenSizePerGarment', 'wovenSizeQty', 'wovenSizeTotal');
    grandTotal += calculateSectionTotal('', 'polybagPerGarment', 'polybagQty', 'polybagTotal');

    document.getElementById('grandTotal').textContent = '₹' + grandTotal.toFixed(2);
    return grandTotal;
}

// Collect form data
function collectFormData() {
    const get = id => document.getElementById(id)?.value || '';

    return {
        orderNumber: get('orderNumber'),
        orderType: get('orderType'),

        fabricType: get('fabricType'),
        fabricConsumption: get('fabricConsumption') || get('fabricAmount'),
        fabricUnit: get('fabricUnit'),
        fabricSupplier: get('fabricSupplier'),
        fabricQty: get('fabricQty'),
        fabricTotal: get('fabricTotal'),

        collarYarnType: get('collarYarnType'),
        collarYarnAmount: get('collarYarnAmount'),
        collarYarnTotal: get('collarYarnTotal'),

        tippingYarnType: get('tippingYarnType'),
        tippingYarnAmount: get('tippingYarnAmount'),
        tippingYarnTotal: get('tippingYarnTotal'),

        buttonTotal: get('buttonTotal'),
        polybagTotal: get('polybagTotal'),

        grandTotal: document.getElementById('grandTotal')?.textContent || '₹0.00'
    };
}

// CSV EXPORT
function createCSVContent(data) {
    const rows = [];
    rows.push(['Twinklub BOM - Bill of Materials']);
    rows.push([]);
    rows.push(['Order Number', data.orderNumber]);
    rows.push(['Order Type', data.orderType]);
    rows.push([]);
    rows.push(['Section', 'Details', 'Consumption', 'Unit', 'Supplier', 'Q.QTY', 'Total']);

    rows.push([
        'Fabric',
        data.fabricType,
        data.fabricConsumption,
        data.fabricUnit,
        data.fabricSupplier,
        data.fabricQty,
        data.fabricTotal
    ]);

    rows.push([]);
    rows.push(['GRAND TOTAL', '', '', '', '', '', data.grandTotal]);

    return rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
}

// TEXT EXPORT
function createTextContent(data) {
    return `
TWINKLUB BOM - BILL OF MATERIALS
===============================

Order Number: ${data.orderNumber}
Order Type  : ${data.orderType}

Fabric:
-------
Type        : ${data.fabricType}
Consumption : ${data.fabricConsumption} ${data.fabricUnit}
Supplier    : ${data.fabricSupplier}
Q.QTY       : ${data.fabricQty}
Total       : ₹${data.fabricTotal}

===============================
GRAND TOTAL: ${data.grandTotal}
===============================
`;
}

// Auto-calc on load
window.addEventListener('load', calculateAllTotals);
