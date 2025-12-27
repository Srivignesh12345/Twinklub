// ================= GLOBAL =================
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    setupCalculationListeners();
    calculateAllTotals();
});

// ================= EVENTS =================
function setupEventListeners() {
    document.getElementById('exportExcelBtn')?.addEventListener('click', e => {
        e.preventDefault();
        exportToExcel();
    });

    document.getElementById('exportPDFBtn')?.addEventListener('click', e => {
        e.preventDefault();
        exportToPDF();
    });

    document.getElementById('exportWordBtn')?.addEventListener('click', e => {
        e.preventDefault();
        exportToWord();
    });
}

// ================= CALCULATION =================
function setupCalculationListeners() {
    ['fabricConsumption', 'fabricAmount', 'fabricQty']
        .forEach(id => {
            document.getElementById(id)?.addEventListener('input', calculateAllTotals);
        });
}

function calculateAllTotals() {
    const consumption =
        parseFloat(document.getElementById('fabricConsumption')?.value ||
        document.getElementById('fabricAmount')?.value || 0);

    const qty = parseFloat(document.getElementById('fabricQty')?.value || 0);
    const total = consumption * qty;

    document.getElementById('fabricTotal').value = total.toFixed(2);
    document.getElementById('grandTotal').textContent = `₹${total.toFixed(2)}`;

    return total;
}

// ================= DATA =================
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
        grandTotal: document.getElementById('grandTotal').textContent
    };
}

// ================= CSV / EXCEL =================
function exportToExcel() {
    const data = collectFormData();

    const rows = [
        ['Twinklub BOM - Bill of Materials'],
        [],
        ['Order Number', data.orderNumber],
        ['Order Type', data.orderType],
        [],
        ['Section', 'Details', 'Consumption', 'Unit', 'Supplier', 'Q.QTY', 'Total'],
        [
            'Fabric',
            data.fabricType,
            data.fabricConsumption,
            data.fabricUnit,
            data.fabricSupplier,
            data.fabricQty,
            data.fabricTotal
        ],
        [],
        ['GRAND TOTAL', '', '', '', '', '', data.grandTotal]
    ];

    if (window.XLSX) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, 'BOM');
        XLSX.writeFile(wb, 'Twinklub_BOM.xlsx');
    } else {
        downloadFile(rows.map(r => r.join(',')).join('\n'), 'csv', 'Twinklub_BOM.csv');
    }
}

// ================= PDF =================
function exportToPDF() {
    const data = collectFormData();

    if (!window.jspdf) {
        downloadFile(createTextContent(data), 'txt', 'Twinklub_BOM.txt');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    doc.text('Twinklub BOM - Bill of Materials', 20, 20);
    doc.text(`Consumption: ${data.fabricConsumption} ${data.fabricUnit}`, 20, 40);
    doc.text(`Q.QTY: ${data.fabricQty}`, 20, 50);
    doc.text(`Total: ${data.grandTotal}`, 20, 60);

    doc.save('Twinklub_BOM.pdf');
}

// ================= WORD =================
function exportToWord() {
    const data = collectFormData();

    if (!window.docx) {
        downloadFile(createTextContent(data), 'txt', 'Twinklub_BOM.txt');
        return;
    }

    const { Document, Packer, Paragraph } = window.docx;

    const doc = new Document({
        sections: [{
            children: [
                new Paragraph('Twinklub BOM - Bill of Materials'),
                new Paragraph(`Consumption: ${data.fabricConsumption} ${data.fabricUnit}`),
                new Paragraph(`Quantity: ${data.fabricQty}`),
                new Paragraph(`Total: ${data.grandTotal}`)
            ]
        }]
    });

    Packer.toBlob(doc).then(blob => {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'Twinklub_BOM.docx';
        link.click();
    });
}

// ================= TEXT / FALLBACK =================
function createTextContent(data) {
    return `
TWINKLUB BOM - BILL OF MATERIALS
--------------------------------
Fabric: ${data.fabricType}
Consumption: ${data.fabricConsumption} ${data.fabricUnit}
Quantity: ${data.fabricQty}
Total: ${data.grandTotal}
--------------------------------
`;
}

function downloadFile(content, type, filename) {
    const blob = new Blob([content], { type: 'text/plain' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}
