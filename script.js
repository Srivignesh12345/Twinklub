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
    /* ================= LOADING BOX ================= */

function showLoadingBox(message = 'Preparing export...') {
    // Remove existing box if any
    hideLoadingBox();

    const box = document.createElement('div');
    box.id = 'exportLoadingBox';
    box.style.position = 'fixed';
    box.style.top = '50%';
    box.style.left = '50%';
    box.style.transform = 'translate(-50%, -50%)';
    box.style.background = '#ffffff';
    box.style.padding = '16px 22px';
    box.style.borderRadius = '10px';
    box.style.boxShadow = '0 10px 30px rgba(0,0,0,0.25)';
    box.style.zIndex = '9999';
    box.style.display = 'flex';
    box.style.alignItems = 'center';
    box.style.gap = '12px';
    box.style.fontFamily = 'Arial, sans-serif';
    box.style.fontSize = '14px';

    box.innerHTML = `
        <div style="
            width:20px;
            height:20px;
            border:3px solid #ddd;
            border-top:3px solid #3498db;
            border-radius:50%;
            animation: spin 1s linear infinite;
        "></div>
        <span>${message}</span>
    `;

    document.body.appendChild(box);

    // Spinner animation
    if (!document.getElementById('loadingSpinnerStyle')) {
        const style = document.createElement('style');
        style.id = 'loadingSpinnerStyle';
        style.innerHTML = `
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);
    }
}

function hideLoadingBox() {
    const box = document.getElementById('exportLoadingBox');
    if (box) box.remove();
}

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
    /* ================= LOADING BOX ================= */

function showLoadingBox(message = 'Preparing export...') {
    // Remove existing box if any
    hideLoadingBox();

    const box = document.createElement('div');
    box.id = 'exportLoadingBox';
    box.style.position = 'fixed';
    box.style.top = '50%';
    box.style.left = '50%';
    box.style.transform = 'translate(-50%, -50%)';
    box.style.background = '#ffffff';
    box.style.padding = '16px 22px';
    box.style.borderRadius = '10px';
    box.style.boxShadow = '0 10px 30px rgba(0,0,0,0.25)';
    box.style.zIndex = '9999';
    box.style.display = 'flex';
    box.style.alignItems = 'center';
    box.style.gap = '12px';
    box.style.fontFamily = 'Arial, sans-serif';
    box.style.fontSize = '14px';

    box.innerHTML = `
        <div style="
            width:20px;
            height:20px;
            border:3px solid #ddd;
            border-top:3px solid #3498db;
            border-radius:50%;
            animation: spin 1s linear infinite;
        "></div>
        <span>${message}</span>
    `;

    document.body.appendChild(box);

    // Spinner animation
    if (!document.getElementById('loadingSpinnerStyle')) {
        const style = document.createElement('style');
        style.id = 'loadingSpinnerStyle';
        style.innerHTML = `
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);
    }
}

function hideLoadingBox() {
    const box = document.getElementById('exportLoadingBox');
    if (box) box.remove();
}

    showLoadingBox('Preparing PDF file...');

    setTimeout(() => {
        const data = collectFormData();

        if (!window.jspdf || !window.jspdf.jsPDF) {
            downloadFile(createTextContent(data), 'txt', 'Twinklub_BOM.txt');
            hideLoadingBox();
            return;
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        /* ================= TITLE ================= */
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(18);
        doc.text('Twinklub BOM - Bill of Materials', 105, 20, { align: 'center' });

        /* ================= ORDER INFO ================= */
        doc.setFontSize(11);
        doc.setFont('helvetica', 'normal');
        doc.text(`Order Number : ${data.orderNumber || '-'}`, 20, 35);
        doc.text(`Order Type      : ${data.orderType || '-'}`, 20, 42);

        /* ================= TABLE ================= */
        doc.autoTable({
            startY: 55,
            theme: 'grid',
            head: [[
                'Section',
                'Details',
                'Consumption',
                'Unit',
                'Supplier',
                'Q.QTY',
                'Total (Rs)'
            ]],
            body: [[
                'Fabric',
                data.fabricType || '-',
                data.fabricConsumption || '-',
                data.fabricUnit || '-',
                data.fabricSupplier || '-',
                data.fabricQty || '-',
                data.fabricTotal || '0.00'
            ]],
            styles: {
                fontSize: 10,
                cellPadding: 4,
                halign: 'center',
                valign: 'middle'
            },
            headStyles: {
                fillColor: [52, 152, 219],
                textColor: 255,
                halign: 'center'
            },
            columnStyles: {
                0: { halign: 'left' },
                1: { halign: 'left' }
            }
        });

        /* ================= GRAND TOTAL ================= */
        const finalY = doc.lastAutoTable.finalY + 12;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(14);
        doc.text(
            `GRAND TOTAL : Rs ${String(data.grandTotal).replace(/[^\d.]/g, '')}/-`,
            105,
            finalY,
            { align: 'center' }
        );

        doc.save('Twinklub_BOM.pdf');
        hideLoadingBox();
    }, 600);
}

// ================= WORD =================
function exportToWord() {
    /* ================= LOADING BOX ================= */

function showLoadingBox(message = 'Preparing export...') {
    // Remove existing box if any
    hideLoadingBox();

    const box = document.createElement('div');
    box.id = 'exportLoadingBox';
    box.style.position = 'fixed';
    box.style.top = '50%';
    box.style.left = '50%';
    box.style.transform = 'translate(-50%, -50%)';
    box.style.background = '#ffffff';
    box.style.padding = '16px 22px';
    box.style.borderRadius = '10px';
    box.style.boxShadow = '0 10px 30px rgba(0,0,0,0.25)';
    box.style.zIndex = '9999';
    box.style.display = 'flex';
    box.style.alignItems = 'center';
    box.style.gap = '12px';
    box.style.fontFamily = 'Arial, sans-serif';
    box.style.fontSize = '14px';

    box.innerHTML = `
        <div style="
            width:20px;
            height:20px;
            border:3px solid #ddd;
            border-top:3px solid #3498db;
            border-radius:50%;
            animation: spin 1s linear infinite;
        "></div>
        <span>${message}</span>
    `;

    document.body.appendChild(box);

    // Spinner animation
    if (!document.getElementById('loadingSpinnerStyle')) {
        const style = document.createElement('style');
        style.id = 'loadingSpinnerStyle';
        style.innerHTML = `
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);
    }
}

function hideLoadingBox() {
    const box = document.getElementById('exportLoadingBox');
    if (box) box.remove();
}

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
