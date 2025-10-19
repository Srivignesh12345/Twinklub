# Twinklub BOM - Bill of Materials System

A comprehensive web-based Bill of Materials (BOM) management system designed specifically for Twinklub's garment manufacturing operations. This application helps calculate material requirements and costs for organic polo t-shirts and other garments.

## 🎯 Features

### Core Functionality
- **Complete BOM Management**: Track all materials required for garment manufacturing
- **Real-time Calculations**: Automatic calculation of material costs and totals
- **Multi-format Export**: Export data to Excel, PDF, and Word formats
- **Responsive Design**: Works seamlessly on desktop, tablet, and mobile devices

### Material Sections
1. **Fabric** - Main fabric requirements with customizable amounts
2. **Collar Yarn** - 100% CTN yarn specifications
3. **Tipping Yarn** - Additional yarn requirements
4. **Collar Measurements** - Size-specific collar dimensions
5. **Cuff Measurements** - Size-specific cuff dimensions
6. **Buttons** - Button material and quantity tracking
7. **Velvet Tapes** - Velvet tape quality and meter requirements
8. **Wash Care Labels** - Label quality and quantity
9. **Woven Size Labels** - Size-specific label requirements
10. **Polybags** - Packaging material specifications

### Export Capabilities
- **Excel Export**: Structured spreadsheet with all BOM data
- **PDF Export**: Professional PDF reports with tables and formatting
- **Word Export**: Formatted Word documents for documentation

## 🚀 Getting Started

### Prerequisites
- Modern web browser (Chrome, Firefox, Safari, Edge)
- No additional software installation required

### Installation
1. Download all files to a local directory
2. Open `index.html` in your web browser
3. The application will load with all features ready to use

### File Structure
```
twinklub-bom/
├── index.html          # Main application file
├── styles.css          # Styling and responsive design
├── script.js           # JavaScript functionality
└── README.md           # This documentation
```

## 📋 How to Use

 # Twinklub BOM — Bill of Materials System

A lightweight web application to calculate and export Bill of Materials (BOM) for garment manufacturing. Built for Twinklub's organic polo workflow but adaptable to other garments.

## Key features

- Real-time material calculations
- Multiple material sections (Fabric, Collar yarn, Buttons, Labels, Packaging, etc.)
- Export to Excel, PDF and Word
- Responsive layout for desktop/tablet/mobile
- Client-side only — no server required

## Preview

Open `index.html` in your browser to view the UI. Styling is in `styles.css` and behavior in `script.js`.

## File structure

```
Project-Twinklub-main/
├── index.html      # Main application UI
├── styles.css      # Styling and visual rules
├── script.js       # JavaScript behavior and exports
└── README.md       # This file
```

## Requirements

- A modern web browser (Chrome, Edge, Firefox, Safari)

## Run (Windows — PowerShell)

1. Open PowerShell in the project folder (e.g., in VS Code terminal or File Explorer -> "Open PowerShell window here").
2. To open directly in your default browser:

```powershell
Start-Process -FilePath .\index.html
```

Optional: run a simple local server (if Python is installed):

```powershell
python -m http.server 5500; # then open http://localhost:5500
```

## Recent changes / Changelog

- 2025-10-19: Hidden the decorative right-side scroll dots (small blue dot) by updating `.scroll-indicator` in `styles.css`.

Revert or customize

- To show the dots again, open `styles.css` and set `.scroll-indicator { display: flex; }`.
- To remove only the blue highlight, change `.scroll-dot.active` (background and box-shadow).

## Troubleshooting

- If style changes don't appear, hard-refresh the browser (Ctrl+F5) or clear cache.
- If export/download fails, check pop-up and download settings in your browser.

---

**Version:** 1.0.1
**Last Updated:** 2025-10-19
**Author:** Twinklub Development Team