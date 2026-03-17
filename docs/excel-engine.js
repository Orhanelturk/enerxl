/**
 * ExcelEngine - Export Project Data to Microsoft Excel
 */
const ExcelEngine = {
    init() {
        const btn = document.getElementById('btn-share-excel');
        if (btn) {
            btn.addEventListener('click', () => this.exportToExcel());
        }
        console.log("Excel Engine Initialized");
    },

    async exportToExcel() {
        // 1. Check if we are running inside Excel
        if (typeof Office !== 'undefined' && Office.context && Office.context.host === Office.HostType.Excel) {
            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const boq = this.gatherBoq();
                    
                    if (boq.length === 0) {
                        alert("No project data available. Please generate a layout first.");
                        return;
                    }

                    // Get current selection to use as starting point, or default to A1
                    const selection = context.workbook.getSelectedRange();
                    selection.load("address, rowIndex, columnIndex");
                    await context.sync();

                    const startRow = selection.rowIndex;
                    const startCol = selection.columnIndex;

                    const range = sheet.getRangeByIndexes(startRow, startCol, boq.length + 1, 3);
                    
                    const values = [["Description", "Quantity", "Unit"]];
                    boq.forEach(item => {
                        values.push([item.desc, item.qty, item.unit]);
                    });

                    range.values = values;
                    range.format.autofitColumns();
                    
                    // Style Header
                    const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, 3);
                    headerRange.format.font.bold = true;
                    headerRange.format.fill.color = "#E2E8F0";
                    headerRange.format.borders.getItem('Bottom').style = 'Continuous';

                    await context.sync();
                    alert("BOQ successfully exported to Excel cells!");
                });
            } catch (error) {
                console.error("Excel.run Error:", error);
                this.showFallbackBoq();
            }
        } else {
            console.log("Not running in Excel Host. Showing fallback BOQ.");
            this.showFallbackBoq();
        }
    },

    gatherBoq() {
        const lStats = window._layoutStats || {};
        const eStats = window._electricalStats || {};
        const cStats = (window.CablesEngine && window.CablesEngine.stats) || {};

        let totalMods = 0;
        let totalInvs = 0;
        Object.values(lStats).forEach(s => {
            totalMods += (s.moduleCount || 0);
            totalInvs += (s.invCount || 0);
        });

        // Sum up from all areas if multiple
        const boq = [];
        if (totalMods > 0) boq.push({ desc: "PV Modules (Poly/Mono)", qty: totalMods, unit: "pcs" });
        if (totalInvs > 0) boq.push({ desc: "Solar Inverters (String/Central)", qty: totalInvs, unit: "pcs" });
        
        if (eStats.totalStations > 0) boq.push({ desc: "MV Substation / Transformers", qty: eStats.totalStations, unit: "pcs" });
        if (eStats.totalPanels > 0) boq.push({ desc: "LV Distribution Panels", qty: eStats.totalPanels, unit: "pcs" });
        if (eStats.totalMainTr > 0) boq.push({ desc: "Main Power Transformers", qty: eStats.totalMainTr, unit: "pcs" });
        
        if (eStats.totalHvIncomers > 0) boq.push({ desc: "HV Switchgear Incomers", qty: eStats.totalHvIncomers, unit: "pcs" });
        if (eStats.totalHvOutgoing > 0) boq.push({ desc: "HV Switchgear Outgoing Bays", qty: eStats.totalHvOutgoing, unit: "pcs" });
        if (eStats.totalMvOutgoing > 0) boq.push({ desc: "MV Point of Connection Bay", qty: eStats.totalMvOutgoing, unit: "pcs" });

        if (cStats.totalLvLength > 0) boq.push({ desc: "LV AC Cabling (AL/XLPE)", qty: Math.round(cStats.totalLvLength), unit: "m" });
        if (cStats.totalMvLength > 0) boq.push({ desc: "MV Collector Cabling (AL/XLPE)", qty: Math.round(cStats.totalMvLength), unit: "m" });
        if (cStats.totalHvLength > 0) boq.push({ desc: "HV Transmission Line/Cable", qty: Math.round(cStats.totalHvLength), unit: "m" });

        return boq;
    },

    showFallbackBoq() {
        const boq = this.gatherBoq();
        if (boq.length === 0) {
            alert("No project quantities found. Please Draw an Area and click Generate first.");
            return;
        }

        // Create a temporary overlay/modal to show the BOQ if not in Excel
        let modal = document.getElementById('boq-fallback-modal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'boq-fallback-modal';
            modal.style = "position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 2rem; border-radius: 12px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1); z-index: 10000; min-width: 400px; max-width: 90vw; border: 1px solid #e2e8f0;";
            document.body.appendChild(modal);
        }

        let html = `<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1.5rem;">
            <h2 style="margin: 0; font-size: 1.25rem; font-weight: 700; color: #1e293b;">Project Bill of Quantities</h2>
            <button onclick="document.getElementById('boq-fallback-modal').style.display='none'" style="background: none; border: none; cursor: pointer; color: #64748b;">✕</button>
        </div>
        <div style="max-height: 400px; overflow-y: auto;">
        <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
            <thead style="background: #f8fafc; text-align: left;">
                <tr>
                    <th style="padding: 0.75rem; border-bottom: 2px solid #e2e8f0;">Description</th>
                    <th style="padding: 0.75rem; border-bottom: 2px solid #e2e8f0;">Qty</th>
                    <th style="padding: 0.75rem; border-bottom: 2px solid #e2e8f0;">Unit</th>
                </tr>
            </thead>
            <tbody>`;
        
        boq.forEach(item => {
            html += `<tr style="border-bottom: 1px solid #f1f5f9;">
                <td style="padding: 0.75rem; color: #334155;">${item.desc}</td>
                <td style="padding: 0.75rem; font-weight: 600; color: #0f172a;">${item.qty.toLocaleString()}</td>
                <td style="padding: 0.75rem; color: #64748b;">${item.unit}</td>
            </tr>`;
        });

        html += `</tbody></table></div>
        <div style="margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid #e2e8f0; font-size: 0.8rem; color: #64748b;">
            <p>Note: Run this application as an Excel Add-in to export directly to cells.</p>
        </div>`;
        
        modal.innerHTML = html;
        modal.style.display = 'block';
    }
};

window.ExcelEngine = ExcelEngine;
window.addEventListener('DOMContentLoaded', () => { ExcelEngine.init(); });
