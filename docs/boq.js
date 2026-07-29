// boq.js
// Handles the dynamic Bill of Quantities (BoQ) generation and UI

const BoQEngine = {
    data: [
        {
            id: 'pv-modules',
            item: 'PV Modules',
            description: 'Solar PV Modules',
            unit: 'kWp',
            qty: 0,
            supplyPrice: 0,
            installPrice: 0,
            subItems: [
                { name: 'PV Modules Supply', unit: 'kWp', price: 120, qty: 0 },
                { name: 'PV Modules Installation', unit: 'kWp', price: 100, qty: 0 }
            ]
        },
        {
            id: 'inverters',
            item: 'Inverters',
            description: 'Solar Inverters',
            unit: 'kW',
            qty: 0,
            supplyPrice: 0,
            installPrice: 0,
            subItems: [
                { name: 'Inverter Supply', unit: 'kW', price: 0, qty: 0 },
                { name: 'Inverter Installation', unit: 'kW', price: 0, qty: 0 }
            ]
        },
        { id: 'mounting', item: 'Mounting Structure', description: '', unit: 'kWp', qty: 0, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'foundations', item: 'Mounting Structure Foundations', description: '', unit: 'EA', qty: 0, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'solar-dc-cables', item: 'Solar DC Cables', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'dc-cables', item: 'DC Cables', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'lv-cables', item: 'LV Cables', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'mv-cables', item: 'MV Cables', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'hv-cables', item: 'HV Cables/Lines', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'grounding', item: 'Grounding', description: '', unit: 'm', qty: 0, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'dc-combiners', item: 'DC Combiners', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'lv-panels', item: 'LV Panels', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { 
            id: 'lv-mv-substation', 
            item: 'LV/MV Substation', 
            description: '', 
            unit: 'LS', 
            qty: 1, 
            supplyPrice: 0, 
            installPrice: 0, 
            subItems: [
                { name: 'Transformer Bay', unit: 'EA', price: 0, qty: 0 },
                { name: 'Line Bay', unit: 'EA', price: 0, qty: 0 },
                { name: 'Busbar System', unit: 'LS', price: 0, qty: 1 }
            ] 
        },
        { 
            id: 'mv-hv-substation', 
            item: 'MV/HV Substation', 
            description: '', 
            unit: 'LS', 
            qty: 1, 
            supplyPrice: 0, 
            installPrice: 0, 
            subItems: [
                { name: 'Power Transformer', unit: 'EA', price: 0, qty: 0 },
                { name: 'MV Switchgear', unit: 'LS', price: 0, qty: 1 },
                { name: 'HV Switchyard', unit: 'LS', price: 0, qty: 1 }
            ] 
        },
        { id: 'delivery-station', item: 'Delivery Station', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'poc', item: 'POC', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'scada', item: 'Monitoring & SCADA', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'substation-building', item: 'Substation Building', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'control-building', item: 'Control Building', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'inv-foundations', item: 'Inverters Stations Fundations', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'fencing', item: 'Fencing', description: '', unit: 'm', qty: 0, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'land-grading', item: 'Land Grading', description: '', unit: 'm2', qty: 0, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'cut-fill', item: 'Cut/Fill', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'flood', item: 'Flood Management', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] },
        { id: 'pm', item: 'Project Management', description: '', unit: 'LS', qty: 1, supplyPrice: 0, installPrice: 0, subItems: [] }
    ],

    systemKwp: 0,
    foundationBreakdown: [
        { name: 'Piles Option 1 - Driving/Ramming into the ground', percent: 100, cost: 20 },
        { name: 'Piles Option 2 - Pre-drilling and Driving', percent: 0, cost: 35 },
        { name: 'Piles Option 3 - Screwing into the ground', percent: 0, cost: 50 },
        { name: 'Piles Option 4 - Pre-drilling + Screwing', percent: 0, cost: 70 },
        { name: 'Piles Option 5 - Pre-Drill + Concrete', percent: 0, cost: 100 },
        { name: 'Piles Option 6 - Concrete Slabs', percent: 0, cost: 150 },
    ],
    mountingBreakdown: [
        { name: 'Ground Single axis Tracking', percent: 100, supplyCost: 75, installCost: 25 },
        { name: 'Ground Fixed', percent: 0, supplyCost: 45, installCost: 15 },
        { name: 'Rooftop Canopy', percent: 0, supplyCost: 100, installCost: 30 },
        { name: 'Rooftop Flushed/Metalic Roof', percent: 0, supplyCost: 20, installCost: 15 },
        { name: 'Rooftop Triangles Single Side', percent: 0, supplyCost: 25, installCost: 15 },
        { name: 'Rooftop Triangles Two Side', percent: 0, supplyCost: 30, installCost: 15 },
        { name: 'Car Parking', percent: 0, supplyCost: 120, installCost: 40 },
        { name: 'Free Drawing', percent: 0, supplyCost: 0, installCost: 0 },
    ],

    init() {
        const btnBoq = document.getElementById('btn-view-boq');
        if (btnBoq) {
            btnBoq.addEventListener('click', () => {
                this.calculateQuantities();
                this.renderMainTable();
                document.getElementById('boq-modal').classList.remove('hidden');
                if (window.lucide) lucide.createIcons();
            });
        }

        const btnExport = document.getElementById('btn-export-boq');
        if (btnExport) {
            btnExport.addEventListener('click', () => {
                this.exportToExcel();
            });
        }
    },

    getInverterSupplyPrice(pNom) {
        if (pNom <= 50) return 80;
        if (pNom <= 100) return 70;
        if (pNom <= 200) return 60;
        if (pNom <= 300) return 50;
        if (pNom <= 400) return 40;
        return 30;
    },

    calculateQuantities() {
        // Get Descriptions from UI
        const pvSelect = document.getElementById('pv-module-select');
        const pvName = pvSelect && pvSelect.selectedIndex >= 0 ? pvSelect.options[pvSelect.selectedIndex].text : 'Solar PV Modules';
        
        const invSelect = document.getElementById('inverter-select');
        const invName = invSelect && invSelect.selectedIndex >= 0 ? invSelect.options[invSelect.selectedIndex].text : 'Solar Inverters';
        
        const mountSelect = document.getElementById('mounting-type-select');
        const mountName = mountSelect && mountSelect.selectedIndex >= 0 ? mountSelect.options[mountSelect.selectedIndex].text : 'Mounting Structure';

        // 1. Get exact DC capacity from the layout summary tab (in MWp, convert to kWp)
        let totalKwp = 0;
        const layoutDcVal = document.getElementById('stats-total-dc')?.innerText;
        if (layoutDcVal) {
            totalKwp = parseFloat(layoutDcVal.replace(/[^\d.-]/g, '')) * 1000;
        } else {
            // Fallback
            totalKwp = (parseFloat(document.getElementById('sys-dc-cap')?.value) || 0) * 1000;
        }

        // 2. Get exact AC capacity from the layout summary tab (in MW, convert to kW)
        let totalInvKw = 0;
        const layoutAcVal = document.getElementById('stats-total-ac')?.innerText;
        if (layoutAcVal) {
            totalInvKw = parseFloat(layoutAcVal.replace(/[^\d.-]/g, '')) * 1000;
        } else {
            // Fallback
            totalInvKw = (parseFloat(document.getElementById('sys-ac-cap')?.value) || 0) * 1000;
        }
        
        this.systemKwp = totalKwp || 1; // avoid division by zero

        const pvItem = this.data.find(d => d.id === 'pv-modules');
        if (pvItem) {
            pvItem.description = pvName;
            pvItem.qty = totalKwp;
            // Set defaults if 0
            if (pvItem.subItems[0]) {
                pvItem.subItems[0].qty = totalKwp;
                if (pvItem.subItems[0].price === 0) pvItem.subItems[0].price = 120;
            }
            if (pvItem.subItems[1]) {
                pvItem.subItems[1].qty = totalKwp;
                if (pvItem.subItems[1].price === 0) pvItem.subItems[1].price = 100;
            }
        }

        const invItem = this.data.find(d => d.id === 'inverters');
        if (invItem) {
            invItem.description = invName;
            invItem.qty = totalInvKw;
            const invP = parseFloat(document.getElementById('inv-pnom')?.value || 125);
            const defaultSupply = this.getInverterSupplyPrice(invP);
            
            if (invItem.subItems[0]) {
                invItem.subItems[0].qty = totalInvKw;
                if (invItem.subItems[0].price === 0) invItem.subItems[0].price = defaultSupply;
            }
            if (invItem.subItems[1]) {
                invItem.subItems[1].qty = totalInvKw;
                if (invItem.subItems[1].price === 0) invItem.subItems[1].price = 5;
            }
        }

        const mountingItem = this.data.find(d => d.id === 'mounting');
        if (mountingItem) {
            mountingItem.description = mountName;
            if (!mountingItem.manualQty) {
                mountingItem.qty = totalKwp;
            }
            let blendedSupplyCost = 0;
            let blendedInstallCost = 0;
            this.mountingBreakdown.forEach(opt => {
                blendedSupplyCost += (opt.percent / 100) * opt.supplyCost;
                blendedInstallCost += (opt.percent / 100) * opt.installCost;
            });
            mountingItem.supplyPrice = blendedSupplyCost;
            mountingItem.installPrice = blendedInstallCost;
        }

        const foundationsItem = this.data.find(d => d.id === 'foundations');
        if (foundationsItem) {
            if (!foundationsItem.manualQty) {
                foundationsItem.qty = Math.ceil(totalKwp / 10);
            }
            let blendedInstallCost = 0;
            this.foundationBreakdown.forEach(opt => {
                blendedInstallCost += (opt.percent / 100) * opt.cost;
            });
            foundationsItem.installPrice = blendedInstallCost;
        }

        // Auto calculate subtotals from subItems if applicable
        this.data.forEach(item => {
            if (item.subItems && item.subItems.length > 0) {
                let sPrice = 0;
                let iPrice = 0;
                item.subItems.forEach(sub => {
                    if (sub.name.toLowerCase().includes('supply')) sPrice += sub.price * sub.qty;
                    else if (sub.name.toLowerCase().includes('install')) iPrice += sub.price * sub.qty;
                    else sPrice += sub.price * sub.qty; // default to supply
                });
                if (item.qty > 0) {
                    item.supplyPrice = sPrice / item.qty;
                    item.installPrice = iPrice / item.qty;
                }
            }
        });
    },

    renderMainTable() {
        const tbody = document.getElementById('boq-main-tbody');
        const tfoot = document.getElementById('boq-main-tfoot');
        const topTotal = document.getElementById('boq-top-total-row');
        if (!tbody || !tfoot) return;

        tbody.innerHTML = '';
        
        let grandTotalSupply = 0;
        let grandTotalInstall = 0;
        
        // First pass to get grand totals for percentage calculation
        this.data.forEach(item => {
            const supplyTotal = item.supplyPrice * item.qty;
            const installTotal = item.installPrice * item.qty;
            grandTotalSupply += supplyTotal;
            grandTotalInstall += installTotal;
        });

        const grandTotal = grandTotalSupply + grandTotalInstall;

        this.data.forEach(item => {
            const tr = document.createElement('tr');
            tr.style.borderBottom = '1px solid #f1f5f9';
            tr.style.cursor = 'pointer';
            tr.className = 'boq-row';
            
            // Hover effect
            tr.onmouseover = () => tr.style.backgroundColor = '#f8fafc';
            tr.onmouseout = () => tr.style.backgroundColor = 'transparent';
            
            tr.onclick = () => this.showSubBoq(item);

            const supplyTotal = item.supplyPrice * item.qty;
            const installTotal = item.installPrice * item.qty;
            const total = supplyTotal + installTotal;
            const percentage = grandTotal > 0 ? (total / grandTotal) * 100 : 0;
            const perKwp = this.systemKwp > 0 ? total / this.systemKwp : 0;

            tr.innerHTML = `
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; color: #0ea5e9; font-weight: 500;">${item.item}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9;">${item.description}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9;">${item.qty > 0 ? parseFloat(item.qty.toFixed(2)) : 0}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9;">${item.unit}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right;">${this.formatCurrency(item.supplyPrice)}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right;">${this.formatCurrency(supplyTotal)}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right;">${this.formatCurrency(item.installPrice)}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right;">${this.formatCurrency(installTotal)}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right; font-weight: 600;">${this.formatCurrency(total)}</td>
                <td style="padding: 0.5rem 1rem; border-right: 1px solid #f1f5f9; text-align: right;">${percentage.toFixed(2)}%</td>
                <td style="padding: 0.5rem 1rem; text-align: right;">${this.formatCurrency(perKwp)}</td>
            `;
            tbody.appendChild(tr);
        });

        // Add Totals Footer
        const grandPerKwp = this.systemKwp > 0 ? grandTotal / this.systemKwp : 0;
        const totalHtml = `
                <td colspan="4" style="padding: 0.75rem 1rem; text-align: right; border-right: 1px solid #cbd5e1;">Grand Total</td>
                <td style="padding: 0.75rem 1rem; border-right: 1px solid #cbd5e1;"></td>
                <td style="padding: 0.75rem 1rem; text-align: right; border-right: 1px solid #cbd5e1;">${this.formatCurrency(grandTotalSupply)}</td>
                <td style="padding: 0.75rem 1rem; border-right: 1px solid #cbd5e1;"></td>
                <td style="padding: 0.75rem 1rem; text-align: right; border-right: 1px solid #cbd5e1;">${this.formatCurrency(grandTotalInstall)}</td>
                <td style="padding: 0.75rem 1rem; text-align: right; border-right: 1px solid #cbd5e1; color: #286944;">${this.formatCurrency(grandTotal)}</td>
                <td style="padding: 0.75rem 1rem; text-align: right; border-right: 1px solid #cbd5e1;">100.00%</td>
                <td style="padding: 0.75rem 1rem; text-align: right;">${this.formatCurrency(grandPerKwp)}</td>
        `;

        tfoot.innerHTML = `<tr>${totalHtml}</tr>`;

        if (topTotal) {
            topTotal.innerHTML = totalHtml;
            topTotal.style.display = 'table-row';
        }
    },

    showSubBoq(item) {
        if (item.id === 'foundations') {
            this.openFoundationsModal(item);
            return;
        }
        if (item.id === 'mounting') {
            this.openMountingModal(item);
            return;
        }
        
        document.getElementById('sub-boq-title').innerText = item.item + ' - Detailed BoQ';
        const tbody = document.getElementById('sub-boq-tbody');
        tbody.innerHTML = '';

        let overallTotal = 0;

        if (!item.subItems || item.subItems.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="padding: 1rem; text-align: center; color: #64748b;">No sub-items configured for this category.</td></tr>';
        } else {
            item.subItems.forEach((sub, idx) => {
                const tr = document.createElement('tr');
                tr.style.borderBottom = '1px solid #f1f5f9';
                
                const total = sub.price * sub.qty;
                overallTotal += total;
                
                tr.innerHTML = `
                    <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">${sub.name}</td>
                    <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                        <input type="number" value="${sub.qty}" min="0" step="any" class="modern-input" style="width: 70px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;" 
                            onchange="BoQEngine.updateSubItem('${item.id}', ${idx}, 'qty', this.value)">
                    </td>
                    <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">${sub.unit}</td>
                    <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                        <span style="font-size: 0.8rem; color: #64748b;">$</span><input type="number" value="${sub.price}" min="0" step="any" class="modern-input" style="width: 80px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;"
                            onchange="BoQEngine.updateSubItem('${item.id}', ${idx}, 'price', this.value)">
                    </td>
                    <td style="padding: 0.35rem 0.6rem; font-weight: 600; font-size: 0.85rem;" id="sub-total-${item.id}-${idx}">${this.formatCurrency(total)}</td>
                `;
                tbody.appendChild(tr);
            });
        }

        const topTotal = document.getElementById('sub-boq-top-total-row');
        if (topTotal) {
            topTotal.innerHTML = `
                <td colspan="4" style="padding: 0.4rem 0.6rem; text-align: right; border-right: 1px solid #cbd5e1; font-size: 0.85rem;">Sub-Component Total</td>
                <td style="padding: 0.4rem 0.6rem; color: #286944; font-size: 0.85rem;" id="sub-boq-grand-total">${this.formatCurrency(overallTotal)}</td>
            `;
            topTotal.style.display = item.subItems && item.subItems.length > 0 ? 'table-row' : 'none';
        }

        document.getElementById('sub-boq-modal').classList.remove('hidden');
    },

    updateSubItem(itemId, subIdx, field, value) {
        const item = this.data.find(d => d.id === itemId);
        if (item && item.subItems[subIdx]) {
            item.subItems[subIdx][field] = parseFloat(value) || 0;
            
            // Recalculate parent
            let sPrice = 0;
            let iPrice = 0;
            item.subItems.forEach(sub => {
                const total = sub.price * sub.qty;
                if (sub.name.toLowerCase().includes('supply')) sPrice += total;
                else if (sub.name.toLowerCase().includes('install')) iPrice += total;
                else sPrice += total; // default to supply
            });
            
            if (item.qty > 0) {
                item.supplyPrice = sPrice / item.qty;
                item.installPrice = iPrice / item.qty;
            } else {
                // If qty is 0, we can't derive unit price easily, so just store totals somehow, or force qty=1
                item.qty = 1;
                item.supplyPrice = sPrice;
                item.installPrice = iPrice;
            }

            // Update sub item total display
            const subTotal = item.subItems[subIdx].price * item.subItems[subIdx].qty;
            const el = document.getElementById(`sub-total-${itemId}-${subIdx}`);
            if (el) el.innerText = this.formatCurrency(subTotal);

            // Update sub-boq grand total if visible
            const gtEl = document.getElementById('sub-boq-grand-total');
            if (gtEl) {
                let overallTotal = 0;
                item.subItems.forEach(s => overallTotal += (s.price * s.qty));
                gtEl.innerText = this.formatCurrency(overallTotal);
            }

            // Re-render main table in background
            this.renderMainTable();
        }
    },

    openFoundationsModal(item) {
        this.currentFoundationsItem = item;
        const tbody = document.getElementById('foundations-tbody');
        if (!tbody) return;
        tbody.innerHTML = '';
        
        const qtyInput = document.getElementById('foundations-qty-input');
        if (qtyInput) qtyInput.value = item.qty;

        this.foundationBreakdown.forEach((opt, idx) => {
            const tr = document.createElement('tr');
            tr.style.borderBottom = '1px solid #f1f5f9';
            tr.innerHTML = `
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">${opt.name}</td>
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                    <input type="number" value="${opt.percent}" min="0" max="100" step="1" class="modern-input" style="width: 70px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;" 
                        onchange="BoQEngine.updateFoundation(${idx}, 'percent', this.value)">
                </td>
                <td id="foundation-qty-${idx}" style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9; text-align: right;">0</td>
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                    <span style="font-size: 0.8rem; color: #64748b;">$</span><input type="number" value="${opt.cost}" min="0" step="any" class="modern-input" style="width: 80px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;"
                        onchange="BoQEngine.updateFoundation(${idx}, 'cost', this.value)">
                </td>
                <td id="foundation-total-${idx}" style="padding: 0.35rem 0.6rem; text-align: right; font-weight: 600;">$0.00</td>
            `;
            tbody.appendChild(tr);
        });
        
        this.updateFoundationsTotal();
        document.getElementById('foundations-modal').classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    },

    updateFoundation(idx, field, value) {
        this.foundationBreakdown[idx][field] = parseFloat(value) || 0;
        this.updateFoundationsTotal();
    },

    updateFoundationsQty(val) {
        if (this.currentFoundationsItem) {
            this.currentFoundationsItem.qty = parseInt(val) || 0;
            this.currentFoundationsItem.manualQty = true;
            this.updateFoundationsTotal();
        }
    },

    updateFoundationsTotal() {
        let totalPercent = 0;
        let totalQty = 0;
        let totalPrice = 0;
        const totalBases = this.currentFoundationsItem ? this.currentFoundationsItem.qty : 0;
        
        this.foundationBreakdown.forEach((opt, idx) => { 
            totalPercent += opt.percent; 
            const qty = Math.round(totalBases * (opt.percent / 100));
            totalQty += qty;
            const price = qty * opt.cost;
            totalPrice += price;
            
            const qtyCell = document.getElementById(`foundation-qty-${idx}`);
            if (qtyCell) qtyCell.textContent = qty.toLocaleString();
            
            const totalCell = document.getElementById(`foundation-total-${idx}`);
            if (totalCell) totalCell.textContent = this.formatCurrency(price);
        });
        
        const spanP = document.getElementById('foundations-total-percent');
        if (spanP) {
            spanP.textContent = `${totalPercent}%`;
            spanP.style.color = totalPercent === 100 ? '#10b981' : '#ef4444';
        }
        
        const spanQ = document.getElementById('foundations-total-qty');
        if (spanQ) spanQ.textContent = totalQty.toLocaleString();
        
        const spanT = document.getElementById('foundations-total-price');
        if (spanT) spanT.textContent = this.formatCurrency(totalPrice);
    },

    saveFoundationsBreakdown() {
        let blendedInstallCost = 0;
        let totalPercent = 0;
        this.foundationBreakdown.forEach(opt => {
            totalPercent += opt.percent;
            blendedInstallCost += (opt.percent / 100) * opt.cost;
        });
        
        if (totalPercent !== 100) {
            alert("Percentages must sum up to exactly 100%.");
            return;
        }

        if (this.currentFoundationsItem) {
            this.currentFoundationsItem.installPrice = blendedInstallCost;
        }
        
        document.getElementById('foundations-modal').classList.add('hidden');
        this.renderMainTable();
    },

    openMountingModal(item) {
        this.currentMountingItem = item;
        const tbody = document.getElementById('mounting-tbody');
        if (!tbody) return;
        tbody.innerHTML = '';
        
        const qtyInput = document.getElementById('mounting-qty-input');
        if (qtyInput) qtyInput.value = item.qty;

        this.mountingBreakdown.forEach((opt, idx) => {
            const tr = document.createElement('tr');
            tr.style.borderBottom = '1px solid #f1f5f9';
            tr.innerHTML = `
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">${opt.name}</td>
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                    <input type="number" value="${opt.percent}" min="0" max="100" step="1" class="modern-input" style="width: 70px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;" 
                        onchange="BoQEngine.updateMounting(${idx}, 'percent', this.value)">
                </td>
                <td id="mounting-qty-${idx}" style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9; text-align: right;">0</td>
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                    <span style="font-size: 0.8rem; color: #64748b;">$</span><input type="number" value="${opt.supplyCost}" min="0" step="any" class="modern-input" style="width: 80px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;"
                        onchange="BoQEngine.updateMounting(${idx}, 'supplyCost', this.value)">
                </td>
                <td style="padding: 0.35rem 0.6rem; border-right: 1px solid #f1f5f9;">
                    <span style="font-size: 0.8rem; color: #64748b;">$</span><input type="number" value="${opt.installCost}" min="0" step="any" class="modern-input" style="width: 80px; padding: 0.2rem 0.4rem; font-size: 0.8rem; height: 2rem;"
                        onchange="BoQEngine.updateMounting(${idx}, 'installCost', this.value)">
                </td>
                <td id="mounting-total-${idx}" style="padding: 0.35rem 0.6rem; text-align: right; font-weight: 600;">$0.00</td>
            `;
            tbody.appendChild(tr);
        });
        
        this.updateMountingTotal();
        document.getElementById('mounting-modal').classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    },

    updateMounting(idx, field, value) {
        this.mountingBreakdown[idx][field] = parseFloat(value) || 0;
        this.updateMountingTotal();
    },

    updateMountingQty(val) {
        if (this.currentMountingItem) {
            this.currentMountingItem.qty = parseFloat(val) || 0;
            this.currentMountingItem.manualQty = true;
            this.updateMountingTotal();
        }
    },

    updateMountingTotal() {
        let totalPercent = 0;
        let totalQty = 0;
        let totalPrice = 0;
        const totalCap = this.currentMountingItem ? this.currentMountingItem.qty : 0;
        
        this.mountingBreakdown.forEach((opt, idx) => { 
            totalPercent += opt.percent; 
            const qty = totalCap * (opt.percent / 100);
            totalQty += qty;
            const price = qty * (opt.supplyCost + opt.installCost);
            totalPrice += price;
            
            const qtyCell = document.getElementById(`mounting-qty-${idx}`);
            if (qtyCell) qtyCell.textContent = qty.toLocaleString(undefined, {maximumFractionDigits: 1});
            
            const totalCell = document.getElementById(`mounting-total-${idx}`);
            if (totalCell) totalCell.textContent = this.formatCurrency(price);
        });
        
        const spanP = document.getElementById('mounting-total-percent');
        if (spanP) {
            spanP.textContent = `${totalPercent}%`;
            spanP.style.color = totalPercent === 100 ? '#10b981' : '#ef4444';
        }
        
        const spanQ = document.getElementById('mounting-total-qty');
        if (spanQ) spanQ.textContent = totalQty.toLocaleString(undefined, {maximumFractionDigits: 1});
        
        const spanT = document.getElementById('mounting-total-price');
        if (spanT) spanT.textContent = this.formatCurrency(totalPrice);
    },

    saveMountingBreakdown() {
        let blendedSupplyCost = 0;
        let blendedInstallCost = 0;
        let totalPercent = 0;
        this.mountingBreakdown.forEach(opt => {
            totalPercent += opt.percent;
            blendedSupplyCost += (opt.percent / 100) * opt.supplyCost;
            blendedInstallCost += (opt.percent / 100) * opt.installCost;
        });
        
        if (totalPercent !== 100) {
            alert("Percentages must sum up to exactly 100%.");
            return;
        }

        if (this.currentMountingItem) {
            this.currentMountingItem.supplyPrice = blendedSupplyCost;
            this.currentMountingItem.installPrice = blendedInstallCost;
        }
        
        document.getElementById('mounting-modal').classList.add('hidden');
        this.renderMainTable();
    },

    formatCurrency(val) {
        if (!val) return '$0.00';
        return '$' + val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    },

    exportToExcel() {
        // Simple CSV export for now
        let csvContent = "data:text/csv;charset=utf-8,";
        csvContent += "Item,Description,Quantity,Unit,Supply Unit Price,Supply Total,Install Unit Price,Install Total,Total Supply & Install,Percentage,$/kWp\n";

        let grandTotalSupply = 0;
        let grandTotalInstall = 0;
        this.data.forEach(item => {
            const supplyTotal = item.supplyPrice * item.qty;
            const installTotal = item.installPrice * item.qty;
            grandTotalSupply += supplyTotal;
            grandTotalInstall += installTotal;
        });
        const grandTotal = grandTotalSupply + grandTotalInstall;

        this.data.forEach(item => {
            const supplyTotal = item.supplyPrice * item.qty;
            const installTotal = item.installPrice * item.qty;
            const total = supplyTotal + installTotal;
            const percentage = grandTotal > 0 ? (total / grandTotal) * 100 : 0;
            const perKwp = this.systemKwp > 0 ? total / this.systemKwp : 0;

            const row = [
                `"${item.item}"`,
                `"${item.description}"`,
                item.qty > 0 ? parseFloat(item.qty.toFixed(2)) : 0,
                `"${item.unit}"`,
                item.supplyPrice.toFixed(2),
                supplyTotal.toFixed(2),
                item.installPrice.toFixed(2),
                installTotal.toFixed(2),
                total.toFixed(2),
                percentage.toFixed(2) + '%',
                perKwp.toFixed(2)
            ].join(",");
            csvContent += row + "\n";
        });

        // Add Grand Total
        const grandPerKwp = this.systemKwp > 0 ? grandTotal / this.systemKwp : 0;
        csvContent += `Grand Total,,,,,${grandTotalSupply.toFixed(2)},,${grandTotalInstall.toFixed(2)},${grandTotal.toFixed(2)},100%,${grandPerKwp.toFixed(2)}\n`;

        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "BoQ_Export.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
};

// Initialize
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => BoQEngine.init());
} else {
    BoQEngine.init();
}
