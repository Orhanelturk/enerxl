/**
 * CablesEngine - Intelligent Electrical Routing & Optimization
 */

const CablesEngine = {
    cablesDB: {
        dc: [
            { size: 4, type: "Cu/XLPE", R: 5.09, rating: 55, vd_factor: 2.5 },
            { size: 6, type: "Cu/XLPE", R: 3.39, rating: 70, vd_factor: 1.6 },
            { size: 10, type: "Cu/XLPE", R: 1.95, rating: 98, vd_factor: 1.0 },
            { size: 16, type: "Cu/XLPE", R: 1.21, rating: 132, vd_factor: 0.6 },
            { size: 25, type: "Cu/XLPE", R: 0.78, rating: 175, vd_factor: 0.4 },
            { size: 35, type: "Cu/XLPE", R: 0.55, rating: 215, vd_factor: 0.3 },
            { size: 50, type: "Cu/XLPE", R: 0.39, rating: 260, vd_factor: 0.2 },
            { size: 70, type: "Cu/XLPE", R: 0.27, rating: 330, vd_factor: 0.15 },
            { size: 95, type: "Cu/XLPE", R: 0.20, rating: 400, vd_factor: 0.1 },
            { size: 120, type: "Cu/XLPE", R: 0.15, rating: 460, vd_factor: 0.08 },
            { size: 150, type: "Cu/XLPE", R: 0.12, rating: 520, vd_factor: 0.06 },
            { size: 185, type: "Cu/XLPE", R: 0.10, rating: 600, vd_factor: 0.05 },
            { size: 240, type: "Cu/XLPE", R: 0.08, rating: 700, vd_factor: 0.04 },
            { size: 300, type: "Cu/XLPE", R: 0.06, rating: 800, vd_factor: 0.03 }
        ],
        lvac_al: [
            { size: 70, type: "Al/XLPE", R: 0.443, X: 0.082, rating: 160 },
            { size: 95, type: "Al/XLPE", R: 0.320, X: 0.082, rating: 195 },
            { size: 120, type: "Al/XLPE", R: 0.253, X: 0.080, rating: 225 },
            { size: 150, type: "Al/XLPE", R: 0.206, X: 0.080, rating: 260 },
            { size: 185, type: "Al/XLPE", R: 0.164, X: 0.080, rating: 295 },
            { size: 240, type: "Al/XLPE", R: 0.125, X: 0.079, rating: 345 },
            { size: 300, type: "Al/XLPE", R: 0.100, X: 0.079, rating: 395 },
            { size: 400, type: "Al/XLPE", R: 0.0778, X: 0.078, rating: 460 },
            { size: 500, type: "Al/XLPE", R: 0.0605, X: 0.077, rating: 530 },
            { size: 630, type: "Al/XLPE", R: 0.0469, X: 0.076, rating: 600 }
        ],
        lvac_cu: [
            { size: 50, type: "Cu/XLPE", R: 0.387, X: 0.082, rating: 190 },
            { size: 70, type: "Cu/XLPE", R: 0.268, X: 0.082, rating: 240 },
            { size: 95, type: "Cu/XLPE", R: 0.193, X: 0.082, rating: 295 },
            { size: 120, type: "Cu/XLPE", R: 0.153, X: 0.080, rating: 340 },
            { size: 150, type: "Cu/XLPE", R: 0.124, X: 0.080, rating: 390 },
            { size: 185, type: "Cu/XLPE", R: 0.0991, X: 0.080, rating: 445 },
            { size: 240, type: "Cu/XLPE", R: 0.0754, X: 0.079, rating: 525 },
            { size: 300, type: "Cu/XLPE", R: 0.0601, X: 0.079, rating: 600 },
            { size: 400, type: "Cu/XLPE", R: 0.0470, X: 0.078, rating: 680 },
            { size: 500, type: "Cu/XLPE", R: 0.0366, X: 0.077, rating: 780 }
        ],
        mv_al: [
            { size: 95, type: "Al/XLPE 33kV", R: 0.320, X: 0.124, rating: 220 },
            { size: 150, type: "Al/XLPE 33kV", R: 0.206, X: 0.117, rating: 280 },
            { size: 240, type: "Al/XLPE 33kV", R: 0.125, X: 0.109, rating: 360 },
            { size: 300, type: "Al/XLPE 33kV", R: 0.100, X: 0.105, rating: 410 },
            { size: 400, type: "Al/XLPE 33kV", R: 0.0778, X: 0.100, rating: 470 },
            { size: 500, type: "Al/XLPE 33kV", R: 0.0605, X: 0.097, rating: 530 },
            { size: 630, type: "Al/XLPE 33kV", R: 0.0469, X: 0.093, rating: 600 },
            { size: 800, type: "Al/XLPE 33kV", R: 0.0367, X: 0.090, rating: 680 }
        ],
        mv_cu: [
            { size: 70, type: "Cu/XLPE 33kV", R: 0.268, X: 0.124, rating: 250 },
            { size: 95, type: "Cu/XLPE 33kV", R: 0.193, X: 0.120, rating: 310 },
            { size: 120, type: "Cu/XLPE 33kV", R: 0.153, X: 0.115, rating: 350 },
            { size: 150, type: "Cu/XLPE 33kV", R: 0.124, X: 0.110, rating: 400 },
            { size: 185, type: "Cu/XLPE 33kV", R: 0.0991, X: 0.105, rating: 450 },
            { size: 240, type: "Cu/XLPE 33kV", R: 0.0754, X: 0.100, rating: 530 },
            { size: 300, type: "Cu/XLPE 33kV", R: 0.0601, X: 0.096, rating: 610 },
            { size: 400, type: "Cu/XLPE 33kV", R: 0.0470, X: 0.092, rating: 700 },
            { size: 500, type: "Cu/XLPE 33kV", R: 0.0366, X: 0.088, rating: 800 }
        ],
        ohl: [
            // 11kV - 33kV
            { size: 'Dog', type: "ACSR 11kV", R: 0.273, X: 0.36, rating: 300 },
            { size: 'Dog', type: "ACSR 33kV", R: 0.273, X: 0.36, rating: 300 },
            { size: 'Panther', type: "ACSR 33kV", R: 0.136, X: 0.35, rating: 450 },
            // 66kV - 132kV
            { size: 'Wolf', type: "ACSR 66kV", R: 0.180, X: 0.35, rating: 350 },
            { size: 'Panther', type: "ACSR 66kV", R: 0.136, X: 0.35, rating: 450 },
            { size: 'Zebra', type: "ACSR 132kV", R: 0.068, X: 0.34, rating: 650 },
            { size: 'Moose', type: "ACSR 132kV", R: 0.055, X: 0.33, rating: 750 },
            // 220kV
            { size: 'Zebra', type: "ACSR 220kV", R: 0.068, X: 0.34, rating: 650 },
            { size: 'Bersimis', type: "ACSR 220kV", R: 0.042, X: 0.32, rating: 900 },
            { size: '2x Zebra (Bundle)', type: "ACSR 220kV", R: 0.034, X: 0.25, rating: 1300 },
            // 400kV
            { size: '2x Moose (Bundle)', type: "ACSR 400kV", R: 0.027, X: 0.24, rating: 1500 },
            { size: '4x Zebra (Bundle)', type: "ACSR 400kV", R: 0.017, X: 0.20, rating: 2600 },
            { size: '4x Moose (Bundle)', type: "ACSR 400kV", R: 0.013, X: 0.19, rating: 3000 }
        ],
        grounding: [
            { size: 35, type: "Cu Bare", R: 0.524, rating: 150 },
            { size: 70, type: "Cu Bare", R: 0.268, rating: 250 },
            { size: 95, type: "Cu Bare", R: 0.193, rating: 320 },
            { size: 120, type: "Cu Bare", R: 0.153, rating: 380 }
        ]
    },

    selections: {
        dc: null,
        lvac: null,
        mv: null,
        hv: null,
        ohl: null,
        grounding: null
    },

    routingGenerated: false,
    routes: {
        lv: [],
        mv: [],
        hv: []
    },

    cachedHighways: null,

    init() {
        console.log("Cables Engine Initialized");
        const btnOpen = document.getElementById('btn-open-cables-lib');
        if (btnOpen) {
            btnOpen.addEventListener('click', () => {
                document.getElementById('cables-lib-modal').classList.remove('hidden');
                this.renderLibrary('dc');
                if (window.lucide) window.lucide.createIcons();
            });
        }
        const tabs = document.querySelectorAll('[data-cable-tab]');
        tabs.forEach(tab => {
            tab.addEventListener('click', (e) => {
                tabs.forEach(t => t.classList.remove('active'));
                e.target.classList.add('active');
                this.renderLibrary(e.target.dataset.cableTab);
            });
        });
        const btnGen = document.getElementById('btn-generate-routing');
        if (btnGen) {
            btnGen.addEventListener('click', () => {
                this.generateRouting();
            });
        }
    },

    renderLibrary(category) {
        const content = document.getElementById('cables-lib-content');
        if (!content) return;
        const db = this.cablesDB[category];
        if (!db) return;
        let html = `<table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 0.85rem; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <thead style="background: #f1f5f9; color: #475569; border-bottom: 2px solid #e2e8f0;">
                    <tr>
                        <th style="padding: 0.75rem;">Size / Name</th>
                        <th style="padding: 0.75rem;">Type</th>
                        <th style="padding: 0.75rem;">Rating (A)</th>
                        <th style="padding: 0.75rem;">R (Ω/km)</th>
                        <th style="padding: 0.75rem;">Action</th>
                    </tr>
                </thead>
                <tbody>`;
        db.forEach((item, idx) => {
            const isSelected = this.selections[category] === idx;
            html += `<tr style="border-bottom: 1px solid #f1f5f9; ${isSelected ? 'background: #ecfdf5;' : ''}">
                    <td style="padding: 0.75rem; font-weight: 600;">${item.size} ${typeof item.size === 'number' ? 'mm²' : ''}</td>
                    <td style="padding: 0.75rem;">${item.type}</td>
                    <td style="padding: 0.75rem;">${item.rating}</td>
                    <td style="padding: 0.75rem;">${item.R}</td>
                    <td style="padding: 0.75rem;"><button onclick="CablesEngine.selectCable('${category}', ${idx})" class="${isSelected ? 'btn-secondary' : 'btn-primary'}" style="padding: 0.25rem 0.75rem; font-size: 0.7rem;">${isSelected ? 'Selected' : 'Select'}</button></td>
                </tr>`;
        });
        html += `</tbody></table>`;
        content.innerHTML = html;
        this.updateStatus();
    },

    selectCable(category, idx) {
        this.selections[category] = (this.selections[category] === idx) ? null : idx;
        this.renderLibrary(category);
    },

    updateStatus() {
        const status = document.getElementById('cable-lib-status');
        if (!status) return;
        const manualCount = Object.values(this.selections).filter(v => v !== null).length;
        status.innerText = manualCount === 0 ? "Auto-select enabled for all categories." : `${manualCount} manual override(s) active. Engine will auto-select the rest.`;
        status.style.color = manualCount === 0 ? "#10b981" : "#f59e0b";
    },

    selectBestCable(category, reqAmps, voltageKv = null) {
        // Resolve dynamic category if it's a generic request
        let dbKey = category;
        const mvMat = document.getElementById('mv-cable-mat')?.value || 'al';
        const hvMat = document.getElementById('hv-cable-mat')?.value || 'al';
        const mvInst = document.getElementById('mv-install-type')?.value || 'ug';
        const hvInst = document.getElementById('main-tx-install-type')?.value || document.getElementById('hv-install-type')?.value || 'ug';

        if (category === 'mv') {
            dbKey = (mvInst === 'ohl') ? 'ohl' : `mv_${mvMat}`;
        } else if (category === 'hv') {
            dbKey = (hvInst === 'ohl') ? 'ohl' : `mv_${hvMat}`; // HV often uses same spec cable but higher voltage rating, for simplicity we use mv_xx or ohl
        } else if (category === 'lvac') {
            const lvMat = document.getElementById('lv-cable-mat')?.value || 'al';
            dbKey = `lvac_${lvMat}`;
        }

        if (this.selections[category] !== null) {
            const selected = this.cablesDB[dbKey] ? this.cablesDB[dbKey][this.selections[category]] : null;
            if (selected) return selected;
        }

        let db = this.cablesDB[dbKey];
        if (!db) return this.cablesDB['dc'][0]; // Fallback

        // Enforce voltage class filtering for OHL based on actual system voltage
        if (dbKey === 'ohl' && voltageKv) {
            let filteredDb = db.filter(c => {
                const match = c.type.match(/(\d+)kV/);
                if (match) return parseInt(match[1]) >= voltageKv;
                return true;
            });
            if (filteredDb.length > 0) db = filteredDb;
        }

        for (let i = 0; i < db.length; i++) {
            if (db[i].rating >= reqAmps) return db[i];
        }
        return db[db.length - 1];
    },

    calculateCableStats(category, loadMw, voltageKv, pf, lengthMts = 150, metadata = {}) {
        // Determine install method and material
        const mvInst = document.getElementById('mv-install-type')?.value || 'ug';
        const hvInst = document.getElementById('main-tx-install-type')?.value || document.getElementById('hv-install-type')?.value || 'ug';

        let installMethod = 'direct';
        if (category === 'mv') installMethod = mvInst === 'ohl' ? 'ohl' : 'direct';
        else if (category === 'hv') installMethod = hvInst === 'ohl' ? 'ohl' : 'direct';

        const soilR = parseFloat(document.getElementById('cable-soil-res')?.value) || 1.5;
        const thermR = parseFloat(document.getElementById('cable-therm-res')?.value) || 1.0;

        let rDerate = 1.0;
        if (installMethod === 'direct' && soilR > 1.2) rDerate -= (soilR - 1.2) * 0.05;
        if (installMethod === 'ducts') rDerate -= 0.15;
        if (thermR > 1.2) rDerate -= (thermR - 1.2) * 0.08;
        if (installMethod === 'ohl') rDerate = 1.0; // OHL doesn't have soil de-rating

        rDerate = Math.max(0.5, rDerate);

        const baseAmps = (loadMw * 1000) / (voltageKv * (category === 'dc' ? 1 : Math.sqrt(3)) * pf);
        const designAmps = baseAmps / rDerate;

        let cable = null;
        let vdPce = 0;
        let vd = 0;

        if (category === 'dc') {
            // Size DC cables primarily by VD (1.5% max) and ampacity
            const db = this.cablesDB['dc'];
            for (let i = 0; i < db.length; i++) {
                let c = db[i];
                vd = 2 * baseAmps * (lengthMts / 1000) * c.R;
                vdPce = (vd / (voltageKv * 1000)) * 100;
                if (c.rating >= designAmps && vdPce <= 1.5) {
                    cable = c;
                    break;
                }
            }
            if (!cable) {
                cable = db[db.length - 1];
                vd = 2 * baseAmps * (lengthMts / 1000) * cable.R;
                vdPce = (vd / (voltageKv * 1000)) * 100;
            }
        } else {
            cable = this.selectBestCable(category, designAmps, voltageKv);
            if (!cable) return null;
            const sinPhi = Math.sin(Math.acos(pf));
            vd = Math.sqrt(3) * baseAmps * (lengthMts / 1000) * ((cable.R * pf) + ((cable.X || 0) * sinPhi));
            vdPce = (vd / (voltageKv * 1000)) * 100;
        }

        const loadingPct = (baseAmps / (cable.rating * rDerate)) * 100;

        // Store calculation in a global sheet registry for UI access
        // ONLY push when a real fromTo routing label is provided.
        // SLD display calls have no label and must NOT pollute the sizing sheet.
        const cableType = `${Array.isArray(cable.size) ? cable.size.join('+') : cable.size} ${typeof cable.size === 'number' ? 'mm²' : ''} ${cable.type}`;
        if (!window._cableSizingSheet) window._cableSizingSheet = [];
        if (metadata && metadata.fromTo && metadata.fromTo !== 'General') {
            window._cableSizingSheet.push({
                category: category.toUpperCase(),
                loadKw: (loadMw * 1000).toFixed(1),
                voltageV: (voltageKv * 1000).toFixed(0),
                amps: baseAmps.toFixed(1),
                length: lengthMts.toFixed(0),
                cableType,
                loading: loadingPct.toFixed(1),
                vdPce: vdPce.toFixed(2),
                fromTo: metadata.fromTo
            });
        }

        return {
            cable, current: baseAmps, loading: loadingPct, vdPce,
            summaryLine1: cableType,
            summaryLine2: `L: ${lengthMts.toFixed(0)}m | Load: ${loadingPct.toFixed(1)}% | VD: ${vdPce.toFixed(2)}%`
        };
    },

    getSmartPath(p1, p2, invInfo) {
        if (!this.cachedHighways) return [p1, p2];
        const rg = this.cachedHighways.rg;
        const azimuth = rg ? rg.azimuth : (parseFloat(document.getElementById('mount-azimuth')?.value) || 180);

        if (rg && this.cachedHighways && window.turf) {
            try {
                const siteCenter = this.cachedHighways.siteCenter;
                const rotateAngle = this.cachedHighways.rotateAngle;
                const slope = this.cachedHighways.slope;
                const yS = this.cachedHighways.yS;
                const vPool = this.cachedHighways.vHighways;
                const hPool = this.cachedHighways.hBlockRoads;

                const pt1 = window.turf.point([p1.lng(), p1.lat()]);
                const pt2 = window.turf.point([p2.lng(), p2.lat()]);
                const pt1r = window.turf.transformRotate(pt1, rotateAngle, { pivot: siteCenter });
                const pt2r = window.turf.transformRotate(pt2, rotateAngle, { pivot: siteCenter });
                const g1x = pt1r.geometry.coordinates[0], g1y = pt1r.geometry.coordinates[1];
                const g2x = pt2r.geometry.coordinates[0], g2y = pt2r.geometry.coordinates[1];
                const snap = (v) => Math.round(v * 1000000) / 1000000;
                const getActualX = (bx, gy) => snap(bx + (gy - yS) * slope);
                const getBaseX = (gx, gy) => snap(gx - (gy - yS) * slope);
                const g1x_base = getBaseX(g1x, g1y), g2x_base = getBaseX(g2x, g2y);
                const g1y_s = snap(g1y), g2y_s = snap(g2y);

                const tableHalfW = (rg.tableW * 1 / 111320 / Math.cos(siteCenter[1] * Math.PI / 180)) * 0.5;
                const tableHalfH = (rg.tableH * 1 / 111320) * 0.5;
                const aisleShift = 0.2 / 111320; // 20cm additional safety from table edge

                // Horizontal Pool Selection
                const wideH = this.cachedHighways.wideHRoads || [];
                let bestH;
                const minY = Math.min(g1y_s, g2y_s), maxY = Math.max(g1y_s, g2y_s);

                // If it's a trunk (MV/HV), we prefer the most direct route that avoids crossing modules
                // If the target is the Hub (DS), we try to pick the road closest to the destination's latitude first
                const hPoolActual = (wideH.length > 0) ? wideH : (hPool.length > 0 ? hPool : [g2y_s, g1y_s]);

                // For trunk cables, prioritize a road that is within the Y-span of start and end
                const strictlyBounded = hPoolActual.filter(h => h >= minY && h <= maxY);
                const candidates = strictlyBounded.length > 0 ? strictlyBounded : hPoolActual;

                bestH = candidates.reduce((a, b) => Math.abs(b - g2y_s) < Math.abs(a - g2y_s) ? b : a);

                // Vertical Pool Selection 
                let vPoolActual = (vPool.length > 0) ? vPool : [g1x_base - tableHalfW - aisleShift, g1x_base + tableHalfW + aisleShift];
                const vL = snap(g1x_base - tableHalfW - aisleShift);
                const vR = snap(g1x_base + tableHalfW + aisleShift);

                // Find best vertical highways (bestV = exit corridor, bestV2 = entry corridor)
                const bestV = vPoolActual.reduce((a, b) => Math.abs(b - g1x_base) < Math.abs(a - g1x_base) ? b : a);
                const bestV2 = vPoolActual.reduce((a, b) => Math.abs(b - g2x_base) < Math.abs(a - g2x_base) ? b : a);

                const bMinY = this.cachedHighways.bMinY, bMaxY = this.cachedHighways.bMaxY;
                const pathLayers = [p1];

                const addS = (bx, gy) => {
                    let gx = getActualX(bx, gy);
                    let pt = window.turf.point([gx, gy]);
                    const p = window.turf.transformRotate(pt, -rotateAngle, { pivot: siteCenter });
                    const latlng = new google.maps.LatLng(p.geometry.coordinates[1], p.geometry.coordinates[0]);
                    const last = pathLayers[pathLayers.length - 1];
                    if (!last || google.maps.geometry.spherical.computeDistanceBetween(last, latlng) > 0.1) pathLayers.push(latlng);
                };

                // RESTORE STEP LADDER FOR BLUE LINES (Daisy Chains)
                // This connects inverters in a simple L-shape (horizontal then vertical)
                // which the user finds visually superior for LV/Daisy-chain routing.
                if (invInfo && invInfo.isDaisyChain) {
                    addS(g2x_base, g1y_s); // Move horizontally to target X
                    pathLayers.push(p2);   // Then drop vertically to target Y (p2)
                    return pathLayers;
                }

                // --- INTELLIGENT ROUTING WITHOUT UNNECESSARY LOOPS ---

                // Detection: Are we already safely within ANY defined highway spine?
                const tolerance = (tableHalfH * 0.5) + aisleShift;
                const isStartInAnyH = hPoolActual.some(h => Math.abs(h - g1y_s) < tolerance);
                const isEndInAnyH = hPoolActual.some(h => Math.abs(h - g2y_s) < tolerance);

                // 1. EXIT FROM SOURCE
                let currentY = g1y_s;
                if (isStartInAnyH) {
                    // Start is cleanly inside a horizontal road. Just slide horizontally to the vertical highway!
                    addS(bestV, g1y_s);
                } else {
                    // Must safely push out of the PV table into a calculated road Y
                    const exitAisleY = (bestH < g1y_s) ? (g1y_s - tableHalfH - aisleShift) : (g1y_s + tableHalfH + aisleShift);
                    addS(g1x_base, exitAisleY);
                    addS(bestV, exitAisleY);
                    currentY = exitAisleY;
                }

                // 2. MAIN TRUNK MOVEMENT (Green and Red Cables)
                if (Math.abs(bestV - bestV2) > 0.000001) {
                    // Need to cross over using the optimally selected horizontal trunk (bestH)
                    addS(bestV, bestH);
                    addS(bestV2, bestH);
                }

                // 3. ENTRY TO DESTINATION
                // Determine safely where we approach the destination's horizontal axis
                let enterY = g2y_s;
                if (!isEndInAnyH) {
                    enterY = (bestH < g2y_s) ? (g2y_s + tableHalfH + aisleShift) : (g2y_s - tableHalfH - aisleShift);
                }

                // Ride the destination's vertical highway up/down to the horizontal approach line
                addS(bestV2, enterY);

                if (!isEndInAnyH) {
                    // Slide horizontally from the vertical highway to exactly beneath/above the target
                    addS(g2x_base, enterY);
                }

                pathLayers.push(p2);
                return pathLayers;


            } catch (e) { console.warn('Smart routing failed:', e); }
        }

        // --- Grid-Aware Fallback Routing ---
        // If smart routing fails or target is fully external, draw a mathematically pure orthogonal
        // L-shape that strictly adheres to the table's rotated grid axes, completely bypassing PV module collision checking.
        if (window.turf && this.cachedHighways && this.cachedHighways.siteCenter) {
            try {
                const siteCenter = this.cachedHighways.siteCenter;
                const gridAngle = this.cachedHighways.rotateAngle || 0;

                const pt1 = window.turf.point([p1.lng(), p1.lat()]);
                const pt2 = window.turf.point([p2.lng(), p2.lat()]);

                // Spin earth flat to grid coordinates
                const pt1r = window.turf.transformRotate(pt1, gridAngle, { pivot: siteCenter });
                const pt2r = window.turf.transformRotate(pt2, gridAngle, { pivot: siteCenter });

                // Construct horizontal then vertical L-shape
                const cornerR = window.turf.point([pt2r.geometry.coordinates[0], pt1r.geometry.coordinates[1]]);

                // Spin corner back to earth coordinates
                const cornerEarth = window.turf.transformRotate(cornerR, -gridAngle, { pivot: siteCenter });
                const cornerLatLng = new google.maps.LatLng(cornerEarth.geometry.coordinates[1], cornerEarth.geometry.coordinates[0]);

                return [p1, cornerLatLng, p2];
            } catch (e) { }
        }

        // Final atomic failsafe: Direct diagonal trace.
        return [p1, p2];
    },

    /**
     * Post-processing: If an orthogonal L-shape corner protrudes completely outside the 
     * project's setback boundary, mathematically slice it off and route exactly along the boundary line.
     * This iteration handles complex multi-segment outside excursions reliably.
     */
    _enforceBoundary(path, setbackPoly) {
        if (!window.turf || !setbackPoly || path.length < 2) return path;

        try {
            const boundaryLine = window.turf.polygonToLine(setbackPoly);
            const fixedPath = [];

            // Helper to add point safely without trailing duplicates
            const addPt = (lng, lat) => {
                const p = new google.maps.LatLng(lat, lng);
                if (fixedPath.length === 0 || !fixedPath[fixedPath.length - 1].equals(p)) {
                    fixedPath.push(p);
                }
            };

            let exitPoint = null;

            for (let i = 0; i < path.length - 1; i++) {
                const A = path[i];
                const B = path[i + 1];
                const ptA = window.turf.point([A.lng(), A.lat()]);
                const ptB = window.turf.point([B.lng(), B.lat()]);

                const lineAB = window.turf.lineString([ptA.geometry.coordinates, ptB.geometry.coordinates]);
                const intersections = window.turf.lineIntersect(lineAB, boundaryLine);

                const ints = [];
                if (intersections && intersections.features.length > 0) {
                    intersections.features.forEach(f => {
                        ints.push({
                            coord: f.geometry.coordinates,
                            dist: window.turf.distance(ptA, f)
                        });
                    });
                    ints.sort((a, b) => a.dist - b.dist);
                }

                let currentInside = window.turf.booleanPointInPolygon(ptA, setbackPoly) ||
                    window.turf.distance(ptA, window.turf.nearestPointOnLine(boundaryLine, ptA)) < 0.001;

                if (!exitPoint && currentInside) {
                    addPt(ptA.geometry.coordinates[0], ptA.geometry.coordinates[1]);
                }

                if (ints.length === 0) {
                    if (currentInside) {
                        // Completely inside, add normally if it's the last point
                        if (i === path.length - 2) addPt(ptB.geometry.coordinates[0], ptB.geometry.coordinates[1]);
                    } else {
                        // Entirely outside. Bridge along boundary if it's the final target destination
                        if (i === path.length - 2 && exitPoint) {
                            const endCoord = ptB.geometry.coordinates; // Final destination
                            const snapEnd = window.turf.nearestPointOnLine(boundaryLine, window.turf.point(endCoord));

                            let sliceCoords = [exitPoint, snapEnd.geometry.coordinates];
                            try {
                                const slice = window.turf.lineSlice(window.turf.point(exitPoint), snapEnd, boundaryLine);
                                if (slice && window.turf.length(slice) < window.turf.distance(window.turf.point(exitPoint), snapEnd) * 3) {
                                    sliceCoords = slice.geometry.coordinates;
                                    if (window.turf.distance(window.turf.point(sliceCoords[sliceCoords.length - 1]), window.turf.point(exitPoint)) <
                                        window.turf.distance(window.turf.point(sliceCoords[0]), window.turf.point(exitPoint))) {
                                        sliceCoords.reverse();
                                    }
                                }
                            } catch (e) { }

                            sliceCoords.forEach(c => addPt(c[0], c[1]));
                            addPt(ptB.geometry.coordinates[0], ptB.geometry.coordinates[1]); // Ensure explicit target is hit
                        }
                    }
                } else {
                    ints.forEach(intersect => {
                        const ic = intersect.coord;
                        if (currentInside) {
                            // Exiting the Polygon!
                            addPt(ic[0], ic[1]);
                            exitPoint = ic;
                            currentInside = false;
                        } else {
                            // Re-entering the Polygon!
                            if (exitPoint) {
                                const snapEnd = window.turf.nearestPointOnLine(boundaryLine, window.turf.point(ic));
                                let sliceCoords = [exitPoint, snapEnd.geometry.coordinates];
                                try {
                                    const slice = window.turf.lineSlice(window.turf.point(exitPoint), snapEnd, boundaryLine);
                                    if (slice && window.turf.length(slice) < window.turf.distance(window.turf.point(exitPoint), snapEnd) * 3) {
                                        sliceCoords = slice.geometry.coordinates;
                                        if (window.turf.distance(window.turf.point(sliceCoords[sliceCoords.length - 1]), window.turf.point(exitPoint)) <
                                            window.turf.distance(window.turf.point(sliceCoords[0]), window.turf.point(exitPoint))) {
                                            sliceCoords.reverse();
                                        }
                                    }
                                } catch (e) { }
                                sliceCoords.forEach(c => addPt(c[0], c[1]));
                            }
                            addPt(ic[0], ic[1]);
                            exitPoint = null;
                            currentInside = true;
                        }
                    });

                    if (currentInside && i === path.length - 2) {
                        addPt(ptB.geometry.coordinates[0], ptB.geometry.coordinates[1]);
                    } else if (!currentInside && i === path.length - 2 && exitPoint) {
                        const endCoord = ptB.geometry.coordinates;
                        const snapEnd = window.turf.nearestPointOnLine(boundaryLine, window.turf.point(endCoord));
                        let sliceCoords = [exitPoint, snapEnd.geometry.coordinates];
                        try {
                            const slice = window.turf.lineSlice(window.turf.point(exitPoint), snapEnd, boundaryLine);
                            if (slice && window.turf.length(slice) < window.turf.distance(window.turf.point(exitPoint), snapEnd) * 3) {
                                sliceCoords = slice.geometry.coordinates;
                                if (window.turf.distance(window.turf.point(sliceCoords[sliceCoords.length - 1]), window.turf.point(exitPoint)) <
                                    window.turf.distance(window.turf.point(sliceCoords[0]), window.turf.point(exitPoint))) {
                                    sliceCoords.reverse();
                                }
                            }
                        } catch (e) { }
                        sliceCoords.forEach(c => addPt(c[0], c[1]));
                        addPt(ptB.geometry.coordinates[0], ptB.geometry.coordinates[1]);
                    }
                }
            }
            return fixedPath;
        } catch (e) {
            console.warn('Boundary enforcement calculation failed gracefully. Retreating to original path.', e);
            return path;
        }
    },

    _prepareRoutingCache(rg, tables, areaSetbackObj) {
        if (!rg || !tables || tables.length === 0) { this.cachedHighways = null; return; }
        const siteCenter = rg.siteCenter, rotateAngle = rg.rotateAngle;
        const mToDeg = 1 / 111320, cosLat = Math.cos(siteCenter[1] * Math.PI / 180);
        const slope = rg.dominantSlope || 0, yS = rg.yStart || 0;
        const snap = (v) => Math.round(v * 1000000) / 1000000;
        const getBaseX = (gx, gy) => snap(gx - (gy - yS) * slope);

        const tableHalfW = (rg.tableW * mToDeg / cosLat) * 0.5;
        const tableHalfH = (rg.tableH * mToDeg) * 0.5;

        const tablesByBlock = {};
        tables.forEach(t => {
            if (!tablesByBlock[t.blockId]) tablesByBlock[t.blockId] = [];
            tablesByBlock[t.blockId].push(t);
        });

        const allBaseXs = [...new Set(tables.map(t => getBaseX(t.gridX, t.gridY)))].sort((a, b) => a - b);
        const allYs = [...new Set(tables.map(t => snap(t.gridY)))].sort((a, b) => a - b);

        // Mathematical Pitch Detection to overcome Staggered Block issues
        // We find the standard "step" between tables, and any gap larger than
        // 105% of this step is mathematically proven to be a Block Gap.

        const xGaps = [];
        for (let i = 0; i < allBaseXs.length - 1; i++) xGaps.push(allBaseXs[i + 1] - allBaseXs[i]);
        const yGaps = [];
        for (let i = 0; i < allYs.length - 1; i++) yGaps.push(allYs[i + 1] - allYs[i]);

        // Find the 'standard' minimum aisle gap (median of lowest 20% to ignore anomalies)
        xGaps.sort((a, b) => a - b);
        yGaps.sort((a, b) => a - b);
        const stX = xGaps.length > 0 ? xGaps[Math.floor(xGaps.length * 0.1)] : 0;
        const stY = yGaps.length > 0 ? yGaps[Math.floor(yGaps.length * 0.1)] : 0;

        const wideVRoads = [];
        for (let i = 0; i < allBaseXs.length - 1; i++) {
            const gap = allBaseXs[i + 1] - allBaseXs[i];
            if (gap > stX * 1.05) wideVRoads.push(snap((allBaseXs[i] + allBaseXs[i + 1]) / 2));
        }

        const wideHRoads = [];
        for (let i = 0; i < allYs.length - 1; i++) {
            const gap = allYs[i + 1] - allYs[i];
            if (gap > stY * 1.05) wideHRoads.push(snap((allYs[i] + allYs[i + 1]) / 2));
        }

        const vCandidates = [...wideVRoads];
        allBaseXs.forEach(bx => {
            vCandidates.push(snap(bx - tableHalfW - (1.5 * mToDeg / cosLat)));
            vCandidates.push(snap(bx + tableHalfW + (1.5 * mToDeg / cosLat)));
        });

        const bRoads = [...wideHRoads];
        allYs.forEach(ty => {
            bRoads.push(snap(ty - tableHalfH - (1.5 * mToDeg)));
            bRoads.push(snap(ty + tableHalfH + (1.5 * mToDeg)));
        });

        const setbackVal = parseFloat(document.getElementById('default-setback')?.value) || 10;
        const marginDegX = (rg.tableW * 0.5 + setbackVal) * mToDeg / cosLat;
        const marginDegY = (rg.tableH * 0.5 + setbackVal) * mToDeg;

        let bMinX = null, bMaxX = null, bMinY = null, bMaxY = null;
        if (allBaseXs.length > 0) {
            bMinX = snap(allBaseXs[0] - marginDegX);
            bMaxX = snap(allBaseXs[allBaseXs.length - 1] + marginDegX);
            vCandidates.push(bMinX);
            vCandidates.push(bMaxX);
            // Keep perimeter out of wideVRoads so it prioritizes internal gaps first
        }
        if (allYs.length > 0) {
            bMinY = snap(allYs[0] - marginDegY);
            bMaxY = snap(allYs[allYs.length - 1] + marginDegY);
            bRoads.push(bMinY);
            bRoads.push(bMaxY);
            // Keep perimeter out of wideHRoads so it prioritizes internal gaps first
        }

        const marginX = 1.0 * mToDeg / cosLat;
        const marginY = 1.0 * mToDeg;
        const isVSafe = (bx) => allBaseXs.every(tx => Math.abs(tx - bx) > tableHalfW + marginX);
        const isHSafe = (y) => allYs.every(ty => Math.abs(ty - y) > tableHalfH + marginY);


        const rotatedUsable = window.turf.transformRotate(window.LayoutEngine.usablePoly, rotateAngle, { pivot: siteCenter });
        const boundaryObj = areaSetbackObj || window.LayoutEngine.usablePoly;
        const rotatedSetback = boundaryObj ? window.turf.transformRotate(boundaryObj, rotateAngle, { pivot: siteCenter }) : null;

        this.cachedHighways = {
            siteCenter, rotateAngle, mToDeg, cosLat, slope, yS, rg,
            vHighways: [...new Set(vCandidates.filter(isVSafe))].sort((a, b) => a - b),
            hBlockRoads: [...new Set(bRoads.filter(isHSafe))].sort((a, b) => a - b),
            wideVRoads: wideVRoads.filter(isVSafe),
            wideHRoads: wideHRoads.filter(isHSafe),
            rotatedUsable,
            rotatedSetback,
            bMinX, bMaxX, bMinY, bMaxY
        };



    },

    generateRouting(silent = false) {
        if (!window.LayoutEngine || window.LayoutEngine.blocks.length === 0) {
            if (!silent) alert("No layout features generated yet.");
            return;
        }
        // Always reset the sizing sheet at the start of each run so stale data never accumulates
        window._cableSizingSheet = [];
        this.routingGenerated = true;
        this.routes = { lv: [], mv: [], hv: [] };

        const mapWindow = window.googleMap;
        if (!mapWindow) return;

        const toLatLng = (c) => {
            if (!c) return null;
            if (typeof c.lat === 'function') return c;
            if (Array.isArray(c)) return new google.maps.LatLng(c[1], c[0]);
            return new google.maps.LatLng(c.lat, c.lng);
        };

        let mvNodes = window.LayoutEngine.blocks.map((poly, i) => {
            let center = null;
            if (poly.item?.center) {
                const c = poly.item.center;
                center = Array.isArray(c)
                    ? new google.maps.LatLng(c[1], c[0])
                    : (typeof c.lat === 'function' ? c : new google.maps.LatLng(c.lat, c.lng));
            } else if (poly.getPath) {
                const pts = poly.getPath();
                if (pts.getLength() > 0) center = pts.getAt(0);
            }
            return {
                id: i,
                center,
                blockId: poly.item?.blockId,
                area: poly.area,
                item: poly.item
            };
        }).filter(n => n.center !== null);

        let invNodes = (window.LayoutEngine.inverters || []).map((poly, i) => {
            // item.center is a turf [lng, lat] array; getPath().getAt(0) is a LatLng fallback
            let center = null;
            if (poly.item?.center) {
                const c = poly.item.center;
                center = Array.isArray(c)
                    ? new google.maps.LatLng(c[1], c[0])
                    : (typeof c.lat === 'function' ? c : new google.maps.LatLng(c.lat, c.lng));
            } else if (poly.getPath) {
                const pts = poly.getPath();
                if (pts.getLength() > 0) center = pts.getAt(0);
            }
            return {
                id: poly.item?.id || i,
                center,
                blockId: poly.item?.blockId,
                area: poly.area,
                item: poly.item
            };
        }).filter(n => n.center !== null); // discard any with unresolvable center

        if (this.mapPolylines) this.mapPolylines.forEach(l => l.setMap(null));
        this.mapPolylines = [];

        // Group everything by Area UID for isolated processing
        const workGroups = new Map();
        [...mvNodes, ...invNodes].forEach(n => {
            const uid = n.area?.__uid || 'main';
            if (!workGroups.has(uid)) workGroups.set(uid, { area: n.area, mv: [], inv: [] });
            if (n.blockId !== undefined && n.id !== undefined && !invNodes.includes(n)) workGroups.get(uid).mv.push(n);
            else workGroups.get(uid).inv.push(n);
        });

        // Refine workGroups (mv/inv lists were mixed)
        workGroups.forEach(g => {
            g.mv = mvNodes.filter(m => (m.area?.__uid || 'main') === (g.area?.__uid || 'main'));
            g.inv = invNodes.filter(i => (i.area?.__uid || 'main') === (g.area?.__uid || 'main'));
        });

        const blocksPerLoop = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1;
        let totalMvLength = 0;

        workGroups.forEach((group, areaUid) => {
            const area = group.area;
            const areaMv = group.mv;
            const areaInv = group.inv;

            // Initialize routing cache for THIS area specifically using its safely isolated grid
            const layoutData = window.LayoutEngine.layoutStore.get(area);
            if (!layoutData || !layoutData.routingGrid) return;

            const areaSetbackPoly = layoutData.setbackLinePoly || window.LayoutEngine.setbackLinePoly;

            this._prepareRoutingCache(layoutData.routingGrid, layoutData.tables, areaSetbackPoly);

            // Process LV (Blue): each inverter connects DIRECTLY to its ACP panel (LV Bus)
            // ACP numbering matches the SLD: invPerPanel = ceil(blockInvs / panelQty)
            const panelQty = parseInt(document.getElementById('tr-sum-panel-tr')?.innerText) || 2;

            // Build per-block inverter index map so we know ACP number for each inverter
            const blockInvIndexMap = new Map(); // blockId → sorted list of inv IDs
            areaInv.forEach(inv => {
                const bId = inv.blockId;
                if (!blockInvIndexMap.has(bId)) blockInvIndexMap.set(bId, []);
                blockInvIndexMap.get(bId).push(inv);
            });
            // Sort each block's inverters numerically by their ID
            blockInvIndexMap.forEach((invList) => {
                invList.sort((a, b) => {
                    const na = parseInt(String(a.item?.id || a.id || '').split('.').pop()) || 0;
                    const nb = parseInt(String(b.item?.id || b.id || '').split('.').pop()) || 0;
                    return na - nb;
                });
            });

            areaInv.forEach(inv => {
                const parentMv = areaMv.find(mv => mv.blockId === inv.blockId);
                if (!parentMv) return;

                // Determine which ACP panel this inverter belongs to
                const blockInvList = blockInvIndexMap.get(inv.blockId) || [];
                const invPerPanel = Math.ceil(blockInvList.length / panelQty);
                const globalInvIndex = blockInvList.indexOf(inv);
                const acpNum = invPerPanel > 0 ? Math.floor(globalInvIndex / invPerPanel) + 1 : 1;

                // Route directly from inverter → ACP (MV Station)
                let path = this.getSmartPath(inv.center, parentMv.center, inv);
                if (areaSetbackPoly) {
                    path = this._enforceBoundary(path, areaSetbackPoly);
                }

                const lvCable = new google.maps.Polyline({
                    path, geodesic: true, strokeColor: '#3b82f6',
                    strokeOpacity: 0.6, strokeWeight: 3, map: mapWindow, zIndex: 2000,
                    editable: false
                });
                lvCable._isCablesEngineRoute = true; // sentinel so stats loop skips foreign polylines

                // Build label: "Block X Inv Y ➔ ACP{N} Block X"
                const rawSourceId = String(inv.item?.id || inv.id || '');
                let sBlock = rawSourceId.split('-')[1]?.split('.')[0]
                    || String(inv.item?.blockId || inv.blockId || 'N/A').replace(/Block /ig, '').trim();
                sBlock = String(sBlock).replace(/Block /ig, '').trim();
                const sNum = rawSourceId.split('.')[1] || '1';
                const pBlock = String(parentMv.item?.blockId || parentMv.blockId || 'N/A').replace(/Block /ig, '').trim();

                lvCable._myRoutingData = { fromTo: `Block ${sBlock} Inv ${sNum} ➔ ACP${acpNum} Block ${pBlock}` };

                google.maps.event.addListener(lvCable, 'click', function () {
                    this.setEditable(!this.getEditable());
                });
                this.mapPolylines.push(lvCable);
                this.routes.lv.push({ from: inv.center, to: parentMv.center });
            });

            // Find linked DS for this area
            let hub = null;
            if (area && area.linkedDsId && window.SiteEngine) {
                const ds = window.SiteEngine.overlays.find(o =>
                    (o.__uid === area.linkedDsId || o.overlayId === area.linkedDsId || o.areaName === area.linkedDsId)
                );
                if (ds) {
                    const bounds = new google.maps.LatLngBounds();
                    if (ds.getPath) ds.getPath().forEach(p => bounds.extend(p));
                    else if (ds.getPosition) bounds.extend(ds.getPosition());
                    hub = bounds.getCenter();
                }
            }
            if (!hub && window.SiteEngine) {
                const globalDs = window.SiteEngine.overlays.find(o => o.subType === 'station');
                if (globalDs) {
                    const bounds = new google.maps.LatLngBounds();
                    if (globalDs.getPath) globalDs.getPath().forEach(p => bounds.extend(p));
                    hub = bounds.getCenter();
                }
            }

            // Split into loops
            let exitLatLng = hub;
            if (hub && area && window.turf) {
                try {
                    const areaGeo = window.SiteEngine.getGeoJSON(area);
                    if (areaGeo) {
                        const hubPt = window.turf.point([hub.lng(), hub.lat()]);
                        const isInside = window.turf.booleanPointInPolygon(hubPt, areaGeo);
                        if (!isInside) {
                            const lines = window.turf.polygonToLine(areaGeo);
                            const nearest = window.turf.nearestPointOnLine(lines, hubPt);
                            exitLatLng = new google.maps.LatLng(nearest.geometry.coordinates[1], nearest.geometry.coordinates[0]);
                        }
                    }
                } catch (e) { console.warn("Exit point computation failed", e); }
            }

            for (let i = 0; i < areaMv.length; i += blocksPerLoop) {
                const loop = areaMv.slice(i, i + blocksPerLoop);
                if (hub) {
                    // Sort ASCENDING distance from hub: closest first → farthest last.
                    // This matches the SLD drawing: Hub → MV1(close) → MV2 → ... → MVn(far)
                    // The red ring-close cable then runs from the farthest station back to hub.
                    loop.sort((a, b) =>
                        google.maps.geometry.spherical.computeDistanceBetween(a.center, hub) -
                        google.maps.geometry.spherical.computeDistanceBetween(b.center, hub)
                    );
                }

                for (let j = 0; j < loop.length - 1; j++) {
                    let path = this.getSmartPath(loop[j].center, loop[j + 1].center, null);
                    if (areaSetbackPoly) {
                        path = this._enforceBoundary(path, areaSetbackPoly);
                    }
                    let len = 0; for (let k = 0; k < path.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(path[k], path[k + 1]);
                    totalMvLength += len;

                    const greenCable = new google.maps.Polyline({
                        path, geodesic: true, strokeColor: '#10b981',
                        strokeOpacity: 0.7, strokeWeight: 4, map: mapWindow, zIndex: 2000,
                        editable: false
                    });
                    greenCable._isCablesEngineRoute = true;
                    // j=0 is the segment leaving the hub end (carries highest load);
                    // each subsequent segment carries one fewer station worth of load.
                    greenCable._segmentIndex = j;         // 0-based position in chain (0 = closest to hub)
                    greenCable._loopSize = loop.length;   // total stations in this loop
                    greenCable._myRoutingData = { fromTo: `Station ${loop[j].blockId} ➔ Station ${loop[j + 1].blockId}` };
                    google.maps.event.addListener(greenCable, 'click', function () {
                        this.setEditable(!this.getEditable());
                    });
                    this.mapPolylines.push(greenCable);
                }

                const lastNode = loop[loop.length - 1];
                if (hub && !lastNode.center.equals(hub)) {
                    // First leg: Internal routing out to the single area exit point (strictly bounded)
                    let path1 = this.getSmartPath(lastNode.center, exitLatLng, null);
                    if (areaSetbackPoly) {
                        path1 = this._enforceBoundary(path1, areaSetbackPoly);
                    }
                    let len = 0; for (let k = 0; k < path1.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(path1[k], path1[k + 1]);

                    // Second leg: Exterior routing from exit point to the DS Hub
                    let path2 = [];
                    if (!exitLatLng.equals(hub)) {
                        // Crucial fix: External routes MUST bypass getSmartPath PV table collision solving
                        // otherwise the red line will intensely route backwards into the array hunting for a "safe PV gap".
                        let p2 = [exitLatLng, hub];
                        if (window.turf && this.cachedHighways && this.cachedHighways.siteCenter) {
                            try {
                                const siteCenter = this.cachedHighways.siteCenter;
                                const gridAngle = this.cachedHighways.rotateAngle || 0;
                                const pt1 = window.turf.point([exitLatLng.lng(), exitLatLng.lat()]);
                                const pt2 = window.turf.point([hub.lng(), hub.lat()]);

                                const pt1r = window.turf.transformRotate(pt1, gridAngle, { pivot: siteCenter });
                                const pt2r = window.turf.transformRotate(pt2, gridAngle, { pivot: siteCenter });

                                const cornerR = window.turf.point([pt2r.geometry.coordinates[0], pt1r.geometry.coordinates[1]]);
                                const cornerEarth = window.turf.transformRotate(cornerR, -gridAngle, { pivot: siteCenter });
                                const cornerLatLng = new google.maps.LatLng(cornerEarth.geometry.coordinates[1], cornerEarth.geometry.coordinates[0]);

                                p2 = [exitLatLng, cornerLatLng, hub];
                            } catch (e) { }
                        }

                        path2 = p2.slice(1); // skip the starting point to prevent duplicate coordinates
                        for (let k = 0; k < p2.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(p2[k], p2[k + 1]);
                    }

                    const fullPath = path1.concat(path2);
                    totalMvLength += len;

                    const redCable = new google.maps.Polyline({
                        path: fullPath, geodesic: true, strokeColor: '#ef4444',
                        strokeOpacity: 0.8, strokeWeight: 5, map: mapWindow, zIndex: 2100,
                        editable: false // Off by default to prevent visual vertex clutter
                    });
                    redCable._isCablesEngineRoute = true;
                    redCable._myRoutingData = { fromTo: `Station ${lastNode.blockId} ➔ Hub / Bus` };

                    // Toggle vertex editing handles natively only when clicked
                    google.maps.event.addListener(redCable, 'click', function () {
                        this.setEditable(!this.getEditable());
                    });

                    this.mapPolylines.push(redCable);
                }
            }
        });

        this.stats = {
            totalMvLength: totalMvLength,
            totalLvLength: 0,
            totalDcLength: 0,
            totalHvLength: 0,
            breakdown: []
        };

        // Clear global sizing sheet
        window._cableSizingSheet = [];

        // Detailed breakdowns for BOQ
        const breakdownMap = new Map();
        const addToBreakdown = (cable, length) => {
            const key = `${cable.size}mm² ${cable.type}`;
            if (!breakdownMap.has(key)) breakdownMap.set(key, { desc: key, qty: 0, unit: 'm' });
            breakdownMap.get(key).qty += length;
        };

        // 1. Calculate DC Lengths based on actual layout distance
        let totalDcLength = 0;
        const modPerStr = parseInt(document.getElementById('inv-mods-str')?.value) || 25;
        const pvP = parseFloat(document.getElementById('pv-pnom')?.value) || 550;
        const pvVmpp = parseFloat(document.getElementById('pv-vmpp')?.value) || 41.95; // V per module at MPP
        const pvVoc = parseFloat(document.getElementById('pv-voc')?.value) || 49.8;  // V per module open circuit
        // String operating voltage (Vmpp * modules per string), converted to kV
        const stringVoltageKv = (pvVmpp * modPerStr) / 1000;
        // String open-circuit voltage for insulation check
        const stringVocKv = (pvVoc * modPerStr) / 1000;
        const dcVoltage = stringVoltageKv; // use Vmpp for sizing and VD calc

        if (window.LayoutEngine && window.LayoutEngine.tables && window.LayoutEngine.inverters && window.LayoutEngine.tables.length > 0) {
            const tables = window.LayoutEngine.tables;
            const inverters = window.LayoutEngine.inverters;
            const mx = parseInt(document.getElementById('mount-mod-table-x')?.value || 25);
            const my = parseInt(document.getElementById('mount-mod-table-y')?.value || 2);
            const modsPerTable = mx * my;
            const stringsPerTable = modsPerTable / modPerStr;
            const pvWidth = (parseFloat(document.getElementById('pv-width')?.value) || 1134) / 1000;
            const tableWidthM = mx * pvWidth;
            const invP = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
            const acdc = parseFloat(document.getElementById('inv-acdc-ratio')?.value) || 1.2;
            const invStringCounter = new Map();

            // Pre-calculate exact string capacity for each inverter to perfectly match SLD
            const invCapacities = new Map();
            const blockStats = new Map();

            tables.forEach(t => {
                const bId = t.item?.blockId || t.blockId;
                if (!blockStats.has(bId)) blockStats.set(bId, { tables: 0, invs: [] });
                blockStats.get(bId).tables += 1;
            });

            inverters.forEach(inv => {
                const bId = inv.item?.blockId || inv.blockId;
                if (blockStats.has(bId)) {
                    blockStats.get(bId).invs.push(inv);
                }
            });

            blockStats.forEach((stats, bId) => {
                const totalStringsInBlock = Math.round(stats.tables * stringsPerTable);
                const totalInvs = stats.invs.length;
                if (totalInvs === 0) return;

                stats.invs.sort((a, b) => {
                    const idA = String(a.item?.id || a.id || '');
                    const idB = String(b.item?.id || b.id || '');
                    return idA.localeCompare(idB, undefined, { numeric: true });
                });

                const baseStrings = Math.floor(totalStringsInBlock / totalInvs);
                const remainder = totalStringsInBlock % totalInvs;

                stats.invs.forEach((inv, index) => {
                    const rawInvId = String(inv.item?.id || inv.id || '');
                    const capacity = baseStrings + (index < remainder ? 1 : 0);
                    invCapacities.set(rawInvId, capacity);
                });
            });

            // Group tables by block to ensure string count matches SLD per block
            const tablesByBlock = new Map();
            tables.forEach(t => {
                const bId = t.item?.blockId || t.blockId;
                if (!tablesByBlock.has(bId)) tablesByBlock.set(bId, []);
                tablesByBlock.get(bId).push(t);
            });

            tablesByBlock.forEach((blockTables, bId) => {
                const blockInvs = inverters.filter(i => (i.item?.blockId || i.blockId) === bId);
                const strLoadMw = (pvP * modPerStr) / 1000000;

                // Match SLD total block string count exactly
                const targetBlockStrings = Math.round(blockTables.length * stringsPerTable);
                const baseStrPerTable = Math.floor(targetBlockStrings / blockTables.length);
                const extraStrCount = targetBlockStrings % blockTables.length;

                blockTables.forEach((table, tableIdx) => {
                    const stringsForThisTable = baseStrPerTable + (tableIdx < extraStrCount ? 1 : 0);

                    // Route EACH string individually to enforce capacity limits
                    for (let s = 0; s < stringsForThisTable; s++) {
                        let closestDistKm = 0.065; // default 65m
                        let targetInv = null;

                        if (blockInvs.length > 0) {
                            let minDist = Infinity;
                            let fallbackInv = null;
                            let fallbackDist = Infinity;

                            blockInvs.forEach(inv => {
                                const rawInvId = String(inv.item?.id || inv.id || '');
                                const currentCount = invStringCounter.get(rawInvId) || 0;
                                const invCenter = inv.item?.center || inv.center;
                                const d = turf.distance(turf.point(table.center), turf.point(invCenter), { units: 'kilometers' });

                                // Track absolute closest as fallback
                                if (d < fallbackDist) {
                                    fallbackDist = d;
                                    fallbackInv = inv;
                                }

                                // Check if inverter has capacity exactly matching SLD
                                const capacityLimit = invCapacities.get(rawInvId) || 999;
                                if (currentCount < capacityLimit && d < minDist) {
                                    minDist = d;
                                    targetInv = inv;
                                }
                            });

                            // If all inverters are somehow full, fallback to the absolute closest
                            if (!targetInv) {
                                targetInv = fallbackInv;
                                minDist = fallbackDist;
                            }
                            closestDistKm = minDist;
                        }

                        const actualStringDistM = (closestDistKm * 1000) + (tableWidthM / 2); // dist to inv + avg wiring on table
                        const rawTargetId = String(targetInv?.item?.id || targetInv?.id || '');

                        if (!invStringCounter.has(rawTargetId)) invStringCounter.set(rawTargetId, 0);
                        const strNum = invStringCounter.get(rawTargetId) + 1;
                        invStringCounter.set(rawTargetId, strNum);

                        let tgtBlock = rawTargetId.split('-')[1]?.split('.')[0] || table.item?.blockId || table.blockId || 'N/A';
                        tgtBlock = String(tgtBlock).replace(/Block /ig, '').trim();
                        const tgtNum = rawTargetId.split('.')[1] || '1';

                        const statsDc = this.calculateCableStats('dc', strLoadMw, dcVoltage, 1.0, actualStringDistM, {
                            fromTo: `Block ${tgtBlock} Inverter ${tgtNum} String ${strNum}`
                        });
                        if (statsDc) {
                            totalDcLength += actualStringDistM;
                            addToBreakdown(statsDc.cable, actualStringDistM);
                        }
                    }
                });
            });
            this.stats.totalDcLength = totalDcLength;
        } else {
            // Fallback if no layout exists yet
            const totalSysDcKw = (parseFloat(document.getElementById('sys-dc-cap')?.value) || 0) * 1000;
            let totalStrings = Math.ceil(totalSysDcKw / (pvP * modPerStr / 1000));
            const avgStrLen = 65;
            const statsDc = this.calculateCableStats('dc', (pvP * modPerStr) / 1000000, dcVoltage, 1.0, avgStrLen, { fromTo: 'General DC Strings (Avg Estimate)' });
            this.stats.totalDcLength = totalStrings * avgStrLen;
            if (statsDc) addToBreakdown(statsDc.cable, this.stats.totalDcLength);
        }

        // Calculate precise LV length by summing map polylines and add to breakdown
        let sumLv = 0;
        const invPnomCw = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
        const lvVolts = parseFloat(document.getElementById('lv-v')?.value) || 400;
        this.mapPolylines.forEach(l => {
            if (l._isCablesEngineRoute && l.get('strokeColor') === '#3b82f6') { // LV Blue - only our own cables
                if (!l._myRoutingData) return; // skip any without routing metadata — no General fallback
                const path = l.getPath();
                let len = 0;
                for (let i = 0; i < path.getLength() - 1; i++) {
                    len += google.maps.geometry.spherical.computeDistanceBetween(path.getAt(i), path.getAt(i + 1));
                }
                sumLv += len;
                const stats = this.calculateCableStats('lvac', invPnomCw / 1000, lvVolts / 1000, 0.9, len, l._myRoutingData);
                if (stats) addToBreakdown(stats.cable, len);
            }
        });
        this.stats.totalLvLength = sumLv;

        // Calculate MV breakdown
        const blocksPerLoopVal = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1;
        const trMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
        const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
        const pf = parseFloat(document.getElementById('sys-pf')?.value) || 0.9;

        this.mapPolylines.forEach(l => {
            const color = l.get('strokeColor');
            if (l._isCablesEngineRoute && (color === '#10b981' || color === '#ef4444')) { // MV only our own
                if (!l._myRoutingData) return; // no General fallback
                const path = l.getPath();
                let len = 0;
                for (let i = 0; i < path.getLength() - 1; i++) {
                    len += google.maps.geometry.spherical.computeDistanceBetween(path.getAt(i), path.getAt(i + 1));
                }

                let segLoadMva;
                if (color === '#10b981') {
                    // Green feeder: segment j (0 = closest to hub) carries all stations DOWNSTREAM of it.
                    // Downstream count = loopSize - segmentIndex - 1
                    // (segment 0 carries loopSize-1 stations; last green segment carries 1 station)
                    const segIdx  = l._segmentIndex  ?? 0;
                    const loopSz  = l._loopSize      ?? blocksPerLoopVal;
                    const downstream = Math.max(1, loopSz - segIdx - 1);
                    segLoadMva = downstream * trMva * pf;
                } else {
                    // Red ring-close cable: sized for full loop load (worst-case fault scenario)
                    segLoadMva = blocksPerLoopVal * trMva * pf;
                }

                const statsMv = this.calculateCableStats('mv', segLoadMva, mvV, pf, len, l._myRoutingData);
                if (statsMv) addToBreakdown(statsMv.cable, len);
            }
        });

        // Placeholder for HV breakdown (moved down)

        this.stats.breakdown = Array.from(breakdownMap.values());

        // --- Draw Editable HV Cables (Orange) from DS to POC ---
        if (window.SiteEngine) {
            const SE = window.SiteEngine;
            // Only pull actively rendered geometric nodes! Ghost objects left in memory from incomplete deletions are purged.
            const pocs = SE.overlays.filter(o => o.subType === 'poc' && typeof o.getMap === 'function' && o.getMap() !== null);
            const stations = SE.overlays.filter(o => o.subType === 'station' && typeof o.getMap === 'function' && o.getMap() !== null);

            const validCableSignatures = new Set();

            stations.forEach(ds => {
                let targetPoc = null;
                const parentArea = SE.overlays.find(o => o.linkedDsId === ds.__uid);

                if (parentArea && parentArea.linkedPocId) {
                    targetPoc = pocs.find(p => p.__uid === parentArea.linkedPocId || p.overlayId === parentArea.linkedPocId || p.areaName === parentArea.linkedPocId);
                }
                if (!targetPoc && pocs.length > 0) targetPoc = pocs[0];

                if (targetPoc) {
                    const getCenter = (eq) => {
                        const bounds = new google.maps.LatLngBounds();
                        if (eq.getPath) eq.getPath().forEach(p => bounds.extend(p));
                        else if (eq.getPosition) bounds.extend(eq.getPosition());
                        return bounds.getCenter();
                    };

                    const dsCenter = getCenter(ds);
                    const pocCenter = getCenter(targetPoc);


                    validCableSignatures.add(`${ds.__uid}-${targetPoc.__uid}`);

                    const existingCable = SE.overlays.find(o => o.category === 'hv-cable' && o.startNode === ds.__uid && o.endNode === targetPoc.__uid);

                    if (!existingCable) {
                        const orangeCable = new google.maps.Polyline({
                            path: [dsCenter, pocCenter],
                            geodesic: true,
                            strokeColor: '#f97316',
                            strokeOpacity: 0.9,
                            strokeWeight: 6,
                            map: mapWindow,
                            editable: false,
                            zIndex: 2200
                        });
                        orangeCable._myRoutingData = { fromTo: `Delivery Station ➔ Point of Connection` };

                        google.maps.event.addListener(orangeCable, 'click', function () {
                            this.setEditable(!this.getEditable());
                        });

                        orangeCable.category = 'hv-cable';
                        orangeCable.subType = 'hv-cable';
                        if (SE.generateUid) orangeCable.__uid = SE.generateUid();
                        orangeCable.startNode = ds.__uid;
                        orangeCable.endNode = targetPoc.__uid;
                        orangeCable.areaName = `HV Cable`;

                        const path = orangeCable.getPath();
                        google.maps.event.addListener(path, 'set_at', function (index) {
                            if (index === 0) {
                                const currentDsCenter = getCenter(ds);
                                if (!this.getAt(0).equals(currentDsCenter)) this.setAt(0, currentDsCenter);
                            } else if (index === this.getLength() - 1) {
                                const currentPocCenter = getCenter(targetPoc);
                                if (!this.getAt(this.getLength() - 1).equals(currentPocCenter)) this.setAt(this.getLength() - 1, currentPocCenter);
                            }
                        });

                        SE.overlays.push(orangeCable);
                        SE.attachOverlayListeners(orangeCable);
                    } else {
                        // Update ends securely if DS or POC moved manually
                        const path = existingCable.getPath();
                        // Turn off the listener momentarily if we really want to force it, but dynamic getCenter() in the listener means it gracefully accepts this now.
                        path.setAt(0, dsCenter);
                        path.setAt(path.getLength() - 1, pocCenter);

                        // Enforce off-by-default for legacy lingering cables
                        if (existingCable.getEditable() !== false && existingCable.setEditable) existingCable.setEditable(false);
                    }
                }
            });

            // Post-Generation Garbage Collection: Destroy Stale/Orphaned HV Cables
            // If the user deleted a Delivery Station or manually overrode an area's mapping to target a different POC,
            // we must cleanly wipe the old orange routing trace off the physical canvas and out of engine memory.
            for (let i = SE.overlays.length - 1; i >= 0; i--) {
                const o = SE.overlays[i];
                if (o.category === 'hv-cable') {
                    const sig = `${o.startNode}-${o.endNode}`;
                    if (!validCableSignatures.has(sig)) {
                        if (o.setMap) o.setMap(null);
                        SE.overlays.splice(i, 1);
                    }
                }
            }
        }

        // Add HV cable lengths to stats and Breakdown
        let sumHv = 0;
        if (window.SiteEngine) {
            const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
            const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
            const hvV = parseFloat(document.getElementById('hv-v')?.value) || 132;
            const transmissionV = (pocLevel === 'MV') ? mvV : hvV;
            const totalProjectMw = (parseFloat(document.getElementById('sys-ac-cap')?.value) || 0);

            window.SiteEngine.overlays.forEach(o => {
                if (o.category === 'hv-cable' && o.getPath) {
                    const path = o.getPath();
                    let len = 0;
                    for (let i = 0; i < path.getLength() - 1; i++) {
                        len += google.maps.geometry.spherical.computeDistanceBetween(path.getAt(i), path.getAt(i + 1));
                    }
                    sumHv += len;
                    const statsHv = this.calculateCableStats('hv', totalProjectMw, transmissionV, 0.9, len, o._myRoutingData || { fromTo: 'Main DS ➔ Grid POC' });
                    if (statsHv) addToBreakdown(statsHv.cable, len);
                }
            });
        }
        if (this.stats) this.stats.totalHvLength = sumHv;

        if (window.SldEngine && !document.getElementById('sld-window').classList.contains('hidden')) window.SldEngine.renderSld();
        if (!silent) alert(`Routing generated successfully!\nTotal estimated MV route length: ${(totalMvLength / 1000).toFixed(2)} km`);
    },

    /**
     * Renders the Detailed Cable Sizing Sheet in a UI Modal
     */
    showSizingSheet() {
        if (!window._cableSizingSheet || window._cableSizingSheet.length === 0) {
            alert("No cables calculated yet. Please generate a layout and ensure the electrical system is properly sized first.");
            return;
        }

        const container = document.getElementById('cable-sizing-content');
        if (!container) return;

        let html = `
            <table style="width: 100%; border-collapse: collapse; background: white; font-size: 0.85rem; text-align: left;">
                <thead style="background: #1e293b; color: white;">
                    <tr>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155;">Category</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155;">Routing Segment (From ➔ To)</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155;">Cable / Line Spec</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155; text-align: right;">Voltage (V)</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155; text-align: right;">Load (kW)</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155; text-align: right;">Length (m)</th>
                        <th style="padding: 0.75rem; border-bottom: 2px solid #334155; text-align: right;">Volt Drop (%)</th>
                    </tr>
                </thead>
                <tbody>
        `;

        const grouped = {};
        // Filter out any residual General entries before display (belt-and-suspenders)
        const cleanSheet = (window._cableSizingSheet || []).filter(row => row.fromTo && row.fromTo !== 'General');
        cleanSheet.forEach(row => {
            if (!grouped[row.category]) grouped[row.category] = [];
            grouped[row.category].push(row);
        });

        // Sort rows naturally
        const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
        Object.keys(grouped).forEach(cat => {
            grouped[cat].sort((a, b) => collator.compare(a.fromTo, b.fromTo));
        });

        const order = ['DC', 'LVAC', 'MV', 'HV', 'GROUNDING'];
        const categories = Object.keys(grouped).sort((a, b) => order.indexOf(a) - order.indexOf(b));

        categories.forEach(cat => {
            html += `
                <tr style="background: #cbd5e1; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='#94a3b8'" onmouseout="this.style.background='#cbd5e1'" onclick="document.querySelectorAll('.cat-row-${cat}').forEach(el => el.style.display = el.style.display === 'none' ? 'table-row' : 'none')">
                    <td colspan="7" style="padding: 0.75rem; font-weight: 700; color: #0f172a; text-transform: uppercase;">
                        <i data-lucide="chevron-down" style="width: 16px; height: 16px; display: inline-block; vertical-align: text-bottom; margin-right: 6px;"></i>
                        ${cat} CABLES (${grouped[cat].length} routing segments)
                    </td>
                </tr>
            `;

            grouped[cat].sort((a, b) => parseFloat(b.loadKw) - parseFloat(a.loadKw)).forEach(row => {
                const isWarning = parseFloat(row.loading) > 80 || parseFloat(row.vdPce) > (cat === 'DC' ? 1.5 : 3.0);
                const color = isWarning ? 'color: #b91c1c; font-weight: 600; background: #fee2e2;' : '';

                html += `
                    <tr class="cat-row-${cat}" style="border-bottom: 1px solid #e2e8f0; ${color} transition: background 0.2s;">
                        <td style="padding: 0.75rem; color: #475569;">${row.category}</td>
                        <td style="padding: 0.75rem;">${row.fromTo}</td>
                        <td style="padding: 0.75rem; font-weight: 600;">${row.cableType}</td>
                        <td style="padding: 0.75rem; text-align: right;">${row.voltageV}</td>
                        <td style="padding: 0.75rem; text-align: right;">${row.loadKw}</td>
                        <td style="padding: 0.75rem; text-align: right;">${row.length}</td>
                        <td style="padding: 0.75rem; text-align: right;">${row.vdPce}</td>
                    </tr>
                `;
            });
        });

        html += `</tbody></table>`;
        container.innerHTML = html;
        if (window.lucide) window.lucide.createIcons();

        document.getElementById('cable-sizing-modal').classList.remove('hidden');
    }
};

window.CablesEngine = CablesEngine;
window.addEventListener('DOMContentLoaded', () => { CablesEngine.init(); });
