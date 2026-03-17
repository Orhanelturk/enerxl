/**
 * CablesEngine - Intelligent Electrical Routing & Optimization
 */

const CablesEngine = {
    cablesDB: {
        dc: [
            { size: 4, type: "Cu/XLPE", R: 5.09, rating: 55, vd_factor: 2.5 },
            { size: 6, type: "Cu/XLPE", R: 3.39, rating: 70, vd_factor: 1.6 },
            { size: 10, type: "Cu/XLPE", R: 1.95, rating: 98, vd_factor: 1.0 }
        ],
        lvac: [
            { size: 70, type: "Al/XLPE", R: 0.443, X: 0.082, rating: 160 },
            { size: 95, type: "Al/XLPE", R: 0.320, X: 0.082, rating: 195 },
            { size: 120, type: "Al/XLPE", R: 0.253, X: 0.080, rating: 225 },
            { size: 150, type: "Al/XLPE", R: 0.206, X: 0.080, rating: 260 },
            { size: 185, type: "Al/XLPE", R: 0.164, X: 0.080, rating: 295 },
            { size: 240, type: "Al/XLPE", R: 0.125, X: 0.079, rating: 345 },
            { size: 300, type: "Al/XLPE", R: 0.100, X: 0.079, rating: 395 },
            { size: 400, type: "Al/XLPE", R: 0.0778, X: 0.078, rating: 460 }
        ],
        mv: [
            { size: 95, type: "Al/XLPE 33kV", R: 0.320, X: 0.124, rating: 220 },
            { size: 150, type: "Al/XLPE 33kV", R: 0.206, X: 0.117, rating: 280 },
            { size: 240, type: "Al/XLPE 33kV", R: 0.125, X: 0.109, rating: 360 },
            { size: 300, type: "Al/XLPE 33kV", R: 0.100, X: 0.105, rating: 410 },
            { size: 400, type: "Al/XLPE 33kV", R: 0.0778, X: 0.100, rating: 470 },
            { size: 500, type: "Al/XLPE 33kV", R: 0.0605, X: 0.097, rating: 530 },
            { size: 630, type: "Al/XLPE 33kV", R: 0.0469, X: 0.093, rating: 600 }
        ],
        ohl: [
            { size: 'Dog', type: "ACSR", R: 0.273, X: 0.36, rating: 300 },
            { size: 'Panther', type: "ACSR", R: 0.136, X: 0.35, rating: 450 },
            { size: 'Zebra', type: "ACSR", R: 0.068, X: 0.34, rating: 650 }
        ]
    },

    selections: {
        dc: null,
        lvac: null,
        mv: null,
        ohl: null
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

    selectBestCable(category, reqAmps) {
        if (this.selections[category] !== null) return this.cablesDB[category][this.selections[category]];
        const db = this.cablesDB[category];
        for (let i = 0; i < db.length; i++) { if (db[i].rating >= reqAmps) return db[i]; }
        return db[db.length - 1];
    },

    calculateCableStats(category, loadMw, voltageKv, pf, lengthMts = 150) {
        if (!this.routingGenerated) return null;
        const soilR = parseFloat(document.getElementById('cable-soil-res')?.value) || 1.5;
        const thermR = parseFloat(document.getElementById('cable-therm-res')?.value) || 1.0;
        const installMethod = document.getElementById('cable-install-method')?.value || 'direct';
        let rDerate = 1.0;
        if (installMethod === 'direct' && soilR > 1.2) rDerate -= (soilR - 1.2) * 0.05;
        if (installMethod === 'ducts') rDerate -= 0.15;
        if (thermR > 1.2) rDerate -= (thermR - 1.2) * 0.08;
        rDerate = Math.max(0.5, rDerate);

        const baseAmps = (loadMw * 1000) / (voltageKv * Math.sqrt(3) * pf);
        const designAmps = baseAmps / rDerate;
        const cable = this.selectBestCable(category, designAmps);
        const loadingPct = (baseAmps / (cable.rating * rDerate)) * 100;
        let vd = 0, vdPce = 0;
        if (category === 'dc') {
            vd = 2 * baseAmps * (lengthMts / 1000) * cable.R;
            vdPce = (vd / (voltageKv * 1000)) * 100;
        } else {
            const sinPhi = Math.sin(Math.acos(pf));
            vd = Math.sqrt(3) * baseAmps * (lengthMts / 1000) * ((cable.R * pf) + ((cable.X || 0) * sinPhi));
            vdPce = (vd / (voltageKv * 1000)) * 100;
        }
        return {
            cable, current: baseAmps, loading: loadingPct, vdPce,
            summaryLine1: `${Array.isArray(cable.size) ? cable.size.join('+') : cable.size} ${typeof cable.size === 'number' ? 'mm²' : ''} ${cable.type}`,
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

                // 1. EXIT FROM SOURCE
                // Goal: Move vertically to the aisle, then horizontally to bestV.
                const exitAisleY = (bestH < g1y_s) ? (g1y_s - tableHalfH - aisleShift) : (g1y_s + tableHalfH + aisleShift);
                addS(g1x_base, exitAisleY);
                addS(bestV, exitAisleY);

                // 2. MAIN TRUNK MOVEMENT (Green and Red Cables)
                // Goal: Move through defined vertical corridors and horizontal roads.
                if (Math.abs(bestV - bestV2) > 0.000001) {
                    addS(bestV, bestH);
                    addS(bestV2, bestH);
                }

                // 3. ENTRY TO DESTINATION
                // Goal: Enter destination X at its nearest aisle Y, then final vertical stub.
                const entryAisleY = (bestH < g2y_s) ? (g2y_s + tableHalfH + aisleShift) : (g2y_s - tableHalfH - aisleShift);
                addS(bestV2, entryAisleY);
                addS(g2x_base, entryAisleY);
                pathLayers.push(p2);
                return pathLayers;


            } catch (e) { console.warn('Smart routing failed:', e); }
        }
        
        // --- Grid-Aware Fallback Routing ---
        // If smart routing fails or missing cache, draw an L-shape that strictly adheres 
        // to the table's rotated grid axes, NOT absolute North/South.
        const gridAngle = (this.cachedHighways && this.cachedHighways.rg) 
            ? -(this.cachedHighways.rotateAngle || 0) 
            : 0;
            
        // Calculate the vector between p1 and p2
        const heading = google.maps.geometry.spherical.computeHeading(p1, p2);
        const dist = google.maps.geometry.spherical.computeDistanceBetween(p1, p2);
        
        // Relative heading difference from our grid's "North"
        const alpha = (heading - gridAngle) * Math.PI / 180;
        
        // Decomposition into grid X and grid Y
        const dxTotal = dist * Math.sin(alpha);
        
        // Offset purely along the grid's X-axis
        const xHeading = dxTotal >= 0 ? (gridAngle + 90) : (gridAngle - 90);
        const ptH = google.maps.geometry.spherical.computeOffset(p1, Math.abs(dxTotal), xHeading);
        
        return [p1, ptH, p2];
    },

    _prepareRoutingCache(rg, tables) {
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
        const rotatedSetback = window.turf.transformRotate(window.LayoutEngine.setbackLinePoly || window.LayoutEngine.usablePoly, rotateAngle, { pivot: siteCenter });

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

        let mvNodes = window.LayoutEngine.blocks.map((poly, i) => ({
            id: i,
            center: toLatLng(poly.item?.center || (poly.getPath ? poly.getPath().getAt(0) : null)),
            blockId: poly.item?.blockId,
            area: poly.area
        }));

        let invNodes = (window.LayoutEngine.inverters || []).map((poly, i) => ({
            id: i,
            center: toLatLng(poly.item?.center || (poly.getPath ? poly.getPath().getAt(0) : null)),
            blockId: poly.item?.blockId,
            area: poly.area
        }));

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

            // Initialize routing cache for THIS area specifically using its saved grid
            const layoutData = window.LayoutEngine.layoutStore.get(area);
            if (!layoutData || !layoutData.routingGrid) return;

            this._prepareRoutingCache(layoutData.routingGrid, layoutData.tables);

            // Process LV (Blue) for this area
            areaInv.forEach(inv => {
                const parentMv = areaMv.find(mv => mv.blockId === inv.blockId);
                if (!parentMv || inv.center.equals(parentMv.center)) return;

                const blockInvs = areaInv.filter(i => i.blockId === inv.blockId);
                const myDistToMv = google.maps.geometry.spherical.computeDistanceBetween(inv.center, parentMv.center);

                let targetNode = parentMv;
                let minDist = myDistToMv;

                blockInvs.forEach(otherInv => {
                    if (otherInv === inv) return;
                    const otherDistToMv = google.maps.geometry.spherical.computeDistanceBetween(otherInv.center, parentMv.center);
                    if (otherDistToMv < myDistToMv - 5) {
                        const distToOther = google.maps.geometry.spherical.computeDistanceBetween(inv.center, otherInv.center);
                        if (distToOther < minDist && distToOther < 250) {
                            minDist = distToOther;
                            targetNode = otherInv;
                        }
                    }
                });

                inv.isDaisyChain = (targetNode !== parentMv);
                const path = this.getSmartPath(inv.center, targetNode.center, inv);
                const line = new google.maps.Polyline({ path, geodesic: true, strokeColor: '#3b82f6', strokeOpacity: 0.6, strokeWeight: 3, map: mapWindow, zIndex: 2000 });
                this.mapPolylines.push(line);
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
                } catch(e) { console.warn("Exit point computation failed", e); }
            }

            for (let i = 0; i < areaMv.length; i += blocksPerLoop) {
                const loop = areaMv.slice(i, i + blocksPerLoop);
                if (hub) {
                    loop.sort((a, b) => google.maps.geometry.spherical.computeDistanceBetween(b.center, hub) - google.maps.geometry.spherical.computeDistanceBetween(a.center, hub));
                }

                for (let j = 0; j < loop.length - 1; j++) {
                    const path = this.getSmartPath(loop[j].center, loop[j + 1].center, null);
                    let len = 0; for (let k = 0; k < path.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(path[k], path[k + 1]);
                    totalMvLength += len;
                    this.mapPolylines.push(new google.maps.Polyline({ path, geodesic: true, strokeColor: '#10b981', strokeOpacity: 0.7, strokeWeight: 4, map: mapWindow, zIndex: 2000 }));
                }

                const lastNode = loop[loop.length - 1];
                if (hub && !lastNode.center.equals(hub)) {
                    // First leg: Internal routing out to the single area exit point
                    const path1 = this.getSmartPath(lastNode.center, exitLatLng, null);
                    let len = 0; for (let k = 0; k < path1.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(path1[k], path1[k + 1]);

                    // Second leg: Exterior routing from exit point to the DS Hub
                    let path2 = [];
                    if (!exitLatLng.equals(hub)) {
                        const p2 = this.getSmartPath(exitLatLng, hub, null);
                        path2 = p2.slice(1); // skip the starting point to prevent duplicate coordinates
                        for (let k = 0; k < p2.length - 1; k++) len += google.maps.geometry.spherical.computeDistanceBetween(p2[k], p2[k + 1]);
                    }

                    const fullPath = path1.concat(path2);
                    totalMvLength += len;
                    this.mapPolylines.push(new google.maps.Polyline({ path: fullPath, geodesic: true, strokeColor: '#ef4444', strokeOpacity: 0.8, strokeWeight: 5, map: mapWindow, zIndex: 2100 }));
                }
            }
        });

        // --- Draw Editable HV Cables (Orange) from DS to POC ---
        if (window.SiteEngine) {
            const SE = window.SiteEngine;
            const pocs = SE.overlays.filter(o => o.subType === 'poc');
            const stations = SE.overlays.filter(o => o.subType === 'station');

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

                    const existingCable = SE.overlays.find(o => o.category === 'hv-cable' && o.startNode === ds.__uid && o.endNode === targetPoc.__uid);

                    if (!existingCable) {
                        const orangeCable = new google.maps.Polyline({
                            path: [dsCenter, pocCenter],
                            geodesic: true,
                            strokeColor: '#f97316',
                            strokeOpacity: 0.9,
                            strokeWeight: 6,
                            map: mapWindow,
                            editable: true,
                            zIndex: 2200
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
                    }
                }
            });
        }

        if (window.SldEngine && !document.getElementById('sld-window').classList.contains('hidden')) window.SldEngine.renderSld();
        if (!silent) alert(`Routing generated successfully!\nTotal estimated MV route length: ${(totalMvLength / 1000).toFixed(2)} km`);
    }
};

window.CablesEngine = CablesEngine;
window.addEventListener('DOMContentLoaded', () => { CablesEngine.init(); });
