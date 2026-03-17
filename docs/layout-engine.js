/**
 * LayoutEngine - Advanced Solar Site Generator (Master Component)
 * 
 * CORE RESPONSIBILITY:
 * 1. Gather site boundaries (Selected Area)
 * 2. Identify Obstacles (Constraint polygons)
 * 3. Incorporate Setbacks/Offsets
 * 4. Apply hierarchical spacing: Module -> Table -> Block -> Gap (Roads)
 * 5. Fit within target capacity or fill available area.
 */

const LayoutEngine = {
    layoutStore: new Map(), // areaObject -> [overlays]
    blocks: [], // Exportable list of generated MV Station blocks for CablesEngine
    inverters: [], // Exportable list of generated Inverters
    isManualConfig: false, // Flag to prevent auto-overwrites if user manually picks

    /**
     * Master generation entry point.
     * Triggered by the main GENERATE button.
     */
    generate() {
        try {
            // 0. Automatic Inverter Configuration Sync
            this.refreshInternalBlockState();

            // 1. Gather all system state
            let state = this._gatherSystemState();
            if (!state.primaryArea) {
                alert("Please select a Project Area on the map first.");
                return;
            }

            // AUTO-ADD ESSENTIAL EQUIPMENT if forgotten
            const SE = window.SiteEngine || {};
            const allAreas = SE.overlays ? SE.overlays.filter(o => o.category === 'area' || (o.getPath && !o.subType && !o.category)) : [];
            const isFirstArea = (allAreas.indexOf(state.primaryArea) === 0 || allAreas.length <= 1);
            
            // Expected Name e.g. "Area 2" -> "DS 2"
            const areaNameStr = state.primaryArea.areaName || `Area ${allAreas.indexOf(state.primaryArea) + 1}`;
            const expectedDsName = areaNameStr.replace(/Area/i, 'DS').trim();
            
            // Has user manually linked a DS, or does one with the expected name exist?
            const hasLinkedDS = state.primaryArea.linkedDsId 
                ? state.pocs.some(o => o.__uid === state.primaryArea.linkedDsId)
                : state.pocs.some(o => o.subType === 'station' && o.areaName === expectedDsName);
                
            let hasGlobalPOC = state.pocs.some(o => o.subType === 'poc');

            if (!hasLinkedDS || (isFirstArea && !hasGlobalPOC)) {
                this._ensureEssentialEquipment(state, hasLinkedDS, hasGlobalPOC, expectedDsName, isFirstArea);
                state = this._gatherSystemState(); // Re-gather with new objects as obstacles
            }

            const area = state.primaryArea;
            if (area) {
                area.lastConfig = state; // Save config for SLD
            }

            // 2. Clear previous generation for THIS area only
            this._clearOldLayout(area);

            // 3. Geometry Preparation
            const usableSiteGeometry = this._processGeometry(state);
            if (!usableSiteGeometry) return;
            this.usablePoly = usableSiteGeometry.usablePoly; // Expose for other engines (Cables)


            // 4. Algorithm Execution
            const layoutPlan = this._calculateOptimalPlacement(usableSiteGeometry, state);
            this.tables = layoutPlan.tables || [];

            // Export routing grid metadata for CablesEngine to compute exact gap positions
            if (layoutPlan.tables.length > 0) {
                const mount = state.mounting;
                const cfg = state.config;
                const azimuth = mount.azimuth;
                const siteCenter = turf.center(usableSiteGeometry.primaryPoly).geometry.coordinates;
                const rotateAngle = 180 - azimuth;

                const isPortrait = mount.orient === 'Portrait';
                const tableW = isPortrait ? (state.pv.w * mount.mx + (mount.mx - 1) * mount.modDistX) : (state.pv.h * mount.mx + (mount.mx - 1) * mount.modDistX);
                const tableH = isPortrait ? (state.pv.h * mount.my + (mount.my - 1) * mount.modDistY) : (state.pv.w * mount.my + (mount.my - 1) * mount.modDistY);
                const stepX = tableW + mount.tableDistX;
                const stepY = tableH + mount.tableDistY;
                const mvGapX = (cfg.divideMV && cfg.everyX > 0) ? (cfg.gapX || 0) : 0;
                const mvGapY = (cfg.divideMV && cfg.everyY > 0) ? (cfg.gapY || 0) : 0;
                const stationX = Math.max(1, cfg.stationX || 10);
                const stationY = Math.max(1, cfg.stationY || 10);

                this.routingGrid = {
                    azimuth, siteCenter, rotateAngle,
                    stepX, stepY, tableW, tableH,
                    mvGapX, mvGapY, stationX, stationY,
                    // Col/row stride of MV gap insertion
                    everyX: Math.max(1, (cfg.everyX || 1) * stationX),
                    everyY: Math.max(1, (cfg.everyY || 1) * stationY),
                    rowTableDist: mount.tableDistY, // gap between rows
                    dominantSlope: layoutPlan.dominantSlope || 0,
                    yStart: layoutPlan.yStart || 0
                };
            }

            // 5. Verification
            if (layoutPlan.tables.length === 0) {
                alert("Could not fit any tables in the selected area with the current constraints.");
                return;
            }

            // 6. Drawing to Map
            this._renderToMap(layoutPlan, area, state);
            
            // Store grids and tables for stable routing later
            const ld = this.layoutStore.get(area);
            if (ld) {
                ld.routingGrid = this.routingGrid;
                ld.tables = layoutPlan.tables;
            }

            // 6b. Rebuild Global Collection for CablesEngine
            this.rebuildGlobalLists();

            // 7. Auto-Generate Electrical Cable Routing
            if (window.CablesEngine) {
                window.CablesEngine.generateRouting(true); // true = silent (no alert)
            }

        } catch (error) {
            console.error("Layout Generation Failed:", error);
            alert("An error occurred during master layout generation. Check console for details.");
        }
    },

    rebuildGlobalLists() {
        this.blocks = [];
        this.inverters = [];
        this.layoutStore.forEach((data, area) => {
            if (data.blocks) this.blocks.push(...data.blocks);
            if (data.inverters) this.inverters.push(...data.inverters);
        });
    },

    /**
     * Clears only the layouts associated with a specific area
     */
    _clearOldLayout(area) {
        const data = this.layoutStore.get(area);
        if (data) {
            if (data.overlays) data.overlays.forEach(overlay => overlay.setMap(null));
            this.layoutStore.delete(area);
        }
    },

    /**
     * Collates input from all UI tabs and the global map state
     */
    _gatherSystemState() {
        this.refreshInternalBlockState(); // Dynamically enforce 200m physical cable constraint

        const SE = window.SiteEngine || {};
        const primaryArea = SE.selectedOverlays.find(o => o.category === 'area' || (o.getPath && !o.subType));
        const obstacles = SE.overlays.filter(o =>
            o.category === 'constraint' ||
            o.category === 'setback' ||
            o.subType === 'setback' ||
            o.category === 'poc' ||
            o.subType === 'poc' ||
            o.subType === 'station'
        );
        const pocs = SE.overlays.filter(o => o.category === 'poc' || o.subType === 'poc' || o.subType === 'station');

        return {
            primaryArea,
            obstacles,
            pocs,
            pv: {
                w: parseFloat(document.getElementById('pv-width').value) / 1000,
                h: parseFloat(document.getElementById('pv-height').value) / 1000,
                p: parseFloat(document.getElementById('pv-pnom').value)
            },
            mounting: {
                type: document.getElementById('mounting-type-select').value,
                orient: document.querySelector('#mounting-orientation-seg .seg-btn.active').dataset.value,
                tilt: parseFloat(document.getElementById('mount-tilt').value),
                azimuth: parseFloat(document.getElementById('mount-azimuth').value),
                modDistX: parseFloat(document.getElementById('mount-mod-dist-x')?.value || "0.02"),
                modDistY: parseFloat(document.getElementById('mount-mod-dist-y')?.value || "0.02"),
                mx: parseInt(document.getElementById('mount-mod-table-x')?.value || "25"),
                my: parseInt(document.getElementById('mount-mod-table-y')?.value || "2"),
                tableDistX: parseFloat(document.getElementById('mount-table-dist-x')?.value || "0.5"), // distance adj tables
                tableDistY: parseFloat(document.getElementById('mount-table-dist-y')?.value || "5"), // row-row distance
                blockConfig: {
                    enabled: document.getElementById('enable-block-config').checked,
                    bx: parseInt(document.getElementById('mount-block-x').value),
                    by: parseInt(document.getElementById('mount-block-y').value),
                    bdx: parseFloat(document.getElementById('mount-block-dist-x').value),
                    bdy: parseFloat(document.getElementById('mount-block-dist-y').value)
                }
            },
            config: {
                objective: document.querySelector('#layout-objective-seg .seg-btn.active').dataset.value,
                targetMW: parseFloat(document.getElementById('cfg-target-mw').value),
                divideMV: document.getElementById('cfg-divide-mv').checked,
                everyX: parseInt(document.getElementById('cfg-dist-x-count')?.value || 1),
                everyY: parseInt(document.getElementById('cfg-dist-y-count')?.value || 1),
                gapX: parseFloat(document.getElementById('cfg-gap-dist-x')?.value || 5),
                gapY: parseFloat(document.getElementById('cfg-gap-dist-y')?.value || 5),
                direction: document.getElementById('cfg-fill-direction')?.value || 'top',
                method: document.querySelector('#layout-method-seg .seg-btn.active').dataset.value,
                resolution: document.querySelector('#draw-resolution-seg .seg-btn.active').dataset.value,
                colorizeBlocks: document.getElementById('cfg-color-blocks').checked || false,
                // Add physical scaling for MV Stations
                stationX: window._lastTablesPerBlock ? window._lastTablesPerBlock.x : 10,
                stationY: window._lastTablesPerBlock ? window._lastTablesPerBlock.y : 10
            },
            electrical: {
                distribution: document.querySelector('#inv-distribution-seg .seg-btn.active').dataset.value,
                strategy: document.getElementById('inv-placement-strategy').value,
                invP: parseFloat(document.getElementById('inv-pnom').value) || 125,
                invDim: {
                    w: parseFloat(document.getElementById('inv-dim-w').value) || 2,
                    l: parseFloat(document.getElementById('inv-dim-l').value) || 0.5
                },
                stationDim: {
                    w: parseFloat(document.getElementById('station-dim-w').value) || 10,
                    l: parseFloat(document.getElementById('station-dim-l').value) || 5
                },
                modPerStr: this.currentBlockConfig?.modsPerStr || 25,
                strPerInv: this.currentBlockConfig?.strsPerInv || 12,
                invPerBlock: this.currentBlockConfig?.invsPerBlock || 8,
                mvCapacities: (document.getElementById('substation-size').value || "9").split(';').map(v => parseFloat(v.trim())).filter(v => !isNaN(v)),
                minLoading: (parseFloat(document.getElementById('min-block-loading')?.value) || 0) / 100,
                useMVStation: document.getElementById('enable-mv-stations').checked,
                useManualHub: document.getElementById('enable-manual-station-loc').checked,
                manualStationLoc: window._manualStationLoc
            }
        };
    },

    /**
     * Internal helper to refresh station dimensions (tables per block)
     * without overwriting user-facing gap configuration.
     */
    refreshInternalBlockState() {
        const substationVal = document.getElementById('substation-size').value || "9";
        const stationMva = parseFloat(substationVal.split(';')[0]) || 9;

        // Fingerprint to detect if hardware specs changed
        const fingerprint = JSON.stringify({
            invP: parseFloat(document.getElementById('inv-pnom')?.value) || 125,
            invV: parseFloat(document.getElementById('inv-vout')?.value) || 400,
            pvP: parseFloat(document.getElementById('pv-pnom')?.value) || 550,
            mva: stationMva
        });

        // If specs changed, force a reset of manual flag and auto-pick best fit
        const isHardwareChange = !this._lastFingerprint || this._lastFingerprint !== fingerprint;

        // Auto-recalc if hardware changed OR if no config exists and NOT manual
        if (isHardwareChange || (!this.currentBlockConfig && !this.isManualConfig)) {
            console.log("Hardware specs changed or missing config, auto-refreshing block design...");
            const options = ElectricalEngine.generateOptions();
            if (options && options.length > 0) {
                const standard = options[0];
                standard._sourceMva = stationMva;
                this.applyBlockConfiguration(standard);
                this._lastFingerprint = fingerprint;
                this.isManualConfig = false; // Hardware change resets manual choice
                return;
            }
        }

        const config = this.currentBlockConfig || { modsPerStr: 25, strsPerInv: 12, invsPerBlock: 8 };
        const rowLength = config.modsPerStr;
        const modsPerTable = rowLength * (parseInt(document.getElementById('mount-mod-table-y')?.value) || 2);

        const totalStationMods = rowLength * config.strsPerInv * config.invsPerBlock;
        const tablesPerStation = Math.ceil(totalStationMods / modsPerTable);

        // Calculate physical table width to enforce 200m DC cable limit per block row
        const modW = (parseFloat(document.getElementById('pv-width')?.value) || 1134) / 1000; // meters
        const modH = (parseFloat(document.getElementById('pv-height')?.value) || 2278) / 1000; // meters
        const modDistX = parseFloat(document.getElementById('mount-mod-dist-x')?.value) || 0;
        const isPortrait = document.querySelector('#mounting-orientation-seg .seg-btn.active')?.dataset.value === 'Portrait';

        const tableW = isPortrait ?
            (modW * rowLength + (rowLength - 1) * modDistX) :
            (modH * rowLength + (rowLength - 1) * modDistX);

        const tableDistX = parseFloat(document.getElementById('mount-table-dist-x')?.value) || 0.5;

        // Formula: width = bX * tableW + (bX - 1) * tableDistX <= 200
        const maxBx = Math.max(1, Math.floor((200 + tableDistX) / (tableW + tableDistX)));

        let bX = Math.ceil(Math.sqrt(tablesPerStation));
        if (bX > maxBx) bX = maxBx;

        const bY = Math.ceil(tablesPerStation / bX);
        window._lastTablesPerBlock = { x: bX, y: bY };
    },

    /**
     * Logic for preparing project geometries
     */
    _processGeometry(state) {
        const SE = window.SiteEngine;
        const primaryPoly = SE.getGeoJSON(state.primaryArea);
        if (!primaryPoly) return null;

        // Hard Subtraction of Obstacles
        let usablePoly = JSON.parse(JSON.stringify(primaryPoly));

        // 1. Always apply an automatic setback to the outer boundary
        try {
            const defaultSetbackVal = parseFloat(document.getElementById('default-setback')?.value) || 10;
            const setbackDist = -defaultSetbackVal / 1000; // negative meters inward in kilometers
            const buffered = turf.buffer(usablePoly, setbackDist, { units: 'kilometers' });
            if (buffered) {
                usablePoly = buffered;
                // Generate a perfect routing ring exactly in the middle of the setback corridor
                const cableDist = -(defaultSetbackVal / 2) / 1000;
                window.LayoutEngine.setbackLinePoly = turf.buffer(JSON.parse(JSON.stringify(primaryPoly)), cableDist, { units: 'kilometers' });
            } else {
                window.LayoutEngine.setbackLinePoly = JSON.parse(JSON.stringify(primaryPoly));
            }
        } catch (e) {
            console.warn("Auto-setback failed", e);
            window.LayoutEngine.setbackLinePoly = usablePoly;
        }

        // 2. Subtract Obstacles from the usable area
        state.obstacles.forEach(obs => {
            const obsGeo = SE.getGeoJSON(obs);
            if (obsGeo) {
                try {
                    const diff = turf.difference(usablePoly, obsGeo);
                    if (diff) usablePoly = diff;
                } catch (e) {
                    console.warn("Obstacle subtraction failed:", e);
                }
            }
        });

        const bbox = turf.bbox(usablePoly);
        return { usablePoly, primaryPoly, bbox };
    },

    _calculateOptimalPlacement(geometry, state) {
        if (state.config.method === 'adaptive') {
            return this._calculateAdaptivePlacement(geometry, state);
        } else {
            return this._calculateBlockPlacement(geometry, state);
        }
    },

    /**
     * Finds the slope (dx/dy) of the longest edge segment located on the selected bounding box side.
     * This is used to "shear" or stagger the grid so the layout overall perfectly parallels the site boundary,
     * WITHOUT rotating the individual PV modules (which keeps them facing South).
     */
    _findDominantSlopeForDirection(poly, direction) {
        let dominantSlope = 0;
        try {
            let coords = [];
            if (poly.geometry.type === 'Polygon') coords = poly.geometry.coordinates[0];
            else if (poly.geometry.type === 'MultiPolygon') coords = poly.geometry.coordinates[0][0];

            if (!coords || coords.length < 3) return 0;

            const bbox = turf.bbox(poly);
            const searchMarginX = (bbox[2] - bbox[0]) * 0.15; // 15% margin
            const searchMarginY = (bbox[3] - bbox[1]) * 0.15;

            let maxLen = -1;
            for (let i = 0; i < coords.length - 1; i++) {
                const pt1 = coords[i];
                const pt2 = coords[i + 1];

                const midX = (pt1[0] + pt2[0]) / 2;
                const midY = (pt1[1] + pt2[1]) / 2;

                let isOnTargetSide = false;
                switch (direction) {
                    case 'left': isOnTargetSide = midX < bbox[0] + searchMarginX; break;
                    case 'right': isOnTargetSide = midX > bbox[2] - searchMarginX; break;
                    case 'bottom': isOnTargetSide = midY < bbox[1] + searchMarginY; break;
                    case 'top': isOnTargetSide = midY > bbox[3] - searchMarginY; break;
                    default: isOnTargetSide = true; break; // 'middle' or unselected
                }

                if (isOnTargetSide) {
                    const dxSeg = pt2[0] - pt1[0];
                    const dySeg = pt2[1] - pt1[1];
                    const len = Math.sqrt(dxSeg * dxSeg + dySeg * dySeg);

                    if (len > maxLen) {
                        maxLen = len;
                        if (Math.abs(dySeg) > 1e-8) {
                            let slope = dxSeg / dySeg;
                            if (Math.abs(slope) < 5) dominantSlope = slope;
                            else dominantSlope = 0;
                        } else dominantSlope = 0;
                    }
                }
            }
        } catch (e) {
            console.error("Slope auto-calculation failed", e);
        }
        return dominantSlope;
    },

    _calculateAdaptivePlacement(geometry, state) {
        let allTables = [];
        const { usablePoly, primaryPoly } = geometry;
        const mount = state.mounting;
        const bc = mount.blockConfig;
        const pv = state.pv;
        const cfg = state.config;

        // Table dimensions
        const isPortrait = mount.orient === 'Portrait';
        const tableW = isPortrait ?
            (pv.w * mount.mx + (mount.mx - 1) * mount.modDistX) :
            (pv.h * mount.mx + (mount.mx - 1) * mount.modDistX);
        const tableH = isPortrait ?
            (pv.h * mount.my + (mount.my - 1) * mount.modDistY) :
            (pv.w * mount.my + (mount.my - 1) * mount.modDistY);

        const stepX = tableW + mount.tableDistX;
        const stepY = tableH + mount.tableDistY;

        // Gaps
        const mvEnabled = cfg.divideMV === true;
        const bcEnabled = bc && bc.enabled === true;
        const mvGapX = mvEnabled ? (cfg.gapX || 0) : 0;
        const mvGapY = mvEnabled ? (cfg.gapY || 0) : 0;
        const bcGapX = bcEnabled ? (bc.bdx || 0) : 0;
        const bcGapY = bcEnabled ? (bc.bdy || 0) : 0;

        const azimuth = mount.azimuth;
        const siteCenter = turf.center(primaryPoly).geometry.coordinates;
        const rotateAngle = 180 - azimuth;
        const rotatedUsable = turf.transformRotate(usablePoly, rotateAngle, { pivot: siteCenter });
        const rb = turf.bbox(rotatedUsable);

        const mToDeg = 1 / 111320;
        const cosLat = Math.cos(siteCenter[1] * Math.PI / 180);
        const dx = stepX * mToDeg / cosLat;
        const dy = stepY * mToDeg;
        const mvgx = mvGapX * mToDeg / cosLat;
        const mvgy = mvGapY * mToDeg;
        const bcgx = bcGapX * mToDeg / cosLat;
        const bcgy = bcGapY * mToDeg;

        const direction = cfg.direction;
        const isBottom = direction === 'bottom';
        const isRightDir = direction === 'right';
        const isEverywhereCenter = direction === 'middle';

        // Find dominant slope of the rotated polygon to shear the adaptive grid uniformly
        // targeted strictly to the requested starting direction edge
        const dominantSlope = this._findDominantSlopeForDirection(rotatedUsable, direction);

        // Row Scan Order
        let yStart = rb[3] - dy / 2;
        let yDir = -1;
        if (isBottom) {
            yStart = rb[1] + dy / 2;
            yDir = 1;
        }

        // Column Alignment Preference
        let xStart = rb[0] + dx / 2;
        let xDir = 1;
        if (isRightDir) {
            xStart = rb[2] - dx / 2;
            xDir = -1;
        }

        // If originating from the Left and we have a positive/negative slope, 
        // the sheared row might drift away from the bounding box edge.
        // We calculate the maximum drift over the Y span and pull xStart back 
        // to ensure the layout strictly touches the left boundary.
        if (direction === 'left') {
            const ySpan = rb[3] - rb[1];
            const maxShear = ySpan * dominantSlope;
            // If shear is positive, rows drift right; shift origin further left
            if (maxShear > 0) xStart -= maxShear;
            // If shear is negative, rows drift left; shift origin further right to keep start tight
            if (maxShear < 0) xStart -= maxShear;
        } else if (direction === 'right') {
            const ySpan = rb[3] - rb[1];
            const maxShear = ySpan * dominantSlope;
            // If shear is negative, rows drift left; shift origin further right to compensate
            if (maxShear < 0) xStart -= maxShear;
            if (maxShear > 0) xStart -= maxShear;
        }

        // Similarly for the bottom/top if we wanted to bound them tightly, but 
        // row scanning inherently fixes Y.

        const everyX = Math.max(1, (cfg.everyX || 1) * (cfg.stationX || 1));
        const everyY = Math.max(1, (cfg.everyY || 1) * (cfg.stationY || 1));
        const bcEveryY = bc.by || 1;
        const bcEveryX = bc.bx || 1;

        const maxRows = Math.ceil((rb[3] - rb[1]) / dy) + 15;
        const maxCols = Math.ceil((rb[2] - rb[0]) / dx) + 15;

        for (let ri = 0; ri < maxRows; ri++) {
            const yOffset = (mvEnabled && cfg.everyY > 0 ? Math.floor(ri / everyY) * mvgy : 0) +
                (bcEnabled && bc.by > 0 ? Math.floor(ri / bcEveryY) * bcgy : 0);

            const currentY = yStart + (ri * yDir * dy) + (yDir * yOffset);
            if (yDir > 0 ? (currentY > rb[3] + dy) : (currentY < rb[1] - dy)) break;

            const rowShearShift = (currentY - yStart) * dominantSlope;

            // Compute dynamic loop bounds to guarantee we completely blanket the polygon's bounding box 
            // no matter how violently rowShearShift pushes the sheared matrix left or right.
            const targetMinX = rb[0] - dx * 2;
            const targetMaxX = rb[2] + dx * 2;

            let minCol, maxCol;
            if (xDir > 0) {
                minCol = Math.floor((targetMinX - xStart - rowShearShift) / dx);
                maxCol = Math.ceil((targetMaxX - xStart - rowShearShift) / dx);
            } else {
                minCol = Math.floor(-(targetMaxX - xStart - rowShearShift) / dx);
                maxCol = Math.ceil(-(targetMinX - xStart - rowShearShift) / dx);
            }

            // Heavily overprovision the loop safely based strictly on physical geometry containment
            for (let colIdx = minCol; colIdx <= maxCol; colIdx++) {
                const xOffset = (mvEnabled && cfg.everyX > 0 ? Math.floor(colIdx / everyX) * mvgx : 0) +
                    (bcEnabled && bc.bx > 0 ? Math.floor(colIdx / bcEveryX) * bcgx : 0);

                const currentX = xStart + rowShearShift + (colIdx * xDir * dx) + (xDir * xOffset);

                // Quick discard outside rendering box to prevent heavy Turf checks
                if (currentX < targetMinX || currentX > targetMaxX) continue;

                const tablePt = turf.point([currentX, currentY]);
                if (turf.booleanPointInPolygon(tablePt, rotatedUsable)) {
                    const realPoint = turf.transformRotate(tablePt, -rotateAngle, { pivot: siteCenter });
                    const tablePoly = this._createTablePolygon(realPoint.geometry.coordinates, tableW, tableH, azimuth);

                    if (turf.booleanWithin(tablePoly, usablePoly)) {
                        allTables.push({
                            geometry: tablePoly,
                            center: realPoint.geometry.coordinates,
                            gridX: currentX,
                            gridY: currentY,
                            colIdx: colIdx,
                            rowIdx: ri
                        });
                    }
                }
            }
        }
        const layoutPlan = this._sortAndFinalize(allTables, geometry, state);
        layoutPlan.dominantSlope = dominantSlope;
        layoutPlan.yStart = yStart;
        return layoutPlan;
    },

    _calculateBlockPlacement(geometry, state) {
        let allTables = [];
        const { usablePoly, primaryPoly } = geometry;
        const mount = state.mounting;
        const bc = mount.blockConfig;
        const pv = state.pv;
        const cfg = state.config;

        const isPortrait = mount.orient === 'Portrait';
        const tableW = isPortrait ?
            (pv.w * mount.mx + (mount.mx - 1) * mount.modDistX) :
            (pv.h * mount.mx + (mount.mx - 1) * mount.modDistX);
        const tableH = isPortrait ?
            (pv.h * mount.my + (mount.my - 1) * mount.modDistY) :
            (pv.w * mount.my + (mount.my - 1) * mount.modDistY);

        const stepX = tableW + mount.tableDistX;
        const stepY = tableH + mount.tableDistY;

        const mvEnabled = cfg.divideMV === true;
        const bcEnabled = bc && bc.enabled === true;
        const mvGapX = mvEnabled ? (cfg.gapX || 0) : 0;
        const mvGapY = mvEnabled ? (cfg.gapY || 0) : 0;
        const bcGapX = bcEnabled ? (bc.bdx || 0) : 0;
        const bcGapY = bcEnabled ? (bc.bdy || 0) : 0;

        const azimuth = mount.azimuth;
        const siteCenter = turf.center(primaryPoly).geometry.coordinates;
        const rotateAngle = 180 - azimuth;
        const rotatedUsable = turf.transformRotate(usablePoly, rotateAngle, { pivot: siteCenter });
        const rb = turf.bbox(rotatedUsable);

        const mToDeg = 1 / 111320;
        const cosLat = Math.cos(siteCenter[1] * Math.PI / 180);
        const dx = stepX * mToDeg / cosLat;
        const dy = stepY * mToDeg;
        const mvgx = mvGapX * mToDeg / cosLat;
        const mvgy = mvGapY * mToDeg;
        const bcgx = bcGapX * mToDeg / cosLat;
        const bcgy = bcGapY * mToDeg;

        const direction = cfg.direction;
        const isBottom = direction === 'bottom';
        const isRightDir = direction === 'right';
        const isMiddle = direction === 'middle';

        const everyX = Math.max(1, (cfg.everyX || 1) * (cfg.stationX || 1));
        const everyY = Math.max(1, (cfg.everyY || 1) * (cfg.stationY || 1));
        const bcEveryY = bc.by || 1;
        const bcEveryX = bc.bx || 1;

        // Determine Start Points
        let yStart = isBottom ? rb[1] + dy / 2 : rb[3] - dy / 2;
        let yDir = isBottom ? 1 : -1;
        let xStart = isRightDir ? rb[2] - dx / 2 : rb[0] + dx / 2;
        let xDir = isRightDir ? -1 : 1;

        if (isMiddle) {
            // Center the grid - find bounding box center and back-calculate yStart/xStart 
            // to align a potential table exactly in the middle.
            const midY = (rb[1] + rb[3]) / 2;
            const midX = (rb[0] + rb[2]) / 2;
            yStart = midY;
            xStart = midX;
            yDir = -1; // Default to Top-Down scan
            xDir = 1;  // Default to Left-Right scan
        }

        const maxCols = Math.ceil((rb[2] - rb[0]) / dx) + 50;
        const maxRows = Math.ceil((rb[3] - rb[1]) / dy) + 50;

        for (let colIdx = 0; colIdx < maxCols; colIdx++) {
            const xOffset = (mvEnabled && cfg.everyX > 0 ? Math.floor(colIdx / everyX) * mvgx : 0) +
                (bcEnabled && bc.bx > 0 ? Math.floor(colIdx / bcEveryX) * bcgx : 0);
            const currentX = xStart + (colIdx * xDir * dx) + (xDir * xOffset);

            if (xDir > 0 ? (currentX > rb[2] + dx) : (currentX < rb[0] - dx)) break;

            for (let rowIdx = 0; rowIdx < maxRows; rowIdx++) {
                const yOffset = (mvEnabled && cfg.everyY > 0 ? Math.floor(rowIdx / everyY) * mvgy : 0) +
                    (bcEnabled && bc.by > 0 ? Math.floor(rowIdx / bcEveryY) * bcgy : 0);
                const currentY = yStart + (rowIdx * yDir * dy) + (yDir * yOffset);

                if (yDir > 0 ? (currentY > rb[3] + dy) : (currentY < rb[1] - dy)) break;

                const pt = turf.point([currentX, currentY]);
                if (turf.booleanPointInPolygon(pt, rotatedUsable)) {
                    const realPoint = turf.transformRotate(pt, -rotateAngle, { pivot: siteCenter });
                    const tablePoly = this._createTablePolygon(realPoint.geometry.coordinates, tableW, tableH, azimuth);

                    if (turf.booleanWithin(tablePoly, usablePoly)) {
                        allTables.push({
                            geometry: tablePoly,
                            center: realPoint.geometry.coordinates,
                            gridX: currentX,
                            gridY: currentY,
                            colIdx: colIdx,
                            rowIdx: rowIdx
                        });
                    }
                }
            }
        }
        
        return this._sortAndFinalize(allTables, geometry, state);
    },

    /**
     * Sorting and Capacity Finalization
     */
    _sortAndFinalize(allTables, geometry, state) {
        const cfg = state.config;

        // Apply Directional Priority Sorting
        switch (cfg.direction) {
            case 'left':
                allTables.sort((a, b) => a.gridX - b.gridX);
                break;
            case 'right':
                allTables.sort((a, b) => b.gridX - a.gridX);
                break;
            case 'top':
                allTables.sort((a, b) => b.gridY - a.gridY);
                break;
            case 'bottom':
                allTables.sort((a, b) => a.gridY - b.gridY);
                break;
            case 'middle':
                const center = turf.center(geometry.primaryPoly).geometry.coordinates;
                allTables.sort((a, b) => {
                    const dA = turf.distance(turf.point(a.center), turf.point(center));
                    const dB = turf.distance(turf.point(b.center), turf.point(center));
                    return dA - dB;
                });
                break;
        }

        // Handle Capacity Objective
        if (cfg.objective === 'capacity') {
            const tableMW = (state.mounting.mx * state.mounting.my * state.pv.p) / 1000000;
            const tablesNeeded = Math.ceil(cfg.targetMW / tableMW);
            allTables = allTables.slice(0, tablesNeeded);
        }

        return this._finalizeResults(allTables, geometry, state);
    },

    /**
     * Precise vertex generation using polar projection
     */
    _createTablePolygon(centerPort, w, h, az) {
        const center = turf.point(centerPort);
        const halfW = w / 2;
        const halfH = h / 2;

        const corners = [
            this._getMetricOffset(center, -halfW, -halfH, az),
            this._getMetricOffset(center, halfW, -halfH, az),
            this._getMetricOffset(center, halfW, halfH, az),
            this._getMetricOffset(center, -halfW, halfH, az),
            this._getMetricOffset(center, -halfW, -halfH, az)
        ];
        return turf.polygon([corners.map(p => p.geometry.coordinates)]);
    },

    _getMetricOffset(startPoint, dx, dy, azimuth) {
        const angleRad = (azimuth) * (Math.PI / 180);
        // Correct vector components for Map bearing logic
        const rx = dx * Math.cos(angleRad) + dy * Math.sin(angleRad);
        const ry = -dx * Math.sin(angleRad) + dy * Math.cos(angleRad);
        const dist = Math.sqrt(rx * rx + ry * ry);
        const bearing = (Math.atan2(rx, ry) * 180 / Math.PI);
        return turf.destination(startPoint, dist, bearing, { units: 'meters' });
    },

    _finalizeResults(tables, geometry, state) {
        const mount = state.mounting;
        const elec = state.electrical;
        const cfg = state.config;

        const modsPerTable = mount.mx * mount.my;

        // Assign Block IDs and calculate block stats
        const blockSummary = {};
        if (cfg.divideMV) {
            const stationX = Math.max(1, cfg.stationX || 10);
            const stationY = Math.max(1, cfg.stationY || 10);

            const blockCoordsToId = new Map();
            let nextBlockNum = 1;

            tables.forEach(table => {
                const bX = Math.floor(table.colIdx / stationX);
                const bY = Math.floor(table.rowIdx / stationY);
                const coordKey = `${bX}_${bY}`;

                if (!blockCoordsToId.has(coordKey)) {
                    blockCoordsToId.set(coordKey, `Block ${nextBlockNum++}`);
                }

                table.blockId = blockCoordsToId.get(coordKey);

                if (!blockSummary[table.blockId]) blockSummary[table.blockId] = { count: 0, mw: 0 };
                blockSummary[table.blockId].count++;
                blockSummary[table.blockId].mw += (modsPerTable * state.pv.p) / 1000000;
            });

            // Optional: Filter blocks based on minimum loading criteria
            if (elec.minLoading > 0) {
                const excludedBlocks = new Set();
                const blockIds = Object.keys(blockSummary);

                blockIds.forEach((blockId, idx) => {
                    const capacityList = elec.mvCapacities || [9];
                    // Match block index to capacity list
                    const mva = capacityList[idx % capacityList.length] || capacityList[capacityList.length - 1] || 1;
                    const loading = blockSummary[blockId].mw / mva;

                    if (loading < elec.minLoading) {
                        excludedBlocks.add(blockId);
                    }
                });

                if (excludedBlocks.size > 0) {
                    tables = tables.filter(t => !excludedBlocks.has(t.blockId));
                    excludedBlocks.forEach(b => delete blockSummary[b]);
                }
            }

            // Update final capacities for tooltips
            tables.forEach(table => {
                if (blockSummary[table.blockId]) {
                    table.blockMW = blockSummary[table.blockId].mw;
                }
            });
        } else if (tables.length > 0) {
            // Default single block for inverter placement if MV is disabled
            const blockId = "Project Block";
            const projectMW = (tables.length * modsPerTable * state.pv.p) / 1000000;
            tables.forEach(table => {
                table.blockId = blockId;
                table.blockMW = projectMW;
            });
            blockSummary[blockId] = {
                count: tables.length,
                mw: projectMW
            };
        }

        const totalModules = tables.length * modsPerTable;
        const totalMW = (totalModules * state.pv.p) / 1000000;

        // NEW: Place Electrical Equipment (Stations/Inverters)
        const electricalPlan = this._placeElectricalEquipment(tables, blockSummary, geometry, state);
        if (electricalPlan.excludedTables.size > 0) {
            tables = tables.filter(t => !electricalPlan.excludedTables.has(t));
        }

        // Save stats for summary box access
        window._layoutStats = window._layoutStats || {};
        const areaName = state.primaryArea?.areaName || 'Project Area';
        const acdcEl = document.getElementById('inv-acdc-ratio');
        const acdcRatio = acdcEl ? parseFloat(acdcEl.value) : 1.2;
        const acMW = (totalMW / acdcRatio).toFixed(2);

        window._layoutStats[areaName] = {
            dcMWp: totalMW.toFixed(2),
            acMW: acMW,
            moduleCount: totalModules,
            blockCount: Object.keys(blockSummary).length || (cfg.divideMV ? 1 : 0),
            invCount: electricalPlan.equipment.length,
            name: areaName,
            areaSqm: Math.round(turf.area(geometry.primaryPoly))
        };

        if (window.updateProjectSummary) window.updateProjectSummary();

        // Update System Tab fields if not in manual override mode
        const chkManual = document.getElementById('chk-manual-cap');
        if (chkManual && !chkManual.checked) {
            const dcInput = document.getElementById('sys-dc-cap');
            const acInput = document.getElementById('sys-ac-cap');
            if (dcInput) dcInput.value = totalMW.toFixed(2);
            if (acInput) acInput.value = acMW;

            // Trigger electrical engine to refresh its summaries based on new actual capacity
            if (window.ElectricalEngine) window.ElectricalEngine.updatePowerDisplay();
        }

        return { tables, totalModules, totalMW, blockSummary, equipment: electricalPlan.equipment };
    },

    /**
     * Places Inverter Stations or row inverters based on gathered state
     */
    _placeElectricalEquipment(tables, blockSummary, geometry, state) {
        const elec = state.electrical;
        const mount = state.mounting;
        const equipment = [];
        const excludedTables = new Set();
        const pocs = state.pocs || [];

        // Calculation constants for table span
        const isPortrait = mount.orient === 'Portrait';
        const tableWidthM = isPortrait ?
            (mount.mx * state.pv.w) + ((mount.mx - 1) * mount.modDistX) :
            (mount.mx * state.pv.h) + ((mount.mx - 1) * mount.modDistX);
        const az = mount.azimuth;

        // Find main collection point (Prioritize DS over POC)
        let mainPoc = null;
        if (pocs.length > 0) {
            const dsIndex = pocs.findIndex(o => o.subType === 'station');
            const targetPoc = (dsIndex !== -1) ? pocs[dsIndex] : pocs[0];

            if (targetPoc.getPosition) {
                mainPoc = [targetPoc.getPosition().lng(), targetPoc.getPosition().lat()];
            } else if (targetPoc.getBounds) {
                const c = targetPoc.getBounds().getCenter();
                mainPoc = [c.lng(), c.lat()];
            } else {
                // Polygon - get centroid via SiteEngine/Turf
                const geo = window.SiteEngine.getGeoJSON(targetPoc);
                if (geo) {
                    const center = turf.centroid(geo);
                    mainPoc = center.geometry.coordinates;
                }
            }
        }

        const blockIds = Object.keys(blockSummary);

        blockIds.forEach(blockId => {
            const blockTables = tables.filter(t => t.blockId === blockId);
            if (blockTables.length === 0) return;

            // --- 1. Calculate the ideal Anchor Point for this block ---
            const centers = blockTables.map(t => t.center);
            const avgX = centers.reduce((a, b) => a + b[0], 0) / centers.length;
            const avgY = centers.reduce((a, b) => a + b[1], 0) / centers.length;
            let blockAnchor = [avgX, avgY];

            if (mainPoc) {
                const distToPoc = turf.distance(blockAnchor, mainPoc);
                // Significantly pull the search anchor toward the POC to prioritize the "front" side of the block
                blockAnchor = turf.destination(blockAnchor, distToPoc * 0.25, turf.bearing(blockAnchor, mainPoc)).geometry.coordinates;
            }

            let stationLoc = null;
            if (elec.useManualHub && elec.manualStationLoc) {
                stationLoc = elec.manualStationLoc;
            } else if (elec.strategy === 'partition') {
                const sortedX = [...blockTables].sort((a, b) => a.center[0] - b.center[0]);
                const sortedY = [...blockTables].sort((a, b) => a.center[1] - b.center[1]);
                const corners = [
                    [sortedX[0].center[0], sortedY[0].center[1]],
                    [sortedX[0].center[0], sortedY[sortedY.length - 1].center[1]],
                    [sortedX[sortedX.length - 1].center[0], sortedY[0].center[1]],
                    [sortedX[sortedX.length - 1].center[0], sortedY[sortedY.length - 1].center[1]]
                ];

                let closestCorner = corners[0];
                let minDist = Infinity;
                corners.forEach(c => {
                    const d = turf.distance(c, blockAnchor);
                    if (d < minDist) { minDist = d; closestCorner = c; }
                });

                let bestTable = null;
                let minDistToCorner = Infinity;
                blockTables.forEach(t => {
                    const d = turf.distance(t.center, closestCorner);
                    if (d < minDistToCorner) { minDistToCorner = d; bestTable = t; }
                });

                if (bestTable) {
                    stationLoc = bestTable.center;
                    if (elec.useMVStation) excludedTables.add(bestTable);
                }
            } else {
                const sortedY = blockTables.map(t => t.center[1]).sort((a, b) => a - b);
                const isNorth = (mainPoc && mainPoc[1] > avgY);
                // Offset ~5m into the road gap
                const roadY = isNorth ? (sortedY[sortedY.length - 1] + 0.00005) : (sortedY[0] - 0.00005);
                stationLoc = [avgX, roadY];
            }

            if (!stationLoc) stationLoc = blockAnchor;

            // --- 2. Place MV Station ---
            if (elec.useMVStation) {
                let currentStationLoc = stationLoc;
                let stationPoly = this._createTablePolygon(currentStationLoc, elec.stationDim.w, elec.stationDim.l, az);

                // Check if station is in safe zone. If not, nudge towards blockAnchor
                if (!turf.booleanWithin(stationPoly, geometry.usablePoly)) {
                    const bearingToCenter = turf.bearing(currentStationLoc, blockAnchor);
                    let iterations = 0;
                    while (!turf.booleanWithin(stationPoly, geometry.usablePoly) && iterations < 15) {
                        currentStationLoc = turf.destination(currentStationLoc, 2, bearingToCenter, { units: 'meters' }).geometry.coordinates;
                        stationPoly = this._createTablePolygon(currentStationLoc, elec.stationDim.w, elec.stationDim.l, az);
                        iterations++;
                    }
                }

                const station = {
                    type: 'MV Station',
                    center: currentStationLoc,
                    w: elec.stationDim.w, l: elec.stationDim.l,
                    blockId: blockId,
                    rating: (elec.mvCapacities[0] || 3.15) + " MVA"
                };
                equipment.push(station);

                // Clear overlaps for station
                blockTables.forEach(t => {
                    if (!excludedTables.has(t) && turf.booleanIntersects(stationPoly, t.geometry)) {
                        excludedTables.add(t);
                    }
                });

                // Update stationLoc for inverter references if it moved
                stationLoc = currentStationLoc;
            }

            // --- 3. Place Inverters ---
            let count = elec.invPerBlock || 1;
            if (!state.config.divideMV) {
                const acdc = parseFloat(document.getElementById('inv-acdc-ratio')?.value) || 1.2;
                const invP = parseFloat(document.getElementById('inv-pnom').value) || 125;
                count = Math.ceil((blockSummary[blockId].mw * 1000) / invP / acdc);
            }

            // Metadata for tooltips
            const invKwAc = parseFloat(document.getElementById('inv-pnom').value) || 125;
            const pvP = parseFloat(document.getElementById('pv-pnom').value) || 550;
            const invKwDc = (elec.modPerStr * elec.strPerInv * pvP) / 1000;

            if (elec.distribution === 'station') {
                const cols = Math.ceil(Math.sqrt(count));
                const spacingX = 3.0; // meters
                const spacingY = 2.5; // meters

                for (let i = 0; i < count; i++) {
                    const localDx = (i % cols - (cols - 1) / 2) * spacingX;
                    const localDy = (Math.floor(i / cols) + 1) * spacingY;

                    const invPt = this._getMetricOffset(turf.point(stationLoc), localDx, localDy, az);
                    let invCenter = invPt.geometry.coordinates;
                    let invPoly = this._createTablePolygon(invCenter, elec.invDim.w, elec.invDim.l, az);

                    // Boundary Check for centralized inverters
                    if (!turf.booleanWithin(invPoly, geometry.usablePoly)) {
                        const bearingToStation = turf.bearing(invCenter, stationLoc);
                        let iterations = 0;
                        while (!turf.booleanWithin(invPoly, geometry.usablePoly) && iterations < 10) {
                            invCenter = turf.destination(invCenter, 1, bearingToStation, { units: 'meters' }).geometry.coordinates;
                            invPoly = this._createTablePolygon(invCenter, elec.invDim.w, elec.invDim.l, az);
                            iterations++;
                        }
                    }

                    const inv = {
                        type: 'Inverter',
                        id: `INV-${blockId}.${i + 1}`,
                        center: invCenter,
                        w: elec.invDim.w, l: elec.invDim.l,
                        azimuth: az,
                        blockId: blockId,
                        dcKw: invKwDc.toFixed(1),
                        acKw: invKwAc.toFixed(1)
                    };
                    equipment.push(inv);

                    // Clear overlaps
                    blockTables.forEach(t => {
                        if (!excludedTables.has(t) && turf.booleanIntersects(invPoly, t.geometry)) {
                            excludedTables.add(t);
                        }
                    });
                }
            } else {
                // Distributed to Rows: Place EXACT count, distributed among rows
                const rowIds = [...new Set(blockTables.map(t => t.rowIdx))].sort((a, b) => a - b);
                if (rowIds.length === 0) return;

                for (let i = 0; i < count; i++) {
                    // Pick row statistically to spread evenly
                    const targetRowIndex = Math.floor(i * rowIds.length / count);
                    const rowId = rowIds[targetRowIndex];

                    const rowTablesInBlock = blockTables.filter(t => t.rowIdx === rowId);
                    const activeRowTables = rowTablesInBlock.filter(t => !excludedTables.has(t));
                    if (activeRowTables.length === 0) continue;

                    const sortedX = [...activeRowTables].sort((a, b) => a.center[0] - b.center[0]);
                    const tableL = sortedX[0];
                    const tableR = sortedX[sortedX.length - 1];

                    const dL = turf.distance(tableL.center, stationLoc);
                    const dR = turf.distance(tableR.center, stationLoc);
                    const targetTable = (dL < dR) ? tableL : tableR;

                    // If multiple inverters fall on same row, shift them slightly
                    const instancesOnThisRow = Math.floor(i / count * rowIds.length) - Math.floor((i - 1) / count * rowIds.length);
                    const rowOffsetMultiplier = (i % Math.max(1, Math.floor(count / rowIds.length))) || 0;

                    const totalOffset = (tableWidthM / 2) + (elec.invDim.w / 2) + 0.35 + (rowOffsetMultiplier * 2.0);
                    const bL = (az - 90 + 360) % 360;
                    const bR = (az + 90 + 360) % 360;

                    const otherRef = (activeRowTables.length > 1)
                        ? (targetTable === tableL ? tableR.center : tableL.center)
                        : blockAnchor;

                    const testL = turf.destination(targetTable.center, 1, bL, { units: 'meters' });
                    const testR = turf.destination(targetTable.center, 1, bR, { units: 'meters' });
                    const choiceBearing = (turf.distance(testL, otherRef) > turf.distance(testR, otherRef)) ? bL : bR;

                    let invCenter = turf.destination(targetTable.center, totalOffset, choiceBearing, { units: 'meters' }).geometry.coordinates;
                    let invAz = (az + 90) % 360;
                    let invPoly = this._createTablePolygon(invCenter, elec.invDim.w, elec.invDim.l, invAz);

                    if (!turf.booleanWithin(invPoly, geometry.usablePoly)) {
                        const bearingToRow = turf.bearing(invCenter, targetTable.center);
                        let iterations = 0;
                        while (!turf.booleanWithin(invPoly, geometry.usablePoly) && iterations < 15) {
                            invCenter = turf.destination(invCenter, 1, bearingToRow, { units: 'meters' }).geometry.coordinates;
                            invPoly = this._createTablePolygon(invCenter, elec.invDim.w, elec.invDim.l, invAz);
                            iterations++;
                        }
                    }

                    equipment.push({
                        type: 'Inverter',
                        id: `INV-${blockId}.${i + 1}`,
                        center: invCenter,
                        w: elec.invDim.w, l: elec.invDim.l,
                        azimuth: invAz,
                        blockId: blockId,
                        rowId: rowId,
                        dcKw: invKwDc.toFixed(1),
                        acKw: invKwAc.toFixed(1)
                    });

                    blockTables.forEach(t => {
                        if (!excludedTables.has(t) && turf.booleanIntersects(invPoly, t.geometry)) {
                            excludedTables.add(t);
                        }
                    });
                }
            }
        });

        return { equipment, excludedTables };
    },

    _renderToMap(plan, area, state) {
        const SE = window.SiteEngine;
        if (!SE) return;

        const stateColorize = state.config.colorizeBlocks;

        const blockColors = [
            '#ff0055', '#39ff14', '#00ffff', '#ffcc00', '#ff00ff', '#2e31ff',
            '#ff6600', '#00ffcc', '#bf00ff', '#ccff00', '#ff0088', '#0099ff',
            '#e60000', '#00e600', '#0000e6', '#e6e600', '#e600e6', '#00e6e6',
            '#8b0000', '#006400', '#00008b', '#8b8b00', '#8b008b', '#008b8b'
        ];

        const overlays = [];
        // Create or get global tooltip element
        let tooltip = document.getElementById('map-tooltip');
        if (!tooltip) {
            tooltip = document.createElement('div');
            tooltip.id = 'map-tooltip';
            tooltip.style.cssText = 'position: fixed; pointer-events: none; z-index: 3000; background: rgba(255, 255, 255, 0.9); backdrop-filter: blur(12px); color: #1e293b; padding: 4px 10px; border-radius: 8px; font-size: 11px; font-family: Inter, sans-serif; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border: 1px solid rgba(255,255,255,0.5); white-space: nowrap; font-weight: 500;';
            document.body.appendChild(tooltip);
        }

        const res = state.config.resolution;
        const totalDC = parseFloat(document.getElementById('sys-dc-cap')?.value) || 0;
        const resThreshold = 50;
        const effectiveRes = (totalDC > resThreshold) ? 'tables' : res;

        const mount = state.mounting;
        const pv = state.pv;
        const az = mount.azimuth;

        plan.tables.forEach(table => {
            const tableCoords = table.geometry.geometry.coordinates[0].map(c => ({ lat: c[1], lng: c[0] }));

            let color = '#2563eb';
            if (table.blockId !== undefined) {
                if (stateColorize) {
                    let hash = 0;
                    const str = table.blockId.toString();
                    for (let i = 0; i < str.length; i++) { hash = str.charCodeAt(i) + ((hash << 5) - hash); }
                    color = blockColors[Math.abs(hash) % blockColors.length];
                } else {
                    color = '#2563eb'; // Default blue
                }
            }

            if (effectiveRes === 'modules') {
                const mx = parseInt(mount.mx) || 1;
                const my = parseInt(mount.my) || 1;
                const modDistX = parseFloat(mount.modDistX) || 0;
                const modDistY = parseFloat(mount.modDistY) || 0;
                const isPortrait = mount.orient === 'Portrait';

                const mX = isPortrait ? pv.w : pv.h;
                const mY = isPortrait ? pv.h : pv.w;

                const centerPt = table.center;

                // Create a completely invisible base JUST for tooltip hovering 
                // in case the user hovers exactly in the gap.
                const tableBase = new google.maps.Polygon({
                    paths: tableCoords,
                    strokeOpacity: 0,
                    fillOpacity: 0,
                    map: SE.map,
                    zIndex: 499
                });
                tableBase.addListener('mouseover', (e) => {
                    if (table.blockId) {
                        tooltip.innerHTML = `<b style="color: #0f172a;">${table.blockId}</b> | <span style="color: #64748b;">Table</span>`;
                        tooltip.style.display = 'block';
                    }
                });
                tableBase.addListener('mousemove', (e) => {
                    tooltip.style.left = (e.domEvent.clientX + 15) + 'px';
                    tooltip.style.top = (e.domEvent.clientY + 15) + 'px';
                });
                tableBase.addListener('mouseout', () => { tooltip.style.display = 'none'; });
                overlays.push(tableBase);

                // Draw Individual Modules at exact physical scale with true gaps
                for (let ix = 0; ix < mx; ix++) {
                    for (let iy = 0; iy < my; iy++) {
                        // Offset: exact placement in the grid
                        const offsetX = (ix - (mx - 1) / 2) * (mX + modDistX);
                        const offsetY = -(iy - (my - 1) / 2) * (mY + modDistY);

                        const modCenter = this._getMetricOffset(centerPt, offsetX, offsetY, az).geometry.coordinates;
                        // Draw module with its true physical dimensions
                        const modPoly = this._createTablePolygon(modCenter, mX, mY, az);
                        const modCoords = modPoly.geometry.coordinates[0].map(c => ({ lat: c[1], lng: c[0] }));

                        const modulePolygon = new google.maps.Polygon({
                            paths: modCoords,
                            strokeColor: '#38bdf8', strokeOpacity: 0.3, strokeWeight: 0.5,
                            fillColor: color, fillOpacity: 0.9, 
                            map: SE.map,
                            zIndex: 501
                        });

                        modulePolygon.addListener('mouseover', (e) => {
                            modulePolygon.setOptions({ strokeColor: '#fff', strokeOpacity: 1, strokeWeight: 2 });
                            if (table.blockId) {
                                tooltip.innerHTML = `<b>${table.blockId}</b> | Row ${iy+1}, Mod ${ix+1}`;
                                tooltip.style.display = 'block';
                            }
                        });
                        modulePolygon.addListener('mousemove', (e) => {
                            tooltip.style.left = (e.domEvent.clientX + 15) + 'px';
                            tooltip.style.top = (e.domEvent.clientY + 15) + 'px';
                        });
                        modulePolygon.addListener('mouseout', () => {
                            modulePolygon.setOptions({ strokeColor: '#38bdf8', strokeOpacity: 0.2, strokeWeight: 0.5 });
                            tooltip.style.display = 'none';
                        });

                        overlays.push(modulePolygon);
                    }
                }
            } else {
                const polygon = new google.maps.Polygon({
                    paths: tableCoords,
                    strokeColor: '#000', strokeOpacity: 0.3, strokeWeight: 1,
                    fillColor: color, fillOpacity: 0.9,
                    map: SE.map,
                    zIndex: 500
                });

                // Table Hover Listeners (Modern Lightweight Tooltip)
                polygon.addListener('mouseover', (e) => {
                    if (table.blockId) {
                        polygon.setOptions({ strokeOpacity: 0.8, strokeWeight: 2 });
                        tooltip.innerHTML = `<b style="color: #0f172a;">${table.blockId}</b> | <span style="color: #64748b;">${table.blockMW.toFixed(2)} MWp</span>`;
                        tooltip.style.display = 'block';
                    }
                });
                polygon.addListener('mousemove', (e) => {
                    tooltip.style.left = (e.domEvent.clientX + 15) + 'px';
                    tooltip.style.top = (e.domEvent.clientY + 15) + 'px';
                });
                polygon.addListener('mouseout', () => {
                    polygon.setOptions({ strokeOpacity: 0.3, strokeWeight: 1 });
                    tooltip.style.display = 'none';
                });

                overlays.push(polygon);
            }
        });

        // DRAW ELECTRICAL EQUIPMENT (Stations/Inverters)
        const areaBlocks = [];
        const areaInverters = [];
        if (plan.equipment) {
            plan.equipment.forEach(item => {
                const center = item.center;
                const azimuth = (item.azimuth !== undefined) ? item.azimuth : state.mounting.azimuth;
                const poly = this._createTablePolygon(center, item.w, item.l, azimuth);
                const coords = poly.geometry.coordinates[0].map(c => ({ lat: c[1], lng: c[0] }));

                const equipmentPoly = new google.maps.Polygon({
                    paths: coords,
                    strokeColor: '#334155', strokeOpacity: 1, strokeWeight: 2,
                    fillColor: item.type === 'MV Station' ? '#64748b' : '#334155',
                    fillOpacity: 1,
                    map: SE.map,
                    zIndex: 1000 // Above tables
                });

                equipmentPoly.isBlock = (item.type === 'MV Station');
                equipmentPoly.item = item;
                equipmentPoly.area = area; // Tag with area for CablesEngine

                if (equipmentPoly.isBlock) {
                    areaBlocks.push(equipmentPoly);
                } else if (item.type === 'Inverter') {
                    areaInverters.push(equipmentPoly);
                }

                equipmentPoly.addListener('mouseover', (e) => {
                    equipmentPoly.setOptions({ strokeColor: '#0f172a', strokeWeight: 3 });
                    let infoHtml = `<b style="color: #0f172a;">${item.type}${item.id ? ': ' + item.id : ''}</b>`;
                    if (item.rating) infoHtml += `<br><span style="color: #64748b; font-weight: 600;">Rating: ${item.rating}</span>`;
                    if (item.type === 'Inverter') {
                        infoHtml += `<br><span style="color: #64748b;">${item.dcKw} kWp DC | ${item.acKw} kW AC</span>`;
                    }
                    infoHtml += `<br><small style="color: #94a3b8;">${item.blockId}</small>`;

                    tooltip.innerHTML = infoHtml;
                    tooltip.style.display = 'block';
                });
                equipmentPoly.addListener('mousemove', (e) => {
                    tooltip.style.left = (e.domEvent.clientX + 15) + 'px';
                    tooltip.style.top = (e.domEvent.clientY + 15) + 'px';
                });
                equipmentPoly.addListener('mouseout', () => {
                    equipmentPoly.setOptions({ strokeColor: '#334155', strokeWeight: 2 });
                    tooltip.style.display = 'none';
                });

                overlays.push(equipmentPoly);
            });
        }

        // Store with metadata
        this.layoutStore.set(area, {
            overlays: overlays,
            blocks: areaBlocks,
            inverters: areaInverters
        });
    },



    /**
     * Applies the selected configuration to the system state
     */
    applyBlockConfiguration(data) {
        console.log("Applying Block Design:", data);
        // We still store these values in a way that _gatherSystemState can find them
        // If the old elements don't exist, we might need to store them in a persistent state object
        window._appliedRotation = data; // Cache for next layout run

        // Update the DC/AC ratio in main UI
        const dcAcInput = document.getElementById('inv-acdc-ratio');
        if (dcAcInput) dcAcInput.value = data.ratio;

        // Force variables that layout logic uses
        // Since we removed the inputs, we'll store them in a global config object or similar
        this.currentBlockConfig = data;

        // Update physical layout suggestions
        const modsPerTable = data.modsPerStr * (parseInt(document.getElementById('mount-mod-table-y')?.value) || 2);
        const totalStationMods = data.modsPerStr * data.strsPerInv * data.invsPerBlock;
        const tablesPerStation = Math.ceil(totalStationMods / modsPerTable);

        // Calculate physical table width to enforce 200m DC cable limit per block row
        const modW = (parseFloat(document.getElementById('pv-width')?.value) || 1134) / 1000; // meters
        const modH = (parseFloat(document.getElementById('pv-height')?.value) || 2278) / 1000; // meters
        const modDistX = parseFloat(document.getElementById('mount-mod-dist-x')?.value) || 0;
        const isPortrait = document.querySelector('#mounting-orientation-seg .seg-btn.active')?.dataset.value === 'Portrait';

        const tableW = isPortrait ?
            (modW * data.modsPerStr + (data.modsPerStr - 1) * modDistX) :
            (modH * data.modsPerStr + (data.modsPerStr - 1) * modDistX);

        const tableDistX = parseFloat(document.getElementById('mount-table-dist-x')?.value) || 0.5;

        const maxBx = Math.max(1, Math.floor((200 + tableDistX) / (tableW + tableDistX)));

        let bX = Math.ceil(Math.sqrt(tablesPerStation));
        if (bX > maxBx) bX = maxBx;

        const bY = Math.ceil(tablesPerStation / bX);
        window._lastTablesPerBlock = { x: bX, y: bY };
    },

    /**
     * Helper to add a DS or POC if user forgot
     */
    _ensureEssentialEquipment(state, hasDS, hasPOC, expectedDsName, isFirstArea) {
        const SE = window.SiteEngine;
        if (!SE || !state.primaryArea) return;

        const areaGeo = SE.getGeoJSON(state.primaryArea);
        if (!areaGeo) return;

        const bbox = turf.bbox(areaGeo); // [minX, minY, maxX, maxY]
        const yMid = (bbox[1] + bbox[3]) / 2;
        const xRight = bbox[2];

        // Placement point at middle-right of area
        const dsPos = [xRight, yMid];
        // POC point (approx 50m further right)
        const pocPos = turf.destination(dsPos, 0.05, 90, { units: 'kilometers' }).geometry.coordinates;

        if (!hasDS) {
            const newDs = this._spawnEquipmentMarkup(dsPos, 'station', expectedDsName, '#ef4444', true);
            if (newDs && state.primaryArea) {
                state.primaryArea.linkedDsId = newDs.__uid; // Auto-link
            }
        }
        
        // Only spawn POC for the very first project area
        if (!hasPOC && isFirstArea) {
            this._spawnEquipmentMarkup(pocPos, 'poc', 'POC 1', '#286944', true);
        }
    },

    /**
     * Internal geometry builder for DS/POC
     */
    _spawnEquipmentMarkup(pos, subType, name, color, skipSelect = false) {
        const SE = window.SiteEngine;
        let w, h;

        if (subType === 'poc') {
            w = parseFloat(document.getElementById('poc-width')?.value) || 30;
            h = parseFloat(document.getElementById('poc-height')?.value) || 15;
        } else {
            w = parseFloat(document.getElementById('ds-width')?.value) || 30;
            h = parseFloat(document.getElementById('ds-height')?.value) || 15;
        }

        const mToDeg = 1 / 111320;
        const cosLat = Math.cos(pos[1] * Math.PI / 180);
        const dLat = (h / 2) * mToDeg;
        const dLng = (w / 2) * mToDeg / cosLat;

        const coords = [
            { lat: pos[1] + dLat, lng: pos[0] - dLng },
            { lat: pos[1] + dLat, lng: pos[0] + dLng },
            { lat: pos[1] - dLat, lng: pos[0] + dLng },
            { lat: pos[1] - dLat, lng: pos[0] - dLng }
        ];

        const poly = new google.maps.Polygon({
            paths: coords,
            fillColor: color,
            fillOpacity: 0.4,
            strokeColor: color,
            strokeWeight: 2,
            map: SE.map,
            editable: true,
            draggable: true,
            zIndex: 1500 // Topmost for grid equipment
        });

        if (!poly.__uid && SE.generateUid) poly.__uid = SE.generateUid();
        poly.category = 'poc';
        poly.subType = subType;
        poly.areaName = String(name);
        poly.type = google.maps.drawing.OverlayType.POLYGON;

        if (skipSelect) {
            SE.overlays.push(poly);
            SE.attachOverlayListeners(poly);
            // Ensure label is created with string name
            setTimeout(() => SE.updateAreaLabel(poly), 50);
        } else {
            SE.processNewOverlay(poly);
        }
        
        return poly;
    }
};

window.LayoutEngine = LayoutEngine;
