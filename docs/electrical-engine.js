/**
 * ElectricalEngine - Advanced Inverter & MV Block Configuration Engine
 * 
 * CORE RESPONSIBILITIES:
 * 1. IEC 62548 String Sizing (Voltage boundaries, Temp coefficients)
 * 2. Combinatorial Optimization for Inverter/Block configurations
 * 3. Connection details generation (Strings -> Inverter -> MV Station)
 * 4. Compliance Verification (Max DC voltage, MPPT range, Max Power)
 */

window.ElectricalEngine = {
    isUpdating: false,

    updatePowerDisplay() {
        if (this.isUpdating) return;
        this.isUpdating = true;

        try {
            const pfField = document.getElementById('sys-pf');
            const pf = Math.max(0.1, parseFloat(pfField?.value) || 0.9);
            const sqrt3 = 1.732;

            // POC Level Logic: Tab Bar Filtering
            const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
            const tabButtons = document.querySelectorAll('[data-elec-tab]');

            tabButtons.forEach(btn => {
                const key = btn.getAttribute('data-elec-tab');
                let show = true;
                if (pocLevel === 'LV') {
                    if (['transformer', 'mvswg', 'mvhvtr', 'hvswg'].includes(key)) show = false;
                } else if (pocLevel === 'MV') {
                    if (['mvhvtr', 'hvswg'].includes(key)) show = false;
                }
                btn.style.display = show ? 'flex' : 'none';

                // If active tab gets hidden, jump back to system
                if (!show && btn.classList.contains('active')) {
                    setTimeout(() => {
                        if (btn.classList.contains('active')) {
                            document.querySelector('[data-elec-tab="system"]')?.click();
                        }
                    }, 10);
                }
            });

            // 1. Sync & Calculate Inverter Side
            const pInv = parseFloat(document.getElementById('inv-pnom')?.value) || 0;
            const vOut = parseFloat(document.getElementById('inv-vout')?.value) || 400;

            // Sync LV Voltage
            const lvVField = document.getElementById('lv-v');
            if (lvVField && lvVField.value != vOut) lvVField.value = vOut;

            // Automatic Inv CB Selection (I = P / (U * sqrt3 * PF) * 1.25 safety factor)
            if (pInv > 0 && vOut > 0 && pf > 0) {
                const iInv = (pInv * 1000) / (vOut * sqrt3 * pf);
                const reqCB = iInv * 1.25;

                // Find next standard size
                const standardSizes = [160, 200, 250, 400, 630, 800, 1000, 1250, 1600, 2000, 2500];
                const bestFit = standardSizes.find(s => s >= reqCB) || 2500;

                const cbInput = document.getElementById('lv-inv-cb');
                const cbSelector = document.querySelector('[data-target="lv-inv-cb"]');
                if (cbInput && cbInput.value != bestFit) cbInput.value = bestFit;
                if (cbSelector && cbSelector.value != bestFit) cbSelector.value = bestFit;
            }

            // 2. LV Main Calculations
            const lvV = vOut;
            const lvInvA = parseFloat(document.getElementById('lv-inv-cb')?.value) || 0;
            const lvMainA = parseFloat(document.getElementById('lv-main-cb')?.value) || 0;

            // Calculate individual CB hints
            const divisor = 1000;
            const invCwKw = (lvV * lvInvA * sqrt3 * pf) / divisor;
            const mainCwKw = (lvV * lvMainA * sqrt3 * pf) / divisor;

            const invHint = document.getElementById('lv-inv-kw-hint');
            const mainHint = document.getElementById('lv-main-kw-hint');
            if (invHint) invHint.innerText = `(~${Math.round(invCwKw)} kW)`;
            if (mainHint) mainHint.innerText = `(~${Math.round(mainCwKw)} kW)`;

            // LV Summary Data
            const totalSysAcKw = (parseFloat(document.getElementById('sys-ac-cap')?.value) || 0) * 1000;
            const invPActual = parseFloat(document.getElementById('inv-pnom')?.value) || 1;
            const totalInverters = Math.ceil(totalSysAcKw / invPActual);

            const invPerPanel = lvInvA > 0 ? Math.floor(lvMainA / lvInvA) : 0;
            const totalPanels = invPerPanel > 0 ? Math.ceil(totalInverters / invPerPanel) : 0;

            const sumAc = document.getElementById('lv-sum-ac');
            const sumInvs = document.getElementById('lv-sum-total-inv');
            const sumInvQty = document.getElementById('lv-sum-inv-qty');
            const sumPanels = document.getElementById('lv-sum-total-panels');

            if (sumAc) sumAc.innerText = `${Math.round(mainCwKw)} kW`;
            if (sumInvs) sumInvs.innerText = totalInverters;
            if (sumInvQty) sumInvQty.innerText = invPerPanel;
            if (sumPanels) sumPanels.innerText = totalPanels;

            // LV POC Updates
            const lvPocBox = document.getElementById('lv-poc-out-box');
            if (lvPocBox) {
                const isLvPoc = pocLevel === 'LV';
                lvPocBox.style.display = isLvPoc ? 'block' : 'none';
                document.querySelectorAll('#elec-tab-lv .poc-summary-row').forEach(r => r.style.display = isLvPoc ? 'flex' : 'none');

                if (isLvPoc) {
                    const lvSumInc = document.getElementById('lv-sum-incomers');
                    const lvSumOut = document.getElementById('lv-sum-outgoing');
                    if (lvSumInc) lvSumInc.innerText = totalInverters;
                    if (lvSumOut) lvSumOut.innerText = document.getElementById('lv-poc-outgoing')?.value || 1;
                }
            }

            // Transformer Summary Data
            const trMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
            const trAcKw = trMva * 1000 * pf;
            const windingType = parseInt(document.getElementById('tr-winding-type')?.value) || 2;

            const panelPerTr = mainCwKw > 0 ? Math.floor(trAcKw / mainCwKw) : 0;
            const panelPerWinding = (windingType === 3) ? Math.floor(panelPerTr / 2) : panelPerTr;
            const symmetryCap = (windingType === 3) ? (panelPerWinding * 2) : panelPerTr;

            // Priority: Actual Layout Stations (if exists) -> Theoretical Estimate
            let totalStations = symmetryCap > 0 ? Math.ceil(totalPanels / symmetryCap) : 0;
            const chkManualCap = document.getElementById('chk-manual-cap');
            if (this.totalLayoutStations > 0 && (!chkManualCap || !chkManualCap.checked)) {
                totalStations = this.totalLayoutStations;
            }

            // Distribute total panels across the final number of stations
            const effectivePanelsPerTr = totalStations > 0 ? Math.ceil(totalPanels / totalStations) : 0;
            const effectivePanelsPerWinding = (windingType === 3) ? Math.ceil(effectivePanelsPerTr / 2) : effectivePanelsPerTr;
            const effectiveTrTotal = (windingType === 3) ? (effectivePanelsPerWinding * 2) : effectivePanelsPerWinding;

            const sumTrTotalPanels = document.getElementById('tr-sum-total-panels');
            const sumTrWinding = document.getElementById('tr-sum-panel-winding');
            const sumTrPanel = document.getElementById('tr-sum-panel-tr');
            const sumTrTotal = document.getElementById('tr-sum-total-stations');

            if (sumTrTotalPanels) sumTrTotalPanels.innerText = totalPanels;
            if (sumTrWinding) sumTrWinding.innerText = effectivePanelsPerWinding;
            if (sumTrPanel) sumTrPanel.innerText = effectiveTrTotal;
            if (sumTrTotal) sumTrTotal.innerText = totalStations;

            // 3. MV Calculations
            const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
            const mvBusA = parseFloat(document.getElementById('mv-bus')?.value) || 630;
            const mvBayA = parseFloat(document.getElementById('mv-bay-cb')?.value) || 630;
            const connType = document.getElementById('mv-conn-type')?.value || 'loop';
            const blocksPerLoopInput = document.getElementById('mv-blocks-per-loop');

            // Power rating helper for MV
            const getMvMw = (amps) => (mvV * amps * sqrt3 * pf) / 1000;
            const fmtHint = (mw) => (mw < 1) ? `(~${(mw * 1000).toFixed(0)} kW)` : `(~${mw.toFixed(1)} MW)`;

            // 1. Bay CB Hint
            const hBay = document.getElementById('mv-bay-cb-hint');
            if (hBay) hBay.innerText = fmtHint(getMvMw(mvBayA));

            // Calculation based on Station (Transformer) current
            const iStation = (trMva * 1000) / (mvV * sqrt3); // On kVA rating
            // Optimal limit: Max blocks such that total current ≤ 75% of CB rating
            let maxBlocksPerLoop = mvBayA > 0 ? Math.floor((mvBayA * 0.75) / iStation) : 1;
            if (maxBlocksPerLoop < 1) maxBlocksPerLoop = 1;

            // Also cap by total stations in project
            if (totalStations > 0) maxBlocksPerLoop = Math.min(maxBlocksPerLoop, totalStations);

            if (connType === 'radial') {
                if (blocksPerLoopInput) {
                    blocksPerLoopInput.value = 1;
                    blocksPerLoopInput.disabled = true;
                }
            } else {
                if (blocksPerLoopInput) {
                    blocksPerLoopInput.disabled = false;
                    // Automatically track manual overrides
                    if (!blocksPerLoopInput.hasAttribute('data-initialized')) {
                        blocksPerLoopInput.addEventListener('input', () => blocksPerLoopInput.setAttribute('data-manual', 'true'));
                        blocksPerLoopInput.setAttribute('data-initialized', 'true');
                    }
                    
                    // Auto-set to the safe max if not manually overridden by user
                    if (!blocksPerLoopInput.hasAttribute('data-manual') || parseInt(blocksPerLoopInput.value) > maxBlocksPerLoop) {
                        blocksPerLoopInput.value = maxBlocksPerLoop;
                    }
                }
            }

            // 2. Loop Load Hint
            const blocksPerLoop = parseInt(blocksPerLoopInput?.value) || 1;
            const hLoop = document.getElementById('mv-blocks-loop-hint');
            if (hLoop) hLoop.innerText = fmtHint(getMvMw(blocksPerLoop * iStation));

            // 3. Bus Hint
            const hBus = document.getElementById('mv-bus-hint');
            if (hBus) hBus.innerText = fmtHint(getMvMw(mvBusA));

            const stationsPerBusLimit = iStation > 0 ? Math.floor(mvBusA / iStation) : 1;
            const totalBuses = stationsPerBusLimit > 0 ? Math.ceil(totalStations / stationsPerBusLimit) : 0;
            const effectiveStationsPerBus = totalBuses > 0 ? Math.ceil(totalStations / totalBuses) : 0;
            const inBaysBus = blocksPerLoop > 0 ? Math.ceil(effectiveStationsPerBus / blocksPerLoop) : 0;

            // Transformer Bay CB Selection (Default to match Bus Rating)
            const trBayCBInput = document.getElementById('mv-tr-bay-cb');
            const trBayCBSelector = document.querySelector('[data-target="mv-tr-bay-cb"]');

            // If user hasn't manually touched it, or we just want to sync it to bus as default
            const trBayBestFit = mvBusA;

            // Removed forceful override so user CAN change it manually
            // if (trBayCBInput && trBayCBInput.value != trBayBestFit) trBayCBInput.value = trBayBestFit;
            // if (trBayCBSelector && trBayCBSelector.value != trBayBestFit) trBayCBSelector.value = trBayBestFit;

            // 4. Tr Bay CB Hint
            const hTrBay = document.getElementById('mv-tr-bay-cb-hint');
            const trBayA = parseFloat(trBayCBInput?.value) || 1250;
            if (hTrBay) hTrBay.innerText = fmtHint(getMvMw(trBayA));

            // Update MV Summary
            const sumMvTotalStats = document.getElementById('mv-sum-total-stations');
            const sumMvStationsBay = document.getElementById('mv-sum-stations-bay');
            const sumMvInBaysBus = document.getElementById('mv-sum-in-bays-bus');
            const sumMvStationsBus = document.getElementById('mv-sum-stations-bus');
            const sumMvTotalBuses = document.getElementById('mv-sum-total-buses');

            if (sumMvTotalStats) sumMvTotalStats.innerText = totalStations;
            if (sumMvStationsBay) sumMvStationsBay.innerText = blocksPerLoop;
            if (sumMvInBaysBus) sumMvInBaysBus.innerText = inBaysBus;
            if (sumMvStationsBus) sumMvStationsBus.innerText = effectiveStationsPerBus;
            if (sumMvTotalBuses) sumMvTotalBuses.innerText = totalBuses;

            // MV POC Updates
            const mvPocBox = document.getElementById('mv-poc-out-box');
            if (mvPocBox) {
                const isMvPoc = pocLevel === 'MV';
                mvPocBox.style.display = isMvPoc ? 'block' : 'none';
                document.querySelectorAll('#elec-tab-mvswg .poc-summary-row').forEach(r => r.style.display = isMvPoc ? 'flex' : 'none');

                if (isMvPoc) {
                    const mvSumInc = document.getElementById('mv-sum-incomers');
                    const mvSumOut = document.getElementById('mv-sum-outgoing');
                    if (mvSumInc) mvSumInc.innerText = totalStations;
                    if (mvSumOut) mvSumOut.innerText = document.getElementById('mv-poc-outgoing')?.value || 1;
                }
            }

            // 4. MV/HV Transformer Calculations (Similar to LV/MV)
            const mvhvWindingType = parseInt(document.getElementById('mvhv-winding-type')?.value) || 2;

            // Use Active Power (MW) for sizing check to match user's benchmark (e.g., 2500A -> ~123 MW @ 33kV)
            // This ensures a 125 MVA TR is seen as capable of handling one 2500A bus section.
            const sizingBusMw = (mvV * mvBusA * sqrt3 * pf) / 1000;

            // Auto-selection logic: Find optimal size to serve the buses
            // Standard substation sizes
            const standardSubstationSizes = [40, 63, 80, 100, 120, 125, 160, 200, 250, 315];

            // Target per Transformer capacity based on bus load
            let targetSizePerTr = sizingBusMw;
            if (mvhvWindingType === 3 && totalBuses >= 2) {
                targetSizePerTr = sizingBusMw * 2; // For 3-winding, we try to fit 2 buses
            }

            const mvhvBestFit = standardSubstationSizes.find(s => s >= targetSizePerTr) || 315;

            const mvhvMvaInput = document.getElementById('mvhv-mva');
            const mvhvMvaSelector = document.querySelector('[data-target="mvhv-mva"]');

            // Removed forceful override so user CAN change the HV transformer from default
            // if (mvhvMvaInput && !mvhvMvaInput.dataset.manual) mvhvMvaInput.value = mvhvBestFit;

            const mvhvMva = parseFloat(mvhvMvaInput?.value) || mvhvBestFit;

            // How many buses per transformer can we actually fit?
            // For 3-winding, each secondary winding takes mvhvMva / 2
            let busesPerTr = 0;
            let busesPerWinding = 0;

            if (mvhvWindingType === 3) {
                busesPerWinding = sizingBusMw > 0 ? Math.floor((mvhvMva / 2) / sizingBusMw) : 0;
                // At least one bus if rating is very close (handles the 125 and 123 case)
                if (busesPerWinding === 0 && (mvhvMva / 2) >= (sizingBusMw * 0.95)) busesPerWinding = 1;
                busesPerTr = busesPerWinding * 2;
            } else {
                busesPerWinding = sizingBusMw > 0 ? Math.floor(mvhvMva / sizingBusMw) : 0;
                if (busesPerWinding === 0 && mvhvMva >= (sizingBusMw * 0.95)) busesPerWinding = 1;
                busesPerTr = busesPerWinding;
            }

            const totalMainTr = busesPerTr > 0 ? Math.ceil(totalBuses / busesPerTr) : 0;
            const effectiveBusesPerTrFinal = totalMainTr > 0 ? Math.ceil(totalBuses / totalMainTr) : 0;
            const effectiveBusesPerWinding = (mvhvWindingType === 3) ? Math.ceil(effectiveBusesPerTrFinal / 2) : effectiveBusesPerTrFinal;
            const effectiveTrTotalBus = (mvhvWindingType === 3) ? (effectiveBusesPerWinding * 2) : effectiveBusesPerWinding;

            const sumMvHvTotalBuses = document.getElementById('mvhv-sum-total-buses');
            const sumMvHvWinding = document.getElementById('mvhv-sum-bus-winding');
            const sumMvHvBusTr = document.getElementById('mvhv-sum-bus-tr');
            const sumMvHvTotalTr = document.getElementById('mvhv-sum-total-tr');

            if (sumMvHvTotalBuses) sumMvHvTotalBuses.innerText = totalBuses;
            if (sumMvHvWinding) sumMvHvWinding.innerText = effectiveBusesPerWinding;
            if (sumMvHvBusTr) sumMvHvBusTr.innerText = effectiveTrTotalBus;
            if (sumMvHvTotalTr) sumMvHvTotalTr.innerText = totalMainTr;

            // 5. HV Calculations
            const useHvSwg = document.getElementById('chk-use-hvswg')?.checked;
            const hvContainer = document.getElementById('hvswg-controls-container');
            if (hvContainer) hvContainer.style.display = useHvSwg ? 'block' : 'none';

            if (useHvSwg) {
                const isHvPoc = pocLevel === 'HV';
                const hvV = parseFloat(document.getElementById('hv-v')?.value) || 132;
                const hvBusA = parseFloat(document.getElementById('hv-bus')?.value) || 2000;

                // Only show POC outgoing in HV if level is HV
                const hvPocBox = document.querySelector('#elec-tab-hvswg .mini-grid');
                if (hvPocBox) hvPocBox.style.display = isHvPoc ? 'grid' : 'none';
                document.querySelectorAll('#elec-tab-hvswg .poc-summary-row').forEach(r => r.style.display = isHvPoc ? 'flex' : 'none');

                const hvPocOut = parseInt(document.getElementById('hv-poc-outgoing')?.value) || 1;

                const getHvMw = (amps) => (hvV * amps * sqrt3 * pf) / 1000;
                const hHvBus = document.getElementById('hv-bus-hint');
                if (hHvBus) hHvBus.innerText = fmtHint(getHvMw(hvBusA));

                const sumHvIncomers = document.getElementById('hv-sum-incomers');
                const sumHvOutgoing = document.getElementById('hv-sum-outgoing');

                if (sumHvIncomers) sumHvIncomers.innerText = totalMainTr;
                if (sumHvOutgoing) sumHvOutgoing.innerText = hvPocOut;
            }

            // Global Export Stats for Excel (BOQ)
            window._electricalStats = {
                totalPanels: totalPanels,
                totalStations: totalStations,
                totalMainTr: totalMainTr,
                totalHvIncomers: totalMainTr,
                totalHvOutgoing: (pocLevel === 'HV' && useHvSwg) ? (parseInt(document.getElementById('hv-poc-outgoing')?.value) || 1) : 0,
                totalMvOutgoing: (pocLevel === 'MV') ? (parseInt(document.getElementById('mv-poc-outgoing')?.value) || 1) : 0,
                totalLvOutgoing: (pocLevel === 'LV') ? (parseInt(document.getElementById('lv-poc-outgoing')?.value) || 1) : 0,
                voltageLevels: { lv: lvV, mv: mvV, hv: (pocLevel === 'HV') ? parseFloat(document.getElementById('hv-v')?.value) : 0 }
            };
        } catch (e) {
            console.error("Error in updatePowerDisplay:", e);
        } finally {
            this.isUpdating = false;
        }
    },

    /**
     * Generates optimal configurations based on physical equipment constraints
     */
    generateOptions() {
        // PV specs
        const pv = {
            p: parseFloat(document.getElementById('pv-pnom').value) || 550,
            voc: parseFloat(document.getElementById('pv-voc').value) || 50,
            vmpp: parseFloat(document.getElementById('pv-vmpp').value) || 42,
            vocCoeff: parseFloat(document.getElementById('pv-voc-coeff').value) || -0.26,
            pmaxCoeff: parseFloat(document.getElementById('pv-pmax-coeff').value) || -0.34
        };

        // Inverter specs
        const inv = {
            p: parseFloat(document.getElementById('inv-pnom').value) || 125,
            vmax: parseFloat(document.getElementById('inv-vmax-dc').value) || 1100,
            vmppMin: parseFloat(document.getElementById('inv-vmpp-min').value) || 200,
            nbMppt: parseInt(document.getElementById('inv-nb-mppt').value) || 12
        };

        // Site conditions
        const site = {
            tMin: parseFloat(document.getElementById('site-temp-min').value) || -5,
            tMax: parseFloat(document.getElementById('site-temp-max').value) || 50,
            pf: parseFloat(document.getElementById('sys-pf')?.value) || 0.9,
            targetMW: parseFloat(document.getElementById('cfg-target-mw').value) || 10
        };

        const stationMva = parseFloat((document.getElementById('substation-size').value || "9").split(';')[0]) || 9;

        // 1. String Sizing Logic
        const res = this.verifyStringSizing(pv, inv, site);
        const nStr = res.recommended;

        // 2. Combinatorial Search for best "Electrical Blocks"
        // We look for configurations that utilize the transformer well while maintaining healthy DC/AC ratios
        const pAcBlockKw = stationMva * 1000 * site.pf;
        const baseInvPerBlock = Math.round(pAcBlockKw / inv.p);

        const options = [];
        // Test variations: inv count near base, string count near target ratios
        const invVariations = [baseInvPerBlock, baseInvPerBlock - 1, baseInvPerBlock + 1];
        const strVariations = [Math.floor(inv.nbMppt * 1.1), inv.nbMppt, Math.ceil(inv.nbMppt * 1.5)];

        const seen = new Set();

        [1.10, 1.20, 1.30, 1.40].forEach(ratio => {
            const targetDcKwPerInv = inv.p * ratio;
            const kwPerStr = (pv.p * nStr) / 1000;
            const strings = Math.round(targetDcKwPerInv / kwPerStr);

            invVariations.forEach(invCount => {
                if (invCount <= 0) return;

                const dcKw = (nStr * strings * invCount * pv.p) / 1000;
                const acKw = invCount * inv.p;
                const actualRatio = dcKw / acKw;
                const numStations = Math.ceil((site.targetMW * 1000) / acKw);

                const totalDcKw = Math.round(dcKw * numStations);
                const totalAcKw = Math.round(acKw * numStations);

                const key = `${nStr}_${strings}_${invCount}`;
                if (!seen.has(key)) {
                    options.push({
                        modsPerStr: nStr,
                        strsPerInv: strings,
                        strsPerMppt: (strings / inv.nbMppt).toFixed(1),
                        invsPerBlock: invCount,
                        ratio: actualRatio.toFixed(3),
                        dcKw: totalDcKw, // Show Total Project DC
                        acKw: totalAcKw, // Show Total Project AC
                        numStations: numStations
                    });
                    seen.add(key);
                }
            });
        });

        // Rank by goodness of ratio (closest to average utility norms 1.2-1.3)
        return options.sort((a, b) => Math.abs(a.ratio - 1.25) - Math.abs(b.ratio - 1.25)).slice(0, 5);
    },

    verifyStringSizing(pv, inv, site) {
        // Temperature corrections
        const vocMaxT = pv.voc * (1 + (pv.vocCoeff / 100) * (site.tMin - 25));
        const tCellMax = site.tMax + 25; // Standard operational delta
        const vmppMinT = pv.vmpp * (1 + (pv.pmaxCoeff / 100) * (tCellMax - 25));

        const nMax = Math.floor(inv.vmax / vocMaxT);
        const nMin = Math.ceil(inv.vmppMin / vmppMinT);

        // Try to match Module Along (Row Length) for clean physical layout
        const preferredLen = parseInt(document.getElementById('mount-mod-table-x')?.value) || 25;
        let recommended = preferredLen;

        // If preferred length is out of electrical bounds, clamp it
        if (recommended > nMax) recommended = nMax;
        if (recommended < nMin) recommended = nMin;

        return { nMin, nMax, recommended, vocMaxT, vmppMinT };
    },

    getVerificationDetails(config) {
        // PV specs
        const pv = {
            p: parseFloat(document.getElementById('pv-pnom').value) || 550,
            voc: parseFloat(document.getElementById('pv-voc').value) || 50,
            vmpp: parseFloat(document.getElementById('pv-vmpp').value) || 42,
            vocCoeff: parseFloat(document.getElementById('pv-voc-coeff').value) || -0.26,
            pmaxCoeff: parseFloat(document.getElementById('pv-pmax-coeff').value) || -0.34
        };
        const inv = {
            p: parseFloat(document.getElementById('inv-pnom').value) || 125,
            vmax: parseFloat(document.getElementById('inv-vmax-dc').value) || 1100,
            vmppMin: parseFloat(document.getElementById('inv-vmpp-min').value) || 200
        };
        const site = {
            tMin: parseFloat(document.getElementById('site-temp-min').value) || -5,
            tMax: parseFloat(document.getElementById('site-temp-max').value) || 50
        };

        const sz = this.verifyStringSizing(pv, inv, site);
        const stringVoc = config.modsPerStr * sz.vocMaxT;
        const stringVmpp = config.modsPerStr * sz.vmppMinT;

        return {
            title: "IEC 62548 Compliance Verification",
            checks: [
                {
                    label: "Voltage Check (@ Tmin)",
                    value: `${stringVoc.toFixed(1)}V`,
                    limit: `${inv.vmax}V`,
                    status: stringVoc <= inv.vmax ? "PASS" : "FAIL",
                    detail: `Max string Voc (${config.modsPerStr} mods * ${sz.vocMaxT.toFixed(2)}V) must be <= Inverter Vmax.`
                },
                {
                    label: "MPPT Range (@ Tmax)",
                    value: `${stringVmpp.toFixed(1)}V`,
                    limit: `>${inv.vmppMin}V`,
                    status: stringVmpp >= inv.vmppMin ? "PASS" : "FAIL",
                    detail: `Min string Vmpp at ${site.tMax}°C must stay within Inverter tracking range.`
                }
            ],
            summary: `Configuration verified for ${site.tMin}°C to ${site.tMax}°C operation.`
        };
    },

    getConnectionFlow(config) {
        const schedule = this.getFullSchedule(config);
        return {
            title: "Project Electrical Schedule & Connections",
            summary: `Total Site: ${schedule.length} MV Blocks | ${config.numStations} Stations estimated for Target MW.`,
            schedule: schedule,
            steps: [
                { part: "Module Level", text: `${config.modsPerStr} PV Modules per String (${((config.modsPerStr * parseFloat(document.getElementById('pv-pnom').value)) / 1000).toFixed(2)} kWp/str).` },
                { part: "Inverter Level", text: `Each Inverter: ${config.strsPerInv} Strings | ${config.dcKw} kWp DC | ${config.acKw} kW AC.` },
                { part: "Block Level", text: `${config.invsPerBlock} Inverters per MV Block.` }
            ]
        };
    },

    getFullSchedule(config) {
        const blocks = [];
        const numBlocks = config.numStations;
        const pvP = parseFloat(document.getElementById('pv-pnom').value) || 550;
        const invP = parseFloat(document.getElementById('inv-pnom').value) || 125;
        const stationMva = parseFloat((document.getElementById('substation-size').value || "9").split(';')[0]) || 9;

        for (let b = 1; b <= numBlocks; b++) {
            const inverters = [];
            for (let i = 1; i <= config.invsPerBlock; i++) {
                inverters.push({
                    id: `INV-${b}.${i}`,
                    strings: config.strsPerInv,
                    dcKw: (config.modsPerStr * config.strsPerInv * pvP) / 1000,
                    acKw: invP
                });
            }
            blocks.push({
                id: `MV Block ${b}`,
                sizeMva: stationMva,
                invCount: config.invsPerBlock,
                totalDc: (config.modsPerStr * config.strsPerInv * config.invsPerBlock * pvP / 1000).toFixed(1),
                inverters: inverters
            });
        }
        return blocks;
    },

    initDesignUI() {
        console.log("Electrical Design UI Initializing...");

        // Tab Switching
        document.querySelectorAll('[data-elec-tab]').forEach(btn => {
            btn.addEventListener('click', () => {
                const tab = btn.getAttribute('data-elec-tab');
                document.querySelectorAll('[data-elec-tab]').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.elec-tab-content').forEach(c => c.classList.remove('active'));

                btn.classList.add('active');
                document.getElementById(`elec-tab-${tab}`).classList.add('active');
            });
        });

        // Capacity Override Logic
        const chkManual = document.getElementById('chk-manual-cap');
        const dcCap = document.getElementById('sys-dc-cap');
        const acCap = document.getElementById('sys-ac-cap');
        if (chkManual && dcCap && acCap) {
            chkManual.addEventListener('change', () => {
                dcCap.disabled = !chkManual.checked;
                acCap.disabled = !chkManual.checked;
            });
        }

        // Power Calculation Logic (Reactive/Automatic)
        const calcInputs = [
            'lv-inv-cb', 'lv-main-cb',
            'mv-v', 'mv-bus', 'mv-bay-cb', 'mv-conn-type', 'mv-blocks-per-loop', 'mv-tr-bay-cb',
            'hv-v', 'hv-bus', 'sys-pf', 'inv-vout', 'inv-pnom', 'tr-winding-type', 'substation-size', 'mvhv-mva',
            'poc-v-level', 'lv-v', 'chk-use-hvswg', 'hv-poc-outgoing', 'poc-connect-type',
            'lv-poc-outgoing', 'mv-poc-outgoing'
        ];
        calcInputs.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.addEventListener(el.tagName === 'SELECT' ? 'change' : 'input', () => this.updatePowerDisplay());
        });

        // Add listener for inverter select to trigger voltage sync
        document.getElementById('inverter-select')?.addEventListener('change', () => {
            setTimeout(() => this.updatePowerDisplay(), 50);
        });

        // Trigger on selector change too
        document.querySelectorAll('.std-selector').forEach(sel => {
            sel.addEventListener('change', () => {
                const targetId = sel.getAttribute('data-target');
                const targetInput = document.getElementById(targetId);
                if (targetInput && sel.value !== "") {
                    targetInput.value = sel.value;
                    // Programmatic change doesn't trigger 'input', so we dispatch it manually
                    targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                }
            });
        });

        this.updatePowerDisplay();

        // --- Centralized Sync Logic ---
        const syncGroups = [
            // { ids: ['id1', 'id2', ...], callback: optional }
            { ids: ['substation-size'], isMvaList: true },
            { ids: ['mv-v', 'elec-mv-v'] },
            { ids: ['grid-v', 'hv-v', 'mvhv-hv-v'] }
        ];

        syncGroups.forEach(group => {
            group.ids.forEach(id => {
                const el = document.getElementById(id);
                if (!el) return;
                el.addEventListener('input', () => {
                    let val = el.value;

                    // Specific handling for the MVA list (Eg: 9;6;3)
                    if (group.isMvaList && id === 'substation-size') {
                        val = val.split(';')[0].trim();
                    }

                    group.ids.forEach(otherId => {
                        if (otherId === id) return;
                        const otherEl = document.getElementById(otherId);
                        if (!otherEl) return;

                        if (group.isMvaList && otherId === 'substation-size') {
                            const parts = otherEl.value.split(';');
                            parts[0] = val;
                            otherEl.value = parts.join(';');
                        } else {
                            otherEl.value = val;
                        }

                        // Update any matching selectors
                        const sel = document.querySelector(`.std-selector[data-target="${otherId}"]`);
                        if (sel) {
                            const exists = Array.from(sel.options).some(o => o.value == val);
                            sel.value = exists ? val : "";
                        }
                    });

                    // Side effects
                    if (group.isMvaList && window.LayoutEngine) {
                        window.LayoutEngine.isManualConfig = false; // Reset to allow auto-suggestion
                    }
                    this.updatePowerDisplay();
                });
            });
        });

        // Panel Toggles
        const trigger = document.getElementById('electrical-sld-trigger');
        const panel = document.getElementById('electrical-design-panel');
        const closeBtn = document.getElementById('close-elec-panel');

        if (trigger) {
            trigger.addEventListener('click', () => {
                const isHidden = panel.classList.contains('hidden');
                if (typeof window.hideAllPanels === 'function') {
                    window.hideAllPanels();
                }
                if (isHidden) {
                    panel.classList.remove('hidden');
                }
            });
        }
        if (closeBtn) closeBtn.addEventListener('click', () => panel.classList.add('hidden'));

        // SLD Window - Now on Main Toolbar
        document.getElementById('btn-view-sld-main')?.addEventListener('click', () => this.toggleSldWindow(true));

        // Area Configuration Modal
        document.getElementById('btn-open-area-config')?.addEventListener('click', () => this.openAreaConfigModal());
        document.getElementById('btn-save-area-config')?.addEventListener('click', () => this.saveAreaConfig());

        this.setupSldDragging();
    },

    openAreaConfigModal() {
        const modal = document.getElementById('area-elec-config-modal');
        const tbody = document.getElementById('area-config-tbody');
        if (!modal || !tbody) return;

        const SE = window.SiteEngine;
        if (!SE) return;

        // Find all primary areas
        const areas = SE.overlays.filter(o => o.category === 'area' || (o.getPath && !o.subType && !o.category));
        const dsStations = SE.overlays.filter(o => o.subType === 'station');
        const pocs = SE.overlays.filter(o => o.subType === 'poc');

        if (areas.length === 0) {
            tbody.innerHTML = `<tr><td colspan="3" style="padding: 3rem; text-align: center; color: #94a3b8;"><i data-lucide="info" style="width: 24px; height: 24px; margin-bottom: 0.5rem; display: block; margin-inline: auto;"></i>No areas identified. Please draw project areas first.</td></tr>`;
            if (window.lucide) lucide.createIcons();
            modal.classList.remove('hidden');
            return;
        }

        // Populate Common Bus Voltage Dropdown
        const commonBusSelect = document.getElementById('common-bus-v');
        if (commonBusSelect) {
            // LV is entered/read in Volts. Divide by 1000 to convert to kV for the dropdown.
            const lv = (parseFloat(document.getElementById('lv-v')?.value) || 0) / 1000;
            const mv = parseFloat(document.getElementById('mv-v')?.value) || 33;
            const hv = parseFloat(document.getElementById('hv-v')?.value) || 132;
            const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
            const useHv = document.getElementById('chk-use-hvswg')?.checked;

            commonBusSelect.innerHTML = '';
            
            // Re-use current value or default to highest available
            const currentVal = this.commonBusV || (pocLevel === 'HV' && useHv ? hv : mv);

            const options = [];
            
            // Push valid voltages descending from the maximum allowed by the POC Level
            if (pocLevel === 'HV' && useHv && hv > 0) options.push(hv);
            if ((pocLevel === 'HV' || pocLevel === 'MV') && mv > 0) options.push(mv);
            if (lv > 0) options.push(lv);

            [...new Set(options)].forEach(v => {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = `${v} kV`;
                if (v === currentVal) opt.selected = true;
                commonBusSelect.appendChild(opt);
            });
        }

        tbody.innerHTML = '';
        areas.forEach((area, index) => {
            // Force UID assignment if missing so mapping can be saved robustly
            if (!area.__uid && SE.generateUid) area.__uid = SE.generateUid();
            
            const areaName = area.areaName || `Area ${index + 1}`;
            const linkedDsId = area.linkedDsId || '';
            const linkedPocId = area.linkedPocId || '';
            const areaUid = area.__uid;

            const row = document.createElement('tr');
            row.style.borderBottom = '1px solid #f1f5f9';
            
            // Generate DS Options
            let dsOptions = `<option value="">-- Auto Route --</option>`;
            dsStations.forEach(ds => {
                if (!ds.__uid && SE.generateUid) ds.__uid = SE.generateUid();
                const id = ds.__uid || ds.overlayId || ds.areaName;
                dsOptions += `<option value="${id}" ${linkedDsId === id ? 'selected' : ''}>${ds.areaName || 'Station'}</option>`;
            });

            // Generate POC Options
            let pocOptions = `<option value="">-- Global POC --</option>`;
            pocs.forEach(poc => {
                if (!poc.__uid && SE.generateUid) poc.__uid = SE.generateUid();
                const id = poc.__uid || poc.overlayId || poc.areaName;
                pocOptions += `<option value="${id}" ${linkedPocId === id ? 'selected' : ''}>${poc.areaName || 'POC'}</option>`;
            });

            row.innerHTML = `
                <td style="padding: 1rem; font-weight: 600; color: #1e293b;">${areaName}</td>
                <td style="padding: 0.75rem;">
                    <select class="modern-area-select area-ds-select" data-area-uid="${areaUid}">
                        ${dsOptions}
                    </select>
                </td>
                <td style="padding: 0.75rem;">
                    <select class="modern-area-select area-poc-select" data-area-uid="${areaUid}">
                        ${pocOptions}
                    </select>
                </td>
            `;
            tbody.appendChild(row);
        });

        modal.classList.remove('hidden');
        if (window.lucide) lucide.createIcons();
    },

    saveAreaConfig() {
        const SE = window.SiteEngine;
        const areas = SE.overlays.filter(o => o.category === 'area' || (o.getPath && !o.subType && !o.category));
        
        const dsSelects = document.querySelectorAll('.area-ds-select');
        const pocSelects = document.querySelectorAll('.area-poc-select');

        dsSelects.forEach(sel => {
            const uid = sel.dataset.areaUid;
            const area = areas.find(a => a.__uid === uid);
            if (area) {
                area.linkedDsId = sel.value;
            }
        });

        pocSelects.forEach(sel => {
            const uid = sel.dataset.areaUid;
            const area = areas.find(a => a.__uid === uid);
            if (area) {
                area.linkedPocId = sel.value;
            }
        });

        // Save common bus V
        const commonBusV = document.getElementById('common-bus-v').value;
        this.commonBusV = parseFloat(commonBusV) || 132;

        document.getElementById('area-elec-config-modal').classList.add('hidden');
        
        // Trigger SLD Refresh if open
        const sldWin = document.getElementById('sld-window');
        if (sldWin && !sldWin.classList.contains('hidden')) {
            if (window.SldEngine) window.SldEngine.renderSld();
        }

        alert("Area design configuration saved. This mapping will be used for Cable Routing and Single Line Diagram (SLD).");
    },

    toggleSldWindow(show) {
        const win = document.getElementById('sld-window');
        if (show) {
            win.classList.remove('hidden');
            lucide.createIcons();
            if (window.SldEngine) {
                window.SldEngine.renderSld();
            }
        } else {
            win.classList.add('hidden');
        }
    },

    setupSldDragging() {
        const win = document.getElementById('sld-window');
        const handle = document.getElementById('sld-handle');
        if (!handle || !win) return;

        let isDragging = false;
        let startX, startY, startLeft, startTop;

        handle.onmousedown = (e) => {
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            startLeft = win.offsetLeft;
            startTop = win.offsetTop;
            document.onmousemove = onDrag;
            document.onmouseup = stopDrag;
        };

        const onDrag = (e) => {
            if (!isDragging) return;
            win.style.left = (startLeft + e.clientX - startX) + "px";
            win.style.top = (startTop + e.clientY - startY) + "px";
            win.style.bottom = "auto";
            win.style.right = "auto";
        };

        const stopDrag = () => {
            isDragging = false;
            document.onmousemove = null;
            document.onmouseup = null;
        };
    }
};

// Auto-init UI
document.addEventListener('DOMContentLoaded', () => {
    ElectricalEngine.initDesignUI();
});
