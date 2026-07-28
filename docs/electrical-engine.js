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
    isSizingCorrect: true,

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
            const lvInvA = parseFloat(document.getElementById('lv-inv-cb')?.value) || 250;
            
            // --- REFINED SMART LV SIZING LOGIC ---
            const trTargetMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
            const trTargetAcKw = trTargetMva * 1000 * pf;
            const windingTypeVal = parseInt(document.getElementById('tr-winding-type')?.value) || 2;
            const invPActualVal = parseFloat(document.getElementById('inv-pnom')?.value) || 1;
            const invIMaxACVal = parseFloat(document.getElementById('inv-imax-ac')?.value) || 0;

            // Target distribution: 1 panel for 2-winding, 2 panels for 3-winding
            const panelsPerTrDefault = (windingTypeVal === 3) ? 2 : 1;
            const totalInvsPerTrNominal = Math.floor(trTargetAcKw / invPActualVal) || 1;
            const targetInvsPerPanel = Math.ceil(totalInvsPerTrNominal / panelsPerTrDefault);

            // Sizing Target must consider BOTH current load AND the breaker space/bus capacity
            const reqCurrentA = targetInvsPerPanel * invIMaxACVal * 1.25;
            const reqBreakerCapacityA = targetInvsPerPanel * lvInvA; 
            const reqMainCB = Math.max(reqCurrentA, reqBreakerCapacityA);

            const standardBusSizes = [630, 800, 1000, 1250, 1600, 2000, 2500, 3200, 4000, 5000, 6300];
            const bestFitMain = standardBusSizes.find(s => s >= reqMainCB) || 6300;

            const mainCBInput = document.getElementById('lv-main-cb');
            const mainCBSelector = document.querySelector('[data-target="lv-main-cb"]');
            
            if (mainCBInput) {
                if (!mainCBInput.hasAttribute('data-initialized')) {
                    mainCBInput.addEventListener('input', () => mainCBInput.setAttribute('data-manual', 'true'));
                    mainCBInput.setAttribute('data-initialized', 'true');
                }
                // Auto-update if not manual
                if (!mainCBInput.hasAttribute('data-manual')) {
                    mainCBInput.value = bestFitMain;
                    if (mainCBSelector) mainCBSelector.value = bestFitMain;
                }
            }

            const lvMainA = parseFloat(mainCBInput?.value) || bestFitMain;

            // Calculate individual CB hints
            const divisor = 1000;
            const invCwKw = (lvV * lvInvA * sqrt3 * pf) / divisor;
            const mainCwKw = (lvV * lvMainA * sqrt3 * pf) / divisor;

            const invHint = document.getElementById('lv-inv-kw-hint');
            const mainHint = document.getElementById('lv-main-kw-hint');
            if (invHint) invHint.innerText = `(~${Math.round(invCwKw)} kW)`;
            if (mainHint) mainHint.innerText = `(~${Math.round(mainCwKw)} kW)`;

            // LV Summary Data (Project-wide)
            const totalSysAcKw = (parseFloat(document.getElementById('sys-ac-cap')?.value) || 0) * 1000;
            const totalInvertersProject = Math.ceil(totalSysAcKw / invPActualVal);
            
            // Effective capacity per panel based on current CB selection
            const invPerPanelActual = lvInvA > 0 ? Math.floor(lvMainA / lvInvA) : 0;
            const totalPanelsProject = invPerPanelActual > 0 ? Math.ceil(totalInvertersProject / invPerPanelActual) : 0;

            const sumAc = document.getElementById('lv-sum-ac');
            const sumInvs = document.getElementById('lv-sum-total-inv');
            const sumInvQty = document.getElementById('lv-sum-inv-qty');
            const sumPanels = document.getElementById('lv-sum-total-panels');

            if (sumAc) sumAc.innerText = `${Math.round(mainCwKw)} kW`;
            if (sumInvs) sumInvs.innerText = totalInvertersProject;
            if (sumInvQty) sumInvQty.innerText = invPerPanelActual;
            if (sumPanels) sumPanels.innerText = totalPanelsProject;

            // Transformer Summary Updates
            const trPerProject = totalInvsPerTrNominal > 0 ? Math.ceil(totalInvertersProject / totalInvsPerTrNominal) : 1;
            const panelsPerStationActual = Math.ceil(totalPanelsProject / trPerProject);

            let totalStationsDefault = panelsPerStationActual > 0 ? Math.ceil(totalPanelsProject / panelsPerStationActual) : 1;
            const chkManualCap = document.getElementById('chk-manual-cap');
            if (this.totalLayoutStations > 0 && (!chkManualCap || !chkManualCap.checked)) {
                totalStationsDefault = this.totalLayoutStations;
            }

            const effectivePanelsPerTr = totalStationsDefault > 0 ? Math.ceil(totalPanelsProject / totalStationsDefault) : 0;
            const effectivePanelsPerWinding = (windingTypeVal === 3) ? Math.ceil(effectivePanelsPerTr / 2) : effectivePanelsPerTr;
            const effectiveTrTotalFinal = (windingTypeVal === 3) ? (effectivePanelsPerWinding * 2) : effectivePanelsPerWinding;

            const sumTrTotalPanels = document.getElementById('tr-sum-total-panels');
            const sumTrWinding = document.getElementById('tr-sum-panel-winding');
            const sumTrPanel = document.getElementById('tr-sum-panel-tr');
            const sumTrTotal = document.getElementById('tr-sum-total-stations');

            if (sumTrTotalPanels) sumTrTotalPanels.innerText = totalPanelsProject;
            if (sumTrWinding) sumTrWinding.innerText = effectivePanelsPerWinding;
            if (sumTrPanel) sumTrPanel.innerText = effectiveTrTotalFinal;
            if (sumTrTotal) sumTrTotal.innerText = totalStationsDefault;

            const totalStations = totalStationsDefault;
            const trMva = trTargetMva;


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

            // Auto-selection logic: Find optimal size to serve the explicit layout
            // Standard substation sizes
            const standardSubstationSizes = [40, 63, 80, 100, 120, 125, 160, 200, 250, 315];

            // Use the ACTUAL calculated bus load from the station map to determine capacity limits
            const sysAcMw = parseFloat(document.getElementById('sys-ac-cap')?.value) || 0;
            const actualStationMw = totalStations > 0 ? (sysAcMw / totalStations) : 0;
            const actualBusMw = Math.max(0.1, effectiveStationsPerBus * actualStationMw);

            let targetSizePerTr = actualBusMw;
            if (mvhvWindingType === 3 && totalBuses >= 2) {
                targetSizePerTr = actualBusMw * 2; // For 3-winding, we try to fit 2 buses
            }

            const mvhvBestFit = standardSubstationSizes.find(s => (s * pf) >= targetSizePerTr) || 315;

            const mvhvMvaInput = document.getElementById('mvhv-mva');
            const mvhvMvaSelector = document.querySelector('[data-target="mvhv-mva"]');

            const mvhvMva = parseFloat(mvhvMvaInput?.value) || mvhvBestFit;

            // How many buses per transformer can we actually fit safely without overloading?
            // For 3-winding, each secondary winding handles exactly (mvhvMva / 2) MVA total
            let busesPerTr = 0;
            let busesPerWinding = 0;
            let trCapMw = mvhvMva * pf;

            if (mvhvWindingType === 3) {
                let windingCap = trCapMw / 2;
                busesPerWinding = actualBusMw > 0 ? Math.floor(windingCap / actualBusMw) : 0;
                // If the transformer chosen by user is extremely tight/small, it MUST still take at least 1 Bus even if overloaded!
                if (busesPerWinding === 0) busesPerWinding = 1;
                busesPerTr = busesPerWinding * 2;
            } else {
                busesPerWinding = actualBusMw > 0 ? Math.floor(trCapMw / actualBusMw) : 0;
                // If it mathematically falls short, bind fallback to 1 bus strictly
                if (busesPerWinding === 0) busesPerWinding = 1;
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
                totalPanels: totalPanelsProject,
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
     * Automated calculation of station block sizing based on target DC/AC ratio
     * and site equipment specs.
     */
    generateOptions() {
        const pv = {
            p: parseFloat(document.getElementById('pv-pnom')?.value) || 550,
            pmaxCoeff: parseFloat(document.getElementById('pv-pmax-coeff')?.value) || -0.34
        };
        const inv = {
            p: parseFloat(document.getElementById('inv-pnom')?.value) || 125,
            nbMppt: parseInt(document.getElementById('inv-nb-mppt')?.value) || 12
        };
        const site = {
            pf: parseFloat(document.getElementById('sys-pf')?.value) || 0.9,
            targetMW: parseFloat(document.getElementById('cfg-target-mw')?.value) || 10
        };
        const stationMva = parseFloat((document.getElementById('substation-size')?.value || "9").split(';')[0]) || 9;

        const nStr = parseInt(document.getElementById('inv-mods-str')?.value) || 25;
        const targetRatio = parseFloat(document.getElementById('inv-acdc-ratio')?.value) || 1.25;

        // Target DC/AC ratio for healthy power density
        const pAcBlockKw = stationMva * 1000 * site.pf;
        const invsPerBlock = Math.max(1, Math.floor(pAcBlockKw / inv.p));
        
        const targetDcKwPerInv = inv.p * targetRatio;
        const kwPerStr = (pv.p * nStr) / 1000;
        const strsPerInv = Math.round(targetDcKwPerInv / kwPerStr);

        return [{
            modsPerStr: nStr,
            strsPerInv: Math.max(1, strsPerInv),
            invsPerBlock: Math.max(1, invsPerBlock),
            ratio: ((nStr * strsPerInv * pv.p / 1000) / inv.p).toFixed(3)
        }];
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
            'hv-v', 'hv-bus', 'sys-pf', 'inv-vout', 'inv-pnom', 'tr-winding-type', 'substation-size', 'mvhv-mva', 'mvhv-winding-type',
            'poc-v-level', 'lv-v', 'chk-use-hvswg', 'hv-poc-outgoing', 'poc-connect-type',
            'lv-poc-outgoing', 'mv-poc-outgoing', 'inv-acdc-ratio',
            'lv-cable-mat', 'mv-cable-mat', 'mv-install-type', 'hv-cable-mat', 'hv-install-type'
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
                    this.updateStringSizing();
                });
            });
        });

        // Link Modules/String to Table Along modules
        const tableX = document.getElementById('mount-mod-table-x');
        const modsStr = document.getElementById('inv-mods-str');
        if (tableX && modsStr) {
            // Initial sync
            if (!modsStr.classList.contains('user-modified')) {
                modsStr.value = tableX.value || 25;
            }
            
            tableX.addEventListener('input', () => {
                if (!modsStr.classList.contains('user-modified')) {
                    modsStr.value = tableX.value;
                    this.updateStringSizing();
                }
            });
            modsStr.addEventListener('input', () => {
                modsStr.classList.add('user-modified');
                this.updateStringSizing();
            });
        }

        // String sizing triggers
        const sizingTriggers = [
            'site-temp-min', 'site-temp-max', 
            'pv-voc', 'pv-vmpp', 'pv-voc-coeff', 'pv-pmax-coeff',
            'inv-vmax-dc', 'inv-vmpp-min', 'inv-vmpp-max',
            'inv-mods-str', 'inv-acdc-ratio', 'inv-strs-per-inv', 'inv-invs-per-block'
        ];
        sizingTriggers.forEach(id => {
            document.getElementById(id)?.addEventListener('input', () => this.updateStringSizing());
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

        this.setupSldDragging();
        this.updateStringSizing();
        this.setupAreaConfigUI();
    },

    openAreaConfigModal() {
        const modal = document.getElementById('area-cfg-modal');
        const tbody = document.getElementById('area-cfg-table-body');
        const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
        const busSelect = document.getElementById('global-bus-level-select');

        // Populate Bus Level Options based on POC Level
        busSelect.innerHTML = '';
        const lvV = document.getElementById('lv-v')?.value || 400;
        const mvV = document.getElementById('mv-v')?.value || 33;
        const hvV = document.getElementById('hv-v')?.value || 132;

        const addOpt = (val, label) => {
            const opt = document.createElement('option');
            opt.value = val;
            opt.innerText = label;
            busSelect.appendChild(opt);
        };

        if (pocLevel === 'LV') {
            addOpt('LV', lvV);
            busSelect.value = 'LV';
        } else if (pocLevel === 'MV') {
            addOpt('LV', lvV);
            addOpt('MV', mvV);
            busSelect.value = 'MV';
        } else {
            addOpt('LV', lvV);
            addOpt('MV', mvV);
            addOpt('HV', hvV);
            busSelect.value = 'HV';
        }

        // Populate DS / POC Options from SiteEngine if available
        let dsOptionsHTML = '<option value="auto">-- Auto Route --</option>';
        let pocOptionsHTML = '<option value="global">-- Global POC --</option>';
        
        let availableDS = [];
        let availablePOC = [];

        if (window.SiteEngine && window.SiteEngine.overlays) {
            window.SiteEngine.overlays.forEach(o => {
                if (o.subType === 'station') availableDS.push(o);
                if (o.subType === 'poc') availablePOC.push(o);
            });
            
            availableDS.forEach((o, idx) => {
                const name = o.areaName || `DS ${idx + 1}`;
                dsOptionsHTML += `<option value="${o.__uid || name}">${name}</option>`;
            });
            availablePOC.forEach((o, idx) => {
                const name = o.areaName || `POC ${idx + 1}`;
                pocOptionsHTML += `<option value="${o.__uid || name}">${name}</option>`;
            });
        }

        // Initialize state if empty
        window.ElectricalEngine.areaMappings = window.ElectricalEngine.areaMappings || {};

        // Loop over Layout Engine Areas
        if (tbody) {
            tbody.innerHTML = '';
            const stats = window._layoutStats || {};
            const areas = Object.keys(stats).length > 0 ? Object.keys(stats) : ['Area 1'];

            areas.forEach(areaName => {
                const mapping = window.ElectricalEngine.areaMappings[areaName] || { ds: 'auto', poc: 'global' };
                const row = document.createElement('tr');
                row.setAttribute('data-area', areaName);
                
                const dsSelect = document.createElement('select');
                dsSelect.className = 'modern-input ds-select';
                dsSelect.style.height = '36px';
                dsSelect.style.padding = '0 0.5rem';
                dsSelect.innerHTML = dsOptionsHTML;
                dsSelect.value = mapping.ds || 'auto';
                
                const pocSelect = document.createElement('select');
                pocSelect.className = 'modern-input poc-select';
                pocSelect.style.height = '36px';
                pocSelect.style.padding = '0 0.5rem';
                pocSelect.innerHTML = pocOptionsHTML;
                pocSelect.value = mapping.poc || 'global';
                
                row.innerHTML = `
                    <td style="padding: 1rem; font-weight: 600; color: #1e293b; border-bottom: 1px solid #f1f5f9;">${areaName}</td>
                    <td class="ds-td" style="padding: 0.5rem 1rem; border-bottom: 1px solid #f1f5f9;"></td>
                    <td class="poc-td" style="padding: 0.5rem 1rem; border-bottom: 1px solid #f1f5f9;"></td>
                `;
                row.querySelector('.ds-td').appendChild(dsSelect);
                row.querySelector('.poc-td').appendChild(pocSelect);
                tbody.appendChild(row);
            });
        }
        
        if (window.ElectricalEngine.globalBusLevel) {
            busSelect.value = window.ElectricalEngine.globalBusLevel;
        }

        if (modal) {
            modal.classList.remove('hidden');
            if (window.lucide) window.lucide.createIcons();
        }
    },

    setupAreaConfigUI() {
        const btnOpen = document.getElementById('btn-area-cfg-open');
        const btnSave = document.getElementById('btn-save-area-cfg');
        
        if (btnOpen) {
            btnOpen.addEventListener('click', () => this.openAreaConfigModal());
        }

        if (btnSave) {
            btnSave.addEventListener('click', () => {
                const modal = document.getElementById('area-cfg-modal');
                const tbody = document.getElementById('area-cfg-table-body');
                const busSelect = document.getElementById('global-bus-level-select');
                
                window.ElectricalEngine.areaMappings = {};
                window.ElectricalEngine.globalBusLevel = busSelect ? busSelect.value : null;
                
                if (tbody) {
                    const rows = tbody.querySelectorAll('tr[data-area]');
                    rows.forEach(row => {
                        const areaName = row.getAttribute('data-area');
                        const dsVal = row.querySelector('.ds-select').value;
                        const pocVal = row.querySelector('.poc-select').value;
                        window.ElectricalEngine.areaMappings[areaName] = { ds: dsVal, poc: pocVal };
                    });
                }
                
                if (modal) modal.classList.add('hidden');
                
                if (window.ElectricalEngine && typeof window.ElectricalEngine.updatePowerDisplay === 'function') {
                    window.ElectricalEngine.updatePowerDisplay();
                }

                if (window.SldEngine && typeof window.SldEngine.renderSld === 'function') {
                    window.SldEngine.renderSld();
                }
                
                if (window.AppAlert) {
                    window.AppAlert("Area design configurations saved successfully.", "success");
                }
            });
        }
    },

    updateStringSizing() {
        // PV Params
        const voc = parseFloat(document.getElementById('pv-voc')?.value) || 0;
        const vmpp = parseFloat(document.getElementById('pv-vmpp')?.value) || 0;
        const vocCoeff = parseFloat(document.getElementById('pv-voc-coeff')?.value) || -0.26;
        const pmaxCoeff = parseFloat(document.getElementById('pv-pmax-coeff')?.value) || -0.34; // Used for Vmpp as proxy
        
        // Site Params
        const tMin = parseFloat(document.getElementById('site-temp-min')?.value) || -5;
        const tMax = parseFloat(document.getElementById('site-temp-max')?.value) || 50;
        
        // Inverter Params
        const invVmax = parseFloat(document.getElementById('inv-vmax-dc')?.value) || 1100;
        const invVmppMin = parseFloat(document.getElementById('inv-vmpp-min')?.value) || 180;
        const invVmppMax = parseFloat(document.getElementById('inv-vmpp-max')?.value) || 1000;
        
        // Sizing
        const nMods = parseInt(document.getElementById('inv-mods-str')?.value) || 0;
        
        const statusEl = document.getElementById('string-sizing-status');
        const detailsEl = document.getElementById('string-sizing-details');
        const boxEl = document.getElementById('string-sizing-box');
        
        if (!voc || !nMods || !statusEl || !detailsEl) return;
        
        // IEC Calculations
        // Voc at Min Temp (Max Voltage)
        const vMaxString = nMods * voc * (1 + (vocCoeff * (tMin - 25)) / 100);
        // Vmpp at Max Temp (Min Voltage) - using Pmax coeff as Vmpp coeff proxy if not specific
        const vMinString = nMods * vmpp * (1 + (pmaxCoeff * (tMax - 25)) / 100);
        
        const isUpperOk = vMaxString <= invVmax;
        const isLowerOk = vMinString >= invVmppMin && vMinString <= invVmppMax;
        
        const isCorrect = isUpperOk && isLowerOk;
        this.isSizingCorrect = isCorrect;
        
        // Update UI
        if (isCorrect) {
            statusEl.innerHTML = `<i data-lucide="check-circle" style="width: 14px; height: 14px; color: #16a34a;"></i><span style="color: #16a34a;">Sizing Correct</span>`;
            if (boxEl) {
                boxEl.style.borderColor = '#bbf7d0';
                boxEl.style.background = '#f0fdf4';
            }
        } else {
            statusEl.innerHTML = `<i data-lucide="x-circle" style="width: 14px; height: 14px; color: #dc2626;"></i><span style="color: #dc2626;">Sizing Conflict</span>`;
            if (boxEl) {
                boxEl.style.borderColor = '#fecaca';
                boxEl.style.background = '#fef2f2';
            }
        }
        
        let details = `• <b>Max Voltage (Cold):</b> ${vMaxString.toFixed(1)}V (Limit ${invVmax}V) ${isUpperOk ? '✅' : '❌'}<br>`;
        details += `• <b>Min Voltage (Hot):</b> ${vMinString.toFixed(1)}V (Min ${invVmppMin}V) ${isLowerOk ? '✅' : '❌'}`;
        
        detailsEl.innerHTML = details;
        if (window.lucide) lucide.createIcons();
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
