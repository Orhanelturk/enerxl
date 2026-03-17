const fs = require('fs');
const file = 'c:/Users/Orhan/OneDrive - Naqaa Energy/Desktop/ENERXL/sld-engine.js';
let code = fs.readFileSync(file, 'utf8');

const prefix = code.split('    drawMvHvSld(isSecondPass = false) {')[0];

const newDrawFunction = `    drawMvHvSld(isSecondPass = false) {
        const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
        if (pocLevel === 'LV') {
            this.drawText("Project connects at LV level.\\nNo MV/HV infrastructure to display.", this.canvas.width / 2, this.canvas.height / 2, "bold 24px sans-serif", "#64748b");
            return;
        }

        // --- Fetch Engineering Data ---
        const totalStations = parseInt(document.getElementById('mv-sum-total-stations')?.innerText) || 0;
        const blocksPerLoop = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1;
        const totalBuses = parseInt(document.getElementById('mv-sum-total-buses')?.innerText) || 0;
        const stationsPerBus = parseInt(document.getElementById('mv-sum-stations-bus')?.innerText) || 0;
        const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
        const mvBusA = parseFloat(document.getElementById('mv-bus')?.value) || 630;
        const mvMbaRating = (mvV * mvBusA * 1.732) / 1000;

        const mvhvTr = parseInt(document.getElementById('mvhv-sum-total-tr')?.innerText) || 0;
        const mvhvMva = parseFloat(document.getElementById('mvhv-mva')?.value) || 100;
        const mvhvWindingType = parseInt(document.getElementById('mvhv-winding-type')?.value) || 2;

        const hvV = parseFloat(document.getElementById('hv-v')?.value) || 132;
        const useHv = document.getElementById('chk-use-hvswg')?.checked && pocLevel === 'HV';
        const hvBusA = useHv ? (parseFloat(document.getElementById('hv-bus')?.value) || 2000) : 0;
        const hvMbaRating = (hvV * hvBusA * 1.732) / 1000;

        const pocOut = parseInt(document.getElementById(useHv ? 'hv-poc-outgoing' : 'mv-poc-outgoing')?.value) || 1;

        // Actual loadings
        const sysAcMw = parseFloat(document.getElementById('sys-ac-cap')?.value) || 0;
        const sysDcMw = parseFloat(document.getElementById('sys-dc-cap')?.value) || 0;
        const actualStationAcMw = totalStations > 0 ? (sysAcMw / totalStations) : 0;
        const actualStationDcMw = totalStations > 0 ? (sysDcMw / totalStations) : 0;

        const invQty = document.getElementById('lv-sum-inv-qty')?.innerText || '-';
        const panelTr = document.getElementById('tr-sum-panel-tr')?.innerText || '-';
        const lvmvMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
        const stationLoading = lvmvMva > 0 ? (actualStationAcMw / lvmvMva * 100) : 0;

        if (totalStations === 0) {
            this.drawText("No MV Stations Generated. Check Layout.", this.canvas.width / 2, this.canvas.height / 2, "bold 20px sans-serif", "#ef4444");
            return;
        }

        // --- Layout Data Calculation (Pass 1) ---
        let trCount = mvhvTr > 0 ? mvhvTr : (totalBuses > 0 ? 1 : 0);
        let bIdx = 0;
        let trData = [];
        let totalRequiredWidth = 0;
        let maxBusHeight = 0;
        let currentStationCountLayout = 0;

        for (let t = 0; t < trCount; t++) {
            let remainBuses = totalBuses - bIdx;
            if (remainBuses <= 0) break;

            let is3W = (mvhvWindingType === 3 && remainBuses >= 2);
            let busesThisTr = is3W ? 2 : 1;

            let trInfo = { buses: [] };
            let trWidth = 0;

            for (let i = 0; i < busesThisTr; i++) {
                let b = bIdx++;
                let remainForBus = Math.min(stationsPerBus, totalStations - currentStationCountLayout);
                if (b === totalBuses - 1) remainForBus = totalStations - currentStationCountLayout;

                let loops = Math.ceil(remainForBus / blocksPerLoop);
                let busW = Math.max(250, loops * 100); 
                trWidth += busW;

                let maxLoopH = 0;
                for (let l = 0; l < loops; l++) {
                    let loopStations = Math.min(blocksPerLoop, remainForBus);
                    let h = loopStations * 70; 
                    if (h > maxLoopH) maxLoopH = h;
                    remainForBus -= loopStations;
                    currentStationCountLayout += loopStations;
                }
                if (maxLoopH > maxBusHeight) maxBusHeight = maxLoopH;

                trInfo.buses.push({
                   loopCount: loops,
                   width: busW,
                   busIdx: b
                });
            }
            if (busesThisTr === 2) {
                trWidth += 60; // gap for 3-winding
            }
            trInfo.width = Math.max(trWidth, 300);
            trData.push(trInfo);
            totalRequiredWidth += trInfo.width + 100; // spacer between TRs
        }

        totalRequiredWidth = Math.max(1200, totalRequiredWidth + 100); // add outer padding left/right

        const pocY = 80;
        let hvBusY = pocY + 150;
        let trY = hvBusY + 150;
        let busY = trY + 120;
        if (pocLevel === 'MV') {
            busY = pocY + 150;
        }

        const requiredHeight = Math.max(1000, busY + maxBusHeight + 150);

        // If first pass and size changed, resize and redraw
        if (!isSecondPass && (this.canvas.width !== totalRequiredWidth || this.canvas.height !== requiredHeight)) {
            this.canvas.width = totalRequiredWidth;
            this.canvas.height = requiredHeight;
            this.ctx.fillStyle = '#ffffff';
            this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
            this.drawMvHvSld(true);
            return;
        }

        // --- ACTUAL DRAWING --- 
        this.ctx.fillStyle = '#ffffff'; 
        this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);

        this.drawText("Project Single Line Diagram", this.canvas.width / 2, 40, "bold 28px sans-serif");

        // --- HV & POC Level ---
        if (pocLevel === 'HV' && useHv) {
            let pocSpacing = (this.canvas.width - 200) / (pocOut + 1);
            for (let p = 0; p < pocOut; p++) {
                let pX = 100 + (pocSpacing * (p + 1));
                this.drawRect(pX - 40, pocY - 20, 80, 40, '#10b981', '#059669', 2);
                this.drawText("POC", pX, pocY, "bold 14px sans-serif", "#fff");
                this.drawLine(pX, pocY + 20, pX, hvBusY); 
                this.drawCB(pX, pocY + 60);
            }

            this.drawLine(100, hvBusY, this.canvas.width - 100, hvBusY, '#8b5cf6', 6);
            this.drawText(\`HV BUS (\${hvV} kV, \${hvBusA} A)\`, 110, hvBusY - 15, "bold 14px sans-serif", "#8b5cf6", "left");
            this.drawText(\`Rated: ~\${hvMbaRating.toFixed(1)} MVA  |  Actual: \${sysAcMw.toFixed(1)} MW (\${Math.min(100, sysAcMw/hvMbaRating*100).toFixed(1)}%)\`, 110, hvBusY + 15, "12px sans-serif", "#64748b", "left");
        }

        // --- MV & TR Level ---
        let currentX = 100;
        let currentStationCount = 0;

        for (let t = 0; t < trData.length; t++) {
            let trInfo = trData[t];
            let trCenterX = currentX + trInfo.width / 2;

            if (pocLevel === 'HV') {
               this.drawLine(trCenterX, hvBusY, trCenterX, trY - 30);
               this.drawCB(trCenterX, hvBusY + 40);
               this.drawTransformer(trCenterX, trY, 20, trInfo.buses.length === 2);

               let trMw = Math.min(sysAcMw, actualStationAcMw * stationsPerBus * trInfo.buses.length);
               this.drawText(\`\${mvhvMva} MVA TR\`, trCenterX + 35, trY - 10, "bold 12px sans-serif", "#f59e0b", "left");
               this.drawText((trInfo.buses.length === 2 ? \`\${hvV}/\${mvV}/\${mvV} kV\` : \`\${hvV}/\${mvV} kV\`), trCenterX + 35, trY + 5, "10px sans-serif", "#64748b", "left");
               this.drawText(\`Act: \${trMw.toFixed(1)} MW\`, trCenterX + 35, trY + 18, "10px sans-serif", "#64748b", "left");
            }

            let busOffsetX = currentX;
            for (let i = 0; i < trInfo.buses.length; i++) {
                let busInfo = trInfo.buses[i];
                let busXStart = busOffsetX;
                let busXEnd = busOffsetX + busInfo.width;

                let b = busInfo.busIdx;
                this.drawLine(busXStart, busY, busXEnd, busY, '#2563eb', 5);
                let busMw = Math.min(sysAcMw, actualStationAcMw * stationsPerBus);

                this.drawText(\`MV BUS \${b + 1} (\${mvV} kV, \${mvBusA} A)\`, busXStart + 10, busY - 12, "bold 11px sans-serif", "#2563eb", "left");
                this.drawText(\`\${busMw.toFixed(1)} MW (\${Math.min(100, busMw/mvMbaRating*100).toFixed(1)}%)\`, busXStart + 10, busY + 12, "10px sans-serif", "#64748b", "left");

                // Connect Bus up to TR/POC
                let connectX = busXStart + busInfo.width / 2;
                if (pocLevel === 'HV') {
                    if (trInfo.buses.length === 2) {
                        this.drawLine(trCenterX, trY + 40, trCenterX, trY + 50);
                        this.drawLine(trCenterX, trY + 50, connectX, trY + 50);
                        this.drawLine(connectX, trY + 50, connectX, busY);
                        this.drawCB(connectX, busY - 20);
                    } else {
                        connectX = trCenterX; 
                        this.drawLine(trCenterX, trY + 40, connectX, busY);
                        this.drawCB(connectX, busY - 20);
                    }
                } else if (pocLevel === 'MV') {
                    this.drawRect(connectX - 30, pocY - 20, 60, 40, '#10b981', '#059669', 2);
                    this.drawText("POC", connectX, pocY, "bold 14px sans-serif", "#fff");
                    this.drawLine(connectX, pocY + 20, connectX, busY);
                    this.drawCB(connectX, busY - 40);
                }

                // Downstream MV Stations
                let remainForBus = Math.min(stationsPerBus, totalStations - currentStationCount);
                if (b === totalBuses - 1) remainForBus = totalStations - currentStationCount;

                let loops = busInfo.loopCount;
                if (loops > 0) {
                    let spacing = busInfo.width / (loops + 1);

                    for (let l = 0; l < loops; l++) {
                        let lx = busXStart + (spacing * (l + 1));
                        let loopStations = Math.min(blocksPerLoop, remainForBus);

                        this.drawLine(lx, busY, lx, busY + 20);
                        this.drawCB(lx, busY + 20);
                        const mvCbA = document.getElementById('mv-bay-cb')?.value || 630;
                        this.drawText(\`\${mvCbA}A CB\`, lx + 12, busY + 20, "10px sans-serif", "#dc2626", "left");
                        this.drawLine(lx, busY + 20, lx, busY + 50);

                        let statY = busY + 70; 
                        for (let s = 0; s < loopStations; s++) {
                            currentStationCount++;
                            this.drawRect(lx - 20, statY - 20, 40, 40, '#f1f5f9', '#3b82f6', 2);
                            this.drawText(\`MV\${currentStationCount}\`, lx, statY - 5, "bold 10px sans-serif", "#1e293b");
                            this.drawText(\`\${actualStationAcMw.toFixed(1)} MW\`, lx, statY + 8, "9px sans-serif", "#64748b");

                            this.drawText(\`Invs: \${invQty} | Pnl: \${panelTr}\\nTr: \${lvmvMva} MVA (\${stationLoading.toFixed(0)}%)\`, lx + 25, statY, "9px sans-serif", "#64748b", "left");

                            if (s < loopStations - 1) {
                                this.drawLine(lx, statY + 20, lx, statY + 60); // connect to next station
                                statY += 80; // smaller spacing
                            }
                        }
                        remainForBus -= loopStations;
                    }
                }
                busOffsetX += busInfo.width + 60; // gap for next bus inside TR
            }
            currentX += trInfo.width + 100;
        }
    }
};

window.addEventListener('DOMContentLoaded', () => {
    // Wait for elements
    setTimeout(() => {
        if (window.SldEngine) window.SldEngine.init();
    }, 100);
});
`;

fs.writeFileSync(file, prefix + newDrawFunction);
