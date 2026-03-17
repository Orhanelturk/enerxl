window.SldEngine = {
    canvas: null,
    ctx: null,
    scale: 1,
    panX: 0,
    panY: 0,
    isDragging: false,
    dragStartX: 0,
    dragStartY: 0,
    type: 'mvhv',

    init() {
        this.canvas = document.getElementById('sld-canvas');
        if (!this.canvas) return;
        this.ctx = this.canvas.getContext('2d');

        // UI Controls
        document.getElementById('sld-type-select')?.addEventListener('change', (e) => {
            this.type = e.target.value;
            this.renderSld();
        });
        document.getElementById('sld-zoom-in')?.addEventListener('click', () => this.zoom(1.2));
        document.getElementById('sld-zoom-out')?.addEventListener('click', () => this.zoom(1 / 1.2));
        document.getElementById('sld-zoom-reset')?.addEventListener('click', () => {
            this.autoFit();
        });

        const btnBackMv = document.getElementById('btn-back-mv');
        if (btnBackMv) {
            btnBackMv.addEventListener('click', () => {
                const sldSelect = document.getElementById('sld-type-select');
                if (sldSelect) sldSelect.value = 'mvhv';
                this.type = 'mvhv';
                this.renderSld();
            });
        }

        // Mouse Pan/Zoom
        const viewport = document.getElementById('sld-viewport');
        if (viewport) {
            viewport.addEventListener('wheel', (e) => {
                e.preventDefault();
                const viewportRect = viewport.getBoundingClientRect();
                const mouseX = e.clientX - viewportRect.left;
                const mouseY = e.clientY - viewportRect.top;

                const zoomFactor = e.deltaY > 0 ? 0.9 : 1.1;
                this.zoom(zoomFactor, mouseX, mouseY);
            });

            viewport.addEventListener('mousedown', (e) => {
                this.isDragging = true;
                this.dragHasMoved = false;
                this.dragStartX = e.clientX - this.panX;
                this.dragStartY = e.clientY - this.panY;
                viewport.style.cursor = 'grabbing';
            });

            window.addEventListener('mousemove', (e) => {
                if (!this.isDragging) return;
                if (Math.abs(e.clientX - this.dragStartX - this.panX) > 2 || Math.abs(e.clientY - this.dragStartY - this.panY) > 2) {
                    this.dragHasMoved = true;
                }
                this.panX = e.clientX - this.dragStartX;
                this.panY = e.clientY - this.dragStartY;
                this.applyTransform();
            });

            window.addEventListener('mouseup', () => {
                this.isDragging = false;
                viewport.style.cursor = 'grab';
            });

            viewport.addEventListener('click', (e) => {
                if (this.dragHasMoved) return; // Prevent triggering click after dragging

                const viewportRect = viewport.getBoundingClientRect();
                const rawX = e.clientX - viewportRect.left;
                const rawY = e.clientY - viewportRect.top;

                const clickX = (rawX - this.panX) / this.scale;
                const clickY = (rawY - this.panY) / this.scale;

                if (this.hitZones) {
                    for (let i = this.hitZones.length - 1; i >= 0; i--) {
                        let hz = this.hitZones[i];
                        if (clickX >= hz.x && clickX <= hz.x + hz.w &&
                            clickY >= hz.y && clickY <= hz.y + hz.h) {

                            if (hz.type === 'mv-block') {
                                const sldSelect = document.getElementById('sld-type-select');
                                if (sldSelect) sldSelect.value = 'lv';
                                this.type = 'lv';
                                this.lvTargetBlock = hz.id;
                                this.renderSld();
                            }
                            break;
                        }
                    }
                }
            });

            viewport.addEventListener('mousemove', (e) => {
                const viewportRect = viewport.getBoundingClientRect();
                const hoverX = (e.clientX - viewportRect.left - this.panX) / this.scale;
                const hoverY = (e.clientY - viewportRect.top - this.panY) / this.scale;
                let isHoveringHitZone = false;

                if (this.hitZones && !this.isDragging) {
                    for (let hz of this.hitZones) {
                        if (hoverX >= hz.x && hoverX <= hz.x + hz.w &&
                            hoverY >= hz.y && hoverY <= hz.y + hz.h) {
                            isHoveringHitZone = true;
                            break;
                        }
                    }
                    viewport.style.cursor = isHoveringHitZone ? 'pointer' : 'grab';
                }
            });
        }
    },

    zoom(factor, mouseX = null, mouseY = null) {
        const oldScale = this.scale;
        this.scale *= factor;
        this.scale = Math.max(0.2, Math.min(this.scale, 5)); // limits

        if (mouseX !== null && mouseY !== null) {
            const scaleChange = this.scale - oldScale;
            this.panX -= (mouseX - this.panX) * (scaleChange / oldScale);
            this.panY -= (mouseY - this.panY) * (scaleChange / oldScale);
        }
        this.applyTransform();
    },

    applyTransform() {
        const workspace = document.getElementById('sld-workspace');
        if (workspace) {
            workspace.style.transform = `translate(${this.panX}px, ${this.panY}px) scale(${this.scale})`;
        }
    },

    autoFit() {
        const viewport = document.getElementById('sld-viewport');
        if (!viewport || !this.canvas) return;
        const logicalW = parseFloat(this.canvas.dataset.logicalW) || this.canvas.width;
        const logicalH = parseFloat(this.canvas.dataset.logicalH) || this.canvas.height;

        const vpW = viewport.clientWidth;
        const vpH = viewport.clientHeight;

        if (logicalW > 0 && logicalH > 0 && vpW > 0) {
            const scaleX = vpW / (logicalW + 100);
            const scaleY = vpH / (logicalH + 100);

            this.scale = Math.max(0.2, Math.min(Math.min(scaleX, scaleY), 1.5));
            this.panX = (vpW - (logicalW * this.scale)) / 2;
            this.panY = Math.max(0, (vpH - (logicalH * this.scale)) / 2);
            this.applyTransform();
        }
    },

    renderSld() {
        if (!this.ctx) this.init();
        document.getElementById('sld-loading').style.display = 'block';

        // Hide canvas temporarily while loading
        this.canvas.style.opacity = '0';

        setTimeout(() => {
            document.getElementById('sld-loading').style.display = 'none';
            this.canvas.style.opacity = '1';

            this.ctx.setTransform(1, 0, 0, 1, 0, 0);
            this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
            this.hitZones = [];

            if (this.type === 'lv') {
                this.drawLVSld();
            } else {
                this.drawMvHvSld();
            }

            const btnBackMv = document.getElementById('btn-back-mv');
            const typeSelect = document.getElementById('sld-type-select');
            if (btnBackMv && typeSelect) {
                btnBackMv.style.display = this.type === 'lv' ? 'flex' : 'none';
                typeSelect.style.display = this.type === 'lv' ? 'none' : 'flex';
            }

            // Let canvas size resolve layout engine, then autoFit
            setTimeout(() => this.autoFit(), 10);
        }, 300);
    },

    drawRect(x, y, w, h, fill, stroke = '#1e293b', width = 2) {
        this.ctx.fillStyle = fill;
        this.ctx.strokeStyle = stroke;
        this.ctx.lineWidth = width;
        this.ctx.fillRect(x, y, w, h);
        if (stroke) this.ctx.strokeRect(x, y, w, h);
    },

    drawLine(x1, y1, x2, y2, color = '#1e293b', width = 2) {
        this.ctx.beginPath();
        this.ctx.moveTo(x1, y1);
        this.ctx.lineTo(x2, y2);
        this.ctx.strokeStyle = color;
        this.ctx.lineWidth = width;
        this.ctx.stroke();
    },

    drawText(text, x, y, font = '12px sans-serif', color = '#1e293b', align = 'center') {
        if (!text) return;
        this.ctx.font = font;
        this.ctx.fillStyle = color;
        this.ctx.textAlign = align;
        this.ctx.textBaseline = 'middle';

        // Handle multiline
        const lines = String(text).split('\n');
        const sizeMatch = font.match(/\d+/);
        const fontSize = sizeMatch ? parseInt(sizeMatch[0], 10) : 12;
        const lineHeight = fontSize * 1.2;

        lines.forEach((line, i) => {
            if (line.trim() !== '') {
                this.ctx.fillText(line, x, y + (i * lineHeight));
            }
        });
    },

    drawCB(x, y, size = 4.5, isClosed = true) {
        this.ctx.strokeStyle = '#1e293b'; // Blackish
        this.ctx.lineWidth = 2;

        if (isClosed) {
            this.ctx.beginPath();
            this.ctx.moveTo(x - size, y - size);
            this.ctx.lineTo(x + size, y + size);
            this.ctx.moveTo(x + size, y - size);
            this.ctx.lineTo(x - size, y + size);
            this.ctx.stroke();
        }
    },

    drawTransformer(x, y, r = 20, is3Winding = false) {
        this.ctx.strokeStyle = '#1e293b';
        this.ctx.lineWidth = 2;

        if (!is3Winding) {
            // Primary winding
            this.ctx.beginPath();
            this.ctx.arc(x, y - r / 1.5, r, 0, Math.PI * 2);
            this.ctx.stroke();

            // Secondary winding
            this.ctx.beginPath();
            this.ctx.arc(x, y + r / 1.5, r, 0, Math.PI * 2);
            this.ctx.stroke();
        } else {
            // 3 Winding: Less intersection area (circles pushed further apart)

            // Top circle
            this.ctx.beginPath();
            this.ctx.arc(x, y - r * 0.7, r, 0, Math.PI * 2);
            this.ctx.stroke();

            // Bottom left
            this.ctx.beginPath();
            this.ctx.arc(x - r * 0.6, y + r * 0.5, r, 0, Math.PI * 2);
            this.ctx.stroke();

            // Bottom right
            this.ctx.beginPath();
            this.ctx.arc(x + r * 0.6, y + r * 0.5, r, 0, Math.PI * 2);
            this.ctx.stroke();
        }
    },

    drawDashedRect(x, y, w, h, fill, stroke = '#1e293b', width = 2) {
        if (fill) {
            this.ctx.fillStyle = fill;
            this.ctx.fillRect(x, y, w, h);
        }
        this.ctx.strokeStyle = stroke;
        this.ctx.lineWidth = width;
        this.ctx.setLineDash([6, 6]);
        this.ctx.strokeRect(x, y, w, h);
        this.ctx.setLineDash([]); // Reset
    },

    drawInverterSymbol(x, y, size = 20) {
        this.ctx.strokeStyle = '#d97706';
        this.ctx.lineWidth = 2;
        this.ctx.fillStyle = '#fef3c7';

        this.ctx.beginPath();
        this.ctx.rect(x - size, y - size, size * 2, size * 2);
        this.ctx.fill();
        this.ctx.stroke();

        this.ctx.beginPath();
        this.ctx.moveTo(x - size, y + size);
        this.ctx.lineTo(x + size, y - size);
        this.ctx.stroke();

        // AC symbol top-left
        this.ctx.font = "bold 14px sans-serif";
        this.ctx.fillStyle = '#d97706';
        this.ctx.textAlign = 'center';
        this.ctx.textBaseline = 'middle';
        this.ctx.fillText("~", x - size / 2, y - size / 2);

        // DC symbol bottom-right
        this.ctx.fillText("=", x + size / 2, y + size / 2 + 2);
    },

    drawLVSld() {
        const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
        const lvmvMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
        const lvV = parseFloat(document.getElementById('inv-vout')?.value) || 400;
        const invPnom = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
        const acdc = parseFloat(document.getElementById('inv-acdc-ratio')?.value) || 1.2;

        // Calculate the theoretical number of inverters specifically for *this* MV station block
        // based on the station's total capacity, not the entire site capacity.
        const invQty = Math.ceil((lvmvMva * 1000) / invPnom / acdc);

        const panelQty = parseInt(document.getElementById('tr-sum-panel-tr')?.innerText) || 2;
        const strPerInv = parseInt(document.getElementById('wiring-str-inv')?.value) || 20;
        const invPerPanel = Math.ceil(invQty / panelQty);

        const minInvWidth = 80;
        const panelW = Math.max(350, invPerPanel * minInvWidth + 40);
        const totalW = Math.max(1400, panelQty * panelW + 200);
        const totalH = 1000;

        // High DPI Setup
        const RES_MULT = 3;
        this.canvas.dataset.logicalW = totalW;
        this.canvas.dataset.logicalH = totalH;
        this.canvas.style.width = totalW + 'px';
        this.canvas.style.height = totalH + 'px';
        this.canvas.width = totalW * RES_MULT;
        this.canvas.height = totalH * RES_MULT;

        this.ctx.setTransform(1, 0, 0, 1, 0, 0);
        this.ctx.scale(RES_MULT, RES_MULT);

        // draw background
        this.ctx.fillStyle = '#ffffff';
        this.ctx.fillRect(0, 0, totalW, totalH);

        const targetName = this.lvTargetBlock ? `MV Station ${this.lvTargetBlock}` : 'Typical MV Station';
        this.drawText(`Low Voltage SLD - ${targetName}`, totalW / 2, 40, "bold 24px sans-serif");

        // draw RMU / MV at the top
        const rmX = totalW / 2;
        const rmY = 120;
        this.drawRect(rmX - 80, rmY - 30, 160, 60, '#f1f5f9', '#1e293b', 2);
        this.drawText("Ring Main Unit (RMU)", rmX, rmY - 10, "bold 14px sans-serif");
        this.drawText(`${mvV} kV In/Out`, rmX, rmY + 10, "12px sans-serif", "#64748b");

        // drop down to MV/LV transformer
        const is3Winding = document.getElementById('tr-winding-type')?.value === '3';
        const trY = rmY + 120;
        this.drawLine(rmX, rmY + 30, rmX, trY - 50, '#1e293b', 3);
        this.drawTransformer(rmX, trY, 30, is3Winding);
        this.drawText(`${lvmvMva} MVA`, rmX + 45, trY - 10, "bold 12px sans-serif", "#f59e0b", "left");
        this.drawText(`${mvV}/${lvV / 1000}${is3Winding ? '/' + (lvV / 1000) : ''} kV`, rmX + 45, trY + 10, "12px sans-serif", "#64748b", "left");

        // split to LV panels
        const panelY = trY + 200;
        const invY = panelY + 200;

        const panelSpread = (panelQty - 1) * panelW;
        let startPanelX = rmX - (panelSpread / 2);

        for (let p = 0; p < panelQty; p++) {
            let px = panelQty === 1 ? rmX : startPanelX + (p * panelW);

            // Correct drawing origin off 2-Winding vs 3-Winding
            let trOutX = rmX;
            let trOutY = trY + 50; // Default 2 winding outgoing
            if (is3Winding) {
                // 3 Winding: 30 is the TR radius. 30*0.5 + 30
                let isLeftWinding = p < (panelQty / 2);
                trOutX = isLeftWinding ? (rmX - 18) : (rmX + 18);
                trOutY = trY + 45;
            }

            // Route from TR to panel center top
            let routeY = trY + 100;
            this.drawLine(trOutX, trOutY, trOutX, routeY);
            this.drawLine(trOutX, routeY, px, routeY);
            this.drawLine(px, routeY, px, panelY);

            // Incoming CB
            this.drawCB(px, panelY - 50);
            this.drawText("Main CB", px + 10, panelY - 50, "10px sans-serif", "#1e293b", "left");

            const actualInv = Math.min(invPerPanel, invQty - (p * invPerPanel));

            // Bus and Envelope definitions
            const invSpread = (actualInv - 1) * minInvWidth;
            let startInvX = actualInv === 1 ? px : px - (invSpread / 2);
            let endInvX = actualInv === 1 ? px : px + (invSpread / 2);

            const boxTop = panelY - 80;

            // Removing Dashed boundary entirely as per user request to favor clean bus.
            this.drawText(`AC Combiner Panel ${p + 1}`, px + 20, boxTop - 15, "bold 12px sans-serif", "#0284c7", "left");
            this.drawText(`${lvV}V | ${actualInv} Inputs`, px + 20, boxTop + 15, "10px sans-serif", "#0284c7", "left");

            // Horizontal Bus
            let busLeft = actualInv === 1 ? px - 40 : startInvX - 10;
            let busRight = actualInv === 1 ? px + 40 : endInvX + 10;
            this.drawLine(busLeft, panelY, busRight, panelY, '#2563eb', 4);

            for (let i = 0; i < actualInv; i++) {
                let ix = actualInv === 1 ? px : startInvX + (i * minInvWidth);

                // Route from Bus to Inverter
                this.drawLine(ix, panelY, ix, invY - 20);
                this.drawCB(ix, panelY + 40, 3); // Outgoing CB

                if (window.CablesEngine && CablesEngine.routingGenerated) {
                    const invPnomCw = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
                    const lvVolts = parseFloat(document.getElementById('lv-v')?.value) || 400;
                    const statsLv = CablesEngine.calculateCableStats('lvac', invPnomCw / 1000, lvVolts / 1000, 0.9, 10); // assumed 10m pad dist
                    if (statsLv) {
                        this.drawText(statsLv.summaryLine1, ix + 5, panelY + 55, "7px sans-serif", "#2563eb", "left");
                        this.drawText(statsLv.summaryLine2, ix + 5, panelY + 65, "7px sans-serif", "#64748b", "left");
                    }
                }

                // Inverter Vector Symbol (Replaces basic box)
                this.drawInverterSymbol(ix, invY, 20);
                this.drawText(`INV ${i + 1}`, ix, invY - 30, "bold 10px sans-serif", "#d97706");
                this.drawText("DC/AC", ix + 25, invY, "8px sans-serif", "#d97706", "left");

                // Strings under inverter
                const strY = invY + 100;
                this.drawLine(ix, invY + 20, ix, strY - 20); // shortened to not pierce rect
                this.drawRect(ix - 25, strY - 20, 50, 40, '#dcfce7', '#16a34a', 2);
                this.drawText(`${strPerInv}x PV`, ix, strY - 8, "bold 10px sans-serif", "#16a34a");
                this.drawText("Strings", ix, strY + 7, "8px sans-serif", "#16a34a");

                if (window.CablesEngine && CablesEngine.routingGenerated) {
                    const invPnomCw = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
                    const statsDc = CablesEngine.calculateCableStats('dc', (invPnomCw / 1000) / strPerInv, 1.5, 1.0, 50); // 1.5kV string 50m avg
                    if (statsDc) {
                        this.drawText(statsDc.summaryLine1, ix + 5, invY + 30, "7px sans-serif", "#16a34a", "left");
                        this.drawText(statsDc.summaryLine2, ix + 5, invY + 40, "7px sans-serif", "#64748b", "left");
                    }
                }
            }
        }
    },

    drawMvHvSld(isSecondPass = false) {
        const EE = window.ElectricalEngine;
        const LE = window.LayoutEngine;
        
        const commonV = (EE && EE.commonBusV) ? EE.commonBusV : 132;
        const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
        const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
        
        if (pocLevel === 'LV') {
            this.drawText("Project connects at LV level.\nNo MV/HV infrastructure to display.", this.canvas.width / 2, this.canvas.height / 2, "bold 24px sans-serif", "#64748b");
            return;
        }

        // 1. Gather Areas and their Blocks
        const areaData = [];
        if (LE && LE.layoutStore) {
            LE.layoutStore.forEach((data, area) => {
                if (data.blocks && data.blocks.length > 0) {
                    areaData.push({
                        area: area,
                        name: area.areaName || "Area",
                        blocks: data.blocks
                    });
                }
            });
        }

        // Search global if layoutStore is empty (first generation fallback)
        if (areaData.length === 0 && LE && LE.blocks.length > 0) {
            areaData.push({
                area: null,
                name: "Main Area",
                blocks: LE.blocks
            });
        }

        if (areaData.length === 0) {
            this.drawText("No project areas with electrical blocks found.", this.canvas.width / 2, this.canvas.height / 2, "bold 20px sans-serif", "#64748b");
            this.drawText("Please draw an area and generate layout first.", this.canvas.width / 2, this.canvas.height / 2 + 40, "14px sans-serif", "#94a3b8");
            return;
        }

        // --- Dynamic Layout Logic for Area-Based SLD ---
        const blocksPerLoop = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1;
        
        let totalW = 400; // Base lateral padding
        let maxBlocksInAnyArea = 0;

        areaData.forEach(data => {
            if (data.blocks.length > maxBlocksInAnyArea) maxBlocksInAnyArea = data.blocks.length;
            const numLoops = Math.ceil(data.blocks.length / (blocksPerLoop || 1));
            // Ensure enough width for all loops in this area
            data.actualBusW = Math.max(400, numLoops * 120); 
            data.allocW = data.actualBusW + 300; // 300px visual safety padding between area buses to prevent bridging
            totalW += data.allocW;
        });

        const maxLoopsDeep = Math.min(blocksPerLoop, maxBlocksInAnyArea);
        const totalH = Math.max(1200, 1000 + (maxLoopsDeep * 120)); // Dynamically extend canvas depth to fit all chained stations

        // High DPI Setup
        const RES_MULT = 3;
        if (!isSecondPass) {
            this.canvas.dataset.logicalW = totalW;
            this.canvas.dataset.logicalH = totalH;
            this.canvas.style.width = totalW + 'px';
            this.canvas.style.height = totalH + 'px';
            this.canvas.width = totalW * RES_MULT;
            this.canvas.height = totalH * RES_MULT;
            this.ctx.setTransform(1, 0, 0, 1, 0, 0);
            this.ctx.scale(RES_MULT, RES_MULT);
            this.drawMvHvSld(true);
            return;
        }
        
        this.ctx.setTransform(1, 0, 0, 1, 0, 0);
        this.ctx.scale(RES_MULT, RES_MULT);
        this.ctx.fillStyle = '#ffffff';
        this.ctx.fillRect(0, 0, totalW, totalH);

        this.drawText("Electrical Single Line Diagram", totalW / 2, 50, "bold 32px sans-serif");
        this.drawText(`Common Project Collection Level: ${commonV} kV`, totalW / 2, 90, "18px sans-serif", "#64748b");

        // 2. Draw POC at the very top
        const pocX = totalW / 2;
        const pocY = 180;
        this.drawRect(pocX - 80, pocY - 30, 160, 60, '#10b981', '#059669', 2);
        this.drawText("Project POC", pocX, pocY - 8, "bold 16px sans-serif", "#fff");
        this.drawText(`${commonV} kV Bus Connection`, pocX, pocY + 12, "12px sans-serif", "#fff");

        // 3. Draw Common Bus
        const commonBusY = pocY + 180;
        const commonBusW = totalW - 400;
        this.drawLine(pocX - commonBusW/2, commonBusY, pocX + commonBusW/2, commonBusY, '#8b5cf6', 8);
        this.drawText(`COMMON BUSBAR | ${commonV} kV`, pocX - commonBusW/2, commonBusY - 25, "bold 16px sans-serif", "#8b5cf6", "left");
        
        this.drawLine(pocX, pocY + 30, pocX, commonBusY);
        this.drawCB(pocX, commonBusY - 80);
        this.drawText("Main CB", pocX + 15, commonBusY - 80, "10px sans-serif", "#1e293b", "left");

        let currentX = 200; // Start with left padding
        // 4. Draw Area-Specific Branches
        areaData.forEach((data, aIdx) => {
            const areaX = currentX + (data.allocW / 2);
            const actualBusW = data.actualBusW;
            currentX += data.allocW; // Advance X ticker for the next area
            
            // Drop from Common Bus to Area
            let branchY = commonBusY + 400;
            this.drawLine(areaX, commonBusY, areaX, branchY);
            this.drawCB(areaX, commonBusY + 80);
            this.drawText(`Fdr ${aIdx+1}`, areaX + 10, commonBusY + 80, "10px sans-serif", "#1e293b", "left");

            // If commonV > mvV, we need a Step-Up Transformer per area
            if (commonV > mvV) {
                const trY = commonBusY + 220;
                this.drawTransformer(areaX, trY, 30, false);
                this.drawText(`TR UNIT ${aIdx + 1}`, areaX + 45, trY - 10, "bold 12px sans-serif", "#f59e0b", "left");
                this.drawText(`${commonV}/${mvV} kV`, areaX + 45, trY + 10, "10px sans-serif", "#64748b", "left");
            }

            // Draw Area Bus (MV level) using isolated width
            const areaBusY = branchY;
            this.drawLine(areaX - actualBusW/2, areaBusY, areaX + actualBusW/2, areaBusY, '#2563eb', 5);
            this.drawText(`${data.name} LOCAL BUS | ${mvV} kV`, areaX - actualBusW/2, areaBusY - 20, "bold 13px sans-serif", "#2563eb", "left");

            // Draw Downstream Blocks for this Area (with Loop Support)
            const blocks = data.blocks;
            const numLoops = Math.ceil(blocks.length / (blocksPerLoop || 1));
            const spacing = actualBusW / (numLoops + 1);
            
            const lvmvMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;

            let bIdx = 0;
            for (let l = 0; l < numLoops; l++) {
                const bx = (areaX - actualBusW/2) + (spacing * (l + 1));
                let by = areaBusY + 150;
                const loopStations = Math.min(blocksPerLoop, blocks.length - bIdx);

                // Draw Feeder Circuit Breaker for this loop
                this.drawLine(bx, areaBusY, bx, by - 25);
                this.drawCB(bx, areaBusY + 50, 4);

                // Daisy-chain stations vertically
                for (let s = 0; s < loopStations; s++) {
                    const block = blocks[bIdx];

                    this.drawRect(bx - 25, by - 25, 50, 50, '#f8fafc', '#3b82f6', 2.5);
                    this.drawText(`MV${bIdx + 1}`, bx, by - 8, "bold 10px sans-serif", "#1e293b");
                    this.drawText("VIEW LV", bx, by + 12, "bold 8px sans-serif", "#3b82f6");
                    
                    // Details label slightly offset to the right so it doesn't overlap the line below
                    this.drawText(`${lvmvMva} MVA\n33/0.6kV`, bx + 35, by, "8px sans-serif", "#64748b", "left");
                    
                    // Hit detection zone for drilling down to LV SLD
                    this.hitZones.push({ 
                        type: 'mv-block', 
                        id: block.item?.id || `Block ${bIdx+1}`, 
                        x: bx - 25, y: by - 25, w: 50, h: 50 
                    });

                    // Draw cable to the next station in this loop
                    if (s < loopStations - 1) {
                        this.drawLine(bx, by + 25, bx, by + 75);
                        by += 100; // Increment Y coordinate for the next station below
                    }
                    
                    bIdx++;
                }
            }
        });
    }
};

window.addEventListener('DOMContentLoaded', () => {
    // Wait for elements
    setTimeout(() => {
        if (window.SldEngine) window.SldEngine.init();
    }, 100);
});
