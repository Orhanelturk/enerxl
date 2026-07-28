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
        if (this.isInitialized) return; // Prevent double init
        
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
            }, { passive: false });

            viewport.addEventListener('mousedown', (e) => {
                this.isDragging = true;
                this.dragHasMoved = false;
                this.dragStartX = e.clientX - this.panX;
                this.dragStartY = e.clientY - this.panY;
                this.clickStartCoords = { x: e.clientX, y: e.clientY };
                viewport.style.cursor = 'grabbing';
            });

            window.addEventListener('mousemove', (e) => {
                if (!this.isDragging) return;
                
                // Track movement for click-pan suppression (10px threshold)
                if (this.clickStartCoords) {
                    const dx = Math.abs(e.clientX - this.clickStartCoords.x);
                    const dy = Math.abs(e.clientY - this.clickStartCoords.y);
                    if (dx > 10 || dy > 10) this.dragHasMoved = true;
                }
                
                this.panX = e.clientX - this.dragStartX;
                this.panY = e.clientY - this.dragStartY;
                this.applyTransform();
            });

            window.addEventListener('mouseup', () => {
                this.isDragging = false;
                viewport.style.cursor = 'grab';
            });

            // Click Interaction for Drilling Down
            this.canvas.addEventListener('click', (e) => {
                if (this.dragHasMoved) return;

                const rect = this.canvas.getBoundingClientRect();
                const logicalW = parseFloat(this.canvas.style.width) || this.canvas.width;
                const logicalH = parseFloat(this.canvas.style.height) || this.canvas.height;
                
                // Map screen coordinates proportionally to logical drawing coordinates
                const clickX = ((e.clientX - rect.left) / rect.width) * logicalW;
                const clickY = ((e.clientY - rect.top) / rect.height) * logicalH;

                if (this.hitZones && this.hitZones.length > 0) {
                    for (let i = this.hitZones.length - 1; i >= 0; i--) {
                        const hz = this.hitZones[i];
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

            // Hover Effect
            this.canvas.addEventListener('mousemove', (e) => {
                const rect = this.canvas.getBoundingClientRect();
                const logicalW = parseFloat(this.canvas.style.width) || this.canvas.width;
                const logicalH = parseFloat(this.canvas.style.height) || this.canvas.height;
                
                const hoverX = ((e.clientX - rect.left) / rect.width) * logicalW;
                const hoverY = ((e.clientY - rect.top) / rect.height) * logicalH;
                
                let isHoveringHitZone = false;

                if (this.hitZones && this.hitZones.length > 0) {
                    for (const hz of this.hitZones) {
                        if (hoverX >= hz.x && hoverX <= hz.x + hz.w &&
                            hoverY >= hz.y && hoverY <= hz.y + hz.h) {
                            isHoveringHitZone = true;
                            break;
                        }
                    }
                }
                viewport.style.cursor = isHoveringHitZone ? 'pointer' : 'grab';
            });
        }
        this.isInitialized = true;
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

    drawText(text, x, y, font = '12px sans-serif', color = '#1e293b', align = 'center', angle = 0) {
        if (!text) return;
        this.ctx.save();
        this.ctx.font = font;
        this.ctx.fillStyle = color;
        this.ctx.textAlign = align;
        this.ctx.textBaseline = 'middle';

        this.ctx.translate(x, y);
        if (angle !== 0) {
            this.ctx.rotate((angle * Math.PI) / 180);
        }

        // Handle multiline
        const lines = String(text).split('\n');
        const sizeMatch = font.match(/\d+/);
        const fontSize = sizeMatch ? parseInt(sizeMatch[0], 10) : 12;
        const lineHeight = fontSize * 1.2;

        lines.forEach((line, i) => {
            if (line.trim() !== '') {
                this.ctx.fillText(line, 0, i * lineHeight);
            }
        });
        
        this.ctx.restore();
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

        // Default fallback nominal values
        const pf = Math.max(0.1, parseFloat(document.getElementById('sys-pf')?.value) || 0.9);
        const iInvA = (invPnom * 1000) / (lvV * 1.732 * pf);
        const trTargetAcKw = lvmvMva * 1000 * pf;
        let invQty = Math.max(1, Math.floor(trTargetAcKw / invPnom));
        const panelQty = parseInt(document.getElementById('tr-sum-panel-tr')?.innerText) || 2;
        
        const modPerStr = parseInt(document.getElementById('inv-mods-str')?.value) || 26;
        const pvP = parseFloat(document.getElementById('pv-pnom')?.value) || 550;
        const nominalStr = Math.max(1, Math.round((invPnom * acdc * 1000) / (pvP * modPerStr)));
        
        let globalStrPerInv = nominalStr; // Safely default to physical math, input doesn't exist.

        // Dynamic Actual Block Evaluation
        let actualBlockInvs = 0;
        let actualBlockStrings = 0;
        let isSpecificBlock = false;

        if (this.lvTargetBlock && window.LayoutEngine && window.LayoutEngine.layoutStore) {
            const mx = parseInt(document.getElementById('mount-mod-table-x')?.value) || 25;
            const my = parseInt(document.getElementById('mount-mod-table-y')?.value) || 2;
            const stringsPerTable = (mx * my) / modPerStr;

            window.LayoutEngine.layoutStore.forEach((data) => {
                const targetUid = String(this.lvTargetBlock).trim();
                if (data.overlays) {
                    data.overlays.forEach(o => {
                        if (o.isTable && o.tableData && String(o.tableData.blockId).trim() === targetUid) {
                            actualBlockStrings += stringsPerTable;
                        }
                    });
                }
                if (data.inverters) {
                    data.inverters.forEach(inv => {
                        if (inv.item && String(inv.item.blockId).trim() === targetUid) {
                            actualBlockInvs++;
                        }
                    });
                }
            });

            if (actualBlockInvs > 0) {
                isSpecificBlock = true;
                invQty = actualBlockInvs;
            }
        }

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
        this.drawText(targetName, totalW / 2, 20, "bold 24px sans-serif"); 

        // draw RMU Ring Switches at the top
        const rmX = totalW / 2;
        const rmY = 135; // Nudged overall RMU structure slightly further down to give title breathing room
        
        // RMU Main Common Bus
        this.drawLine(rmX - 40, rmY, rmX + 40, rmY, '#1e293b', 3);
        
        // Left Ring Switch (Incoming / To SWG)
        this.drawLine(rmX - 40, rmY, rmX - 40, rmY - 20, '#1e293b', 2); // Vertical up
        this.drawLine(rmX - 40, rmY - 20, rmX - 52, rmY - 35, '#1e293b', 2); // Switch swinging outward to the left
        this.drawLine(rmX - 40, rmY - 35, rmX - 40, rmY - 65, '#1e293b', 2); // Vertical top (collinear, length doubled per request)
        this.drawLine(rmX - 40, rmY - 65, rmX - 120, rmY - 65, '#1e293b', 2); // Feeder wire out (extended)
        
        // Filled Left Arrow head
        this.ctx.fillStyle = '#1e293b';
        this.ctx.beginPath();
        this.ctx.moveTo(rmX - 120, rmY - 65);
        this.ctx.lineTo(rmX - 110, rmY - 70);
        this.ctx.lineTo(rmX - 110, rmY - 60);
        this.ctx.closePath();
        this.ctx.fill();

        // Right Ring Switch (Outgoing / Loop Next)
        let isFinalOne = false;
        
        // Calculate Loop topology text
        let rmuPrev = "To MV SWG";
        let rmuNext = "To Ring End";
        if (this.lvTargetBlock) {
            const activeNum = parseInt(String(this.lvTargetBlock).replace('Block ', ''));
            if (!isNaN(activeNum)) {
                const bPerLoop = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 5;
                const posInLoop = (activeNum - 1) % bPerLoop;
                if (posInLoop > 0) rmuPrev = `From MV Station ${activeNum - 1}`;
                
                if (posInLoop < bPerLoop - 1) {
                    rmuNext = `To MV Station ${activeNum + 1}`;
                } else {
                    rmuNext = ""; // LEAVE EMPTY IF FINAL ONE
                    isFinalOne = true;
                }
            }
        }

        this.drawLine(rmX + 40, rmY, rmX + 40, rmY - 20, '#1e293b', 2); // Vertical up
        this.drawLine(rmX + 40, rmY - 20, rmX + 52, rmY - 35, '#1e293b', 2); // Switch swinging outward to the right
        this.drawLine(rmX + 40, rmY - 35, rmX + 40, rmY - 65, '#1e293b', 2); // Vertical top (collinear, length doubled)
        this.drawLine(rmX + 40, rmY - 65, rmX + 120, rmY - 65, '#1e293b', 2); // Feeder wire out (extended)
        
        if (!isFinalOne) {
            // Filled Right Arrow head ONLY if not the final block
            this.ctx.fillStyle = '#1e293b';
            this.ctx.beginPath();
            this.ctx.moveTo(rmX + 120, rmY - 65);
            this.ctx.lineTo(rmX + 110, rmY - 70);
            this.ctx.lineTo(rmX + 110, rmY - 60);
            this.ctx.closePath();
            this.ctx.fill();
        }
        
        this.drawText(rmuPrev, rmX - 125, rmY - 65, "10px sans-serif", "#1e293b", "right");
        if (rmuNext !== "") {
            this.drawText(rmuNext, rmX + 125, rmY - 65, "10px sans-serif", "#1e293b", "left");
        }

        // Drop down to MV/LV transformer with Protection CB
        const is3Winding = document.getElementById('tr-winding-type')?.value === '3';
        
        // TR Bay Circuit Breaker
        this.drawLine(rmX, rmY, rmX, rmY + 30, '#1e293b', 3);
        this.drawCB(rmX, rmY + 30, 4);
        const trMvCbRating = parseInt(document.getElementById('mv-tr-bay-cb')?.value) || 1250;
        this.drawText(`${trMvCbRating} A`, rmX + 15, rmY + 30, "10px sans-serif", "#1e293b", "left");
        
        const trY = rmY + 100; // Pulled TR up significantly to tighten upper CB line
        this.drawLine(rmX, rmY + 30, rmX, trY - 50, '#1e293b', 3);
        
        this.drawTransformer(rmX, trY, 30, is3Winding); // Draw the actual transformer SVG
        
        // Calculate Utilization percentage based on user's exact constraint layout
        const trAcLimits = lvmvMva * pf;
        const totalAcMw = (invQty * invPnom) / 1000;
        const utilPercent = trAcLimits > 0 ? ((totalAcMw / trAcLimits) * 100).toFixed(1) : 0;
        
        let trTextX = rmX + 45;
        if (is3Winding) trTextX += 30; // Push text explicitly further right for 3-winding
        
        this.drawText(`${lvmvMva} MVA`, trTextX, trY - 15, "bold 12px sans-serif", "#f59e0b", "left");
        this.drawText(`${mvV}/${lvV / 1000}${is3Winding ? '/' + (lvV / 1000) : ''} kV`, trTextX, trY + 5, "12px sans-serif", "#64748b", "left");
        this.drawText(`${utilPercent}%`, trTextX, trY + 20, "bold 10px sans-serif", utilPercent > 99 ? "#dc2626" : "#16a34a", "left");

        // split to LV panels
        const panelY = trY + 140; // Pulled entire Combiner infrastructure up to tighten TR lower line
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
            let routeY = trY + 70; // Adjusted split origin closer to TR
            this.drawLine(trOutX, trOutY, trOutX, routeY);
            this.drawLine(trOutX, routeY, px, routeY);
            this.drawLine(px, routeY, px, panelY);

            const actualInv = Math.min(invPerPanel, invQty - (p * invPerPanel));

            // Incoming CB
            const mainCbVal = parseInt(document.getElementById('lv-main-cb')?.value) || 2500;
            const mainCbLoad = mainCbVal > 0 ? ((actualInv * iInvA) / mainCbVal) * 100 : 0;

            this.drawCB(px, panelY - 50);
            this.drawText(`${mainCbVal} A`, px + 10, panelY - 55, "10px sans-serif", "#1e293b", "left");
            this.drawText(`${mainCbLoad.toFixed(1)}%`, px + 10, panelY - 43, "8px sans-serif", mainCbLoad > 99 ? "#dc2626" : "#64748b", "left");

            // Bus and Envelope definitions
            const invSpread = (actualInv - 1) * minInvWidth;
            let startInvX = actualInv === 1 ? px : px - (invSpread / 2);
            let endInvX = actualInv === 1 ? px : px + (invSpread / 2);

            // Bus bounds calculation
            let busLeft = actualInv === 1 ? px - 40 : startInvX - 10;
            let busRight = actualInv === 1 ? px + 40 : endInvX + 10;

            // Panel Title formatting over the precise structural bus-bar tip
            let totalBusAcMw = (actualInv * invPnom) / 1000;
            let lvBusAmpLoad = (totalBusAcMw * 1000) / ((lvV/1000) * 1.732 * pf);
            let busLoadPct = mainCbVal > 0 ? (lvBusAmpLoad / mainCbVal) * 100 : 0;

            this.drawText(`ACP ${p + 1}`, busLeft, panelY - 24, "bold 12px sans-serif", "#0284c7", "left");
            this.drawText(`${lvV}V | ${mainCbVal} A  (${busLoadPct.toFixed(1)}%)`, busLeft, panelY - 10, "10px sans-serif", busLoadPct > 99 ? "#dc2626" : "#64748b", "left");

            // Horizontal Bus draw
            this.drawLine(busLeft, panelY, busRight, panelY, '#2563eb', 4);

            for (let i = 0; i < actualInv; i++) {
                const globalInvIndex = (p * invPerPanel) + i; // which inverter is this in the block (0-indexed)
                let ix = actualInv === 1 ? px : startInvX + (i * minInvWidth);

                // Dynamically distribute strings asymmetrically across all inverters
                let localStrPerInv = globalStrPerInv;
                if (isSpecificBlock && actualBlockInvs > 0 && actualBlockStrings > 0) {
                    const totalRoundedStrings = Math.round(actualBlockStrings);
                    const baseStrings = Math.floor(totalRoundedStrings / actualBlockInvs);
                    const remainder = totalRoundedStrings % actualBlockInvs;
                    // Distribute leftover strings starting from the first inverter
                    localStrPerInv = baseStrings + (globalInvIndex < remainder ? 1 : 0);
                }

                // Route from Bus to Inverter
                const invCbVal = parseInt(document.getElementById('lv-inv-cb')?.value) || 250;
                const invCbLoad = invCbVal > 0 ? (iInvA / invCbVal) * 100 : 0;
                
                this.drawLine(ix, panelY, ix, invY - 20);
                this.drawCB(ix, panelY + 40, 3); // Outgoing CB
                
                this.drawText(`${invCbVal} A`, ix + 5, panelY + 36, "8px sans-serif", "#1e293b", "left");
                this.drawText(`${invCbLoad.toFixed(1)}%`, ix + 5, panelY + 46, "7px sans-serif", invCbLoad > 99 ? "#dc2626" : "#64748b", "left");

                if (window.CablesEngine) {
                    const invPnomCw = parseFloat(document.getElementById('inv-pnom')?.value) || 125;
                    const lvVolts = parseFloat(document.getElementById('lv-v')?.value) || 400;
                    const statsLv = CablesEngine.calculateCableStats('lvac', invPnomCw / 1000, lvVolts / 1000, 0.9, 15); // assumed 15m pad dist
                    if (statsLv) {
                        this.drawText(statsLv.summaryLine1, ix + 7, panelY + 60, "bold 7px sans-serif", "#2563eb", "left", -90);
                        this.drawText(statsLv.summaryLine2, ix + 14, panelY + 60, "6px sans-serif", "#64748b", "left", -90);
                    }
                }

                let invName = `INV-${globalInvIndex + 1}`;
                if (this.lvTargetBlock) {
                    invName = `INV-${String(this.lvTargetBlock).replace('Block ', '')}.${globalInvIndex + 1}`;
                }

                // Inverter Vector Symbol (Replaces basic box)
                this.drawInverterSymbol(ix, invY, 20);
                
                const invAcKw = invPnom;
                const invDcKw = (localStrPerInv * modPerStr * pvP) / 1000;
                const invRatio = (invDcKw / invAcKw).toFixed(2);

                this.drawText(invName, ix + 22, invY - 14, "bold 8px sans-serif", "#d97706", "left");
                this.drawText(`${invDcKw.toFixed(1)} kWdc\n${invAcKw.toFixed(1)} kWac\nDC/AC: ${invRatio}`, ix + 22, invY - 3, "6px sans-serif", "#64748b", "left");

                // Strings under inverter
                const strY = invY + 100;
                this.drawLine(ix, invY + 20, ix, strY - 25); // shortened to not pierce rect

                if (window.CablesEngine) {
                    const modPerStr = parseInt(document.getElementById('inv-mods-str')?.value) || 25;
                    const pvP = parseFloat(document.getElementById('pv-pnom')?.value) || 550;
                    const strAmps = 13; // Nominal string current
                    const statsDc = CablesEngine.calculateCableStats('dc', (pvP * modPerStr / 1000000) * 0.6, 0.6, 1.0, 50); // placeholder 600V, 50m
                    if (statsDc) {
                         this.drawText(`DC String: ${statsDc.cable.size}mm² Cu`, ix + 5, invY + 45, "bold 6px sans-serif", "#16a34a", "left", -90);
                    }
                }

                this.drawRect(ix - 30, strY - 25, 60, 50, '#dcfce7', '#16a34a', 2);
                this.drawText(`${localStrPerInv} Strings`, ix, strY - 10, "bold 10px sans-serif", "#16a34a");
                this.drawText(`${modPerStr} Mod/Str`, ix, strY + 4, "bold 9px sans-serif", "#16a34a");
                this.drawText("DC Array", ix, strY + 18, "8px sans-serif", "#16a34a");

                // User requested removal of DC Cable text here to prevent visual clutter
            }
        }
    },

    drawMvHvSld(isSecondPass = false) {
        const EE = window.ElectricalEngine;
        const LE = window.LayoutEngine;
        
        const pocLevel = document.getElementById('poc-v-level')?.value || 'HV';
        const mvV = parseFloat(document.getElementById('mv-v')?.value) || 33;
        let commonV = pocLevel === 'MV' ? mvV : (parseFloat(document.getElementById('hv-v')?.value) || 132);
        
        // Final fallback override
        if (commonV < mvV) commonV = mvV;
        
        if (pocLevel === 'LV') {
            this.drawText("Project connects at LV level.\nNo MV/HV infrastructure to display.", this.canvas.width / 2, this.canvas.height / 2, "bold 24px sans-serif", "#64748b");
            return;
        }

        // 1. Gather all blocks grouped by Area
        let areasMap = {};
        let totalStatsBlocks = 0;

        if (LE && LE.layoutStore && LE.layoutStore.size > 0) {
            let fallbackIdx = 1;
            const SE = window.SiteEngine;
            let allAreas = [];
            if (SE && SE.overlays) {
                allAreas = SE.overlays.filter(o => o.category === 'area' || (o.getPath && !o.subType && !o.category));
            }

            LE.layoutStore.forEach((data, areaObj) => {
                if (data.blocks && data.blocks.length > 0) {
                    const idx = allAreas.indexOf(areaObj);
                    const name = areaObj.areaName || `Area ${idx >= 0 ? idx + 1 : fallbackIdx++}`;
                    let resolvedDsId = areaObj.linkedDsId || 'default_ds';

                    if (window.ElectricalEngine && window.ElectricalEngine.areaMappings && window.ElectricalEngine.areaMappings[name]) {
                        const mapDS = window.ElectricalEngine.areaMappings[name].ds;
                        if (mapDS && mapDS !== 'auto') resolvedDsId = mapDS;
                    }

                    const uid = areaObj.__uid || (Math.random() + "");
                    const uniqueKey = `${name}_${uid}`; // Ensure unique string key
                    areasMap[uniqueKey] = {
                        name: name,
                        blocks: data.blocks,
                        linkedDsId: resolvedDsId,
                        linkedPocId: areaObj.linkedPocId || 'default_poc',
                        sldConfig: data.sldConfig || null
                    };
                    totalStatsBlocks += data.blocks.length;
                }
            });
        }
        
        // Fallback for debug/manual modes
        if (totalStatsBlocks === 0) {
            let fakeBlocks = [];
            if (LE && LE.blocks && LE.blocks.length > 0) {
                fakeBlocks = [...LE.blocks];
            } else if (parseInt(document.getElementById('mv-sum-total-stations')?.innerText) > 0) {
                const fakeQ = parseInt(document.getElementById('mv-sum-total-stations')?.innerText);
                for(let i=0; i<fakeQ; i++) fakeBlocks.push({ id: `Block ${i+1}` });
            }
            if (fakeBlocks.length > 0) {
                areasMap['Project Layout_1'] = { name: 'Project Layout', blocks: fakeBlocks };
            }
        }

        if (Object.keys(areasMap).length === 0) {
            this.drawText("No project areas with electrical blocks found.", this.canvas.width / 2, this.canvas.height / 2, "bold 20px sans-serif", "#64748b");
            this.drawText("Please draw an area and generate layout first.", this.canvas.width / 2, this.canvas.height / 2 + 40, "14px sans-serif", "#94a3b8");
            return;
        }

        // --- Fetch Engineering Configuration ---
        const blocksPerLoop = parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1;
        const is3WindingSystem = parseInt(document.getElementById('mvhv-winding-type')?.value) === 3;
        
        // Ensure total stations are available as a fallback
        let globalTotalStations = 0;
        Object.values(areasMap).forEach(ad => globalTotalStations += ad.blocks.length);
        const stationsPerBus = parseInt(document.getElementById('mv-sum-stations-bus')?.innerText) || globalTotalStations || 1;
        
        const standardLvmvMva = parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15;
        const pf = parseFloat(document.getElementById('sys-pf')?.value) || 0.9;
        
        // --- Calculate Topology (Mapped by DS -> Transformers -> Bus -> Blocks) ---
        let dsGroups = {};
        for (const [uniqueKey, areaData] of Object.entries(areasMap)) {
            const dsId = areaData.linkedDsId || 'default_ds';
            if (!dsGroups[dsId]) dsGroups[dsId] = { dsName: `DS/Common Bus`, areas: [] };
            dsGroups[dsId].areas.push(areaData);
        }

        if (window.SiteEngine) {
            Object.keys(dsGroups).forEach(dsId => {
                const ds = window.SiteEngine.overlays.find(o => (o.__uid === dsId || o.overlayId === dsId));
                if (ds && ds.areaName) dsGroups[dsId].dsName = ds.areaName;
            });
        }

        let dsLayouts = [];
        let totalW = 400; // Base lateral padding
        let busIterator = 1;
        let maxDeep = 1;

        for (const [dsId, group] of Object.entries(dsGroups)) {
            let dsTrData = [];
            let dsWidth = 0;
            let trIterator = 1;

            // 1. Preserve segregation: Pre-build independent physical field loops (feeders) per Area.
            // This prevents Area 1 and Area 2 RMU blocks from chaining onto the exact same grid feeder.
            let allLoopsToDistribute = [];
            let unifiedSldConfig = null;
            let combinedAreaNames = [];
            
            for (let a = 0; a < group.areas.length; a++) {
                const areaData = group.areas[a];
                combinedAreaNames.push(areaData.name);
                
                const areaSldConfig = areaData.sldConfig || {
                    blocksPerLoop: parseInt(document.getElementById('mv-sum-stations-bay')?.innerText) || 1,
                    is3WindingSystem: parseInt(document.getElementById('mvhv-winding-type')?.value) === 3,
                    stationsPerBus: parseInt(document.getElementById('mv-sum-stations-bus')?.innerText) || 1,
                    transformerRating: parseFloat((document.getElementById('substation-size')?.value || "3.15").split(';')[0]) || 3.15,
                    mvhvMva: 100
                };
                if (!unifiedSldConfig) unifiedSldConfig = areaSldConfig;

                let numLoopsForArea = Math.ceil(areaData.blocks.length / Math.max(1, areaSldConfig.blocksPerLoop));
                for(let l = 0; l < numLoopsForArea; l++) {
                    const lStart = l * areaSldConfig.blocksPerLoop;
                    const lEnd = Math.min(lStart + areaSldConfig.blocksPerLoop, areaData.blocks.length);
                    allLoopsToDistribute.push({
                         areaName: areaData.name,
                         blocks: areaData.blocks.slice(lStart, lEnd),
                         sldConfig: areaSldConfig
                    });
                }
            }

            if (!unifiedSldConfig) unifiedSldConfig = { blocksPerLoop: 1, stationsPerBus: 1, transformerRating: 3.15, mvhvMva: 100 };
            if (commonV <= mvV) unifiedSldConfig.is3WindingSystem = false;

            const dsLabelName = group.dsName && group.dsName !== "DS/Common Bus" ? group.dsName : combinedAreaNames.join(' + ');

            let numAreaBuses = 1;
            if (commonV > mvV) {
                let totalStations = allLoopsToDistribute.reduce((sum, loop) => sum + loop.blocks.length, 0);
                numAreaBuses = Math.ceil(totalStations / Math.max(1, unifiedSldConfig.stationsPerBus));
                if (numAreaBuses < 1) numAreaBuses = 1;
            }

            const areaBusesData = [];
            let loopsDistributed = 0;

            for (let i = 0; i < numAreaBuses; i++) {
                let busLoopsArray = [];
                let currentBusStationCount = 0;

                while (loopsDistributed < allLoopsToDistribute.length) {
                    const loopObj = allLoopsToDistribute[loopsDistributed];
                    
                    if (numAreaBuses === 1 || i === numAreaBuses - 1) {
                         busLoopsArray.push(loopObj);
                         currentBusStationCount += loopObj.blocks.length;
                         loopsDistributed++;
                    } else if (busLoopsArray.length === 0 || (currentBusStationCount + loopObj.blocks.length <= unifiedSldConfig.stationsPerBus)) {
                         busLoopsArray.push(loopObj);
                         currentBusStationCount += loopObj.blocks.length;
                         loopsDistributed++;
                    } else {
                         break;
                    }
                }

                let numLoops = busLoopsArray.length;
                let maxLoopDeep = Math.max(...busLoopsArray.map(l => l.blocks.length));
                maxDeep = Math.max(maxDeep, maxLoopDeep || 1);
                let busW = Math.max(400, numLoops * 120); 

                areaBusesData.push({
                   loopCount: numLoops,
                   width: busW,
                   busIdx: busIterator++,
                   areaTitle: numAreaBuses > 1 ? `${dsLabelName} MV Bus ${i+1}` : `${dsLabelName} MV Bus`,
                   loopsInfo: busLoopsArray,
                   sldConfig: unifiedSldConfig,
                   blocks: busLoopsArray.reduce((acc, l) => acc.concat(l.blocks), []) // Flat map for compatibility down-chain
                });
            }

            const busesPerTr = unifiedSldConfig.is3WindingSystem ? 2 : 1;
            const numAreaTrs = Math.ceil(areaBusesData.length / busesPerTr);

            for (let i = 0; i < numAreaTrs; i++) {
                const trBuses = areaBusesData.slice(i * busesPerTr, (i + 1) * busesPerTr);
                
                let trWidth = 0;
                trBuses.forEach(b => trWidth += b.width);
                if (trBuses.length === 2) trWidth += 200; // spacer gap
                trWidth = Math.max(trWidth, 500);

                dsTrData.push({
                    is3Winding: (unifiedSldConfig.is3WindingSystem && trBuses.length >= 2),
                    width: trWidth,
                    areaLabel: numAreaTrs > 1 ? `${dsLabelName} (TR ${trIterator})` : dsLabelName,
                    buses: trBuses,
                    sldConfig: unifiedSldConfig
                });
                dsWidth += trWidth + 200; 
                trIterator++;
            }

            dsWidth = Math.max(800, dsWidth - 200); 
            totalW += dsWidth + 300; 

            dsLayouts.push({
                dsId: dsId,
                dsName: group.dsName,
                width: dsWidth,
                trData: dsTrData
            });
        }

        const totalH = Math.max(1200, 1000 + (maxDeep * 140)); // Resize canvas safely vertically ensuring depth fits

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

        // Compute absolute project load to display HV Bus utilization
        let totalHVacMw = 0;
        dsLayouts.forEach(ds => {
            ds.trData.forEach(tr => {
                tr.buses.forEach(b => {
                    const rating = b.sldConfig.transformerRating || standardLvmvMva;
                    totalHVacMw += b.blocks.length * rating * pf;
                });
            });
        });

        const uiHvBusA = parseFloat(document.getElementById('hv-bus')?.value) || 2000;
        const actualHvAmpA = (totalHVacMw * 1000) / (commonV * 1.732 * pf);
        const pocLoadPct = uiHvBusA > 0 ? ((actualHvAmpA / uiHvBusA) * 100) : 0;
        
        // Calculate exact horizontal total spread of all deployed DS Nodes to group them dynamically
        const totalSysSpreadWidth = dsLayouts.reduce((sum, ds) => sum + ds.width, 0) + (Math.max(0, dsLayouts.length - 1) * 300);
        let currentDsX = pocX - (totalSysSpreadWidth / 2);

        // Find exact leftmost and rightmost transformer connection dropping points
        let leftmostTrX = pocX;
        let rightmostTrX = pocX;

        let tempDsX = currentDsX;
        for (let d = 0; d < dsLayouts.length; d++) {
            let dsInfo = dsLayouts[d];
            let dsCenterX = tempDsX + (dsInfo.width / 2);
            let trCurrentX = dsCenterX - (dsInfo.width / 2);
            for (let t = 0; t < dsInfo.trData.length; t++) {
                let trInfo = dsInfo.trData[t];
                let trCenterX = trCurrentX + trInfo.width / 2;
                if (d === 0 && t === 0) leftmostTrX = trCenterX;
                if (d === dsLayouts.length - 1 && t === dsInfo.trData.length - 1) rightmostTrX = trCenterX;
                trCurrentX += trInfo.width;
            }
            tempDsX += dsInfo.width + 300;
        }

        // 3. Draw The SINGLE Common Master Bus (or POC Cable Routing if MV-only)
        const commonBusY = pocY + 180;
        const commonBusW = totalW - 400; // Legacy width for HV

        if (commonV > mvV) {
            // Standard HV Collection Bus
            const masterBusLabel = `HV BB1 (Master Bus)`;
            this.drawLine(pocX - commonBusW/2, commonBusY, pocX + commonBusW/2, commonBusY, '#8b5cf6', 8);
            this.drawText(masterBusLabel, pocX - commonBusW/2, commonBusY - 24, "bold 16px sans-serif", "#8b5cf6", "left");
            this.drawText(`${commonV} kV | ${uiHvBusA} A   (${pocLoadPct.toFixed(1)}%)`, pocX - commonBusW/2, commonBusY - 8, "12px sans-serif", pocLoadPct > 99 ? "#dc2626" : "#64748b", "left");
            
            // Connect POC to Master Bus
            this.drawLine(pocX, pocY + 30, pocX, commonBusY);
            this.drawCB(pocX, commonBusY - 80);
            this.drawText(`${uiHvBusA} A`, pocX + 15, commonBusY - 80, "10px sans-serif", "#1e293b", "left");
            
            const outGoingNum = parseInt(document.getElementById('hv-poc-outgoing')?.value) || 1;
            const finalLoadPct = pocLoadPct / outGoingNum;
            this.drawText(`${finalLoadPct.toFixed(1)}%`, pocX + 15, commonBusY - 68, "8px sans-serif", finalLoadPct > 99 ? "#dc2626" : "#64748b", "left");
        } else {
            // MV-Only Topology: Delivery Stations cable directly to POC. No intermediary "Collector Switchgear" required.
            this.drawLine(pocX, pocY + 30, pocX, commonBusY, '#10b981', 3); // Drop cable from POC
            
            // Mathematically secure the geometric gap if the underlying array padding shifts trCenterX off of pocX
            const drawMinX = Math.min(pocX, leftmostTrX);
            const drawMaxX = Math.max(pocX, rightmostTrX);
            if (drawMinX !== drawMaxX) {
                this.drawLine(drawMinX, commonBusY, drawMaxX, commonBusY, '#10b981', 3); 
            }
        }

        // 4. Draw Delivery Stations iteratively (directly beneath the Master Bus)
        for (let d = 0; d < dsLayouts.length; d++) {
            let dsInfo = dsLayouts[d];
            let dsCenterX = currentDsX + (dsInfo.width / 2);
            
            // Draw a subtle "DS grouping" text above the transformers
            if (dsInfo.dsName && dsInfo.dsName !== "DS/Common Bus") {
                this.drawText(`[ ${dsInfo.dsName} Delivery Substation ]`, dsCenterX, commonBusY + 25, "bold 14px sans-serif", "#8b5cf6");
            }
            
            // Draw Split System Layout iterating across designated Transformer -> Bus geometry for this DS
            let trCurrentX = dsCenterX - (dsInfo.width / 2);
            for (let t = 0; t < dsInfo.trData.length; t++) {
                let trInfo = dsInfo.trData[t];
                let trCenterX = trCurrentX + trInfo.width / 2;
                
                // Extracted local variable override handling for transformer calculations
                const localTrMva = trInfo.sldConfig.transformerRating || standardLvmvMva;

                if (commonV > mvV) {
                   const dropToY = trInfo.is3Winding ? commonBusY + 169 : commonBusY + 170;
                   this.drawLine(trCenterX, commonBusY, trCenterX, dropToY);
                   this.drawCB(trCenterX, commonBusY + 80);
                   
                   const uiMvhvMva = trInfo.sldConfig.mvhvMva || 100; // Use captured snapshot value!
                   const totalTrMw = trInfo.buses.reduce((sum, b) => sum + (b.blocks.length * localTrMva * pf), 0);
                   const trLoadPct = uiMvhvMva > 0 ? ((totalTrMw / uiMvhvMva) * 100) : 0;
                   
                   let trHvCb = Math.max(630, Math.ceil((uiMvhvMva * 1000) / (commonV * 1.732 * pf) / 100) * 100);
                   this.drawText(`${trHvCb} A`, trCenterX + 15, commonBusY + 80, "10px sans-serif", "#1e293b", "left");
                   this.drawText(`${trLoadPct.toFixed(1)}%`, trCenterX + 15, commonBusY + 92, "8px sans-serif", trLoadPct > 99 ? "#dc2626" : "#64748b", "left");
                   
                   this.drawTransformer(trCenterX, commonBusY + 220, 30, trInfo.is3Winding);
                   
                   let trTextX = trCenterX + 45;
                   if (trInfo.is3Winding) trTextX += 30;
                   
                   this.drawText(`${uiMvhvMva} MVA`, trTextX, commonBusY + 205, "bold 12px sans-serif", "#f59e0b", "left");
                   this.drawText(`${commonV}/${mvV}${trInfo.is3Winding ? '/' + mvV : ''} kV`, trTextX, commonBusY + 225, "12px sans-serif", "#64748b", "left");
                   this.drawText(`${trLoadPct.toFixed(1)}%`, trTextX, commonBusY + 240, "bold 10px sans-serif", trLoadPct > 99 ? "#dc2626" : "#16a34a", "left");
                }

                let totalBusWidthInsideTr = trInfo.buses.reduce((sum, b) => sum + b.width, 0) + (Math.max(0, trInfo.buses.length - 1) * 200);
                let busOffsetX = trCenterX - (totalBusWidthInsideTr / 2);
                
                for (let i = 0; i < trInfo.buses.length; i++) {
                    let busInfo = trInfo.buses[i];
                    let busXStart = busOffsetX;
                    let busXEnd = busXStart + busInfo.width;
                    let connectX = busXStart + busInfo.width / 2;
                    let localBusTrMva = busInfo.sldConfig.transformerRating || standardLvmvMva;

                    const areaBusY = (commonV > mvV) ? commonBusY + 400 : commonBusY + 180;
                    const uiMvBusA = parseFloat(document.getElementById('mv-bus')?.value) || 3200;
                    const actualMvAmpA = (busInfo.blocks.length * localBusTrMva * 1000) / (mvV * 1.732 * pf);
                    const mvBusLoadPct = uiMvBusA > 0 ? ((actualMvAmpA / uiMvBusA) * 100) : 0;

                    this.drawLine(busXStart, areaBusY, busXEnd, areaBusY, '#2563eb', 5);
                    this.drawText(`${busInfo.areaTitle}`, busXStart, areaBusY - 24, "bold 13px sans-serif", "#2563eb", "left");
                    this.drawText(`${mvV} kV | ${uiMvBusA} A   (${mvBusLoadPct.toFixed(1)}%)`, busXStart, areaBusY - 10, "11px sans-serif", mvBusLoadPct > 99 ? "#dc2626" : "#64748b", "left");

                    // Connect Bus visually back UP to Transformer node output
                    if (commonV > mvV) {
                        if (trInfo.is3Winding && trInfo.buses.length === 2) { 
                            let outX = i === 0 ? trCenterX - 18 : trCenterX + 18;
                            this.drawLine(outX, commonBusY + 265, outX, commonBusY + 300); 
                            this.drawLine(outX, commonBusY + 300, connectX, commonBusY + 300); 
                            this.drawLine(connectX, commonBusY + 300, connectX, areaBusY);
                            this.drawCB(connectX, areaBusY - 50);
                        } else { 
                            let startY = trInfo.is3Winding ? commonBusY + 265 : commonBusY + 270;
                            connectX = trCenterX; 
                            this.drawLine(trCenterX, startY, connectX, areaBusY);
                            this.drawCB(connectX, areaBusY - 50);
                        }
                    } else { 
                        // MV-ONLY (No step-up Tr). Just connect straight up to POC horizontal bounding loop.
                        this.drawLine(connectX, commonBusY, connectX, areaBusY, '#10b981', 3);
                        
                        // Overwrite the bottom stub to draw the Breaker in standardized MV format.
                        this.drawLine(connectX, areaBusY - 70, connectX, areaBusY, '#1e293b', 2);
                        this.drawCB(connectX, areaBusY - 60);
                    }
                    this.drawText(`${uiMvBusA} A`, connectX + 15, areaBusY - 50, "10px sans-serif", "#1e293b", "left");
                    this.drawText(`${mvBusLoadPct.toFixed(1)}%`, connectX + 15, areaBusY - 38, "8px sans-serif", mvBusLoadPct > 99 ? "#dc2626" : "#64748b", "left");

                    // Draw Downstream Blocks specifically designated for this Bus
                    let spacing = busInfo.width / (busInfo.loopCount + 1);

                    for (let l = 0; l < busInfo.loopCount; l++) {
                        let bx = busXStart + (spacing * (l + 1));
                        let by = areaBusY + 150;
                        
                        const loopObj = busInfo.loopsInfo[l];
                        const loopStations = loopObj.blocks.length;
                        const localBusTrMva = loopObj.sldConfig.transformerRating || standardLvmvMva;

                        this.drawLine(bx, areaBusY, bx, by - 25);
                        this.drawCB(bx, areaBusY + 50, 4);

                        const uiMvBayCb = parseFloat(document.getElementById('mv-bay-cb')?.value) || 630;
                        const loopAmps = (loopStations * localBusTrMva * 1000) / (mvV * 1.732 * pf);
                        const feederLoad = uiMvBayCb > 0 ? ((loopAmps / uiMvBayCb) * 100) : 0;
                        
                        this.drawText(`${uiMvBayCb} A`, bx + 12, areaBusY + 50, "10px sans-serif", "#1e293b", "left");
                        this.drawText(`${feederLoad.toFixed(1)}%`, bx + 12, areaBusY + 62, "8px sans-serif", feederLoad > 99 ? "#dc2626" : "#64748b", "left");

                        // Add MV Cable sizing text
                        if (window.CablesEngine) {
                            const statsMv = CablesEngine.calculateCableStats('mv', (loopStations * localBusTrMva * pf), mvV, pf, 200); // 200m avg feeder
                            if (statsMv) {
                                this.drawText(statsMv.summaryLine1, bx - 8, areaBusY + 100, "bold 7px sans-serif", "#1e293b", "right", -90);
                            }
                        }

                        for (let s = 0; s < loopStations; s++) {
                            const block = loopObj.blocks[s];

                            this.drawRect(bx - 25, by - 25, 50, 50, '#f8fafc', '#3b82f6', 2.5);
                            this.drawText(`MV${s + 1}`, bx, by - 8, "bold 10px sans-serif", "#1e293b");
                            this.drawText("VIEW LV", bx, by + 12, "bold 8px sans-serif", "#3b82f6");
                            
                            const blockMw = block.item?.dcMw || block.dcMw;
                            if (blockMw) {
                                this.drawText(`${localBusTrMva} MVA\n${blockMw}`, bx + 27, by - 2, "8px sans-serif", "#64748b", "left");
                            } else {
                                this.drawText(`${localBusTrMva} MVA`, bx + 27, by + 2, "8px sans-serif", "#64748b", "left");
                            }
                            
                            const blockId = block.item?.blockId || block.blockId || block.item?.id || block.id || `Block ${s+1}`;
                            this.hitZones.push({ 
                                type: 'mv-block', 
                                id: blockId, 
                                x: bx - 30, y: by - 30, w: 60, h: 60 
                            });

                            if (s < loopStations - 1) {
                                this.drawLine(bx, by + 25, bx, by + 75);
                                
                                if (window.CablesEngine) {
                                    const stationsDownstream = loopStations - (s + 1);
                                    const statsMvFeeder = CablesEngine.calculateCableStats('mv', (stationsDownstream * localBusTrMva * pf), mvV, pf, 100);
                                    if (statsMvFeeder) {
                                        this.drawText(statsMvFeeder.summaryLine1, bx + 5, by + 50, "6px sans-serif", "#64748b", "left", -90);
                                    }
                                }

                                by += 100; 
                            }
                        }
                    }
                    busOffsetX += busInfo.width + 200; 
                }
                trCurrentX += trInfo.width + 200;
            }
            currentDsX += dsInfo.width + 300;
        }
    }
};

window.addEventListener('DOMContentLoaded', () => {
    // Wait for elements
    setTimeout(() => {
        if (window.SldEngine) window.SldEngine.init();
    }, 100);
});
