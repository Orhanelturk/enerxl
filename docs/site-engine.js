
/**
 * Site Engine - Handles all Site Configuration, Drawing, and Geography Logic
 */

const SiteEngine = {
    map: null,
    drawingManager: null,
    overlays: [],
    selectedOverlays: [],
    labels: [],
    pendingOverlay: null,
    clipboard: null,
    pasteMode: false,
    isShiftDown: false,
    isMarqueeActive: false,
    dragStarted: false,
    dragStartPositions: new Map(),
    nextOverlayId: 1, // Unique ID counter for stable references

    generateUid() {
        return 'ov-' + (this.nextOverlayId++) + '-' + Math.random().toString(36).substr(2, 5);
    },

    // Configuration Styles
    styles: {
        area: { fillColor: '#FFD700', fillOpacity: 0.25, strokeColor: '#FFD700', strokeWeight: 2, editable: false, draggable: false, zIndex: 1, selectedColor: '#FFFF00' },
        constraint: { fillColor: '#ef4444', fillOpacity: 0.15, strokeColor: '#ef4444', strokeWeight: 1.5, editable: false, draggable: false, zIndex: 2, selectedColor: '#800000' },
        setback: { fillColor: '#64748b', fillOpacity: 0.45, strokeColor: '#475569', strokeWeight: 1, editable: false, draggable: false, zIndex: 3, selectedColor: '#334155' },
        selected: { strokeWeight: 4 },
        marquee: { fillColor: '#3b82f6', fillOpacity: 0.1, strokeColor: '#3b82f6', strokeWeight: 1, clickable: false, zIndex: 1000 },
        measure: { strokeColor: '#10b981', strokeOpacity: 0.8, strokeWeight: 4, editable: true, draggable: true, zIndex: 999 },
        topo_excluded: { fillColor: '#78350f', fillOpacity: 0.6, strokeColor: '#451a03', strokeWeight: 1, zIndex: 5 }
    },

    topoPoints: [],
    topoMesh: [],

    init(map) {
        this.map = map;
        this.initDrawingManager();
        this.setupMapListeners();
        this.setupModalListeners();
        this.setupActionListeners();
    },

    initDrawingManager() {
        this.drawingManager = new google.maps.drawing.DrawingManager({
            drawingControl: false,
            polygonOptions: this.styles.area,
            rectangleOptions: this.styles.area,
            circleOptions: this.styles.area,
            polylineOptions: this.styles.measure,
            markerOptions: { draggable: true }
        });
        this.drawingManager.setMap(this.map);

        google.maps.event.addListener(this.drawingManager, 'overlaycomplete', (event) => {
            const overlay = event.overlay;
            if (this.isMarqueeActive) {
                this.handleMarqueeComplete(overlay);
                overlay.setMap(null);
                this.setDrawingMode(null);
                document.getElementById('tool-marquee').classList.remove('active');
                return;
            }
            if (event.type === google.maps.drawing.OverlayType.RECTANGLE && !this.isMarqueeActive) {
                const b = overlay.getBounds(); const ne = b.getNorthEast(); const sw = b.getSouthWest();
                const paths = [[{ lat: ne.lat(), lng: sw.lng() }, { lat: ne.lat(), lng: ne.lng() }, { lat: sw.lat(), lng: ne.lng() }, { lat: sw.lat(), lng: sw.lng() }]];
                const poly = new google.maps.Polygon({ paths: paths, ...(window.currentCategory === 'area' ? this.styles.area : this.styles.constraint), map: this.map });
                overlay.setMap(null);
                poly.type = google.maps.drawing.OverlayType.POLYGON;
                poly.category = window.currentCategory || 'area';
                this.processNewOverlay(poly);
            } else if (event.type === google.maps.drawing.OverlayType.MARKER && (window.currentShape === 'poc' || window.currentShape === 'station')) {
                const latLng = overlay.getPosition();
                overlay.setMap(null); // Remove marker
                const shape = window.currentShape;
                let w, h, rot;
                if (shape === 'poc') {
                    w = parseFloat(document.getElementById('poc-width')?.value) || 30;
                    h = parseFloat(document.getElementById('poc-height')?.value) || 15;
                    rot = parseFloat(document.getElementById('poc-rotation')?.value) || 0;
                } else {
                    w = parseFloat(document.getElementById('ds-width')?.value) || 30;
                    h = parseFloat(document.getElementById('ds-height')?.value) || 15;
                    rot = parseFloat(document.getElementById('ds-rotation')?.value) || 0;
                }
                const mToDeg = 1 / 111320;
                const cosLat = Math.cos(latLng.lat() * Math.PI / 180);
                const dLat = (h / 2) * mToDeg;
                const dLng = (w / 2) * mToDeg / cosLat;

                // Construct basic unrotated polygon ring
                const ring = [
                    [latLng.lng() - dLng, latLng.lat() + dLat],
                    [latLng.lng() + dLng, latLng.lat() + dLat],
                    [latLng.lng() + dLng, latLng.lat() - dLat],
                    [latLng.lng() - dLng, latLng.lat() - dLat],
                    [latLng.lng() - dLng, latLng.lat() + dLat]
                ];
                let geoPoly = turf.polygon([ring]);

                // Apply rotation if any
                if (rot !== 0) {
                    geoPoly = turf.transformRotate(geoPoly, rot, { pivot: [latLng.lng(), latLng.lat()] });
                }

                const coords = geoPoly.geometry.coordinates[0].map(c => ({ lat: c[1], lng: c[0] }));
                const poly = new google.maps.Polygon({
                    paths: coords,
                    ...this.styles.constraint,
                    map: this.map,
                    editable: true, draggable: true,
                    fillColor: shape === 'poc' ? '#286944' : '#ef4444'
                });
                poly.type = google.maps.drawing.OverlayType.POLYGON;
                poly.category = 'poc';
                poly.subType = shape;
                this.processNewOverlay(poly);
            } else if (event.type === google.maps.drawing.OverlayType.POLYLINE && window.currentShape === 'measure') {
                overlay.type = 'measure';
                overlay.category = 'measure';
                this.processNewOverlay(overlay);
                this.updateMeasureLabel(overlay);

                // Add listeners to update label when edited
                const path = overlay.getPath();
                google.maps.event.addListener(path, 'set_at', () => this.updateMeasureLabel(overlay));
                google.maps.event.addListener(path, 'insert_at', () => this.updateMeasureLabel(overlay));
                google.maps.event.addListener(path, 'remove_at', () => this.updateMeasureLabel(overlay));

                this.setDrawingMode(null);
                const btnMeasure = document.getElementById('tool-measure');
                if (btnMeasure) btnMeasure.classList.remove('active');
                return;
            } else {
                overlay.type = event.type;
                overlay.category = window.currentCategory || 'area';
                if (window.currentShape === 'poc') overlay.subType = 'poc';
                if (window.currentShape === 'station') overlay.subType = 'station';
                this.processNewOverlay(overlay);
            }
            this.setDrawingMode(null);
            document.querySelectorAll('.draw-tool-btn').forEach(b => b.classList.remove('active'));
        });
    },

    setupMapListeners() {
        this.map.addListener('click', (e) => {
            if (this.pasteMode && this.clipboard) {
                this.paste(e.latLng);
            } else {
                this.clearSelection();
                this.closeAllMenus();
            }
        });
        window.addEventListener('keydown', (e) => { if (e.key === 'Shift') this.isShiftDown = true; });
        window.addEventListener('keyup', (e) => { if (e.key === 'Shift') this.isShiftDown = false; });
    },

    setupModalListeners() {
        const saveBtn = document.getElementById('btn-save-name');
        saveBtn.addEventListener('click', () => {
            const input = document.getElementById('area-name-input');
            if (this.pendingOverlay && input.value) {
                this.pendingOverlay.areaName = input.value;
                this.updateAreaLabel(this.pendingOverlay);
                document.getElementById('naming-modal').classList.add('hidden');
                input.value = '';
                this.pendingOverlay = null;
            }
        });
        document.getElementById('btn-generate-grid').addEventListener('click', () => {
            this.generateGrid();
            document.getElementById('grid-modal').classList.add('hidden');
        });
    },

    setupActionListeners() {
        document.getElementById('btn-delete').addEventListener('click', () => this.deleteSelected());
        document.getElementById('btn-move-toggle').addEventListener('click', () => this.toggleMoveMode());
        document.getElementById('btn-rotate-toggle').addEventListener('click', () => {
            const popup = document.getElementById('rotation-popup');
            const isHidden = popup.classList.contains('hidden');
            popup.classList.toggle('hidden', !isHidden);
            document.getElementById('btn-rotate-toggle').classList.toggle('active', isHidden);

            // Reset slider on open
            if (isHidden) {
                const slider = document.getElementById('rotation-slider');
                if (slider) {
                    slider.value = 0;
                    this.totalRotation = 0;
                    const valDisplay = document.getElementById('rotation-value');
                    if (valDisplay) valDisplay.textContent = '0°';
                }
            }
        });
        document.getElementById('btn-close-edit').addEventListener('click', () => this.clearSelection());
        document.getElementById('btn-copy').addEventListener('click', () => this.copy());
        document.getElementById('btn-paste-toggle').addEventListener('click', () => this.togglePasteMode());
        document.getElementById('btn-grid').addEventListener('click', () => {
            const canGrid = this.selectedOverlays.length > 0 && this.selectedOverlays.every(o => o.category === 'constraint');
            if (canGrid) document.getElementById('grid-modal').classList.remove('hidden');
            else alert("Grid is only available for obstacles/constraints.");
        });
        document.getElementById('btn-align-menu').addEventListener('click', (e) => this.toggleMenu('align-menu', e));
        document.getElementById('btn-distribute-menu').addEventListener('click', (e) => this.toggleMenu('distribute-menu', e));
        document.querySelectorAll('.sub-btn').forEach(btn => btn.addEventListener('click', () => this.handleLayoutAction(btn.dataset.action)));

        // Rotation Slider Listener
        const slider = document.getElementById('rotation-slider');
        const valDisplay = document.getElementById('rotation-value');
        this.totalRotation = 0;

        slider?.addEventListener('input', (e) => {
            const currentVal = parseInt(e.target.value);
            const delta = currentVal - this.totalRotation;
            this.rotateSelected(delta);
            this.totalRotation = currentVal;
            valDisplay.textContent = `${this.totalRotation}°`;
        });

        // Topography Listeners
        document.getElementById('btn-create-mesh')?.addEventListener('click', () => this.createTopographyMesh());
        document.getElementById('btn-apply-topo')?.addEventListener('click', () => this.runTopographyAnalysis());
        document.getElementById('btn-clear-topo')?.addEventListener('click', () => this.clearTopoConstraints());
    },

    toggleMenu(id, e) {
        if (e) e.stopPropagation();
        const menu = document.getElementById(id);
        const isHidden = menu.classList.contains('hidden');
        this.closeAllMenus();
        if (isHidden) menu.classList.remove('hidden');
    },

    closeAllMenus() { document.querySelectorAll('.edit-sub-menu').forEach(m => m.classList.add('hidden')); },

    setDrawingMode(shape) {
        window.currentShape = shape;
        this.isMarqueeActive = (shape === 'marquee');
        let mode = null;
        if (shape === 'polygon') mode = google.maps.drawing.OverlayType.POLYGON;
        if (shape === 'rectangle' || shape === 'marquee') mode = google.maps.drawing.OverlayType.RECTANGLE;
        if (shape === 'circle') mode = google.maps.drawing.OverlayType.CIRCLE;
        if (shape === 'poc' || shape === 'station') mode = google.maps.drawing.OverlayType.MARKER;
        if (shape === 'measure') mode = google.maps.drawing.OverlayType.POLYLINE;
        this.drawingManager.setDrawingMode(mode);
        if (this.isMarqueeActive) this.drawingManager.setOptions({ rectangleOptions: this.styles.marquee });
        else if (shape === 'measure') this.drawingManager.setOptions({ polylineOptions: this.styles.measure });
        else this.setCategory(window.currentCategory || 'area');
        if (mode === google.maps.drawing.OverlayType.MARKER) {
            const isPOC = shape === 'poc';
            this.drawingManager.setOptions({
                markerOptions: {
                    icon: {
                        path: google.maps.SymbolPath.CIRCLE, scale: 8,
                        fillColor: isPOC ? '#286944' : '#ef4444',
                        fillOpacity: 1, strokeColor: '#FFFFFF', strokeWeight: 2.5
                    }
                }
            });
        }
    },

    setCategory(cat) {
        window.currentCategory = cat;
        const style = cat === 'area' ? this.styles.area : this.styles.constraint;
        this.drawingManager.setOptions({ polygonOptions: style, rectangleOptions: style, circleOptions: style });
    },

    processNewOverlay(overlay) {
        if (!overlay.__uid) overlay.__uid = this.generateUid();
        this.overlays.push(overlay);
        this.attachOverlayListeners(overlay);

        if (!overlay.areaName && (overlay.category === 'area' || overlay.subType) && overlay.subType !== 'topo_exclusion' && overlay.subType !== 'obstacle_point') {
            this.showNamingModal(overlay);
        }
        if (!this.isMarqueeActive) this.selectOverlay(overlay, this.isShiftDown);
    },

    attachOverlayListeners(overlay) {
        overlay.addListener('click', (e) => {
            if (e) e.stop();
            if (this.pasteMode && this.clipboard) this.paste(e.latLng);
            else this.selectOverlay(overlay, this.isShiftDown);
        });

        // Group Drag Logic
        overlay.addListener('dragstart', () => {
            if (!this.selectedOverlays.includes(overlay)) return;
            this.dragStarted = true;
            this.dragStartPositions.clear();
            this.selectedOverlays.forEach(o => {
                let p;
                if (o.getPosition) p = o.getPosition();
                else if (o.getCenter) p = o.getCenter();
                else if (o.getBounds) p = o.getBounds().getCenter();
                else if (o.getPaths) p = o.getPaths().getAt(0).getAt(0);
                else if (o.getPath) p = o.getPath().getAt(0);

                if (o.getPaths) {
                    const paths = o.getPaths().getArray().map(path => path.getArray().map(pt => ({ lat: pt.lat(), lng: pt.lng() })));
                    this.dragStartPositions.set(o, paths);
                } else if (o.getPath) {
                    this.dragStartPositions.set(o, o.getPath().getArray().map(pt => ({ lat: pt.lat(), lng: pt.lng() })));
                } else {
                    this.dragStartPositions.set(o, { lat: p.lat(), lng: p.lng() });
                }
            });
            const startPos = overlay.getPosition ? overlay.getPosition() : (overlay.getCenter ? overlay.getCenter() : (overlay.getBounds ? overlay.getBounds().getCenter() : (overlay.getPaths ? overlay.getPaths().getAt(0).getAt(0) : overlay.getPath().getAt(0))));
            this.anchorPoint = { lat: startPos.lat(), lng: startPos.lng() };
        });

        overlay.addListener('drag', () => {
            if (overlay.subType === 'station' || overlay.subType === 'poc') this.updateConnectedHvCables(overlay);
            
            if (!this.dragStarted) return;
            const currentPos = overlay.getPosition ? overlay.getPosition() : (overlay.getCenter ? overlay.getCenter() : (overlay.getBounds ? overlay.getBounds().getCenter() : (overlay.getPaths ? overlay.getPaths().getAt(0).getAt(0) : overlay.getPath().getAt(0))));
            const dl = currentPos.lat() - this.anchorPoint.lat;
            const dg = currentPos.lng() - this.anchorPoint.lng;

            this.selectedOverlays.forEach(o => {
                const start = this.dragStartPositions.get(o);
                if (!start) return;

                if (o.getPaths) {
                    const newPaths = start.map(path => path.map(p => ({ lat: p.lat + dl, lng: p.lng + dg })));
                    o.setPaths(newPaths);
                } else if (o.getPath) {
                    const newPath = start.map(p => ({ lat: p.lat + dl, lng: p.lng + dg }));
                    o.setPath(newPath);
                } else if (o.setPosition) {
                    o.setPosition({ lat: start.lat + dl, lng: start.lng + dg });
                } else if (o.setCenter) {
                    o.setCenter({ lat: start.lat + dl, lng: start.lng + dg });
                }

                if (o.category === 'area' || o.subType) this.updateAreaLabel(o);
            });
        });

        overlay.addListener('dragend', () => {
            this.dragStarted = false;
            if (overlay.category === 'area' || overlay.subType) this.updateAreaLabel(overlay);
            if (overlay.subType === 'station' || overlay.subType === 'poc') this.updateConnectedHvCables(overlay);
        });

        if (overlay.getPath) {
            const onChange = () => {
                if (overlay.category === 'area' || overlay.subType) this.updateAreaLabel(overlay);
                if (overlay.subType === 'station' || overlay.subType === 'poc') this.updateConnectedHvCables(overlay);
            };
            const path = overlay.getPath();
            google.maps.event.addListener(path, 'set_at', onChange);
            google.maps.event.addListener(path, 'insert_at', onChange);
            google.maps.event.addListener(path, 'remove_at', onChange);
        }
    },

    updateConnectedHvCables(overlay) {
        if (!overlay.__uid) return;
        const hvCables = this.overlays.filter(o => o.category === 'hv-cable' && (o.startNode === overlay.__uid || o.endNode === overlay.__uid));
        if (!hvCables.length) return;

        const getCenter = (eq) => {
            const bounds = new google.maps.LatLngBounds();
            if (eq.getPath) eq.getPath().forEach(p => bounds.extend(p));
            else if (eq.getPosition) bounds.extend(eq.getPosition());
            return bounds.getCenter();
        };
        const newCenter = getCenter(overlay);

        hvCables.forEach(cable => {
            if (!cable.getPath) return;
            const path = cable.getPath();
            if (cable.startNode === overlay.__uid) {
                path.setAt(0, newCenter);
            }
            if (cable.endNode === overlay.__uid) {
                path.setAt(path.getLength() - 1, newCenter);
            }
        });
    },

    handleMarqueeComplete(marqueeRect) {
        const bounds = marqueeRect.getBounds();
        if (!this.isShiftDown) this.clearSelection();
        this.overlays.forEach(overlay => {
            let pos;
            if (overlay.getPosition) pos = overlay.getPosition();
            else if (overlay.getCenter) pos = overlay.getCenter();
            else if (overlay.getBounds) pos = overlay.getBounds().getCenter();
            else if (overlay.getPaths) pos = overlay.getPaths().getAt(0).getAt(0);
            else if (overlay.getPath) {
                const b = new google.maps.LatLngBounds();
                overlay.getPath().forEach(p => b.extend(p));
                pos = b.getCenter();
            }
            if (pos && bounds.contains(pos)) this.selectOverlay(overlay, true);
        });
    },

    applySetback() {
        if (!this.selectedOverlays.length) { alert("Select a project area or constraint first."); return; }

        const distanceVal = parseFloat(document.querySelector('#tab-constraints .modern-input').value);
        const type = document.querySelector('#tab-constraints .seg-btn.active').dataset.type;
        const distInKm = distanceVal / 1000;

        this.selectedOverlays.forEach(overlay => {
            const geojson = this.getGeoJSON(overlay);
            if (!geojson) return;

            try {
                let frame;
                if (type === 'inner') {
                    const innerHole = turf.buffer(geojson, -distInKm, { units: 'kilometers' });
                    frame = turf.difference(geojson, innerHole);
                } else {
                    const outerRing = turf.buffer(geojson, distInKm, { units: 'kilometers' });
                    frame = turf.difference(outerRing, geojson);
                }

                if (frame) this.spawnFromGeoJSON(frame, 'setback', null, this.styles.setback);
            } catch (e) {
                console.error("Setback calculation failed", e);
                alert("Could not apply setback to this shape. Try a smaller distance.");
            }
        });
    },

    showNamingModal(overlay) {
        this.pendingOverlay = overlay;
        const modal = document.getElementById('naming-modal');
        const title = document.getElementById('modal-title');
        const input = document.getElementById('area-name-input');
        modal.classList.remove('hidden');
        input.focus();
        let defaultName = "";
        if (overlay.category === 'area') { title.innerText = "Name Project Area"; defaultName = `Area ${this.overlays.filter(o => o.category === 'area').length}`; }
        else if (overlay.subType === 'poc') { title.innerText = "Name POC"; defaultName = `POC ${this.overlays.filter(o => o.subType === 'poc').length}`; }
        else if (overlay.subType === 'station') { title.innerText = "Name DS"; defaultName = `DS ${this.overlays.filter(o => o.subType === 'station').length}`; }
        input.value = defaultName;
        input.select();
    },

    updateMeasureLabel(overlay) {
        if (!overlay.getPath) return;
        const path = overlay.getPath().getArray();
        let distance = 0;
        for (let i = 0; i < path.length - 1; i++) {
            distance += google.maps.geometry.spherical.computeDistanceBetween(path[i], path[i + 1]);
        }

        let displayDist = distance > 1000 ? (distance / 1000).toFixed(2) + ' km' : distance.toFixed(1) + ' m';
        overlay.areaName = displayDist; // Re-use area name property for simple label rendering compatibility

        if (!overlay.labelMarker) {
            overlay.labelMarker = new google.maps.Marker({
                map: this.map,
                label: { text: displayDist, color: '#10b981', fontWeight: '700', fontSize: '11px', className: 'map-label-bg' },
                icon: { path: google.maps.SymbolPath.CIRCLE, scale: 0 },
                zIndex: 1000
            });
        }

        // Put label on the last point of the path
        if (path.length > 0) {
            overlay.labelMarker.setPosition(path[path.length - 1]);
        }

        overlay.labelMarker.setLabel({
            ...overlay.labelMarker.getLabel(),
            text: displayDist
        });
    },

    updateAreaLabel(overlay) {
        if (!overlay.labelMarker) {
            let color = '#856404'; if (overlay.subType === 'poc') color = '#286944'; if (overlay.subType === 'station') color = '#ef4444'; if (overlay.subType === 'obstacle_point') color = '#ef4444';
            overlay.labelMarker = new google.maps.Marker({
                map: this.map,
                label: { text: overlay.areaName || "", color: color, fontWeight: '700', fontSize: '11px', className: 'map-label-bg' },
                icon: { path: google.maps.SymbolPath.CIRCLE, scale: 0 },
                zIndex: 1000
            });

        }
        const pos = this.getLabelPosition(overlay);
        if (pos) overlay.labelMarker.setPosition(pos);
        if (overlay.labelMarker.getLabel()) {
            overlay.labelMarker.setLabel({
                ...overlay.labelMarker.getLabel(),
                text: String(overlay.areaName || "")
            });
        }
    },

    getLabelPosition(overlay) {
        if (overlay.getPosition) return overlay.getPosition();
        if (overlay.getCenter) return overlay.getCenter();
        if (overlay.getBounds) return { lat: overlay.getBounds().getNorthEast().lat(), lng: overlay.getBounds().getCenter().lng() };

        let path = null;
        if (overlay.getPaths) path = overlay.getPaths().getAt(0).getArray();
        else if (overlay.getPath) path = overlay.getPath().getArray();

        if (!path || path.length === 0) return null;
        let top = path[0]; path.forEach(p => { if (p.lat() > top.lat()) top = p; });
        return top;
    },

    selectOverlay(overlay, multi = false) {
        if (!multi && !this.selectedOverlays.includes(overlay)) this.clearSelection();
        if (this.selectedOverlays.includes(overlay)) { if (multi) this.deselectOverlay(overlay); return; }
        this.selectedOverlays.push(overlay);
        if (overlay.setOptions) {
            if (overlay.getPosition && !overlay.getPath) { // Marker/Point
                const icon = overlay.getIcon();
                if (icon && typeof icon === 'object') {
                    overlay.originalIcon = overlay.originalIcon || { ...icon };
                    overlay.setIcon({
                        ...icon,
                        strokeWeight: 4,
                        strokeColor: '#FFFFFF',
                        scale: icon.scale + 2
                    });
                }
                overlay.setOptions({ zIndex: 100 });
            } else {
                let style = this.styles.area; if (overlay.category === 'constraint') style = this.styles.constraint; if (overlay.category === 'setback') style = this.styles.setback;
                const isMoveMode = document.getElementById('btn-move-toggle')?.classList.contains('active');
                overlay.setOptions({
                    strokeColor: style.selectedColor,
                    strokeWeight: 4,
                    zIndex: 10,
                    editable: !isMoveMode,
                    draggable: isMoveMode
                });
            }
        }
        document.getElementById('floating-edit-bar').classList.remove('hidden');
        document.getElementById('btn-move-toggle').classList.toggle('active', !!overlay.get('draggable'));

        this.updateToolbarVisibility();
    },

    updateToolbarVisibility() {
        if (this.selectedOverlays.length === 0) return;

        const hasArea = this.selectedOverlays.some(o => o.category === 'area');
        const gridBtn = document.getElementById('btn-grid');
        const copyBtn = document.getElementById('btn-copy');
        const rotateBtn = document.getElementById('btn-rotate-toggle');
        const alignBtn = document.getElementById('btn-align-menu');
        const distBtn = document.getElementById('btn-distribute-menu');
        const pasteBtn = document.getElementById('btn-paste-toggle');
        const dividers = document.querySelectorAll('.edit-bar-divider');

        const showExtended = true; // Show all tools for all selected shapes

        if (gridBtn) gridBtn.style.display = (showExtended && !hasArea) ? 'flex' : 'none';
        if (copyBtn) copyBtn.style.display = (showExtended && !hasArea) ? 'flex' : 'none';
        if (rotateBtn) rotateBtn.style.display = 'flex'; // Always show rotation if something is selected
        if (alignBtn) alignBtn.style.display = 'flex';
        if (distBtn) distBtn.style.display = 'flex';
        if (pasteBtn) pasteBtn.style.display = 'flex';

        dividers.forEach((d) => {
            d.style.display = 'block';
        });

        // Ensure rotation slider hides if its toggle is hidden
        if (!showExtended) {
            document.getElementById('rotation-popup')?.classList.add('hidden');
            rotateBtn?.classList.remove('active');
        }
    },

    deselectOverlay(overlay) {
        if (overlay.getPosition && !overlay.getPath) {
            if (overlay.originalIcon) overlay.setIcon(overlay.originalIcon);
            overlay.setOptions({ draggable: false, zIndex: 10 });
        } else {
            let style = this.styles.area; if (overlay.category === 'constraint') style = this.styles.constraint; if (overlay.category === 'setback') style = this.styles.setback;
            if (overlay.setOptions) overlay.setOptions({ strokeColor: style.strokeColor, strokeWeight: style.strokeWeight, zIndex: style.zIndex, editable: false, draggable: false });
        }
        this.selectedOverlays = this.selectedOverlays.filter(o => o !== overlay);
        if (this.selectedOverlays.length > 0) this.updateToolbarVisibility();
        if (this.selectedOverlays.length === 0) {
            document.getElementById('floating-edit-bar').classList.add('hidden');
            document.getElementById('rotation-popup').classList.add('hidden');
            document.getElementById('btn-rotate-toggle')?.classList.remove('active');
        }
    },

    clearSelection() {
        this.selectedOverlays.forEach(overlay => {
            if (overlay.getPosition && !overlay.getPath) {
                if (overlay.originalIcon) overlay.setIcon(overlay.originalIcon);
                overlay.setOptions({ draggable: false, zIndex: 10 });
            } else {
                let style = this.styles.area; if (overlay.category === 'constraint') style = this.styles.constraint; if (overlay.category === 'setback') style = this.styles.setback;
                if (overlay.setOptions) overlay.setOptions({ strokeColor: style.strokeColor, strokeWeight: style.strokeWeight, zIndex: style.zIndex, editable: false, draggable: false });
            }
        });
        this.selectedOverlays = [];
        
        // Global Map Click Reset: Force hide all vertex editing handles for EVERY object on the map
        if (this.overlays) {
            this.overlays.forEach(overlay => {
                if (overlay.setEditable) overlay.setEditable(false);
            });
        }
        if (window.CablesEngine && window.CablesEngine.mapPolylines) {
            window.CablesEngine.mapPolylines.forEach(p => {
                if (p.setEditable) p.setEditable(false);
            });
        }
        
        document.getElementById('floating-edit-bar').classList.add('hidden');
        document.getElementById('rotation-popup').classList.add('hidden');

        const slider = document.getElementById('rotation-slider');
        if (slider) {
            slider.value = 0;
            this.totalRotation = 0;
            const valDisplay = document.getElementById('rotation-value');
            if (valDisplay) valDisplay.textContent = '0°';
        }
        document.getElementById('btn-rotate-toggle')?.classList.remove('active');

        this.closeAllMenus();
        this.pasteMode = false;
        document.getElementById('btn-paste-toggle').classList.remove('active');
        this.map.setOptions({ draggableCursor: null });
    },

    copy() {
        if (this.selectedOverlays.length === 0 || this.selectedOverlays.some(o => o.category === 'area')) return;
        this.clipboard = this.selectedOverlays.map(o => ({
            category: o.category, subType: o.subType, type: o.type, geojson: this.getGeoJSON(o),
            style: o.category === 'area' ? this.styles.area : (o.category === 'constraint' ? this.styles.constraint : this.styles.setback)
        }));
    },

    togglePasteMode() {
        if (!this.clipboard) { alert("Copy something first"); return; }
        this.pasteMode = !this.pasteMode;
        document.getElementById('btn-paste-toggle').classList.toggle('active', this.pasteMode);
        this.map.setOptions({ draggableCursor: this.pasteMode ? 'copy' : null });
    },

    paste(latLng) {
        if (!this.clipboard) return;
        const features = this.clipboard.map(c => c.geojson);
        const groupGeo = turf.featureCollection(features);
        const groupCenter = turf.center(groupGeo);
        const targetPoint = turf.point([latLng.lng(), latLng.lat()]);
        const distance = turf.distance(groupCenter, targetPoint, { units: 'meters' });
        const bearing = turf.bearing(groupCenter, targetPoint);
        this.clipboard.forEach(item => {
            const finalGeo = turf.transformTranslate(item.geojson, distance, bearing, { units: 'meters' });
            this.spawnFromGeoJSON(finalGeo, item.category, item.subType, item.style);
        });
        this.pasteMode = false;
        document.getElementById('btn-paste-toggle').classList.remove('active');
        this.map.setOptions({ draggableCursor: null });
    },

    generateGrid() {
        if (this.selectedOverlays.length === 0 || this.selectedOverlays.some(o => o.category === 'area')) return;
        const xCount = parseInt(document.getElementById('grid-x-count').value);
        const yCount = parseInt(document.getElementById('grid-y-count').value);
        const xSpace = parseFloat(document.getElementById('grid-x-space').value);
        const ySpace = parseFloat(document.getElementById('grid-y-space').value);
        this.selectedOverlays.forEach(overlay => {
            const baseGeo = this.getGeoJSON(overlay);
            const style = overlay.category === 'area' ? this.styles.area : this.styles.constraint;
            for (let i = 0; i < xCount; i++) {
                for (let j = 0; j < yCount; j++) {
                    if (i === 0 && j === 0) continue;
                    let copy = turf.transformTranslate(baseGeo, i * xSpace, 90, { units: 'meters' });
                    copy = turf.transformTranslate(copy, j * ySpace, 180, { units: 'meters' });
                    this.spawnFromGeoJSON(copy, overlay.category, overlay.subType, style);
                }
            }
        });
    },

    spawnFromGeoJSON(geojson, category, subType, style) {
        if (!geojson || !geojson.geometry) return;
        let paths = [];
        if (geojson.geometry.type === 'Polygon') {
            paths = geojson.geometry.coordinates.map(ring => ring.map(c => ({ lat: c[1], lng: c[0] })));
        } else if (geojson.geometry.type === 'MultiPolygon') {
            paths = geojson.geometry.coordinates[0].map(ring => ring.map(c => ({ lat: c[1], lng: c[0] })));
        } else return;

        const polygon = new google.maps.Polygon({ paths: paths, ...style, map: this.map });
        polygon.category = category; polygon.subType = subType; polygon.type = google.maps.drawing.OverlayType.POLYGON;
        this.processNewOverlay(polygon);
    },

    handleLayoutAction(action) {
        if (this.selectedOverlays.length < 2) { alert("Select at least 2 items to align/distribute"); return; }
        const bounds = this.selectedOverlays.map(o => {
            const b = new google.maps.LatLngBounds();
            if (o.getPosition) b.extend(o.getPosition());
            else if (o.getCenter) b.extend(o.getCenter());
            else if (o.getPaths) o.getPaths().forEach(p => p.forEach(pt => b.extend(pt)));
            else if (o.getPath) o.getPath().forEach(p => b.extend(p));
            return { b, center: b.getCenter(), ne: b.getNorthEast(), sw: b.getSouthWest() };
        });

        if (action.startsWith('align-')) {
            let target;
            if (action === 'align-top') target = Math.max(...bounds.map(b => b.ne.lat()));
            if (action === 'align-bottom') target = Math.min(...bounds.map(b => b.sw.lat()));
            if (action === 'align-left') target = Math.min(...bounds.map(b => b.sw.lng()));
            if (action === 'align-right') target = Math.max(...bounds.map(b => b.ne.lng()));
            if (action === 'align-center') target = bounds.reduce((a, b) => a + b.center.lat(), 0) / bounds.length;

            this.selectedOverlays.forEach((o, i) => {
                const b = bounds[i];
                let dl = 0, dg = 0;
                if (action === 'align-top') dl = target - b.ne.lat();
                if (action === 'align-bottom') dl = target - b.sw.lat();
                if (action === 'align-left') dg = target - b.sw.lng();
                if (action === 'align-right') dg = target - b.ne.lng();
                if (action === 'align-center') dl = target - b.center.lat();
                this.moveOverlayRelative(o, dl, dg);
            });
        } else if (action === 'dist-h' || action === 'dist-v') {
            const dim = action === 'dist-h' ? 'lng' : 'lat';
            const sorted = [...this.selectedOverlays].sort((x, y) => {
                const getP = (o) => o.getPosition ? o.getPosition()[dim]() : (o.getCenter ? o.getCenter()[dim]() : (o.getPaths ? o.getPaths().getAt(0).getAt(0)[dim]() : o.getPath().getAt(0)[dim]()));
                return getP(x) - getP(y);
            });

            const getP = (o) => o.getPosition ? o.getPosition() : (o.getCenter ? o.getCenter() : (o.getPaths ? o.getPaths().getAt(0).getAt(0) : o.getPath().getAt(0)));
            const start = getP(sorted[0]);
            const end = getP(sorted[sorted.length - 1]);
            const total = action === 'dist-h' ? end.lng() - start.lng() : end.lat() - start.lat();
            const step = total / (sorted.length - 1);

            sorted.forEach((o, i) => {
                if (i === 0 || i === sorted.length - 1) return;
                const current = getP(o);
                const targetVal = (action === 'dist-h' ? start.lng() : start.lat()) + (i * step);
                const delta = targetVal - (action === 'dist-h' ? current.lng() : current.lat());
                if (action === 'dist-h') this.moveOverlayRelative(o, 0, delta);
                else this.moveOverlayRelative(o, delta, 0);
            });
        }
        this.closeAllMenus();
    },

    moveOverlayRelative(o, dl, dg) {
        if (o.getPaths) {
            const newPaths = o.getPaths().getArray().map(p => p.getArray().map(pt => ({ lat: pt.lat() + dl, lng: pt.lng() + dg })));
            o.setPaths(newPaths);
        } else if (o.getPath) {
            const newPath = o.getPath().getArray().map(pt => ({ lat: pt.lat() + dl, lng: pt.lng() + dg }));
            o.setPath(newPath);
        } else if (o.setPosition) {
            const p = o.getPosition(); o.setPosition({ lat: p.lat() + dl, lng: p.lng() + dg });
        } else if (o.setCenter) {
            const p = o.getCenter(); o.setCenter({ lat: p.lat() + dl, lng: p.lng() + dg });
        }
        if (o.category === 'area' || o.subType) this.updateAreaLabel(o);
    },

    getGeoJSON(overlay) {
        if (overlay.type === google.maps.drawing.OverlayType.CIRCLE || (overlay.getRadius && overlay.getCenter)) {
            const center = overlay.getCenter();
            return turf.circle([center.lng(), center.lat()], overlay.getRadius(), { steps: 64, units: 'meters' });
        }
        if (overlay.getPaths) {
            const coords = overlay.getPaths().getArray().map(path => {
                const ring = path.getArray().map(p => [p.lng(), p.lat()]);
                ring.push(ring[0]);
                return ring;
            });
            return turf.polygon(coords);
        }
        if (overlay.getPath) {
            const ring = overlay.getPath().getArray().map(p => [p.lng(), p.lat()]);
            ring.push(ring[0]);
            return turf.polygon([ring]);
        }
        if (overlay.getBounds) {
            const b = overlay.getBounds(); const ne = b.getNorthEast(); const sw = b.getSouthWest();
            const ring = [[sw.lng(), ne.lat()], [ne.lng(), ne.lat()], [ne.lng(), sw.lat()], [sw.lng(), sw.lat()], [sw.lng(), ne.lat()]];
            return turf.polygon([ring]);
        }
        if (overlay.getPosition) return turf.point([overlay.getPosition().lng(), overlay.getPosition().lat()]);
        return null;
    },

    async deleteSelected() {
        if (this.selectedOverlays.length > 0) {
            if (window.AppConfirm) {
                const proceed = await window.AppConfirm(`Are you sure you want to delete ${this.selectedOverlays.length} selected item(s)?`, 'Delete', 'danger');
                if (!proceed) return;
            }
            this.selectedOverlays.forEach(o => {
                // If it's an area, clear the layout first
                if (o.category === 'area' || (o.getPath && !o.subType)) {
                    if (window.LayoutEngine) {
                        window.LayoutEngine._clearOldLayout(o);
                    }
                }
                o.setMap(null);
                if (o.labelMarker) o.labelMarker.setMap(null);
                this.overlays = this.overlays.filter(item => item !== o);
            });
            this.clearSelection();
        }
    },

    toggleMoveMode() {
        if (this.selectedOverlays.length === 0) return;
        const currentOverlay = this.selectedOverlays[0];
        const drag = !currentOverlay.get('draggable');

        this.selectedOverlays.forEach(o => {
            if (o.setDraggable) o.setDraggable(drag); // Markers
            if (o.setEditable) o.setEditable(!drag); // Polygons

            // Fallback for setOptions
            if (o.setOptions) {
                o.setOptions({ draggable: drag, editable: !drag });
            }
        });
        document.getElementById('btn-move-toggle').classList.toggle('active', drag);
    },

    rotateSelected(angle) {
        if (this.selectedOverlays.length === 0) return;
        const gos = this.selectedOverlays.filter(o => this.getGeoJSON(o)).map(o => this.getGeoJSON(o));
        if (gos.length === 0) return;
        const pivot = turf.center(turf.featureCollection(gos));
        this.selectedOverlays.forEach(overlay => {
            const geojson = this.getGeoJSON(overlay); if (!geojson) return;
            const rotated = turf.transformRotate(geojson, angle, { pivot: pivot });
            if (overlay.getPaths) {
                const newPaths = rotated.geometry.coordinates.map(ring => ring.map(c => ({ lat: c[1], lng: c[0] })));
                overlay.setPaths(newPaths);
            } else if (overlay.setPath) {
                const coords = rotated.geometry.coordinates[0].map(c => ({ lat: c[1], lng: c[0] })); coords.pop(); overlay.setPath(coords);
            } else if (overlay.getBounds && !overlay.getPosition) {
                // Compatibility for older rectangles: Convert to polygon on the fly
                const style = overlay.category === 'area' ? this.styles.area : this.styles.constraint;
                const newPaths = rotated.geometry.coordinates.map(ring => ring.map(c => ({ lat: c[1], lng: c[0] })));
                const poly = new google.maps.Polygon({
                    paths: newPaths,
                    ...style,
                    strokeColor: style.selectedColor,
                    strokeWeight: 4,
                    map: this.map,
                    draggable: overlay.get('draggable'),
                    editable: overlay.get('editable'),
                    zIndex: 10
                });
                poly.type = google.maps.drawing.OverlayType.POLYGON;
                poly.category = overlay.category;
                poly.subType = overlay.subType;
                if (overlay.labelMarker) { poly.labelMarker = overlay.labelMarker; overlay.labelMarker = null; }

                const mainIdx = this.overlays.indexOf(overlay);
                if (mainIdx !== -1) this.overlays[mainIdx] = poly;
                const selIdx = this.selectedOverlays.indexOf(overlay);
                if (selIdx !== -1) this.selectedOverlays[selIdx] = poly;

                overlay.setMap(null);
                overlay.addListener = () => { };
                this.attachOverlayListeners(poly);
            } else if (overlay.setPosition) {
                const newPos = rotated.geometry.coordinates; overlay.setPosition({ lat: newPos[1], lng: newPos[0] });
            }
            if (overlay.category === 'area' || overlay.subType) this.updateAreaLabel(overlay);
        });
    },

    createTopographyMesh() {
        const input = document.getElementById('topo-data-input').value;
        if (!input) { alert("Please paste topography data first."); return; }

        const lines = input.split('\n').filter(l => l.trim());
        this.topoPoints = lines.map(line => {
            const parts = line.split(/[,\t\s]+/).filter(p => !isNaN(p));
            if (parts.length < 3) return null;
            return { lat: parseFloat(parts[0]), lng: parseFloat(parts[1]), elev: parseFloat(parts[2]) };
        }).filter(p => p);

        if (this.topoPoints.length < 3) { alert("At least 3 valid points are required."); return; }

        try {
            const points = turf.featureCollection(this.topoPoints.map(p => turf.point([p.lng, p.lat], { z: p.elev })));
            const tin = turf.tin(points, 'z');

            this.topoMesh = tin.features.map(f => ({
                geometry: f.geometry,
                properties: f.properties
            }));

            alert(`Terrain model generated with ${this.topoMesh.length} zones. Ready for analysis.`);
        } catch (e) {
            console.error(e);
            alert("Mesh generation failed. Check data format.");
        }
    },

    runTopographyAnalysis() {
        if (!this.topoMesh || this.topoMesh.length === 0) {
            if (this.topoPoints && this.topoPoints.length >= 3) this.createTopographyMesh();
            else { alert("Import topography data first."); return; }
        }

        const maxSlope = parseFloat(document.getElementById('topo-max-slope').value);
        const applyGrading = document.getElementById('topo-apply-grading').checked;
        const gradingSlopeThreshold = parseFloat(document.getElementById('topo-grading-slope').value);

        // Reset existing topo constraints before analysis
        this.clearTopoConstraints();

        let failedCount = 0;
        let totalCut = 0;
        let totalFill = 0;

        this.topoMesh.forEach(triangle => {
            const coords = triangle.geometry.coordinates[0];

            // Extract points and elevations.
            // Turf.tin often places elevation in the 3rd coordinate (Z) of each vertex.
            const pts = coords.slice(0, 3).map(c => {
                // Try Z coordinate first
                if (c.length >= 3) return { lat: c[1], lng: c[0], elev: c[2] };
                // Fallback to searching points
                return this.topoPoints.find(p => Math.abs(p.lng - c[0]) < 0.0001 && Math.abs(p.lat - c[1]) < 0.0001);
            });

            if (pts.some(p => !p)) return;
            const [p1, p2, p3] = pts;

            // Calculate Approximate Slope (%)
            const d12 = turf.distance([p1.lng, p1.lat], [p2.lng, p2.lat], { units: 'meters' });
            const d23 = turf.distance([p2.lng, p2.lat], [p3.lng, p3.lat], { units: 'meters' });
            const d31 = turf.distance([p3.lng, p3.lat], [p1.lng, p1.lat], { units: 'meters' });

            const s12 = (Math.abs(p1.elev - p2.elev) / (d12 || 1)) * 100;
            const s23 = (Math.abs(p2.elev - p3.elev) / (d23 || 1)) * 100;
            const s31 = (Math.abs(p3.elev - p1.elev) / (d31 || 1)) * 100;
            const avgSlope = (s12 + s23 + s31) / 3;

            // NEW LOGIC: Grade it if it's LESS than the threshold
            const isGradingZone = applyGrading && avgSlope < gradingSlopeThreshold;

            // Exclude it if it's MORE than the max slope AND we aren't grading it
            if (avgSlope > maxSlope && !isGradingZone) {
                const mapCoords = coords.map(c => ({ lat: c[1], lng: c[0] }));
                const constraint = new google.maps.Polygon({
                    paths: mapCoords,
                    ...this.styles.topo_excluded,
                    map: this.map
                });
                constraint.category = 'constraint';
                constraint.subType = 'topo_exclusion';
                constraint.areaName = `Slope Excl (${avgSlope.toFixed(1)}%)`;
                this.processNewOverlay(constraint);
                failedCount++;
            }

            // Volume calculation for grading zone
            if (isGradingZone) {
                const area = turf.area(triangle.geometry);
                const avgElev = (p1.elev + p2.elev + p3.elev) / 3;
                const deviations = [p1.elev - avgElev, p2.elev - avgElev, p3.elev - avgElev];
                deviations.forEach(d => {
                    const vol = (Math.abs(d) * (area / 3));
                    if (d > 0) totalCut += vol;
                    else totalFill += vol;
                });
            }
        });

        // Update UI
        document.getElementById('topo-stats').classList.remove('hidden');
        document.getElementById('topo-failed-count').innerText = failedCount;

        if (applyGrading) {
            document.getElementById('grading-stats').classList.remove('hidden');
            document.getElementById('topo-cut-vol').innerText = Math.round(totalCut).toLocaleString();
            document.getElementById('topo-fill-vol').innerText = Math.round(totalFill).toLocaleString();
        } else {
            document.getElementById('grading-stats').classList.add('hidden');
        }
    },

    clearTopoConstraints() {
        // If an area is selected, only clear inside it. Else, clear all topo exclusions.
        let toRemove = [];
        const selection = this.selectedOverlays[0];

        if (selection && selection.category === 'area') {
            const areaGeo = this.getGeoJSON(selection);
            toRemove = this.overlays.filter(o => {
                if (o.subType !== 'topo_exclusion') return false;
                const pt = this.getLabelPosition(o); // Use centroid for intersection check
                if (!pt) return false;
                const point = turf.point([pt.lng(), pt.lat()]);
                return turf.booleanPointInPolygon(point, areaGeo);
            });
        } else {
            toRemove = this.overlays.filter(o => o.subType === 'topo_exclusion');
        }

        toRemove.forEach(o => {
            o.setMap(null);
            if (o.labelMarker) o.labelMarker.setMap(null);
            this.overlays = this.overlays.filter(item => item !== o);
        });

        // Update stats
        const failedDisplay = document.getElementById('topo-failed-count');
        if (failedDisplay) failedDisplay.innerText = this.overlays.filter(o => o.subType === 'topo_exclusion').length;
    },

    createAreaFromCoords() {
        const input = document.getElementById('area-coords-input').value;
        if (!input) return;
        const lines = input.split('\n').filter(l => l.trim());
        const coords = lines.map(line => {
            const parts = line.split(/[,\t\s]+/).filter(p => !isNaN(p));
            if (parts.length < 2) return null;
            return { lat: parseFloat(parts[0]), lng: parseFloat(parts[1]) };
        }).filter(c => c);

        if (coords.length < 3) { alert("At least 3 valid coordinates are required."); return; }

        const poly = new google.maps.Polygon({ paths: coords, ...this.styles.area, map: this.map });
        poly.category = 'area'; poly.type = google.maps.drawing.OverlayType.POLYGON;
        this.processNewOverlay(poly);

        // Auto-fit to new area
        const bounds = new google.maps.LatLngBounds();
        coords.forEach(c => bounds.extend(c));
        this.map.fitBounds(bounds);
    },

    importKmlKmz(file) {
        if (!file) return;
        const ext = file.name.split('.').pop().toLowerCase();

        const processKmlText = (text) => {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(text, "text/xml");

            // Look for polygons specifically, often within <Polygon> -> <outerBoundaryIs> -> <LinearRing> -> <coordinates>
            const polyTags = xmlDoc.getElementsByTagName("Polygon");
            let areaCreated = false;
            let combinedBounds = new google.maps.LatLngBounds();

            if (polyTags.length > 0) {
                for (let i = 0; i < polyTags.length; i++) {
                    const coordsNode = polyTags[i].getElementsByTagName("coordinates")[0];
                    if (!coordsNode) continue;

                    const coordString = coordsNode.textContent.trim();
                    const points = coordString.split(/\s+/);

                    const coords = points.map(p => {
                        const parts = p.split(',').map(Number);
                        if (parts.length >= 2) return { lat: parts[1], lng: parts[0] };
                        return null;
                    }).filter(c => c);

                    if (coords.length >= 3) {
                        const poly = new google.maps.Polygon({ paths: coords, ...this.styles.area, map: this.map });
                        poly.category = 'area';
                        poly.type = google.maps.drawing.OverlayType.POLYGON;
                        this.processNewOverlay(poly);
                        areaCreated = true;

                        coords.forEach(c => combinedBounds.extend(c));
                    }
                }
            } else {
                // Fallback to purely looking for raw coordinates tags if no Polygon wrapping
                const coordsTags = xmlDoc.getElementsByTagName("coordinates");
                for (let i = 0; i < coordsTags.length; i++) {
                    const coordString = coordsTags[i].textContent.trim();
                    const points = coordString.split(/\s+/);

                    const coords = points.map(p => {
                        const parts = p.split(',').map(Number);
                        if (parts.length >= 2) return { lat: parts[1], lng: parts[0] };
                        return null;
                    }).filter(c => c);

                    if (coords.length >= 3) {
                        const poly = new google.maps.Polygon({ paths: coords, ...this.styles.area, map: this.map });
                        poly.category = 'area';
                        poly.type = google.maps.drawing.OverlayType.POLYGON;
                        this.processNewOverlay(poly);
                        areaCreated = true;

                        coords.forEach(c => combinedBounds.extend(c));
                    }
                }
            }

            if (!areaCreated) {
                alert("No valid polygonal coordinates found in the file.");
            } else {
                this.map.fitBounds(combinedBounds);
            }
        };

        if (ext === 'kmz') {
            if (typeof JSZip === 'undefined') {
                alert("KMZ processing engine is currently loading. Please try again in a few seconds.");
                return;
            }
            const reader = new FileReader();
            reader.onload = function (e) {
                JSZip.loadAsync(e.target.result).then(function (zip) {
                    const kmlFile = Object.keys(zip.files).find(name => name.toLowerCase().endsWith('.kml'));
                    if (kmlFile) {
                        zip.files[kmlFile].async("string").then(function (kmlText) {
                            processKmlText(kmlText);
                        });
                    } else {
                        alert("Error: No inner KML document found inside the KMZ archive.");
                    }
                }).catch(function (err) {
                    alert("Error extracting KMZ file: " + err.message);
                });
            };
            reader.readAsArrayBuffer(file);
        } else if (ext === 'kml') {
            const reader = new FileReader();
            reader.onload = function (e) {
                processKmlText(e.target.result);
            };
            reader.readAsText(file);
        } else {
            alert("Unsupported format. Please select a .kml or .kmz file.");
        }
    },

    /**
     * Imports all features from a KML/KMZ file as Obstacle overlays.
     * - Polygon placemarks  → red polygon area constraints
     * - Point placemarks    → red dot marker objects
     * Placemark <name> tags are used for auto-labelling each feature.
     */
    importKmlKmzAsObstacle(file) {
        if (!file) return;
        const ext = file.name.split('.').pop().toLowerCase();
        const self = this;

        const processKmlAsObstacles = (text) => {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(text, "text/xml");

            let polyCount = 0;
            let pointCount = 0;
            const combinedBounds = new google.maps.LatLngBounds();

            // Helper: parse a coordinate string into {lat, lng} objects
            const parseCoordString = (str) => str.trim().split(/\s+/).map(p => {
                const parts = p.split(',').map(Number);
                return parts.length >= 2 ? { lat: parts[1], lng: parts[0] } : null;
            }).filter(c => c && !isNaN(c.lat) && !isNaN(c.lng));

            // Walk every Placemark to pick up name + geometry
            const placemarks = xmlDoc.getElementsByTagName("Placemark");

            if (placemarks.length > 0) {
                for (let i = 0; i < placemarks.length; i++) {
                    const pm = placemarks[i];
                    // Read optional name
                    const nameEl = pm.getElementsByTagName("name")[0];
                    const label = nameEl ? nameEl.textContent.trim() : null;

                    // --- Polygon obstacle ---
                    const polyEl = pm.getElementsByTagName("Polygon")[0];
                    if (polyEl) {
                        const coordsNode = polyEl.getElementsByTagName("coordinates")[0];
                        if (coordsNode) {
                            const coords = parseCoordString(coordsNode.textContent);
                            if (coords.length >= 3) {
                                const poly = new google.maps.Polygon({
                                    paths: coords, ...self.styles.constraint, map: self.map
                                });
                                poly.category = 'constraint';
                                poly.type = google.maps.drawing.OverlayType.POLYGON;
                                if (label) poly.areaName = label;
                                self.processNewOverlay(poly);
                                coords.forEach(c => combinedBounds.extend(c));
                                polyCount++;
                            }
                        }
                        continue; // Move on to next placemark
                    }

                    // --- Point obstacle (rendered as dot marker) ---
                    const pointEl = pm.getElementsByTagName("Point")[0];
                    if (pointEl) {
                        const coordsNode = pointEl.getElementsByTagName("coordinates")[0];
                        if (coordsNode) {
                            const parts = coordsNode.textContent.trim().split(',').map(Number);
                            if (parts.length >= 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
                                const position = { lat: parts[1], lng: parts[0] };
                                const marker = new google.maps.Marker({
                                    position,
                                    map: self.map,
                                    title: label || 'Obstacle Object',
                                    icon: {
                                        path: google.maps.SymbolPath.CIRCLE,
                                        scale: 7,
                                        fillColor: '#ef4444',
                                        fillOpacity: 1,
                                        strokeColor: '#ffffff',
                                        strokeWeight: 2
                                    }
                                });
                                marker.category = 'constraint';
                                marker.subType = 'obstacle_point';
                                marker.type = google.maps.drawing.OverlayType.MARKER;
                                if (label) marker.areaName = label;
                                self.processNewOverlay(marker);
                                combinedBounds.extend(position);
                                pointCount++;
                            }
                        }
                    }
                }

                // Fallback: if Placemarks had no geometry, try raw tags
                if (polyCount === 0 && pointCount === 0) {
                    const coordsTags = xmlDoc.getElementsByTagName("coordinates");
                    for (let i = 0; i < coordsTags.length; i++) {
                        const coords = parseCoordString(coordsTags[i].textContent);
                        if (coords.length >= 3) {
                            const poly = new google.maps.Polygon({
                                paths: coords, ...self.styles.constraint, map: self.map
                            });
                            poly.category = 'constraint';
                            poly.type = google.maps.drawing.OverlayType.POLYGON;
                            self.processNewOverlay(poly);
                            coords.forEach(c => combinedBounds.extend(c));
                            polyCount++;
                        }
                    }
                }
            } else {
                // No Placemarks at all — scan raw Polygon tags as before
                const polyTags = xmlDoc.getElementsByTagName("Polygon");
                for (let i = 0; i < polyTags.length; i++) {
                    const coordsNode = polyTags[i].getElementsByTagName("coordinates")[0];
                    if (!coordsNode) continue;
                    const coords = parseCoordString(coordsNode.textContent);
                    if (coords.length >= 3) {
                        const poly = new google.maps.Polygon({
                            paths: coords, ...self.styles.constraint, map: self.map
                        });
                        poly.category = 'constraint';
                        poly.type = google.maps.drawing.OverlayType.POLYGON;
                        self.processNewOverlay(poly);
                        coords.forEach(c => combinedBounds.extend(c));
                        polyCount++;
                    }
                }
            }

            const total = polyCount + pointCount;
            if (total === 0) {
                alert("No valid obstacle geometry found in the file.");
            } else {
                const parts = [];
                if (polyCount > 0) parts.push(`${polyCount} polygon area${polyCount > 1 ? 's' : ''}`);
                if (pointCount > 0) parts.push(`${pointCount} point object${pointCount > 1 ? 's' : ''}`);
                alert(`Imported: ${parts.join(' + ')}.`);
                self.map.fitBounds(combinedBounds);
            }
        };

        if (ext === 'kmz') {
            if (typeof JSZip === 'undefined') {
                alert("KMZ processing engine is currently loading. Please try again in a few seconds.");
                return;
            }
            const reader = new FileReader();
            reader.onload = function (e) {
                JSZip.loadAsync(e.target.result).then(function (zip) {
                    const kmlFile = Object.keys(zip.files).find(name => name.toLowerCase().endsWith('.kml'));
                    if (kmlFile) {
                        zip.files[kmlFile].async("string").then(function (kmlText) {
                            processKmlAsObstacles(kmlText);
                        });
                    } else {
                        alert("Error: No inner KML document found inside the KMZ archive.");
                    }
                }).catch(function (err) {
                    alert("Error extracting KMZ file: " + err.message);
                });
            };
            reader.readAsArrayBuffer(file);
        } else if (ext === 'kml') {
            const reader = new FileReader();
            reader.onload = function (e) {
                processKmlAsObstacles(e.target.result);
            };
            reader.readAsText(file);
        } else {
            alert("Unsupported format. Please select a .kml or .kmz file.");
        }
    },


    applyMountingConfig() {
        const config = {
            type: document.getElementById('mounting-type-select').value,
            orientation: document.querySelector('#mounting-orientation-seg .seg-btn.active').dataset.value,
            tilt: parseFloat(document.getElementById('mount-tilt').value),
            azimuth: parseFloat(document.getElementById('mount-azimuth').value),
            moduleSpacingX: parseFloat(document.getElementById('mount-mod-dist-x').value),
            moduleSpacingY: parseFloat(document.getElementById('mount-mod-dist-y').value),
            tableConfig: {
                x: parseInt(document.getElementById('mount-mod-table-x').value),
                y: parseInt(document.getElementById('mount-mod-table-y').value),
                distX: parseFloat(document.getElementById('mount-table-dist-x').value),
                distY: parseFloat(document.getElementById('mount-table-dist-y').value)
            },
            blockConfig: {
                x: parseInt(document.getElementById('mount-block-x').value),
                y: parseInt(document.getElementById('mount-block-y').value),
                distX: parseFloat(document.getElementById('mount-block-dist-x').value),
                distY: parseFloat(document.getElementById('mount-block-dist-y').value)
            }
        };

        console.log("Mounting Configuration Applied:", config);
        alert(`Configuration for ${config.type} applied. Ready to generate layout.`);
    },

    /**
     * Captures specialized layout configuration and executes the generation algorithm
     */
    runLayoutGeneration() {
        const layoutConfig = {
            objective: document.querySelector('#layout-objective-seg .seg-btn.active').dataset.value,
            targetMW: parseFloat(document.getElementById('cfg-target-mw').value),
            divideMV: document.getElementById('cfg-divide-mv').checked,
            blockHierarchy: {
                everyX: parseInt(document.getElementById('cfg-dist-x-count').value),
                everyY: parseInt(document.getElementById('cfg-dist-y-count').value),
                gapX: parseFloat(document.getElementById('cfg-gap-dist-x').value),
                gapY: parseFloat(document.getElementById('cfg-gap-dist-y').value)
            },
            direction: document.getElementById('cfg-fill-direction').value,
            method: document.querySelector('#layout-method-seg .seg-btn.active').dataset.value,
            resolution: document.querySelector('#draw-resolution-seg .seg-btn.active').dataset.value
        };

        console.log("Starting Site Layout Generation with Config:", layoutConfig);

        if (this.selectedOverlays.length === 0) {
            alert("Please select a project area on the map first.");
            return;
        }

        alert("Layout generation algorithm started. Calculating optimal positions based on site geometry...");
        // Call spatial engine logic here
    },

    /**
     * Activates edge selection mode to align the layout to map geometry
     */
    startEdgeSelection() {
        this.map.setOptions({ draggableCursor: 'crosshair' });
        // Logic to listen for polyline clicks or closest segment matching
        console.log("Edge selection mode active.");
    }
};

window.SiteEngine = SiteEngine;
