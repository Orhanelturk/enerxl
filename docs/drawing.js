
// Category styles for drawing
const categoryStyles = {
    area: {
        fillColor: '#286944',
        fillOpacity: 0.3,
        strokeColor: '#286944',
        strokeWeight: 2,
        clickable: true,
        editable: true,
        zIndex: 1
    },
    obstacle: {
        fillColor: '#ef4444',
        fillOpacity: 0.3,
        strokeColor: '#ef4444',
        strokeWeight: 2,
        clickable: true,
        editable: true,
        zIndex: 2
    },
    poc: {
        icon: {
            path: google.maps.SymbolPath.CIRCLE,
            fillColor: '#286944',
            fillOpacity: 1,
            strokeColor: '#ffffff',
            strokeWeight: 2,
            scale: 8
        }
    }
};

window.currentCategory = 'area';
window.currentShape = null;
window.overlays = [];
window.selectedOverlays = [];
window.__overlayRegistry = new Set();

window.initDrawing = function (map) {
    const dm = new google.maps.drawing.DrawingManager({
        drawingMode: null,
        drawingControl: false,
        polygonOptions: categoryStyles.area,
        rectangleOptions: categoryStyles.area,
        circleOptions: categoryStyles.area,
        markerOptions: {
            draggable: true
        }
    });
    dm.setMap(map);
    window.drawingManager = dm;

    google.maps.event.addListener(dm, 'overlaycomplete', function (event) {
        const ov = event.overlay;
        ov.__cat = window.currentCategory;
        ov.__selected = false;
        window.overlays.push(ov);
        window.__overlayRegistry.add(ov);

        // Re-select if needed or just stop drawing
        window.setDrawingMode(null);

        ov.addListener('click', () => {
            selectOverlay(ov);
        });
    });

    map.addListener('click', () => {
        clearSelection();
    });
};

window.setDrawingMode = function (shape) {
    window.currentShape = shape;
    if (!window.drawingManager) return;

    let mode = null;
    if (shape === 'polygon') mode = google.maps.drawing.OverlayType.POLYGON;
    if (shape === 'rectangle') mode = google.maps.drawing.OverlayType.RECTANGLE;
    if (shape === 'circle') mode = google.maps.drawing.OverlayType.CIRCLE;
    if (shape === 'marker' || shape === 'poc' || shape === 'station') mode = google.maps.drawing.OverlayType.MARKER;

    window.drawingManager.setDrawingMode(mode);

    // Update UI state
    document.querySelectorAll('.draw-tool-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.shape === shape);
    });
};

window.setDrawingCategory = function (cat) {
    window.currentCategory = cat;
    if (!window.drawingManager) return;

    const style = categoryStyles[cat] || categoryStyles.area;
    window.drawingManager.setOptions({
        polygonOptions: style,
        rectangleOptions: style,
        circleOptions: style
    });

    // Update UI state
    document.querySelectorAll('.cat-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.cat === cat);
    });
};

function selectOverlay(ov) {
    clearSelection();
    ov.__selected = true;

    // Save original style
    if (ov.get('strokeColor')) {
        ov.__oldStrokeColor = ov.get('strokeColor');
        ov.__oldStrokeWeight = ov.get('strokeWeight');

        ov.setOptions({
            strokeColor: '#ffb703',
            strokeWeight: 3
        });
    }
    window.selectedOverlays = [ov];
}

function clearSelection() {
    window.selectedOverlays.forEach(ov => {
        ov.__selected = false;
        if (ov.setOptions) {
            ov.setOptions({
                strokeColor: ov.__oldStrokeColor || (ov.__cat === 'area' ? '#286944' : '#ef4444'),
                strokeWeight: ov.__oldStrokeWeight || 2
            });
        }
    });
    window.selectedOverlays = [];
}
