
/**
 * LayersEngine - Handles map layer visibility and UI toggles
 */
const LayersEngine = {
    init() {
        this.setupListeners();
    },

    setupListeners() {
        // Toggle Panel
        const trigger = document.getElementById('layers-trigger');
        const panel = document.getElementById('layers-panel');
        if (trigger && panel) {
            trigger.addEventListener('click', (e) => {
                e.stopPropagation();
                panel.classList.toggle('hidden');
                // Auto-hide other panels if needed
                if (!panel.classList.contains('hidden')) {
                    document.getElementById('project-panel')?.classList.add('hidden');
                }
            });
        }
        
        // Prevent clicks inside panel from closing things globally if you have such listeners
        panel?.addEventListener('click', (e) => e.stopPropagation());

        // Layer Toggles
        document.querySelectorAll('.layer-item').forEach(item => {
            item.addEventListener('click', () => {
                const layer = item.dataset.layer;
                const isActive = item.classList.toggle('active');
                this.toggleLayer(layer, isActive);
                
                // Refresh Lucide icons for the newly state-changed icons if necessary
                if (window.lucide) lucide.createIcons();
            });
        });
    },

    toggleLayer(layer, show) {
        const map = show ? window.SiteEngine.map : null;

        switch (layer) {
            case 'areas':
                window.SiteEngine.overlays
                    .filter(o => o.category === 'area')
                    .forEach(o => {
                        o.setMap(map);
                        if (o.labelMarker) o.labelMarker.setMap(map);
                        
                        // Also show/hide associated tables
                        if (window.LayoutEngine && window.LayoutEngine.layoutStore) {
                            const related = window.LayoutEngine.layoutStore.get(o);
                            if (related) related.forEach(table => table.setMap(map));
                        }
                    });
                break;

            case 'obstacles':
                window.SiteEngine.overlays
                    .filter(o => o.category === 'constraint' && !['poc', 'station', 'topo_exclusion'].includes(o.subType))
                    .forEach(o => o.setMap(map));
                break;

            case 'setbacks':
                window.SiteEngine.overlays
                    .filter(o => o.category === 'setback')
                    .forEach(o => o.setMap(map));
                break;

            case 'lv_cables':
                if (window.CablesEngine && window.CablesEngine.mapPolylines) {
                    window.CablesEngine.mapPolylines
                        .filter(p => p.strokeColor === '#3b82f6')
                        .forEach(p => p.setMap(map));
                }
                break;

            case 'internal_mv':
                if (window.CablesEngine && window.CablesEngine.mapPolylines) {
                    window.CablesEngine.mapPolylines
                        .filter(p => p.strokeColor === '#10b981')
                        .forEach(p => p.setMap(map));
                }
                break;

            case 'outgoing_mv':
                if (window.CablesEngine && window.CablesEngine.mapPolylines) {
                    window.CablesEngine.mapPolylines
                        .filter(p => p.strokeColor === '#ef4444')
                        .forEach(p => p.setMap(map));
                }
                break;

            case 'inverters':
                if (window.LayoutEngine && window.LayoutEngine.inverters) {
                    window.LayoutEngine.inverters.forEach(inv => inv.setMap(map));
                }
                break;

            case 'mv_stations':
                if (window.LayoutEngine && window.LayoutEngine.blocks) {
                    window.LayoutEngine.blocks.forEach(block => {
                        block.setMap(map);
                        if (block.labelMarker) block.labelMarker.setMap(map);
                    });
                }
                break;

            case 'delivery_station':
                window.SiteEngine.overlays
                    .filter(o => o.subType === 'station')
                    .forEach(o => {
                        o.setMap(map);
                        if (o.labelMarker) o.labelMarker.setMap(map);
                    });
                break;

            case 'poc':
                window.SiteEngine.overlays
                    .filter(o => o.subType === 'poc')
                    .forEach(o => {
                        o.setMap(map);
                        if (o.labelMarker) o.labelMarker.setMap(map);
                    });
                break;
        }
    }
};

// Initialize after other engines are loaded
window.addEventListener('load', () => {
    LayersEngine.init();
});
