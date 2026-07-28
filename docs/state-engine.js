/**
 * StateEngine - Handles Auto-Saving and Restoring Project State inside the Excel Document
 */
const StateEngine = {
    init() {
        console.log("State Engine Initialized");
        
        // Setup Auto-Save loop
        setInterval(() => {
            if (this._hasDataToSave()) {
                this.saveState(true); // silent auto-save
            }
        }, 15000); // Check every 15 seconds
    },

    _showSaveToast() {
        const toast = document.getElementById('save-toast');
        if (toast) {
            toast.classList.add('show');
            setTimeout(() => {
                toast.classList.remove('show');
            }, 2500); // Hide after 2.5 seconds
        }
    },

    _hasDataToSave() {
        if (!window.SiteEngine) return false;
        return window.SiteEngine.overlays.length > 0;
    },

    saveState(silent = false) {
        if (!window.SiteEngine || typeof Office === 'undefined' || !Office.context || !Office.context.document) return;

        try {
            const state = {
                version: "1.0.0", // Tracking for future software updates
                map: {
                    center: window.googleMap ? window.googleMap.getCenter().toJSON() : null,
                    zoom: window.googleMap ? window.googleMap.getZoom() : 16,
                },
                inputs: {},
                overlays: [],
                generatedAreaUids: []
            };

            // 0. Record which areas already had generated layouts
            if (window.LayoutEngine && window.LayoutEngine.layoutStore) {
                window.LayoutEngine.layoutStore.forEach((data, area) => {
                    if (area.__uid) state.generatedAreaUids.push(area.__uid);
                });
            }

            // 1. Save all relevant inputs (Exclude search bar)
            const inputsToSave = document.querySelectorAll('input, select');
            inputsToSave.forEach(el => {
                if (!el.id || el.id.includes('search') || el.id.includes('pac-input')) return;
                
                if (el.type === 'checkbox' || el.type === 'radio') {
                    state.inputs[el.id] = el.checked;
                } else {
                    state.inputs[el.id] = el.value;
                }
            });

            // 2. Save Overlays (Areas, Obstacles, Setbacks, POCs, etc.)
            window.SiteEngine.overlays.forEach(overlay => {
                if (overlay.category === 'measure') return; // Don't save temporary measure lines

                const ovData = {
                    category: overlay.category,
                    subType: overlay.subType,
                    type: overlay.type,
                    areaName: overlay.areaName,
                    __uid: overlay.__uid
                };

                // Polylines
                if (overlay.getPath && overlay.type === google.maps.drawing.OverlayType.POLYLINE) {
                    const path = overlay.getPath();
                    ovData.path = [];
                    for (let i = 0; i < path.getLength(); i++) {
                        ovData.path.push(path.getAt(i).toJSON());
                    }
                } 
                // Polygons (Areas, Obstacles, POCs, Stations)
                else if (overlay.getPaths && overlay.type === google.maps.drawing.OverlayType.POLYGON) {
                    const paths = overlay.getPaths();
                    ovData.paths = [];
                    for (let i = 0; i < paths.getLength(); i++) {
                        const path = paths.getAt(i);
                        const points = [];
                        for (let j = 0; j < path.getLength(); j++) {
                            points.push(path.getAt(j).toJSON());
                        }
                        ovData.paths.push(points);
                    }
                } 
                // Markers
                else if (overlay.getPosition && overlay.type === google.maps.drawing.OverlayType.MARKER) {
                    ovData.position = overlay.getPosition().toJSON();
                }

                state.overlays.push(ovData);
            });

            const stateStr = JSON.stringify(state);

            // 3. Avoid unnecessary saves if state hasn't changed
            const existing = Office.context.document.settings.get("enerxl_project_state");
            if (existing === stateStr) return;

            // 4. Save to Excel Document Settings
            Office.context.document.settings.set("enerxl_project_state", stateStr);
            Office.context.document.settings.saveAsync((res) => {
                if (res.status === Office.AsyncResultStatus.Succeeded) {
                    if (!silent) console.log("Project state auto-saved to Excel.");
                    this._showSaveToast();
                } else {
                    console.error("Failed to save state to Excel:", res.error);
                }
            });

        } catch(e) {
            console.error("Error saving project state:", e);
        }
    },

    loadState() {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.document) {
            console.log("Not running inside Excel. Cannot load document state.");
            return false; // Not in Excel
        }

        const stateStr = Office.context.document.settings.get("enerxl_project_state");
        if (!stateStr) {
            console.log("New document or no saved settings found.");
            return false;
        }

        try {
            const state = JSON.parse(stateStr);
            console.log(`Loading saved project state (Version: ${state.version || 'Legacy'})...`);

            // Future-proofing: Here we can add migration scripts if version > '1.0.0'
            // e.g. if (state.version === '1.0.0' && currentAppVersion === '2.0.0') { migrateConfig(state); }

            // 1. Restore Map View
            if (window.googleMap && state.map) {
                if (state.map.center) window.googleMap.setCenter(state.map.center);
                if (state.map.zoom) window.googleMap.setZoom(state.map.zoom);
            }

            // 2. Restore Inputs
            if (state.inputs) {
                for (const [id, value] of Object.entries(state.inputs)) {
                    const el = document.getElementById(id);
                    if (el) {
                        if (el.type === 'checkbox' || el.type === 'radio') {
                            el.checked = value;
                        } else {
                            el.value = value;
                        }
                    }
                }

                // Restore Electrical State
                if (window.ElectricalEngine) {
                    window.ElectricalEngine.updatePowerDisplay();
                    if (window.ElectricalEngine.updateStringSizing) {
                        window.ElectricalEngine.updateStringSizing();
                    }
                }
            }

            // 3. Restore Overlays
            if (window.SiteEngine && state.overlays) {
                // Clear any existing generic overlays just in case
                window.SiteEngine.overlays.forEach(ov => {
                    ov.setMap(null);
                    if (ov.labelMarker) ov.labelMarker.setMap(null);
                });
                window.SiteEngine.overlays = [];

                state.overlays.forEach(ovData => {
                    let overlay = null;
                    const style = window.SiteEngine.styles[ovData.category] || window.SiteEngine.styles.area;

                    if (ovData.type === google.maps.drawing.OverlayType.POLYGON) {
                        overlay = new google.maps.Polygon({
                            paths: ovData.paths,
                            ...style,
                            map: window.googleMap
                        });
                    } else if (ovData.type === google.maps.drawing.OverlayType.POLYLINE) {
                        overlay = new google.maps.Polyline({
                            path: ovData.path,
                            ...style,
                            map: window.googleMap
                        });
                    } else if (ovData.type === google.maps.drawing.OverlayType.MARKER) {
                        overlay = new google.maps.Marker({
                            position: ovData.position,
                            map: window.googleMap
                        });
                    }

                    if (overlay) {
                        overlay.category = ovData.category;
                        overlay.subType = ovData.subType;
                        overlay.type = ovData.type;
                        if (ovData.areaName) overlay.areaName = ovData.areaName;
                        if (ovData.__uid) overlay.__uid = ovData.__uid;

                        // Add back to SiteEngine
                        window.SiteEngine.overlays.push(overlay);
                        window.SiteEngine.attachOverlayListeners(overlay);
                        
                        // Force label render if needed
                        if ((overlay.category === 'area' || overlay.subType) && overlay.subType !== 'topo_exclusion') {
                            window.SiteEngine.updateAreaLabel(overlay);
                        }
                    }
                });
            }

            // 4. Auto-Regenerate Layouts
            if (state.generatedAreaUids && state.generatedAreaUids.length > 0) {
                // Short delay to let the map render overlays before Turf.js calculations
                setTimeout(() => {
                    if (window.SiteEngine && window.LayoutEngine) {
                        state.generatedAreaUids.forEach(uid => {
                            const area = window.SiteEngine.overlays.find(o => o.__uid === uid);
                            if (area) {
                                window.SiteEngine.clearSelection();
                                window.SiteEngine.selectOverlay(area, false);
                                window.LayoutEngine.generate();
                            }
                        });
                        // Clear final selection so user starts cleanly
                        window.SiteEngine.clearSelection();
                    }
                }, 800);
            }

            return true;
        } catch (e) {
            console.error("Error loading project state from Excel:", e);
            return false;
        }
    }
};

window.StateEngine = StateEngine;
if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            console.log("Excel Add-in Host Ready. Waiting for Engine Init...");
            // Ensure Map and core engines exist, then try loading.
            // A short delay to allow Maps API and SiteEngine to finish DOM rendering
            setTimeout(() => {
                StateEngine.loadState();
                StateEngine.init();
            }, 1000);
        }
    });
} else {
    console.log("Not running in Office environment.");
}

