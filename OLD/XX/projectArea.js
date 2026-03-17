window.ProjectAreaTool = {
  _map: null,
  _dm: null,
  _type: "pv",

  _pv: [],
  _obstacles: [],
  _objects: [],

  _areaCounter: 1,

  init(map) {
    this._map = map;

    const panel = document.getElementById("sub-bar-area");
    if (!panel) return;

    this._dm = new google.maps.drawing.DrawingManager({
      drawingMode: null,
      drawingControl: false,
      polygonOptions: { editable: true, clickable: true },
      rectangleOptions: { editable: true, clickable: true },
      circleOptions: { editable: true, clickable: true },
      polylineOptions: { editable: true, clickable: true }
    });

    this._dm.setMap(map);

    this._applyDrawingStyle();
    this._wirePanel();
    this._wireDrawingComplete();
  },

  _styleFor(type) {
    if (type === "obstacle") {
      return {
        fillColor: "#ef4444",
        fillOpacity: 0.18,
        strokeColor: "#7f1d1d",
        strokeOpacity: 1,
        strokeWeight: 2,
        zIndex: 12
      };
    }
    if (type === "object") {
      return {
        fillColor: "#60a5fa",
        fillOpacity: 0.16,
        strokeColor: "#1d4ed8",
        strokeOpacity: 1,
        strokeWeight: 2,
        zIndex: 11
      };
    }
    return {
      fillColor: "#facc15",
      fillOpacity: 0.18,
      strokeColor: "#facc15",
      strokeOpacity: 1,
      strokeWeight: 2,
      zIndex: 10
    };
  },

  _applyDrawingStyle() {
    if (!this._dm) return;
    const s = this._styleFor(this._type);

    this._dm.setOptions({
      polygonOptions: { editable: true, clickable: true, ...s },
      rectangleOptions: { editable: true, clickable: true, ...s },
      circleOptions: { editable: true, clickable: true, ...s },
      polylineOptions: {
        editable: true,
        clickable: true,
        strokeColor: s.strokeColor,
        strokeOpacity: s.strokeOpacity,
        strokeWeight: s.strokeWeight,
        zIndex: s.zIndex
      }
    });
  },

  _wirePanel() {
    const setType = (t) => {
      this._type = t;
      document.querySelectorAll("#sub-bar-area [data-pa-type]").forEach(b => {
        b.classList.toggle("active", b.dataset.paType === t);
      });
      this._applyDrawingStyle();
    };

    document.getElementById("pa-type-pv").onclick = () => setType("pv");
    document.getElementById("pa-type-ob").onclick = () => setType("obstacle");
    document.getElementById("pa-type-obj").onclick = () => setType("object");

    document.querySelectorAll("#sub-bar-area .pa-section").forEach(sec => {
      const head = sec.querySelector(".pa-section-head");
      if (head) head.onclick = () => sec.classList.toggle("collapsed");
    });

    const setDrawActive = (id) => {
      document.querySelectorAll("#sub-bar-area .pa-icon-btn").forEach(b => b.classList.remove("active"));
      const el = document.getElementById(id);
      if (el) el.classList.add("active");
    };

    document.getElementById("pa-draw-polygon").onclick = () => {
      setDrawActive("pa-draw-polygon");
      this._dm.setDrawingMode(google.maps.drawing.OverlayType.POLYGON);
    };
    document.getElementById("pa-draw-rect").onclick = () => {
      setDrawActive("pa-draw-rect");
      this._dm.setDrawingMode(google.maps.drawing.OverlayType.RECTANGLE);
    };
    document.getElementById("pa-draw-circle").onclick = () => {
      setDrawActive("pa-draw-circle");
      this._dm.setDrawingMode(google.maps.drawing.OverlayType.CIRCLE);
    };
    document.getElementById("pa-draw-line").onclick = () => {
      setDrawActive("pa-draw-line");
      this._dm.setDrawingMode(google.maps.drawing.OverlayType.POLYLINE);
    };

    document.getElementById("pa-stop").onclick = () => {
      this._dm.setDrawingMode(null);
      document.querySelectorAll("#sub-bar-area .pa-icon-btn").forEach(b => b.classList.remove("active"));
    };

    document.getElementById("pa-clear-type").onclick = () => {
      this._clearByType(this._type);
      this._dm.setDrawingMode(null);
      document.querySelectorAll("#sub-bar-area .pa-icon-btn").forEach(b => b.classList.remove("active"));
    };

    document.getElementById("pa-draw-coords").onclick = () => {
      const txt = document.getElementById("pa-coords").value || "";
      const pts = this._parseCoords(txt);
      if (pts.length < 3) return;

      this._clearByType("pv");

      const poly = new google.maps.Polygon({
        paths: pts,
        editable: true,
        clickable: true
      });
      poly.__paShape = "polygon";
      poly.setOptions(this._styleFor("pv"));
      poly.setMap(this._map);

      if (!poly.__name) poly.__name = "Area " + (this._areaCounter++);
      this._attachLabelAboveOverlay(poly);

      this._pv.push(poly);

      const b = new google.maps.LatLngBounds();
      pts.forEach(p => b.extend(p));
      this._map.fitBounds(b);
    };

    document.getElementById("pa-clear-coords").onclick = () => {
      document.getElementById("pa-coords").value = "";
    };

    setType("pv");
  },

  _wireDrawingComplete() {
    google.maps.event.addListener(this._dm, "polygoncomplete", (o) => this._storeOverlay(o, "polygon"));
    google.maps.event.addListener(this._dm, "rectanglecomplete", (o) => this._storeOverlay(o, "rectangle"));
    google.maps.event.addListener(this._dm, "circlecomplete", (o) => this._storeOverlay(o, "circle"));
    google.maps.event.addListener(this._dm, "polylinecomplete", (o) => this._storeOverlay(o, "polyline"));
  },

  _storeOverlay(overlay, shape) {
    overlay.__paShape = shape;

    const s = this._styleFor(this._type);
    try {
      if (shape === "polyline") {
        overlay.setOptions({
          strokeColor: s.strokeColor,
          strokeOpacity: s.strokeOpacity,
          strokeWeight: s.strokeWeight,
          zIndex: s.zIndex
        });
      } else {
        overlay.setOptions(s);
      }
    } catch (_) {}

    if (this._type === "pv") {
      if (!overlay.__name) overlay.__name = "Area " + (this._areaCounter++);
      this._attachLabelAboveOverlay(overlay);
      this._pv.push(overlay);
    } else if (this._type === "obstacle") {
      this._obstacles.push(overlay);
    } else {
      this._objects.push(overlay);
    }

    this._dm.setDrawingMode(null);
    document.querySelectorAll("#sub-bar-area .pa-icon-btn").forEach(b => b.classList.remove("active"));
  },

  _attachLabelAboveOverlay(overlay) {
    if (!overlay || overlay.__paShape === "polyline") return;

    const self = this;

    class AreaLabel extends google.maps.OverlayView {
      constructor(ov, text) {
        super();
        this.ov = ov;
        this.text = text;
        this.div = null;
      }
      onAdd() {
        this.div = document.createElement("div");
        this.div.className = "area-label";
        this.div.textContent = this.text;
        this.div.style.position = "absolute";
        this.div.style.transform = "translate(-50%,-100%)";
        this.div.style.pointerEvents = "none";
        this.getPanes().floatPane.appendChild(this.div);
      }
      draw() {
        if (!this.div) return;
        const proj = this.getProjection();
        if (!proj) return;
        const p = self._topPixelOfOverlay(this.ov, proj);
        if (!p) return;
        this.div.style.left = p.x + "px";
        this.div.style.top = (p.y - 12) + "px";
      }
      onRemove() {
        if (this.div) this.div.remove();
        this.div = null;
      }
      setText(t) {
        this.text = t;
        if (this.div) this.div.textContent = t;
      }
    }

    if (overlay.__label) {
      try { overlay.__label.setMap(null); } catch (_) {}
      overlay.__label = null;
    }

    const label = new AreaLabel(overlay, overlay.__name || "");
    label.setMap(this._map);
    overlay.__label = label;

    const redraw = () => { try { label.draw(); } catch (_) {} };

    overlay.__labelListeners = overlay.__labelListeners || [];
    const shape = overlay.__paShape;

    if (shape === "polygon" && overlay.getPath) {
      const path = overlay.getPath();
      overlay.__labelListeners.push(google.maps.event.addListener(path, "set_at", redraw));
      overlay.__labelListeners.push(google.maps.event.addListener(path, "insert_at", redraw));
      overlay.__labelListeners.push(google.maps.event.addListener(path, "remove_at", redraw));
    } else if (shape === "rectangle") {
      overlay.__labelListeners.push(google.maps.event.addListener(overlay, "bounds_changed", redraw));
    } else if (shape === "circle") {
      overlay.__labelListeners.push(google.maps.event.addListener(overlay, "center_changed", redraw));
      overlay.__labelListeners.push(google.maps.event.addListener(overlay, "radius_changed", redraw));
    }

    overlay.__labelListeners.push(google.maps.event.addListener(this._map, "zoom_changed", redraw));
    overlay.__labelListeners.push(google.maps.event.addListener(this._map, "center_changed", redraw));
  },

  _topPixelOfOverlay(overlay, proj) {
    const shape = overlay.__paShape;

    if (shape === "polygon" && overlay.getPath) {
      const arr = overlay.getPath().getArray();
      if (!arr || !arr.length) return null;
      let best = null;
      for (const ll of arr) {
        const px = proj.fromLatLngToDivPixel(ll);
        if (!best || px.y < best.y) best = px;
      }
      return best;
    }

    if (shape === "rectangle" && overlay.getBounds) {
      const b = overlay.getBounds();
      if (!b) return null;
      const ne = b.getNorthEast();
      const sw = b.getSouthWest();
      const nw = new google.maps.LatLng(ne.lat(), sw.lng());
      return proj.fromLatLngToDivPixel(nw);
    }

    if (shape === "circle" && overlay.getCenter && overlay.getRadius) {
      const c = overlay.getCenter();
      const r = overlay.getRadius();
      if (!c || !r) return null;
      const north = google.maps.geometry.spherical.computeOffset(c, r, 0);
      return proj.fromLatLngToDivPixel(north);
    }

    return null;
  },

  _clearByType(type) {
    const arr = type === "pv" ? this._pv : (type === "obstacle" ? this._obstacles : this._objects);

    arr.forEach(o => {
      try {
        if (o.__labelListeners && o.__labelListeners.length) {
          o.__labelListeners.forEach(l => {
            try { google.maps.event.removeListener(l); } catch (_) {}
          });
          o.__labelListeners = [];
        }
        if (o.__label) {
          try { o.__label.setMap(null); } catch (_) {}
          o.__label = null;
        }
        o.setMap(null);
      } catch (_) {}
    });

    arr.length = 0;
    if (type === "pv") this._areaCounter = 1;
  },

  _parseCoords(txt) {
    const lines = txt.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const pts = [];
    for (const line of lines) {
      const parts = line.split(/[,\s]+/).map(s => s.trim()).filter(Boolean);
      if (parts.length < 2) continue;
      const lat = parseFloat(parts[0]);
      const lng = parseFloat(parts[1]);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) continue;
      pts.push({ lat, lng });
    }
    return pts;
  }
};
