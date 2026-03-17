window.UI = {
  _tools: {},
  _tooltip: null,
  _tooltipVisible: false,

  register(toolId, buttonId, panelId) {
    this._tools[toolId] = { buttonId, panelId };
  },

  init() {
    if (this._inited) return;
    this._inited = true;
    // Auto-register tools from DOM (so it works even if Map.js forgot to register)
    this._autoRegisterFromDOM();

    // Capture-phase click so close/toggle works even if other scripts stop propagation
    document.addEventListener("click", (e) => {
      const closeBtn = e.target.closest("[data-close-tool]");
      if (closeBtn) {
        e.preventDefault();
        e.stopPropagation();
        this.close(closeBtn.dataset.closeTool);
        return;
      }

      const btn = e.target.closest("[data-tool]");
      if (!btn) return;

      e.preventDefault();
      e.stopPropagation();

      const toolId = btn.dataset.tool;

      // if not registered (dynamic buttons), register by convention on the fly
      if (!this._tools[toolId]) {
        this._registerByConvention(toolId, btn.id);
      }

      this.toggle(toolId);
    }, true);

    window.addEventListener("keydown", (e) => {
      if (e.key === "Escape") this.closeAll();
    });

    this.enableTooltips();
  },

  _autoRegisterFromDOM() {
    document.querySelectorAll("[data-tool]").forEach(btn => {
      const toolId = btn.dataset.tool;
      if (!toolId) return;

      // need an id to control "active" state
      if (!btn.id) return;

      if (this._tools[toolId]) return;

      this._registerByConvention(toolId, btn.id);
    });
  },

  _registerByConvention(toolId, buttonId) {
    let panelId;

    if (toolId === "search") panelId = "sub-bar-search";
    else if (toolId === "area") panelId = "sub-bar-area";
    else panelId = "panel-" + toolId;

    const panel = document.getElementById(panelId);

    // Only register if panel exists (prevents broken tools)
    if (panel) {
      this.register(toolId, buttonId, panelId);
    }
  },

  close(toolId) {
    const t = this._tools[toolId];
    if (!t) return;

    const btn = document.getElementById(t.buttonId);
    const panel = document.getElementById(t.panelId);

    if (btn) btn.classList.remove("active");
    if (panel) panel.classList.add("hidden");
  },

  closeAll() {
    Object.keys(this._tools).forEach(k => this.close(k));
  },

  toggle(toolId) {
    const t = this._tools[toolId];
    if (!t) return;

    const btn = document.getElementById(t.buttonId);
    const panel = document.getElementById(t.panelId);
    if (!btn || !panel) return;

    const isOpen = btn.classList.contains("active");

    // Close everything first (your requirement)
    this.closeAll();

    // If it was open, now it's closed
    if (isOpen) return;

    // Open selected
    btn.classList.add("active");
    panel.classList.remove("hidden");

    // IMPORTANT: do NOT apply saved size / do NOT add resize handle / do NOT force width
    panel.style.width = "";
    panel.style.height = "";
  },

  enableTooltips() {
    const tip = document.getElementById("tooltip");
    if (!tip) return;
    this._tooltip = tip;

    const show = (el) => {
      const text = el.getAttribute("data-tip");
      if (!text) return;
      tip.textContent = text;
      tip.classList.add("show");
      this._tooltipVisible = true;
      this._positionTooltip(el);
    };

    const hide = () => {
      tip.classList.remove("show");
      tip.style.transform = "translate(-9999px,-9999px)";
      this._tooltipVisible = false;
    };

    document.addEventListener("pointerenter", (e) => {
      const el = e.target.closest && e.target.closest("[data-tip]");
      if (el) show(el);
    }, true);

    document.addEventListener("pointerleave", (e) => {
      const el = e.target.closest && e.target.closest("[data-tip]");
      if (el) hide();
    }, true);

    document.addEventListener("scroll", () => {
      if (!this._tooltipVisible) return;
      const hovered = document.querySelector("[data-tip]:hover");
      if (!hovered) return;
      this._positionTooltip(hovered);
    }, true);

    window.addEventListener("resize", () => {
      if (!this._tooltipVisible) return;
      const hovered = document.querySelector("[data-tip]:hover");
      if (!hovered) return;
      this._positionTooltip(hovered);
    });
  },

  _positionTooltip(el) {
    const tip = this._tooltip;
    if (!tip) return;
    const r = el.getBoundingClientRect();
    const x = r.left + r.width / 2 - tip.offsetWidth / 2;
    const y = r.bottom + 10;
    tip.style.transform = `translate(${x}px, ${y}px)`;
  }
};

(function () {
  const S = (window.SystemInputs = window.SystemInputs || {});
  if (S._wired) return;
  S._wired = true;

  const state = {
    capMode: "all",
    selectingPoc: false,
    nextPocId: 1,
    pocs: [],
    customVoltages: [],
    inverter: null
  };

  function q(id) { return document.getElementById(id); }
  function latOf(x) { return (x && typeof x.lat === "function") ? x.lat() : (x ? x.lat : null); }
  function lngOf(x) { return (x && typeof x.lng === "function") ? x.lng() : (x ? x.lng : null); }

  function setCapMode(mode) {
    state.capMode = mode === "desired" ? "desired" : "all";
    const allBtn = q("si-cap-all");
    const desBtn = q("si-cap-desired");
    if (allBtn) allBtn.classList.toggle("active", state.capMode === "all");
    if (desBtn) desBtn.classList.toggle("active", state.capMode === "desired");

    const inp = q("si-desired-mwp");
    if (inp) {
      inp.disabled = state.capMode !== "desired";
      if (state.capMode !== "desired") inp.value = "";
    }
  }

  function setPocSelecting(on) {
    state.selectingPoc = !!on;
    const btn = q("si-select-poc");
    const st = q("si-poc-status");
    if (btn) btn.classList.toggle("active", state.selectingPoc);

    if (st) {
      st.textContent = state.selectingPoc
        ? ("Click on the map to place POC " + state.nextPocId)
        : "";
    }

    try {
      window.dispatchEvent(new CustomEvent("enerx:select-poc", {
        detail: { active: state.selectingPoc, nextId: state.nextPocId }
      }));
    } catch (_) {}
  }

  function tryCreatePocMarker(latLng, id) {
    try {
      if (!window.google || !google.maps) return null;
      const map = window.map || window._map || null;
      if (!map) return null;

      const pos = (latLng && typeof latLng.lat === "function")
        ? latLng
        : new google.maps.LatLng(latOf(latLng), lngOf(latLng));

      const icon = {
        path: "M12 24C6.8 19.3 4 15.6 4 10.5C4 5.9 8 2 12 2C16 2 20 5.9 20 10.5C20 15.6 17.2 19.3 12 24Z",
        fillColor: "#7c3aed",
        fillOpacity: 0.9,
        strokeColor: "#111",
        strokeWeight: 1,
        scale: 1
      };

      const marker = new google.maps.Marker({
        map,
        position: pos,
        icon,
        label: {
          text: "POC " + id,
          color: "#111",
          fontWeight: "600",
          fontSize: "12px"
        },
        draggable: true
      });

      return marker;
    } catch (_) {
      return null;
    }
  }

  function renderPocTable() {
    const body = document.querySelector("#si-poc-table tbody");
    if (!body) return;

    body.innerHTML = "";
    state.pocs.forEach((p) => {
      const tr = document.createElement("tr");

      const tdId = document.createElement("td");
      tdId.textContent = "POC " + p.id;

      const tdLat = document.createElement("td");
      tdLat.textContent = (typeof p.lat === "number") ? p.lat.toFixed(6) : "";

      const tdLng = document.createElement("td");
      tdLng.textContent = (typeof p.lng === "number") ? p.lng.toFixed(6) : "";

      const tdV = document.createElement("td");
      const vInp = document.createElement("input");
      vInp.className = "si-cell-input";
      vInp.type = "number";
      vInp.step = "0.1";
      vInp.placeholder = "kV";
      vInp.value = p.voltageKv || "";
      vInp.addEventListener("input", () => {
        p.voltageKv = vInp.value;
        refreshVoltageDerived();
      });
      tdV.appendChild(vInp);

      const tdAct = document.createElement("td");
      const rm = document.createElement("button");
      rm.className = "si-mini-btn";
      rm.type = "button";
      rm.textContent = "Remove";
      rm.addEventListener("click", () => {
        if (p.marker && p.marker.setMap) p.marker.setMap(null);
        state.pocs = state.pocs.filter(x => x !== p);
        renderPocTable();
        refreshVoltageDerived();
      });
      tdAct.appendChild(rm);

      tr.appendChild(tdId);
      tr.appendChild(tdLat);
      tr.appendChild(tdLng);
      tr.appendChild(tdV);
      tr.appendChild(tdAct);

      body.appendChild(tr);
    });
  }

  function getLvFromInverter(inv) {
    if (!inv) return null;
    const v = inv.acVoltageNom ?? inv.VOutConv ?? inv.vacNom ?? inv.VacNom ?? null;
    const n = (typeof v === "string") ? parseFloat(v) : v;
    return isFinite(n) ? n : null;
  }

  function buildVoltageTable() {
    const body = document.querySelector("#si-voltage-table tbody");
    if (!body) return;
    body.innerHTML = "";

    const rows = [
      { key: "LV", locked: true },
      { key: "POC", locked: true },
      ...state.customVoltages
    ];

    rows.forEach((r) => {
      const tr = document.createElement("tr");

      const tdName = document.createElement("td");
      const nameInp = document.createElement("input");
      nameInp.className = "si-cell-input";
      nameInp.type = "text";
      nameInp.value = r.key || "";
      nameInp.disabled = !!r.locked;
      nameInp.addEventListener("input", () => { r.key = nameInp.value; });
      tdName.appendChild(nameInp);

      const tdV = document.createElement("td");
      const vInp = document.createElement("input");
      vInp.className = "si-cell-input";
      vInp.type = "number";
      vInp.step = "0.1";
      vInp.placeholder = r.locked ? "auto" : "V/kV";
      vInp.disabled = !!r.locked;
      vInp.value = (r.value ?? "") === null ? "" : (r.value ?? "");
      vInp.addEventListener("input", () => { r.value = vInp.value; });
      tdV.appendChild(vInp);

      const tdAct = document.createElement("td");
      if (!r.locked) {
        const rm = document.createElement("button");
        rm.className = "si-mini-btn";
        rm.type = "button";
        rm.textContent = "Remove";
        rm.addEventListener("click", () => {
          state.customVoltages = state.customVoltages.filter(x => x !== r);
          buildVoltageTable();
        });
        tdAct.appendChild(rm);
      } else {
        tdAct.textContent = "";
      }

      tr.appendChild(tdName);
      tr.appendChild(tdV);
      tr.appendChild(tdAct);
      body.appendChild(tr);

      if (r.locked && r.key === "LV") r._inp = vInp;
      if (r.locked && r.key === "POC") r._inp = vInp;
    });

    refreshVoltageDerived();
  }

  function refreshVoltageDerived() {
    const body = document.querySelector("#si-voltage-table tbody");
    if (!body) return;

    const lv = getLvFromInverter(state.inverter);
    const lvCell = body.querySelector('tr:nth-child(1) input[type="number"]');
    if (lvCell) lvCell.value = (lv == null) ? "" : lv;

    const pocCell = body.querySelector('tr:nth-child(2) input[type="number"]');
    if (pocCell) {
      const vals = state.pocs.map(p => parseFloat(p.voltageKv)).filter(n => isFinite(n));
      if (vals.length === 1 && state.pocs.length === 1) pocCell.value = vals[0];
      else pocCell.value = "";
    }
  }

  function fillInverterManufacturers() {
    const mSel = q("si-inv-mfr");
    const modelSel = q("si-inv-model");
    if (!mSel || !modelSel) return;

    mSel.innerHTML = '<option value="">--</option>';
    modelSel.innerHTML = '<option value="">--</option>';
    modelSel.disabled = true;

    const lib = window.inverterLibrary;
    if (!Array.isArray(lib) || !lib.length) return;

    const makers = [...new Set(lib.map(x => x.manufacturer || x.Manufacturer).filter(Boolean))].sort();
    makers.forEach(m => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      mSel.appendChild(opt);
    });

    mSel.addEventListener("change", () => {
      const chosen = mSel.value;
      modelSel.innerHTML = '<option value="">--</option>';
      modelSel.disabled = !chosen;

      const subset = lib.filter(x => (x.manufacturer || x.Manufacturer) === chosen);
      subset.sort((a,b) => (a.acPower || 0) - (b.acPower || 0));

      subset.forEach(inv => {
        const label = inv.model || inv.name || "Model";
        const kw = inv.acPower || "";
        const opt = document.createElement("option");
        opt.value = inv.name || label;
        opt.textContent = (kw ? ("(" + kw + " kW) ") : "") + label;
        opt._invRef = inv;
        modelSel.appendChild(opt);
      });
    });

    modelSel.addEventListener("change", () => {
      const chosenName = modelSel.value;
      const inv = lib.find(x => (x.name === chosenName) || (x.model === chosenName));
      state.inverter = inv || null;

      const hint = q("si-inv-lv-hint");
      const lv = getLvFromInverter(state.inverter);
      if (hint) hint.textContent = lv ? ("LV detected: " + lv + " V") : "";

      refreshVoltageDerived();
    });
  }

  function fillModuleManufacturers() {
    const mSel = q("si-mod-mfr");
    const modelSel = q("si-mod-model");
    if (!mSel || !modelSel) return;

    mSel.innerHTML = '<option value="">--</option>';
    modelSel.innerHTML = '<option value="">--</option>';
    modelSel.disabled = true;

    const lib = window.moduleLibrary;
    if (!Array.isArray(lib) || !lib.length) return;

    const makers = [...new Set(lib.map(x => x.manufacturer || x.Manufacturer).filter(Boolean))].sort();
    makers.forEach(m => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      mSel.appendChild(opt);
    });

    mSel.addEventListener("change", () => {
      const chosen = mSel.value;
      modelSel.innerHTML = '<option value="">--</option>';
      modelSel.disabled = !chosen;

      const subset = lib.filter(x => (x.manufacturer || x.Manufacturer) === chosen);
      subset.sort((a,b) => (a.PNom || a.pnom || 0) - (b.PNom || b.pnom || 0));

      subset.forEach(mod => {
        const label = mod.model || mod.Model || mod.name || "Model";
        const wp = mod.PNom || mod.pnom || mod.watt || "";
        const opt = document.createElement("option");
        opt.value = label;
        opt.textContent = (wp ? ("(" + wp + " W) ") : "") + label;
        modelSel.appendChild(opt);
      });
    });
  }

  S.refreshLibraries = function () {
    fillModuleManufacturers();
    fillInverterManufacturers();
  };

  S.onMapPOCSelected = function (latLng) {
    if (!state.selectingPoc) return false;

    const lat = latOf(latLng);
    const lng = lngOf(latLng);
    if (!isFinite(lat) || !isFinite(lng)) return false;

    const id = state.nextPocId++;
    const marker = tryCreatePocMarker(latLng, id);

    state.pocs.push({
      id,
      lat,
      lng,
      voltageKv: "",
      marker
    });

    setPocSelecting(false);
    renderPocTable();
    refreshVoltageDerived();
    return true;
  };

  S.getState = function () {
    return JSON.parse(JSON.stringify({
      capMode: state.capMode,
      desiredMWp: q("si-desired-mwp")?.value || "",
      pocs: state.pocs.map(p => ({ id: p.id, lat: p.lat, lng: p.lng, voltageKv: p.voltageKv || "" })),
      inverterType: (document.querySelector('input[name="si-inverter-type"]:checked')?.value) || "string",
      dcCombiner: !!q("si-dc-combiner")?.checked,
      mvStation: !!q("si-mv-station")?.checked,
      voltageLevels: [
        { level: "LV", voltage: (getLvFromInverter(state.inverter) ?? "") },
        { level: "POC", voltage: (state.pocs.length === 1 ? (state.pocs[0].voltageKv || "") : "") },
        ...state.customVoltages.map(v => ({ level: v.key || "", voltage: v.value || "" }))
      ]
    }));
  };

  S.init = function () {
    const panel = q("panel-system-inputs");
    if (!panel || S._inited) return;
    S._inited = true;

    q("si-cap-all")?.addEventListener("click", () => setCapMode("all"));
    q("si-cap-desired")?.addEventListener("click", () => setCapMode("desired"));
    setCapMode("all");

    q("si-select-poc")?.addEventListener("click", () => setPocSelecting(!state.selectingPoc));
    q("si-clear-pocs")?.addEventListener("click", () => {
      state.pocs.forEach(p => { if (p.marker && p.marker.setMap) p.marker.setMap(null); });
      state.pocs = [];
      state.nextPocId = 1;
      setPocSelecting(false);
      renderPocTable();
      refreshVoltageDerived();
    });

    q("si-add-voltage")?.addEventListener("click", () => {
      state.customVoltages.push({ key: "Custom", value: "" });
      buildVoltageTable();
    });

    q("si-clear-voltage")?.addEventListener("click", () => {
      state.customVoltages = [];
      buildVoltageTable();
    });

    buildVoltageTable();
    S.refreshLibraries();

    window.addEventListener("enerx:module-library-updated", () => S.refreshLibraries());
    window.addEventListener("enerx:inverter-library-updated", () => S.refreshLibraries());
  };
})();

(function () {
  if (!window.UI || window.UI._systemInputsPatched) return;
  window.UI._systemInputsPatched = true;
  const oldInit = window.UI.init.bind(window.UI);
  window.UI.init = function () {
    oldInit();
    if (window.SystemInputs && typeof window.SystemInputs.init === "function") {
      window.SystemInputs.init();
    }
  };
})();

(function () {
  const run = () => {
    if (window.UI && typeof window.UI.init === "function") window.UI.init();
  };
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", run, { once: true });
  } else {
    run();
  }
})();
