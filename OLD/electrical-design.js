
  const SQRT3 = Math.sqrt(3);
  const SVG_TEXT_SIZE = 5;
  const SVG_LINE_HEIGHT = 5;

  function formatNumber(v, d) {
    if (isNaN(v) || v === null) return "-";
    return Number(v).toFixed(d);
  }

  function showError(msg) {
    const box = document.getElementById("errorBox");
    box.textContent = msg;
    box.style.display = msg ? "block" : "none";
  }

  function showDesignWarning(msg) {
  const box = document.getElementById("warningBox");
  if (!box) return;
  box.textContent = msg || "";
  box.style.display = msg ? "block" : "none";
}


  function roundUpToSeries(value, series) {
    if (!isFinite(value) || value <= 0) return series[0];
    for (let i = 0; i < series.length; i++) {
      if (value <= series[i]) return series[i];
    }
    return series[series.length - 1];
  }

  function initUserEditedTracking() {
    const ids = [
      "dcSize","mvBusType", "mvBusCurrent", "mvBusCoupler"
    ];
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("change", () => {
        el.dataset.userEdited = "true";
      });
    });
  }

  function suggestAndApplyDefaults(cfg) {
    const {
      acExportMw,
      mvVoltage
    } = cfg || {};

    const mvRequiredMva = acExportMw || 0;
    const mvRequiredCurrent = (mvVoltage > 0)
      ? (mvRequiredMva * 1e6) / (SQRT3 * mvVoltage * 1e3)
      : NaN;

    const mvBusTypeEl = document.getElementById("mvBusType");
    const mvBusCurrentEl = document.getElementById("mvBusCurrent");
    const mvBusCouplerEl = document.getElementById("mvBusCoupler");

    if (mvBusCurrentEl && !mvBusCurrentEl.dataset.userEdited &&
        (!mvBusCurrentEl.value || parseFloat(mvBusCurrentEl.value) <= 0)) {
      const mvCurrentSeries = [630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000];
      const suggested = roundUpToSeries(mvRequiredCurrent, mvCurrentSeries) || 2000;
      mvBusCurrentEl.value = String(Math.round(suggested));
    }

    if (mvBusTypeEl && !mvBusTypeEl.dataset.userEdited) {
      let type = mvBusTypeEl.value || "single";
      if (acExportMw > 150) type = "b&h";
      else if (acExportMw > 40) type = "double";
      else type = "single";
      mvBusTypeEl.value = type;
    }

    if (mvBusCouplerEl && !mvBusCouplerEl.dataset.userEdited) {
      const type = mvBusTypeEl ? mvBusTypeEl.value : "single";
      mvBusCouplerEl.checked = (type === "double" || type === "b&h");
    }
  }

  function resetForm() {
    document.getElementById("dcSize").value = "100";
    document.getElementById("acExport").value = "";
    document.getElementById("dcAcRatio").value = "1.2";
    document.getElementById("gridVoltage").value = "132";
    document.getElementById("mvVoltage").value = "33";
    document.getElementById("lvVoltage").value = "0.8";
    document.getElementById("invRating").value = "300";
    document.getElementById("invPerMvTxf").value = "";
    document.getElementById("mvTxfMva").value = "9.0";
    document.getElementById("mvTxfLoading").value = "85";
    document.getElementById("gridTxfMva").value = "125";
    document.getElementById("gridTxfLoading").value = "90";
    document.getElementById("gridTxfCount").value = "";
    document.getElementById("gridN1").checked = false;
    document.getElementById("mvBusType").value = "single";
    document.getElementById("mvBusCurrent").value = "2000";
    document.getElementById("mvBusCoupler").checked = false;
    document.getElementById("hvBusCoupler").checked = false;

    const allInputs = document.querySelectorAll("input, select");
    allInputs.forEach(el => { delete el.dataset.userEdited; });

    updateLoopMaxOptions();
    updateGridTransformerFiltering();

    showError("");
    showDesignWarning("");
    const results = document.getElementById("resultsContainer");
    results.innerHTML =
      '<div class="summary-mini"><strong>Waiting for inputs...</strong><br>' +
      'Enter data on the left and click <em>Calculate SLD Concept</em>.</div>';
  }

  function createSvgElement(type, attrs) {
    const el = document.createElementNS("http://www.w3.org/2000/svg", type);
    Object.keys(attrs || {}).forEach(k => el.setAttribute(k, attrs[k]));
    return el;
  }

  const sldPanZoomState = {
    svg: null,
    root: null,
    scale: 1,
    offsetX: 0,
    offsetY: 0,
    isPanning: false,
    startX: 0,
    startY: 0
  };

  function applyPanZoomTransform() {
    const st = sldPanZoomState;
    if (!st.root) return;
    st.root.setAttribute(
      "transform",
      "translate(" + st.offsetX + "," + st.offsetY + ") scale(" + st.scale + ")"
    );
  }

  function setupPanZoom(svg, root) {
    const st = sldPanZoomState;
    st.svg = svg;
    st.root = root;
    st.scale = 1;
    st.offsetX = 0;
    st.offsetY = 0;
    applyPanZoomTransform();

    if (svg.dataset.panzoomAttached === "true") return;
    svg.dataset.panzoomAttached = "true";

    svg.addEventListener("wheel", function(e) {
      e.preventDefault();

      const pt = svg.createSVGPoint();
      pt.x = e.clientX;
      pt.y = e.clientY;
      const cursor = pt.matrixTransform(svg.getScreenCTM().inverse());

      const factor = e.deltaY < 0 ? 1.25 : 0.8;
      const newScale = Math.min(20, Math.max(0.1, st.scale * factor));

      if (newScale === st.scale) return;

      st.offsetX = cursor.x - (cursor.x - st.offsetX) * (newScale / st.scale);
      st.offsetY = cursor.y - (cursor.y - st.offsetY) * (newScale / st.scale);
      st.scale = newScale;
      applyPanZoomTransform();
    }, { passive: false });

    svg.addEventListener("mousedown", function(e) {
      if (e.button !== 0) return;
      st.isPanning = true;
      st.startX = e.clientX;
      st.startY = e.clientY;
    });

    svg.addEventListener("mousemove", function(e) {
      if (!st.isPanning) return;

      const speed = st.scale < 1 ? 1 : st.scale;
      const dx = (e.clientX - st.startX) * speed;
      const dy = (e.clientY - st.startY) * speed;

      st.startX = e.clientX;
      st.startY = e.clientY;

      st.offsetX += dx;
      st.offsetY += dy;

      applyPanZoomTransform();
    });

    svg.addEventListener("mouseup", function() {
      st.isPanning = false;
    });
    svg.addEventListener("mouseleave", function() {
      st.isPanning = false;
    });
  }

  function drawSLD(params) {
    const {
      hasSeparateGridTxf,
      gridVoltage,
      mvVoltage,
      hvBusType,
      mvBusType,
      gridTxfMva,
      mvTxfMva,
      numGridTransformers,
      mvLoops,
      gtLoopMap,
      hvBusCoupler,
      mvBusCoupler,
      mvSections,
      mvBusCurrent,
      lvBlocks    
    } = params;

    const outFeeders = parseInt(document.getElementById("gridOutFeeders").value);

    const svg = document.getElementById("sldCanvas");
    if (!svg) return;
    while (svg.firstChild) svg.removeChild(svg.firstChild);

    const width = Math.max(2000, mvLoops.length * 250);
    const height = 420;

    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.setAttribute("preserveAspectRatio", "xMidYMid meet");

    const baseStroke = "#0f172a";
    const baseFill = "#ffffff";
    const fontFamily = "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
    const centerX = width / 2;

    const root = createSvgElement("g", { id: "sldRoot" });
    svg.appendChild(root);

    function addText(x, y, text, anchor = "start") {
      const t = createSvgElement("text", {
        x, y,
        "font-size": SVG_TEXT_SIZE,
        "font-family": fontFamily,
        "fill": baseStroke,
        "text-anchor": anchor,
        "dominant-baseline": "middle"
      });
      t.textContent = text;
      root.appendChild(t);
    }

    function addMultilineText(lines, x, y, anchor = "start") {
      const t = createSvgElement("text", {
        x, y,
        "font-size": SVG_TEXT_SIZE,
        "font-family": fontFamily,
        "fill": baseStroke,
        "text-anchor": anchor
      });
      lines.forEach((line, idx) => {
        const span = createSvgElement("tspan", {
          x,
          dy: idx === 0 ? 0 : SVG_LINE_HEIGHT
        });
        span.textContent = line;
        t.appendChild(span);
      });
      root.appendChild(t);
    }

    function drawBusbar(x1, x2, y, label) {
  root.appendChild(createSvgElement("line", {
    x1, y1: y, x2, y2: y,
    stroke: baseStroke,
    "stroke-width": 4,
    "stroke-linecap": "round"
  }));

  // TEXT ON THE LEFT SIDE
  addText(x1, y - 14, label, "start");
}



    function drawBreaker(x, y) {
      const s = 3;
      root.appendChild(createSvgElement("rect", {
        x: x - s / 2, y: y - s / 2,
        width: s, height: s,
        fill: baseFill,
        stroke: baseStroke,
        "stroke-width": 0.8
      }));
    }

    function drawGridTransformer(x, y, index, hvKv, mvKv, ratingMva) {
      const r = 6;
      root.appendChild(createSvgElement("circle", {
        cx: x, cy: y - 5, r,
        stroke: baseStroke, "stroke-width": 1.2, fill: "none"
      }));
      root.appendChild(createSvgElement("circle", {
        cx: x, cy: y + 5, r,
        stroke: baseStroke, "stroke-width": 1.2, fill: "none"
      }));

      const labelLines = [
        "GT-" + index,
        `${formatNumber(hvKv, 0)}/${formatNumber(mvKv, 0)}kV`,
        `${formatNumber(ratingMva, 0)} MVA`
      ];
      addMultilineText(labelLines, x + 12, y - SVG_LINE_HEIGHT, "start");
    }

    function drawMvTransformer(x, y, id, mvKv, lvKv, mva, mw) {
      const r = 5;
      root.appendChild(createSvgElement("circle", {
        cx: x - 4, cy: y, r,
        stroke: baseStroke, "stroke-width": 1, fill: "none"
      }));
      root.appendChild(createSvgElement("circle", {
        cx: x + 4, cy: y, r,
        stroke: baseStroke, "stroke-width": 1, fill: "none"
      }));

      const lines = [
        "TX-" + id,
        `${formatNumber(mvKv, 0)}/${formatNumber(lvKv, 1)}kV`,
        `${formatNumber(mva, 1)} MVA`,
        `${formatNumber(mw, 1)} (MW)`
      ];
      const oldSize = SVG_TEXT_SIZE;
const smallSize = 4;       // smaller text
const smallLine = 4;       // closer spacing

addMultilineTextCustomFont(lines, x, y - 20, "middle", smallSize, smallLine);


    }
function addMultilineTextCustomFont(lines, x, y, anchor, size, lineHeight) {
  const t = createSvgElement("text", {
    x, y,
    "font-size": size,
    "font-family": "system-ui",
    "fill": "#0f172a",
    "text-anchor": anchor
  });

  lines.forEach((line, idx) => {
    const span = createSvgElement("tspan", {
      x,
      dy: idx === 0 ? 0 : lineHeight
    });
    span.textContent = line;
    t.appendChild(span);
  });

  root.appendChild(t);
}

    function drawVerticalLoop(loop, colX, mvBusY, mvVoltage, lvVoltage, mvTxfMva) {
      const cbY = mvBusY + 20;
      const trunkTopY = cbY + 14;

      root.appendChild(createSvgElement("line", {
        x1: colX, y1: mvBusY,
        x2: colX, y2: cbY - 6,
        stroke: baseStroke, "stroke-width": 1
      }));
      drawBreaker(colX, cbY);

      root.appendChild(createSvgElement("line", {
        x1: colX, y1: cbY + 6,
        x2: colX, y2: trunkTopY,
        stroke: baseStroke, "stroke-width": 1
      }));

      const txCount = loop.transformers.length;
      if (txCount === 0) return;

      const perTxStep = 36;
      const trunkBottomY = trunkTopY + 16 + (txCount - 1) * perTxStep;

      root.appendChild(createSvgElement("line", {
        x1: colX, y1: trunkTopY,
        x2: colX, y2: trunkBottomY,
        stroke: baseStroke, "stroke-width": 1.1
      }));

      const branchLen = 12;

      loop.transformers.forEach((tx, idx) => {
        const ty = trunkTopY + 16 + idx * perTxStep;
        const cbX = colX + branchLen / 2;
        const txX = colX + branchLen;

        root.appendChild(createSvgElement("line", {
          x1: colX, y1: ty,
          x2: cbX, y2: ty,
          stroke: baseStroke, "stroke-width": 1
        }));
        drawBreaker(cbX, ty);
        root.appendChild(createSvgElement("line", {
          x1: cbX, y1: ty,
          x2: txX, y2: ty,
          stroke: baseStroke, "stroke-width": 1
        }));

        drawMvTransformer(txX + 6, ty, tx.id, mvVoltage, lvVoltage, mvTxfMva, tx.mw);
      });

      const labelY = trunkBottomY + 20;
      addText(colX, labelY, `Loop ${loop.id} (${formatNumber(loop.totalMw, 1)} MW)`, "middle");
    }

    // -------------------------
    // MV-only plant (no separate HV transformer)
    // -------------------------
    if (!hasSeparateGridTxf || !numGridTransformers) {
      const mvBusY = 80;
      const mvBusX1 = centerX - 260;
      const mvBusX2 = centerX + 260;

      drawBusbar(mvBusX1, mvBusX2, mvBusY,
        `MV POI BUS (${formatNumber(mvVoltage, 0)} kV, ${mvBusType})`);

      const loops = mvLoops || [];
      const loopCount = loops.length;

      if (loopCount > 0) {
        const baseStep = 80;
        const span = baseStep * Math.max(loopCount - 1, 0);
        const startX = loopCount === 1 ? centerX : centerX - span / 2;
        const gap = loopCount > 1 ? span / (loopCount - 1) : 0;

        loops.forEach((loop, i) => {
          const colX = loopCount === 1 ? centerX : startX + i * gap;
          drawVerticalLoop(loop, colX, mvBusY, mvVoltage, 0.4, mvTxfMva);
        });
      }

      

      setupPanZoom(svg, root);
      return;
    }

    // --- Compute busbar ratings & actual loading ---
    const mvBusRatedMVA = (SQRT3 * mvVoltage * mvBusCurrent) / 1000;
    const hvRatedMVA = gridTxfMva * numGridTransformers;
    let hvActualMw = 0;
    mvLoops.forEach(loop => hvActualMw += loop.totalMw);

    // -------------------------
    // HV + MV plant
    // -------------------------
    const gtCount = Math.max(numGridTransformers, 1);
    const gtY = 90;
    const mvBusY = 170;

    const sections = Math.max(1, mvSections || 1);
    const gtToSection = new Array(gtCount).fill(0);
    for (let i = 0; i < gtCount; i++) {
      gtToSection[i] = i % sections;
    }

    const loopsPerSection = new Array(sections).fill(0);
    const actualMvLoad = new Array(sections).fill(0);

    for (let i = 0; i < gtCount; i++) {
      const loops = gtLoopMap[i] || [];
      loopsPerSection[gtToSection[i]] += loops.length;
      actualMvLoad[gtToSection[i]] += loops.reduce((sum, loop) => sum + loop.totalMw, 0);
    }

    let maxLoopsInSection = Math.max(...loopsPerSection);
    if (!isFinite(maxLoopsInSection) || maxLoopsInSection < 1) maxLoopsInSection = 1;

    const loopSpacing = 90;
    const sectionMargin = 160;
    const sectionWidth = maxLoopsInSection > 1
      ? (maxLoopsInSection - 1) * loopSpacing + sectionMargin
      : 420;

    const couplerGap = 30;
    const totalSpan = sections * sectionWidth + (sections - 1) * couplerGap;
    const startX = centerX - totalSpan / 2;

    // --- HV bus ---
    const hvBusY = 40;
    const hvBusX1 = startX;
    const hvBusX2 = startX + totalSpan;

    const hvLabel =
      `HV BUS (${formatNumber(gridVoltage,0)} kV)
Rated: ${formatNumber(hvRatedMVA,1)} MVA
Actual: ${formatNumber(hvActualMw,1)} MW`;

    drawBusbar(hvBusX1, hvBusX2, hvBusY, hvLabel);

    // --- Draw outgoing feeders to grid (north/up) ---
// --- Draw outgoing feeders to grid (north / centered) ---
if (outFeeders && outFeeders > 0) {

  const feederSpacing = 80; // horizontal spacing between feeders

  // center of HV bus
  const busCenterX = (hvBusX1 + hvBusX2) / 2;

  // total span occupied by feeders
  const totalWidth = (outFeeders - 1) * feederSpacing;

  // start position (centered)
  const startX = busCenterX - totalWidth / 2;

  for (let f = 0; f < outFeeders; f++) {
    const fx = startX + f * feederSpacing;

    // vertical line from HV bus up to breaker
    root.appendChild(createSvgElement("line", {
      x1: fx, y1: hvBusY,
      x2: fx, y2: hvBusY - 20,
      stroke: baseStroke, "stroke-width": 1
    }));

    // breaker
    drawBreaker(fx, hvBusY - 25);

    // vertical outgoing line
    root.appendChild(createSvgElement("line", {
      x1: fx, y1: hvBusY - 30,
      x2: fx, y2: hvBusY - 60,
      stroke: baseStroke, "stroke-width": 1
    }));

    // label above feeder
    addText(fx, hvBusY - 70, `Grid Out-${f+1}`, "middle");
  }
}



    // --- HV Bus Coupler (large square) ---
if (hvBusCoupler) {
  const bcX = (hvBusX1 + hvBusX2) / 2;

  // left bus segment
  root.appendChild(createSvgElement("line", {
    x1: bcX - 12, y1: hvBusY,
    x2: bcX - 4,  y2: hvBusY,
    stroke: baseStroke, "stroke-width": 2
  }));

  // large HV coupler breaker
  const s = 10; // ← BIG SIZE HERE
  root.appendChild(createSvgElement("rect", {
    x: bcX - s/2,
    y: hvBusY - s/2,
    width: s,
    height: s,
    fill: baseFill,
    stroke: baseStroke,
    "stroke-width": 1.2
  }));

  // right bus segment
  root.appendChild(createSvgElement("line", {
    x1: bcX + 4, y1: hvBusY,
    x2: bcX + 12, y2: hvBusY,
    stroke: baseStroke, "stroke-width": 2
  }));
}




    const sectionBusInfo = [];
    for (let s = 0; s < sections; s++) {
      const x1 = startX + s * (sectionWidth + couplerGap);
      const x2 = x1 + sectionWidth;
      const cx = (x1 + x2) / 2;

      const mvLabel =
        `MV BUS ${String.fromCharCode(65 + s)} (${formatNumber(mvVoltage,0)} kV)
Rated: ${formatNumber(mvBusRatedMVA,1)} MVA
Actual: ${formatNumber(actualMvLoad[s],1)} MW`;

      drawBusbar(x1, x2, mvBusY, mvLabel);
      sectionBusInfo.push({ x1, x2, cx });

      if (mvBusCoupler && s < sections - 1) {
        const nextX1 = startX + (s + 1) * (sectionWidth + couplerGap);
        const bcX = (x2 + nextX1) / 2;
        root.appendChild(createSvgElement("line", {
          x1: x2, y1: mvBusY,
          x2: bcX - 3, y2: mvBusY,
          stroke: baseStroke, "stroke-width": 1
        }));
        drawBreaker(bcX, mvBusY);
        root.appendChild(createSvgElement("line", {
          x1: bcX + 3, y1: mvBusY,
          x2: nextX1, y2: mvBusY,
          stroke: baseStroke, "stroke-width": 1
        }));
      }
    }

    // Draw GTs above their sections
    for (let i = 0; i < gtCount; i++) {
      const sectionIndex = gtToSection[i];
      const busInfo = sectionBusInfo[sectionIndex];
      const gtX = busInfo.cx;
      const hvCbY = (hvBusY + gtY) / 2;

      root.appendChild(createSvgElement("line", {
        x1: gtX, y1: hvBusY,
        x2: gtX, y2: hvCbY - 6,
        stroke: baseStroke, "stroke-width": 1
      }));
      drawBreaker(gtX, hvCbY);
      root.appendChild(createSvgElement("line", {
        x1: gtX, y1: hvCbY + 6,
        x2: gtX, y2: gtY - 12,
        stroke: baseStroke, "stroke-width": 1
      }));

      drawGridTransformer(gtX, gtY, i + 1, gridVoltage, mvVoltage, gridTxfMva);

      // MV-side breaker BELOW GT
// First small line segment from GT downwards
root.appendChild(createSvgElement("line", {
  x1: gtX, y1: gtY + 12,
  x2: gtX, y2: gtY + 28,
  stroke: baseStroke, "stroke-width": 1
}));

// MV-side breaker
drawBreaker(gtX, gtY + 32);

// Line from breaker to MV bus
root.appendChild(createSvgElement("line", {
  x1: gtX, y1: gtY + 36,
  x2: gtX, y2: mvBusY,
  stroke: baseStroke, "stroke-width": 1
}));

    }

    // Draw loops
    const leftMargin = 40;
    const map = gtLoopMap || [];

    for (let i = 0; i < gtCount; i++) {
      const colLoops = map[i] || [];
      if (colLoops.length === 0) continue;

      const sectionIndex = gtToSection[i];
      const busInfo = sectionBusInfo[sectionIndex];

      const startXLoop = busInfo.x1 + leftMargin;

      colLoops.forEach((loop, idx) => {
        const colX = startXLoop + idx * 90;
        drawVerticalLoop(loop, colX, mvBusY, mvVoltage, 0.4, mvTxfMva);
      });
    }

    setupPanZoom(svg, root);
  }

  function calculateSLD() {

    showError("");
// ======================================
// DC STRINGING AUTO-CALC ENGINE
// ======================================
function autoComputeDcStringing(invAcKw, dcAcRatio, moduleWp, modulesPerStringUser, stringsPerInvUser) {

  const targetRatio = dcAcRatio;
  const targetDcW = invAcKw * 1000 * dcAcRatio;

  // CASE 1 — user provided both → trust user
  if (modulesPerStringUser > 0 && stringsPerInvUser > 0) {
    return {
      modulesPerString: modulesPerStringUser,
      stringsPerInv: stringsPerInvUser
    };
  }

  // CASE 2 — user gave modules per string → compute strings
  if (modulesPerStringUser > 0 && (!stringsPerInvUser || stringsPerInvUser <= 0)) {
    const ms = modulesPerStringUser;
    const stringW = ms * moduleWp;
    const ideal = targetDcW / stringW;

    const c1 = Math.floor(ideal);
    const c2 = Math.ceil(ideal);

    function ratio(c) { return (c * stringW) / (invAcKw * 1000); }

    const best = Math.abs(ratio(c1) - targetRatio) < Math.abs(ratio(c2) - targetRatio) ? c1 : c2;

    return {
      modulesPerString: ms,
      stringsPerInv: best
    };
  }

  // CASE 3 — user gave strings per inverter → compute modules per string
  if (stringsPerInvUser > 0 && (!modulesPerStringUser || modulesPerStringUser <= 0)) {
    const spi = stringsPerInvUser;
    const idealModules = targetDcW / (spi * moduleWp);

    let m1 = Math.max(12, Math.min(36, Math.floor(idealModules)));
    let m2 = Math.max(12, Math.min(36, Math.ceil(idealModules)));

    function ratio(ms) { return (spi * ms * moduleWp) / (invAcKw * 1000); }

    const bestMs = Math.abs(ratio(m1) - targetRatio) < Math.abs(ratio(m2) - targetRatio) ? m1 : m2;

    return {
      modulesPerString: bestMs,
      stringsPerInv: spi
    };
  }

  // CASE 4 — neither provided → choose best from typical sets
  const typical = [22, 24, 26, 28, 30];
  let best = null;
  let bestErr = 999;

  for (let ms of typical) {
    const stringW = ms * moduleWp;
    const ideal = targetDcW / stringW;

    const c1 = Math.floor(ideal);
    const c2 = Math.ceil(ideal);

    function ratio(c) { return (c * stringW) / (invAcKw * 1000); }

    const r1 = ratio(c1);
    const r2 = ratio(c2);

    const cand1Err = Math.abs(r1 - targetRatio);
    const cand2Err = Math.abs(r2 - targetRatio);

    if (cand1Err < bestErr) { bestErr = cand1Err; best = {ms, spi: c1}; }
    if (cand2Err < bestErr) { bestErr = cand2Err; best = {ms, spi: c2}; }
  }

  return {
    modulesPerString: best.ms,
    stringsPerInv: best.spi
  };
}

    const dcSize = parseFloat(document.getElementById("dcSize").value);
    const acExportUser = parseFloat(document.getElementById("acExport").value);
    const dcAcRatio = parseFloat(document.getElementById("dcAcRatio").value);
    const gridVoltage = parseFloat(document.getElementById("gridVoltage").value);
    const mvVoltage = parseFloat(document.getElementById("mvVoltage").value);
    const lvVoltageInput = parseFloat(document.getElementById("lvVoltage").value);
    const lvVoltage = (!isNaN(lvVoltageInput) && lvVoltageInput > 0) ? lvVoltageInput : 0.4;

    const loopMaxMwInput = parseFloat(document.getElementById("loopMaxMw").value);
    let loopMaxMw = (!isNaN(loopMaxMwInput) && loopMaxMwInput > 0) ? loopMaxMwInput : 25;

    const invRatingKw = parseFloat(document.getElementById("invRating").value);
    const invPerMvTxfUser = parseFloat(document.getElementById("invPerMvTxf").value);
    const mvTxfMva = parseFloat(document.getElementById("mvTxfMva").value);
    const mvTxfLoading = parseFloat(document.getElementById("mvTxfLoading").value);
    const gridTxfMva = parseFloat(document.getElementById("gridTxfMva").value);
    const gridTxfLoading = parseFloat(document.getElementById("gridTxfLoading").value);
    const gridTxfCountUser = parseFloat(document.getElementById("gridTxfCount").value);
    const requireN1 = document.getElementById("gridN1").checked;

    const mvBusType = document.getElementById("mvBusType").value;
    const hvBusType = document.getElementById("hvBusType").value;
    const hvBusCoupler = document.getElementById("hvBusCoupler").checked;
    const mvBusCoupler = document.getElementById("mvBusCoupler").checked;
    const mvBusCurrent = parseFloat(document.getElementById("mvBusCurrent").value);

    if (isNaN(dcSize) || dcSize <= 0) return showError("Please enter a valid Plant DC size (MWp).");
    if (isNaN(dcAcRatio) || dcAcRatio <= 0) return showError("Please enter a valid DC/AC ratio.");
    if (isNaN(invRatingKw) || invRatingKw <= 0) return showError("Please enter a valid Inverter AC rating (kW).");
    if (isNaN(mvTxfMva) || mvTxfMva <= 0) return showError("Please enter a valid LV/MV transformer rating (MVA).");
    if (isNaN(gridVoltage) || gridVoltage <= 0) return showError("Please enter a valid Grid/POI voltage (kV).");
    if (isNaN(mvVoltage) || mvVoltage <= 0) return showError("Please enter a valid MV collection level (kV).");

    let acExportMw = acExportUser;
    if (isNaN(acExportMw) || acExportMw <= 0) acExportMw = dcSize / dcAcRatio;

    const acExportKw = acExportMw * 1000;
    let numInverters = Math.ceil(acExportKw / invRatingKw);
    if (!isFinite(numInverters) || numInverters < 1) numInverters = 1;

    const inverterAcMw = invRatingKw / 1000;
    const totalInverterAcMw = numInverters * inverterAcMw;

    let mvTxfInvertersPerUnit;
    if (!isNaN(invPerMvTxfUser) && invPerMvTxfUser > 0) {
      mvTxfInvertersPerUnit = Math.max(1, Math.floor(invPerMvTxfUser));
    } else {
      const mvTxfEffectiveMVA = mvTxfMva * (mvTxfLoading / 100);
      mvTxfInvertersPerUnit = Math.max(1, Math.floor(mvTxfEffectiveMVA / inverterAcMw));
    }
    const numMvTransformers = Math.ceil(numInverters / mvTxfInvertersPerUnit);

    const hasSeparateGridTxf = gridVoltage > mvVoltage && !isNaN(gridTxfMva) && gridTxfMva > 0;

    let numGridTransformers = 0;
    let gridTxfEffectiveMW = 0;

    if (hasSeparateGridTxf) {
      const gridTxfEffectiveMVA = gridTxfMva * (gridTxfLoading / 100);
      gridTxfEffectiveMW = gridTxfEffectiveMVA;

      if (!isNaN(gridTxfCountUser) && gridTxfCountUser > 0) {
        numGridTransformers = Math.floor(gridTxfCountUser);
      } else {
        numGridTransformers = Math.ceil(acExportMw / gridTxfEffectiveMW);
      }
      if (numGridTransformers < 1) numGridTransformers = 1;

      if (requireN1) {
        while ((numGridTransformers - 1) * gridTxfEffectiveMW < acExportMw) {
          numGridTransformers++;
          if (numGridTransformers > 20) break;
        }
      }
    }

    suggestAndApplyDefaults({
      acExportMw,
      mvVoltage
    });

    const invertersPerMvTxf = [];



// --- AUTO DC STRINGING / LV CALC ---
const moduleWp = parseFloat(document.getElementById("moduleWp").value);
const modulesPerStringUser = parseInt(document.getElementById("modulesPerString").value);
const stringsPerInverterUser = parseInt(document.getElementById("stringsPerInverter")?.value);

let lvBlocks = [];

for (let i = 0; i < numMvTransformers; i++) {

  const invCount = 0; // <-- will fill AFTER splitting inverters, don't worry yet

  // AC rating per inverter
  const invAcKw = invRatingKw;

  // auto compute DC values
  const auto = autoComputeDcStringing(
    invAcKw,
    dcAcRatio,
    moduleWp,
    modulesPerStringUser,
    stringsPerInverterUser
  );

  const ms = auto.modulesPerString;
  const spi = auto.stringsPerInv;

  lvBlocks.push({
    id: i + 1,
    modulesPerString: ms,
    stringsPerInverter: spi,
    moduleWp,
    invCount: 0, // will fill later
    stringsTotal: 0,
    modulesTotal: 0,
    dcMwBlock: 0
  });
}




    const basePerTx = Math.floor(numInverters / numMvTransformers);
    let remainderInv = numInverters % numMvTransformers;

    for (let i = 0; i < numMvTransformers; i++) {
      invertersPerMvTxf.push(basePerTx + (i < remainderInv ? 1 : 0));
    }

    // Fill DC/LV block counts now that we know invPerMvTxf
for (let i = 0; i < numMvTransformers; i++) {

  const invCount = invertersPerMvTxf[i];
  const ms = lvBlocks[i].modulesPerString;
  const spi = lvBlocks[i].stringsPerInverter;
  const mwBlock = (invCount * spi * ms * moduleWp) / 1e6;

  lvBlocks[i].invCount = invCount;
  lvBlocks[i].stringsTotal = invCount * spi;
  lvBlocks[i].modulesTotal = invCount * spi * ms;
  lvBlocks[i].dcMwBlock = mwBlock;
}

// ---------- LV PANEL DISTRIBUTION ----------
const lvPanelRatingKw = parseFloat(document.getElementById("lvPanelRating").value);

for (let i = 0; i < numMvTransformers; i++) {
  
  const invCount = lvBlocks[i].invCount;   // <-- use already computed invCount
  const invKw = invRatingKw;
  const totalKw = invCount * invKw;

  const panelCapacityKw = lvPanelRatingKw;
  const numPanels = Math.ceil(totalKw / panelCapacityKw);

  const panels = [];
  let remainingKw = totalKw;

  for (let p = 0; p < numPanels; p++) {
    const panelKw = Math.min(panelCapacityKw, remainingKw);
    const invInPanel = Math.floor(panelKw / invKw);
    remainingKw -= panelKw;

    panels.push({
      id: p + 1,
      inverterCount: invInPanel,
      inverterKw: invKw,
      panelKw: panelKw
    });
  }

  lvBlocks[i].lvPanels = panels;
  lvBlocks[i].numLvPanels = numPanels;
}

function getInvertersForPanel(invCount, modulesPerString, stringsPerInverter, moduleWp) {
  const invArray = [];
  for (let i = 1; i <= invCount; i++) {
    const dcW = modulesPerString * stringsPerInverter * moduleWp;
    invArray.push({
      id: i,
      modulesPerString,
      stringsPerInverter,
      moduleWp,
      dcKw: dcW / 1000
    });
  }
  return invArray;
}



    const mvLoops = [];
    let currentLoop = { id: 1, transformers: [], totalMw: 0 };

    for (let i = 0; i < numMvTransformers; i++) {
      const txMw = invertersPerMvTxf[i] * inverterAcMw;
      if (currentLoop.totalMw > 0 && currentLoop.totalMw + txMw > loopMaxMw) {
        mvLoops.push(currentLoop);
        currentLoop = { id: mvLoops.length + 1, transformers: [], totalMw: 0 };
      }
      currentLoop.transformers.push({ id: i + 1, mw: txMw });
      currentLoop.totalMw += txMw;
    }
    if (currentLoop.transformers.length > 0) mvLoops.push(currentLoop);

    let gtLoopMap = [];
    if (hasSeparateGridTxf && numGridTransformers > 0) {
      gtLoopMap = new Array(numGridTransformers).fill(null).map(() => []);
      let gi = 0;
      mvLoops.forEach(loop => {
        gtLoopMap[gi].push(loop);
        gi = (gi + 1) % numGridTransformers;
      });
    }

    const totalMvMva = numMvTransformers * mvTxfMva;
    const mvBusIscKA = mvBusCurrent;
    const mvBusFaultMva = (SQRT3 * mvVoltage * mvBusIscKA) || 0;
    const mvMaxConnectedMva = 0.6 * mvBusFaultMva;
    let mvSections = Math.max(1, Math.ceil(totalMvMva / mvMaxConnectedMva));

    if (hasSeparateGridTxf && numGridTransformers > mvSections) {
      mvSections = numGridTransformers;
    }

    const results = document.getElementById("resultsContainer");
    let html = "";

    html += `
      <div class="summary-mini">
        <strong>${formatNumber(dcSize,2)} MWp</strong> |
        <strong>${formatNumber(acExportMw,2)} MW AC</strong> |
        Grid <strong>${formatNumber(gridVoltage,0)} kV</strong> → MV <strong>${formatNumber(mvVoltage,0)} kV</strong> |
        Inverters: <strong>${numInverters}</strong> × ${formatNumber(inverterAcMw,3)} MW |
        MV TX: <strong>${numMvTransformers}</strong> × ${formatNumber(mvTxfMva,1)} MVA |
        ${hasSeparateGridTxf ? `GT: <strong>${numGridTransformers}</strong> × ${formatNumber(gridTxfMva,1)} MVA |` : "" }
        Loops: <strong>${mvLoops.length}</strong> |
        MV Sections: <strong>${mvSections}</strong>
      </div>
    `;

    html += `
  <div class="summary-mini">
    <strong>DC Stringing:</strong><br>
    Modules/String: <strong>${lvBlocks[0].modulesPerString}</strong> |
    Strings/Inverter: <strong>${lvBlocks[0].stringsPerInverter}</strong><br>
    Example Block: <strong>${formatNumber(lvBlocks[0].dcMwBlock,2)} MWp</strong> |
    ${lvBlocks[0].modulesTotal} modules |
    ${lvBlocks[0].stringsTotal} strings
  </div>
`;
html += `
<div class="summary-mini collapsible-container">

  <div class="collapsible-header" onclick="
    const b = this.nextElementSibling;
    b.style.display = (b.style.display === 'none') ? 'block' : 'none';
    this.querySelector('.arrow').textContent = (b.style.display === 'none') ? '▶' : '▼';
  ">
    <span><strong>LV Panel Distribution</strong></span>
    <span class="arrow">▶</span>
  </div>

  <div class="collapsible-body" style="display:none; max-height:250px; overflow-y:auto;">
    ${lvBlocks.map(block => `
      <div class="lv-panel-block">
        <strong>MV TX-${block.id}</strong> → ${block.numLvPanels} LV Panels<br>

        ${block.lvPanels.map((p,idx) => `
          <div class="lv-panel-line">
            <div class="lv-panel-header" onclick="
              const b = this.nextElementSibling;
              b.style.display = (b.style.display === 'none') ? 'block' : 'none';
              this.querySelector('.arrow').textContent = (b.style.display === 'none') ? '▶' : '▼';
            ">
              <span>Panel ${p.id}: ${p.inverterCount} × ${p.inverterKw} kW = ${p.panelKw} kW</span>
              <span class='arrow'>▶</span>
            </div>

            <div class="lv-panel-details" style="display:none; padding-left:14px; margin-bottom:6px;">
              ${getInvertersForPanel(
                  p.inverterCount,
                  block.modulesPerString,
                  block.stringsPerInverter,
                  block.moduleWp
              ).map(inv => `
                <div class="inv-line">
  Inverter ${inv.id}: 
  ${inv.modulesPerString} modules × ${inv.stringsPerInverter} strings 
  (${inv.dcKw.toFixed(2)} kW DC)
</div>

              `).join("")}
            </div>
          </div>
        `).join("")}
      </div>
    `).join("")}
  </div>

</div>
`;



    html += `
      </div></div></div> 
      <div class="sld-full">
        <div class="sld-wrapper-wide">
          <svg id="sldCanvas"></svg>
        </div>
      </div>
    `;

    results.innerHTML = html;

    drawSLD({
      hasSeparateGridTxf,
      gridVoltage,
      mvVoltage,
      hvBusType,
      mvBusType,
      gridTxfMva,
      mvTxfMva,
      numGridTransformers,
      mvLoops,
      gtLoopMap,
      hvBusCoupler,
      mvBusCoupler,
      mvSections,
      mvBusCurrent,
      lvBlocks    
    });
  }

  function initCollapsibles() {
    document.querySelectorAll(".wizard-toggle").forEach(toggle => {
      const targetId = toggle.getAttribute("data-target");
      const body = document.getElementById(targetId);
      let collapsed = toggle.getAttribute("data-collapsed") === "true";
      if (collapsed && body) body.classList.add("hidden");
      toggle.addEventListener("click", () => {
        collapsed = !collapsed;
        toggle.setAttribute("data-collapsed", collapsed ? "true" : "false");
        if (!body) return;
        if (collapsed) body.classList.add("hidden");
        else body.classList.remove("hidden");
      });
    });
  }

  function updateLoopMaxOptions() {
    const mv = parseFloat(document.getElementById("mvVoltage").value) || 0;
    const select = document.getElementById("loopMaxMw");
    const prev = parseFloat(select.value) || null;

    let options;
    if (mv > 0 && mv < 15) {
      // 11 kV family
      options = [3, 5, 7, 8, 10];
    } else if (mv >= 15 && mv < 30) {
      // 22 kV family
      options = [7, 10, 12, 15];
    } else {
      // 33–34.5 kV family
      options = [20, 25, 30, 35];
    }

    select.innerHTML = "";
    options.forEach(v => {
      const opt = document.createElement("option");
      opt.value = String(v);
      opt.textContent = v + " MW";
      select.appendChild(opt);
    });

    let selectedValue = options[0];
    if (prev && options.includes(prev)) selectedValue = prev;
    select.value = String(selectedValue);
  }

  function updateGridTransformerFiltering() {
    const gridSel = document.getElementById("gridVoltage");
    const mvSel = document.getElementById("mvVoltage");
    const tfSel = document.getElementById("gridTxfMva");

    const hv = parseFloat(gridSel.value) || 0;
    const mv = parseFloat(mvSel.value) || 0;

    // If POI <= MV: grid transformer not used → disable entire select
    if (hv <= mv) {
      tfSel.disabled = true;
      tfSel.style.opacity = "0.5";
      return;
    } else {
      tfSel.disabled = false;
      tfSel.style.opacity = "1";
    }

    // HV connection: enable/disable individual options based on realistic ranges
    const opts = Array.from(tfSel.options);
    let firstValid = null;

    opts.forEach(opt => {
      const mva = parseFloat(opt.value);
      let valid = true;

      if (hv > 0 && hv <= 132) {
        // 66–132 kV: typically 20–200 MVA
        valid = (mva >= 20 && mva <= 200);
      } else if (hv > 132 && hv <= 220) {
        // 150–220 kV: 50–315 MVA
        valid = (mva >= 50 && mva <= 315);
      } else if (hv > 220) {
        // 300–400–500 kV: 250–500 MVA
        valid = (mva >= 250 && mva <= 500);
      }

      opt.disabled = !valid;
      opt.style.color = valid ? "" : "#999";

      if (valid && firstValid === null) firstValid = opt.value;
    });

    if (tfSel.options[tfSel.selectedIndex]?.disabled && firstValid !== null) {
      tfSel.value = firstValid;
    }
  }


  function applyGlobalFiltering() {
  const gridSel = document.getElementById("gridVoltage");
  const mvSel   = document.getElementById("mvVoltage");
  const lvSel   = document.getElementById("lvVoltage");

  const mvBusTypeEl    = document.getElementById("mvBusType");
  const mvBusCurrentEl = document.getElementById("mvBusCurrent");
  const mvBusCouplerEl = document.getElementById("mvBusCoupler");
  const hvBusCouplerEl = document.getElementById("hvBusCoupler");

  const gridTxfSel     = document.getElementById("gridTxfMva");
  const gridTxfLoadEl  = document.getElementById("gridTxfLoading");
  const gridTxfCountEl = document.getElementById("gridTxfCount");
  const gridN1El       = document.getElementById("gridN1");
  const gridOutFeedSel = document.getElementById("gridOutFeeders");

  const lvPanelSel     = document.getElementById("lvPanelRating");

  const dcSize   = parseFloat(document.getElementById("dcSize").value) || 0;
  const acExportUser = parseFloat(document.getElementById("acExport").value);
  const dcAcRatio = parseFloat(document.getElementById("dcAcRatio").value) || 1.0;

  let acExportMw = acExportUser;
  if (isNaN(acExportMw) || acExportMw <= 0) {
    acExportMw = dcSize > 0 ? (dcSize / dcAcRatio) : 0;
  }
  const plantKw = acExportMw * 1000;

  const poi = parseFloat(gridSel.value) || 0;
  const mv  = parseFloat(mvSel.value)   || 0;

  const isLvPoi = poi > 0 && poi <= 1.0;     // 0.4–0.8 kV
  const isMvPoi = poi > 1.0 && poi <= 36.0;  // 3.3–34.5 kV
  const isHvPoi = poi > 36.0;                // 66–500 kV

  // Helper to enable/disable a control
  function setEnabled(el, enabled) {
    if (!el) return;
    el.disabled = !enabled;
    el.style.opacity = enabled ? "1" : "0.5";
  }

  // -----------------------
  // 1) LV POI: disable ALL MV & HV
  // -----------------------
  if (isLvPoi) {
    // Disable MV selection and MV equipment
    setEnabled(mvSel, false);
    setEnabled(mvBusTypeEl, false);
    setEnabled(mvBusCurrentEl, false);
    setEnabled(mvBusCouplerEl, false);

    // Disable grid transformer & HV stuff
    setEnabled(gridTxfSel, false);
    setEnabled(gridTxfLoadEl, false);
    setEnabled(gridTxfCountEl, false);
    setEnabled(gridN1El, false);
    setEnabled(hvBusCouplerEl, false);
    setEnabled(gridOutFeedSel, false);

    // No special filtering for LV panel ratings here (just size-based below)
  } 
  // -----------------------
  // 2) MV POI: disable HV, keep MV
  // -----------------------
  else if (isMvPoi) {
    // MV stays enabled
    setEnabled(mvSel, true);
    setEnabled(mvBusTypeEl, true);
    setEnabled(mvBusCurrentEl, true);
    setEnabled(mvBusCouplerEl, true);

    // No grid transformer (POI already at MV)
    setEnabled(gridTxfSel, false);
    setEnabled(gridTxfLoadEl, false);
    setEnabled(gridTxfCountEl, false);
    setEnabled(gridN1El, false);

    // HV bus coupler & feeders conceptually not used
    setEnabled(hvBusCouplerEl, false);
    setEnabled(gridOutFeedSel, false);
  }
  // -----------------------
  // 3) HV POI: everything relevant ON
  // -----------------------
  else if (isHvPoi) {
    setEnabled(mvSel, true);
    setEnabled(mvBusTypeEl, true);
    setEnabled(mvBusCurrentEl, true);
    setEnabled(mvBusCouplerEl, true);

    // Grid transformer & HV features enabled
    setEnabled(gridTxfSel, true);
    setEnabled(gridTxfLoadEl, true);
    setEnabled(gridTxfCountEl, true);
    setEnabled(gridN1El, true);
    setEnabled(hvBusCouplerEl, true);
    setEnabled(gridOutFeedSel, true);
  }

  // -----------------------
  // 4) Filter MV voltage options based on POI
  // -----------------------
  if (mvSel) {
    const mvOptions = Array.from(mvSel.options);
    mvOptions.forEach(opt => {
      const mvVal = parseFloat(opt.value) || 0;
      let valid = true;

      if (isLvPoi) {
        // no MV system if POI is LV
        valid = false;
      } else if (isMvPoi) {
        // MV collection must not exceed POI level
        if (mvVal > poi) valid = false;
      } else if (isHvPoi) {
        // any MV is allowed under HV
        valid = true;
      }

      opt.disabled = !valid;
      opt.style.color = valid ? "" : "#999";
    });

    // If current MV is invalid, pick first valid
    if (mvSel.options[mvSel.selectedIndex]?.disabled) {
      const firstValid = mvOptions.find(o => !o.disabled);
      if (firstValid) mvSel.value = firstValid.value;
    }
  }

  // -----------------------
  // 5) Filter LV panel ratings: disable huge panels on tiny plants
  // -----------------------
  if (lvPanelSel) {
    const opts = Array.from(lvPanelSel.options);
    let firstValidPanel = null;
    opts.forEach(opt => {
      const panelKw = parseFloat(opt.value) || 0;
      // allow panels up to 2× plant capacity, disable extreme ones:
      const valid = (plantKw <= 0) ? true : (panelKw <= plantKw * 2);

      opt.disabled = !valid;
      opt.style.color = valid ? "" : "#999";

      if (valid && firstValidPanel === null) {
        firstValidPanel = opt.value;
      }
    });

    if (lvPanelSel.options[lvPanelSel.selectedIndex]?.disabled && firstValidPanel !== null) {
      lvPanelSel.value = firstValidPanel;
    }
  }

  // -----------------------
  // 6) Design warning messages (non-blocking)
  // -----------------------
  let warn = "";

  if (isMvPoi && acExportMw > 0 && acExportMw < 1) {
    warn = `Notice: Plant capacity ${formatNumber(acExportMw,2)} MW is very small for a MV grid connection (${poi.toFixed(1)} kV). ` +
           `Check if a LV (0.4–0.8 kV) POI might be more appropriate.`;
  }

  if (isHvPoi && acExportMw > 0 && acExportMw < 20) {
    warn = `Notice: Plant capacity ${formatNumber(acExportMw,2)} MW is quite small for a HV grid connection (${poi.toFixed(0)} kV). ` +
           `Typically HV POI is used for ≥ 20–50 MW plants.`;
  }

  showDesignWarning(warn);
}

  function autoFillAfterStep1IfPossible() {
    const dcSize = parseFloat(document.getElementById("dcSize").value);
    const acExportUser = parseFloat(document.getElementById("acExport").value);
    const dcAcRatio = parseFloat(document.getElementById("dcAcRatio").value);
    const mvVoltage = parseFloat(document.getElementById("mvVoltage").value);

    if (isNaN(dcSize) || dcSize <= 0) return;
    if (isNaN(dcAcRatio) || dcAcRatio <= 0) return;

    let acExportMw = acExportUser;
    if (isNaN(acExportMw) || acExportMw <= 0) acExportMw = dcSize / dcAcRatio;

    suggestAndApplyDefaults({
      acExportMw,
      mvVoltage
    });
  }


  // Auto-fill Plant DC size in Electrical Design from PV Layout total DC
function initTotalDcSyncFromPvLayout() {
  const dcInput = document.getElementById("dcSize");
  const totCell = document.getElementById("tot-dc"); // from PV layout summary

  // If either side is missing, just bail out gracefully
  if (!dcInput || !totCell) return;

  // Treat typing as user override – once they type, we stop auto-updating
  dcInput.addEventListener("input", () => {
    dcInput.dataset.userEdited = "true";
  });

  function applyTotalDcFromLayout() {
    // Don't overwrite if user has manually edited Plant DC
    if (dcInput.dataset.userEdited === "true") return;

    const raw = (totCell.textContent || totCell.innerText || "").trim();
    if (!raw) return;

    // Strip non-numeric stuff (commas, units, etc.)
    const numericKwp = parseFloat(
      raw
        .replace(/[^0-9.,\-]/g, "") // keep digits, dot, comma, minus
        .replace(/,/g, "")          // drop thousand separators
    );

    if (!isFinite(numericKwp) || numericKwp <= 0) return;

    // PV layout total is in kWp → convert to MWp for Electrical Design
    const numericMwp = numericKwp / 1000;

    dcInput.value = numericMwp.toString();

    // Re-apply downstream defaults based on DC size
    autoFillAfterStep1IfPossible();
    applyGlobalFiltering();
  }

  // 1) Initial fill – in case user already has a PV layout before opening Electrical Design
  applyTotalDcFromLayout();

  // 2) React to future changes in PV layout total DC
  const observer = new MutationObserver(applyTotalDcFromLayout);
  observer.observe(totCell, {
    childList: true,
    characterData: true,
    subtree: true
  });

  // Optional: expose a manual trigger if you want to call it when you show the Electrical Design window
  window.refreshElectricalDesignDcFromLayout = applyTotalDcFromLayout;
}


  document.getElementById("btnCalculate").addEventListener("click", calculateSLD);
document.getElementById("btnReset").addEventListener("click", () => {
  resetForm();
  applyGlobalFiltering();
});

document.getElementById("dcSize").addEventListener("change", () => {
  autoFillAfterStep1IfPossible();
  applyGlobalFiltering();
});

document.getElementById("acExport").addEventListener("change", applyGlobalFiltering);
document.getElementById("dcAcRatio").addEventListener("change", applyGlobalFiltering);
document.getElementById("invRating").addEventListener("change", applyGlobalFiltering);
document.getElementById("mvTxfMva").addEventListener("change", applyGlobalFiltering);
document.getElementById("mvTxfLoading").addEventListener("change", applyGlobalFiltering);

document.getElementById("mvVoltage").addEventListener("change", () => {
  updateLoopMaxOptions();
  updateGridTransformerFiltering();
  applyGlobalFiltering();
});

document.getElementById("gridVoltage").addEventListener("change", () => {
  updateGridTransformerFiltering();
  applyGlobalFiltering();
});

initCollapsibles();
initUserEditedTracking();
updateLoopMaxOptions();
updateGridTransformerFiltering();
applyGlobalFiltering();
initTotalDcSyncFromPvLayout();   // NEW: bridge PV Layout → Electrical Design


// Wire the Electrical Design / SLD sheet to the left rail button
(function wireElectricalDesignSheet() {
  const edSheet  = document.getElementById("ed-sheet");
  const edButton = document.querySelector('.tool-btn[data-tool="electrical"]');
  const edClose  = document.getElementById("ed-close");

  if (!edSheet || !edButton || !edClose) {
    console.warn("Electrical Design UI elements not found in taskpane.");
    return;
  }

  function openSheet() {
    // show SLD sheet
    edSheet.classList.add("open");
    edSheet.setAttribute("aria-modal", "true");

    // optionally hide other sheets so they don't overlap
    const pvSheet   = document.getElementById("pv-sheet");
    const zoneSheet = document.getElementById("zone-sheet");
    if (pvSheet)   pvSheet.classList.remove("open");
    if (zoneSheet) zoneSheet.classList.remove("open");
  }

  function closeSheet() {
    edSheet.classList.remove("open");
    edSheet.setAttribute("aria-modal", "false");
  }

  // Left-rail “Electrical Design” button
  edButton.addEventListener("click", (e) => {
    e.preventDefault();
    openSheet();
  });

  // “X” close button on the SLD window
  edClose.addEventListener("click", (e) => {
    e.preventDefault();
    closeSheet();
  });
})();


