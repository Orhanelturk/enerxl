// zoning.js — Zoning logic: balanced substation siting, ITS/inverter pads, trenches & PV rebuild
// Exposes: window.mountZoning({ google, map, document })
window.mountZoning = function mountZoning(ctx){
  const { google, map, document } = ctx;
  const zoneContent = document.getElementById('zone-content');
  if (!zoneContent) return;

  // ---------- UI: inject form ----------
  zoneContent.innerHTML = `
    <div class="zn-section">
      <h4>Topology</h4>
      <div class="zn-note">
        Using <span class="zn-badge">String inverters</span> → pad‑mount transformers (ITS) → substation(s).
      </div>
      <div class="zn-two" style="margin-top:6px">
        <label style="font-size:12px;color:#334;display:flex;align-items:center;gap:8px">
          <input id="zn-align-az" type="checkbox" checked>
          Align pads to PV azimuth
        </label>
        <label style="font-size:12px;color:#334;display:flex;align-items:center;gap:8px">
          <input id="zn-balance" type="checkbox" checked>
          Balance substation loading (±%)
        </label>
      </div>
      <div class="zn-two" style="margin-top:6px">
        <div class="pv-field"><label>Balance tolerance (±%)</label><input id="zn-baltol" type="number" min="0" step="1" value="10"></div>
        <div class="pv-field">
          <label>Routing</label>
          <select id="zn-route">
            <option value="lshape" selected>Azimuth L‑shape (with detours)</option>
          </select>
        </div>
      </div>
    </div>

    <div class="zn-section">
      <h4>Ratings</h4>
      <div class="zn-three">
        <div class="pv-field"><label>DC/AC ratio</label><input id="zn-dcac" type="number" min="1.0" step="0.01" value="1.30"></div>
        <div class="pv-field"><label>Inverter AC (kW)</label><input id="zn-inv-ac" type="number" min="10" step="1" value="110"></div>
        <div class="pv-field"><label>ITS rating (kVA)</label><input id="zn-its-kva" type="number" min="100" step="50" value="2500"></div>
      </div>
      <div class="zn-two" style="margin-top:6px">
        <div class="pv-field"><label>Substation rating (MVA)</label><input id="zn-sub-mva" type="number" min="5" step="5" value="100"></div>
        <div class="pv-field"><label>Phase</label>
          <select id="zn-phase">
            <option value="3" selected>3‑phase</option>
            <option value="1">1‑phase</option>
          </select>
        </div>
      </div>
    </div>

    <div class="zn-section">
      <h4>Pads & Clearances</h4>
      <div class="zn-three">
        <div class="pv-field"><label>INV pad W×L (m)</label>
          <div class="zn-two"><input id="zn-inv-w" type="number" min="0.5" step="0.1" value="2.0"><input id="zn-inv-l" type="number" min="0.5" step="0.1" value="1.2"></div>
        </div>
        <div class="pv-field"><label>ITS pad W×L (m)</label>
          <div class="zn-two"><input id="zn-its-w" type="number" min="1" step="0.1" value="4.5"><input id="zn-its-l" type="number" min="1" step="0.1" value="3.0"></div>
        </div>
        <div class="pv-field"><label>Sub pad W×L (m)</label>
          <div class="zn-two"><input id="zn-sub-w" type="number" min="5" step="0.5" value="60"><input id="zn-sub-l" type="number" min="5" step="0.5" value="40"></div>
        </div>
      </div>
      <div class="zn-three" style="margin-top:6px">
        <div class="pv-field"><label>INV→PV clear (m)</label><input id="zn-gap-inv" type="number" min="0" step="0.1" value="1.5"></div>
        <div class="pv-field"><label>ITS→PV clear (m)</label><input id="zn-gap-its" type="number" min="0" step="0.1" value="3"></div>
        <div class="pv-field"><label>SUB→PV clear (m)</label><input id="zn-gap-sub" type="number" min="0" step="0.1" value="10"></div>
      </div>
      <div class="zn-three" style="margin-top:6px">
        <div class="pv-field"><label>Trench→PV clear (m)</label><input id="zn-gap-trench" type="number" min="0" step="0.1" value="2"></div>
        <div class="pv-field"><label>Block–block min gap (m)</label><input id="zn-gap-block" type="number" min="0" step="0.1" value="4"></div>
        <div class="pv-field"><label>Trench width (m)</label><input id="zn-trench-w" type="number" min="0.2" step="0.1" value="1.0"></div>
      </div>
    </div>

    <div class="zn-section">
      <h4>Preview (from current PV layout)</h4>
      <div class="zn-preview">
        <div class="zn-kv"><label>Total DC (kWp)</label><span id="zn-out-dc">—</span></div>
        <div class="zn-kv"><label>Total AC (kW)</label><span id="zn-out-ac">—</span></div>
        <div class="zn-kv"><label>Est. # Substations</label><span id="zn-out-subs">—</span></div>
        <div class="zn-kv"><label>Est. # ITS (pad‑mount)</label><span id="zn-out-its">—</span></div>
        <div class="zn-kv"><label>Est. # String inverters</label><span id="zn-out-inv">—</span></div>
      </div>
      <div class="zn-note" style="margin-top:6px">
        Counts are estimates; final placement respects pad sizes & clearances.
      </div>
    </div>
  `;

  // ---------- State ----------
  const state = { poi: null, poiMarker: null, placed: { subs:[], its:[], invs:[], trenches:[] } };
  window.__zoningCfg = window.__zoningCfg || {}; // exposed for debugging

  // ---------- Helpers ----------
  function getNumber(id){ const el=document.getElementById(id); return Number(el?.value||0); }
  function getBool(id){ const el=document.getElementById(id); return !!el?.checked; }
  function setText(id, txt){ const el=document.getElementById(id); if(el) el.textContent = txt; }

  function readCfg(){
    const cfg = {
      dcac: getNumber('zn-dcac'),
      inv_ac_kw: getNumber('zn-inv-ac'),
      its_kva: getNumber('zn-its-kva'),
      sub_mva: getNumber('zn-sub-mva'),
      balance: getBool('zn-balance'),
      bal_tol: Math.max(0, getNumber('zn-baltol'))/100, // e.g. 0.10
      align_az: getBool('zn-align-az'),
      route: document.getElementById('zn-route')?.value || 'lshape',
      trench_w: Math.max(0.2, getNumber('zn-trench-w')),
      phase: document.getElementById('zn-phase')?.value || '3',
      pads: {
        inv: { w:getNumber('zn-inv-w'),  l:getNumber('zn-inv-l'),  clr:getNumber('zn-gap-inv') },
        its: { w:getNumber('zn-its-w'),  l:getNumber('zn-its-l'),  clr:getNumber('zn-gap-its') },
        sub: { w:getNumber('zn-sub-w'),  l:getNumber('zn-sub-l'),  clr:getNumber('zn-gap-sub') }
      },
      gaps: {
        trench_to_pv: getNumber('zn-gap-trench'),
        block_block: getNumber('zn-gap-block')
      }
    };
    window.__zoningCfg = cfg;
    return cfg;
  }

  function currentTotalsFromSummary(){
    // Read your on-map Summary values rendered by drawing.js / taskpane.html
    const dcTxt = document.getElementById('tot-dc')?.textContent || '0';
    const acTxt = document.getElementById('tot-ac')?.textContent || '0';
    const dc = Number(dcTxt.replace(/,/g,'')); const ac = Number(acTxt.replace(/,/g,''));
    return { dc_kwp: dc, ac_kw: ac };
  }

  function refreshPreview(){
    const cfg = readCfg();
    let { dc_kwp, ac_kw } = currentTotalsFromSummary();
    // Recompute AC from DC if dcac provided (AC ≈ DC / (DC/AC))
    if (cfg.dcac > 0 && dc_kwp>0) ac_kw = dc_kwp / cfg.dcac;

    const nSubs = cfg.sub_mva>0 ? Math.ceil(ac_kw / (cfg.sub_mva * 1000)) : 0;
    const nITS  = cfg.its_kva>0 ? Math.ceil(ac_kw / cfg.its_kva) : 0; // kW ≈ kVA
    const nInv  = cfg.inv_ac_kw>0 ? Math.ceil(ac_kw / cfg.inv_ac_kw) : 0;

    setText('zn-out-dc',  dc_kwp.toFixed(1));
    setText('zn-out-ac',  ac_kw.toFixed(1));
    setText('zn-out-subs', String(nSubs));
    setText('zn-out-its',  String(nITS));
    setText('zn-out-inv',  String(nInv));
  }

  // ---------- POI picking ----------
  const poiBtn = document.getElementById('zn-poi-pick');
  poiBtn?.addEventListener('click', ()=>{
    if (!google || !map) return;
    map.setOptions({ draggableCursor:'crosshair' });
    google.maps.event.addListenerOnce(map, 'click', (ev)=>{
      if (state.poiMarker) state.poiMarker.setMap(null);
      state.poi = ev.latLng;
      state.poiMarker = new google.maps.Marker({
        map, position: ev.latLng,
        icon: { path: google.maps.SymbolPath.BACKWARD_CLOSED_ARROW, scale: 5, strokeColor:'#2563eb' },
        title:'POI (Grid Tie)'
      });
      window.__poi = ev.latLng;
      map.setOptions({ draggableCursor:null });
    });
  });

  // ---------- Utility from drawing layer (replicated/local) ----------
  function makeLocal(origin){
    return {
      toXY(p){
        const d = google.maps.geometry.spherical.computeDistanceBetween(origin, p);
        const h = google.maps.geometry.spherical.computeHeading(origin, p) * Math.PI/180;
        return { x: d*Math.sin(h), y: d*Math.cos(h) };
      },
      fromXY(x,y){
        const d = Math.hypot(x,y);
        const h = Math.atan2(x,y) * 180/Math.PI;
        return google.maps.geometry.spherical.computeOffset(origin, d, h);
      }
    };
  }
  function contains(poly, latLng){ return google.maps.geometry.poly.containsLocation(latLng, poly); }
  function centroidLL(path){
    let lat=0,lng=0; const n=path.length;
    path.forEach(p=>{ lat+=p.lat(); lng+=p.lng(); });
    return new google.maps.LatLng(lat/n, lng/n);
  }
  function safeContain(poly, latLng, tol=0.02){
    if (contains(poly, latLng)) return true;
    const center = centroidLL(poly.getPath().getArray());
    const heading = google.maps.geometry.spherical.computeHeading(latLng, center);
    const inward = google.maps.geometry.spherical.computeOffset(latLng, tol, heading);
    return contains(poly, inward);
  }
  function getAllOverlays(){ return window.__overlayRegistry ? Array.from(window.__overlayRegistry) : []; } // added by taskpane script
  function areasOnMap(){
    return getAllOverlays().filter(ov=>ov.getMap && ov.getMap() && ov.__cat==='area');
  }
  function modulesOnMap(){
    // modules created by simple fill carry __kwp and __parentArea
    return getAllOverlays().filter(ov=>{
      try{
        return ov.getMap && ov.getMap() && typeof ov.__kwp === 'number' && ov.__kwp>0 && ov.getPath;
      }catch{ return false; }
    });
  }
  function pathOf(ov){
    if (ov.getPath) return ov.getPath().getArray();
    if (ov.getBounds){
      const b=ov.getBounds(), ne=b.getNorthEast(), sw=b.getSouthWest();
      const nw = new google.maps.LatLng(ne.lat(), sw.lng());
      const se = new google.maps.LatLng(sw.lat(), ne.lng());
      return [nw, ne, se, sw, nw];
    }
    return [];
  }
  function polygonFromPath(paths, opts){
    return new google.maps.Polygon(Object.assign({
      map, paths, clickable:true, draggable:false, editable:false,
      strokeColor: '#334155', strokeOpacity:0.9, strokeWeight:1.2, fillColor:'#94a3b8', fillOpacity:0.25
    }, opts||{}));
  }

  // Corridor from path (similar to roads) so trenches are polygons the PV fill will avoid
  function averageHeading(h1,h2){ const r1=h1*Math.PI/180, r2=h2*Math.PI/180; const x=Math.cos(r1)+Math.cos(r2), y=Math.sin(r1)+Math.sin(r2); return Math.atan2(y,x)*180/Math.PI; }
  function headingAtIndex(points, i){
    const n=points.length;
    if(i===0) return google.maps.geometry.spherical.computeHeading(points[0], points[1]);
    if(i===n-1) return google.maps.geometry.spherical.computeHeading(points[n-2], points[n-1]);
    const h1=google.maps.geometry.spherical.computeHeading(points[i-1], points[i]);
    const h2=google.maps.geometry.spherical.computeHeading(points[i], points[i+1]);
    return averageHeading(h1,h2);
  }
  function corridorFromPath(points, widthMeters){
    const half=widthMeters/2, left=[], right=[];
    const n=points.length;
    for(let i=0;i<n;i++){
      const h = headingAtIndex(points, i);
      const leftPt  = google.maps.geometry.spherical.computeOffset(points[i], half, h - 90);
      const rightPt = google.maps.geometry.spherical.computeOffset(points[i], half, h + 90);
      left.push(leftPt); right.push(rightPt);
    }
    right.reverse();
    return left.concat(right);
  }

  // Oriented rectangle around a center
  function orientedRect(centerLL, w, l, azDeg){
    const origin = centerLL;
    const prj = makeLocal(origin);
    const theta = azDeg * Math.PI/180;
    const cosT = Math.cos(theta), sinT = Math.sin(theta);
    const hw=w/2, hl=l/2;
    function rot(u,v){ return { x: cosT*u + sinT*v, y: -sinT*u + cosT*v }; }
    const pts = [ [-hw,-hl], [ hw,-hl], [ hw, hl], [-hw, hl] ].map(([u,v])=>{
      const p = rot(u,v);
      return prj.fromXY(p.x, p.y);
    });
    pts.push(pts[0]);
    return pts;
  }

  function dominantPVAzimuthDeg(){
    // Use module azimuths if available; fall back to Summary az or south=180
    const mods = modulesOnMap();
    if (mods.length){
      const sampleStep = Math.max(1, Math.floor(mods.length/1500));
      const azs = [];
      for (let i=0;i<mods.length;i+=sampleStep){
        const a = Number(mods[i].__az || mods[i].__azimuth || 0);
        if (!isNaN(a)) azs.push(((a%360)+360)%360);
      }
      if (azs.length){
        azs.sort((a,b)=>a-b);
        return azs[Math.floor(azs.length/2)];
      }
    }
    // try summary header
    const cell = document.querySelector('#sum-body tr td.num:last-child');
    if (cell){
      const v = Number(cell.textContent||'180');
      if(!isNaN(v)) return ((v%360)+360)%360;
    }
    return 180;
  }

  // ---------- Balanced k-means (weighted, capacity-awareness, simple) ----------
  function kmeansBalanced(pointsXY, k, capKW, tolFrac, maxIter=10){
    // pointsXY: [{x,y,w,..}], capKW per cluster, tolFrac (e.g., 0.10)
    if (k<=0) return { centers: [], assign: [] };
    // KMeans++ init
    const centers = [];
    const used = new Set();
    function dist2(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return dx*dx+dy*dy; }
    // pick first random
    centers.push(pointsXY[Math.floor(Math.random()*pointsXY.length)]);
    // next centers
    for(let c=1;c<k;c++){
      const d2 = pointsXY.map(p=>{
        let m = Infinity;
        for(const q of centers){ const d = dist2(p,q); if(d<m) m=d; }
        return m;
      });
      const sum = d2.reduce((a,b)=>a+b,0) || 1;
      let r = Math.random()*sum, idx=0;
      for(;idx<d2.length-1;idx++){ r -= d2[idx]; if(r<=0) break; }
      centers.push(pointsXY[idx]);
    }
    // iterative assignment
    let assign = new Array(pointsXY.length).fill(-1);
    for(let it=0; it<maxIter; it++){
      // reset cluster loads
      const caps = new Array(k).fill(0);
      const lists = Array.from({length:k}, ()=>[]);
      // assignment with soft capacity
      for(let i=0;i<pointsXY.length;i++){
        const p = pointsXY[i];
        // sort centers by distance
        const order = centers.map((c,ci)=>({ci,d2:dist2(p,c)})).sort((a,b)=>a.d2-b.d2);
        let chosen = order[0].ci;
        for(const cand of order){
          if (caps[cand.ci] + p.w <= capKW*(1+tolFrac)){
            chosen = cand.ci;
            break;
          }
        }
        assign[i] = chosen;
        caps[chosen] += p.w;
        lists[chosen].push(p);
      }
      // recompute centers (weighted)
      let changed = false;
      for(let ci=0; ci<k; ci++){
        if (!lists[ci].length) continue;
        const sumW = lists[ci].reduce((a,b)=>a+b.w,0);
        const nx = lists[ci].reduce((a,b)=>a+b.x*b.w,0) / (sumW||1);
        const ny = lists[ci].reduce((a,b)=>a+b.y*b.w,0) / (sumW||1);
        const old = centers[ci];
        if (Math.hypot(nx-old.x, ny-old.y) > 0.1) changed = true;
        centers[ci] = { x:nx, y:ny, w:0 };
      }
      if (!changed) break;
    }
    return { centers, assign };
  }

  // ---------- Collect module sample ----------
  function sampleModulePoints(cfg){
    const mods = modulesOnMap();
    if (!mods.length) return null;
    // domain origin
    const allAreas = areasOnMap();
    const origin = allAreas.length ? centroidLL(pathOf(allAreas[0])) : centroidLL(pathOf(mods[0]));
    const prj = makeLocal(origin);

    const step = Math.max(1, Math.floor(mods.length / 2500)); // cap ~2.5k points
    const pts = [];
    for(let i=0;i<mods.length;i+=step){
      const m = mods[i];
      const path = pathOf(m);
      if (!path.length) continue;
      const c = centroidLL(path);
      const xy = prj.toXY(c);
      const wKW = (Number(m.__kwp||0) / (cfg.dcac||1)); // AC kW per module approx
      pts.push({ x:xy.x, y:xy.y, w: Math.max(0.0001, wKW), mref:m, ll:c });
    }
    return { pts, origin, prj };
  }

  // ---------- Place pads (sub/its/inv) & cleanup ----------
  function placePad(centerLL, dims, azDeg, kind, name, color, extra={}){
    const rect = orientedRect(centerLL, dims.w, dims.l, azDeg);
    const poly = polygonFromPath(rect, {
      strokeColor: color || '#111827',
      fillColor: color || '#111827',
      fillOpacity: 0.18,
      zIndex: 1200
    });
    poly.__type='polygon'; poly.__cat='object'; poly.__name=name; poly.__padType = kind;
    // label-like title
    poly.set('title', name);
    return poly;
  }

  function removeModulesUnder(padPoly){
    const mods = modulesOnMap();
    let removed = 0;
    for(const m of mods){
      const mPath = pathOf(m);
      // quick: any module vertex within pad OR pad center within module
      const padCenter = centroidLL(padPoly.getPath().getArray());
      let hit = contains(padPoly, padCenter);
      if (!hit){
        for(const pt of mPath){ if (contains(padPoly, pt)) { hit = true; break; } }
      }
      if (hit){
        try{ m.setMap(null); removed++; }catch(_){}
      }
    }
    return removed;
  }

  function pointInsideAnyArea(latlng){
    const areas = areasOnMap();
    for(const a of areas){
      const poly = new google.maps.Polygon({ paths: pathOf(a) });
      if (contains(poly, latlng)) return true;
    }
    return false;
  }

  function nudgeInside(latlng){
    // if a point is outside, walk toward global centroid until it's inside
    const areas = areasOnMap();
    if (!areas.length) return latlng;
    const allPts = areas.flatMap(a=>pathOf(a));
    const gC = centroidLL(allPts);
    let cur = latlng, tries=0;
    while(!pointInsideAnyArea(cur) && tries<40){
      const d = google.maps.geometry.spherical.computeDistanceBetween(cur, gC);
      const h = google.maps.geometry.spherical.computeHeading(cur, gC);
      cur = google.maps.geometry.spherical.computeOffset(cur, Math.max(1, d*0.25), h);
      tries++;
    }
    return cur;
  }

  // ---------- Trench planning (azimuth L-shape with lateral detours) ----------
  function buildDetouredLPath(aLL, bLL, azDeg, clearMeters){
    // Construct a 2‑segment path aligned with azimuth axes (U axis = azimuth; V axis = azimuth+90)
    const origin = aLL;
    const prj = makeLocal(origin);
    const theta = azDeg * Math.PI/180;
    const cosT = Math.cos(theta), sinT = Math.sin(theta);
    const toUV = (p)=>{ const xy = prj.toXY(p); return { u: cosT*xy.x - sinT*xy.y, v: sinT*xy.x + cosT*xy.y }; };
    const fromUV = (u,v)=>{ const x = cosT*u + sinT*v, y = -sinT*u + cosT*v; return prj.fromXY(x,y); };

    const A = toUV(aLL), B = toUV(bLL);
    // primary L: A (uA,vA) → (uB,vA) → B
    let path = [ aLL, fromUV(B.u, A.v), bLL ];
    // simple detour: try offsets left/right if corridor hits modules
    function pathHitsPV(nodes){
      const corridor = corridorFromPath(nodes, readCfg().trench_w + 2*clearMeters);
      const trenchPoly = new google.maps.Polygon({ paths: corridor });
      const mods = modulesOnMap();
      for(const m of mods){
        const mPath = pathOf(m);
        // check any mod vertex inside corridor
        for(const pt of mPath){ if (safeContain(trenchPoly, pt)) return true; }
      }
      return false;
    }
    if (!pathHitsPV(path)) return path;

    // try lateral offsets (± step) on the elbow
    const step = Math.max(1.0, readCfg().gaps.block_block || 4);
    const maxShift = step*6;
    for(let s=step; s<=maxShift; s+=step){
      // try shifting elbow along +V and -V first, then along +U/-U
      const candidates = [
        { u:B.u, v:A.v + s },
        { u:B.u, v:A.v - s },
        { u:B.u + s, v:A.v },
        { u:B.u - s, v:A.v }
      ];
      for(const c of candidates){
        const elbow = fromUV(c.u, c.v);
        const cand = [ aLL, elbow, bLL ];
        if (!pathHitsPV(cand)) return cand;
      }
    }
    // fallback: straight if unavoidable
    return [aLL, bLL];
  }

  function drawTrenchPath(nodesLL, w, name){
    const corridor = corridorFromPath(nodesLL, w);
    const poly = polygonFromPath(corridor, {
      strokeColor:'#0f766e', fillColor:'#0ea5e9', fillOpacity:0.18, strokeOpacity:0.9, zIndex:1100
    });
    poly.__type='polygon'; poly.__cat='object'; poly.__name=name||'Trench';
    return poly;
  }

  // ---------- Orchestration ----------
  const runBtn = document.getElementById('zn-run');
  runBtn?.addEventListener('click', async ()=>{
    refreshPreview();
    const cfg = readCfg();

    // 0) Basic validations
    const mods = modulesOnMap();
    if (!mods.length){ alert('Please build the PV layout first.'); return; }

    // 1) Collect weighted sample of module centroids
    const sample = sampleModulePoints(cfg);
    if (!sample || !sample.pts.length){ alert('No PV modules found on map.'); return; }
    const { pts, origin, prj } = sample;

    // 2) Totals & counts
    let { dc_kwp, ac_kw } = currentTotalsFromSummary();
    if (cfg.dcac > 0 && dc_kwp>0) ac_kw = dc_kwp / cfg.dcac;

    const subCount = cfg.sub_mva>0 ? Math.max(1, Math.ceil(ac_kw / (cfg.sub_mva*1000))) : 1;
    const subCapKW = cfg.sub_mva*1000;

    // 3) Dominant azimuth for pad rotation
    const domAz = cfg.align_az ? dominantPVAzimuthDeg() : 0;

    // 4) Balanced k-means for substations
    const kres = kmeansBalanced(pts, subCount, subCapKW, cfg.balance ? cfg.bal_tol : 1.0);
    const subCentersXY = kres.centers;
    const subCentersLL = subCentersXY.map(c=>prj.fromXY(c.x, c.y)).map(c=>nudgeInside(c));

    // 5) Create substation pads (remove overlapping PV)
    const placedSubs = [];
    for(let i=0;i<subCentersLL.length;i++){
      const center = subCentersLL[i];
      const pad = placePad(center, cfg.pads.sub, domAz, 'SUB', `Substation ${i+1}`, '#f59e0b');
      placedSubs.push({ pad, center });
      removeModulesUnder(pad);
    }
    state.placed.subs = placedSubs.map(s=>s.pad);

    // 6) Partition points to nearest sub (capacity‑aware assignment we already have in kres.assign)
    const groupMap = new Map();
    pts.forEach((p, idx)=>{
      const ci = kres.assign[idx];
      if (!groupMap.has(ci)) groupMap.set(ci, []);
      groupMap.get(ci).push(p);
    });

    // 7) ITS per sub cluster
    const placedITS = [];
    for(const [ci, groupPts] of groupMap.entries()){
      const clusterKW = groupPts.reduce((a,b)=>a+b.w,0);
      const itsCount = cfg.its_kva>0 ? Math.max(1, Math.ceil(clusterKW / cfg.its_kva)) : 1;
      const itsRes = kmeansBalanced(groupPts, itsCount, cfg.its_kva, 1.0); // soft capacity
      const itsCenters = itsRes.centers.map(c=>prj.fromXY(c.x, c.y)).map(c=>nudgeInside(c));

      for(let j=0;j<itsCenters.length;j++){
        const center = itsCenters[j];
        const pad = placePad(center, cfg.pads.its, domAz, 'ITS', `ITS ${ci+1}-${j+1}`, '#10b981');
        placedITS.push({ pad, center, subIndex: ci });
        removeModulesUnder(pad);
      }
    }
    state.placed.its = placedITS.map(x=>x.pad);

    // 8) Inverters near each ITS: row along azimuth beside ITS pad
    const placedInvs = [];
    for(const its of placedITS){
      // its kw share ~ evenly from its cluster (approx)
      const invCount = cfg.inv_ac_kw>0 ? Math.max(1, Math.ceil(cfg.its_kva / cfg.inv_ac_kw)) : 1;
      const hw = cfg.pads.its.w/2;
      const offset = cfg.pads.its.clr + cfg.pads.inv.l/2 + 0.5;
      const theta = domAz * Math.PI/180;
      const cosT = Math.cos(theta), sinT=Math.sin(theta);

      // place a line of inverters to one side of ITS along azimuth
      for(let n=0;n<invCount;n++){
        const along = (n - (invCount-1)/2) * (cfg.pads.inv.w + cfg.pads.inv.clr);
        // vector perpendicular to azimuth to sit beside ITS
        const sideX = Math.cos((domAz+90)*Math.PI/180)*offset;
        const sideY = Math.sin((domAz+90)*Math.PI/180)*offset;
        const alongX = cosT*along, alongY = sinT*along;
        const itsXY = prj.toXY(its.center);
        const invCenter = prj.fromXY(itsXY.x + sideX + alongX, itsXY.y + sideY + alongY);
        const pad = placePad(invCenter, cfg.pads.inv, domAz, 'INV', `INV ${its.subIndex+1}-${n+1}`, '#2563eb');
        placedInvs.push(pad);
        removeModulesUnder(pad);
      }
    }
    state.placed.invs = placedInvs;

    // 9) Trenches
    const trenchW = cfg.trench_w;
    const trenchClr = cfg.gaps.trench_to_pv || 0;
    const placedTrenches = [];

    // ITS → nearest Sub
    for(const its of placedITS){
      const sub = placedSubs[its.subIndex];
      const nodes = buildDetouredLPath(its.center, sub.center, domAz, trenchClr);
      const tr = drawTrenchPath(nodes, trenchW, `TR ITS→SUB ${its.subIndex+1}`);
      placedTrenches.push(tr);
    }

    // Sub → POI (if given)
    if (state.poi || window.__poi){
      const poi = state.poi || window.__poi;
      for(let i=0;i<placedSubs.length;i++){
        const sub = placedSubs[i];
        const nodes = buildDetouredLPath(sub.center, poi, domAz, trenchClr);
        const tr = drawTrenchPath(nodes, trenchW, `TR SUB${i+1}→POI`);
        placedTrenches.push(tr);
      }
    }
    state.placed.trenches = placedTrenches;

    // 10) Rebuild PV around blockers (pads + trenches)
    //    Mark all areas as selected then trigger your existing "Build initial layout"
    const areas = areasOnMap();
    areas.forEach(a=>{ a.__selected = true; });
    const buildBtn = document.getElementById('pv-build');
    if (buildBtn) buildBtn.click(); // leverages your existing fill; blockers (__cat='object') are respected

    console.log('Zoning done:', {
      subs: state.placed.subs.length,
      its: state.placed.its.length,
      invs: state.placed.invs.length,
      trenches: state.placed.trenches.length
    });
  });

  // Clear zoning overlays
  const clearBtn = document.getElementById('zn-clear');
  clearBtn?.addEventListener('click', ()=>{
    for(const k of ['subs','its','invs','trenches']){
      (state.placed[k]||[]).forEach(ov=>{ try{ ov.setMap(null); }catch{} });
      state.placed[k] = [];
    }
    if (state.poiMarker){ state.poiMarker.setMap(null); state.poiMarker=null; }
    window.__poi = null;
  });

  // Live preview on change
  zoneContent.querySelectorAll('input,select').forEach(inp=>{
    inp.addEventListener('input', refreshPreview);
    inp.addEventListener('change', refreshPreview);
  });

  // Initial fill
  refreshPreview();
};
