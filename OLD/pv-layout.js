// PV layout logic extracted into a separate file.
// This script mounts all PV-related behavior once the DOM is ready.
(function () {
  function ready(fn){ if (document.readyState !== 'loading') fn(); else document.addEventListener('DOMContentLoaded', fn); }

  ready(function () {
    // ===== Basic DOM refs (PV sheet & controls) =====
    const drawer = document.getElementById('drawer');
    const pvBtn = document.querySelector('[data-tool="pv"]');
    const toolsBtn = document.querySelector('[data-tool="tools"]');
    const searchBtn = document.querySelector('[data-tool="search"]');
    const searchWrap = document.getElementById('search-wrap');

    const pvSheet = document.getElementById('pv-sheet');
    const pvClose = document.getElementById('pv-close');
    const panBrowse = document.getElementById('pv-pan-browse');
    const panInput = document.getElementById('pv-pan-input');
    const panName = document.getElementById('pv-pan-name');

    const catBackdrop = document.getElementById('pv-cat-backdrop');
    const catClose = document.getElementById('pv-cat-close');
    const openCatalog = document.getElementById('open-catalog');
    const catGrid = document.getElementById('pv-catalog');
    const structPicked = document.getElementById('struct-picked');

    const pvLib = document.getElementById('pv-lib');
    const mdName = document.getElementById('md-name');
    const mdPwr  = document.getElementById('md-pwr');
    const mdLen  = document.getElementById('md-len');
    const mdWid  = document.getElementById('md-wid');
    const pvH = document.getElementById('pv-h');
    const pvV = document.getElementById('pv-v');
    const tpbH = document.getElementById('tab-per-blk-h');
    const tpbV = document.getElementById('tab-per-blk-v');
    const orientSeg = document.getElementById('pv-orient');

    const gbTables = document.getElementById('gb-tables');
    const gbBlocks = document.getElementById('gb-blocks');

    const gapModX = document.getElementById('gap-mod-x');
    const gapModY = document.getElementById('gap-mod-y');
    const gapTabX = document.getElementById('gap-tab-x');
    const gapTabY = document.getElementById('gap-tab-y');
    const gapBlkX = document.getElementById('gap-blk-x');
    const gapBlkY = document.getElementById('gap-blk-y');
    const tilt = document.getElementById('tilt');
    const azimuth = document.getElementById('azimuth');
    const edgeMode = document.getElementById('edge-mode'); // kept for compatibility
    const fillStart = document.getElementById('fill-start');
    const fillMethod = document.getElementById('fill-method');
    const targetKwp = document.getElementById('pv-target-kwp');
    const showModules = document.getElementById('show-modules');
    const pvBuild = document.getElementById('pv-build');
    const pvLock = document.getElementById('pv-lock');

    const alignEdgeBtn = document.getElementById('align-edge');
    const edgeHelp = document.getElementById('edge-help');

    const canvas = document.getElementById('pv-preview');
    const ctx = canvas?.getContext('2d');
    const meta = document.getElementById('pv-preview-meta');
    const zoom = document.getElementById('pv-zoom');
    const zoomIn = document.getElementById('pv-zoom-in');
    const zoomOut = document.getElementById('pv-zoom-out');
    let previewScale = 1;
    let panX = 0, panY = 0;
    let dragging = false, dragStartX = 0, dragStartY = 0;
    const arrangeDetails = document.getElementById('pv-arrange');

    // ===== Helpers to close other UI panels =====
    function closeDrawer(){ drawer?.classList.remove('open'); }
    function closeSearch(){ searchWrap?.classList.remove('open'); searchWrap?.setAttribute('aria-expanded','false'); }

    // ===== PV Sheet open/close =====
    pvBtn?.addEventListener('click', ()=>{
      const willOpen = !pvSheet.classList.contains('open');
      if (willOpen){ closeDrawer(); closeSearch(); }
      pvSheet.classList.toggle('open');
      if (pvSheet.classList.contains('open')) { setTimeout(fitCanvasToContainer, 50); }
    });
    pvClose?.addEventListener('click', ()=> pvSheet.classList.remove('open'));
    toolsBtn?.addEventListener('click', ()=> pvSheet.classList.remove('open'));
    searchBtn?.addEventListener('click', ()=> pvSheet.classList.remove('open'));

    // ===== Module file browse =====
    panBrowse?.addEventListener('click', ()=> panInput?.click());
    panInput?.addEventListener('change', ()=>{ if(panName) panName.textContent = panInput.files?.[0]?.name || 'No file selected'; });

    // ===== Module orientation =====
    orientSeg?.querySelectorAll('button').forEach(btn=>{
      btn.addEventListener('click', ()=>{
        orientSeg.querySelectorAll('button').forEach(b=>b.classList.remove('active'));
        btn.classList.add('active');
        drawPreview();
      });
    });

    // ===== Library presets =====
    const presets = {
      GEN600:{ name:'Generic 600 Wp', pwr:600, len:2100, wid:1134 },
      GEN550:{ name:'Generic 550 Wp', pwr:550, len:2094, wid:1134 },
      GEN450:{ name:'Generic 450 Wp', pwr:450, len:1960, wid:992 },
    };
    pvLib?.addEventListener('change', ()=>{
      const p = presets[pvLib.value];
      if (p){ mdName.value=p.name; mdPwr.value=p.pwr; mdLen.value=p.len; mdWid.value=p.wid; drawPreview(); }
    });

    // ===== Structure Catalog defaults =====
    const structDefaults = {
      'ground-fixed': {
        name: 'Ground — Fixed Tilt',
        tilt: 20, azimuth: 180, orient: 'landscape',
        pvH: 12, pvV: 2, tabX: 0.6, tabY: 6.0, blkX: 3.0, blkY: 8.0, tpbH: 10, tpbV: 3
      },
      'ground-tracker': {
        name: 'Ground — Single Axis',
        tilt: 0, azimuth: 180, orient: 'landscape',
        pvH: 2, pvV: 30, tabX: 5, tabY: .5, blkX: 3.0, blkY: 8.0, tpbH: 12, tpbV: 2
      },
      'roof-flush': {
        name: 'Rooftop — Flush',
        tilt: 10, azimuth: 180, orient: 'landscape',
        pvH: 1, pvV: 3, tabX: 0, tabY: 0.8, blkX: 1.2, blkY: 1.5, tpbH: 5, tpbV: 3
      },
      'roof-triangle': {
        name: 'Rooftop — Triangle',
        tilt: 15, azimuth: 180, orient: 'landscape',
        pvH: 2, pvV: 3, tabX: 0.4, tabY: 1.0, blkX: 1.5, blkY: 2.0, tpbH: 4, tpbV: 3
      },
      'roof-ballasted': {
        name: 'Rooftop — Ballasted',
        tilt: 10, azimuth: 180, orient: 'landscape',
        pvH: 2, pvV: 3, tabX: 0.4, tabY: 0.9, blkX: 1.2, blkY: 1.8, tpbH: 5, tpbV: 3
      },
      'carport': {
        name: 'Specialty — Carport',
        tilt: 4, azimuth: 180, orient: 'landscape',
        pvH: 3, pvV: 10, tabX: 0.3, tabY: 6.0, blkX: 3.0, blkY: 7.5, tpbH: 4, tpbV: 1
      },
      'canopy': {
        name: 'Specialty — Solar Canopy',
        tilt: 5, azimuth: 180, orient: 'portralandscapeit',
        pvH: 3, pvV: 8, tabX: 0.3, tabY: 3.0, blkX: 2.5, blkY: 4.0, tpbH: 3, tpbV: 2
      }
    };

    function openCat(){ if(catBackdrop){ catBackdrop.style.display='block'; catBackdrop.setAttribute('aria-hidden','false'); } }
    function closeCat(){ if(catBackdrop){ catBackdrop.style.display='none'; catBackdrop.setAttribute('aria-hidden','true'); } }
    openCatalog?.addEventListener('click', openCat);
    catClose?.addEventListener('click', closeCat);
    catBackdrop?.addEventListener('click', e=>{ if(e.target===catBackdrop) closeCat(); });
    catGrid?.querySelectorAll('.pv-card').forEach(card=>{
      card.addEventListener('click', ()=>{
        const key = card.getAttribute('data-struct'); const s = structDefaults[key]; if(!s) return;
        if(structPicked) structPicked.textContent = s.name;
        tilt.value = s.tilt; azimuth.value = s.azimuth;
        pvH.value = s.pvH; pvV.value = s.pvV;
        gapTabX.value=s.tabX; gapTabY.value=s.tabY;
        gapBlkX.value=s.blkX; gapBlkY.value=s.blkY;
        tpbH.value=s.tpbH; tpbV.value=s.tpbV;
        orientSeg?.querySelectorAll('button').forEach(b=>b.classList.toggle('active', b.dataset.val===s.orient));
        drawPreview();
        closeCat();
      });
    });

    // ===== PV target controls =====
    function updateTargetControls(){
      const mode = document.querySelector('input[name="pv-target"]:checked')?.value;
      if (mode === 'cap'){
        targetKwp.disabled = false;
      } else {
        targetKwp.disabled = true;
      }
    }
    document.querySelectorAll('input[name="pv-target"]').forEach(r=>r.addEventListener('change', updateTargetControls));
    updateTargetControls();

    // ===== Grouping related enable/disable =====
    function updateGapControls(){
      const useTables = !!(gbTables && gbTables.checked);
      const useBlocks = !!(gbBlocks && gbBlocks.checked);

      if (gapTabX) gapTabX.disabled = !useTables;
      if (gapTabY) gapTabY.disabled = !useTables;
      if (pvH) pvH.disabled = !useTables;
      if (pvV) pvV.disabled = !useTables;

      if (gapBlkX) gapBlkX.disabled = !useBlocks;
      if (gapBlkY) gapBlkY.disabled = !useBlocks;
      if (tpbH)    tpbH.disabled    = !useBlocks;
      if (tpbV)    tpbV.disabled    = !useBlocks;
    }
    gbTables?.addEventListener('change', ()=>{ updateGapControls(); drawPreview(); });
    gbBlocks?.addEventListener('change', ()=>{ updateGapControls(); drawPreview(); });
    updateGapControls();

    // ===== Live preview inputs =====
    [mdPwr, mdLen, mdWid, pvH, pvV, tpbH, tpbV, gapModX, gapModY, gapTabX, gapTabY, gapBlkX, gapBlkY, tilt, azimuth, gbTables, gbBlocks].forEach(inp=>{
      inp?.addEventListener?.('input', drawPreview);
      inp?.addEventListener?.('change', drawPreview);
    });

    // ===== Canvas sizing & zoom/pan =====
    function fitCanvasToContainer(){
      if (!canvas) return;
      const w = canvas.clientWidth || 300;
      const h = canvas.clientHeight || 200;
      if (canvas.width !== w) canvas.width = w;
      if (canvas.height !== h) canvas.height = h;
      drawPreview();
    }
    arrangeDetails?.addEventListener('toggle', ()=>{ if(arrangeDetails.open) { setTimeout(fitCanvasToContainer, 50); } });

    zoom?.addEventListener('input', ()=>{ previewScale=parseFloat(zoom.value); drawPreview(); });
    zoomIn?.addEventListener('click', ()=>{ zoom.value=Math.min(20, (parseFloat(zoom.value)+0.2)).toFixed(2); previewScale=parseFloat(zoom.value); drawPreview(); });
    zoomOut?.addEventListener('click', ()=>{ zoom.value=Math.max(0.5, (parseFloat(zoom.value)-0.2)).toFixed(2); previewScale=parseFloat(zoom.value); drawPreview(); });
    canvas?.addEventListener('wheel', (e)=>{
      e.preventDefault();
      const delta = e.deltaY>0 ? -0.5 : 0.5;
      const old = parseFloat(zoom.value);
      const next = Math.min(20, Math.max(0.5, old + delta));
      zoom.value = next.toFixed(2); previewScale = parseFloat(zoom.value);
      drawPreview();
    }, {passive:false});
    canvas?.addEventListener('mousedown', (e)=>{ dragging=true; dragStartX=e.clientX; dragStartY=e.clientY; });
    window.addEventListener('mousemove', (e)=>{ if(!dragging) return; panX += (e.clientX - dragStartX); panY += (e.clientY - dragStartY); dragStartX=e.clientX; dragStartY=e.clientY; drawPreview(); });
    window.addEventListener('mouseup', ()=>{ dragging=false; });

    // ===== Preview configuration & drawing =====
    function getPreviewCfg(){
      const orient = orientSeg?.querySelector('.active')?.getAttribute('data-val') || 'portrait';
      const H = Math.max(1, Number(pvH.value||1));
      const V = Math.max(1, Number(pvV.value||1));
      const modLenM = Number(mdLen.value||0)/1000;
      const modWidM = Number(mdWid.value||0)/1000;
      const modX = orient==='portrait' ? modWidM : modLenM;
      const modY = orient==='portrait' ? modLenM : modWidM;

      const useTables = !!(gbTables && gbTables.checked);
      const useBlocks = !!(gbBlocks && gbBlocks.checked);

      const cfg = {
        module:{ name: mdName.value||'Module', pwr: Number(mdPwr.value||0), len: modLenM, wid: modWidM, orient, modX, modY },
        table:{ modH: H, modV: V, gapModX: Number(gapModX.value||0), gapModY: Number(gapModY.value||0) },
        block:{ tH: Math.max(1, Number(tpbH.value||1)), tV: Math.max(1, Number(tpbV.value||1)) },
        gaps:{
          tabX: useTables ? Number(gapTabX.value||0) : 0,
          tabY: useTables ? Number(gapTabY.value||0) : 0,
          blkX: useBlocks ? Number(gapBlkX.value||0) : 0,
          blkY: useBlocks ? Number(gapBlkY.value||0) : 0
        },
        orientation:{ tilt: Number(tilt.value||0), azimuth: Number(azimuth.value||0) }
      };

      if (!useTables) {
        cfg.table.modH = 1;
        cfg.table.modV = 1;
        cfg.gaps.tabX = Number(gapModX.value||0);
        cfg.gaps.tabY = Number(gapModY.value||0);
        cfg.table.gapModX = 0;
        cfg.table.gapModY = 0;
      }

      const modXdim = modX;
      const modYdim = modY * Math.cos((cfg.orientation.tilt||0) * Math.PI/180);
      cfg._modXdim = modXdim;
      cfg._modYdim = modYdim;

      cfg.table.w = cfg.table.modH * modXdim + (cfg.table.modH - 1) * cfg.table.gapModX;
      cfg.table.h = cfg.table.modV * modYdim + (cfg.table.modV - 1) * cfg.table.gapModY;
      cfg.table.kwp = (cfg.table.modH * cfg.table.modV * cfg.module.pwr) / 1000;

      if (!useBlocks) {
        cfg.block.tH = 1;
        cfg.block.tV = 1;
        cfg.gaps.blkX = 0;
        cfg.gaps.blkY = 0;
      }

      cfg.block.w = (cfg.block.tH * cfg.table.w) + ((cfg.block.tH - 1) * cfg.gaps.tabX);
      cfg.block.h = (cfg.block.tV * cfg.table.h) + ((cfg.block.tV - 1) * cfg.gaps.tabY);

      return cfg;
    }

    function drawPreview(){
      if (!canvas || !ctx) return;
      const cfg = getPreviewCfg();
      const useTables = !!(gbTables && gbTables.checked);
      const W = canvas.width, Hc = canvas.height;
      ctx.clearRect(0,0,W,Hc);

      const colsB = 2, rowsB = 2;
      const padPx = 12;

      const contentW = colsB*cfg.block.w + (colsB-1)*cfg.gaps.blkX;
      const contentH = rowsB*cfg.block.h + (rowsB-1)*cfg.gaps.blkY;

      const theta = (cfg.orientation.azimuth||0) * Math.PI/180;
      const cosT = Math.cos(theta), sinT = Math.sin(theta);
      const uvToXY = (u,v)=>({ x: cosT*u + sinT*v, y: -sinT*u + cosT*v });

      const cornersUV = [{u:0,v:0},{u:contentW,v:0},{u:contentW,v:contentH},{u:0,v:contentH}];
      const cornersXY = cornersUV.map(p=>uvToXY(p.u,p.v));
      const minX = Math.min(...cornersXY.map(p=>p.x));
      const maxX = Math.max(...cornersXY.map(p=>p.x));
      const minY = Math.min(...cornersXY.map(p=>p.y));
      const maxY = Math.max(...cornersXY.map(p=>p.y));
      const rotW = maxX - minX, rotH = maxY - minY;

      const scaleBase = Math.min((W - 2*padPx)/rotW, (Hc - 2*padPx)/rotH);
      const scale = Math.max(0.0001, scaleBase * previewScale);

      const offX = (W - rotW*scale)/2 - (minX*scale) + panX;
      const offY = (Hc - rotH*scale)/2 - (minY*scale) + panY;

      function pathRectUV(u0,v0,w,h){
        const p0 = uvToXY(u0,     v0);
        const p1 = uvToXY(u0+w,   v0);
        const p2 = uvToXY(u0+w,   v0+h);
        const p3 = uvToXY(u0,     v0+h);
        ctx.beginPath();
        ctx.moveTo(offX + p0.x*scale, offY + p0.y*scale);
        ctx.lineTo(offX + p1.x*scale, offY + p1.y*scale);
        ctx.lineTo(offX + p2.x*scale, offY + p2.y*scale);
        ctx.lineTo(offX + p3.x*scale, offY + p3.y*scale);
        ctx.closePath();
      }

      ctx.lineWidth = 1;

      const tW = cfg.table.w, tH = cfg.table.h, tGapX = cfg.gaps.tabX, tGapY = cfg.gaps.tabY;

      for (let br=0;br<rowsB;br++){
        for (let bc=0;bc<colsB;bc++){
          const bu = bc*(cfg.block.w + cfg.gaps.blkX);
          const bv = br*(cfg.block.h + cfg.gaps.blkY);

          ctx.strokeStyle = '#9ca3af';
          pathRectUV(bu, bv, cfg.block.w, cfg.block.h);
          ctx.stroke();

          for(let rr=0; rr<cfg.block.tV; rr++){
            for(let cc=0; cc<cfg.block.tH; cc++){
              const tu = bu + cc*(tW + tGapX);
              const tv = bv + rr*(tH + tGapY);

              if (useTables) {
                ctx.fillStyle = '#e2e8f0';
                ctx.strokeStyle = '#64748b';
                pathRectUV(tu, tv, tW, tH);
                ctx.fill();
                ctx.stroke();
              }

              const mWpx = cfg._modXdim * scale;
              const mHpx = cfg._modYdim * scale;
              if (mWpx > 4 && mHpx > 4){
                for (let mr=0; mr<cfg.table.modV; mr++){
                  for (let mc=0; mc<cfg.table.modH; mc++){
                    const mu = tu + mc*(cfg._modXdim + cfg.table.gapModX);
                    const mv = tv + mr*(cfg._modYdim + cfg.table.gapModY);
                    ctx.fillStyle = '#f8fafc';
                    ctx.strokeStyle = '#94a3b8';
                    pathRectUV(mu, mv, cfg._modXdim, cfg._modYdim);
                    ctx.fill(); ctx.stroke();
                  }
                }
              }
            }
          }
        }
      }

      meta.textContent = `Table: ${cfg.table.w.toFixed(2)}m × ${cfg.table.h.toFixed(2)}m · ${cfg.table.kwp.toFixed(2)} kWp/table · Block: ${cfg.block.tH*cfg.block.tV} tables`;
    }
    drawPreview();

    // ===== Fill Method Toggle (Adaptive / Block) =====
    const fillMethodSeg = document.getElementById("fill-method");
    if (fillMethodSeg) {
      fillMethodSeg.querySelectorAll("button").forEach(btn => {
        btn.addEventListener("click", () => {
          fillMethodSeg.querySelectorAll("button")
            .forEach(b => b.classList.remove("active"));
          btn.classList.add("active");
          console.log("Fill method changed to:", btn.dataset.val);
        });
      });
    }

    // ===== Overlay registry patch (for later selection/clear) =====
    window.__overlayRegistry = window.__overlayRegistry || new Set();
    function patchOverlayRegistry(){
      const g = window._google; if(!g) return;
      ['Polygon','Rectangle','Circle','Polyline'].forEach(klass=>{
        const C = g.maps[klass]; if(!C || C.__patched) return;
        const proto = C.prototype; const orig = proto.setMap;
        proto.setMap = function(map){ try{ window.__overlayRegistry.add(this); }catch(_e){} return orig.apply(this, arguments); };
        C.__patched = true;
      });
    }
    window.addEventListener('load', ()=>setTimeout(patchOverlayRegistry, 0));

    function getOverlaysBy(pred){
      const out=[]; window.__overlayRegistry?.forEach(ov=>{ try{ if(ov.getMap && ov.getMap() && pred(ov)) out.push(ov); }catch(_e){} });
      return out;
    }
    


        function getAreasSelected(){
      const sel = getOverlaysBy(ov=>ov.__cat==='area' && ov.__selected===true);
      if(sel.length) return sel;
      return [];
    }

    // Normalized polygon path for any overlay (polygon / rectangle / circle)
        function polygonPathFromOverlay(ov) {
      const g = window._google;
      if (!ov) return [];

      // 1) Generic polygon (also used for rotated rectangles we converted)
      if (ov.getPath) {
        return ov.getPath().getArray();
      }

      // 2) Rectangle → 5‑point closed path
      if (ov.getBounds && g?.maps?.drawing && ov.__type === g.maps.drawing.OverlayType.RECTANGLE) {
        const b  = ov.getBounds();
        const ne = b.getNorthEast();
        const sw = b.getSouthWest();
        const nw = new g.maps.LatLng(ne.lat(), sw.lng());
        const se = new g.maps.LatLng(sw.lat(), ne.lng());
        return [nw, ne, se, sw, nw];
      }

      // 3) Circle → approximated with 64 segments
      if (ov.getCenter && ov.getRadius && g?.maps?.drawing && ov.__type === g.maps.drawing.OverlayType.CIRCLE) {
        const c   = ov.getCenter();
        const r   = ov.getRadius();
        const pts = [];
        const N   = 64;
        for (let i = 0; i < N; i++) {
          pts.push(g.maps.geometry.spherical.computeOffset(c, r, i * (360 / N)));
        }
        pts.push(pts[0]); // close
        return pts;
      }

      return [];
    }

    // Make it available globally (DXF exporter uses this)
    window.polygonPathFromOverlay = polygonPathFromOverlay;

    // expose helper globally for DXF + other tools
    window.polygonPathFromOverlay = polygonPathFromOverlay;


    function makeLocal(origin){
      const g=window._google;
      return {
        toXY(p){
          const d = g.maps.geometry.spherical.computeDistanceBetween(origin, p);
          const h = g.maps.geometry.spherical.computeHeading(origin, p) * Math.PI/180;
          return { x: d*Math.sin(h), y: d*Math.cos(h) };
        },
        fromXY(x,y){
          const d = Math.hypot(x,y);
          const h = Math.atan2(x,y) * 180/Math.PI;
          return g.maps.geometry.spherical.computeOffset(origin, d, h);
        }
      };
    }
    function centroid(path){
      const g=window._google;
      let lat=0,lng=0; const n=path.length;
      path.forEach(p=>{ lat+=p.lat(); lng+=p.lng(); });
      return new g.maps.LatLng(lat/n, lng/n);
    }
    function contains(poly, latLng){ return window._google.maps.geometry.poly.containsLocation(latLng, poly); }

    function findNearestEdge(areas, latLng){
      const g=window._google;
      let best={dist2:Infinity, angle:0};
      areas.forEach(ov=>{
        const path = polygonPathFromOverlay(ov);
        if (!path.length) return;
        const org = centroid(path);
        const prj = makeLocal(org);
        const P = prj.toXY(latLng);
        const pts = path.map(p=>prj.toXY(p));
        const n = pts.length;
        for (let i=0;i<n;i++){
          const a=pts[i], b=pts[(i+1)%n];
          const vx=b.x-a.x, vy=b.y-a.y;
          const wx=P.x-a.x, wy=P.y-a.y;
          const denom = vx*vx+vy*vy || 1;
          let t=(wx*vx+wy*vy)/denom; if(t<0)t=0; if(t>1)t=1;
          const qx=a.x+t*vx, qy=a.y+t*vy;
          const dx=P.x-qx, dy=P.y-qy;
          const d2=dx*dx+dy*dy;
          if (d2<best.dist2){
            const heading = Math.atan2(vx, vy)*180/Math.PI;
            best={dist2:d2, angle:(heading+360)%360, edgeIndex:i, ov};
          }
        }
      });
      return best;
    }

    // ===== Align to edge (map click) =====
    let edgePickListener=null;
    alignEdgeBtn?.addEventListener('click', ()=>{
      const g=window._google, map=window._map;
      if(!g||!map){ alert('Map not ready'); return; }
      const areas = getAreasSelected();
      if(!areas.length){ alert('Select a PV Area first.'); return; }
      if(edgePickListener){ g.maps.event.removeListener(edgePickListener); edgePickListener=null; }
      alignEdgeBtn.textContent = 'Click an edge on map…';
      edgeHelp.style.display='inline';
      map.setOptions({ draggableCursor:'crosshair' });
      edgePickListener = g.maps.event.addListener(map, 'click', (e)=>{
        const pick = findNearestEdge(areas, e.latLng);
        const az = Math.round((pick.angle+360)%360);
        azimuth.value = az;
        drawPreview();
        g.maps.event.removeListener(edgePickListener); edgePickListener=null;
        map.setOptions({ draggableCursor:null });
        edgeHelp.style.display='none';
        alignEdgeBtn.textContent = 'Pick edge on map';
      });
    });

    // ===== Layout creation / editing lock =====
    // Shared list of PV table polygons (used by summary + DXF)
    window.layoutOverlays = window.layoutOverlays || [];
    const layoutOverlays = window.layoutOverlays;

    function clearLayoutForAreas(areas){
      const toRemove = new Set(areas);
      // mutate in place so window.layoutOverlays and local stay in sync
      for (let i = layoutOverlays.length - 1; i >= 0; i--) {
        const ov = layoutOverlays[i];
        if (toRemove.has(ov.__parentArea)) {
          try { ov.setMap(null); } catch (_e) {}
          layoutOverlays.splice(i, 1);
        }
      }
    }

    pvLock?.addEventListener('click', ()=>{
      if (pvLock.dataset.state==='unlocked'){
        pvLock.dataset.state='locked'; pvLock.textContent='Unlock editing';
        layoutOverlays.forEach(ov=>{ try{ ov.setEditable(false); ov.setDraggable(false);}catch(_e){} });
      }else{
        pvLock.dataset.state='unlocked'; pvLock.textContent='Lock editing';
        layoutOverlays.forEach(ov=>{ try{ ov.setEditable(true); ov.setDraggable(true);}catch(_e){} });
      }
    });

    pvBuild?.addEventListener('click', buildSimpleFill);

    // === Inset helpers (LL = LatLng space) ===
    function avgHeading(h1, h2){
      const r1 = h1*Math.PI/180, r2 = h2*Math.PI/180;
      const x = Math.cos(r1)+Math.cos(r2), y = Math.sin(r1)+Math.sin(r2);
      return Math.atan2(y, x) * 180/Math.PI;
    }
    function signedAreaApproxLL(path){
      let a = 0, n = path.length;
      for (let i=0;i<n;i++){
        const p = path[i], q = path[(i+1)%n];
        a += (q.lng() - p.lng()) * (q.lat() + p.lat());
      }
      return -a; // ccw positive
    }
    function insetPolygonLL(pathLL, distMeters){
      const g = window._google;
      if (!g || !pathLL || pathLL.length < 3) return null;

      const closed = pathLL[0].equals(pathLL[pathLL.length-1]) ? pathLL.slice(0,-1) : pathLL.slice();
      if (closed.length < 3) return null;

      const ccw = signedAreaApproxLL(closed) > 0;
      const N = closed.length;
      const out = [];
      for (let i=0; i<N; i++){
        const prev = closed[(i-1+N)%N], curr = closed[i], next = closed[(i+1)%N];
        const hPrev = window._google.maps.geometry.spherical.computeHeading(prev, curr);
        const hNext = window._google.maps.geometry.spherical.computeHeading(curr, next);
        const hAvg  = avgHeading(hPrev, hNext);
        const normal = ccw ? (hAvg - 90) : (hAvg + 90);
        out.push(window._google.maps.geometry.spherical.computeOffset(curr, distMeters, normal));
      }
      return out;
    }

    // === Helper: scanline intervals in UV for a given v ===
    function scanlineIntervals(polyUV, vRow){
      const intersections = [];
      const n = polyUV.length;
      for (let i = 0; i < n; i++){
        const p = polyUV[i];
        const q = polyUV[(i+1) % n];
        const v1 = p.v, v2 = q.v;
        if ((v1 <= vRow && v2 > vRow) || (v2 <= vRow && v1 > vRow)) {
          const t = (vRow - v1) / (v2 - v1);
          const u = p.u + t * (q.u - p.u);
          intersections.push(u);
        }
      }
      intersections.sort((a,b)=>a-b);
      const segments = [];
      for (let i=0; i+1<intersections.length; i+=2){
        segments.push({ u0: intersections[i], u1: intersections[i+1] });
      }
      return segments;
    }

    // ===== Core: Simple Fill builder (TABLES ONLY, 2 GW cap, Adaptive vs Block) =====
    function buildSimpleFill(){
      const g = window._google, map = window._map;
      if (!g || !map) { alert('Map not ready'); return; }

      const areas = getAreasSelected();
      if (!areas.length) { alert('Select a PV Area first.'); return; }

      clearLayoutForAreas(areas);

      const cfg = getPreviewCfg();
      const userSetback = parseFloat(document.getElementById('in-setback-boundary')?.value || '1');

      const TAB_W = cfg.table.w;
      const TAB_H = cfg.table.h;
      const TAB_GX = cfg.gaps.tabX;
      const TAB_GY = cfg.gaps.tabY;

      const BLK_W = cfg.block.w;
      const BLK_H = cfg.block.h;
      const BLK_GX = cfg.gaps.blkX;
      const BLK_GY = cfg.gaps.blkY;

      const TABLE_MOD_H = cfg.table.modH;
      const TABLE_MOD_V = cfg.table.modV;
      let BLOCK_TAB_H = cfg.block.tH;
      let BLOCK_TAB_V = cfg.block.tV;

      const STYLE = {
        strokeColor: '#1e3a8a',
        strokeOpacity: 0.7,
        strokeWeight: 0.6,
        fillColor: '#3a7afe',
        fillOpacity: 0.9,
        zIndex: 1000
      };

      // capacity limit = 2,000,000 kWp (2 GW)
      const MAX_KWP = 2_000_000;
      const KWP_PER_TABLE = (TABLE_MOD_H * TABLE_MOD_V * (cfg.module.pwr || 0)) / 1000;

      let totalTables = 0;
      let totalKwp = 0;
      let hitCap = false;

      const theta = (cfg.orientation.azimuth || 0) * Math.PI / 180;
      const cosT = Math.cos(theta), sinT = Math.sin(theta);

      function uvFromLL(latLng, proj){
        const xy = proj.toXY(latLng);
        const u = cosT * xy.x - sinT * xy.y;
        const v = sinT * xy.x + cosT * xy.y;
        return { u, v };
      }
      function llFromUV(u,v, proj){
        const x = cosT * u + sinT * v;
        const y = -sinT * u + cosT * v;
        return proj.fromXY(x,y);
      }

      function safeContain(poly, latLng, tolerance = 0.02) {
        if (contains(poly, latLng)) return true;
        const center = centroid(poly.getPath().getArray());
        const heading = g.maps.geometry.spherical.computeHeading(latLng, center);
        const inward = g.maps.geometry.spherical.computeOffset(latLng, tolerance, heading);
        return contains(poly, inward);
      }

      function intersectsAnyBlocker(cornersLL, blockers){
        for (const blocker of blockers){
          for (const pt of cornersLL){ if (safeContain(blocker, pt)) return true; }
          const blockerCenter = centroid(blocker.getPath().getArray());
          const testPoly = new g.maps.Polygon({ paths: cornersLL });
          if (safeContain(testPoly, blockerCenter)) return true;
        }
        return false;
      }

      const method = fillMethod?.querySelector('.active')?.getAttribute('data-val') || 'adaptive';
      const useBlocks = (method === 'block') && (cfg.block.tH > 1 || cfg.block.tV > 1);
      if (!useBlocks) {
        BLOCK_TAB_H = 1;
        BLOCK_TAB_V = 1;
      }

      // blockers once
      const allOverlays = window.__overlayRegistry ? Array.from(window.__overlayRegistry) : [];
      const blockers = allOverlays
        .filter(ov => (ov.__cat === 'constraint' || ov.__cat === 'object') && ov.getMap())
        .map(ov => new g.maps.Polygon({ paths: polygonPathFromOverlay(ov) }));

      // ===== ADAPTIVE MODE (scanline inside polygon) =====
      if (method === 'adaptive') {
        areas.forEach(ovArea => {
          if (hitCap) return;

          let path = polygonPathFromOverlay(ovArea);
          if (!path.length) return;

          const setback = Math.max(0, userSetback || 0);
          if (setback > 0) {
            const inset = insetPolygonLL(path, setback);
            if (!inset || inset.length < 3) return;
            path = inset;
          }

          const polyArea = new g.maps.Polygon({ paths: path });
          const origin = centroid(path);
          const proj = makeLocal(origin);

          const uvPoly = path.map(p => uvFromLL(p, proj));
          const n = uvPoly.length;
          if (n < 3) return;

          const vVals = uvPoly.map(p => p.v);
          const minV = Math.min(...vVals);
          const maxV = Math.max(...vVals);
          const rowPitch = TAB_H + TAB_GY;
          const TOL = 0.01;

          for (let rowV = minV; rowV <= maxV - TAB_H + TOL; rowV += rowPitch) {
            if (hitCap) break;

            const segments = scanlineIntervals(uvPoly, rowV);
            for (const seg of segments) {
              if (hitCap) break;

              const uStart = seg.u0;
              const uEnd   = seg.u1;
              // small inward margin to avoid grazing edges
              const margin = 0.05;
              let u = uStart + margin;

              while (u + TAB_W <= uEnd - margin + 1e-6) {
                if (hitCap) break;

                const tableCornersUV = [
                  { u: u,          v: rowV },
                  { u: u + TAB_W,  v: rowV },
                  { u: u + TAB_W,  v: rowV + TAB_H },
                  { u: u,          v: rowV + TAB_H }
                ];
                const tableCornersLL = tableCornersUV.map(p => llFromUV(p.u, p.v, proj));

                if (!tableCornersLL.every(pt => safeContain(polyArea, pt))) {
                  u += TAB_W + TAB_GX;
                  continue;
                }
                if (intersectsAnyBlocker(tableCornersLL, blockers)) {
                  u += TAB_W + TAB_GX;
                  continue;
                }

                if (totalKwp + KWP_PER_TABLE > MAX_KWP) {
                  console.warn('DC layout limit of 2,000,000 kWp reached. Stopping further tables (adaptive).');
                  hitCap = true;
                  break;
                }

                const tablePoly = new g.maps.Polygon(Object.assign({
                  map,
                  paths: tableCornersLL,
                  clickable: true,
                  draggable: (pvLock.dataset.state !== 'locked'),
                  editable: (pvLock.dataset.state !== 'locked')
                }, STYLE));

                tablePoly.__areaName = ovArea.__name || 'Area';
                tablePoly.__kwp = KWP_PER_TABLE;
                tablePoly.__tilt = cfg.orientation.tilt;
                tablePoly.__az = cfg.orientation.azimuth;
                tablePoly.__parentArea = ovArea;
                layoutOverlays.push(tablePoly);

                totalTables++;
                totalKwp += KWP_PER_TABLE;

                u += TAB_W + TAB_GX;
              }
            }
          }
        });

        if (hitCap) {
          alert('DC layout limit of 2,000,000 kWp (2 GW) reached. Not all possible tables were drawn (adaptive).');
        }
        console.log(`✅ Adaptive layout: ${totalTables} tables, ${totalKwp.toFixed(1)} kWp DC.`);
        updateSummary();
        return;
      }

      // ===== BLOCK MODE (original stepping with blocks) =====
      areas.forEach(ovArea => {
        if (hitCap) return;

        let path = polygonPathFromOverlay(ovArea);
        if (!path.length) return;

        const setback = Math.max(0, userSetback || 0);
        if (setback > 0) {
          const inset = insetPolygonLL(path, setback);
          if (!inset || inset.length < 3) return;
          path = inset;
        }

        const polyArea = new g.maps.Polygon({ paths: path });
        const origin = centroid(path);
        const proj = makeLocal(origin);

        const uvPoly = path.map(p => uvFromLL(p, proj));
        const uVals = uvPoly.map(p => p.u);
        const vVals = uvPoly.map(p => p.v);
        const minU = Math.min(...uVals);
        const maxU = Math.max(...uVals);
        const minV = Math.min(...vVals);
        const maxV = Math.max(...vVals);

        const startU = minU;
        const startV = minV;

        const STEP_W = useBlocks ? (BLK_W + BLK_GX) : (TAB_W + TAB_GX);
        const STEP_H = useBlocks ? (BLK_H + BLK_GY) : (TAB_H + TAB_GY);
        const TOL = 0.01;

        const nRows = Math.floor((maxV - minV + TOL) / STEP_H);
        const nCols = Math.floor((maxU - minU + TOL) / STEP_W);

        for (let r = 0; r < nRows; r++) {
          const bv = Math.min(startV + r * STEP_H, maxV - (useBlocks ? BLK_H : TAB_H));
          if (hitCap) break;
          for (let c = 0; c < nCols; c++) {
            const bu = Math.min(startU + c * STEP_W, maxU - (useBlocks ? BLK_W : TAB_W));
            if (hitCap) break;

            const buClamped = Math.min(bu, maxU - (useBlocks ? BLK_W : TAB_W) + 0.001);
            const bvClamped = Math.min(bv, maxV - (useBlocks ? BLK_H : TAB_H) + 0.001);

            const blockCornersUV = [
              { u: buClamped, v: bvClamped },
              { u: buClamped + (useBlocks ? BLK_W : TAB_W), v: bvClamped },
              { u: buClamped + (useBlocks ? BLK_W : TAB_W), v: bvClamped + (useBlocks ? BLK_H : TAB_H) },
              { u: buClamped, v: bvClamped + (useBlocks ? BLK_H : TAB_H) }
            ];
            const blockCornersLL = blockCornersUV.map(p => llFromUV(p.u, p.v, proj));
            if (!blockCornersLL.every(pt => safeContain(polyArea, pt))) continue;
            if (intersectsAnyBlocker(blockCornersLL, blockers)) continue;

            for (let tr = 0; tr < BLOCK_TAB_V; tr++) {
              const tabV = bvClamped + tr * (TAB_H + TAB_GY);
              if (hitCap) break;
              for (let tc = 0; tc < BLOCK_TAB_H; tc++) {
                const tabU = buClamped + tc * (TAB_W + TAB_GX);

                const tableCornersUV = [
                  { u: tabU, v: tabV },
                  { u: tabU + TAB_W, v: tabV },
                  { u: tabU + TAB_W, v: tabV + TAB_H },
                  { u: tabU, v: tabV + TAB_H }
                ];
                const tableCornersLL = tableCornersUV.map(p => llFromUV(p.u, p.v, proj));

                if (!tableCornersLL.every(pt => safeContain(polyArea, pt))) continue;
                if (intersectsAnyBlocker(tableCornersLL, blockers)) continue;

                if (totalKwp + KWP_PER_TABLE > MAX_KWP) {
                  console.warn('DC layout limit of 2,000,000 kWp reached. Stopping further tables (block).');
                  hitCap = true;
                  break;
                }

                const tablePoly = new g.maps.Polygon(Object.assign({
                  map,
                  paths: tableCornersLL,
                  clickable: true,
                  draggable: (pvLock.dataset.state !== 'locked'),
                  editable: (pvLock.dataset.state !== 'locked')
                }, STYLE));

                tablePoly.__areaName = ovArea.__name || 'Area';
                tablePoly.__kwp = KWP_PER_TABLE;
                tablePoly.__tilt = cfg.orientation.tilt;
                tablePoly.__az = cfg.orientation.azimuth;
                tablePoly.__parentArea = ovArea;
                layoutOverlays.push(tablePoly);

                totalTables++;
                totalKwp += KWP_PER_TABLE;
              }
            }
          }
        }
      });

      if (hitCap) {
        alert('DC layout limit of 2,000,000 kWp (2 GW) reached. Not all possible tables were drawn (block).');
      }

      console.log(`✅ Block layout: ${totalTables} tables, ${totalKwp.toFixed(1)} kWp DC.`);
      updateSummary();
    }

    // ===== Summary rendering =====
    function updateSummary() {
      const g = window._google;
      if (!g) return;

      const sumBody = document.getElementById('sum-body');
      const totDcEl = document.getElementById('tot-dc');
      const totAcEl = document.getElementById('tot-ac');
      const totM2El = document.getElementById('tot-m2');

      if (!sumBody) return;

      sumBody.innerHTML = '';

      const areas = [];
      window.layoutOverlays.forEach(poly => {
        const name = poly.__areaName || 'Area';
        let rec = areas.find(a => a.name === name);
        if (!rec) {
          rec = { name, modules: 0, kwp: 0, area: 0, tilt: 0, az: 0, count: 0 };
          areas.push(rec);
        }
        rec.modules++; // now effectively "tables"
        rec.kwp += poly.__kwp || 0;
        rec.count++;
        try {
          rec.area += g.maps.geometry.spherical.computeArea(poly.getPath());
        } catch (_) {}
        if (poly.__tilt) rec.tilt = poly.__tilt;
        if (poly.__az) rec.az = poly.__az;
      });

      if (!areas.length) {
        sumBody.innerHTML = `<tr><td colspan="6" class="muted" style="padding:12px 10px">No areas yet</td></tr>`;
        return;
      }

      let totalDC = 0, totalAC = 0, totalArea = 0;
      const rows = [];
      areas.forEach(a => {
        totalDC += a.kwp;
        totalArea += a.area;
        rows.push(`
          <tr>
            <td>${a.name}</td>
            <td class="num">${a.kwp.toFixed(1)}</td>
            <td class="num">${(a.kwp*0.8).toFixed(1)}</td>
            <td class="num">${(a.area).toFixed(0)}</td>
            <td class="num">${a.tilt.toFixed(0)}</td>
            <td class="num">${a.az.toFixed(0)}</td>
          </tr>`);
      });

      sumBody.innerHTML = rows.join('');
      totDcEl.textContent = totalDC.toFixed(1);
      totAcEl.textContent = (totalDC*0.8).toFixed(1);
      totM2El.textContent = totalArea.toFixed(0);
    }

    // ===== Export to Excel =====
    async function exportSummaryToExcel() {
      try {
        const table = document.getElementById('sum-body');
        if (!table || !table.rows.length) {
          alert('No summary data to export.');
          return;
        }

        const data = [['Area Name', 'DC kWp', 'AC kWp', 'Area (m²)', 'Tilt (°)', 'Azimuth (°)']];
        for (const row of table.rows) {
          const cells = [...row.cells].map(td => td.textContent.trim());
          if (cells.length >= 6 && !cells[0].includes('No areas')) data.push(cells);
        }

        await Excel.run(async context => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const startCell = sheet.getRange("B5");
          const range = startCell.getResizedRange(data.length - 1, data[0].length - 1);
          range.values = data;
          range.format.autofitColumns();
          range.format.autofitRows();
          await context.sync();
        });

        console.log('✅ Summary exported to Excel.');
      } catch (err) {
        console.error('Export failed:', err);
        alert('Export failed. Check console for details.');
      }
    }
    document.getElementById('btn-export-excel')?.addEventListener('click', exportSummaryToExcel);

    // ===== Summary collapse toggle =====
    const summaryEl = document.getElementById('summary');
    const sumToggle = document.getElementById('sum-toggle');
    function setSummary(open){
      summaryEl.classList.toggle('collapsed', !open);
      summaryEl.setAttribute('aria-expanded', String(open));
    }
    setSummary(false);
    sumToggle?.addEventListener('click', ()=>{
      const willOpen = summaryEl.classList.contains('collapsed');
      setSummary(willOpen);
    });

    // ===== Initial sizing =====
    fitCanvasToContainer();
  });


function exportToDXF() {
  const g = window._google;
  if (!g) {
    console.error("Map not ready");
    return;
  }

  // ========= HANDLE GENERATOR (group code 5) =========
  let handleCounter = 0;
  const nextHandle = () => (++handleCounter).toString(16).toUpperCase();

  // ========= BASIC HELPERS =========
  function makeLocal(origin) {
    return {
      toXY(p) {
        const d = g.maps.geometry.spherical.computeDistanceBetween(origin, p);
        const h = g.maps.geometry.spherical.computeHeading(origin, p) * Math.PI / 180;
        return { x: d * Math.sin(h), y: d * Math.cos(h) }; // X = east, Y = north
      }
    };
  }

  function centroid(path) {
    let lat = 0, lng = 0;
    const n = path.length || 1;
    path.forEach(p => { lat += p.lat(); lng += p.lng(); });
    return new g.maps.LatLng(lat / n, lng / n);
  }

  function getPathForExport(ov) {
    if (!ov) return [];

    // If you have a shared helper in your app, use it
    if (typeof window.polygonPathFromOverlay === "function") {
      try {
        const path = window.polygonPathFromOverlay(ov);
        if (path && path.length) return path;
      } catch (e) {
        console.warn("polygonPathFromOverlay failed:", e);
      }
    }

    // Polygon / Polyline style overlays
    if (ov.getPath) {
      return ov.getPath().getArray();
    }

    // Rectangle → 4 corners
    if (ov.getBounds) {
      const b = ov.getBounds();
      const ne = b.getNorthEast();
      const sw = b.getSouthWest();
      const nw = new g.maps.LatLng(ne.lat(), sw.lng());
      const se = new g.maps.LatLng(sw.lat(), ne.lng());
      return [nw, ne, se, sw, nw];
    }

    // Circle → approximate with 64-sided polygon
    if (ov.getCenter && ov.getRadius) {
      const c = ov.getCenter();
      const r = ov.getRadius();
      const pts = [];
      const N = 64;
      for (let i = 0; i < N; i++) {
        pts.push(g.maps.geometry.spherical.computeOffset(c, r, i * (360 / N)));
      }
      pts.push(pts[0]);
      return pts;
    }

    return [];
  }

  function latLngToXY(latlng, proj) {
    const xy = proj.toXY(latlng);
    return {
      x: Number(xy.x.toFixed(3)),
      y: Number(xy.y.toFixed(3))
    };
  }

  // Emit one closed POLYLINE (R12) with VERTEX and SEQEND
  function polyToDXF(points, layer) {
    if (!points || !points.length) return "";

    // Ensure closed polyline
    const closed = points.slice();
    const first = closed[0];
    const last = closed[closed.length - 1];
    if (first.x !== last.x || first.y !== last.y) {
      closed.push({ x: first.x, y: first.y });
    }

    const polyHandle = nextHandle();
    const vertexHandles = closed.map(() => nextHandle());
    const seqHandle = nextHandle();

    // Simple color scheme by layer (ACI index)
    let colorIndex = 7; // white by default
    if (layer === "PV_AREAS") colorIndex = 3;          // green
    else if (layer === "PV_TABLES") colorIndex = 5;    // blue
    else if (layer === "PV_OBSTACLES") colorIndex = 1; // red

    let s = "";

    // POLYLINE entity (header)
    s += "0\nPOLYLINE\n";
    s += "5\n" + polyHandle + "\n";
    s += "8\n" + layer + "\n";
    s += "62\n" + colorIndex + "\n"; // color by entity
    s += "66\n1\n";                  // vertices follow
    s += "70\n1\n";                  // closed polyline
    s += "10\n0.0\n20\n0.0\n30\n0.0\n";

    // VERTEX entities
    closed.forEach((pt, i) => {
      s += "0\nVERTEX\n";
      s += "5\n" + vertexHandles[i] + "\n";
      s += "8\n" + layer + "\n";
      s += "10\n" + pt.x + "\n";
      s += "20\n" + pt.y + "\n";
      s += "30\n0.0\n";
    });

    // SEQEND
    s += "0\nSEQEND\n";
    s += "5\n" + seqHandle + "\n";
    s += "8\n" + layer + "\n";

    return s;
  }

  // Emit a TEXT entity at (x,y) with given content on given layer + optional color
  function textToDXF(x, y, text, layer, colorIndex) {
    const h = 2.5; // text height
    const handle = nextHandle();
    let s = "";
    s += "0\nTEXT\n";
    s += "5\n" + handle + "\n";
    s += "8\n" + layer + "\n";
    if (typeof colorIndex === "number") {
      s += "62\n" + colorIndex + "\n";
    }
    s += "10\n" + x + "\n";
    s += "20\n" + y + "\n";
    s += "30\n0.0\n";
    s += "40\n" + h + "\n";
    s += "1\n" + text + "\n";
    s += "50\n0.0\n";
    return s;
  }

  // Emit a LINE entity
  function lineToDXF(x1, y1, x2, y2, layer, colorIndex) {
    const handle = nextHandle();
    let s = "";
    s += "0\nLINE\n";
    s += "5\n" + handle + "\n";
    s += "8\n" + layer + "\n";
    if (typeof colorIndex === "number") {
      s += "62\n" + colorIndex + "\n";
    }
    s += "10\n" + x1 + "\n";
    s += "20\n" + y1 + "\n";
    s += "30\n0.0\n";
    s += "11\n" + x2 + "\n";
    s += "21\n" + y2 + "\n";
    s += "31\n0.0\n";
    return s;
  }

  // ========= COLLECT ALL GEOMETRY FIRST =========
  // We'll also tag each item with a "kind" so we know which are PV_AREAS
  const exportItems = []; // { layer, pathLL, kind, ptsXY }

  // Areas + obstacles from overlay registry
  const overlays = Array.from(window.__overlayRegistry || []);
  overlays.forEach(ov => {
    try {
      if (!ov.getMap || !ov.getMap()) return;

      let layer = null;
      let kind = null;

      if (ov.__cat === "area") {
        layer = "PV_AREAS";
        kind = "area";
      } else if (ov.__cat === "constraint" || ov.__cat === "object") {
        layer = "PV_OBSTACLES";
        kind = "obstacle";
      }

      if (!layer) return;

      const pathLL = getPathForExport(ov);
      if (!pathLL || !pathLL.length) return;

      exportItems.push({ layer, pathLL, kind });
    } catch (err) {
      console.warn("DXF overlay error:", err);
    }
  });

  // PV tables from layoutOverlays
  (window.layoutOverlays || []).forEach(poly => {
    try {
      if (!poly.getMap || !poly.getMap()) return;

      const pathLL = getPathForExport(poly) || [];
      if (!pathLL.length) return;

      exportItems.push({ layer: "PV_TABLES", pathLL, kind: "table" });
    } catch (err) {
      console.warn("DXF table error:", err);
    }
  });

  if (!exportItems.length) {
    console.warn("DXF export: no geometry found.");
    return;
  }

  // ========= GLOBAL ORIGIN (one projection for everything) =========
  const allPoints = [];
  exportItems.forEach(item => {
    item.pathLL.forEach(p => allPoints.push(p));
  });

  const origin = centroid(allPoints);
  const proj = makeLocal(origin);

  // Convert all to XY and track global extents
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;

  exportItems.forEach(item => {
    const ptsXY = item.pathLL.map(p => latLngToXY(p, proj));
    item.ptsXY = ptsXY;
    ptsXY.forEach(pt => {
      if (pt.x < minX) minX = pt.x;
      if (pt.x > maxX) maxX = pt.x;
      if (pt.y < minY) minY = pt.y;
      if (pt.y > maxY) maxY = pt.y;
    });
  });

  // ========= BUILD ENTITIES STRING =========
  let entities = "";

  // For table + numbering: store all PV_AREAS vertices as rows
  const coordRows = []; // { idx, lat, lng }
  let globalPointIndex = 1;

  exportItems.forEach(item => {
    const layer = item.layer;
    let ptsXY = item.ptsXY;
    let pathLL = item.pathLL;

    if (!ptsXY || !ptsXY.length) return;

    // For numbering / table on PV_AREAS, avoid duplicate last point if closed
    let uniqueLL = pathLL;
    let uniqueXY = ptsXY;
    if (pathLL.length > 1) {
      const f = pathLL[0];
      const l = pathLL[pathLL.length - 1];
      if (f.lat() === l.lat() && f.lng() === l.lng()) {
        uniqueLL = pathLL.slice(0, -1);
        uniqueXY = ptsXY.slice(0, -1);
      }
    }

    // Add polygon entity
    entities += polyToDXF(uniqueXY, layer);

    // If PV_AREAS: number each vertex and record coordinates
    if (layer === "PV_AREAS") {
      for (let i = 0; i < uniqueLL.length; i++) {
        const ll = uniqueLL[i];
        const pt = uniqueXY[i];

        const idx = globalPointIndex++;
        coordRows.push({
          idx,
          lat: ll.lat(),
          lng: ll.lng()
        });

        // Number mark near the vertex (small offset so it doesn't hide the corner)
        const markX = pt.x + 2;
        const markY = pt.y + 2;
        // colorIndex 5 = blue
        entities += textToDXF(markX, markY, String(idx), "PV_AREA_POINT_ID", 5);
      }
    }
  });

  // ========= COORDINATE TABLE AT BOTTOM-RIGHT =========
  if (coordRows.length > 0) {
    const marginX = 20;
    const marginY = 10;
    const rowStep = 4; // vertical spacing

    // bottom-right of drawing: (maxX, minY)
    const x0 = maxX + marginX;
    const firstRowY = minY - marginY; // header baseline

    // Header (blue)
    entities += textToDXF(x0, firstRowY, "Pt   Lat,Lng", "PV_AREA_COORDS", 5);

    // Rows (blue)
    coordRows.forEach((row, i) => {
      const y = firstRowY - rowStep * (i + 1);
      const text = row.idx + "   " +
        row.lat.toFixed(6) + "," + row.lng.toFixed(6);
      entities += textToDXF(x0, y, text, "PV_AREA_COORDS", 5);
    });

    // Simple frame rectangle around the table (blue)
    const totalRows = coordRows.length + 1; // header + rows
    const topY = firstRowY + 2;
    const bottomY = firstRowY - rowStep * totalRows - 2;
    const x1 = x0 + 80; // table width

    entities += lineToDXF(x0, topY, x1, topY, "PV_AREA_TABLE", 5);
    entities += lineToDXF(x1, topY, x1, bottomY, "PV_AREA_TABLE", 5);
    entities += lineToDXF(x1, bottomY, x0, bottomY, "PV_AREA_TABLE", 5);
    entities += lineToDXF(x0, bottomY, x0, topY, "PV_AREA_TABLE", 5);

    // Horizontal line under header
    const headerLineY = firstRowY - rowStep * 0.5;
    entities += lineToDXF(x0, headerLineY, x1, headerLineY, "PV_AREA_TABLE", 5);
  }

  // ========= DXF HEADER (R12 minimal) =========
  const header = `0
SECTION
2
HEADER
9
$ACADVER
1
AC1009
0
ENDSEC
0
SECTION
2
ENTITIES
`;

  // ========= FOOTER =========
  const footer = `0
ENDSEC
0
EOF
`;

  // ========= SAVE FILE =========
  const dxf = header + entities + footer;
  const blob = new Blob([dxf], { type: "application/dxf" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "pv_layout.dxf";
  a.click();
  URL.revokeObjectURL(url);

  console.log(
    "DXF Export COMPLETE. Handles:",
    handleCounter,
    "items:",
    exportItems.length,
    "points:",
    coordRows.length
  );
}

window.exportToDXF = exportToDXF;









})();
