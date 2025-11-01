// drawing.js — classic script (no ES module export); exposes window.mountDrawing(ctx)
// ctx: { google, map, document }
window.mountDrawing = function mountDrawing(ctx){
  const { google, map, document } = ctx;

  // ==== Constants (drawing/layout) ====
  const CAPACITY_DENSITY_KWP_PER_SQM = 0.20;  // DC kWp per m²
  const DC_TO_AC_RATIO = 1.15;                // DC/AC ratio for AC estimate
  const DEFAULT_TILT = 10, DEFAULT_AZIMUTH = 180;

  // ==== HtmlLabel overlay ====
  class HtmlLabel extends google.maps.OverlayView{
    constructor(position, text){ super(); this.position = position; this.text = text; this.div = null; }
    onAdd(){ this.div=document.createElement('div'); this.div.className='area-label'; this.div.textContent=this.text; this.getPanes().overlayMouseTarget.appendChild(this.div); }
    draw(){ if(!this.div) return; const proj=this.getProjection(); if(!proj) return; const pt=proj.fromLatLngToDivPixel(this.position); this.div.style.position='absolute'; this.div.style.left=(pt.x - this.div.offsetWidth/2)+'px'; this.div.style.top=(pt.y - this.div.offsetHeight/2)+'px'; }
    onRemove(){ this.div?.parentNode?.removeChild(this.div); this.div=null; }
    setPosition(latlng){ this.position=latlng; this.draw(); }
    setText(t){ this.text=t; if(this.div) this.div.textContent=t; }
  }

  // ==== UI: Search/Drawer toggles ====
  const searchWrap=document.getElementById('search-wrap'), input=document.getElementById('search-input');
  const byTool=(t)=>document.querySelector(`[data-tool="${t}"]`);
  const drawer=document.getElementById('drawer');

  // Popovers (declared now, wired later)
  const popRotate=document.getElementById('pop-rotate');
  const popAlign=document.getElementById('pop-align');
  const popDistrib=document.getElementById('pop-distribute');
  const popGrid=document.getElementById('pop-grid');
  const popSetback=document.getElementById('pop-setback');
  const popRoads=document.getElementById('pop-roads');

  const btnRotate=document.getElementById('act-rotate');
  const btnAlign=document.getElementById('act-align');
  const btnDistrib=document.getElementById('act-distribute');
  const btnGrid=document.getElementById('act-grid');
  const btnSetback=document.getElementById('act-setback');
  const btnRoads=document.getElementById('act-roads');
  const btnBuilding=document.getElementById('act-building');

  ['search','tools','select','draw'].forEach(t=>{
    const el=byTool(t); if(!el) return;
    el.onclick=()=>{
      if(t==='search'){
        drawer.classList.remove('open');
        const open=searchWrap.classList.toggle('open');
        searchWrap.setAttribute('aria-expanded',String(open));
        if(open) setTimeout(()=>{input.focus(); input.select();},0);
        // close popovers
        [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none');
        [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active'));
      }else if(t==='tools'){
        searchWrap.classList.remove('open'); searchWrap.setAttribute('aria-expanded','false');
        drawer.classList.toggle('open');
        if(!drawer.classList.contains('open')){
          [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none');
          [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active'));
        }
      }else{
        document.querySelectorAll('.tool-btn').forEach(b=>b.classList.remove('is-primary'));
        el.classList.add('is-primary');
        if(t==='select') setMode('select');
        if(t==='draw'){ drawer.classList.add('open'); }
      }
    };
  });
  drawer.querySelector('[data-close-drawer]').onclick=()=>{ drawer.classList.remove('open'); [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none'); [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active')); };

  // ==== Category & shape buttons ====
  let category='area';
  const seg=document.getElementById('seg-category');
  seg.querySelectorAll('button').forEach(btn=>{
    btn.onclick=()=>{ seg.querySelectorAll('button').forEach(b=>b.classList.remove('active')); btn.classList.add('active'); category=btn.dataset.cat; updateDrawingOptions(); };
  });

  const shapeBtns={
    polygon:document.getElementById('shape-poly'),
    rectangle:document.getElementById('shape-rect'),
    circle:document.getElementById('shape-circle')
  };
  let drawKind=null; // 'polygon'|'rectangle'|'circle'|'road-free' (polyline)
  Object.entries(shapeBtns).forEach(([k,btn])=>{
    btn.onclick=()=>{
      if(drawKind===k){ btn.classList.remove('active'); drawKind=null; dm.setDrawingMode(null); setMode('select'); return; }
      Object.values(shapeBtns).forEach(b=>b.classList.remove('active'));
      btn.classList.add('active'); drawKind=k; setMode('draw'); startDrawing();
    };
  });

  // ==== Mode buttons ====
  const btnMove=document.getElementById('act-move');
  const btnMarq=document.getElementById('act-marquee');
  const btnMulti=document.getElementById('act-multisel');
  const btnPaste=document.getElementById('act-paste');
  let multiLatch=false, pasteArmed=false;

  document.getElementById('act-select').onclick=()=>{ setMode('select'); btnMove.classList.remove('active'); btnMarq.classList.remove('active'); };
  btnMove.onclick=()=>{ if(mode==='move'){ setMode('select'); btnMove.classList.remove('active'); } else { setMode('move'); btnMove.classList.add('active'); btnMarq.classList.remove('active'); } };
  btnMarq.onclick=()=>{ if(mode==='marquee'){ setMode('select'); btnMarq.classList.remove('active'); } else { setMode('marquee'); btnMarq.classList.add('active'); btnMove.classList.remove('active'); } };
  btnMulti.onclick=()=>{ multiLatch=!multiLatch; btnMulti.classList.toggle('active', multiLatch); };

  document.getElementById('act-copy').onclick=handleCopy;
  btnPaste.onclick=()=>{
    if(pasteArmed){ pasteArmed=false; btnPaste.classList.remove('active'); map.setOptions({draggable:true}); map.set('draggableCursor',null); setMode('select'); return; }
    if(!copyBuffer?.items?.length) return;
    pasteArmed=true; btnPaste.classList.add('active'); setMode('paste'); map.setOptions({draggable:false}); map.set('draggableCursor','copy');
  };
  document.getElementById('act-delete').onclick=handleDelete;

  // ==== Popovers placement ====
  const appEl = document.getElementById('app');
  function placePopoverNextToDrawer(pop, btn){
    if(!drawer.classList.contains('open')) return;
    if(pop.parentElement !== appEl) appEl.appendChild(pop);
    const drect = drawer.getBoundingClientRect();
    const arect = appEl.getBoundingClientRect();
    const left = (drect.right - arect.left) + 8;
    const top  = (drect.top - arect.top);
    Object.assign(pop.style,{left:left+'px', top:top+'px'});
  }
  function togglePop(pop, btn){
    const on = pop.style.display!=='none';
    [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none');
    [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active'));
    if(!on){
      placePopoverNextToDrawer(pop, btn);
      pop.style.display='block';
      btn.classList.add('active');
    }
  }
  btnRotate.onclick=()=>togglePop(popRotate, btnRotate);
  btnAlign.onclick=()=>togglePop(popAlign, btnAlign);
  btnDistrib.onclick=()=>togglePop(popDistrib, btnDistrib);
  btnGrid.onclick=()=>togglePop(popGrid, btnGrid);
  btnSetback.onclick=()=>togglePop(popSetback, btnSetback);
  btnRoads.onclick=()=>togglePop(popRoads, btnRoads);
  btnBuilding.onclick=()=>{
    // convenience: switch to Objects + draw rectangle
    seg.querySelectorAll('button').forEach(b=>b.classList.remove('active'));
    const obBtn = [...seg.querySelectorAll('button')].find(b=>b.dataset.cat==='object'); if(obBtn){ obBtn.classList.add('active'); category='object'; }
    Object.values(shapeBtns).forEach(b=>b.classList.remove('active'));
    shapeBtns.rectangle.classList.add('active');
    drawKind='rectangle'; setMode('draw'); startDrawing();
  };
  document.addEventListener('click',(e)=>{
    const inside = e.target.closest('.popover') || e.target.closest('#act-rotate,#act-align,#act-distribute,#act-grid,#act-setback,#act-roads');
    if(!inside){
      [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none');
      [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active'));
    }
  });

  // ==== State ====
  const areas=[]; const constraints=[]; const objects=[];
  let areaCounter=0, constraintCounter=0, objectCounter=0, roadsCounter=0;
  let mode='select'; let selected=[];
  let suppressMapClickOnce=false; /* prevent map click from clearing marquee selection */
  let focusIndex=-1;
  let boundarySetbackMeters = 0;

  // ==== Drawing manager ====
  const dm=new google.maps.drawing.DrawingManager({
    drawingMode:null, drawingControl:false,
    polygonOptions: styleFor('area'),
    rectangleOptions: styleFor('area'),
    circleOptions: styleFor('area'),
    polylineOptions: { strokeColor:'#1f2937', strokeOpacity:0.85, strokeWeight:3, clickable:true, draggable:false, editable:false }
  });
  dm.setMap(map);

  function styleFor(cat){
    if(cat==='area')       return { fillColor:'#3a7afe',  fillOpacity:0.15, strokeColor:'#3a7afe',  strokeWeight:2, clickable:true, editable:false, draggable:false };
    if(cat==='constraint') return { fillColor:'#fbd1d1',  fillOpacity:0.12, strokeColor:'#d83b3b',  strokeWeight:2, clickable:true, editable:false, draggable:false };
    // objects style
    return                 { fillColor:'#cfd4dd',  fillOpacity:0.18, strokeColor:'#4b5563',  strokeWeight:2, clickable:true, editable:false, draggable:false };
  }
  function selectedStyleOf(cat){
    if(cat==='area')       return { fillColor:'#ffcf66', fillOpacity:0.18, strokeColor:'#ffb703', strokeWeight:3 };
    if(cat==='constraint') return { fillColor:'#ffd1d1', fillOpacity:0.18, strokeColor:'#ff7b7b', strokeWeight:3 };
    return                 { fillColor:'#e6eaf1', fillOpacity:0.22, strokeColor:'#64748b', strokeWeight:3 };
  }
  function applySelectedStyle(ov,on){
    if(on){ ov.__savedStyle = { fillColor:ov.get('fillColor'), fillOpacity:ov.get('fillOpacity'), strokeColor:ov.get('strokeColor'), strokeWeight:ov.get('strokeWeight') }; ov.setOptions(selectedStyleOf(ov.__cat)); }
    else if(ov.__savedStyle){ ov.setOptions(ov.__savedStyle); ov.__savedStyle=null; }
  }
  function updateDrawingOptions(){
    dm.setOptions({ polygonOptions:styleFor(category), rectangleOptions:styleFor(category), circleOptions:styleFor(category) });
    if(drawKind) startDrawing();
  }
  function startDrawing(){
    if(!drawKind){ dm.setDrawingMode(null); return; }
    const kind =
      drawKind==='polygon'   ? google.maps.drawing.OverlayType.POLYGON   :
      drawKind==='rectangle' ? google.maps.drawing.OverlayType.RECTANGLE :
      drawKind==='circle'    ? google.maps.drawing.OverlayType.CIRCLE    :
      drawKind==='road-free' ? google.maps.drawing.OverlayType.POLYLINE  :
      null;
    dm.setDrawingMode(kind);
  }
  function stopDrawing(){ dm.setDrawingMode(null); }

  // ==== overlaycomplete handler ====
  google.maps.event.addListener(dm,'overlaycomplete',e=>{
    const ov=e.overlay, type=e.type;

    // Special case: free-draw road (polyline -> corridor polygon)
    if(type===google.maps.drawing.OverlayType.POLYLINE && drawKind==='road-free'){
      const width=Number(document.getElementById('road-width').value||'4');
      const setback=Number(document.getElementById('road-setback').value||'0');
      const path=e.overlay.getPath().getArray();
      createRoadCorridorFromPath(path, width + 2*setback, 'Road ' + (++roadsCounter));
      e.overlay.setMap(null);
      drawKind=null; stopDrawing(); setMode('select');
      return;
    }

    // Normal overlays
    ov.__type=type; ov.__cat=category; ov.__selected=false;
    if(category==='area'){
      const name=`Area ${++areaCounter}`; ov.__name=name;
      ov.__tilt = DEFAULT_TILT; ov.__azimuth = DEFAULT_AZIMUTH;
      const label=new HtmlLabel(getLabelNW(ov), name); label.setMap(map); ov.__label=label;
      areas.push({overlay:ov,type,name,label});
    }else if(category==='constraint'){
      const name=`Constraint ${++constraintCounter}`; ov.__name=name; ov.__label=null;
      constraints.push({overlay:ov,type,name,label:null});
    }else{
      const name=`Object ${++objectCounter}`; ov.__name=name; ov.__label=null;
      objects.push({overlay:ov,type,name,label:null});
    }
    bindEvents(ov);
    if(category==='area') updateSummary();
    if(drawKind && drawKind!=='road-free'){ startDrawing(); } else { stopDrawing(); }
    clickSelect(ov, false);
  });

  function bindEvents(ov){
    const refresh=()=>{
      if(ov.__label) ov.__label.setPosition(getLabelNW(ov));
      if(ov.__cat==='area') updateSummary();
    };
    ov.addListener('click',(ev)=>clickSelect(ov, multiLatch || ev?.domEvent?.ctrlKey || ev?.domEvent?.metaKey || ev?.domEvent?.shiftKey));
    ov.addListener('mousedown',(ev)=>{
      if(mode==='move'){
        if(!ov.__selected && !multiLatch) clearSelection();
        if(!ov.__selected) toggleSelection(ov,true);
        applyModeFlags();
      }
    });
    if(ov.getPath){
      ov.getPath().addListener('insert_at', refresh);
      ov.getPath().addListener('remove_at', refresh);
      ov.getPath().addListener('set_at', refresh);
      ov.addListener('mouseup', refresh);
    }else if(ov.getBounds){ ov.addListener('bounds_changed', refresh); }
    if(ov.getRadius){ ov.addListener('center_changed', refresh); ov.addListener('radius_changed', refresh); }
    ov.addListener('dragend', refresh);

    // group drag
    ov.addListener('dragstart', ()=>onDragStart(ov));
    ov.addListener('drag', ()=>onDragMove(ov));
    ov.addListener('dragend', ()=>onDragEnd(ov));
  }

  // ==== Selection helpers ====
  function clickSelect(ov, additive){
    if(additive){ toggleSelection(ov, false); }
    else{ clearSelection(); toggleSelection(ov, true); }
    focusIndex = selected.indexOf(ov);
    applyModeFlags();
  }
  function toggleSelection(ov, forceOn=false){
    const i=selected.indexOf(ov);
    if(i>=0 && !forceOn){
      selected.splice(i,1); ov.__selected=false; ov.setEditable(false); ov.setDraggable(false); applySelectedStyle(ov,false);
    }else if(i<0){
      selected.push(ov); ov.__selected=true; ov.setEditable(mode==='select'); ov.setDraggable(mode==='move'); applySelectedStyle(ov,true);
    }
    updateFocusAfterSelection();
    showGroupHandle();
  }
  function clearSelection(){
    selected.forEach(o=>{o.__selected=false; o.setEditable(false); o.setDraggable(false); applySelectedStyle(o,false);});
    selected=[]; focusIndex=-1; showGroupHandle();
  }
  function updateFocusAfterSelection(){
    if(selected.length===0){ focusIndex=-1; return; }
    if(focusIndex<0) focusIndex=0;
    if(focusIndex>=selected.length) focusIndex=selected.length-1;
  }
  function deselectFocused(){
    if(focusIndex>=0 && focusIndex<selected.length){
      const ov = selected[focusIndex];
      toggleSelection(ov,false);
      applyModeFlags();
    }
  }
  function applyModeFlags(){ selected.forEach(o=>{ o.setEditable(mode==='select'); o.setDraggable(mode==='move'); }); showGroupHandle(); }

  map.addListener('click',(e)=>{
    if(suppressMapClickOnce){ suppressMapClickOnce=false; return; }
    if(mode==='paste' && pasteArmed && copyBuffer?.items?.length){ placeGroupAt(e.latLng); return; }
    if(!multiLatch) clearSelection();
  });

  function setMode(m){
    mode=m;
    if(m!=='marquee') removeMarquee();
    if(m!=='paste'){ pasteArmed=false; btnPaste.classList.remove('active'); map.setOptions({draggable:true}); map.set('draggableCursor',null); }
    applyModeFlags();
    btnMarq.classList.toggle('active', m==='marquee');
    btnMove.classList.toggle('active', m==='move');
    if(m==='draw'){ startDrawing(); }
  }

  // ==== Group handle ====
  let groupHandle=null;
  function showGroupHandle(){
    if(groupHandle){ groupHandle.setMap(null); groupHandle=null; }
    if(selected.length<2 || mode!=='move') return;
    const center=getGroupCenter(selected);
    groupHandle=new google.maps.Marker({position:center,map,draggable:true,icon:{path:google.maps.SymbolPath.CIRCLE,scale:0},label:{text:'✥',color:'#1e2433',fontSize:'16px',fontWeight:'700'}});
    let start=null;
    groupHandle.addListener('dragstart',()=>{ start=groupHandle.getPosition(); });
    groupHandle.addListener('dragend',()=>{
      const end=groupHandle.getPosition();
      const d=google.maps.geometry.spherical.computeDistanceBetween(start,end);
      const heading=google.maps.geometry.spherical.computeHeading(start,end);
      selected.forEach(ov=>offsetOverlay(ov,d,heading));
      selected.forEach(o=>{ if(o.__label) o.__label.setPosition(getLabelNW(o)); });
      updateSummary();
    });
  }
  function getGroupCenter(arr){ let lat=0,lng=0; arr.forEach(o=>{const c=getCenter(o); lat+=c.lat(); lng+=c.lng();}); return new google.maps.LatLng(lat/arr.length,lng/arr.length); }

  // ==== Copy / Paste / Delete ====
  let copyBuffer=null;
  function handleCopy(){
    if(!selected.length) return;
    const centers = selected.map(s=>getCenter(s));
    const gc = new google.maps.LatLng(
      centers.reduce((a,c)=>a+c.lat(),0)/centers.length,
      centers.reduce((a,c)=>a+c.lng(),0)/centers.length
    );
    copyBuffer = {
      groupCenter: {lat:gc.lat(), lng:gc.lng()},
      items: selected.map(s=>serializeWithCenter(s))
    };
  }
  function serializeWithCenter(ov){
    const c=getCenter(ov);
    if(ov.__type==='polygon'){
      return {type:'polygon', cat:ov.__cat, center:{lat:c.lat(),lng:c.lng()}, path: ov.getPath().getArray().map(ll=>({lat:ll.lat(),lng:ll.lng()}))};
    }
    if(ov.__type==='rectangle'){
      const b=ov.getBounds(), ne=b.getNorthEast(), sw=b.getSouthWest();
      return {type:'rectangle', cat:ov.__cat, center:{lat:c.lat(),lng:c.lng()}, ne:{lat:ne.lat(),lng:ne.lng()}, sw:{lat:sw.lat(),lng:sw.lng()}};
    }
    if(ov.__type==='circle'){
      return {type:'circle', cat:ov.__cat, center:{lat:c.lat(),lng:c.lng()}, radius:ov.getRadius()};
    }
  }
  function placeGroupAt(targetLatLng){
    const GCsrc = new google.maps.LatLng(copyBuffer.groupCenter.lat, copyBuffer.groupCenter.lng);
    const GCdst = targetLatLng;
    copyBuffer.items.forEach(item=>{
      const Csrc=new google.maps.LatLng(item.center.lat, item.center.lng);
      const d_gc = google.maps.geometry.spherical.computeDistanceBetween(GCsrc, Csrc);
      const h_gc = google.maps.geometry.spherical.computeHeading(GCsrc, Csrc);
      const Cdst = google.maps.geometry.spherical.computeOffset(GCdst, d_gc, h_gc);
      const ov = rebuildShapeAt(item, Cdst);
      ov.__cat=item.cat; ov.__type=item.type;
      if(ov.__cat==='area'){
        const nm=`Area ${++areaCounter}`; ov.__name=nm;
        ov.__tilt = DEFAULT_TILT; ov.__azimuth = DEFAULT_AZIMUTH;
        const label=new HtmlLabel(getLabelNW(ov),nm); label.setMap(map); ov.__label=label;
        areas.push({overlay:ov,type:ov.__type,name:nm,label}); updateSummary();
      }else if(ov.__cat==='constraint'){
        ov.__name=`Constraint ${++constraintCounter}`; ov.__label=null;
        constraints.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
      }else{
        ov.__name=`Object ${++objectCounter}`; ov.__label=null;
        objects.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
      }
      bindEvents(ov);
    });
  }
  function duplicateSelectionAt(targetLatLng){
    if(!selected.length) return;
    const centers = selected.map(s=>getCenter(s));
    const gc = new google.maps.LatLng(
      centers.reduce((a,c)=>a+c.lat(),0)/centers.length,
      centers.reduce((a,c)=>a+c.lng(),0)/centers.length
    );
    const GCsrc = gc;
    const GCdst = targetLatLng;
    const items = selected.map(s=>serializeWithCenter(s));
    clearSelection();
    items.forEach(item=>{
      const Csrc=new google.maps.LatLng(item.center.lat, item.center.lng);
      const d_gc = google.maps.geometry.spherical.computeDistanceBetween(GCsrc, Csrc);
      const h_gc = google.maps.geometry.spherical.computeHeading(GCsrc, Csrc);
      const Cdst = google.maps.geometry.spherical.computeOffset(GCdst, d_gc, h_gc);
      const ov = rebuildShapeAt(item, Cdst);
      ov.__cat=item.cat; ov.__type=item.type;
      if(ov.__cat==='area'){
        const nm=`Area ${++areaCounter}`; ov.__name=nm;
        ov.__tilt = item.tilt ?? DEFAULT_TILT; ov.__azimuth = item.azimuth ?? DEFAULT_AZIMUTH;
        const label=new HtmlLabel(getLabelNW(ov),nm); label.setMap(map); ov.__label=label;
        areas.push({overlay:ov,type:ov.__type,name:nm,label}); updateSummary();
      }else if(ov.__cat==='constraint'){
        ov.__name=`Constraint ${++constraintCounter}`; ov.__label=null;
        constraints.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
      }else{
        ov.__name=`Object ${++objectCounter}`; ov.__label=null;
        objects.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
      }
      bindEvents(ov);
      toggleSelection(ov,true);
    });
    applyModeFlags();
    updateSummary();
  }
  function rebuildShapeAt(item, newCenter){
    if(item.type==='polygon'){
      const Csrc=new google.maps.LatLng(item.center.lat, item.center.lng);
      const newPath=item.path.map(p=>{
        const pt=new google.maps.LatLng(p.lat,p.lng);
        const d=google.maps.geometry.spherical.computeDistanceBetween(Csrc, pt);
        const h=google.maps.geometry.spherical.computeHeading(Csrc, pt);
        return google.maps.geometry.spherical.computeOffset(newCenter, d, h);
      });
      return new google.maps.Polygon(Object.assign(styleFor(item.cat), {map, paths:newPath}));
    }
    if(item.type==='rectangle'){
      const Csrc=new google.maps.LatLng(item.center.lat, item.center.lng);
      const neSrc=new google.maps.LatLng(item.ne.lat, item.ne.lng);
      const swSrc=new google.maps.LatLng(item.sw.lat, item.sw.lng);
      const dNE=google.maps.geometry.spherical.computeDistanceBetween(Csrc, neSrc);
      const hNE=google.maps.geometry.spherical.computeHeading(Csrc, neSrc);
      const dSW=google.maps.geometry.spherical.computeDistanceBetween(Csrc, swSrc);
      const hSW=google.maps.geometry.spherical.computeHeading(Csrc, swSrc);
      const ne=google.maps.geometry.spherical.computeOffset(newCenter, dNE, hNE);
      const sw=google.maps.geometry.spherical.computeOffset(newCenter, dSW, hSW);
      return new google.maps.Rectangle(Object.assign(styleFor(item.cat), {map, bounds:new google.maps.LatLngBounds(sw, ne)}));
    }
    if(item.type==='circle'){
      return new google.maps.Circle(Object.assign(styleFor(item.cat), {map, center:newCenter, radius:item.radius}));
    }
  }
  function handleDelete(){
    if(!selected.length) return;
    selected.forEach(ov=>{
      const arr=(ov.__cat==='area')?areas:(ov.__cat==='constraint'?constraints:objects);
      const idx=arr.findIndex(x=>x.overlay===ov);
      if(idx>=0){ if(arr[idx].label) arr[idx].label.setMap(null); arr.splice(idx,1); }
      ov.setMap(null);
    });
    selected=[]; focusIndex=-1; updateSummary(); showGroupHandle();
  }

  // ==== Summary ====
  function fmt(n,d=1){ return Number(n).toLocaleString('en-US',{maximumFractionDigits:d}); }
  function updateSummary(){
    const tbody=document.getElementById('sum-body'); tbody.innerHTML='';
    let totM2=0, totDC=0, totAC=0;
    if(!areas.length){
      tbody.innerHTML=`<tr><td colspan="6" class="muted" style="padding:12px 10px">No areas yet</td></tr>`;
    }else{
      areas.forEach(({overlay})=>{
        const name=overlay.__name||'Area', tilt=overlay.__tilt??DEFAULT_TILT, az=overlay.__azimuth??DEFAULT_AZIMUTH;
        const m2=computeArea(overlay,overlay.__type); const dc=m2*CAPACITY_DENSITY_KWP_PER_SQM; const ac=dc/DC_TO_AC_RATIO;
        totM2+=m2; totDC+=dc; totAC+=ac;
        const tr=document.createElement('tr');
        tr.innerHTML=`<td class="name">${name}</td><td class="num">${fmt(dc)}</td><td class="num">${fmt(ac)}</td><td class="num">${fmt(m2)}</td><td class="num">${fmt(tilt,0)}°</td><td class="num">${fmt(az,0)}°</td>`;
        tbody.appendChild(tr);
      });
    }
    document.getElementById('tot-dc').textContent=fmt(totDC);
    document.getElementById('tot-ac').textContent=fmt(totAC);
    document.getElementById('tot-m2').textContent=fmt(totM2);
  }
  function computeArea(ov,type){
    if(type==='polygon') return google.maps.geometry.spherical.computeArea(ov.getPath());
    if(type==='rectangle'){
      const b=ov.getBounds(), ne=b.getNorthEast(), sw=b.getSouthWest();
      const nw=new google.maps.LatLng(ne.lat(),sw.lng()), se=new google.maps.LatLng(sw.lat(),ne.lng());
      return google.maps.geometry.spherical.computeArea([ne,se,sw,nw]);
    }
    if(type==='circle'){ const r=ov.getRadius(); return Math.PI*r*r; }
    return 0;
  }

  // Sync header/footer scroll with body
  (function(){
    const head=document.getElementById('sum-head');
    const body=document.getElementById('sum-bodywrap');
    const foot=document.getElementById('sum-foot');
    body.addEventListener('scroll', ()=>{ head.scrollLeft=body.scrollLeft; foot.scrollLeft=body.scrollLeft; });
  })();

  // ==== Marquee selection ====
  let marqueeDiv=null, marqueeActive=false, marqueeStart=null, wasDraggable=true, marqueeAdditive=false;
  const projOverlay=new google.maps.OverlayView(); projOverlay.onAdd=function(){}; projOverlay.draw=function(){}; projOverlay.onRemove=function(){}; projOverlay.setMap(map);
  document.getElementById('map').addEventListener('mousedown',(ev)=>{
    if(mode!=='marquee') return;
    marqueeActive=true;
    const rect=map.getDiv().getBoundingClientRect();
    marqueeStart={x:ev.clientX-rect.left, y:ev.clientY-rect.top};
    marqueeAdditive = multiLatch || ev.ctrlKey || ev.metaKey || ev.shiftKey;
    createMarquee(marqueeStart,marqueeStart);
    wasDraggable = map.get('draggable')!==false;
    map.setOptions({draggable:false});
    const onMove=(e)=>{ if(!marqueeActive) return; const r=map.getDiv().getBoundingClientRect(); const p2={x:e.clientX-r.left, y:e.clientY-r.top}; createMarquee(marqueeStart,p2); };
    const onUp=()=>{ marqueeActive=false; document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp); finalizeMarquee(); map.setOptions({draggable:wasDraggable}); };
    document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp);
  });
  function createMarquee(p1,p2){
    if(!marqueeDiv){ marqueeDiv=document.createElement('div'); marqueeDiv.className='marquee'; map.getDiv().appendChild(marqueeDiv); }
    const x=Math.min(p1.x,p2.x), y=Math.min(p1.y,p2.y), w=Math.abs(p1.x-p2.x), h=Math.abs(p1.y-p2.y);
    Object.assign(marqueeDiv.style,{left:x+'px',top:y+'px',width:w+'px',height:h+'px'});
  }
  function finalizeMarquee(){
    if(!marqueeDiv) return;
    const proj=projOverlay.getProjection();
    if(proj){
      const left=parseFloat(marqueeDiv.style.left), top=parseFloat(marqueeDiv.style.top), width=parseFloat(marqueeDiv.style.width), height=parseFloat(marqueeDiv.style.height);
      const x1 = left, y1 = top, x2 = left + width, y2 = top + height;
      const sw=proj.fromContainerPixelToLatLng(new google.maps.Point(Math.min(x1,x2), Math.max(y1,y2)));
      const ne=proj.fromContainerPixelToLatLng(new google.maps.Point(Math.max(x1,x2), Math.min(y1,y2)));
      const box=new google.maps.LatLngBounds(sw, ne);
      if(!marqueeAdditive) clearSelection();
      [...areas, ...constraints, ...objects].forEach(({overlay})=>{
        const b=getBounds(overlay);
        if(b && box.intersects(b)){
          toggleSelection(overlay, marqueeAdditive ? false : true);
        }
      });
      applyModeFlags();
      focusIndex = selected.length ? selected.length-1 : -1;
      suppressMapClickOnce=true;
    }
    removeMarquee();
  }
  function removeMarquee(){ marqueeDiv?.remove(); marqueeDiv=null; }
  function getBounds(ov){
    if(ov.getBounds) return ov.getBounds();
    if(ov.getPath){ const b=new google.maps.LatLngBounds(); ov.getPath().forEach(p=>b.extend(p)); return b; }
    if(ov.getCenter && ov.getRadius){
      const c=ov.getCenter(), r=ov.getRadius();
      const ne=google.maps.geometry.spherical.computeOffset(c, r*Math.SQRT1_2, 45);
      const sw=google.maps.geometry.spherical.computeOffset(c, r*Math.SQRT1_2, 225);
      return new google.maps.LatLngBounds(sw, ne);
    }
    return null;
  }

  // ==== Keyboard shortcuts ====
  document.addEventListener('keydown',(e)=>{
    if(e.key==='Escape'){
      if(drawKind){ Object.values(shapeBtns).forEach(b=>b.classList.remove('active')); drawKind=null; dm.setDrawingMode(null); }
      pasteArmed=false; btnPaste.classList.remove('active'); map.setOptions({draggable:true}); map.set('draggableCursor',null);
      if(mode==='marquee'){ setMode('select'); btnMarq.classList.remove('active'); }
      [popRotate,popAlign,popDistrib,popGrid,popSetback,popRoads].forEach(p=>p.style.display='none'); [btnRotate,btnAlign,btnDistrib,btnGrid,btnSetback,btnRoads].forEach(b=>b.classList.remove('active'));
      return;
    }

    // ignore when typing
    const t = e.target;
    const typing = t && (t.tagName==='INPUT' || t.tagName==='TEXTAREA' || t.isContentEditable);
    if(typing) return;

    const ctrlMeta = e.ctrlKey || e.metaKey;

    // tool switches
    if(e.key.toLowerCase()==='s'){ setMode('select'); }
    else if(e.key.toLowerCase()==='v'){ setMode('move'); btnMove.classList.add('active'); btnMarq.classList.remove('active'); }
    else if(e.key.toLowerCase()==='m'){ setMode('marquee'); btnMarq.classList.add('active'); btnMove.classList.remove('active'); }
    else if(e.key.toLowerCase()==='r'){ btnRotate.click(); }
    else if(e.key.toLowerCase()==='l'){ btnAlign.click(); }
    else if(e.key.toLowerCase()==='d'){ btnDistrib.click(); }

    // select all
    if(ctrlMeta && e.key.toLowerCase()==='a'){
      e.preventDefault();
      clearSelection();
      [...areas, ...constraints, ...objects].forEach(({overlay})=>toggleSelection(overlay,true));
      applyModeFlags();
      focusIndex = selected.length ? 0 : -1;
    }

    // Copy / Paste
    if(ctrlMeta && e.key.toLowerCase()==='c'){ e.preventDefault(); handleCopy(); }
    if(ctrlMeta && e.key.toLowerCase()==='v'){ e.preventDefault(); btnPaste.click(); }

    // Duplicate (Ctrl/Cmd + D)
    if(ctrlMeta && e.key.toLowerCase()==='d'){
      e.preventDefault();
      if(selected.length){
        const gc=getGroupCenter(selected);
        const east=google.maps.geometry.spherical.computeOffset(gc, 5, 90);
        const target=google.maps.geometry.spherical.computeOffset(east, 5, 180);
        duplicateSelectionAt(target);
      }
    }

    // Delete
    if(e.key==='Delete' || e.key==='Backspace'){ e.preventDefault(); handleDelete(); }

    // Nudge
    if(e.key.startsWith('Arrow')){
      e.preventDefault();
      const step = e.shiftKey ? 10 : 1; // meters
      const heading = (e.key==='ArrowUp')?0:(e.key==='ArrowRight')?90:(e.key==='ArrowDown')?180:270;
      selected.forEach(ov=>{
        offsetOverlay(ov, step, heading);
        if(ov.__label) ov.__label.setPosition(getLabelNW(ov));
      });
      showGroupHandle();
      updateSummary();
    }

    // Cycle focus
    if(e.key===']' && selected.length){ focusIndex = (focusIndex+1) % selected.length; }
    if(e.key==='[' && selected.length){ focusIndex = (focusIndex-1+selected.length) % selected.length; }

    // Deselect focused item (Ctrl/Cmd + Backspace)
    if(ctrlMeta && (e.key==='Backspace')){ e.preventDefault(); deselectFocused(); }
  });

  // ==== Geometry helpers ====
  function offsetOverlay(ov, meters, heading){
    if(ov.getPath){
      const path = ov.getPath();
      for(let i=0;i<path.getLength();i++){
        path.setAt(i, google.maps.geometry.spherical.computeOffset(path.getAt(i), meters, heading));
      }
    }else if(ov.getBounds){
      const b=ov.getBounds();
      const ne = google.maps.geometry.spherical.computeOffset(b.getNorthEast(), meters, heading);
      const sw = google.maps.geometry.spherical.computeOffset(b.getSouthWest(), meters, heading);
      ov.setBounds(new google.maps.LatLngBounds(sw, ne));
    }else if(ov.getCenter){
      ov.setCenter( google.maps.geometry.spherical.computeOffset(ov.getCenter(), meters, heading) );
    }
  }
  function getCenter(ov){
    if(ov.getCenter) return ov.getCenter();
    if(ov.getBounds) return ov.getBounds().getCenter();
    if(ov.getPath){
      let lat=0,lng=0; const n=ov.getPath().getLength();
      for(let i=0;i<n;i++){ const p=ov.getPath().getAt(i); lat+=p.lat(); lng+=p.lng(); }
      return new google.maps.LatLng(lat/n, lng/n);
    }
  }

  // ==== Group drag ====
  let groupDrag=null; // {anchor, startCenter, snaps: Map(overlay -> snapshot)}
  function snapshotForGroup(ov){
    if(ov.__type==='polygon'){
      return {type:'polygon', path: ov.getPath().getArray().map(ll=>({lat:ll.lat(),lng:ll.lng()}))};
    }
    if(ov.__type==='rectangle'){
      const b=ov.getBounds(), ne=b.getNorthEast(), sw=b.getSouthWest();
      return {type:'rectangle', ne:{lat:ne.lat(),lng:ne.lng()}, sw:{lat:sw.lat(),lng:sw.lng()}};
    }
    if(ov.__type==='circle'){
      const c=ov.getCenter();
      return {type:'circle', center:{lat:c.lat(),lng:c.lng()}, radius: ov.getRadius()};
    }
  }
  function applySnapshotOffset(ov, snap, meters, heading){
    if(snap.type==='polygon'){
      const newPath = snap.path.map(p=>google.maps.geometry.spherical.computeOffset(new google.maps.LatLng(p.lat,p.lng), meters, heading));
      ov.setPaths(newPath);
    }else if(snap.type==='rectangle'){
      const ne = google.maps.geometry.spherical.computeOffset(new google.maps.LatLng(snap.ne.lat, snap.ne.lng), meters, heading);
      const sw = google.maps.geometry.spherical.computeOffset(new google.maps.LatLng(snap.sw.lat, snap.sw.lng), meters, heading);
      ov.setBounds(new google.maps.LatLngBounds(sw, ne));
    }else if(snap.type==='circle'){
      const c = google.maps.geometry.spherical.computeOffset(new google.maps.LatLng(snap.center.lat, snap.center.lng), meters, heading);
      ov.setCenter(c);
    }
  }
  function onDragStart(ov){
    if(mode!=='move' || !ov.__selected || selected.length<2) return;
    groupDrag = { anchor: ov, startCenter: getCenter(ov), snaps: new Map() };
    selected.forEach(o=>{ if(o!==ov) groupDrag.snaps.set(o, snapshotForGroup(o)); });
  }
  function onDragMove(ov){
    if(!groupDrag || groupDrag.anchor!==ov) return;
    const cur = getCenter(ov);
    const d = google.maps.geometry.spherical.computeDistanceBetween(groupDrag.startCenter, cur);
    const h = google.maps.geometry.spherical.computeHeading(groupDrag.startCenter, cur);
    groupDrag.snaps.forEach((snap, other)=>{
      applySnapshotOffset(other, snap, d, h);
      if(other.__label) other.__label.setPosition(getLabelNW(other));
    });
    if(groupHandle) groupHandle.setPosition(getGroupCenter(selected));
  }
  function onDragEnd(ov){
    if(!groupDrag || groupDrag.anchor!==ov) return;
    groupDrag=null;
    updateSummary();
    showGroupHandle();
  }

  // ==== Rotate / Align / Distribute / Grid ====
  const slider=document.getElementById('rotate-range'), sliderDeg=document.getElementById('rotate-deg');
  let rotateSession=null;
  slider.addEventListener('input', ()=>{
    if(!selected.length){ sliderDeg.textContent='0°'; slider.value=0; return; }
    if(!rotateSession){
      rotateSession = { center:getGroupCenter(selected), start:0, snapshot:selected.map(s=>serializeWithCenter(s)) };
    }
    const angle = Number(slider.value); sliderDeg.textContent = angle+'°';
    const delta = angle - rotateSession.start;
    applyRotation(delta, rotateSession.center, rotateSession.snapshot);
  });
  slider.addEventListener('change', ()=>{ if(rotateSession){ rotateSession.start = Number(slider.value); }});
  document.getElementById('act-rotate').addEventListener('click', ()=>{ rotateSession=null; slider.value=0; sliderDeg.textContent='0°'; });

  function applyRotation(deltaDeg, centerLatLng, snapshot){
    selected.forEach((ov, i)=>{
      const snap = snapshot[i];
      replaceOverlayWithRotation(ov, snap, centerLatLng, deltaDeg);
      if(ov.__label) ov.__label.setPosition(getLabelNW(ov));
    });
    updateSummary();
  }
  function rotatePointAround(point, center, deltaDeg){
    const d = google.maps.geometry.spherical.computeDistanceBetween(center, point);
    const h = google.maps.geometry.spherical.computeHeading(center, point);
    return google.maps.geometry.spherical.computeOffset(center, d, h + deltaDeg);
  }
  function replaceOverlayWithRotation(ov, snap, center, deltaDeg){
    if(snap.type==='polygon'){
      const rotated = snap.path.map(p=>rotatePointAround(new google.maps.LatLng(p.lat, p.lng), center, deltaDeg));
      ov.setPaths(rotated);
    }else if(snap.type==='rectangle'){
      const ne = new google.maps.LatLng(snap.ne.lat, snap.ne.lng);
      const sw = new google.maps.LatLng(snap.sw.lat, snap.sw.lng);
      const nw = new google.maps.LatLng(ne.lat(), sw.lng());
      const se = new google.maps.LatLng(sw.lat(), ne.lng());
      const corners = [ne,se,sw,nw];
      const rotated = corners.map(pt=>rotatePointAround(pt, center, deltaDeg));
      if(ov instanceof google.maps.Rectangle){
        const poly = new google.maps.Polygon(Object.assign(styleFor(ov.__cat), {map, paths:rotated}));
        poly.__type='polygon'; poly.__cat=ov.__cat; poly.__name=ov.__name; poly.__label=ov.__label; poly.__tilt=ov.__tilt; poly.__azimuth=ov.__azimuth;
        if(poly.__label) poly.__label.setMap(map);
        const idx = selected.indexOf(ov);
        if(idx>=0) selected[idx]=poly;
        ov.setMap(null);
        bindEvents(poly); applySelectedStyle(poly,true);
      }else{
        ov.setPaths(rotated);
      }
    }else if(snap.type==='circle'){
      // rotate circle center around group center
      const cOld = new google.maps.LatLng(snap.center.lat, snap.center.lng);
      const cNew = rotatePointAround(cOld, center, deltaDeg);
      ov.setCenter(cNew);
    }
  }

  popAlign.querySelectorAll('button').forEach(b=>{
    b.addEventListener('click', ()=>{
      if(selected.length<2) return;
      const ref = selected[0]; const rb = getBounds(ref);
      const refTop = rb.getNorthEast().lat(), refBottom = rb.getSouthWest().lat();
      const refRight = rb.getNorthEast().lng(), refLeft = rb.getSouthWest().lng();
      selected.slice(1).forEach(ov=>{
        const bds = getBounds(ov); if(!bds) return;
        const top = bds.getNorthEast().lat(), bottom = bds.getSouthWest().lat();
        const right = bds.getNorthEast().lng(), left = bds.getSouthWest().lng();
        let meters=0, heading=0;
        if(b.dataset.align==='top'){
          const p1=new google.maps.LatLng(top, (left+right)/2);
          const p2=new google.maps.LatLng(refTop, (left+right)/2);
          meters = google.maps.geometry.spherical.computeDistanceBetween(p1,p2);
          heading = google.maps.geometry.spherical.computeHeading(p1,p2);
        }
        if(b.dataset.align==='bottom'){
          const p1=new google.maps.LatLng(bottom, (left+right)/2);
          const p2=new google.maps.LatLng(refBottom, (left+right)/2);
          meters = google.maps.geometry.spherical.computeDistanceBetween(p1,p2);
          heading = google.maps.geometry.spherical.computeHeading(p1,p2);
        }
        if(b.dataset.align==='left'){
          const p1=new google.maps.LatLng((top+bottom)/2, left);
          const p2=new google.maps.LatLng((top+bottom)/2, refLeft);
          meters = google.maps.geometry.spherical.computeDistanceBetween(p1,p2);
          heading = google.maps.geometry.spherical.computeHeading(p1,p2);
        }
        if(b.dataset.align==='right'){
          const p1=new google.maps.LatLng((top+bottom)/2, right);
          const p2=new google.maps.LatLng((top+bottom)/2, refRight);
          meters = google.maps.geometry.spherical.computeDistanceBetween(p1,p2);
          heading = google.maps.geometry.spherical.computeHeading(p1,p2);
        }
        offsetOverlay(ov, meters, heading);
        if(ov.__label) ov.__label.setPosition(getLabelNW(ov));
      });
      updateSummary();
      popAlign.style.display='none'; btnAlign.classList.remove('active');
    });
  });

  popDistrib.querySelectorAll('button').forEach(b=>{
    b.addEventListener('click', ()=>{
      if(selected.length<3) return;
      const arr=[...selected];
      const axis=b.dataset.distribute;
      arr.sort((a,b)=>{
        const ca=getCenter(a), cb=getCenter(b);
        return axis==='horizontal' ? ca.lng()-cb.lng() : ca.lat()-cb.lat();
      });
      const first=arr[0], last=arr[arr.length-1];
      const c1=getCenter(first), cN=getCenter(last);
      const total=google.maps.geometry.spherical.computeDistanceBetween(c1,cN);
      const heading=google.maps.geometry.spherical.computeHeading(c1,cN);
      for(let i=1;i<arr.length-1;i++){
        const target=google.maps.geometry.spherical.computeOffset(c1, total*(i/(arr.length-1)), heading);
        const cur=getCenter(arr[i]);
        const d=google.maps.geometry.spherical.computeDistanceBetween(cur,target);
        const h=google.maps.geometry.spherical.computeHeading(cur,target);
        offsetOverlay(arr[i], d, h);
        if(arr[i].__label) arr[i].__label.setPosition(getLabelNW(arr[i]));
      }
      updateSummary();
      popDistrib.style.display='none'; btnDistrib.classList.remove('active');
    });
  });

  document.getElementById('grid-apply').addEventListener('click', ()=>{
    if(!selected.length) return;
    const cols=parseInt(document.getElementById('grid-cols').value||'1',10);
    const rows=parseInt(document.getElementById('grid-rows').value||'1',10);
    const sx=parseFloat(document.getElementById('grid-sx').value||'0');
    const sy=parseFloat(document.getElementById('grid-sy').value||'0');
    const gc=getGroupCenter(selected);
    const buffer = selected.map(s=>serializeWithCenter(s));
    for(let r=0;r<rows;r++){
      for(let c=0;c<cols;c++){
        if(r===0 && c===0) continue;
        const dx = c*sx, dy = r*sy;
        const east = google.maps.geometry.spherical.computeOffset(gc, dx, 90);
        const target = google.maps.geometry.spherical.computeOffset(east, dy, 180);
        buffer.forEach(item=>{
          const ov = rebuildShapeAt(item, targetShift(item.center, gc, target));
          ov.__cat=item.cat; ov.__type=item.type;
          if(ov.__cat==='area'){
            const nm=`Area ${++areaCounter}`; ov.__name=nm;
            ov.__tilt = item.tilt ?? DEFAULT_TILT; ov.__azimuth = item.azimuth ?? DEFAULT_AZIMUTH;
            const label=new HtmlLabel(getLabelNW(ov),nm); label.setMap(map); ov.__label=label;
            areas.push({overlay:ov,type:ov.__type,name:nm,label}); updateSummary();
          }else if(ov.__cat==='constraint'){
            ov.__name=`Constraint ${++constraintCounter}`; ov.__label=null; constraints.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
          }else{
            ov.__name=`Object ${++objectCounter}`; ov.__label=null; objects.push({overlay:ov,type:ov.__type,name:ov.__name,label:null});
          }
          bindEvents(ov);
        });
      }
    }
    popGrid.style.display='none'; btnGrid.classList.remove('active');
  });
  function targetShift(itemCenter, groupCenterSrc, gridTargetCenter){
    const Csrc=new google.maps.LatLng(itemCenter.lat, itemCenter.lng);
    const Gsrc=new google.maps.LatLng(groupCenterSrc.lat(), groupCenterSrc.lng());
    const d=google.maps.geometry.spherical.computeDistanceBetween(Gsrc, Csrc);
    const h=google.maps.geometry.spherical.computeHeading(Gsrc, Csrc);
    return google.maps.geometry.spherical.computeOffset(gridTargetCenter, d, h);
  }

  // ==== Objects & Setbacks ====
  // 1) Boundary setback (store only)
  document.getElementById('in-setback-boundary').addEventListener('change', (e)=>{
    const v = Number(e.target.value||'0');
    boundarySetbackMeters = Math.max(0, v);
  });

  // 2) Roads
  document.getElementById('road-free').addEventListener('click', ()=>{
    drawKind='road-free'; setMode('draw'); startDrawing();
    popRoads.style.display='none'; btnRoads.classList.remove('active');
  });
  document.getElementById('road-perim').addEventListener('click', ()=>{
    const width=Number(document.getElementById('road-width').value||'4');
    const setback=Number(document.getElementById('road-setback').value||'0');
    const W = width + 2*setback;
    const areaOv = selected.find(o=>o.__cat==='area') || (areas.length? areas[areas.length-1].overlay : null);
    if(!areaOv){ alert('Draw or select an Area first.'); return; }
    const boundaryPath = getPerimeterPath(areaOv);
    if(!boundaryPath || boundaryPath.length<2){ alert('Area boundary not available.'); return; }
    const d = width/2 + setback;
    const innerPath = offsetClosedPathInward(boundaryPath, d);
    createRoadCorridorFromPath(innerPath, W, 'Perimeter Road ' + (++roadsCounter));
    popRoads.style.display='none'; btnRoads.classList.remove('active');
  });
  document.getElementById('road-h-apply').addEventListener('click', ()=>{
    const n=parseInt(document.getElementById('road-h-n').value||'1',10);
    const width=Number(document.getElementById('road-width').value||'4');
    const setback=Number(document.getElementById('road-setback').value||'0');
    if(n<=0) return;
    const b = getAreasUnionBounds();
    if(!b){ alert('Draw or select an Area first.'); return; }
    const latMin=b.getSouthWest().lat(), latMax=b.getNorthEast().lat();
    const lngMin=b.getSouthWest().lng(), lngMax=b.getNorthEast().lng();
    for(let i=1;i<=n;i++){
      const lat = latMin + (i/(n+1))*(latMax-latMin);
      const path=[ new google.maps.LatLng(lat, lngMin), new google.maps.LatLng(lat, lngMax) ];
      createRoadCorridorFromPath(path, width + 2*setback, 'Road H' + i);
    }
    popRoads.style.display='none'; btnRoads.classList.remove('active');
  });
  document.getElementById('road-v-apply').addEventListener('click', ()=>{
    const n=parseInt(document.getElementById('road-v-n').value||'1',10);
    const width=Number(document.getElementById('road-width').value||'4');
    const setback=Number(document.getElementById('road-setback').value||'0');
    if(n<=0) return;
    const b = getAreasUnionBounds();
    if(!b){ alert('Draw or select an Area first.'); return; }
    const latMin=b.getSouthWest().lat(), latMax=b.getNorthEast().lat();
    const lngMin=b.getSouthWest().lng(), lngMax=b.getNorthEast().lng();
    for(let i=1;i<=n;i++){
      const lng = lngMin + (i/(n+1))*(lngMax-lngMin);
      const path=[ new google.maps.LatLng(latMin, lng), new google.maps.LatLng(latMax, lng) ];
      createRoadCorridorFromPath(path, width + 2*setback, 'Road V' + i);
    }
    popRoads.style.display='none'; btnRoads.classList.remove('active');
  });

  function getAreasUnionBounds(){
    const list = selected.filter(o=>o.__cat==='area');
    const bucket = list.length ? list : areas.map(a=>a.overlay);
    if(!bucket.length) return null;
    const b=new google.maps.LatLngBounds();
    bucket.forEach(ov=>{ const bb=getBounds(ov); if(bb) b.union(bb); });
    return b;
  }

  function getPerimeterPath(ov){
    if(ov.__type===google.maps.drawing.OverlayType.POLYGON || ov.getPath){
      const arr = ov.getPath().getArray();
      return (arr.length && (arr[0].equals(arr[arr.length-1]))) ? arr.slice(0,-1) : arr.slice();
    }
    if(ov.__type===google.maps.drawing.OverlayType.RECTANGLE || ov.getBounds){
      const b=ov.getBounds(), ne=b.getNorthEast(), sw=b.getSouthWest();
      const nw=new google.maps.LatLng(ne.lat(), sw.lng());
      const se=new google.maps.LatLng(sw.lat(), ne.lng());
      return [nw, ne, se, sw, nw];
    }
    if(ov.__type===google.maps.drawing.OverlayType.CIRCLE || ov.getCenter){
      const c=ov.getCenter(), r=ov.getRadius();
      const pts=[]; const N=64;
      for(let i=0;i<N;i++){ pts.push( google.maps.geometry.spherical.computeOffset(c, r, i*(360/N)) ); }
      pts.push(pts[0]);
      return pts;
    }
    return null;
  }

  // Build a corridor polygon around a (poly)line path
  function createRoadCorridorFromPath(path, widthMeters, name){
    if(!path || path.length<2) return;
    const polyPoints = corridorFromPath(path, widthMeters);
    const poly = new google.maps.Polygon(Object.assign(styleFor('object'), {map, paths:polyPoints}));
    poly.__type='polygon'; poly.__cat='object'; poly.__name=name||('Road ' + (++roadsCounter)); poly.__label=null;
    objects.push({overlay:poly,type:'polygon',name:poly.__name,label:null});
    bindEvents(poly);
    clearSelection(); toggleSelection(poly,true); applyModeFlags();
  }
  // Corridor from path: offsets left/right by width/2 (approximation with averaged heading per vertex)
  function corridorFromPath(points, widthMeters){
    const half=widthMeters/2;
    const left=[], right=[];
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
  function headingAtIndex(points, i){
    const n=points.length;
    if(i===0) return google.maps.geometry.spherical.computeHeading(points[0], points[1]);
    if(i===n-1) return google.maps.geometry.spherical.computeHeading(points[n-2], points[n-1]);
    const h1=google.maps.geometry.spherical.computeHeading(points[i-1], points[i]);
    const h2=google.maps.geometry.spherical.computeHeading(points[i], points[i+1]);
    return averageHeading(h1,h2);
  }
  function averageHeading(h1,h2){
    const r1=h1*Math.PI/180, r2=h2*Math.PI/180;
    const x=Math.cos(r1)+Math.cos(r2), y=Math.sin(r1)+Math.sin(r2);
    return Math.atan2(y,x)*180/Math.PI;
  }

  // Offset a closed path inward (approximate) then return open polyline for corridor centerline
  function offsetClosedPathInward(path, dist){
    const closed = (path.length>1 && path[0].equals(path[path.length-1])) ? path.slice(0,-1) : path.slice();
    const ccw = signedAreaApprox(closed) > 0; // ccw => inward = left normal
    const out=[]; const L=closed.length;
    for(let i=0;i<L;i++){
      const prev=closed[(i-1+L)%L], curr=closed[i], next=closed[(i+1)%L];
      const hPrev=google.maps.geometry.spherical.computeHeading(prev,curr);
      const hNext=google.maps.geometry.spherical.computeHeading(curr,next);
      const hAvg=averageHeading(hPrev,hNext);
      const normal = ccw ? (hAvg - 90) : (hAvg + 90);
      out.push( google.maps.geometry.spherical.computeOffset(curr, dist, normal) );
    }
    return out;
  }
  function signedAreaApprox(path){
    // planar approximation sufficient for orientation detection
    let a=0; const n=path.length;
    for(let i=0;i<n;i++){
      const p=path[i], q=path[(i+1)%n];
      a += (q.lng()-p.lng()) * (q.lat()+p.lat());
    }
    return -a; // ccw positive
  }

  // ==== Helpers used in multiple places ====
  function getLabelNW(ov){
    let bounds=null;
    if(ov.getBounds) bounds=ov.getBounds();
    else if(ov.getPath){
      bounds=new google.maps.LatLngBounds();
      ov.getPath().forEach(p=>bounds.extend(p));
    }else if(ov.getCenter && ov.getRadius){
      const c=ov.getCenter(), r=ov.getRadius();
      const ne=google.maps.geometry.spherical.computeOffset(c, r*Math.SQRT1_2, 45);
      const sw=google.maps.geometry.spherical.computeOffset(c, r*Math.SQRT1_2, 225);
      bounds=new google.maps.LatLngBounds(sw, ne);
    }
    if(!bounds) return map.getCenter();
    const ne=bounds.getNorthEast(), sw=bounds.getSouthWest();
    const nw=new google.maps.LatLng(ne.lat(), sw.lng());
    return google.maps.geometry.spherical.computeOffset(nw, 4, 315);
  }

  // Keep popovers stuck next to drawer while window resizes
  window.addEventListener('resize', ()=>{
    [ [popRotate,btnRotate], [popAlign,btnAlign], [popDistrib,btnDistrib], [popGrid,btnGrid], [popSetback,btnSetback], [popRoads,btnRoads] ].forEach(([p,b])=>{
      if(p.style.display!=='none') placePopoverNextToDrawer(p,b);
    });
  });
};
