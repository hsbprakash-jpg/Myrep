/* ================= Interactive charts (Apache ECharts) ==================
   A chart page drawn with Apache ECharts instead of hand-built SVG: a
   tooltip on every mark, a legend that hides and shows series, and
   click-to-drill down a hierarchy. The figures are the builder's own —
   chartData() computes them — so scope, YTD month, units, comparison bases
   and the top-8-plus-Other fold are exactly what Drag & drop charts shows.
   Clicking a category pins it as a page filter (every other page follows)
   and moves the chart to the next level of that hierarchy; Undo steps back.
   Excel and PPT go through the same chartAOA / exportPPT paths, the chart
   rendered as SVG so the deck gets a picture, not a table. */
(function(){
  'use strict';
  const LSKEY='iwpbsg_ichart';
  const IC={groupBy:'micaL2', splitBy:'', l2:'', measure:'', mUser:false, type:'column', drill:[]};
  try{ Object.assign(IC, JSON.parse(localStorage.getItem(LSKEY)||'{}')); }catch(e){}
  IC.drill=[];                       // a drill path belongs to a session, not a stored setting
  function save(){ try{ const o={...IC}; delete o.drill; localStorage.setItem(LSKEY,JSON.stringify(o)); }catch(e){} }
  function $id(x){ return document.getElementById(x); }
  /* the levels of each hierarchy, coarse to fine: a click on a level drills
     to the next one the file carries and the page has not already pinned */
  const HIER=[['business','region','country'],
              ['micaL1','micaL2','micaL3','micaL4','micaL5','mica'],
              ['prodL1','prodL2','prodL3','prodL4','prodL5','prodL6','prodL7','prodCode'],
              ['cgL1','cgL2','segCode'],
              ['funcL1','funcL2','funcCode']];
  const IC_TYPES=[['column','Column'],['stack','Stacked column'],['hbar','Horizontal bar'],['line','Line']];
  let chart=null, last=null;

  /* ---- mount: a nav entry after Drag & drop charts, a page after its view ---- */
  function mount(){
    if($id('ichartView')) return true;
    const nb=$id('navBuilder'), bv=$id('builderView');
    if(!nb||!bv||typeof S==='undefined') return false;
    const st=document.createElement('style');
    st.textContent=`
#icOpts{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;margin-top:10px}
#icOpts label{display:flex;flex-direction:column;gap:3px;font-size:11px;color:var(--text-muted);font-weight:600}
#icOpts select{padding:6px 8px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface-2);
  color:var(--text-primary);font:inherit;font-size:12.5px;min-width:150px}
#icOpts .icbtns{display:flex;gap:8px;margin-left:auto}
#icCrumb[hidden]{display:none}
#icCrumb{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:10px;font-size:12px;color:var(--text-secondary)}
#icCrumb b{color:var(--text-primary)}
#icCrumb button{padding:4px 10px;font-size:11.5px}
#icChart{width:100%;height:460px}
#icChart svg{display:block}`;
    document.head.appendChild(st);
    const btn=document.createElement('button');
    btn.id='navIchart';
    btn.innerHTML='<svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">'+
      '<rect x="1" y="9" width="3.4" height="6" rx="1" fill="currentColor" opacity=".55"/>'+
      '<rect x="6.3" y="5" width="3.4" height="10" rx="1" fill="currentColor"/>'+
      '<rect x="11.6" y="7" width="3.4" height="8" rx="1" fill="currentColor" opacity=".55"/>'+
      '<path d="M1.5 6.5 L6.5 2.5 L10 4.5 L15 1" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>'+
      '<span data-lbl="nav_ichart">Interactive charts</span>';
    nb.insertAdjacentElement('afterend',btn);
    btn.addEventListener('click',()=>setView('ichart'));
    const eyebrow=(bv.querySelector('.eyebrow')||{}).textContent||'';
    const v=document.createElement('div'); v.id='ichartView'; v.hidden=true;
    v.innerHTML=`
    <div class="pagehead"><div>
      <div class="eyebrow">${esc(eyebrow)}</div>
      <h2><span data-lbl="ichart_title">Interactive Charts</span> <span class="unitsfx" data-unitsfx></span></h2>
      <p class="lede2" data-lbl="ichart_lede">Hover a bar for its figures, click a legend entry to hide a series, click a bar to drill into the next level. Same figures as Drag &amp; drop charts.</p>
    </div></div>
    <div class="panel">
      <h4>Chart</h4>
      <p class="subtitle">Charts respect the filters in the left pane · top 8 categories, rest folded into Other</p>
      <div id="icOpts">
        <label>Group by <select id="icGroup" aria-label="Group by"></select></label>
        <label>Split by <select id="icSplit" aria-label="Split by"></select></label>
        <label>KPI scope <select id="icScope" aria-label="KPI scope"></select></label>
        <label>Measure <select id="icMeasure" aria-label="Measure"></select></label>
        <label>Chart <select id="icType" aria-label="Chart type"></select></label>
        <span class="icbtns">
          <button class="btn-quiet" id="icXls">⤓ Excel</button>
          <button class="btn-quiet" id="icPpt">⤓ PPT</button>
        </span>
      </div>
      <div id="icCrumb" hidden>
        <span id="icCrumbTxt"></span>
        <button class="btn-quiet" id="icUndo">↩ Undo drill</button>
        <button class="btn-quiet" id="icClear">✕ Clear drill</button>
      </div>
    </div>
    <div class="panel">
      <div class="panel-head"><div><h4 id="icTitle">Chart</h4><p class="subtitle" id="icSub"></p></div></div>
      <div class="chart-wrap"><div id="icChart" role="img" aria-label="Interactive chart"></div></div>
      <div class="note" id="icNote"></div>
    </div>`;
    bv.insertAdjacentElement('afterend',v);
    wire();
    cfgHook();
    return true;
  }
  function wire(){
    const on=(id,fn)=>$id(id).addEventListener('change',e=>{ fn(e.target.value); save(); renderIchart(); });
    on('icGroup',v=>{ IC.groupBy=v; if(IC.splitBy===v) IC.splitBy=''; });
    on('icSplit',v=>{ IC.splitBy=v; });
    on('icScope',v=>{ IC.l2=v; });
    on('icMeasure',v=>{ IC.measure=v; IC.mUser=true; });
    on('icType',v=>{ IC.type=v; });
    $id('icUndo').addEventListener('click',undo);
    $id('icClear').addEventListener('click',clearDrill);
    $id('icXls').addEventListener('click',()=>{
      if(last) exportWB([{name:'Chart data', aoa:chartAOA(last.cfg)}], brandSlug()+'_interactive_chart');
    });
    $id('icPpt').addEventListener('click',()=>{
      if(last) exportPPT([{title:chartTitle(last.cfg), el:$id('icChart'), aoa:chartAOA(last.cfg), note:filterNote()}],
        brandSlug()+'_interactive_chart');
    });
    window.addEventListener('resize',()=>{ if(chart&&S.view==='ichart') chart.resize(); });
    try{ window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change',()=>{ if(S.view==='ichart') renderIchart(); }); }catch(e){}
  }
  /* the view answers to the config pack's Views sheet like every other page */
  function cfgHook(){
    const nb=$id('navIchart'); if(!nb) return;
    let on=true; try{ on=cfgViewOn('ichart'); }catch(e){}
    nb.style.display=on?'':'none';
    if(!on&&S.view==='ichart') S.view='summary';
    try{ const v=cfgL('nav_ichart'); if(v!=null) nb.querySelector('span').textContent=v; }catch(e){}
  }

  /* ---- the selection, as the builder's cfg ------------------------------- */
  function isDim(k){ return !!k && !PERIOD_FIELDS.has(k) && k!=='metric' && k!=='cmp'; }
  function pinned(k){ const s=S.filters[k]; return !!(s&&s.size===1); }
  function fieldOpts(){
    const m=S.model, out=[];
    try{ if(liveMetrics(filteredRows()).length>1) out.push('metric'); }catch(e){}
    if(cmpChipAvail(m)) out.push('cmp');
    if(m.cyMonths&&m.cyMonths.length) out.push('perM','perQ','perY');
    for(const k of Object.keys(FIELD_LABELS)) if(k in m.dimCols) out.push(k);
    return out;
  }
  function measureNow(opts){
    return (IC.mUser&&opts.includes(IC.measure))? IC.measure : (opts.includes('basis')?'basis':'cy');
  }
  function cfgOf(){
    const split=IC.splitBy? [IC.splitBy] : [];
    /* the data type is 'stack' whenever there is a split: chartData keeps
       negative segments for a stack and drops them for a column, and a
       grouped bar wants the true figure. How it is drawn is IC.type. */
    return {groupBy:[IC.groupBy], splitBy:split, measure:measureNow(measureOpts()), mUser:IC.mUser,
            type:split.length?'stack':'column', l2:IC.l2};
  }
  function fillControls(){
    const m=S.model, fo=fieldOpts();
    if(!fo.includes(IC.groupBy)) IC.groupBy=fo.includes('micaL2')?'micaL2':(fo.find(isDim)||fo[0]||'');
    const lbl=k=>FIELD_LABELS[k]||ALL_DIM_LABELS[k]||k;
    $id('icGroup').innerHTML=fo.map(k=>`<option value="${esc(k)}"${k===IC.groupBy?' selected':''}>${esc(lbl(k))}</option>`).join('');
    const so=fo.filter(k=>k!==IC.groupBy&&!PERIOD_FIELDS.has(k)&&k!=='metric');
    if(IC.splitBy&&!so.includes(IC.splitBy)) IC.splitBy='';
    $id('icSplit').innerHTML=`<option value="">— none —</option>`+
      so.map(k=>`<option value="${esc(k)}"${k===IC.splitBy?' selected':''}>${esc(lbl(k))}</option>`).join('');
    const l2s=[...new Set(m.rows.map(r=>r.d.micaL2).filter(Boolean))].sort();
    let mets=[]; try{ mets=liveMetrics(m.rows); }catch(e){}
    const sc=$id('icScope');
    sc.innerHTML=`<option value="">Total (all KPIs)</option>`+
      (mets.length? `<optgroup label="Metrics">${mets.map(d2=>`<option value="m:${esc(d2.name)}">${esc(d2.name)}</option>`).join('')}</optgroup>`:'')+
      l2s.map(l2=>`<option value="${esc(l2)}">${esc(l2)}</option>`).join('');
    if(![...sc.options].some(o=>o.value===IC.l2)) IC.l2='';
    sc.value=IC.l2;
    const mo=measureOpts(), mea=measureNow(mo);
    $id('icMeasure').innerHTML=mo.map(k=>`<option value="${esc(k)}"${k===mea?' selected':''}>${esc(measureLabel(k))}</option>`).join('');
    if(!IC_TYPES.some(t=>t[0]===IC.type)) IC.type='column';
    $id('icType').innerHTML=IC_TYPES.map(([v,l])=>`<option value="${v}"${v===IC.type?' selected':''}>${l}</option>`).join('');
  }

  /* ---- drill: a click pins the category and steps one level finer -------- */
  function nextLevel(k){
    const m=S.model, fam=HIER.find(h=>h.includes(k)); if(!fam) return '';
    return fam.slice(fam.indexOf(k)+1).find(x=>(x in m.dimCols)&&!pinned(x)&&(x in FIELD_LABELS))||'';
  }
  function onClick(p){
    if(!last||!isDim(IC.groupBy)) return;
    const c=last.d.cats[p.dataIndex]; if(!c||c.name==='Other') return;
    const dim=IC.groupBy;
    const prev=S.filters[dim]? new Set(S.filters[dim]) : null;
    IC.drill.push({dim, val:c.name, prev, prevGroup:dim});
    S.filters[dim]=new Set([c.name]);
    const nx=nextLevel(dim); if(nx) IC.groupBy=nx;
    if(IC.splitBy===IC.groupBy) IC.splitBy='';
    save(); render();
  }
  function popDrill(){
    const d=IC.drill.pop(); if(!d) return false;
    if(d.prev&&d.prev.size) S.filters[d.dim]=d.prev; else delete S.filters[d.dim];
    IC.groupBy=d.prevGroup;
    return true;
  }
  function undo(){ if(popDrill()){ save(); render(); } }
  function clearDrill(){ let n=0; while(popDrill()) n++; if(n){ save(); render(); } }
  /* a step whose filter was changed elsewhere on the page is no longer a drill */
  function pruneDrill(){ IC.drill=IC.drill.filter(d=>{ const s=S.filters[d.dim]; return s&&s.size===1&&s.has(d.val); }); }

  /* ---- draw ---------------------------------------------------------------- */
  function renderIchart(){
    if(!mount()||!S.model) return;
    const view=$id('ichartView'); if(view.hidden) return;   // setView draws it when it opens
    fillControls(); pruneDrill();
    const lbl=k=>FIELD_LABELS[k]||ALL_DIM_LABELS[k]||k;
    $id('icCrumb').hidden=!IC.drill.length;
    $id('icCrumbTxt').innerHTML=IC.drill.length? 'Drilled: '+IC.drill.map(d=>`<b>${esc(lbl(d.dim))}</b> = ${esc(d.val)}`).join(' › ') : '';
    const title=$id('icTitle'), sub=$id('icSub'), note=$id('icNote');
    if(typeof echarts==='undefined'){
      title.textContent='Chart engine not loaded'; sub.textContent='';
      note.textContent='This HTML file may be truncated — re-download the full dashboard file.'; return;
    }
    const cfg=cfgOf();
    let d;
    try{ d=chartData(cfg); }catch(e){ title.textContent='Chart'; sub.textContent=''; note.textContent='Could not compute this chart: '+(e&&e.message||e); return; }
    last={cfg,d};
    const {cats, segNames, paired, multi, clamped, pnlOnly, isPer}=d;
    title.textContent=chartTitle(cfg);
    sub.textContent=filterNote();
    const el=$id('icChart');
    if(!cats.length){ note.textContent='No rows for this selection under the current filters.'; if(chart) chart.clear(); return; }

    const css=k=>getComputedStyle(document.documentElement).getPropertyValue(k).trim();
    const cat=[1,2,3,4,5,6,7].map(i=>css('--cat'+i)), stk=[1,2,3,4,5].map(i=>css('--stk'+i));
    const grey=css('--series-2'), ink=css('--text-primary'), ink2=css('--text-secondary'), muted=css('--text-muted');
    const grid=css('--grid'), surf=css('--surface-1'), border=css('--border');
    const type=IC.type, horiz=type==='hbar', isLine=type==='line';
    const stackable=!!segNames&&!paired&&!multi&&type==='stack';   // bases are scenarios of one number: never stacked
    const names=cats.map(c=>c.name);
    const radius=v=>{ if(v==null||stackable||isLine) return undefined;
      const r=[4,4,0,0], n=[0,0,4,4], hr=[0,4,4,0], hn=[4,0,0,4];
      return horiz? (v<0?hn:hr) : (v<0?n:r); };
    function mk(name,data,color,extra){
      const s={name, type:isLine?'line':'bar', color, emphasis:{focus:'series'},
        data:data.map((v,i)=>{
          const p={value:v};
          const st={};
          if(!isLine){ const br=radius(v); if(br) st.borderRadius=br;
            if(!segNames&&cats[i].name==='Other') st.color=grey; }
          if(Object.keys(st).length) p.itemStyle=st;
          return p; })};
      if(isLine) Object.assign(s,{symbol:'circle',symbolSize:8,lineStyle:{width:2},itemStyle:{borderColor:surf,borderWidth:2},connectNulls:false});
      else Object.assign(s,{barMaxWidth:horiz?18:28, itemStyle:{borderColor:surf,borderWidth:1}});
      return Object.assign(s,extra||{});
    }
    let series;
    if(segNames){
      const cols=paired? segNames.map((sn,i)=>i===0?cat[0]:i===1?grey:cat[i%7])
               : stackable? segNames.map((sn,i)=>stk[sn==='Other'?4:Math.min(i,4)])
               : segNames.map((sn,i)=>sn==='Other'?grey:cat[i%7]);
      series=segNames.map((sn,i)=>mk(sn, cats.map(c=>(c.segs&&c.segs[sn]!=null)?c.segs[sn]:null), cols[i], stackable?{stack:'total'}:{}));
    } else {
      series=[mk(measureLabel(cfg.measure), cats.map(c=>c.val), cat[0], {})];
    }
    const slotW=Math.max(56, Math.floor((el.clientWidth-100)/Math.max(1,names.length))-8);
    const catAxis={type:'category', data:names, axisTick:{show:false}, axisLine:{lineStyle:{color:border}},
      axisLabel:Object.assign({color:ink2, interval:0},
        horiz? {width:150, overflow:'break', lineHeight:13}
        : names.length>8? {rotate:30, width:110, overflow:'truncate'}
        : {width:slotW, overflow:'break', lineHeight:14})};
    if(horiz) catAxis.inverse=true;                  // first category at the top, as a table reads
    const valAxis={type:'value', axisLabel:{color:ink2, formatter:v=>fmtChart(v)}, splitLine:{lineStyle:{color:grid}}, axisLine:{show:false}};
    const opt={
      backgroundColor:'transparent', animationDuration:250,
      textStyle:{fontFamily:'"Segoe UI",system-ui,-apple-system,Helvetica,Arial,sans-serif'},
      aria:{enabled:true},
      tooltip:{trigger:'axis', axisPointer:{type:isLine?'line':'shadow'}, valueFormatter:v=>v==null?'–':fmtFull(v),
        backgroundColor:surf, borderColor:border, textStyle:{color:ink, fontSize:12}},
      legend:{show:series.length>1, bottom:0, type:'scroll', icon:isLine?'circle':'roundRect',
        textStyle:{color:ink2}, pageTextStyle:{color:ink2}, pageIconColor:ink2, pageIconInactiveColor:border},
      toolbox:{right:0, top:0, itemSize:14, iconStyle:{borderColor:muted}, emphasis:{iconStyle:{borderColor:css('--hsbc-red')}},
        feature:{saveAsImage:{name:brandSlug()+'_interactive_chart', title:'Save as image', backgroundColor:surf}}},
      grid:{left:horiz?24:8, right:24, top:30, bottom:series.length>1?36:8, containLabel:true},
      xAxis: horiz? valAxis : catAxis,
      yAxis: horiz? catAxis : valAxis,
      series
    };
    const rowsH=segNames&&!stackable? 14*segNames.length+18 : 34;
    el.style.height=(horiz? Math.max(320, 70+names.length*rowsH) : 460)+'px';
    if(!chart){ chart=echarts.init(el,null,{renderer:'svg'}); chart.on('click',onClick); }
    /* a redraw can land while the pointer is still on a bar (click-to-drill):
       hide the tooltip first and replace the parts rather than the whole option,
       so the tooltip's own DOM survives and no timer finds it gone */
    try{ chart.dispatchAction({type:'hideTip'}); }catch(e){}
    chart.resize(); chart.setOption(opt,{replaceMerge:['series','xAxis','yAxis','legend','grid']}); chart.resize();

    const notes=[];
    if(cats.some(c=>c.name==='Other')) notes.push('top 8 categories by size, the rest folded into Other');
    if(isPer&&IC.mUser&&cfg.measure!=='cy') notes.push('a period axis charts each period’s own actuals and forecast — the measure does not apply');
    if(pnlOnly) notes.push('P&L and balances never total together — this comparison covers the P&L side of the scope');
    if(clamped) notes.push('negative components excluded from stacked segments');
    if(type==='stack'&&!stackable) notes.push(segNames? 'comparison bases are never stacked — drawn side by side' : 'add a Split by to stack');
    if(isDim(IC.groupBy)){
      const nx=nextLevel(IC.groupBy);
      notes.push(nx? `click a ${isLine?'point':'bar'} to drill into ${lbl(nx)}` : `click a ${isLine?'point':'bar'} to filter the page to that ${lbl(IC.groupBy)}`);
    }
    notes.push(`${cats.length} categor${cats.length===1?'y':'ies'}`);
    note.textContent=notes.join(' · ')+'.';
  }

  /* ---- hooks into the page: view switching, render, config -------------- */
  const _setView=setView;
  setView=function(v){
    mount();
    const ic=v==='ichart', vw=$id('ichartView'), nb=$id('navIchart');
    if(vw) vw.hidden=!ic;                 // shown before render() so the chart measures a real box
    if(nb) nb.classList.toggle('active',ic);
    _setView(v);
  };
  const _render=render;
  render=function(){
    if(S.view!=='ichart'){ _render(); return; }
    if(!S.model) return;
    try{ renderTplPane(); }catch(e){}
    renderIchart();
    // the page furniture render() keeps in step on every other view
    try{
      const scTr=scopeTrail();
      document.querySelectorAll('.eyebrow:not(.exec-eyebrow)').forEach(el=>el.textContent=cfgS('eyebrow')+(scTr? ' — '+scTr : ''));
      renderKpiNav(); renderCountryNav();
      $id('navBizperf').hidden=!bizPerfAvail();
      $id('rowCount').textContent=`${filteredRows().length.toLocaleString()} rows in scope`;
    }catch(e){}
  };
  if(typeof applyCfgDom==='function'){
    const _acd=applyCfgDom;
    applyCfgDom=function(){ const r=_acd.apply(this,arguments); cfgHook(); return r; };
  }
  window.__ichartRender=renderIchart;
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mount); else mount();
})();
