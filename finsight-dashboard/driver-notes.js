/* ================= Driver notes from minutes ============================
   Drivers written in a meeting, not in the config pack: Zoom minutes, a
   transcript (.vtt, or the text export), or a Word write-up (.docx) are read
   in the browser, cut into sentences, and every sentence that names a line
   the file carries is matched the way the config's own free commentary is —
   comMatchList decides which line it is about. Minutes are noisy, so nothing
   is stitched until it has been reviewed: a table shows each sentence, the
   line it was matched to (changeable), how it matched, and the terms that
   were ambiguous; in a transcript only sentences that carry a driver cue
   (driven by, offset, timing, one-off …) are ticked to begin with. Accepted
   notes join comInLoaded(), so the House commentary, the Financial Summary,
   All commentaries and the exports weave them in exactly as they weave in the
   config's paragraphs — scope-following, digesting and the per-block cap
   included — and each clause can carry its source: "(minutes, 3 Sep)".
   Word files are unzipped natively (DecompressionStream); PDF is not read
   here — save it as Word or text. */
(function(){
  'use strict';
  const LSKEY='iwpbsg_minutes';
  function $id(x){ return document.getElementById(x); }
  const MN={docs:[], prefer:false, mark:true};
  try{ Object.assign(MN, JSON.parse(localStorage.getItem(LSKEY)||'{}')); }catch(e){}
  if(!Array.isArray(MN.docs)) MN.docs=[];
  function save(){ try{ localStorage.setItem(LSKEY, JSON.stringify(MN)); }catch(e){} }

  /* ---- the accepted notes, as commentary items ---------------------------- */
  const SRC=new Map();                      // sentence -> source marker
  function accepted(){
    const out=[]; SRC.clear();
    for(const d of MN.docs) for(const it of (d.items||[])){
      if(!it.on) continue;
      out.push({view:'', line:it.line||'', text:it.text, src:d.src});
      SRC.set(it.text, d.src);
    }
    return out;
  }
  function counts(d){ const n=(d.items||[]).length, on=(d.items||[]).filter(i=>i.on).length; return {n,on}; }

  /* ---- reading the file ---------------------------------------------------- */
  const unesc=s=>String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(+n)).replace(/&amp;/g,'&');
  const U8=new TextDecoder('utf-8');
  const u16=(b,i)=>b[i]|(b[i+1]<<8), u32=(b,i)=>(b[i]|(b[i+1]<<8)|(b[i+2]<<16))+(b[i+3]*16777216);
  async function inflateRaw(bytes){
    if(typeof DecompressionStream!=='function')
      throw new Error('This browser cannot unzip a Word file — upload the minutes as .txt instead.');
    const ab=await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
    return new Uint8Array(ab);
  }
  /* one entry out of a zip: walk the central directory, then the local header */
  async function zipEntry(b, want){
    let eocd=-1;
    for(let i=b.length-22;i>=Math.max(0,b.length-70000);i--){ if(b[i]===0x50&&b[i+1]===0x4b&&b[i+2]===0x05&&b[i+3]===0x06){ eocd=i; break; } }
    if(eocd<0) throw new Error('not a zip file');
    const n=u16(b,eocd+10); let p=u32(b,eocd+16);
    for(let k=0;k<n;k++){
      if(u32(b,p)!==0x02014b50) break;
      const method=u16(b,p+10), csize=u32(b,p+20), nl=u16(b,p+28), el=u16(b,p+30), cl=u16(b,p+32), lho=u32(b,p+42);
      const name=U8.decode(b.subarray(p+46,p+46+nl));
      if(name===want){
        const lnl=u16(b,lho+26), lel=u16(b,lho+28), start=lho+30+lnl+lel;
        const data=b.subarray(start,start+csize);
        return U8.decode(method===0? data : await inflateRaw(data));
      }
      p+=46+nl+el+cl;
    }
    return null;
  }
  async function docxText(file){
    const b=new Uint8Array(await file.arrayBuffer());
    const xml=await zipEntry(b,'word/document.xml');
    if(!xml) throw new Error('No document text found inside '+file.name);
    const paras=[]; const re=/<w:p[\s>][\s\S]*?<\/w:p>/g; let m;
    while((m=re.exec(xml))){
      const p=m[0].replace(/<w:tab\/>/g,' ').replace(/<w:br\/>/g,' ');
      const t=(p.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g)||[]).map(x=>x.replace(/<[^>]+>/g,'')).join('');
      const s=unesc(t).replace(/\s+/g,' ').trim(); if(s) paras.push(s);
    }
    return paras.join('\n');
  }
  /* cues become speaker turns, one paragraph each */
  function turnsToText(cues){
    const paras=[]; let spk=null, buf=[];
    for(const c of cues){ if(c.spk!==spk&&buf.length){ paras.push(buf.join(' ')); buf=[]; } spk=c.spk; buf.push(c.txt); }
    if(buf.length) paras.push(buf.join(' '));
    return paras.join('\n');
  }
  const SPK=/^([A-Z][\w .'’-]{1,40}?):\s+(.*)$/;
  function vttText(t){
    const cues=[]; let cur=null;
    for(const raw of t.split(/\r?\n/)){
      const ln=raw.trim();
      if(!ln){ if(cur){ cues.push(cur); cur=null; } continue; }
      if(/^WEBVTT/.test(ln)||/^NOTE\b/.test(ln)||/^\d+$/.test(ln)) continue;
      if(/-->/.test(ln)){ if(cur) cues.push(cur); cur={spk:'',txt:''}; continue; }
      if(!cur) cur={spk:'',txt:''};
      let s=ln.replace(/<v\s+([^>]+)>/i,(_,nm)=>{ cur.spk=nm.trim(); return ''; }).replace(/<\/v>/ig,'').replace(/<[^>]+>/g,'').trim();
      const m=s.match(SPK); if(m&&!cur.txt){ cur.spk=m[1]; s=m[2]; }
      cur.txt+=(cur.txt?' ':'')+s;
    }
    if(cur) cues.push(cur);
    return turnsToText(cues);
  }
  function looksTranscript(t){
    const ls=t.split(/\r?\n/).filter(l=>l.trim()); if(ls.length<4) return false;
    const n=ls.filter(l=>/-->/.test(l)||SPK.test(l.replace(/^\[?\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?\]?\s*/,''))).length;
    return n/ls.length>0.4;
  }
  function transcriptText(t){
    const cues=[];
    for(const raw of t.split(/\r?\n/)){
      const ln=raw.trim(); if(!ln||/-->/.test(ln)||/^\d+$/.test(ln)||/^WEBVTT/.test(ln)) continue;
      const m=ln.replace(/^\[?\d{1,2}:\d{2}(?::\d{2})?(?:\.\d+)?\]?\s*/,'').match(SPK);
      cues.push(m? {spk:m[1],txt:m[2]} : {spk:cues.length?cues[cues.length-1].spk:'',txt:ln});
    }
    return turnsToText(cues);
  }
  async function readFile(file){
    const ext=(file.name.split('.').pop()||'').toLowerCase();
    if(ext==='docx') return {text:await docxText(file), kind:'doc'};
    if(ext==='vtt') return {text:vttText(await file.text()), kind:'transcript'};
    if(ext==='pdf') throw new Error('PDF is not read here — save the minutes as Word (.docx) or text (.txt) and upload that.');
    if(ext==='doc') throw new Error('Old .doc files are not read — save as .docx and upload that.');
    const t=await file.text();
    return looksTranscript(t)? {text:transcriptText(t), kind:'transcript'} : {text:t, kind:'doc'};
  }

  /* ---- sentences, and which line each is about --------------------------- */
  const CUE=/\b(?:driven|drives?|driver|due to|because|owing|reflect(?:s|ing|ed)?|offset|partly|partially|higher|lower|up|down|ahead|behind|above|below|increas(?:e|ed|es|ing)|decreas(?:e|ed|es|ing)|grew|growth|declin(?:e|ed|es|ing)|fell|rose|one[- ]off|timing|favou?rable|adverse|impact(?:ed|s)?|mix|margin|volume|pricing|rates?|spread|attrition|run[- ]rate|seasonal|delay(?:ed|s)?|slippage|write[- ]?(?:back|off)|provision|release|uplift|shortfall|overrun|saving|savings)\b/i;
  const FILLER=/\b(?:um+|uh+|erm|you know|i mean|sort of|kind of|basically|actually|okay so|so yeah)\b[,]?\s*/gi;
  const SKIP=/^(?:action(?:s| items?)?|ai|next steps?|attendees?|present|apologies|agenda|minutes|date|time|recording|chair|note taker|circulated)\b/i;
  /* sentences, each knowing which paragraph (speaker turn) it came from */
  function sentences(text){
    const out=[], seen=new Set();
    const paras=String(text).split(/\n+/).map(x=>x.trim()).filter(Boolean);
    // a document's title is not a driver: the first line, if it does not end
    // like a sentence, is a heading — "IWPB Singapore — performance review"
    // would otherwise pin itself to the business line it names
    if(paras.length>1&&!/[.!?;:]$/.test(paras[0])&&paras[0].split(/\s+/).length<=14) paras.shift();
    paras.forEach((para,pi)=>{
      let p=para.replace(FILLER,'').replace(/\s+/g,' ').trim();
      if(!p||SKIP.test(p)) return;
      p=p.replace(/^[-•*–—\d.)]+\s*/,'');
      for(const s of p.split(/(?<=[.!?;])\s+(?=[A-Z(“"'\d])/)){
        let t=s.trim().replace(/^[,;:\-–—\s]+/,'');
        if(t.split(/\s+/).length<6) continue;                    // a fragment
        if(t.length>420) t=t.slice(0,417).replace(/\s+\S*$/,'')+'…';
        if(!/[.!?]$/.test(t)) t+='.';
        const k=t.toLowerCase(); if(seen.has(k)) continue; seen.add(k);
        out.push({t, pi});
      }
    });
    return out;
  }
  /* a sentence that opens on a pronoun continues the one before it, in the
     same paragraph: "NII is up. It's driven by the deposit margin" is about
     NII, whatever else the second sentence happens to name */
  const FOLLOWS=/^(?:it|it's|its|that|that's|this|these|those|they|which|so it|and it)\b/i;
  function analyse(text, kind){
    const rows=[]; let prev=null;
    for(const {t,pi} of sentences(text)){
      let hits=[]; try{ hits=comMatchList(t)||[]; }catch(e){}
      const cue=CUE.test(t);
      const follows=prev&&prev.pi===pi&&FOLLOWS.test(t);
      if(!hits.length&&!(follows&&cue)) { if(!hits.length) prev=null; continue; }   // names nothing the file carries
      const best=hits.find(h=>!h.passing)||hits[0];
      let line, how, also;
      if(follows&&(!best||best.passing||!hits.some(h=>h.subject))){
        line=prev.line; how='follows on from the sentence before';
        also=hits.slice(0,2).map(h=>h.name);
      } else {
        line=best.name; how=best.how+(best.passing?' (in passing)':'');
        also=hits.filter(h=>h!==best).slice(0,2).map(h=>h.name);
      }
      let vag=[]; try{ vag=(comVague(t)||[]).map(v=>v.term+' → '+(v.options||[]).slice(0,3).join(' / ')); }catch(e){}
      const passing=!!(best&&best.passing)&&how.indexOf('follows')<0;
      const row={text:t, line, how, also, vague:vag, cue, pi,
        on: kind==='transcript'? (cue&&!passing) : !passing};
      rows.push(row); prev=row;
    }
    rows.forEach(r=>{ delete r.pi; });
    return rows;
  }
  const MONS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  function dateOf(file){
    const n=file.name;
    let m=n.match(/(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/);
    if(m) return `${+m[3]} ${MONS[Math.max(0,Math.min(11,+m[2]-1))]}`;
    m=n.match(/(\d{1,2})[-_ ]?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i);
    if(m) return `${+m[1]} ${MONS[MONS.findIndex(x=>x.toLowerCase()===m[2].toLowerCase())]}`;
    const d=new Date(file.lastModified||Date.now());
    return `${d.getDate()} ${MONS[d.getMonth()]}`;
  }

  /* ---- the review table ---------------------------------------------------- */
  function lineOptions(cur){
    let names=[]; try{ const ix=comIndex(); names=ix? ix.names.slice().sort((a,b)=>a.localeCompare(b)) : []; }catch(e){}
    if(cur&&!names.includes(cur)) names.unshift(cur);
    return `<option value="">(whole view)</option>`+names.map(n=>`<option value="${esc(n)}"${n===cur?' selected':''}>${esc(n)}</option>`).join('');
  }
  function openReview(doc, isNew){
    closeReview();
    const ov=document.createElement('div'); ov.id='mnOverlay';
    const rows=doc.items;
    ov.innerHTML=`<div id="mnBox" role="dialog" aria-modal="true" aria-labelledby="mnTitle">
      <div id="mnHead">
        <h3 id="mnTitle">Driver notes from ${esc(doc.name)}</h3>
        <p class="subtitle">${rows.length} sentence${rows.length===1?'':'s'} name a line the file carries
          ${doc.kind==='transcript'? ' · a transcript: only sentences with a driver cue are ticked to begin with' : ''}
          · untick what does not belong, change the line where the match is wrong.
          A term listed as ambiguous is settled by an Alias row in the config pack.</p>
        <div id="mnCtl">
          <label>Meeting date <input type="text" id="mnDate" value="${esc(doc.date||'')}" size="8"></label>
          <button class="btn-quiet" id="mnAll">Tick all</button>
          <button class="btn-quiet" id="mnNone">Untick all</button>
          <span class="note" id="mnCount"></span>
        </div>
      </div>
      <div id="mnBody">
        <table class="mntab"><thead><tr><th></th><th>Sentence</th><th>Line</th><th>Matched by</th><th>Also names</th><th>Ambiguous</th></tr></thead>
        <tbody>${rows.map((r,i)=>`<tr class="${r.on?'':'mn-off'}" data-i="${i}">
          <td><input type="checkbox" ${r.on?'checked':''} aria-label="Use this sentence"></td>
          <td class="sent">${esc(r.text)}${r.cue?'':' <span class="mn-nocue" title="No driver wording (driven by, offset, timing …)">no cue</span>'}</td>
          <td><select aria-label="Line">${lineOptions(r.line)}</select></td>
          <td class="muted">${esc(r.how)}</td>
          <td class="muted">${esc((r.also||[]).join(', '))}</td>
          <td class="muted">${esc((r.vague||[]).join('; '))}</td></tr>`).join('')}
        ${rows.length?'':'<tr><td colspan="6" class="muted">Nothing in this file names a line of the loaded extract.</td></tr>'}</tbody></table>
      </div>
      <div id="mnFoot">
        <button class="btn-primary" id="mnUse">${isNew?'Use these notes':'Save'}</button>
        <button class="btn-quiet" id="mnCancel">${isNew?'Discard':'Close'}</button>
        <span class="note">Accepted notes are kept in this browser only and survive a config re-upload; Forget removes them.</span>
      </div></div>`;
    document.body.appendChild(ov);
    const cnt=()=>{ const on=[...ov.querySelectorAll('tbody input[type=checkbox]')].filter(c=>c.checked).length;
      $id('mnCount').textContent=`${on} of ${rows.length} ticked`; };
    ov.querySelectorAll('tbody input[type=checkbox]').forEach(cb=>cb.addEventListener('change',()=>{ cb.closest('tr').classList.toggle('mn-off',!cb.checked); cnt(); }));
    $id('mnAll').onclick=()=>{ ov.querySelectorAll('tbody input[type=checkbox]').forEach(c=>{ c.checked=true; c.closest('tr').classList.remove('mn-off'); }); cnt(); };
    $id('mnNone').onclick=()=>{ ov.querySelectorAll('tbody input[type=checkbox]').forEach(c=>{ c.checked=false; c.closest('tr').classList.add('mn-off'); }); cnt(); };
    $id('mnCancel').onclick=closeReview;
    ov.addEventListener('click',e=>{ if(e.target===ov) closeReview(); });
    $id('mnUse').onclick=()=>{
      ov.querySelectorAll('tbody tr[data-i]').forEach(tr=>{
        const r=rows[+tr.dataset.i];
        r.on=tr.querySelector('input[type=checkbox]').checked;
        r.line=tr.querySelector('select').value;
      });
      doc.date=($id('mnDate').value||'').trim()||doc.date;
      doc.src=`minutes, ${doc.date}`;
      const i=MN.docs.findIndex(d=>d.name===doc.name);
      if(i>=0) MN.docs[i]=doc; else MN.docs.push(doc);
      save(); closeReview(); applyChange();
      status(`${counts(doc).on} driver note${counts(doc).on===1?'':'s'} from ${doc.name} now weave into the commentary.`);
    };
    cnt();
  }
  function closeReview(){ const o=$id('mnOverlay'); if(o) o.remove(); }

  /* ---- the side panel ------------------------------------------------------- */
  function status(t, err){ const s=$id('minStatus'); if(s){ s.textContent=t||''; s.style.color=err? 'var(--status-critical)' : ''; } }
  function paintList(){
    const host=$id('minList'); if(!host) return;
    host.innerHTML=MN.docs.map((d,i)=>{ const c=counts(d);
      return `<div class="mn-doc"><span class="mn-name" title="${esc(d.name)}">${esc(d.name)}</span>
        <span class="muted">${c.on} of ${c.n} · ${esc(d.date||'')}</span>
        <button class="linklike" data-rev="${i}">Review</button> ·
        <button class="linklike" data-forget="${i}">Forget</button></div>`; }).join('')
      ||'<div class="note">No minutes loaded. Drivers written in a meeting join the commentary from here.</div>';
    host.querySelectorAll('[data-rev]').forEach(b=>b.onclick=()=>openReview(MN.docs[+b.dataset.rev], false));
    host.querySelectorAll('[data-forget]').forEach(b=>b.onclick=()=>{ MN.docs.splice(+b.dataset.forget,1); save(); applyChange(); status('Forgotten.'); });
  }
  /* the Commentary templates note counts what is loaded: the minutes count too */
  function noteHook(){
    if(typeof window.comShow!=='function'||window.comShow.__minutes) return;
    const _cs=window.comShow;
    const w=function(){
      const r=_cs.apply(this,arguments);
      const n=$id('comInNote'), mine=accepted().length;
      if(n&&mine&&n.textContent.indexOf('from minutes')<0){
        const add=`${mine} from minutes`;
        n.textContent=/^No written commentary/.test(n.textContent)
          ? `\u2713 ${add} \u2014 leading the blocks they name.`
          : n.textContent.replace(' \u2014 leading', ` \u00b7 ${add} \u2014 leading`);
      }
      return r;
    };
    w.__minutes=true; window.comShow=w;
  }
  function applyChange(){
    paintList();
    try{ calcCacheClear(); }catch(e){}
    try{ if(S.model) render(); }catch(e){}
    try{ noteHook(); window.comShow(); }catch(e){}
  }
  async function handleFile(file){
    if(!S.model){ status('Load the extract first — notes are matched against its line names.', true); return; }
    status(`Reading ${file.name}…`);
    try{
      const {text,kind}=await readFile(file);
      const items=analyse(text, kind);
      const doc={name:file.name, date:dateOf(file), kind, items, src:''};
      doc.src=`minutes, ${doc.date}`;
      status(items.length? `${items.length} candidate sentence${items.length===1?'':'s'} in ${file.name} — review them.`
                         : `Nothing in ${file.name} names a line of the loaded extract.`);
      openReview(doc, true);
    }catch(e){ status('Could not read '+file.name+': '+(e&&e.message||e), true); }
  }
  function mount(){
    if($id('minPanel')) return true;
    const tp=$id('tplPanel'); if(!tp||typeof S==='undefined'||typeof comInLoaded!=='function') return false;
    const st=document.createElement('style');
    st.textContent=`
#minPanel{padding:2px 16px 12px;display:flex;flex-direction:column;gap:7px;font-size:12px}
#minPanel label{display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--text-muted);margin:0 2px}
#minPanel .mn-doc{display:flex;flex-wrap:wrap;gap:4px 6px;align-items:baseline;font-size:11.5px;padding:4px 0;border-top:1px solid var(--border)}
#minPanel .mn-name{font-weight:600;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#minPanel .linklike{background:none;border:none;color:var(--hsbc-red);cursor:pointer;font:inherit;font-size:11.5px;padding:0}
#minPanel .linklike:hover{text-decoration:underline}
#mnOverlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;display:flex;align-items:center;justify-content:center;padding:24px}
#mnBox{background:var(--surface-1);color:var(--text-primary);border:1px solid var(--border);border-radius:10px;max-width:1140px;width:100%;
  max-height:92vh;display:flex;flex-direction:column;box-shadow:0 12px 40px rgba(0,0,0,.35)}
#mnHead{padding:16px 20px 10px;border-bottom:1px solid var(--border)}
#mnHead h3{margin:0 0 4px;font-size:15px}
#mnHead .subtitle{margin:0 0 8px}
#mnCtl{display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:12px}
#mnCtl label{display:flex;gap:6px;align-items:center;color:var(--text-muted)}
#mnCtl input{padding:4px 6px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface-2);color:var(--text-primary);font:inherit}
#mnCtl .btn-quiet{padding:4px 10px;font-size:11.5px}
#mnBody{overflow:auto;padding:0 20px}
table.mntab{width:100%;border-collapse:collapse;font-size:12.5px}
table.mntab th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.6px;color:var(--text-muted);padding:8px 8px 6px;position:sticky;top:0;background:var(--surface-1);border-bottom:1px solid var(--border-strong)}
table.mntab td{padding:7px 8px;border-bottom:1px solid var(--border);vertical-align:top}
table.mntab td.sent{min-width:340px;line-height:1.4}
table.mntab td.muted{color:var(--text-muted);font-size:11.5px}
table.mntab select{max-width:220px;padding:4px 6px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface-2);color:var(--text-primary);font:inherit;font-size:12px}
table.mntab tr.mn-off td{opacity:.5}
.mn-nocue{font-size:10px;color:var(--text-muted);border:1px solid var(--border);border-radius:10px;padding:0 6px;margin-left:4px;white-space:nowrap}
#mnFoot{padding:12px 20px;border-top:1px solid var(--border);display:flex;gap:10px;align-items:center;flex-wrap:wrap}
#mnFoot .btn-primary{padding:7px 18px;font-size:13px}`;
    document.head.appendChild(st);
    const row=document.createElement('div');
    row.className='sect sect-fold'; row.id='sectMinRow'; row.title='Collapse / expand';
    row.style.cssText='display:flex;align-items:center;gap:6px;cursor:pointer';
    row.innerHTML='<span data-lbl="sect_minutes">Driver notes from minutes</span><span id="minCaret" style="margin-left:auto;font-size:10px;color:inherit">▾</span>';
    const panel=document.createElement('div'); panel.id='minPanel';
    panel.innerHTML=`
      <button class="btn-quiet" id="btnMinUpload">⤒ Upload minutes (docx / txt / vtt)</button>
      <input type="file" id="minInput" accept=".docx,.txt,.text,.vtt,.pdf,.doc" hidden>
      <div id="minList"></div>
      <label><input type="checkbox" id="minPrefer"${MN.prefer?' checked':''}> prefer minutes over the config's commentary</label>
      <label><input type="checkbox" id="minMark"${MN.mark?' checked':''}> mark the source in the text</label>
      <div class="note" id="minStatus"></div>`;
    tp.insertAdjacentElement('afterend',row);
    row.insertAdjacentElement('afterend',panel);
    let folded=false; try{ folded=localStorage.getItem('iwpbsg_min_collapsed')==='1'; }catch(e){}
    const fold=f=>{ panel.hidden=f; $id('minCaret').textContent=f?'▸':'▾'; try{ localStorage.setItem('iwpbsg_min_collapsed',f?'1':'0'); }catch(e){} };
    fold(folded);
    row.addEventListener('click',()=>fold(!panel.hidden));
    const fi=$id('minInput');
    $id('btnMinUpload').addEventListener('click',()=>fi.click());
    fi.addEventListener('change',()=>{ const f=fi.files&&fi.files[0]; if(f) handleFile(f); fi.value=''; });
    $id('minPrefer').addEventListener('change',e=>{ MN.prefer=e.target.checked; save(); applyChange(); });
    $id('minMark').addEventListener('change',e=>{ MN.mark=e.target.checked; save(); applyChange(); });
    paintList();
    return true;
  }

  /* ---- hooks: the notes join the loaded commentary; a clause names its source */
  const _loaded=comInLoaded;
  comInLoaded=function(){
    const base=_loaded.apply(this,arguments), mine=accepted();
    if(!mine.length) return base;
    return MN.prefer? mine.concat(base) : base.concat(mine);
  };
  if(typeof comDigest==='function'){
    const _digest=comDigest;
    comDigest=function(txt,line,cap){
      const d=_digest.apply(this,arguments);
      if(d&&MN.mark&&SRC.has(txt)&&d.indexOf('(minutes, ')<0) return d+' ('+SRC.get(txt)+')';
      return d;
    };
  }
  window.__minutesIngest=function(text, kind, name){ const doc={name:name||'pasted', date:dateOf({name:name||'',lastModified:Date.now()}), kind:kind||'doc', items:analyse(text,kind||'doc'), src:''}; doc.src=`minutes, ${doc.date}`; return doc; };
  window.__minutesOpen=openReview;
  window.__minutesState=MN;
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mount); else mount();
  setTimeout(noteHook,0);
})();
