/* ================= Enrich commentary =====================================
   Commentary written outside the config pack — any Word write-up (.docx),
   plain text (.txt) or a transcript (.vtt, or a "Name: …" text export), one
   file or many at once — is read in the browser, cut into sentences, and
   every sentence that names a line the file carries is matched the way the
   config's own free commentary is: comMatchList decides which line it is
   about. The sentences enrich the commentary the moment the file is read;
   a review table (optional, from the panel or the status line) shows each
   sentence, the line it was matched to (changeable), how it matched, and
   the terms that were ambiguous, so a wrong match is a tick away from
   fixed. In a transcript only sentences that carry a driver cue (driven by,
   offset, timing, one-off …) are used to begin with; a sentence that names
   no line is kept, unticked, as whole-view commentary the reviewer can turn
   on. Accepted sentences are the ONLY written commentary the page uses —
   the config pack's Commentary sheet is not read for narrative any more —
   and the House commentary, the Financial Summary, All commentaries, the
   widgets and the exports weave them in with scope-following, digesting and
   the per-block cap intact. The sources are named in a line under each
   commentary block; an inline marker on every clause is a checkbox away.
   Every upload is stamped with the scope it was made under (business,
   region, country — parents filled in from the data). At that scope or a
   narrower one the sentences read as written; at a wider scope they roll up
   as one clause per line that names each place — "Singapore: driven by
   deposit margin … · Hong Kong: behind on fees …"; under a different
   country they are held back.
   Word files are unzipped natively (DecompressionStream); PDF is not read
   here — save it as Word or text.
   AI review (on when a model is found): after the rules have run, the
   sentences and the extract's line names go to a language model running on
   this desktop — Ollama or LM Studio, reached at localhost, so nothing leaves
   the machine — or to Claude / an approved gateway if the page is pointed at
   one. The model says, per sentence, whether it is a driver, which line it
   belongs to and how sure it is; confident answers are stitched with no
   ticking, doubtful ones are left unticked with the reason shown in Review.
   If no model answers, the rules' decisions stand. */
(function(){
  'use strict';
  const LSKEY='iwpbsg_enrich';
  function $id(x){ return document.getElementById(x); }
  const EN={docs:[], mark:false,
    ai:{on:true, endpoint:'', model:'', key:'', threshold:0.7, rewrite:false, detected:null}};
  try{ const st=JSON.parse(localStorage.getItem(LSKEY)||'{}'); const ai={...EN.ai, ...(st.ai||{})}; Object.assign(EN, st); EN.ai=ai; }catch(e){}
  if(!Array.isArray(EN.docs)) EN.docs=[];
  function save(){ try{ localStorage.setItem(LSKEY, JSON.stringify(EN)); }catch(e){} }

  /* ---- the accepted sentences, as commentary items, at the page's scope ---- */
  const SRC=new Map();                      // sentence -> source label
  const ROLL=new Map();                     // rolled-up clause -> how many places it names
  const GEO=['business','region','country'];
  let DIGEST=null;                          // the page's own comDigest, before it was wrapped
  function scopeNow(){
    const one=k=>{ const st=S&&S.filters&&S.filters[k]; return (st&&st.size===1)? String([...st][0]) : ''; };
    return {business:one('business'), region:one('region'), country:one('country')};
  }
  /* the scope an upload is made under, with its parents filled in from the
     rows, so a country upload knows its region and business */
  function scopeStamp(){
    const sc=scopeNow(), m=S&&S.model; if(!m) return sc;
    const only=(k,pred)=>{ const v=new Set(); for(const r of m.rows){ if(pred(r)&&r.d[k]) v.add(String(r.d[k])); if(v.size>1) break; } return v.size===1? [...v][0] : ''; };
    if(sc.country){
      if(!sc.region) sc.region=only('region', r=>String(r.d.country||'')===sc.country);
      if(!sc.business) sc.business=only('business', r=>String(r.d.country||'')===sc.country);
    } else if(sc.region&&!sc.business) sc.business=only('business', r=>String(r.d.region||'')===sc.region);
    return sc;
  }
  function scopeLabel(sc){ return (sc&&(sc.country||sc.region||sc.business))||''; }
  function accepted(opts){
    const all=!!(opts&&opts.all), cur=all? {} : scopeNow();
    const out=[]; SRC.clear(); ROLL.clear();
    const groups=new Map();                 // line -> the narrower places' sentences
    for(const d of EN.docs){
      const ds=d.scope||{};
      // a different place altogether: held back
      if(!all&&GEO.some(k=>cur[k]&&ds[k]&&cur[k]!==ds[k])) continue;
      // narrower than the page: rolled up under the place's name
      const rollK=all? '' : ['country','region','business'].find(k=>ds[k]&&!cur[k]);
      const label=rollK? ds[rollK] : '';
      for(const it of (d.items||[])){
        if(!it.on) continue;
        if(!label){ out.push({view:'', line:it.line||'', text:it.text, src:d.label}); SRC.set(it.text, d.label); continue; }
        const key=it.line||'';
        (groups.get(key)||groups.set(key,[]).get(key)).push({label, text:it.text, src:d.label});
      }
    }
    for(const [line,parts] of groups){
      // one clause per line, each place named once, in upload order
      const byLabel=new Map(); for(const p of parts) if(!byLabel.has(p.label)) byLabel.set(p.label,p);
      const segs=[...byLabel.values()].map(p=>{
        let dg=''; try{ dg=DIGEST? DIGEST(p.text, line) : ''; }catch(e){}
        if(!dg) dg=String(p.text).replace(/[.\s]+$/,'');
        // "Singapore: ahead of plan", not "Singapore: was ahead of plan"
        dg=dg.replace(/^(?:was|were|is|are|has been|have been|had been)\s+/i,'');
        return `${p.label}: ${dg}`;
      });
      const text=segs.join(' \u00b7 ')+'.';
      const srcs=[...new Set([...byLabel.values()].map(p=>p.src))].join(', ');
      out.push({view:'', line, text, src:srcs}); SRC.set(text, srcs); ROLL.set(text, byLabel.size);
    }
    return out;
  }
  function counts(d){ const n=(d.items||[]).length, on=(d.items||[]).filter(i=>i.on).length; return {n,on}; }

  /* ---- reading a file -------------------------------------------------------- */
  const unesc=s=>String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(+n)).replace(/&amp;/g,'&');
  const U8=new TextDecoder('utf-8');
  const u16=(b,i)=>b[i]|(b[i+1]<<8), u32=(b,i)=>(b[i]|(b[i+1]<<8)|(b[i+2]<<16))+(b[i+3]*16777216);
  async function inflateRaw(bytes){
    if(typeof DecompressionStream!=='function')
      throw new Error('This browser cannot unzip a Word file — upload the text as .txt instead.');
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
    if(ext==='pdf') throw new Error('PDF is not read here — save it as Word (.docx) or text (.txt) and upload that.');
    if(ext==='doc') throw new Error('Old .doc files are not read — save as .docx and upload that.');
    if(['xlsx','xls','csv','pptx','zip'].includes(ext)) throw new Error('Not a text document — upload Word, text or a transcript.');
    const t=await file.text();
    return looksTranscript(t)? {text:transcriptText(t), kind:'transcript'} : {text:t, kind:'doc'};
  }

  /* ---- sentences, and which line each is about --------------------------- */
  const CUE=/\b(?:driven|drives?|driver|due to|because|owing|reflect(?:s|ing|ed)?|offset|partly|partially|higher|lower|up|down|ahead|behind|above|below|increas(?:e|ed|es|ing)|decreas(?:e|ed|es|ing)|grew|growth|declin(?:e|ed|es|ing)|fell|rose|one[- ]off|timing|favou?rable|adverse|impact(?:ed|s)?|mix|margin|volume|pricing|rates?|spread|attrition|run[- ]rate|seasonal|delay(?:ed|s)?|slippage|write[- ]?(?:back|off)|provision|release|uplift|shortfall|overrun|saving|savings)\b/i;
  const FILLER=/\b(?:um+|uh+|erm|you know|i mean|sort of|kind of|basically|actually|okay so|so yeah)\b[,]?\s*/gi;
  const SKIP=/^(?:action(?:s| items?)?|ai|next steps?|attendees?|present|apologies|agenda|minutes|date|time|recording|chair|note taker|circulated|prepared by|version|page \d)\b/i;
  /* sentences, each knowing which paragraph (speaker turn) it came from */
  function sentences(text){
    const out=[], seen=new Set();
    const paras=String(text).split(/\n+/).map(x=>x.trim()).filter(Boolean);
    // a document's title is not commentary: the first line, if it does not
    // end like a sentence, is a heading — "IWPB Singapore — performance
    // review" would otherwise pin itself to the business line it names
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
      if(!hits.length&&!(follows&&cue)){
        // names nothing the file carries: kept as commentary for the view as
        // a whole, unticked — the reviewer decides whether it belongs
        rows.push({text:t, line:'', how:'no line named', also:[], vague:[], cue, pi, on:false, loose:true});
        prev=null; continue;
      }
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
  /* the source label a clause carries: the file's name, made readable */
  function labelOf(name){
    let s=String(name).replace(/\.[^.]+$/,'').replace(/[_\-]+/g,' ').replace(/\s+/g,' ').trim();
    if(s.length>32) s=s.slice(0,31).replace(/\s+\S*$/,'')+'…';
    return s||'source';
  }

  /* ---- AI review: a model on this desktop, or Claude, decides ---------------- */
  /* Where the model is found, in order: an endpoint typed under Advanced (or
     set in the config pack), else the one found running on this desktop, else
     none — and then the rules decide alone. The provider is read off the URL:
     the Anthropic Messages API for Claude, OpenAI-style chat for the rest. */
  const AI_LOCAL=['http://localhost:11434','http://127.0.0.1:11434','http://localhost:1234','http://127.0.0.1:1234'];
  function cfgAi(k,d){ try{ const v=cfgS('ai_'+k); return (v==null||String(v).trim()==='')? d : String(v).trim(); }catch(e){ return d; } }
  function providerOf(url){ return /api\.anthropic\.com|\/v1\/messages/.test(url)? 'anthropic' : 'openai'; }
  function aiCfg(){
    const a=EN.ai||{};
    const th=parseFloat(a.threshold!=null? a.threshold : cfgAi('threshold','0.7'));
    const threshold=isFinite(th)? Math.min(1,Math.max(0,th)) : 0.7;
    const rewrite=a.rewrite!=null? !!a.rewrite : /^y/i.test(cfgAi('rewrite','N'));
    const on=a.on!==false;
    const custom=String(a.endpoint||cfgAi('endpoint','')).trim();
    if(custom){
      const provider=providerOf(custom);
      return {on, source:'custom', provider, endpoint:custom, key:a.key||cfgAi('key',''),
              model:a.model||cfgAi('model','')||(provider==='anthropic'? 'claude-opus-5' : ''), threshold, rewrite};
    }
    const d=a.detected;
    if(d&&d.endpoint&&d.model)
      return {on, source:'desktop', provider:'openai', endpoint:d.endpoint, model:a.model||d.model, key:'', threshold, rewrite, server:d.server};
    return {on:false, wanted:on, source:'none', provider:'openai', endpoint:'', model:'', key:'', threshold, rewrite};
  }
  /* look for a model server on this desktop: Ollama and LM Studio both list
     their models at /v1/models; the first chat-capable one is taken */
  let AI_PROBING=null;
  async function aiProbe(){
    if(AI_PROBING) return AI_PROBING;
    AI_PROBING=(async()=>{
      for(const base of AI_LOCAL){
        try{
          const ctl=new AbortController(); const tm=setTimeout(()=>ctl.abort(),1500);
          const r=await fetch(base+'/v1/models',{signal:ctl.signal}); clearTimeout(tm);
          if(!r.ok) continue;
          const j=await r.json(); const ids=(j.data||j.models||[]).map(m=>m.id||m.name).filter(Boolean);
          if(!ids.length) continue;
          const chat=ids.filter(i=>!/embed|whisper|clip|rerank|vision/i.test(i));
          const pick=chat.find(i=>/instruct|chat/i.test(i))||chat[0]||ids[0];
          EN.ai.detected={endpoint:base+'/v1/chat/completions', model:pick, models:ids.slice(0,30), at:Date.now(),
                          server:/:1234/.test(base)? 'LM Studio' : 'Ollama'};
          save(); return EN.ai.detected;
        }catch(e){}
      }
      EN.ai.detected=null; save(); return null;
    })();
    try{ return await AI_PROBING; } finally{ AI_PROBING=null; }
  }
  function aiPrompt(doc, names){
    const m=S.model; let period='';
    try{ period=`${cmpLabels(m).period}; YTD ${ytdLabel()}; current year ${m.yearCY}, prior year ${m.yearPY}`; }catch(e){}
    const sys=[
      'You classify sentences from a finance write-up or meeting transcript for a management-reporting dashboard.',
      'The dashboard writes commentary line by line (revenue and cost lines, products, segments). A sentence is worth keeping only if it explains WHY a figure moved or WHAT drove it — a driver, a cause, an offset, timing, a one-off, mix, volume, margin, rates.',
      'Drop: chit-chat, questions, logistics, action items, agenda lines, restatements of a figure with no reason, and anything about a future plan that does not explain the current numbers.',
      'For each kept sentence pick exactly one line from the list of line names — the one the sentence explains. Use the empty string only when the sentence is a general remark about the whole business.',
      'A sentence that opens on a pronoun continues the sentence before it: it belongs to that sentence\'s line.',
      'Confidence is 0 to 1: how sure you are that the sentence is a driver AND that the line is right.',
      'If asked to reword, rewrite the sentence in the concise, past-tense style of a finance pack, keeping every fact and figure, removing speech filler — otherwise return the sentence unchanged.',
      'Answer with JSON only, no prose, exactly: {"decisions":[{"i":<index>,"use":<true|false>,"line":"<one of the line names or empty>","confidence":<0-1>,"reason":"<short>","text":"<sentence, reworded only if asked>"}]}'
    ].join(' ');
    const user=JSON.stringify({
      reporting_period: period,
      reword: !!aiCfg().rewrite,
      line_names: names,
      sentences: doc.items.map((it,i)=>({i, text:it.text, rules_line:it.line||'', rules_use:!!it.on}))
    });
    return {sys,user};
  }
  function aiParse(txt){
    const t=String(txt||'').trim();
    const a=t.indexOf('{'), b=t.lastIndexOf('}');
    if(a<0||b<a) throw new Error('the model did not answer with JSON');
    const o=JSON.parse(t.slice(a,b+1));
    if(!o||!Array.isArray(o.decisions)) throw new Error('the model\'s JSON has no decisions');
    return o.decisions;
  }
  const AI_SCHEMA={type:'object', additionalProperties:false, required:['decisions'],
    properties:{decisions:{type:'array', items:{type:'object', additionalProperties:false,
      required:['i','use','line','confidence','reason','text'],
      properties:{i:{type:'integer'}, use:{type:'boolean'}, line:{type:'string'}, confidence:{type:'number'},
                  reason:{type:'string'}, text:{type:'string'}}}}}};
  async function aiCall(c, sys, user){
    const ctl=new AbortController(); const tm=setTimeout(()=>ctl.abort(), 120000);
    try{
      if(c.provider==='anthropic'){
        // raw Messages API from the browser — no SDK can ride inside this file
        const r=await fetch(c.endpoint,{method:'POST', signal:ctl.signal,
          headers:{'content-type':'application/json','x-api-key':c.key,'anthropic-version':'2023-06-01',
                   'anthropic-dangerous-direct-browser-access':'true'},
          body:JSON.stringify({model:c.model, max_tokens:8000, system:sys,
            messages:[{role:'user', content:user}],
            output_config:{format:{type:'json_schema', schema:AI_SCHEMA}}})});
        if(!r.ok) throw new Error(`HTTP ${r.status} from ${c.endpoint}`);
        const j=await r.json();
        if(j.stop_reason==='refusal') throw new Error('the model declined this request');
        return (j.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('');
      }
      // OpenAI-style chat completions: Ollama, LM Studio, and most internal gateways
      const headers={'content-type':'application/json'};
      if(c.key) headers['authorization']='Bearer '+c.key;
      const r=await fetch(c.endpoint,{method:'POST', signal:ctl.signal, headers,
        body:JSON.stringify({model:c.model, temperature:0, stream:false, response_format:{type:'json_object'},
          messages:[{role:'system', content:sys},{role:'user', content:user}]})});
      if(!r.ok) throw new Error(`HTTP ${r.status} from ${c.endpoint}`);
      const j=await r.json();
      const ch=j.choices&&j.choices[0];
      return (ch&&ch.message&&(ch.message.content||''))||j.response||'';
    } finally{ clearTimeout(tm); }
  }
  function aiFailWhy(e, c){
    const msg=(e&&e.message)||String(e);
    if(/Failed to fetch|NetworkError|Load failed|aborted/i.test(msg)){
      return c.provider==='anthropic'
        ? 'could not reach the endpoint (network or CORS)'
        : 'no model answered at '+c.endpoint+' — is it running, and does it allow this page? (Ollama: start with OLLAMA_ORIGINS=* · LM Studio: switch CORS on)';
    }
    return msg;
  }
  /* the decisions, applied: confident answers stitch themselves, doubtful ones
     stay unticked with the reason on show; the model never adds a sentence */
  async function aiReview(doc){
    const c=aiCfg(); if(!c.on||!doc.items.length) return null;
    if(!c.model) throw new Error('no model name set');
    let names=[]; try{ const ix=comIndex(); names=ix? ix.names.slice(0,400) : []; }catch(e){}
    const {sys,user}=aiPrompt(doc, names);
    const raw=await aiCall(c, sys, user);
    const dec=aiParse(raw);
    const byI=new Map(dec.map(d=>[+d.i, d]));
    let applied=0, unsure=0;
    doc.items.forEach((it,i)=>{
      const d=byI.get(i); if(!d) return;
      const conf=Math.max(0,Math.min(1, +d.confidence||0));
      const line=(typeof d.line==='string'&&(d.line===''||names.includes(d.line)))? d.line : it.line;
      it.ai={use:!!d.use, line, confidence:conf, reason:String(d.reason||'').slice(0,160)};
      if(!!d.use&&conf>=c.threshold){
        it.on=true; it.line=line; applied++;
        if(c.rewrite&&d.text&&String(d.text).trim()&&String(d.text).trim()!==it.text){ it.orig=it.orig||it.text; it.text=String(d.text).trim(); it.reworded=true; }
      } else if(!d.use&&conf>=c.threshold){
        it.on=false; applied++;
      } else { it.on=false; unsure++; }
    });
    doc.ai={provider:c.provider, model:c.model, applied, unsure, at:Date.now()};
    return doc.ai;
  }
  async function aiTest(){
    const c=aiCfg(), out=$id('enrichAiStatus');
    if(!c.endpoint||!c.model){ out.textContent='No model to test — none found on this desktop and no endpoint typed above.'; return; }
    out.textContent='Testing '+c.endpoint+' …';
    try{
      const raw=await aiCall(c, 'Answer with JSON only: {"decisions":[{"i":0,"use":true,"line":"","confidence":1,"reason":"test","text":"ok"}]}', '{"sentences":[{"i":0,"text":"test"}]}');
      aiParse(raw);
      out.textContent='✓ '+c.model+' answered at '+c.endpoint+'.';
    }catch(e){ out.textContent='✗ '+aiFailWhy(e,c); }
  }

  /* ---- the review table ---------------------------------------------------- */
  function lineOptions(cur){
    let names=[]; try{ const ix=comIndex(); names=ix? ix.names.slice().sort((a,b)=>a.localeCompare(b)) : []; }catch(e){}
    if(cur&&!names.includes(cur)) names.unshift(cur);
    return `<option value="">(whole view)</option>`+names.map(n=>`<option value="${esc(n)}"${n===cur?' selected':''}>${esc(n)}</option>`).join('');
  }
  function openReview(docs){
    closeReview();
    const ov=document.createElement('div'); ov.id='enOverlay';
    const total=docs.reduce((s,d)=>s+d.items.length,0);
    const anyTr=docs.some(d=>d.kind==='transcript');
    const body=docs.map((d,di)=>`
      <tr class="en-file"><td colspan="6">
        <b>${esc(d.name)}</b> <span class="muted">· ${d.kind==='transcript'?'transcript':'document'}${scopeLabel(d.scope)? ' · uploaded under '+esc([d.scope.business,d.scope.region,d.scope.country].filter(Boolean).join(' › ')) : ' · whole book'} · ${d.items.length} sentence${d.items.length===1?'':'s'}${d.items.some(i=>i.loose)? ', '+d.items.filter(i=>i.loose).length+' naming no line' : ''}${d.ai? ` · reviewed by ${esc(d.ai.model)}: ${d.ai.applied} decided, ${d.ai.unsure} unsure` : ''}</span>
        <label>Source label <input type="text" data-lbl-d="${di}" value="${esc(d.label||'')}" size="22"></label>
      </td></tr>`+
      d.items.map((r,i)=>`<tr class="${r.on?'':'en-off'}" data-d="${di}" data-i="${i}">
        <td><input type="checkbox" ${r.on?'checked':''} aria-label="Use this sentence"></td>
        <td class="sent">${esc(r.text)}${r.reworded?` <span class="en-nocue" title="Reworded by the model — original: ${esc(r.orig||'')}">reworded</span>`:''}${r.cue?'':' <span class="en-nocue" title="No driver wording (driven by, offset, timing …)">no cue</span>'}</td>
        <td><select aria-label="Line">${lineOptions(r.line)}</select></td>
        <td class="muted">${esc(r.how)}</td>
        <td class="muted">${esc((r.also||[]).join(', '))}</td>
        <td class="muted">${esc((r.vague||[]).join('; '))}</td>${docs.some(x=>x.ai)? `<td class="muted">${r.ai? `${r.ai.use?'use':'drop'} · ${Math.round(r.ai.confidence*100)}%${r.ai.line&&r.ai.line!==r.line?' · '+esc(r.ai.line):''}<br>${esc(r.ai.reason)}` : ''}</td>` : ''}</tr>`).join('')+
      (d.items.length?'':`<tr><td colspan="6" class="muted">Nothing in this file names a line of the loaded extract.</td></tr>`)
    ).join('');
    ov.innerHTML=`<div id="enBox" role="dialog" aria-modal="true" aria-labelledby="enTitle">
      <div id="enHead">
        <h3 id="enTitle">Enrich commentary — ${docs.length===1? esc(docs[0].name) : docs.length+' files'}</h3>
        <p class="subtitle">${total} sentence${total===1?'':'s'} · ticked ones enrich the commentary now
          ${anyTr? ' · in a transcript only sentences with a driver cue start ticked' : ''}
          · a sentence that names no line is kept unticked as whole-view commentary
          · untick what does not belong, change the line where the match is wrong; an ambiguous term is settled by an Alias row in the config pack.</p>
        <div id="enCtl">
          <button class="btn-quiet" id="enAll">Tick all</button>
          <button class="btn-quiet" id="enNone">Untick all</button>
          <span class="note" id="enCount"></span>
        </div>
      </div>
      <div id="enBody">
        <table class="entab"><thead><tr><th></th><th>Sentence</th><th>Line</th><th>Matched by</th><th>Also names</th><th>Ambiguous</th>${docs.some(d=>d.ai)?'<th>AI</th>':''}</tr></thead>
        <tbody>${body}</tbody></table>
      </div>
      <div id="enFoot">
        <button class="btn-primary" id="enUse">Save</button>
        <button class="btn-quiet" id="enCancel">Close</button>
        <span class="note">Sentences are kept in this browser only and survive a config re-upload; Forget removes a file's.</span>
      </div></div>`;
    document.body.appendChild(ov);
    const cnt=()=>{ const on=[...ov.querySelectorAll('tr[data-i] input[type=checkbox]')].filter(c=>c.checked).length;
      $id('enCount').textContent=`${on} of ${total} ticked`; };
    ov.querySelectorAll('tr[data-i] input[type=checkbox]').forEach(cb=>cb.addEventListener('change',()=>{ cb.closest('tr').classList.toggle('en-off',!cb.checked); cnt(); }));
    $id('enAll').onclick=()=>{ ov.querySelectorAll('tr[data-i] input[type=checkbox]').forEach(c=>{ c.checked=true; c.closest('tr').classList.remove('en-off'); }); cnt(); };
    $id('enNone').onclick=()=>{ ov.querySelectorAll('tr[data-i] input[type=checkbox]').forEach(c=>{ c.checked=false; c.closest('tr').classList.add('en-off'); }); cnt(); };
    $id('enCancel').onclick=closeReview;
    ov.addEventListener('click',e=>{ if(e.target===ov) closeReview(); });
    $id('enUse').onclick=()=>{
      ov.querySelectorAll('tr[data-i]').forEach(tr=>{
        const r=docs[+tr.dataset.d].items[+tr.dataset.i];
        r.on=tr.querySelector('input[type=checkbox]').checked;
        r.line=tr.querySelector('select').value;
      });
      ov.querySelectorAll('[data-lbl-d]').forEach(inp=>{ const d=docs[+inp.dataset.lblD]; d.label=(inp.value||'').trim()||labelOf(d.name); });
      let on=0;
      for(const d of docs){
        const i=EN.docs.findIndex(x=>x.name===d.name);
        if(i>=0) EN.docs[i]=d; else EN.docs.push(d);
        on+=counts(d).on;
      }
      save(); closeReview(); applyChange();
      status(`${on} sentence${on===1?'':'s'} from ${docs.length===1? docs[0].name : docs.length+' files'} ${on===1?'enriches':'enrich'} the commentary.`);
    };
    cnt();
  }
  function closeReview(){ const o=$id('enOverlay'); if(o) o.remove(); }

  /* ---- the side panel ------------------------------------------------------- */
  function status(t, err, review){
    const s=$id('enrichStatus'); if(!s) return;
    s.textContent=(t||'')+(review&&review.length?' ':''); s.style.color=err? 'var(--status-critical)' : '';
    if(review&&review.length){
      const b=document.createElement('button'); b.className='linklike'; b.textContent='Review';
      b.style.marginLeft='6px'; b.onclick=()=>openReview(review);
      s.appendChild(b);
    }
  }
  function paintList(){
    const host=$id('enrichList'); if(!host) return;
    host.innerHTML=EN.docs.map((d,i)=>{ const c=counts(d);
      return `<div class="en-doc"><span class="en-name" title="${esc(d.name)}">${esc(d.name)}</span>
        <span class="muted">${c.on} of ${c.n} · ${esc(d.label||'')}${scopeLabel(d.scope)? ' · '+esc(scopeLabel(d.scope)) : ''}</span>
        <button class="linklike" data-rev="${i}">Review</button> ·
        <button class="linklike" data-forget="${i}">Forget</button></div>`; }).join('')
      ||'<div class="note">No sources loaded. Anything written about the numbers — a write-up, notes, a transcript — can enrich the commentary from here.</div>';
    host.querySelectorAll('[data-rev]').forEach(b=>b.onclick=()=>openReview([EN.docs[+b.dataset.rev]]));
    host.querySelectorAll('[data-forget]').forEach(b=>b.onclick=()=>{ EN.docs.splice(+b.dataset.forget,1); save(); applyChange(); status('Forgotten.'); });
    const fa=$id('btnEnrichForgetAll'); if(fa) fa.hidden=EN.docs.length<2;
    const ra=$id('btnEnrichReviewAll'); if(ra) ra.hidden=EN.docs.length<2;
  }
  /* the Commentary templates note reports the uploads — the only written
     commentary in play */
  function noteText(){
    const mine=accepted().length;
    if(!mine) return 'No uploaded commentary yet — add a write-up, notes or a transcript under Enrich commentary; every block reads its own figures until then.';
    let held=0; try{ held=comInAllUnscoped().length-comInAll().length; }catch(e){}
    return `\u2713 ${mine} from uploaded source${EN.docs.length===1?'':'s'}`
      +(held>0? ` \u00b7 ${held} held back by the current filters \u2014 written about lines outside this scope` : '')
      +' \u2014 leading the blocks they name.';
  }
  function noteHook(){
    if(typeof window.comShow!=='function'||window.comShow.__enrich) return;
    const _cs=window.comShow;
    const w=function(){ const r=_cs.apply(this,arguments); const n=$id('comInNote'); if(n) n.textContent=noteText(); return r; };
    w.__enrich=true; window.comShow=w;
  }
  /* the sources named once under each commentary block, not on every clause */
  function paintSources(){
    const cur=scopeNow();
    const labels=[...new Set(EN.docs.filter(d=>counts(d).on&&!GEO.some(k=>cur[k]&&d.scope&&d.scope[k]&&cur[k]!==d.scope[k]))
      .map(d=>d.label+(scopeLabel(d.scope)&&scopeLabel(d.scope)!==scopeLabel(cur)? ' ('+scopeLabel(d.scope)+')' : '')))];
    for(const id of ['houseDisc','fsumDisc','caDisc']){
      const dz=$id(id); if(!dz) continue;
      let p=$id(id+'Src');
      if(!labels.length){ if(p) p.remove(); continue; }
      if(!p){ p=document.createElement('p'); p.id=id+'Src'; p.className='note en-sources'; dz.insertAdjacentElement('afterend',p); }
      p.textContent='Commentary enriched from: '+labels.join(', ')+'.';
      p.hidden=dz.hidden; p.style.display=dz.style.display;
    }
  }
  function applyChange(){
    paintList();
    try{ calcCacheClear(); }catch(e){}
    try{ if(S.model) render(); }catch(e){}
    try{ noteHook(); window.comShow(); }catch(e){}
    paintSources();
  }
  const _render=render;
  render=function(){ const r=_render.apply(this,arguments); try{ paintSources(); }catch(e){} return r; };
  /* every AI commentary block carries Retrieve and Regenerate on its header
     row; Enrich sits beside them, and opens the same file picker */
  if(typeof aiComment==='function'){
    const _aic=aiComment;
    aiComment=function(host){
      const r=_aic.apply(this,arguments);
      try{
        const row=host&&host.querySelector&&host.querySelector('.aicom-btns');
        if(row&&!row.hidden&&!row.querySelector('.aicom-e')){
          const b=document.createElement('button'); b.type='button'; b.className='aicom-e';
          b.title='Add a write-up, notes or a transcript — its sentences enrich this commentary';
          b.textContent='⤒ Enrich';
          b.addEventListener('click',e=>{ e.stopPropagation(); const fi=$id('enrichInput'); if(fi) fi.click(); });
          row.insertBefore(b, row.firstChild);
        }
      }catch(e){}
      return r;
    };
  }
  async function handleFiles(files){
    files=[...(files||[])]; if(!files.length) return;
    if(!S.model){ status('Load the extract first — sentences are matched against its line names.', true); return; }
    const docs=[], bad=[];
    for(const file of files){
      status(`Reading ${file.name}…`);
      try{
        const {text,kind}=await readFile(file);
        docs.push({name:file.name, label:labelOf(file.name), kind, items:analyse(text, kind), scope:scopeStamp()});
      }catch(e){ bad.push(file.name+': '+(e&&e.message||e)); }
    }
    // then the model, if one is switched on: it decides, the rules stand if it cannot
    let aiNote='';
    if(aiCfg().source==='none'&&aiCfg().wanted&&docs.length){ status('Looking for a model on this desktop…'); await aiProbe(); }
    if(aiCfg().on&&docs.length){
      const c=aiCfg();
      status(`Asking ${c.model} at ${c.endpoint}…`);
      let decided=0, unsure=0, fail='';
      for(const d of docs){
        try{ const r=await aiReview(d); if(r){ decided+=r.applied; unsure+=r.unsure; } }
        catch(e){ fail=aiFailWhy(e,c); break; }
      }
      aiNote=fail? ` · AI review unavailable (${fail}), the rules' decisions stand`
                 : ` · ${c.model} decided ${decided}${unsure? `, left ${unsure} unsure for review` : ''}`;
    }
    if(!docs.length){ status('Could not read '+bad.join(' · '), true); return; }
    // the sentences enrich the commentary now; review is a link away
    let on=0, loose=0;
    for(const d of docs){
      const i=EN.docs.findIndex(x=>x.name===d.name);
      if(i>=0) EN.docs[i]=d; else EN.docs.push(d);
      on+=counts(d).on; loose+=d.items.filter(x=>x.loose).length;
    }
    save(); applyChange();
    const who=docs.length===1? docs[0].name : docs.length+' files';
    const sl=scopeLabel(docs[0].scope);
    status((on? `${on} sentence${on===1?'':'s'} from ${who} ${on===1?'enriches':'enrich'} the ${sl? sl+' ' : ''}commentary${sl? ' and roll up to '+(docs[0].scope.region&&docs[0].scope.country? docs[0].scope.region+' and the whole book' : 'the whole book') : ''}` : `Nothing in ${who} names a line of the loaded extract`)
      +(loose&&!aiNote? ` · ${loose} naming no line kept unticked` : '')
      +aiNote
      +(bad.length? ` · skipped ${bad.join(' · ')}` : '')+'.', !!bad.length&&!on, docs);
  }
  function mount(){
    if($id('enrichPanel')) return true;
    const tp=$id('tplPanel'); if(!tp||typeof S==='undefined'||typeof comInLoaded!=='function') return false;
    const st=document.createElement('style');
    st.textContent=`
#enrichPanel{padding:2px 16px 12px;display:flex;flex-direction:column;gap:7px;font-size:12px}
#enrichPanel label{display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--text-muted);margin:0 2px}
#enrichPanel .en-doc{display:flex;flex-wrap:wrap;gap:4px 6px;align-items:baseline;font-size:11.5px;padding:4px 0;border-top:1px solid var(--border)}
#enrichPanel .en-name{font-weight:600;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#enrichPanel .linklike{background:none;border:none;color:var(--hsbc-red);cursor:pointer;font:inherit;font-size:11.5px;padding:0}
#enrichPanel .linklike:hover{text-decoration:underline}
#enrichAdv summary{cursor:pointer;font-size:11.5px;font-weight:600;color:var(--text-muted)}
#enrichAdv{border-top:1px solid var(--border);padding-top:6px;display:flex;flex-direction:column;gap:6px}
#enrichAdv input[type=text],#enrichAdv input[type=password],#enrichAdv input[type=number]{flex:1;min-width:0;border:1px solid var(--border);border-radius:6px;padding:4px 7px;font:inherit;font-size:12px;background:var(--surface-1);color:var(--text-primary)}
#enrichAiLine .linklike{background:none;border:none;color:var(--hsbc-red);cursor:pointer;font:inherit;font-size:11.5px;padding:0}
.aicom-e{background:var(--fold-bg);border:none;color:var(--fold-red);cursor:pointer;font:inherit;font-size:10px;font-weight:700;padding:2px 9px;border-radius:5px;letter-spacing:.5px;text-transform:uppercase}
.aicom-e:hover{background:var(--fold-bg-h)}
.en-drag{outline:2px dashed var(--hsbc-red);outline-offset:-2px}
.en-sources{font-size:11.5px;color:var(--text-muted);margin:6px 0 0}
#enOverlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;display:flex;align-items:center;justify-content:center;padding:24px}
#enBox{background:var(--surface-1);color:var(--text-primary);border:1px solid var(--border);border-radius:10px;max-width:1140px;width:100%;
  max-height:92vh;display:flex;flex-direction:column;box-shadow:0 12px 40px rgba(0,0,0,.35)}
#enHead{padding:16px 20px 10px;border-bottom:1px solid var(--border)}
#enHead h3{margin:0 0 4px;font-size:15px}
#enHead .subtitle{margin:0 0 8px}
#enCtl{display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:12px}
#enCtl .btn-quiet{padding:4px 10px;font-size:11.5px}
#enBody{overflow:auto;padding:0 20px}
table.entab{width:100%;border-collapse:collapse;font-size:12.5px}
table.entab th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.6px;color:var(--text-muted);padding:8px 8px 6px;position:sticky;top:0;background:var(--surface-1);border-bottom:1px solid var(--border-strong);z-index:1}
table.entab td{padding:7px 8px;border-bottom:1px solid var(--border);vertical-align:top}
table.entab td.sent{min-width:340px;line-height:1.4}
table.entab td.muted{color:var(--text-muted);font-size:11.5px}
table.entab select{max-width:220px;padding:4px 6px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface-2);color:var(--text-primary);font:inherit;font-size:12px}
table.entab tr.en-off td{opacity:.5}
table.entab tr.en-file td{background:var(--surface-0);padding:9px 8px;font-size:12px}
table.entab tr.en-file label{display:inline-flex;gap:6px;align-items:center;margin-left:14px;color:var(--text-muted);font-size:11.5px}
table.entab tr.en-file input{padding:3px 6px;border:1px solid var(--border-strong);border-radius:4px;background:var(--surface-2);color:var(--text-primary);font:inherit;font-size:12px}
.en-nocue{font-size:10px;color:var(--text-muted);border:1px solid var(--border);border-radius:10px;padding:0 6px;margin-left:4px;white-space:nowrap}
#enFoot{padding:12px 20px;border-top:1px solid var(--border);display:flex;gap:10px;align-items:center;flex-wrap:wrap}
#enFoot .btn-primary{padding:7px 18px;font-size:13px}`;
    document.head.appendChild(st);
    const row=document.createElement('div');
    row.className='sect sect-fold'; row.id='sectEnrichRow'; row.title='Collapse / expand';
    row.style.cssText='display:flex;align-items:center;gap:6px;cursor:pointer';
    row.innerHTML='<span data-lbl="sect_enrich">Enrich commentary</span><span id="enrichCaret" style="margin-left:auto;font-size:10px;color:inherit">▾</span>';
    const panel=document.createElement('div'); panel.id='enrichPanel';
    panel.innerHTML=`
      <button class="btn-quiet" id="btnEnrichUpload">⤒ Add sources (docx / txt / vtt · one or many)</button>
      <input type="file" id="enrichInput" accept=".docx,.txt,.text,.vtt,.md,.pdf,.doc" multiple hidden>
      <div id="enrichList"></div>
      <div style="display:flex;gap:10px"><button class="linklike" id="btnEnrichReviewAll" hidden>Review all</button>
        <button class="linklike" id="btnEnrichForgetAll" hidden>Forget all sources</button></div>
      <div class="note" id="enrichStatus"></div>
      <div class="note" id="enrichAiLine"></div>
      <details id="enrichAdv"><summary>Advanced</summary>
        <label><input type="checkbox" id="enrichMark"${EN.mark?' checked':''}> also mark the source on each clause</label>
        <label>Model endpoint <input type="text" id="aiEndpoint" placeholder="blank = the model found on this desktop"></label>
        <label>Model <input type="text" id="aiModel" placeholder="blank = as found"></label>
        <label>Key <input type="password" id="aiKey" placeholder="only a gateway or Claude needs one"></label>
        <label>Stitch when at least <input type="number" id="aiThreshold" min="0" max="1" step="0.05" style="width:64px"> sure</label>
        <label><input type="checkbox" id="aiRewrite"> reword sentences into pack style</label>
        <div style="display:flex;gap:8px;align-items:center"><button class="btn-quiet" id="btnAiTest">Test the model</button><span class="note" id="enrichAiStatus"></span></div>
        <div class="note">A desktop model (Ollama, LM Studio) keeps everything on this machine; Ollama must be started with OLLAMA_ORIGINS=* so a page opened from a file may call it. An endpoint typed here is used instead — Claude when it is the Anthropic Messages API, OpenAI-style chat otherwise.</div>
      </details>`;
    tp.insertAdjacentElement('afterend',row);
    row.insertAdjacentElement('afterend',panel);
    let folded=false; try{ folded=localStorage.getItem('iwpbsg_enrich_collapsed')==='1'; }catch(e){}
    const fold=f=>{ panel.hidden=f; $id('enrichCaret').textContent=f?'▸':'▾'; try{ localStorage.setItem('iwpbsg_enrich_collapsed',f?'1':'0'); }catch(e){} };
    fold(folded);
    row.addEventListener('click',()=>fold(!panel.hidden));
    const fi=$id('enrichInput');
    $id('btnEnrichUpload').addEventListener('click',()=>fi.click());
    fi.addEventListener('change',()=>{ const fs=[...fi.files]; fi.value=''; if(fs.length) handleFiles(fs); });
    $id('btnEnrichForgetAll').addEventListener('click',()=>{ EN.docs=[]; save(); applyChange(); status('All sources forgotten.'); });
    $id('btnEnrichReviewAll').addEventListener('click',()=>openReview(EN.docs));
    // files can be dropped on the panel, or on the commentary itself
    const dropOn=el=>{ if(!el) return;
      ['dragenter','dragover'].forEach(ev=>el.addEventListener(ev,e=>{ if(e.dataTransfer&&[...e.dataTransfer.types].includes('Files')){ e.preventDefault(); el.classList.add('en-drag'); } }));
      ['dragleave','drop'].forEach(ev=>el.addEventListener(ev,e=>{ el.classList.remove('en-drag'); }));
      el.addEventListener('drop',e=>{ const fs=e.dataTransfer&&e.dataTransfer.files; if(fs&&fs.length){ e.preventDefault(); handleFiles(fs); } });
    };
    dropOn(panel); dropOn($id('housePanel'));
    paintList();
    return true;
  }

  /* ---- hooks: the uploads are the written commentary; a clause names its source */
  // the config pack's Commentary sheet is no longer read for narrative: every
  // path that merged it (comInAll, comInAllUnscoped, comInLoaded) now serves
  // the accepted sentences alone, with the same scope-following and caching
  comInLoaded=function(){ return accepted(); };
  comInAllUnscoped=function(){ return accepted({all:true}).filter(c=>String(c.text||'').trim()); };
  let EN_CACHE=null, EN_M=null, EN_K='';
  comInAll=function(){
    const m=S&&S.model, src=accepted();
    const k=Object.entries(S.filters||{}).filter(([,set])=>set&&set.size)
        .map(([d,set])=>d+':'+[...set].sort().join('|')).sort().join(';')
      +'\u0000'+src.map(c=>c.text||'').join('\u0001');
    if(EN_CACHE&&EN_M===m&&EN_K===k) return EN_CACHE;
    let out=src.filter(c=>String(c.text||'').trim());
    try{ out=out.filter(c=>comScopeOK(c.text)); }catch(e){}
    EN_CACHE=out; EN_M=m; EN_K=k;
    return out;
  };
  if(typeof comDigest==='function'){
    const _digest=comDigest; DIGEST=_digest;
    comDigest=function(txt,line,cap){
      const n=ROLL.get(txt);
      if(n&&cap==null){ let cm=150; try{ cm=comClauseMax(); }catch(e){} cap=cm*n; }
      const d=_digest.call(this, txt, line, cap);
      if(d&&EN.mark&&SRC.has(txt)){ const tag='('+SRC.get(txt)+')'; if(d.indexOf(tag)<0) return d+' '+tag; }
      return d;
    };
  }
  window.__enrichAnalyse=function(text, kind, name){ return {name:name||'pasted', label:labelOf(name||'pasted'), kind:kind||'doc', items:analyse(text,kind||'doc')}; };
  window.__enrichOpen=openReview;
  window.__enrichState=EN;
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mount); else mount();
  setTimeout(noteHook,0);
})();
