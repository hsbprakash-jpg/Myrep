const http=require("http");
const seen=[];
function decide(body, reword){
  const sents=body.sentences||[];
  return sents.map(s=>{
    const t=s.text;
    if(/see my screen|next slide/i.test(t)) return {i:s.i,use:false,line:'',confidence:0.95,reason:'logistics, not a driver',text:t};
    if(/Insurance fees helped/i.test(t)) return {i:s.i,use:true,line:'Insurance fees',confidence:0.45,reason:'weak driver, no cause given',text:t};
    if(/It's driven by the deposit margin/i.test(t)) return {i:s.i,use:true,line:'Net interest income',confidence:0.9,reason:'continues NII sentence',text:reword?'NII was driven by wider deposit margins as rates stayed higher and savings balances grew.':t};
    const use=/driven|because|due to|offset|timing|one-off|below forecast|ahead of|higher than|rose|savings/i.test(t);
    return {i:s.i,use,line:s.rules_line||'',confidence:use?0.88:0.8,reason:use?'explains a movement':'no cause given',text:t};
  });
}
http.createServer((req,res)=>{
  const cors={'Access-Control-Allow-Origin':'*','Access-Control-Allow-Headers':'*','Access-Control-Allow-Methods':'POST, OPTIONS'};
  if(req.method==='OPTIONS'){ res.writeHead(204,cors); return res.end(); }
  let b=''; req.on('data',c=>b+=c); req.on('end',()=>{
    let j={}; try{ j=JSON.parse(b); }catch(e){}
    seen.push({url:req.url, headers:req.headers, body:j});
    let userTxt='';
    if(req.url.startsWith('/v1/messages')) userTxt=(j.messages||[]).map(m=>typeof m.content==='string'?m.content:'').join('');
    else userTxt=(j.messages||[]).filter(m=>m.role==='user').map(m=>m.content).join('');
    let u={}; try{ u=JSON.parse(userTxt); }catch(e){}
    const out=JSON.stringify({decisions:decide(u, !!u.reword)});
    res.writeHead(200,{...cors,'content-type':'application/json'});
    if(req.url.startsWith('/v1/messages')) res.end(JSON.stringify({id:'msg_1',type:'message',role:'assistant',stop_reason:'end_turn',content:[{type:'text',text:out}]}));
    else res.end(JSON.stringify({choices:[{message:{role:'assistant',content:out}}]}));
  });
}).listen(11499,()=>console.log("mock llm on 11499"));
process.on('SIGTERM',()=>{ require('fs').writeFileSync('mock-seen.json',JSON.stringify(seen,null,1)); process.exit(0); });
