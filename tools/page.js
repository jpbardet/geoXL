
'use strict';
let library=null,file=null,busy=false,url=null;
const $=id=>document.getElementById(id);
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function say(id,text,error=false){$(id).textContent=text;$(id).classList.toggle('error',error);}
function reset(){if(url)URL.revokeObjectURL(url);url=null;$('download').classList.add('hidden');$('progress').innerHTML='';$('progress').classList.remove('error');}
function ready(){$('run').disabled=busy||!library||!file;$('workbook').disabled=busy;}
async function loadLibrary(){
  busy=true;library=null;reset();ready();$('load').classList.add('hidden');$('dot').className='dot';
  say('libraryStatus','Loading the current __LABEL__ version…');
  try{
    const r=await fetch('./library/current.json',{cache:'no-store',credentials:'omit'});
    if(!r.ok)throw Error('Could not reach the __LABEL__ release record.');
    const release=await r.json();
    if(!/^releases\/[A-Za-z0-9._-]+\/library\.txt$/.test(release.library))throw Error('Invalid published library path.');
    const response=await fetch('./library/'+release.library,{cache:'no-store',credentials:'omit'});
    if(!response.ok)throw Error('Could not download the __LABEL__ library.');
    const lib=await GeoUpdater.readLibrary(await response.text());
    if(lib.hash!==release.sha256||lib.version!==release.version)throw Error('The downloaded library failed its integrity check.');
    library=lib;$('dot').className='dot ok';
    say('libraryStatus','Current version: __LABEL__ '+lib.version+' · '+lib.all.length+' functions · verified');
  }catch(e){$('dot').className='dot err';say('libraryStatus',e.message+' Check your internet connection and try again.',true);$('load').classList.remove('hidden');}
  finally{busy=false;ready();}
}
$('load').onclick=loadLibrary;
$('workbook').onchange=()=>{reset();file=null;const files=Array.from($('workbook').files);
  if(files.length===1&&/\.xlsx$/i.test(files[0].name)&&!files[0].name.startsWith('~$')){file=files[0];say('selection','Selected: '+file.name);}
  else if(files.length===1&&/\.(xlsm|xls|xlsb)$/i.test(files[0].name))say('selection','This is not an .xlsx workbook. Open it in Excel, choose File › Save As › Excel Workbook (.xlsx), and choose the new file.',true);
  else say('selection','Choose one .xlsx workbook.',true);
  ready();};
async function dottedNames(bytes){
  try{const z=await JSZip.loadAsync(bytes);const w=await z.file('xl/workbook.xml').async('string');
    const doc=new DOMParser().parseFromString(w,'application/xml');
    const nodes=Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/spreadsheetml/2006/main','definedName'));
    const aliases=GeoUpdater.dottedAliases(library,nodes);const seen=new Map();
    for(const n of nodes){const name=n.getAttribute('name'),t=aliases.get(name.toLowerCase());if(t&&!seen.has(name.toLowerCase()))seen.set(name.toLowerCase(),[name,t]);}
    return [...seen.values()];}catch(_){return [];}
}
$('run').onclick=async()=>{if(busy||!library||!file)return;busy=true;ready();reset();
  try{
    say('progress','Working on '+file.name+'…');
    const bytes=await file.arrayBuffer();
    const renamed=await dottedNames(bytes);
    const r=await GeoUpdater.updateWorkbook(bytes,file.name,library);
    const out=file.name.replace(/\.xlsx$/i,'')+' (__LABEL__ '+library.version+').xlsx';
    url=URL.createObjectURL(new Blob([r.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    $('download').href=url;$('download').download=out;$('download').textContent='Download '+out;$('download').classList.remove('hidden');
    const c=r.report.counts,prev=r.report.previousVersion&&r.report.previousVersion!=='unknown'?' (previously '+esc(r.report.previousVersion)+')':'';
    const what=c.updated?('updated'+prev+': '+c.updated+' functions updated, '+c.added+' added'):('installed: '+c.added+' functions added');
    let h='<div class=result><b>Done.</b> Your original workbook <b>'+esc(file.name)+'</b> has not been changed. The download is a new copy with <b>__LABEL__ '+esc(library.version)+'</b> '+what+'.</div>';
    if(renamed.length){const list='<ul>'+renamed.map(([a,b])=>'<li><code>'+esc(a)+'</code> → <code>'+esc(b)+'</code></li>').join('')+'</ul>',head=renamed.length+' older name'+(renamed.length>1?'s':'')+' converted to the current names, and the formulas using them updated';
      h+='<div class=note>'+(renamed.length>5?'<details><summary>'+head+'</summary>'+list+'</details>':head+':'+list)+'</div>';}
    const notes=r.report.warnings.filter(w=>!/version changed from a text value/i.test(w));
    if(notes.length)h+='<div class=note>'+notes.map(esc).join('<br>')+'</div>';
    h+='<p class=small>Open the downloaded workbook in Excel; it recalculates when it opens. Type <code>=__VERSIONFN__()</code> in any cell to check the installed __LABEL__ version.</p>';
    $('progress').innerHTML=h;
  }catch(e){say('progress',e.message+' Your original workbook was not changed.',true);}
  finally{busy=false;ready();}};
const q=$('q');
q.addEventListener('input',()=>{const t=q.value.trim().toLowerCase();
  document.querySelectorAll('.fn').forEach(s=>s.classList.toggle('hidden',!!t&&!s.textContent.toLowerCase().includes(t)));
  $('resultcount').textContent=(t?document.querySelectorAll('.fn:not(.hidden)').length+' of ':'')+'__TOTAL__ functions';
  document.querySelectorAll('main h2').forEach(h=>{let n=h.nextElementSibling,any=false;while(n&&n.tagName!=='H2'){if(!n.classList.contains('hidden'))any=true;n=n.nextElementSibling;}h.classList.toggle('hidden',!any);});
  document.querySelectorAll('nav .fnlink').forEach(a=>a.classList.toggle('hidden',$(a.dataset.fn).classList.contains('hidden')));
  document.querySelectorAll('nav details').forEach(d=>{const shown=d.querySelectorAll('.fnlink:not(.hidden)').length;d.classList.toggle('hidden',!!t&&!shown);if(t)d.open=shown>0;d.querySelector('.count').textContent=t?shown:d.querySelectorAll('.fnlink').length;});});
$('expandAll').onclick=()=>document.querySelectorAll('nav details').forEach(d=>d.open=true);
$('collapseAll').onclick=()=>document.querySelectorAll('nav details').forEach(d=>d.open=false);
function mark(id){document.querySelectorAll('nav .fnlink.current').forEach(a=>a.classList.remove('current'));
  const a=document.querySelector('nav .fnlink[data-fn="'+id+'"]');if(a){a.classList.add('current');a.closest('details').open=true;}
  const g=document.querySelector('nav details[data-group="'+id+'"]');if(g)g.open=true;}
const PANELS=['install','guide','reference'];
function show(which,anchor){
  for(const p of PANELS){$(p).classList.toggle('hidden',p!==which);$('tab-'+p).setAttribute('aria-selected',String(p===which));}
  if(anchor){const el=document.getElementById(anchor);if(el)requestAnimationFrame(()=>el.scrollIntoView());if(which==='reference')mark(anchor);}}
function route(){const h=decodeURIComponent(location.hash.slice(1)),el=h&&document.getElementById(h);
  if(!h||PANELS.includes(h)){show(h||'install');return;}
  for(const p of PANELS)if(el&&$(p).contains(el)){show(p,h);return;}}
for(const p of PANELS)$('tab-'+p).onclick=()=>{history.replaceState(null,'','#'+p);show(p);window.scrollTo(0,0);};
document.querySelectorAll('.tab').forEach((t,i,all)=>t.addEventListener('keydown',e=>{if(e.key==='ArrowRight'||e.key==='ArrowLeft'){const n=all[(i+(e.key==='ArrowRight'?1:all.length-1))%all.length];n.focus();n.click();}}));
window.addEventListener('hashchange',route);route();
loadLibrary();
