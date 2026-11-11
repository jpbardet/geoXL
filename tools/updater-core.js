/* geoXL / geolabXL workbook updater. The page sets window.GeoConfig before loading this script. No formula evaluation or executable remote code. */
(function (root) {
'use strict';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CFG = root.GeoConfig || {};
const INFO = CFG.info || '_GeoEXCEL_UpdateInfo';
const FUTURE = new Set(('LAMBDA LET ISOMITTED IFS SWITCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY XLOOKUP XMATCH XMATCH BYROW BYCOL MAP REDUCE SCAN MAKEARRAY VSTACK HSTACK TOCOL TOROW TAKE DROP CHOOSECOLS CHOOSEROWS EXPAND WRAPROWS WRAPCOLS TEXTSPLIT TEXTBEFORE TEXTAFTER TEXTJOIN CONCAT MAXIFS MINIFS SINGLE ANCHORARRAY STDEV.S STDEV.P VAR.S VAR.P RANK.EQ').split(' ')); // Excel 2010 and later functions are stored with the _xlfn. prefix
FUTURE.add('COT');
const FLOW = new Set(['FILTER','SORT','SORTBY','UNIQUE']);
const BUILTINS = new Set(('ABS ACOS ACOSH ASIN ASINH ATAN ATAN2 ATANH AVERAGE AVERAGEIF AVERAGEIFS BESSELJ CEILING CHAR CHOOSE CLEAN CODE COLUMN COLUMNS COS COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DEGREES ERF ERROR.TYPE EVEN EXACT EXP FACT FALSE FIND FLOOR GEOMEAN HARMEAN IF IFERROR IFNA INDEX INDIRECT INT INTERCEPT ISBLANK ISERR ISERROR ISLOGICAL ISNA ISNUMBER ISTEXT LARGE LEFT LEN LN LOG LOG10 LOOKUP LOWER MATCH MAX MEDIAN MID MIN MOD MROUND N NA NOT ODD OFFSET OR PI POWER PRODUCT PROPER QUOTIENT RADIANS RAND RANK REPLACE REPT RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SIGN SIN SINH SLOPE SMALL SQRT STDEV SUBSTITUTE SUM SUMIF SUMIFS SUMPRODUCT SUMSQ TAN TANH TEXT TRANSPOSE TRIM TRUE TRUNC UPPER VALUE VAR AND LINEST').split(' '));
BUILTINS.add('MMULT'); // Stream-function integration uses ordinary Excel matrix multiplication.
function fail(s) { throw new Error(s); }
function tokenize(s) {
  const out=[]; let i=0;
  while(i<s.length) {
    if (/\s/.test(s[i])) { i++; continue; }
    if(s.slice(i,i+2)==='//') { const n=s.indexOf('\n',i); i=n<0?s.length:n; continue; }
    if(s.slice(i,i+2)==='/*') { const n=s.indexOf('*/',i+2); if(n<0) fail('Unclosed comment.'); i=n+2; continue; }
    const m=s.slice(i).match(/^("(?:[^"]|"")*"|#[A-Za-z0-9/]+[!?]?|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|[A-Za-z_][\w.]*|<=|>=|<>|[=&<>+*/^(),;:{}\[\]\-])/);
    if(!m) fail('Unsupported source syntax near: '+s.slice(i,i+60));
    out.push({v:m[1],start:i,end:i+m[1].length}); i+=m[1].length;
  }
  return out;
}
class Parser {
  constructor(tokens) { this.t=tokens; this.i=0; }
  peek() { return this.t[this.i]?.v || ''; }
  pop() { return this.t[this.i++]?.v || ''; }
  want(v) { if(this.pop()!==v) fail('Expected '+v+' near token '+this.i+'. Check parentheses and commas.'); }
  expression(min=0) {
    const t=this.pop(); if(!t) fail('Incomplete expression.'); let n;
    if(t==='(') { n={k:'group',a:this.expression()}; this.want(')'); }
    else if(t==='[') { const v=this.pop(); this.want(']'); n={k:'optional',v}; }
    else if(t==='-'||t==='+') n={k:'unary',v:t,a:this.expression(4)};
    else if(t==='{') {
      const rows=[]; let row=[];
      while(true) { row.push(this.expression()); const sep=this.pop(); if(sep===',') continue;
        rows.push(row); row=[]; if(sep==='}') break; if(sep!==';') fail('Invalid array constant.'); }
      n={k:'array',rows};
    } else if(t.startsWith('"')||t.startsWith('#')||/^[\d.]/.test(t)) n={k:'literal',v:t};
    else if(/^[A-Za-z_]/.test(t)) {
      n={k:'name',v:t};
      if(this.peek()==='(') { this.pop(); const args=[];
        if(this.peek()!==')') while(true) {
          args.push([',',')'].includes(this.peek())?{k:'omitted'}:this.expression());
          if(this.peek()!==',') break; this.pop();
        }
        this.want(')'); n={k:'call',v:t,args};
      }
    } else fail('Unexpected token '+t+'.');
    const prec={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':2,'-':2,'*':3,'/':3,'^':4,':':5};
    while(this.peek() in prec && prec[this.peek()]>=min) { const op=this.pop(),p=prec[op];
      n={k:'binary',v:op,a:n,b:this.expression(op==='^'?p:p+1)}; }
    return n;
  }
}
function declarations(source) {
  const p=new Parser(tokenize(source)), result=[];
  while(p.peek()) {
    const start=p.t[p.i].start,name=p.pop();
    if(!/^[A-Za-z_][\w.]*$/.test(name)) fail('Expected a definition name.');
    p.want('='); const formulaStart=p.t[p.i]?.start; const ast=p.expression();
    const formulaEnd=p.t[p.i-1].end; p.want(';'); const end=p.t[p.i-1].end;
    result.push({name,ast,start,end,formula:source.slice(formulaStart,formulaEnd)});
  }
  return result;
}
function localName(n) {
  if(!['name','optional'].includes(n.k)||!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n.v)) fail('Invalid local parameter name.');
  if(/^[rc]$/i.test(n.v)) fail('Local name '+n.v+' is reserved by Excel. Rename it before updating.');
  return n.v;
}
function compile(ast,defs,module,env=new Map()) {
  const go=(n,e=env)=>compile(n,defs,module,e);
  if(ast.k==='literal') return ast.v;
  if(ast.k==='omitted') return '';
  if(ast.k==='group') return '('+go(ast.a)+')';
  if(ast.k==='binary') return go(ast.a)+ast.v+go(ast.b);
  if(ast.k==='unary') return ast.v+go(ast.a);
  if(ast.k==='array') return '{'+ast.rows.map(row=>row.map(n=>go(n)).join(',')).join(';')+'}';
  if(ast.k==='name'||ast.k==='optional') {
    const key=ast.v.toLowerCase();
    if(env.has(key)) return '_xlpm.'+env.get(key);
    if(['true','false'].includes(key)) return key.toUpperCase();
    if(defs.has(key)) return (module?module+'.':'')+defs.get(key).name;
    fail('Unbound source name: '+ast.v+'. Pass worksheet data as an argument.');
  }
  const name=ast.v.toUpperCase(),args=ast.args;
  if(name==='LAMBDA') {
    if(args.length<1) fail('LAMBDA needs a body.');
    const inner=new Map(env),params=[];let optional=false;
    for(const a of args.slice(0,-1)) { const v=localName(a),key=v.toLowerCase();
      if(inner.has(key)&&params.some(p=>p.toLowerCase().endsWith('.'+key))) fail('Duplicate LAMBDA parameter '+v);
      if(a.k==='optional') optional=true; else if(optional) fail('Required parameter follows an optional parameter.');
      inner.set(key,v); params.push((a.k==='optional'?'_xlop.':'_xlpm.')+v); }
    return '_xlfn.LAMBDA('+params.concat(go(args.at(-1),inner)).join(',')+')';
  }
  if(name==='LET') {
    if(args.length<3||args.length%2!==1) fail('LET must contain name/value pairs and a result.');
    const inner=new Map(env),items=[];
    for(let i=0;i<args.length-1;i+=2) {
      const v=localName(args[i]); const value=go(args[i+1],inner);
      inner.set(v.toLowerCase(),v);items.push('_xlpm.'+v,value);
    }
    return '_xlfn.LET('+items.concat(go(args.at(-1),inner)).join(',')+')';
  }
  const key=ast.v.toLowerCase();let target;
  if(env.has(key)) target='_xlpm.'+env.get(key);
  else if(defs.has(key)) {
    const d=defs.get(key),params=d.ast.args.slice(0,-1),required=params.filter(a=>a.k!=='optional').length;
    if(args.length<required||args.length>params.length) fail('Wrong number of arguments calling '+d.name);
    target=(module?module+'.':'')+d.name;
  } else if(FUTURE.has(name)) target='_xlfn.'+(FLOW.has(name)?'_xlws.':'')+name;
  else if(BUILTINS.has(name)) target=name;
  else fail('Unsupported Excel function '+ast.v+'. The updater needs a compatibility entry before this release can be installed.');
  return target+'('+args.map(a=>go(a)).join(',')+')';
}
// Excel opens a package only when every entry is a part with a declared content type (Default by extension
// or Override by name). Refuse to write anything else, so that a damaged copy is never handed to a user.
async function checkPackage(zip){
  const names=Object.keys(zip.files);
  const dirs=names.filter(n=>zip.files[n].dir||n.endsWith('/'));
  if(dirs.length)fail('Package check failed: folder entries '+dirs.join(', '));
  if(!zip.file('[Content_Types].xml'))fail('Package check failed: [Content_Types].xml is missing');
  const ct=xml(await zip.file('[Content_Types].xml').async('string'));
  const defaults=new Set(Array.from(ct.getElementsByTagName('Default')).map(e=>(e.getAttribute('Extension')||'').toLowerCase()));
  const overrides=new Set(Array.from(ct.getElementsByTagName('Override')).map(e=>(e.getAttribute('PartName')||'').replace(/^\//,'').toLowerCase()));
  for(const n of names){
    if(n==='[Content_Types].xml')continue;
    const ext=(n.split('.').pop()||'').toLowerCase();
    if(!defaults.has(ext)&&!overrides.has(n.toLowerCase()))fail('Package check failed: no content type for '+n);
  }
}

async function readLibrary(source) {
  if(source.length>2000000) fail('Library is too large (limit 2 MB).');
  const all=declarations(source),defs=new Map();
  for(const d of all) {
    const constant=n=>n.k==='literal'||(n.k==='unary'&&constant(n.a))||(n.k==='array'&&n.rows.every(row=>row.length===n.rows[0].length&&row.every(constant)));
    if((d.ast.k!=='call'||d.ast.v.toUpperCase()!=='LAMBDA') && !constant(d.ast)) fail('Expected a LAMBDA, constant value or rectangular constant array: '+d.name);
    if(d.name.length>230) fail('Use local function names without a module prefix: '+d.name);
    if(defs.has(d.name.toLowerCase())) fail('Duplicate function '+d.name);
    defs.set(d.name.toLowerCase(),d);
    // Match only the adjacent documentation block, never across an earlier */.
    const prefix=source.slice(0,d.start),m=prefix.match(/\/\*\*((?:(?!\*\/)[\s\S])*)\*\/\s*$/);
    d.title=m?m[1].trim().replace(/\s+/g,' '):d.name;
  }
  if(!defs.size) fail('No LAMBDA definitions found.');
  const version=source.match(/(?:geoXL|GeoEXCEL|geolabXL)\s+version\s+([\w.\-]+)/i)?.[1]||'not labelled';
  for(const d of all) {
    d.compiled=compile(d.ast,defs,'');
    if(d.compiled.length>8192) fail(d.name+' exceeds 8,192 characters after Excel compatibility conversion.');
  }
  const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(source))),n=>n.toString(16).padStart(2,'0')).join('');
  return {source,defs,all,version,hash};
}
function xml(s) {
  if(/<!DOCTYPE|<!ENTITY/i.test(s)) fail('XML entities or document types are unsupported.');
  const d=new DOMParser().parseFromString(s,'application/xml');
  if(d.getElementsByTagName('parsererror').length) fail('Malformed workbook XML.'); return d;
}
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\r/g,'&#13;').replace(/\n/g,'&#10;'); }
function encodeFormula(s) { return s.replace(/_x[0-9a-f]{4}_/gi,m=>'_x005F_'+m.slice(1)); }
function decodeFormula(s) { return s.replace(/_x([0-9a-f]{4})_/gi,(_,h)=>String.fromCharCode(parseInt(h,16))); }
function findNamesBlock(w) {
  return w.match(/<((?:[\w.-]+:)?definedNames)\b(?:[^>"]|"[^"]*")*\/>|<((?:[\w.-]+:)?definedNames)\b(?:[^>"]|"[^"]*")*>[^]*?<\/\2\s*>/);
}
function patchNames(w,entries,replaceNames) {
  const block=findNamesBlock(w),root=w.match(/<((?:[\w.-]+:)?workbook)\b/);if(!root)fail('Workbook XML root not found.');
  const prefix=root[1].includes(':')?root[1].split(':')[0]+':':'';
  let kept='';
  if(block) {
    const open=block[0].match(/^<(?:[^>"]|"[^"]*")+>/)[0];
    if(!open.endsWith('/>')) {
      kept=block[0].slice(open.length).replace(/<\/(?:[\w.-]+:)?definedNames\s*>$/,'');
      kept=kept.replace(/<(?:[\w.-]+:)?definedName\b(?:[^>"]|"[^"]*")*(?:\/>|>[^]*?<\/(?:[\w.-]+:)?definedName\s*>)/g,s=>{
        const e=xml('<wrap xmlns="'+NS+'">'+s.replace(/(<\/?)[\w.-]+:/g,'$1')+'</wrap>').documentElement.firstElementChild;
        return replaceNames.has((e.getAttribute('name')||'').toLowerCase()+'|'+(e.getAttribute('localSheetId')||''))?'':s;
      });
    }
  }
  const added=entries.map(e=>'<'+prefix+'definedName name="'+esc(e.name)+'"'+(e.local!==null?' localSheetId="'+esc(e.local)+'"':'')+' comment="'+esc(e.comment.slice(0,255))+'">'+esc(encodeFormula(e.formula))+'</'+prefix+'definedName>').join('');
  const replacement='<'+prefix+'definedNames>'+kept+added+'</'+prefix+'definedNames>';
  if(block) return w.slice(0,block.index)+replacement+w.slice(block.index+block[0].length);
  // ECMA workbook order: definedNames follows sheets/functionGroups/externalReferences,
  // and precedes calcPr, pivotCaches and all later optional elements.
  const after=w.match(/<\/(?:[\w.-]+:)?(?:externalReferences|functionGroups|sheets)\s*>/g);
  if(!after?.length) fail('Workbook sheets element not found.');
  const end=w.lastIndexOf(after.at(-1))+after.at(-1).length;
  return w.slice(0,end)+replacement+w.slice(end);
}
function patchRecalc(w) {
  const m=w.match(/<((?:[\w.-]+:)?calcPr)\b[^>]*(?:\/>|>[^]*?<\/\1\s*>)/);
  if(m) return w.replace(m[0],m[0].replace(/\s(?:fullCalcOnLoad|forceFullCalc)="[^"]*"/g,'').replace(/^(<[^\s/>]+)/,'$1 fullCalcOnLoad="1" forceFullCalc="1"'));
  const close=w.match(/<\/((?:[\w.-]+:)?definedNames)\s*>/);const prefix=close[1].includes(':')?close[1].split(':')[0]+':':'';
  const at=close.index+close[0].length;return w.slice(0,at)+'<'+prefix+'calcPr fullCalcOnLoad="1" forceFullCalc="1"/>'+w.slice(at);
}
function base64decode(s) {const b=atob(s.replace(/\s/g,''));return Uint8Array.from(b,c=>c.charCodeAt(0));}
function decodeProject(s) {
  const bytes=base64decode(s);
  for(const encoding of ['utf-16le','utf-8']) try {return {data:JSON.parse(new TextDecoder(encoding).decode(bytes)),encoding};} catch(_) {}
  fail('Excel Labs module data cannot be decoded.');
}
function encodeProject(data,encoding) {
  const s=JSON.stringify(data);let a;
  if(encoding==='utf-16le') {a=new Uint8Array(s.length*2);for(let i=0;i<s.length;i++){a[2*i]=s.charCodeAt(i)&255;a[2*i+1]=s.charCodeAt(i)>>8;}}
  else a=new TextEncoder().encode(s);
  let b='';for(let i=0;i<a.length;i+=8192)b+=String.fromCharCode(...a.subarray(i,i+8192));return btoa(b);
}
async function patchAFE(zip,lib,legacy,warnings,aliases=dottedAliases(lib)) {
  const changed=[];
  for(const path of Object.keys(zip.files).filter(n=>/^customXml\/item\d+\.xml$/i.test(n))) {
    const source=await zip.file(path).async('string');
    if(!source.includes('AFEJSONBlob')) continue;
    const doc=xml(source),node=doc.documentElement;
    if(node.localName!=='AFEJSONBlob')fail('Unrecognized Excel Labs record.');
    const {data,encoding}=decodeProject(node.textContent);
    if(data.schema!=='http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.2'||!Array.isArray(data.files))fail('Unsupported Excel Labs schema. Workbook left unchanged.');
    let touched=false;
    for(const file of data.files) {
      if(typeof file.path!=='string'||typeof file.text!=='string')fail('Unsupported Excel Labs file record.');
      const leaf=file.path.split('/').pop();
      if(['geoexcel','geoxl'].includes(leaf.toLowerCase())) {
        const old=declarations(file.text),obsolete=old.filter(d=>!lib.defs.has(d.name.toLowerCase())&&!aliases.has(d.name.toLowerCase()));
        file.text=lib.source+'\n'+obsolete.map(d=>'// Retained obsolete definition; review before removal.\n'+d.name+' = '+d.formula+';').join('\n');
        touched=true;
        if(obsolete.length)warnings.push('Retained '+obsolete.length+' older definitions in the Excel Labs module.');
      } else if(leaf==='Workbook') {
        // Only parse a Workbook module if it actually mentions a managed definition.
        const candidates=[...file.text.matchAll(/^\s*([A-Za-z_][\w.]*)\s*=/gm)].map(m=>m[1]);
        const owned=n=>aliases.has(n.toLowerCase())||/^geoexcel[._]/i.test(n)||lib.defs.has(n.toLowerCase())||(legacy&&lib.defs.has(n.toLowerCase().replace(/\./g,'_')));
        if(!candidates.some(owned))continue;
        let result=file.text;
        for(const d of declarations(file.text).reverse()) {
          if(aliases.has(d.name.toLowerCase())) {
            const target=aliases.get(d.name.toLowerCase());
            const replacement=lib.defs.has(target.toLowerCase())?'':target+' = '+migrateFormula(d.formula,aliases).formula+';';
            result=result.slice(0,d.start)+replacement+result.slice(d.end);touched=true;continue;
          }
          let key=lib.defs.has(d.name.toLowerCase())?d.name.toLowerCase():d.name.replace(/^geoexcel[._]/i,'').replace(/\./g,'_').toLowerCase();
          if(!owned(d.name)||!lib.defs.has(key))continue;
          // Workbook-module definitions need fully qualified internal library calls.
          const item=lib.defs.get(key),flat=new Map(lib.all.map(d=>[d.name.toLowerCase(),d]));
          let formula=compile(item.ast,flat,'').replace(/_xlfn\._xlws\.|_xlfn\.|_xlpm\./g,'').replace(/_xlop\.([A-Za-z_][\w]*)/g,'[$1]');
          result=result.slice(0,d.start)+d.name+' = '+formula+';'+result.slice(d.end);touched=true;
        }
        file.text=result;
      }
    }
    if(touched) {
      const payload=encodeProject(data,encoding);
      const match=source.match(/(<(?:[\w.-]+:)?AFEJSONBlob\b[^>]*>)[^]*?(<\/(?:[\w.-]+:)?AFEJSONBlob\s*>)/);
      if(!match)fail('Excel Labs XML payload not found.');
      zip.file(path,source.replace(match[0],match[1]+payload+match[2]),{createFolders:false});changed.push(path);
    }
  }
  return changed;
}
function dIsLambda(formula){return /^_xlfn\.LAMBDA\(/i.test(formula);}
// Older dotted functions whose arguments differ from the underscore function of the same name:
// they continue under a different current name with the old arguments.
const RENAMED={'spt.bearing':'SPT_bearing_direct','pmt.settlement':'PMT_settlement_basic'};
// Calls already renamed by an earlier update but still written with the old arguments:
// name -> {argument count: current function with those arguments}.
const BY_ARGUMENTS={'spt_bearing':{4:'SPT_bearing_direct'},'pmt_settlement':{8:'PMT_settlement_basic'}};
// Rename only aliases that map unambiguously to a name in the selected library.
function dottedAliases(lib,existing=[]) {
  const aliases=new Map();
  for(const d of lib.all) {
    const alias=d.name.replace(/_/g,'.').toLowerCase();
    if(alias===d.name.toLowerCase()||lib.defs.has(alias))continue;
    if(aliases.has(alias)&&aliases.get(alias)!==d.name)fail('Ambiguous older alias: '+alias);
    aliases.set(alias,d.name);
  }
  for(const node of existing) {
    const name=node.getAttribute('name'),key=name.toLowerCase(),target=name.replace(/^(?:geoexcel|geoxl)\./i,'').replace(/\./g,'_');
    if(!name.includes('.')||/^_+(?:xlnm|xlchart|xlfn|xlpm|xlop)\./i.test(name))continue;
    if(lib.defs.has(key))fail('The selected library still defines a dotted name: '+name);
    const canonical=lib.all.find(d=>d.name.toLowerCase()===target.toLowerCase());
    if(canonical)aliases.set(key,canonical.name);
    else if(/^\s*=?\s*(?:_xlfn\.)?LAMBDA\s*\(/i.test(decodeFormula(node.textContent)))aliases.set(key,target);
  }
  // A defined older dotted function keeps its own arguments: map it to the current function with the same count.
  for(const node of existing) {
    const key=node.getAttribute('name').toLowerCase(),alt=RENAMED[key];
    if(!alt||!aliases.has(key))continue;
    const count=lambdaParameterCount(decodeFormula(node.textContent));
    const pick=[aliases.get(key),alt].map(n=>lib.defs.get(n.toLowerCase())).find(d=>d&&lambdaParameterCount(d.compiled)===count);
    if(pick)aliases.set(key,pick.name);
  }
  // 0.16 promotes the tested lateral-pile v2 definitions under their original names.
  // Migrate formulas before retiring comparison definitions, including scoped names.
  const promoted={
    lateralpile_cohesionless_reaction_shear_moment_v2:'lateralpile_cohesionless_reaction_shear_moment',
    lateralpile_cohesive_reaction_shear_moment_v2:'lateralpile_cohesive_reaction_shear_moment',
    lateralpile_integrate_v2:'lateralpile_integrate'
  };
  for(const [old,target] of Object.entries(promoted)) {
    if(!lib.defs.has(target.toLowerCase())||!lib.defs.has('lateralpile_integrate'))continue;
    for(const spelling of [old,old.replace(/_/g,'.')])aliases.set(spelling,target);
    for(const prefix of ['geoexcel.','geoxl.'])for(const spelling of [old,old.replace(/_/g,'.')])aliases.set(prefix+spelling,target);
    for(const node of existing) {
      const key=node.getAttribute('name').toLowerCase();
      if(key.replace(/^(?:geoexcel|geoxl)[._]/,'').replace(/\./g,'_')===old)aliases.set(key,target);
    }
  }
  // 0.28 renames the block stencils (seepage_confined_block, seepage_unconfined_block, seepage_relax).
  const renamedBlocks={seepage_stencil_block:'seepage_confined_block',seepage_stencil_unconfined_block:'seepage_unconfined_block',seepage_stencil_relax:'seepage_relax'};
  for(const [old,target] of Object.entries(renamedBlocks)) {
    if(!lib.defs.has(target)||lib.defs.has(old))continue;
    aliases.set(old,lib.defs.get(target).name);
  }
  aliases.libNames=new Map(lib.all.map(d=>[d.name.toLowerCase(),d.name]));
  aliases.argumentRules={};
  for(const [name,rule] of Object.entries(BY_ARGUMENTS)) {
    if(!lib.defs.has(name))continue;
    const ok=Object.fromEntries(Object.entries(rule).filter(([,t])=>lib.defs.has(t.toLowerCase())).map(([n,t])=>[n,lib.defs.get(t.toLowerCase()).name]));
    if(Object.keys(ok).length)aliases.argumentRules[name]=ok;
  }
  return aliases;
}
function lambdaParameterCount(formula) {
  const t=formulaTokens(formula.replace(/^\s*=\s*/,''));let i=0;
  if(!t.length||!/^(?:_xlfn\.)?LAMBDA$/i.test(t[0].v)||t[1]?.v!=='(')return -1;
  let depth=0,commas=0;
  for(i=1;i<t.length;i++){const v=t[i].v;if(v==='('||v==='{')depth++;else if(v===')'||v==='}'){depth--;if(!depth)break;}else if(v===','&&depth===1)commas++;}
  return commas;   // arguments before the body
}
function formulaTokens(s) {
  const tokens=[];let i=0;
  while(i<s.length) {
    const start=i,c=s[i];
    if(c==='"'||c==="'") {
      i++;let closed=false;
      while(i<s.length){if(s[i]===c){if(s[i+1]===c){i+=2;continue;}i++;closed=true;break;}i++;}
      if(!closed)fail('Unclosed quoted token in workbook formula.');
      tokens.push({v:s.slice(start,i),start,end:i,kind:'quoted'});continue;
    }
    if(c==='[') {
      let depth=1;i++;
      while(i<s.length&&depth){if(s[i]==="'"&&i+1<s.length){i+=2;continue;}if(s[i]==='[')depth++;if(s[i]===']')depth--;i++;}
      if(depth)fail('Unclosed bracket in workbook formula.');
      tokens.push({v:s.slice(start,i),start,end:i,kind:'bracket'});continue;
    }
    const number=s.slice(i).match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/);
    if(number){i+=number[0].length;tokens.push({v:number[0],start,end:i,kind:'number'});continue;}
    const name=s.slice(i).match(/^[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*/);
    if(name){i+=name[0].length;tokens.push({v:name[0],start,end:i,kind:'name'});continue;}
    i++;if(!/\s/.test(c))tokens.push({v:c,start,end:i,kind:'symbol'});
  }
  return tokens;
}
function migrateFormula(s,aliases) {
  const t=formulaTokens(s),scopes=[],stack=[],issues=[],changes=[],argumentRenames=[],moved=[],prefixedCalls=[];
  const normalized=v=>v.replace(/^_xlfn\./i,'').replace(/^_xlpm\./i,'').toLowerCase();
  for(let i=0;i<t.length;i++) {
    if(t[i].v==='(')stack.push({start:i,name:i?normalized(t[i-1].v):'',splits:[i+1],braces:0});
    else if(t[i].v==='{'){if(stack.length)stack.at(-1).braces++;}
    else if(t[i].v==='}'){if(stack.length)stack.at(-1).braces--;}
    else if(t[i].v===','&&stack.length&&!stack.at(-1).braces)stack.at(-1).splits.push(i+1);
    else if(t[i].v===')') {
      const frame=stack.pop();if(!frame)fail('Unmatched formula parenthesis.');
      const rule=(aliases.argumentRules||{})[frame.name.replace(/\./g,'_')];
      if(rule&&frame.start>0&&t[frame.start-1].kind==='name'&&t[frame.start-2]?.v!=='!') {
        const count=i===frame.start+1?0:frame.splits.length,target=rule[count];
        if(target&&target.toLowerCase()!==frame.name)argumentRenames.push({index:frame.start-1,target});
      }
      if(frame.name==='let'||frame.name==='lambda') {
        const names=new Set(),ends=[...frame.splits.slice(1).map(n=>n-1),i];
        const stride=frame.name==='let'?2:1;
        for(let a=0;a<frame.splits.length-1;a+=stride){const start=frame.splits[a];if(ends[a]===start+1&&t[start]?.kind==='name')names.add(normalized(t[start].v));}
        scopes.push({start:frame.start,end:i,names});
      }
    }
  }
  if(stack.length)fail('Unclosed formula parenthesis.');
  for(let i=0;i<t.length;i++) {
    const token=t[i];if(token.kind!=='name')continue;
    const key=token.v.replace(/^_xlfn\./i,'').toLowerCase();
    // GeoEXCEL.name (Excel Labs module), GEOXL.NAME / _xldudf_GEOXL_NAME (former JavaScript add-in)
    const prefixed=key.match(/^(?:_xldudf_(?:geoxl|geoexcel)_|(?:geoxl|geoexcel)\.)(.+)$/);
    const fromPrefix=prefixed&&t[i+1]?.v==='('?aliases.libNames?.get(prefixed[1].replace(/\./g,'_')):null;
    if(fromPrefix) {
      let start=token.start;
      if(t[i-1]?.v==='!'&&t[i-2]){start=t[i-2].start;if(t[i-3]?.kind==='bracket'&&t[i-3].end===t[i-2].start)start=t[i-3].start;}   // drop the FUNCRES.XLAM!/[1]! qualifier
      if(!scopes.some(scope=>scope.start<i&&i<scope.end&&scope.names.has(normalized(token.v)))){changes.push({start,end:token.end,target:fromPrefix});prefixedCalls.push(token.v.replace(/^_xldudf_(geoxl|geoexcel)_/i,(m,p)=>p.toUpperCase()+'.')+' → '+fromPrefix);}
      continue;
    }
    const target=aliases.get(key)||([...aliases.values()].find(n=>key.includes('.')&&n.toLowerCase()===key.replace(/\./g,'_')));
    if(!target||BUILTINS.has(token.v.toUpperCase()))continue;
    if(scopes.some(scope=>scope.start<i&&i<scope.end&&scope.names.has(normalized(token.v))))continue;
    if(t[i+1]?.v==='!')continue; // sheet/workbook qualifier, not a function name
    if(t[i-1]?.v==='!') {
      const qualifier=t[i-2];
      if((qualifier?.kind==='bracket'&&qualifier.v!=='[0]')||(qualifier?.kind==='quoted'&&qualifier.v.includes('['))) {
        issues.push('External workbook reference requires review: '+token.v);continue;
      }
    }
    changes.push({start:token.start,end:token.end,target});
  }
  if(t.some(token=>normalized(token.v)==='indirect')&&t.some(token=>token.kind==='quoted'&&[...aliases.keys()].some(alias=>token.v.toLowerCase().includes(alias))))issues.push('INDIRECT text contains an old name; cannot safely retire that definition.');
  for(const r of argumentRenames) {
    const token=t[r.index];
    if(scopes.some(scope=>scope.start<r.index&&r.index<scope.end&&scope.names.has(normalized(token.v))))continue;
    const k=changes.findIndex(c=>c.start===token.start);if(k>=0)changes.splice(k,1);   // the argument rule wins over a plain rename
    changes.push({start:token.start,end:token.end,target:r.target});moved.push(token.v+' → '+r.target);
  }
  changes.sort((a,b)=>a.start-b.start);
  let result=s;for(const c of changes.slice().reverse())result=result.slice(0,c.start)+c.target+result.slice(c.end);
  return {formula:result,count:changes.length,issues,moved,prefixed:prefixedCalls};
}
function migrateFormulaXml(source,tags,aliases,part,stats) {
  const alternatives=tags.join('|');
  const pattern=new RegExp('(<((?:[\\w.-]+:)?(?:'+alternatives+'))\\b(?!(?:[^>"]|"[^"]*")*\\/>)(?:[^>"]|"[^"]*")*>)([^]*?)(<\\/\\2\\s*>)','g');
  return source.replace(pattern,(whole,open,tag,body,close)=>{
    if(!body||body.includes('<'))return whole;
    // Decode XML entities with the same parser used by workbook validation.
    const formula=decodeFormula(xml('<v>'+body+'</v>').documentElement.textContent);
    const result=migrateFormula(formula,aliases);
    if(result.issues.length)stats.review.push({part,issues:result.issues});
    for(const m of result.moved||[])stats.moved.set(m,(stats.moved.get(m)||0)+1);
    for(const m of result.prefixed||[])stats.prefixed.set(m,(stats.prefixed.get(m)||0)+1);
    if(!result.count)return whole;
    stats.formulas++;stats.references+=result.count;stats.parts.add(part);
    return open+esc(encodeFormula(result.formula))+close;
  });
}
async function migrateDottedReferences(zip,lib,workbook,aliases=dottedAliases(lib)) {
  const stats={formulas:0,references:0,parts:new Set(),review:[],moved:new Map(),prefixed:new Map()};
  workbook=migrateFormulaXml(workbook,['definedName'],aliases,'xl/workbook.xml',stats);
  for(const path of Object.keys(zip.files)) {
    let tags;
    if(/^xl\/worksheets\/[^/]+\.xml$/i.test(path))tags=['f','formula','formula1','formula2'];
    else if(/^xl\/tables\/[^/]+\.xml$/i.test(path))tags=['calculatedColumnFormula','totalsRowFormula'];
    else if(/^xl\/charts\/[^/]+\.xml$/i.test(path))tags=['f'];
    else continue;
    const before=await zip.file(path).async('string'),after=migrateFormulaXml(before,tags,aliases,path,stats);
    if(after!==before){xml(after);zip.file(path,after,{createFolders:false});}
  }
  return {workbook,stats};
}

async function updateWorkbook(bytes,filename,lib,options={}) {
  if(!/\.(xlsx|xlsm)$/i.test(filename))fail('Only .xlsx and .xlsm files are supported.');
  const zip=await JSZip.loadAsync(bytes,{checkCRC32:true});
  if(Object.keys(zip.files).some(n=>/^_xmlsignatures\//i.test(n)))fail('Digitally signed workbook: updating would invalidate its signature.');
  if(!zip.file('xl/workbook.xml'))fail('Not a supported Excel workbook (encrypted and binary files cannot be updated).');
  const original=await zip.file('xl/workbook.xml').async('string'),doc=xml(original);
  if(doc.documentElement.namespaceURI!==NS)fail('Strict-format workbooks are not supported; save as standard .xlsx first.');
  const existing=Array.from(doc.getElementsByTagNameNS(NS,'definedName'));
  const versionNode=existing.find(n=>n.getAttribute('name')===INFO);
  let previous=null;try{previous=JSON.parse(decodeFormula(versionNode.textContent).slice(1,-1).replace(/""/g,'"'));}catch(_){}
  const entries=[],replace=new Set(),warnings=[],counts={added:0,updated:0,legacy:0,obsolete:0};
  // Each library stamps the workbooks it updates. A page may refuse workbooks stamped only by the other library.
  if(CFG.otherInfo){
    const hasOwn=existing.some(n=>n.getAttribute('name')===INFO),hasOther=existing.some(n=>n.getAttribute('name')===CFG.otherInfo);
    if(hasOther&&!hasOwn){if(CFG.refuseOther)fail(CFG.refuseMessage);else if(CFG.otherMessage)warnings.push(CFG.otherMessage);}
  }
  if(lib.defs.has('bearing_capacity_inclined')&&existing.some(n=>n.getAttribute('name').toLowerCase()==='bearing_inclinationfactors')) warnings.push('bearing_inclinationFactors was removed in geoXL 0.23. This copy keeps your existing definition unchanged; replace its calls with bearing_capacity_inclined(..., out = 2), which returns {ic; iq; iγ}.');
  if(lib.defs.has('bearing_capacity')&&existing.some(n=>n.getAttribute('name').toLowerCase()==='bearing_capacity' && /LAMBDA\s*\(\s*(?:_xlpm\.)?coh\s*,/i.test(decodeFormula(n.textContent)))) warnings.push('bearing_capacity has a new argument order in geoXL 0.22: author comes first, and inclined loads use bearing_capacity_inclined. Review existing bearing_capacity worksheet formulas before using the updated workbook.');
  const byName=new Map();for(const n of existing){const key=n.getAttribute('name').toLowerCase();if(!byName.has(key))byName.set(key,[]);byName.get(key).push(n);}
  const aliases=dottedAliases(lib,existing);
  function add(name,formula,title,n=null) {
    const local=n?.getAttribute('localSheetId')??null,key=name.toLowerCase()+'|'+(local||'');
    if(replace.has(key))return;
    if(n && dIsLambda(formula) && !/^\s*=?\s*(?:_xlfn\.)?LAMBDA\s*\(/i.test(decodeFormula(n.textContent))) {
      const oldVersion=name.toLowerCase()==='version'&&/^"(?:geoxl|geoexcel)\s+[0-9.]+/i.test(decodeFormula(n.textContent));
      if(!oldVersion)fail('Name conflict: '+name+' already exists and is not a LAMBDA.');
      warnings.push('Library version changed from a text value to a LAMBDA; use =version() to display it.');
    }
    replace.add(key);entries.push({name,formula,local,comment:title}); if(n)counts.updated++;else counts.added++;
  }
  for(const d of lib.all) {
    const target=d.name,nodes=byName.get(target.toLowerCase())||[];
    const global=nodes.find(n=>!n.hasAttribute('localSheetId'));
    add(target,d.compiled,d.title,global);
    for(const node of nodes.filter(n=>n.hasAttribute('localSheetId')))add(node.getAttribute('name'),d.compiled,d.title,node);
    // Prefix underscores were used by the previous workbook-ready distribution.
    for(const node of byName.get(('GeoEXCEL_'+d.name).toLowerCase())||[])add(node.getAttribute('name'),d.compiled,d.title,node);
  }
  counts.removedDottedNames=0;
  for(const node of existing) {
    const name=node.getAttribute('name'),target=aliases.get(name.toLowerCase());
    if(!target)continue;
    const local=node.getAttribute('localSheetId')??null;
    replace.add(name.toLowerCase()+'|'+(local||''));counts.removedDottedNames++;
    if(!entries.some(e=>e.name.toLowerCase()===target.toLowerCase()&&e.local===local)) {
      const conflict=(byName.get(target.toLowerCase())||[]).find(n=>(n.getAttribute('localSheetId')??null)===local);
      if(conflict&&!lib.defs.has(target.toLowerCase()))fail('Two custom definitions would merge: '+name+' and '+target);
      const definition=lib.defs.get(target.toLowerCase());
      entries.push({name:target,local,formula:definition?definition.compiled:migrateFormula(decodeFormula(node.textContent),aliases).formula,comment:definition?definition.title:(node.getAttribute('comment')||'Migrated older dotted LAMBDA')});
    }
  }
  const obsolete=existing.filter(n=>/^geoexcel[._]/i.test(n.getAttribute('name'))&&!replace.has(n.getAttribute('name').toLowerCase()+'|'+(n.getAttribute('localSheetId')||'')));
  counts.obsolete=obsolete.length;
  if(obsolete.length)warnings.push('Retained older names absent from this release: '+obsolete.map(n=>n.getAttribute('name')).join(', '));
  const info={version:lib.version,sha256:lib.hash,updated:new Date().toISOString()};
  replace.add(INFO.toLowerCase()+'|');entries.push({name:INFO,local:null,formula:'"'+JSON.stringify(info).replace(/"/g,'""')+'"',comment:(CFG.label||'GeoEXCEL')+' installed release and source hash.'});
  let workbook=patchRecalc(patchNames(original,entries,replace));
  const migration=await migrateDottedReferences(zip,lib,workbook,aliases);
  workbook=migration.workbook;xml(workbook);
  counts.renamedFormulas=migration.stats.formulas;counts.renamedReferences=migration.stats.references;
  if(migration.stats.prefixed.size){const total=[...migration.stats.prefixed.values()].reduce((a,b)=>a+b,0),names=[...new Set([...migration.stats.prefixed.keys()].map(k=>k.split(' → ')[1]))];
    warnings.push(total+' call'+(total>1?'s':'')+' with the old GeoEXCEL. / GEOXL. prefix converted to the plain names ('+names.slice(0,8).join(', ')+(names.length>8?', …':'')+').');}
  for(const [m,n] of migration.stats.moved)warnings.push(n+' call'+(n>1?'s':'')+' written with the older arguments renamed: '+m);
  if(migration.stats.review.length)fail('Cannot safely retire old names: '+migration.stats.review.map(r=>r.part+': '+r.issues.join('; ')).join(' | '));
  zip.file('xl/workbook.xml',workbook,{createFolders:false});
  const changed=[...new Set(['xl/workbook.xml',...migration.stats.parts,...await patchAFE(zip,lib,!!options.legacy,warnings,aliases)])];
  for(const n of Object.keys(zip.files))if(zip.files[n].dir)zip.remove(n); // an .xlsx package has no folder entries: Excel reports one as unreadable content
  const output=await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  const check=await JSZip.loadAsync(output,{checkCRC32:true});
  await checkPackage(check);
  const actual=Array.from(xml(await check.file('xl/workbook.xml').async('string')).getElementsByTagNameNS(NS,'definedName'));
  for(const node of actual)if(aliases.has(node.getAttribute('name').toLowerCase()))fail('Obsolete dotted definition remains: '+node.getAttribute('name'));
  for(const e of entries) {
    const found=actual.find(n=>n.getAttribute('name').toLowerCase()===e.name.toLowerCase()&&(n.getAttribute('localSheetId')??null)===e.local);
    if(!found||decodeFormula(found.textContent)!==e.formula)fail('Saved-name verification failed for '+e.name);
  }
  // Every unrelated ZIP entry is verified byte-for-byte, including worksheets and VBA.
  let preserved=0;
  const originalZip=await JSZip.loadAsync(bytes);
  for(const n of Object.keys(originalZip.files)) if(!originalZip.files[n].dir&&!changed.includes(n)) {
    const a=await originalZip.file(n).async('uint8array'),b=await check.file(n).async('uint8array');
    if(a.length!==b.length||!a.every((v,i)=>v===b[i]))fail('Preservation check failed: '+n);preserved++;
  }
  return {bytes:output,report:{file:filename,status:'updated',previousVersion:previous?.version||'unknown',version:lib.version,counts,formulaMigrationReview:migration.stats.review,changedParts:changed,unchangedPartsVerified:preserved,warnings:[...new Set(warnings)],nativeExcelVerified:false}};
}
function sourceURL(input) {
  const u=new URL(input);if(u.protocol!=='https:')fail('Use an HTTPS GitHub or Gist address.');
  if(u.hostname==='github.com') {const a=u.pathname.split('/');if(a[3]!=='blob')fail('Use a link to the library text file, not the repository homepage.');u.hostname='raw.githubusercontent.com';u.pathname='/'+a.slice(1,3).concat(a.slice(4)).join('/');u.search='';}
  if(!['raw.githubusercontent.com','gist.githubusercontent.com','gist.github.com','api.github.com'].includes(u.hostname))fail('Choose a GitHub/Gist source, or load a local text file.');
  if(u.hostname==='gist.github.com') {const id=u.pathname.split('/').filter(Boolean).at(-1);if(!/^[a-f0-9]{20,}$/i.test(id))fail('Invalid Gist address.');return 'https://api.github.com/gists/'+id;}
  return u.href;
}
async function fetchLibrary(input,chooseFile) {
  const url=sourceURL(input);let response;
  try {response=await fetch(url,{cache:'no-store',credentials:'omit',referrerPolicy:'no-referrer'});}catch(_){fail('Could not download the library. Download its text file from GitHub and choose it locally instead.');}
  if(!response.ok)fail('GitHub download failed ('+response.status+'). Use a public source or select the downloaded text file.');
  let text;
  if(url.startsWith('https://api.github.com/gists/')) {
    const data=await response.json(),files=Object.values(data.files||{}).filter(f=>/\.(txt|afe|lambda)$/i.test(f.filename));
    if(!files.length)fail('No library text file found in this Gist.');
    const file=files.length===1?files[0]:await chooseFile(files);
    if(!file)fail('Choose a library file.');
    if(file.truncated) {const r=await fetch(sourceURL(file.raw_url),{cache:'no-store',credentials:'omit'});if(!r.ok)fail('Could not download the complete Gist file.');text=await r.text();}
    else text=file.content;
  } else text=await response.text();
  return readLibrary(text);
}
root.GeoUpdater={tokenize,declarations,readLibrary,updateWorkbook,sourceURL,fetchLibrary,compile,decodeProject,INFO,dottedAliases,migrateFormula,migrateFormulaXml};
})(typeof window!=='undefined'?window:globalThis);
