/* GESMS V10.3 - GitHub 군우명단 Excel 자동연동 */
(function(){
'use strict';

const OWNER='cjy3757-stack';
const REPO='gesms-v10';
const BRANCHES=['main','master'];
const MATCH=name=>/^(?:GESMS[_ -]*)?군우명단(?:\([^)]*\)|[_ -].*)?\.xlsx$/i.test(name)
  || /^members(?:[_ -].*)?\.xlsx$/i.test(name);

const text=v=>String(v??'').trim();
const norm=v=>text(v).replace(/\s+/g,'').replace(/[()_\-]/g,'').toLowerCase();
const pad=n=>String(n).padStart(2,'0');

function colIndex(ref){
  let n=0;
  for(const ch of (ref.match(/[A-Z]+/i)||['A'])[0].toUpperCase()) n=n*26+ch.charCodeAt(0)-64;
  return n-1;
}

async function sharedStrings(zip){
  const out=[];
  const f=zip.file('xl/sharedStrings.xml');
  if(!f)return out;
  const d=new DOMParser().parseFromString(await f.async('text'),'application/xml');
  [...d.getElementsByTagNameNS('*','si')].forEach(si=>{
    out.push([...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join(''));
  });
  return out;
}

async function workbookSheets(zip){
  const wb=zip.file('xl/workbook.xml'), rels=zip.file('xl/_rels/workbook.xml.rels');
  if(!wb||!rels)return [{name:'sheet1',path:'xl/worksheets/sheet1.xml'}];
  const wd=new DOMParser().parseFromString(await wb.async('text'),'application/xml');
  const rd=new DOMParser().parseFromString(await rels.async('text'),'application/xml');
  const relMap=new Map([...rd.getElementsByTagNameNS('*','Relationship')].map(r=>[r.getAttribute('Id'),r.getAttribute('Target')]));
  return [...wd.getElementsByTagNameNS('*','sheet')].map(s=>{
    const rid=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target=relMap.get(rid)||'worksheets/sheet1.xml';
    target=target.replace(/^\/+/,'').replace(/^xl\//,'').replace(/^(\.\.\/)+/,'');
    return {name:s.getAttribute('name')||'',path:'xl/'+target};
  });
}

async function parseSheet(zip,path,shared){
  const f=zip.file(path);
  if(!f)return [];
  const d=new DOMParser().parseFromString(await f.async('text'),'application/xml');
  const rows=[];
  [...d.getElementsByTagNameNS('*','row')].forEach(row=>{
    const out=[];
    [...row.getElementsByTagNameNS('*','c')].forEach(c=>{
      const idx=colIndex(c.getAttribute('r')||'A1');
      const type=c.getAttribute('t');
      const v=c.getElementsByTagNameNS('*','v')[0]?.textContent??'';
      let val=v;
      if(type==='s')val=shared[Number(v)]??'';
      else if(type==='inlineStr')val=[...c.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('');
      out[idx]=val;
    });
    rows[Number(row.getAttribute('r')||rows.length+1)-1]=out;
  });
  return rows;
}

async function parseXlsx(buf){
  if(!window.JSZip)throw new Error('jszip.min.js를 불러오지 못했습니다.');
  const zip=await JSZip.loadAsync(buf);
  const shared=await sharedStrings(zip);
  const sheets=await workbookSheets(zip);
  const preferred=[...sheets].sort((a,b)=>{
    const score=x=>/군우|교인|명단|member/i.test(x.name)?0:1;
    return score(a)-score(b);
  });
  let last=[];
  for(const s of preferred){
    const rows=await parseSheet(zip,s.path,shared);
    if(rows.length){
      last=rows;
      try{findHeader(rows);return {rows,sheet:s.name};}catch(e){}
    }
  }
  if(last.length)return {rows:last,sheet:preferred[0]?.name||'sheet1'};
  throw new Error('읽을 수 있는 Excel 시트가 없습니다.');
}

function findHeader(rows){
  const aliases=['이름','성명','군우명','교인명'];
  for(let i=0;i<Math.min(rows.length,40);i++){
    const vals=(rows[i]||[]).map(text);
    if(vals.some(v=>aliases.map(norm).includes(norm(v))))return {index:i,headers:vals};
  }
  throw new Error('이름/성명 열을 찾지 못했습니다.');
}

function aliasIndex(headers,aliases){
  const hs=headers.map(norm);
  for(const a of aliases){
    const i=hs.indexOf(norm(a));
    if(i>=0)return i;
  }
  return -1;
}

function monthDay(v){
  if(v===null||v===undefined||v==='')return '';
  const s=text(v);
  const num=Number(s);
  if(Number.isFinite(num)&&num>20000){
    const d=new Date(Date.UTC(1899,11,30)+Math.round(num*86400000));
    return `${pad(d.getUTCMonth()+1)}월 ${pad(d.getUTCDate())}일`;
  }
  let m=s.match(/(?:\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if(m)return `${pad(m[1])}월 ${pad(m[2])}일`;
  m=s.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일?/);
  if(m)return `${pad(m[1])}월 ${pad(m[2])}일`;
  m=s.match(/^(\d{1,2})[-/.](\d{1,2})$/);
  if(m)return `${pad(m[1])}월 ${pad(m[2])}일`;
  return s;
}

function buildData(rows){
  const h=findHeader(rows);
  const ix={
    family:aliasIndex(h.headers,['가족코드','가정코드','가족','가정','family']),
    head:aliasIndex(h.headers,['가족대표','가정대표','세대주','가족대표자']),
    name:aliasIndex(h.headers,['이름','성명','군우명','교인명']),
    solar:aliasIndex(h.headers,['양력생일','양력 생일','양력생년월일']),
    birth:aliasIndex(h.headers,['생일','생년월일']),
    kind:aliasIndex(h.headers,['생일구분','음양력','생일종류']),
    position:aliasIndex(h.headers,['직분','직책','계급']),
    phone:aliasIndex(h.headers,['전화번호','휴대폰','핸드폰','연락처']),
    group1:aliasIndex(h.headers,['소속1','소속 1','그룹1','부서1']),
    group2:aliasIndex(h.headers,['소속2','소속 2','그룹2','부서2']),
    group3:aliasIndex(h.headers,['소속3','소속 3','그룹3','부서3']),
    group4:aliasIndex(h.headers,['소속4','소속 4','그룹4','부서4']),
    note:aliasIndex(h.headers,['비고','메모','특이사항'])
  };
  const get=(r,i)=>i>=0?text(r[i]):'';
  const members=[];
  for(let i=h.index+1;i<rows.length;i++){
    const r=rows[i]||[],name=get(r,ix.name);
    if(!name)continue;
    const birthday=monthDay(get(r,ix.solar)||get(r,ix.birth));
    const kind=get(r,ix.kind);
    members.push({
      family:get(r,ix.family)||String(members.length+1),
      name,
      birthday,
      birthdayType:kind?`${kind.replace(/[()]/g,'').trim()}) ${birthday.replace(/\s+/g,'')}`:birthday,
      position:get(r,ix.position),phone:get(r,ix.phone),
      group1:get(r,ix.group1),group2:get(r,ix.group2),
      group3:get(r,ix.group3),group4:get(r,ix.group4),
      note:get(r,ix.note),_head:get(r,ix.head)
    });
  }
  if(!members.length)throw new Error('군우명단에 데이터가 없습니다.');

  const map=new Map();
  members.forEach(m=>{if(!map.has(m.family))map.set(m.family,[]);map.get(m.family).push(m)});
  const families=[...map].map(([code,list])=>{
    const head=list.find(x=>x._head)?._head||list[0].name;
    const clean=list.map(({_head,...m})=>m);
    const groups=[...new Set(clean.flatMap(m=>[m.group1,m.group2,m.group3,m.group4]).filter(Boolean))];
    return {code,label:`${head} 가정`,head,members:clean,groups,size:clean.length};
  });
  const cleanMembers=members.map(({_head,...m})=>m);
  const positions={},groups={},birthMonths={};
  cleanMembers.forEach(m=>{
    const p=m.position||'미기재';positions[p]=(positions[p]||0)+1;
    [m.group1,m.group2,m.group3,m.group4].filter(Boolean).forEach(g=>groups[g]=(groups[g]||0)+1);
    const mm=Number((m.birthday.match(/(\d{1,2})월/)||[])[1]||0);
    if(mm)birthMonths[mm]=(birthMonths[mm]||0)+1;
  });
  return {members:cleanMembers,families,stats:{positions,groups,birthMonths}};
}

async function githubFiles(){
  let last='';
  for(const branch of BRANCHES){
    try{
      const u=`https://api.github.com/repos/${OWNER}/${REPO}/contents/?ref=${branch}&t=${Date.now()}`;
      const r=await fetch(u,{cache:'no-store',headers:{Accept:'application/vnd.github+json'}});
      if(!r.ok){last=`GitHub API HTTP ${r.status}`;continue}
      const list=await r.json();
      return {branch,list:Array.isArray(list)?list:[]};
    }catch(e){last=e.message||String(e)}
  }
  throw new Error(last||'GitHub 파일목록을 읽지 못했습니다.');
}

async function fetchWorkbook(){
  let apiError='';
  try{
    const {list}=await githubFiles();
    const candidates=list.filter(x=>x.type==='file'&&MATCH(x.name));
    candidates.sort((a,b)=>{
      const rank=n=>n==='GESMS_군우명단.xlsx'?0:n==='군우명단.xlsx'?1:n==='members.xlsx'?2:9;
      return rank(a.name)-rank(b.name)||a.name.localeCompare(b.name,'ko');
    });
    if(candidates[0]){
      const f=candidates[0];
      const r=await fetch(`${f.download_url}${f.download_url.includes('?')?'&':'?'}v=${Date.now()}`,{cache:'no-store'});
      if(!r.ok)throw new Error(`${f.name}: HTTP ${r.status}`);
      return {name:f.name,...await parseXlsx(await r.arrayBuffer())};
    }
    apiError='저장소 최상위에 GESMS_군우명단.xlsx가 없습니다.';
  }catch(e){apiError=e.message||String(e)}

  for(const name of ['GESMS_군우명단.xlsx','군우명단.xlsx','members.xlsx']){
    try{
      const u=new URL(name,location.href).href;
      const r=await fetch(`${u}?v=${Date.now()}`,{cache:'no-store'});
      if(!r.ok)continue;
      return {name,...await parseXlsx(await r.arrayBuffer())};
    }catch(e){}
  }
  throw new Error(apiError||'군우명단 Excel을 불러오지 못했습니다.');
}

window.loadLatestMemberExcel=async function(){
  const fallback=window.APP_DATA||{};
  try{
    const wb=await fetchWorkbook();
    const fresh=buildData(wb.rows);
    window.APP_DATA=Object.assign({},fallback,fresh);
    window.GESMS_MEMBER_EXCEL={ok:true,name:wb.name,sheet:wb.sheet,count:fresh.members.length,loadedAt:new Date().toISOString()};
    console.info(`[GESMS V10.3] ${wb.name} / ${wb.sheet}: ${fresh.members.length}명`);
    return true;
  }catch(e){
    window.GESMS_MEMBER_EXCEL={ok:false,error:e.message||String(e)};
    console.warn('[GESMS V10.3] Excel 연동 실패, data.js 예비명단 사용:',e);
    return false;
  }
};
})();