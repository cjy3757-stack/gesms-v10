/* GESMS 군우명단 Excel 자동연동
 * GitHub 저장소 최상위의 GESMS_군우명단.xlsx를 캐시 없이 읽습니다.
 * 실패하면 기존 data.js 명단을 그대로 사용합니다.
 */
(function(){
  'use strict';

  const FILE_NAMES = [
    'GESMS_군우명단.xlsx',
    '군우명단.xlsx',
    'members.xlsx'
  ];

  const text = v => String(v ?? '').trim();
  const norm = v => text(v).replace(/\s+/g,'').replace(/[()_\-]/g,'').toLowerCase();
  const pad = n => String(n).padStart(2,'0');

  function colIndex(ref){
    let n=0;
    for(const ch of (ref.match(/[A-Z]+/i)||['A'])[0].toUpperCase()){
      n=n*26+ch.charCodeAt(0)-64;
    }
    return n-1;
  }

  async function parseXlsx(buffer){
    if(!window.JSZip) throw new Error('JSZip을 불러오지 못했습니다.');
    const zip = await JSZip.loadAsync(buffer);

    const shared = [];
    const sx = zip.file('xl/sharedStrings.xml');
    if(sx){
      const doc = new DOMParser().parseFromString(await sx.async('text'),'application/xml');
      [...doc.getElementsByTagNameNS('*','si')].forEach(si=>{
        shared.push([...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join(''));
      });
    }

    const workbook = zip.file('xl/workbook.xml');
    const rels = zip.file('xl/_rels/workbook.xml.rels');
    let sheetPath = 'xl/worksheets/sheet1.xml';

    if(workbook && rels){
      const wbDoc = new DOMParser().parseFromString(await workbook.async('text'),'application/xml');
      const relDoc = new DOMParser().parseFromString(await rels.async('text'),'application/xml');
      const sheets = [...wbDoc.getElementsByTagNameNS('*','sheet')];
      const preferred = sheets.find(s=>/군우|교인|명단|member/i.test(s.getAttribute('name')||'')) || sheets[0];
      const rid = preferred?.getAttribute('r:id') || preferred?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
      if(rid){
        const rel = [...relDoc.getElementsByTagNameNS('*','Relationship')].find(r=>r.getAttribute('Id')===rid);
        const target = rel?.getAttribute('Target');
        if(target){
          sheetPath = target.startsWith('/') ? target.slice(1) : ('xl/' + target.replace(/^(\.\.\/)+/,''));
        }
      }
    }

    const sheetFile = zip.file(sheetPath) || zip.file('xl/worksheets/sheet1.xml');
    if(!sheetFile) throw new Error('군우명단 시트를 찾지 못했습니다.');

    const doc = new DOMParser().parseFromString(await sheetFile.async('text'),'application/xml');
    const rows = [];
    [...doc.getElementsByTagNameNS('*','row')].forEach(row=>{
      const out = [];
      [...row.getElementsByTagNameNS('*','c')].forEach(c=>{
        const idx = colIndex(c.getAttribute('r')||'A1');
        const type = c.getAttribute('t');
        const v = c.getElementsByTagNameNS('*','v')[0]?.textContent ?? '';
        let value = v;
        if(type==='s') value = shared[Number(v)] ?? '';
        else if(type==='inlineStr') value = [...c.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('');
        out[idx] = value;
      });
      rows[Number(row.getAttribute('r')||rows.length+1)-1] = out;
    });
    return rows;
  }

  function findHeader(rows){
    const nameAliases = ['이름','성명','군우명','교인명'];
    for(let i=0;i<Math.min(rows.length,30);i++){
      const vals=(rows[i]||[]).map(text);
      if(vals.some(v=>nameAliases.map(norm).includes(norm(v)))){
        return {index:i,headers:vals};
      }
    }
    throw new Error('이름 또는 성명 열을 찾지 못했습니다.');
  }

  function aliasIndex(headers, aliases){
    const ns = headers.map(norm);
    for(const a of aliases){
      const i = ns.indexOf(norm(a));
      if(i>=0) return i;
    }
    return -1;
  }

  function excelDateToMonthDay(v){
    if(v===null || v===undefined || v==='') return '';
    if(typeof v==='number' || (/^\d+(\.\d+)?$/.test(text(v)) && Number(v)>20000)){
      const n=Number(v);
      const d=new Date(Date.UTC(1899,11,30)+Math.round(n*86400000));
      return `${pad(d.getUTCMonth()+1)}월 ${pad(d.getUTCDate())}일`;
    }
    const s=text(v);
    let m=s.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if(m) return `${pad(m[2])}월 ${pad(m[3])}일`;
    m=s.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일?/);
    if(m) return `${pad(m[1])}월 ${pad(m[2])}일`;
    m=s.match(/^(\d{1,2})[-/.](\d{1,2})$/);
    if(m) return `${pad(m[1])}월 ${pad(m[2])}일`;
    return s;
  }

  function monthFromBirthday(v){
    const m=text(v).match(/(\d{1,2})\s*월/);
    return m ? Number(m[1]) : 0;
  }

  function buildData(rows){
    const h=findHeader(rows);
    const ix = {
      family: aliasIndex(h.headers,['가족코드','가정코드','가족','가정','family']),
      familyHead: aliasIndex(h.headers,['가족대표','가정대표','세대주','가족대표자']),
      name: aliasIndex(h.headers,['이름','성명','군우명','교인명']),
      solarBirthday: aliasIndex(h.headers,['양력생일','양력 생일','양력생년월일']),
      birthday: aliasIndex(h.headers,['생일','생년월일','생일자']),
      birthdayType: aliasIndex(h.headers,['생일구분','생일 구분','음양력','생일종류']),
      position: aliasIndex(h.headers,['직분','직책','계급']),
      phone: aliasIndex(h.headers,['전화번호','휴대폰','핸드폰','연락처']),
      group1: aliasIndex(h.headers,['소속1','소속 1','그룹1','부서1']),
      group2: aliasIndex(h.headers,['소속2','소속 2','그룹2','부서2']),
      group3: aliasIndex(h.headers,['소속3','소속 3','그룹3','부서3']),
      group4: aliasIndex(h.headers,['소속4','소속 4','그룹4','부서4']),
      note: aliasIndex(h.headers,['비고','메모','특이사항'])
    };

    const get=(r,i)=>i>=0?text(r[i]):'';
    const members=[];

    for(let i=h.index+1;i<rows.length;i++){
      const r=rows[i]||[];
      const name=get(r,ix.name);
      if(!name) continue;

      const family = get(r,ix.family) || String(members.length+1);
      const rawSolar = get(r,ix.solarBirthday);
      const rawBirthday = get(r,ix.birthday);
      const birthday = excelDateToMonthDay(rawSolar || rawBirthday);
      const kind = get(r,ix.birthdayType);
      const birthdayType = kind
        ? `${kind.replace(/[()]/g,'').trim()}) ${birthday.replace(/\s+/g,'')}`
        : birthday;

      members.push({
        family,
        name,
        birthday,
        birthdayType,
        position:get(r,ix.position),
        phone:get(r,ix.phone),
        group1:get(r,ix.group1),
        group2:get(r,ix.group2),
        group3:get(r,ix.group3),
        group4:get(r,ix.group4),
        note:get(r,ix.note),
        _familyHead:get(r,ix.familyHead)
      });
    }

    if(!members.length) throw new Error('군우 데이터가 없습니다.');

    const familyMap=new Map();
    for(const m of members){
      if(!familyMap.has(m.family)) familyMap.set(m.family,[]);
      familyMap.get(m.family).push(m);
    }

    const families=[...familyMap.entries()].map(([code,list])=>{
      const explicit = list.find(x=>x._familyHead)?._familyHead;
      const head = explicit || list[0].name;
      const groups=[...new Set(list.flatMap(m=>[m.group1,m.group2,m.group3,m.group4]).filter(Boolean))];
      const clean=list.map(({_familyHead,...m})=>m);
      return {code,label:`${head} 가정`,head,members:clean,groups,size:clean.length};
    });

    const cleanMembers=members.map(({_familyHead,...m})=>m);
    const positions={}, groups={}, birthMonths={};
    for(const m of cleanMembers){
      const p=m.position||'미기재';
      positions[p]=(positions[p]||0)+1;
      [m.group1,m.group2,m.group3,m.group4].filter(Boolean).forEach(g=>groups[g]=(groups[g]||0)+1);
      const mon=monthFromBirthday(m.birthday);
      if(mon) birthMonths[mon]=(birthMonths[mon]||0)+1;
    }

    return {members:cleanMembers,families,stats:{positions,groups,birthMonths}};
  }

  async function fetchWorkbook(){
    let last='';
    for(const name of FILE_NAMES){
      try{
        const url=new URL(name,location.href).href;
        const r=await fetch(`${url}${url.includes('?')?'&':'?'}v=${Date.now()}`,{cache:'no-store'});
        if(!r.ok){last=`${name}: HTTP ${r.status}`;continue;}
        const buf=await r.arrayBuffer();
        if(buf.byteLength<100){last=`${name}: 파일이 비어 있습니다.`;continue;}
        return {name,rows:await parseXlsx(buf)};
      }catch(e){last=e.message||String(e);}
    }
    throw new Error(last||'군우명단 Excel을 불러오지 못했습니다.');
  }

  window.loadLatestMemberExcel = async function(){
    const fallback=window.APP_DATA||{};
    try{
      const wb=await fetchWorkbook();
      const fresh=buildData(wb.rows);
      window.APP_DATA=Object.assign({},fallback,fresh);
      window.GESMS_MEMBER_EXCEL={ok:true,name:wb.name,count:fresh.members.length};
      console.info(`[GESMS] ${wb.name}에서 군우 ${fresh.members.length}명을 불러왔습니다.`);
      return true;
    }catch(e){
      window.GESMS_MEMBER_EXCEL={ok:false,error:e.message||String(e)};
      console.warn('[GESMS] 군우명단 Excel 자동연동 실패. data.js 예비명단을 사용합니다.',e);
      return false;
    }
  };
})();
