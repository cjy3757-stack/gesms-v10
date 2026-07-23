/* GESMS V10.0 Release - robust Excel office data loader */
(function(){
  'use strict';
  const OWNER='cjy3757-stack';
  const REPO='gesms-v10';
  const BRANCHES=['main','master'];
  const NOTICE_MATCH=name=>/^공지사항(?:\([^)]*\)|[_ -].*)?\.xlsx$/i.test(name)||name==='notice.xlsx';
  const EVENT_MATCH=name=>/^교회일정(?:\([^)]*\)|[_ -].*)?\.xlsx$/i.test(name)||name==='events.xlsx';
  const text=v=>String(v??'').trim();
  const pad=n=>String(n).padStart(2,'0');

  function excelDate(v){
    if(!v&&v!==0)return '';
    if(typeof v==='string' && /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(v)){
      const m=v.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
    }
    const n=Number(v);if(!Number.isFinite(n))return text(v);
    const d=new Date(Date.UTC(1899,11,30)+Math.round(n*86400000));
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`;
  }
  function excelTime(v){
    if(v===null||v===undefined||v==='')return '';
    if(typeof v==='string')return text(v);
    const n=Number(v);if(!Number.isFinite(n))return text(v);
    const mins=Math.round((n-Math.floor(n))*1440)%1440;
    return `${pad(Math.floor(mins/60))}:${pad(mins%60)}`;
  }
  function colIndex(ref){let n=0;for(const ch of (ref.match(/[A-Z]+/i)||['A'])[0].toUpperCase())n=n*26+ch.charCodeAt(0)-64;return n-1;}
  async function parseXlsx(buffer){
    if(!window.JSZip)throw new Error('Excel 압축 해제 도구를 불러오지 못했습니다.');
    const zip=await JSZip.loadAsync(buffer);
    const sx=zip.file('xl/sharedStrings.xml');
    const shared=[];
    if(sx){
      const doc=new DOMParser().parseFromString(await sx.async('text'),'application/xml');
      [...doc.getElementsByTagNameNS('*','si')].forEach(si=>shared.push([...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('')));
    }
    const sheetFile=zip.file('xl/worksheets/sheet1.xml');
    if(!sheetFile)throw new Error('첫 번째 Excel 시트를 찾지 못했습니다.');
    const doc=new DOMParser().parseFromString(await sheetFile.async('text'),'application/xml');
    const rows=[];
    [...doc.getElementsByTagNameNS('*','row')].forEach(row=>{
      const out=[];
      [...row.getElementsByTagNameNS('*','c')].forEach(c=>{
        const idx=colIndex(c.getAttribute('r')||'A1');
        const type=c.getAttribute('t');
        const v=c.getElementsByTagNameNS('*','v')[0]?.textContent??'';
        let value=v;
        if(type==='s')value=shared[Number(v)]??'';
        else if(type==='inlineStr')value=[...c.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('');
        out[idx]=value;
      });
      rows[Number(row.getAttribute('r')||rows.length+1)-1]=out;
    });
    return rows;
  }
  function findHeader(rows,required){
    for(let i=0;i<Math.min(rows.length,20);i++){
      const vals=(rows[i]||[]).map(text);
      if(required.every(h=>vals.includes(h)))return {index:i,headers:vals};
    }
    throw new Error(`필수 열(${required.join(', ')})을 찾지 못했습니다.`);
  }
  async function githubFiles(){
    let last='';
    for(const branch of BRANCHES){
      try{
        const url=`https://api.github.com/repos/${OWNER}/${REPO}/contents/?ref=${branch}&t=${Date.now()}`;
        const r=await fetch(url,{cache:'no-store',headers:{Accept:'application/vnd.github+json'}});
        if(!r.ok){last=`GitHub API HTTP ${r.status}`;continue;}
        const list=await r.json();
        return Array.isArray(list)?list:[];
      }catch(e){last=e.message||String(e);}
    }
    throw new Error(last||'GitHub 파일 목록을 읽지 못했습니다.');
  }
  async function fetchBuffer(url,label){
    const r=await fetch(`${url}${url.includes('?')?'&':'?'}v=${Date.now()}`,{cache:'no-store'});
    if(!r.ok)throw new Error(`${label}: HTTP ${r.status}`);
    const buf=await r.arrayBuffer();
    if(buf.byteLength<100)throw new Error(`${label}: 파일이 비어 있습니다.`);
    return buf;
  }
  async function fetchWorkbook(kind){
    const matcher=kind==='notice'?NOTICE_MATCH:EVENT_MATCH;
    const label=kind==='notice'?'공지사항':'교회일정';
    let apiError='';
    try{
      const files=await githubFiles();
      const candidates=files.filter(x=>x.type==='file'&&matcher(x.name));
      const preferredNames=kind==='notice'
        ? ['공지사항.xlsx','notice.xlsx']
        : ['교회일정.xlsx','events.xlsx'];
      const rank=name=>{
        const exact=preferredNames.indexOf(name);
        if(exact>=0)return exact;
        if(kind==='notice'&&name.startsWith('공지사항'))return 10;
        if(kind==='event'&&name.startsWith('교회일정'))return 10;
        return 99;
      };
      candidates.sort((a,b)=>rank(a.name)-rank(b.name)||a.name.localeCompare(b.name,'ko'));
      const item=candidates[0];
      if(item){
        const url=item.download_url || `https://raw.githubusercontent.com/${OWNER}/${REPO}/main/${item.path.split('/').map(encodeURIComponent).join('/')}`;
        return {rows:await parseXlsx(await fetchBuffer(url,item.name)),name:item.name};
      }
      apiError=`${label} Excel 파일이 저장소 최상위에 없습니다.`;
    }catch(e){apiError=e.message||String(e);}

    // GitHub API가 일시적으로 제한될 때를 위한 동일 출처 보조 경로
    const aliases=kind==='notice'?['공지사항.xlsx','notice.xlsx']:['교회일정.xlsx','events.xlsx'];
    let last=apiError;
    for(const name of aliases){
      try{
        const url=new URL(name,location.href).href;
        return {rows:await parseXlsx(await fetchBuffer(url,name)),name};
      }catch(e){last=e.message||String(e);}
    }
    throw new Error(last||`${label} 파일을 불러오지 못했습니다.`);
  }
  function rowsToNotices(rows){
    const h=findHeader(rows,['시작일','제목','내용']);
    const ix=n=>h.headers.indexOf(n);const out=[];
    for(let i=h.index+1;i<rows.length;i++){
      const r=rows[i]||[];const title=text(r[ix('제목')]);const body=text(r[ix('내용')]);
      if(!title&&!body)continue;
      out.push({start:excelDate(r[ix('시작일')]),end:excelDate(r[ix('종료일')]),important:text(r[ix('중요')]).toUpperCase()==='Y'?'Y':'N',title,body,category:text(r[ix('구분')])});
    }
    return out;
  }
  function rowsToEvents(rows){
    const h=findHeader(rows,['날짜','행사명']);
    const ix=n=>h.headers.indexOf(n);const out=[];
    for(let i=h.index+1;i<rows.length;i++){
      const r=rows[i]||[];const title=text(r[ix('행사명')]);if(!title)continue;
      out.push({date:excelDate(r[ix('날짜')]),time:excelTime(r[ix('시간')]),category:text(r[ix('구분')])||'일정',title,location:text(r[ix('장소')]),owner:text(r[ix('담당')]),note:text(r[ix('비고')])});
    }
    return out;
  }
  function setStatus(msg,ok){
    const el=document.getElementById('officeUpdateStatus');if(!el)return;
    el.textContent=msg;el.style.color=ok?'#176b3a':'#9e1b32';
  }
  window.loadLatestOfficeExcel=async function(){
    setStatus('공지사항과 교회일정 최신 파일을 확인하는 중입니다…',true);
    const results=await Promise.allSettled([fetchWorkbook('notice'),fetchWorkbook('event')]);
    const messages=[];
    if(results[0].status==='fulfilled'){
      window.NOTICE_DATA=rowsToNotices(results[0].value.rows);
      const today=new Date().toISOString().slice(0,10);
      const activeCount=window.NOTICE_DATA.filter(n=>(!n.start||n.start<=today)&&(!n.end||n.end>=today)).length;
      messages.push(`현재 공지 ${activeCount}건 / 전체 ${window.NOTICE_DATA.length}건 (${results[0].value.name})`);
    }else messages.push(`공지 실패: ${results[0].reason?.message||results[0].reason}`);
    if(results[1].status==='fulfilled'){
      window.EVENT_DATA=rowsToEvents(results[1].value.rows);messages.push(`일정 ${window.EVENT_DATA.length}건 (${results[1].value.name})`);
    }else messages.push(`일정 실패: ${results[1].reason?.message||results[1].reason}`);
    if(typeof window.refreshOfficeDataFromExcel==='function')window.refreshOfficeDataFromExcel();
    const ok=results.some(r=>r.status==='fulfilled');
    setStatus(`${messages.join(' · ')} · ${new Date().toLocaleTimeString('ko-KR')} 확인`,ok);
    return ok;
  };
})();
