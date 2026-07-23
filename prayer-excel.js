/* GESMS V10.1 - unified evangelism workbook module */
(function(){
'use strict';
const UNIFIED_FILE='GESMS_전도영혼구원_표준자료양식_V3.0.xlsx';
const FALLBACK_PRAYER_FILES=['기도대상자관리_표준양식_V2.0.xlsx','기도대상자관리.xlsx'];
const STAGES=['기도중','관계형성','첫만남','초청','교회방문','새가족','등록교인','군우편입'];
const $=s=>document.querySelector(s), text=v=>String(v??'').trim(), esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad=n=>String(n).padStart(2,'0');
let targets=[], projects=[];
function excelDate(v){if(v===null||v===undefined||v==='')return '';if(typeof v==='string'&&/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(v)){const m=v.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;}const n=Number(v);if(!Number.isFinite(n))return text(v);const d=new Date(Date.UTC(1899,11,30)+Math.round(n*86400000));return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`;}
function colIndex(ref){let n=0;for(const ch of (ref.match(/[A-Z]+/i)||['A'])[0].toUpperCase())n=n*26+ch.charCodeAt(0)-64;return n-1;}
async function workbookSheets(buffer){
  const zip=await JSZip.loadAsync(buffer), shared=[];
  const sx=zip.file('xl/sharedStrings.xml');
  if(sx){const doc=new DOMParser().parseFromString(await sx.async('text'),'application/xml');[...doc.getElementsByTagNameNS('*','si')].forEach(si=>shared.push([...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('')))}
  const wbxml=zip.file('xl/workbook.xml'); if(!wbxml)throw new Error('Excel 통합문서 정보를 찾지 못했습니다.');
  const wdoc=new DOMParser().parseFromString(await wbxml.async('text'),'application/xml');
  const rel=zip.file('xl/_rels/workbook.xml.rels'); if(!rel)throw new Error('Excel 시트 연결정보를 찾지 못했습니다.');
  const rdoc=new DOMParser().parseFromString(await rel.async('text'),'application/xml');
  const rels={}; [...rdoc.getElementsByTagNameNS('*','Relationship')].forEach(r=>rels[r.getAttribute('Id')]=r.getAttribute('Target'));
  const result={};
  for(const s of [...wdoc.getElementsByTagNameNS('*','sheet')]){
    const name=s.getAttribute('name'), rid=s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')||s.getAttribute('r:id');
    let target=rels[rid]||''; target='xl/'+target.replace(/^\/?xl\//,'');
    const sf=zip.file(target); if(!sf)continue;
    const doc=new DOMParser().parseFromString(await sf.async('text'),'application/xml'), rows=[];
    [...doc.getElementsByTagNameNS('*','row')].forEach(row=>{const out=[];[...row.getElementsByTagNameNS('*','c')].forEach(c=>{const idx=colIndex(c.getAttribute('r')||'A1'),type=c.getAttribute('t'),v=c.getElementsByTagNameNS('*','v')[0]?.textContent??'';let value=v;if(type==='s')value=shared[Number(v)]??'';else if(type==='inlineStr')value=[...c.getElementsByTagNameNS('*','t')].map(t=>t.textContent||'').join('');out[idx]=value});rows[Number(row.getAttribute('r')||rows.length+1)-1]=out});
    result[name]=rows;
  }
  return result;
}
function rowsToObjects(rows, required){let hi=-1,headers=[];for(let i=0;i<Math.min(rows?.length||0,25);i++){const a=(rows[i]||[]).map(text);if(required.every(h=>a.includes(h))){hi=i;headers=a;break}}if(hi<0)throw new Error(required.join('·')+' 열을 찾지 못했습니다.');const out=[];for(let i=hi+1;i<rows.length;i++){const r=rows[i]||[];if(!r.some(v=>text(v)))continue;const o={};headers.forEach((h,j)=>{if(h)o[h]=r[j]});out.push(o)}return out;}
function rowsToTargets(rows){return rowsToObjects(rows,['관리번호','이름','담당자']).filter(o=>text(o['이름'])&& !['미사용','중지'].includes(text(o['사용여부']))).map(o=>{['기도시작일','최근만남일','다음연락일','교회방문일','새가족등록일','등록교인일','군우편입일','기도응답일'].forEach(k=>o[k]=excelDate(o[k]));return o});}
function rowsToProjects(rows){return rowsToObjects(rows,['프로젝트번호','사업명','분야']).filter(o=>text(o['사업명'])).map(o=>{['시작일','종료일'].forEach(k=>o[k]=excelDate(o[k]));let p=Number(o['진행률(%)']);if(Number.isFinite(p)){if(p<=1)p*=100;o['진행률(%)']=Math.max(0,Math.min(100,Math.round(p)))}else o['진행률(%)']=0;return o});}
async function fetchBuffer(name){const r=await fetch('./'+encodeURIComponent(name)+'?v='+Date.now(),{cache:'no-store'});if(!r.ok)throw new Error(`${name}: HTTP ${r.status}`);return r.arrayBuffer();}
async function fetchDefault(){
  try{const sheets=await workbookSheets(await fetchBuffer(UNIFIED_FILE));return {name:UNIFIED_FILE,targets:rowsToTargets(sheets['기도대상자관리']),projects:rowsToProjects(sheets['일만상상프로젝트'])};}
  catch(unifiedError){
    let last=unifiedError.message||String(unifiedError);
    for(const name of FALLBACK_PRAYER_FILES){try{const sheets=await workbookSheets(await fetchBuffer(name));const first=Object.values(sheets)[0]||[];return {name,targets:rowsToTargets(sheets['기도대상자관리']||first),projects:[],warning:'통합 V3.0 파일을 찾지 못해 기존 기도대상자 파일을 사용했습니다.'};}catch(e){last=e.message||String(e)}}
    throw new Error(last);
  }
}
function y(o,k){return text(o[k]).toUpperCase()==='Y';}
function currentStage(o){const direct=text(o['진행상태']);if(STAGES.includes(direct))return direct;for(let i=STAGES.length-1;i>=0;i--)if(y(o,STAGES[i]))return STAGES[i];return direct||'기도 준비';}
function getGoal(){return Math.max(1,Number(localStorage.getItem('gesms_soul_goal')||100));}
function filtered(){const q=text($('#prayerSearch')?.value).toLowerCase(),owner=text($('#prayerOwner')?.value),group=text($('#prayerGroup')?.value),state=text($('#prayerState')?.value);return targets.filter(o=>{const hay=[o['이름'],o['담당자'],o['관계구분'],o['관계상세'],o['기도제목구분'],o['기도제목상세'],o['소속부서/소그룹']].map(text).join(' ').toLowerCase();return(!q||hay.includes(q))&&(!owner||text(o['담당자'])===owner)&&(!group||text(o['소속부서/소그룹'])===group)&&(!state||currentStage(o)===state)});}
function daysFrom(date){if(!date)return null;const a=new Date(date+'T00:00:00'),b=new Date();b.setHours(0,0,0,0);return Math.round((a-b)/86400000);}
function renderFilters(){const owners=[...new Set(targets.map(o=>text(o['담당자'])).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ko'));const groups=[...new Set(targets.map(o=>text(o['소속부서/소그룹'])).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ko'));const fill=(sel,arr,label)=>{if(!sel)return;const old=sel.value;sel.innerHTML=`<option value="">${label}</option>`+arr.map(v=>`<option>${esc(v)}</option>`).join('');sel.value=old};fill($('#prayerOwner'),owners,'모든 담당자');fill($('#prayerGroup'),groups,'모든 소그룹');}
function renderStats(){const goal=getGoal(),counts={};STAGES.forEach(k=>counts[k]=targets.filter(o=>currentStage(o)===k).length);const current=targets.length,visit=counts['교회방문'],newc=counts['새가족'],registered=counts['등록교인']+counts['군우편입'];$('#soulGoal').value=goal;$('#soulGoalKpi').textContent=goal;$('#soulCurrentKpi').textContent=current;$('#soulVisitKpi').textContent=visit;$('#soulNewKpi').textContent=newc;const max=Math.max(1,...Object.values(counts));$('#soulStageChart').innerHTML=STAGES.map(k=>`<div class="stage-row"><span>${k}</span><i style="width:${counts[k]/max*100}%"></i><b>${counts[k]}</b></div>`).join('');const rate=Math.min(100,Math.round(registered/goal*100));$('#soulSummaryText').innerHTML=`현재 <b>${current}명</b>을 위해 기도하며, 등록·군우편입 <b>${registered}명</b>으로 연간 목표의 <b>${rate}%</b>를 이루었습니다.`;$('#homeSoulProgress').textContent=`${registered}/${goal}명 · ${rate}%`;$('#homeSoulMeta').textContent=`기도대상 ${current}명 · 교회방문 ${visit}명 · 새가족 ${newc}명`;const today=targets.filter(o=>currentStage(o)==='기도중').slice(0,5);$('#homePrayerCount').textContent=`${targets.filter(o=>currentStage(o)==='기도중').length}명`;$('#homePrayerNames').textContent=today.length?today.map(o=>text(o['이름'])).join(', '):'등록된 기도중 대상자가 없습니다.';const contacts=targets.filter(o=>{const d=daysFrom(o['다음연락일']);return d!==null&&d>=0&&d<=7}).sort((a,b)=>text(a['다음연락일']).localeCompare(text(b['다음연락일'])));$('#homeContactCount').textContent=`${contacts.length}명`;$('#homeContactNames').textContent=contacts.length?contacts.slice(0,4).map(o=>`${text(o['이름'])}(${o['다음연락일'].slice(5)})`).join(', '):'7일 이내 연락 대상이 없습니다.';}
function renderList(){const list=filtered();$('#prayerCount').textContent=`${list.length}명`;$('#prayerList').innerHTML=list.length?list.map(o=>{const stage=currentStage(o),phone=text(o['연락처']),contact=o['다음연락일']?`다음 연락 ${o['다음연락일']}`:'다음 연락 미정';return `<article class="prayer-person"><div class="prayer-head"><div class="prayer-avatar">${esc(text(o['이름']).slice(0,1))}</div><div><h4>${esc(o['이름'])}</h4><p>담당 ${esc(o['담당자']||'-')} · ${esc(o['소속부서/소그룹']||'소속 미정')} · ${esc(contact)}</p></div><span class="stage-badge">${esc(stage)}</span></div><div class="prayer-detail"><div class="detail-grid"><div><span>관계</span><b>${esc([o['관계구분'],o['관계상세']].filter(Boolean).join(' · ')||'-')}</b></div><div><span>연락처</span><b>${esc(phone||'-')}</b></div><div><span>기도 시작</span><b>${esc(o['기도시작일']||'-')}</b></div><div><span>최근 만남</span><b>${esc(o['최근만남일']||'-')}</b></div><div><span>다음 연락</span><b>${esc(o['다음연락일']||'-')}</b></div><div><span>진행상태</span><b>${esc(stage)}</b></div></div><div class="prayer-topic"><b>${esc(o['기도제목구분']||'기도제목')}</b><br>${esc(o['기도제목상세']||'-')}</div>${text(o['비고'])?`<div class="prayer-topic" style="border-color:var(--navy);background:#eef3f8"><b>비고</b><br>${esc(o['비고'])}</div>`:''}</div></article>`}).join(''):'<div class="no-results">조건에 맞는 기도대상자가 없습니다.</div>';document.querySelectorAll('.prayer-head').forEach(x=>x.onclick=()=>x.parentElement.classList.toggle('open'));}
function money(v){const n=Number(v);return Number.isFinite(n)?n.toLocaleString('ko-KR')+'원':'-';}
function renderProjects(){
  const el=$('#projectList'), summary=$('#projectSummary'); if(!el||!summary)return;
  const states=['예정','진행중','완료','보류'], counts=Object.fromEntries(states.map(s=>[s,projects.filter(p=>text(p['상태'])===s).length]));
  const budget=projects.reduce((a,p)=>a+(Number(p['예산'])||0),0), spent=projects.reduce((a,p)=>a+(Number(p['실제지출'])||0),0);
  summary.innerHTML=`<div class="soul-kpis"><div><span>전체</span><b>${projects.length}</b></div><div><span>진행중</span><b>${counts['진행중']}</b></div><div><span>완료</span><b>${counts['완료']}</b></div><div><span>예산 집행</span><b>${budget?Math.round(spent/budget*100):0}%</b></div></div>`;
  el.innerHTML=projects.length?projects.map(p=>`<article class="prayer-person open"><div class="prayer-head"><div class="prayer-avatar">🌱</div><div><h4>${esc(p['사업명'])}</h4><p>${esc(p['분야']||'-')} · ${esc(p['담당부서']||'-')} · 담당 ${esc(p['담당자']||'-')}</p></div><span class="stage-badge">${esc(p['상태']||'-')}</span></div><div class="prayer-detail"><div class="detail-grid"><div><span>기간</span><b>${esc(p['시작일']||'-')} ~ ${esc(p['종료일']||'-')}</b></div><div><span>진행률</span><b>${esc(p['진행률(%)'])}%</b></div><div><span>참여/수혜</span><b>${esc(p['참여인원']||0)}명 / ${esc(p['수혜인원']||0)}명</b></div><div><span>예산/지출</span><b>${money(p['예산'])} / ${money(p['실제지출'])}</b></div></div><div class="prayer-topic"><b>핵심활동</b><br>${esc(p['핵심활동']||p['사업목적']||'-')}</div>${text(p['결과/열매'])?`<div class="prayer-topic" style="border-color:var(--navy);background:#eef3f8"><b>결과·열매</b><br>${esc(p['결과/열매'])}</div>`:''}</div></article>`).join(''):'<div class="no-results">등록된 일만상상 프로젝트가 없습니다.</div>';
}
function setStatus(msg,error=false){const el=$('#prayerExcelStatus');if(!el)return;el.textContent=msg;el.classList.toggle('error',error);}
function apply(data){targets=data.targets||[];projects=data.projects||[];window.PRAYER_TARGETS=targets;window.MISSION_PROJECTS=projects;renderFilters();renderStats();renderList();renderProjects();setStatus(`${data.name} · 기도대상 ${targets.length}명 · 프로젝트 ${projects.length}건${data.warning?' · '+data.warning:''}`);}
async function load(){setStatus(`${UNIFIED_FILE} 최신 파일을 확인하는 중입니다.`);try{apply(await fetchDefault())}catch(e){targets=[];projects=[];renderFilters();renderStats();renderList();renderProjects();setStatus(`자동 불러오기 실패: ${e.message||e}`,true)}}
window.initPrayerMinistry=function(){['prayerSearch','prayerOwner','prayerGroup','prayerState'].forEach(id=>{const el=$('#'+id);if(el)el.addEventListener(id==='prayerSearch'?'input':'change',renderList)});$('#refreshPrayerExcel')?.addEventListener('click',load);$('#saveSoulGoal')?.addEventListener('click',()=>{localStorage.setItem('gesms_soul_goal',Math.max(1,Number($('#soulGoal').value||100)));renderStats()});$('#selectPrayerExcel')?.addEventListener('click',()=>$('#prayerExcelInput').click());$('#prayerExcelInput')?.addEventListener('change',async e=>{const f=e.target.files[0];if(!f)return;try{const sheets=await workbookSheets(await f.arrayBuffer());apply({name:f.name,targets:rowsToTargets(sheets['기도대상자관리']),projects:rowsToProjects(sheets['일만상상프로젝트'])})}catch(err){setStatus(`파일 읽기 실패: ${err.message||err}`,true)}});$('#downloadPrayerTemplate')?.addEventListener('click',()=>{location.href='./'+encodeURIComponent(UNIFIED_FILE)});load();};
})();
