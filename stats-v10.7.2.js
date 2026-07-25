/* GESMS V10.7.2 - 목양 통계 대시보드 (고정 영역 통합판) */
(function(){
  'use strict';
  const VERSION='V10.7.2';
  const $=s=>document.querySelector(s);
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const year=()=>new Date().getFullYear();
  const ageOf=m=>{const y=Number(m&&m.birthYear);return Number.isFinite(y)&&y>1900?year()-y:null};
  const pct=(n,d)=>d?Math.round(n/d*100):0;

  function ensureStyles(){
    if($('#gesmsV1072Style'))return;
    const s=document.createElement('style');s.id='gesmsV1072Style';s.textContent=`
    #pastoralStatsDashboard{margin-top:16px;margin-bottom:16px}
    .ps-title{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;flex-wrap:wrap;margin-bottom:13px}
    .ps-title h3{margin:0;color:var(--navy)}.ps-badge{padding:6px 10px;border-radius:999px;background:#eef5fb;color:#24517a;font-size:12px;font-weight:900}
    .ps-kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:9px;margin-bottom:13px}
    .ps-kpi{background:#fff;border:1px solid var(--line);border-radius:15px;padding:13px;box-shadow:var(--shadow)}.ps-kpi span{display:block;color:var(--muted);font-size:11px}.ps-kpi b{display:block;color:var(--red);font-size:22px;margin-top:4px}
    .ps-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:13px}.ps-card{background:#fff;border:1px solid var(--line);border-radius:17px;padding:16px;box-shadow:var(--shadow)}.ps-card.wide{grid-column:1/-1}.ps-card h4{margin:0 0 12px;color:var(--navy)}
    .ps-bars{display:grid;gap:8px}.ps-row{display:grid;grid-template-columns:minmax(82px,120px) 1fr 46px;gap:8px;align-items:center;font-size:12px}.ps-track{height:13px;border-radius:999px;background:#edf0f2;overflow:hidden}.ps-fill{height:100%;border-radius:999px;background:linear-gradient(90deg,#c8304a,#e47b8c)}.ps-fill.navy{background:linear-gradient(90deg,#18304a,#4b81aa)}.ps-fill.gold{background:linear-gradient(90deg,#9a6a0b,#d3a22b)}
    .ps-donut-wrap{display:grid;grid-template-columns:150px 1fr;gap:16px;align-items:center}.ps-donut{width:145px;height:145px;border-radius:50%;display:grid;place-items:center;position:relative}.ps-donut:after{content:'';position:absolute;width:82px;height:82px;border-radius:50%;background:#fff}.ps-donut strong{z-index:1;color:var(--navy);font-size:20px}.ps-legend{display:grid;gap:8px}.ps-legend div{display:flex;justify-content:space-between;border-bottom:1px dashed var(--line);padding-bottom:6px}
    .ps-summary{padding:13px;border-radius:13px;background:#fff8e8;border:1px solid #ecd79e;line-height:1.75;font-size:13px}.ps-warning{padding:12px;border-radius:13px;background:#fff4e5;color:#7a5917;font-size:12px}
    @media(max-width:900px){.ps-kpis{grid-template-columns:repeat(3,1fr)}}@media(max-width:700px){.ps-kpis{grid-template-columns:repeat(2,1fr)}.ps-grid{grid-template-columns:1fr}.ps-card.wide{grid-column:auto}.ps-donut-wrap{grid-template-columns:1fr;justify-items:center}.ps-legend{width:100%}}
    `;document.head.appendChild(s);
  }
  function countBy(arr,fn){const o={};arr.forEach(x=>{const k=fn(x);if(Array.isArray(k)){k.filter(Boolean).forEach(v=>o[v]=(o[v]||0)+1)}else if(k)o[k]=(o[k]||0)+1});return o}
  function bars(entries,cls=''){entries=Array.isArray(entries)?entries:Object.entries(entries||{});if(!entries.length)return '<div class="ps-warning">표시할 자료가 없습니다.</div>';const max=Math.max(1,...entries.map(x=>Number(x[1])||0));return `<div class="ps-bars">${entries.map(([k,v])=>`<div class="ps-row"><span>${esc(k)}</span><div class="ps-track"><div class="ps-fill ${cls}" style="width:${(Number(v)||0)/max*100}%"></div></div><b>${Number(v)||0}명</b></div>`).join('')}</div>`}
  function ensureRoot(){const page=$('#statsPage');if(!page)return null;let root=$('#pastoralStatsDashboard');if(!root){root=document.createElement('section');root.id='pastoralStatsDashboard';const weekly=$('#weeklyAttendanceCard');if(weekly)weekly.insertAdjacentElement('afterend',root);else page.appendChild(root)}return root}
  function render(){
    ensureStyles();const root=ensureRoot();if(!root)return false;
    const D=window.APP_DATA||{}, members=Array.isArray(D.members)?D.members:[], families=Array.isArray(D.families)?D.families:[];
    if(!members.length){root.innerHTML='<div class="ps-card"><h3>📊 목양 통계 대시보드</h3><div class="ps-warning">군우명단을 불러오는 중입니다. 잠시 후 다시 표시됩니다.</div></div>';return false}
    const known=members.filter(m=>Number.isFinite(ageOf(m))), ages=known.map(ageOf), avg=ages.length?ages.reduce((a,b)=>a+b,0)/ages.length:0, oldest=ages.length?Math.max(...ages):0, youngest=ages.length?Math.min(...ages):0;
    const ageBands=[['0~12세',a=>a<=12],['13~19세',a=>a>=13&&a<=19],['20~29세',a=>a>=20&&a<=29],['30~39세',a=>a>=30&&a<=39],['40~49세',a=>a>=40&&a<=49],['50~59세',a=>a>=50&&a<=59],['60~69세',a=>a>=60&&a<=69],['70~79세',a=>a>=70&&a<=79],['80세 이상',a=>a>=80]].map(([k,f])=>[k,ages.filter(f).length]);
    const pos=Object.entries(countBy(members,m=>m.position||'미기재')).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'ko'));
    const groups=Object.entries(countBy(members,m=>[m.group1,m.group2,m.group3,m.group4])).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'ko'));
    const gender=countBy(members,m=>m.gender||'미기재'), male=gender['남']||0,female=gender['여']||0,unknown=members.length-male-female,total=Math.max(1,members.length),md=male/total*360,fd=(male+female)/total*360;
    const familySizes=[['1인 가정',families.filter(f=>f.size===1).length],['2인 가정',families.filter(f=>f.size===2).length],['3인 가정',families.filter(f=>f.size===3).length],['4인 가정',families.filter(f=>f.size===4).length],['5인 이상',families.filter(f=>f.size>=5).length]];
    const senior=known.filter(m=>ageOf(m)>=65).length,next=known.filter(m=>ageOf(m)<=29).length;
    root.innerHTML=`<div class="ps-title"><div><h3>📊 목양 통계 대시보드</h3><p class="stats-note">최신 군우명단 Excel의 출생연도·성별·직분·소속을 자동 집계합니다.</p></div><span class="ps-badge">GESMS ${VERSION}</span></div>
    <div class="ps-kpis"><div class="ps-kpi"><span>총 군우</span><b>${members.length}명</b></div><div class="ps-kpi"><span>총 가정</span><b>${families.length}가정</b></div><div class="ps-kpi"><span>평균 연령</span><b>${avg.toFixed(1)}세</b></div><div class="ps-kpi"><span>최고령</span><b>${oldest}세</b></div><div class="ps-kpi"><span>최연소</span><b>${youngest}세</b></div><div class="ps-kpi"><span>평균 가족원</span><b>${families.length?(members.length/families.length).toFixed(1):0}명</b></div></div>
    <div class="ps-grid">
    <article class="ps-card wide"><h4>연령대별 인원</h4>${bars(ageBands)}</article>
    <article class="ps-card"><h4>남녀 비율</h4><div class="ps-donut-wrap"><div class="ps-donut" style="background:conic-gradient(#18304a 0 ${md}deg,#c8304a ${md}deg ${fd}deg,#d8d8d8 ${fd}deg 360deg)"><strong>${members.length}명</strong></div><div class="ps-legend"><div><span>남성</span><b>${male}명 · ${pct(male,members.length)}%</b></div><div><span>여성</span><b>${female}명 · ${pct(female,members.length)}%</b></div>${unknown?`<div><span>미기재</span><b>${unknown}명</b></div>`:''}</div></div></article>
    <article class="ps-card"><h4>시니어·다음세대</h4>${bars([['65세 이상',senior],['70세 이상',known.filter(m=>ageOf(m)>=70).length],['80세 이상',known.filter(m=>ageOf(m)>=80).length],['0~29세',next],['0~12세',known.filter(m=>ageOf(m)<=12).length],['13~19세',known.filter(m=>ageOf(m)>=13&&ageOf(m)<=19).length]],'gold')}</article>
    <article class="ps-card"><h4>가족 규모</h4>${bars(familySizes,'navy')}</article>
    <article class="ps-card"><h4>직분별 인원</h4>${bars(pos.slice(0,20),'navy')}</article>
    <article class="ps-card"><h4>소속별 인원</h4>${bars(groups.slice(0,24),'gold')}</article>
    <article class="ps-card wide"><h4>목양 요약</h4><div class="ps-summary">전체 군우 <b>${members.length}명</b>, 전체 가정 <b>${families.length}가정</b>입니다.<br>출생연도가 입력된 <b>${known.length}명</b>의 평균 연령은 <b>${avg.toFixed(1)}세</b>입니다.<br>65세 이상은 <b>${senior}명(${pct(senior,known.length)}%)</b>, 다음세대 0~29세는 <b>${next}명(${pct(next,known.length)}%)</b>입니다.<br>남성 <b>${male}명</b>, 여성 <b>${female}명</b>입니다.</div>${known.length<members.length?`<div class="ps-warning" style="margin-top:10px">출생연도 미입력 ${members.length-known.length}명은 연령 통계에서 제외되었습니다.</div>`:''}</article>
    </div>`;
    return true;
  }
  function retry(){let n=0;const go=()=>{if(render())return;if(n++<40)setTimeout(go,250)};go()}
  document.addEventListener('DOMContentLoaded',retry);window.addEventListener('gesms-members-loaded',()=>setTimeout(render,20));
  document.addEventListener('click',e=>{const b=e.target.closest('[data-page="statsPage"]');if(b)setTimeout(render,30)});
  window.GESMS_RENDER_PASTORAL_STATS=render;
})();
