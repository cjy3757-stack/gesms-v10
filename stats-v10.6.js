/* GESMS V10.6 - 목양 통계 대시보드 */
(function(){
  'use strict';
  const VERSION='V10.6';
  const $=s=>document.querySelector(s);
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const currentYear=()=>new Date().getFullYear();
  const ageOf=m=>m.birthYear?currentYear()-Number(m.birthYear):null;
  const pct=(n,d)=>d?Math.round(n/d*100):0;

  function addStyles(){
    if($('#gesmsV106Style')) return;
    const style=document.createElement('style');
    style.id='gesmsV106Style';
    style.textContent=`
      .v106-dashboard{margin-bottom:18px}
      .v106-title{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap;margin-bottom:14px}
      .v106-title h3{margin:0;color:var(--navy)}
      .v106-badge{background:#fff0f3;color:var(--red);font-weight:900;border-radius:999px;padding:7px 11px;font-size:12px}
      .v106-kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}
      .v106-kpi{background:#fff;border:1px solid var(--line);border-radius:16px;padding:14px;box-shadow:var(--shadow)}
      .v106-kpi span{display:block;font-size:12px;color:var(--muted)}
      .v106-kpi b{display:block;color:var(--red);font-size:24px;margin-top:4px}
      .v106-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
      .v106-card{background:#fff;border:1px solid var(--line);border-radius:19px;padding:17px;box-shadow:var(--shadow)}
      .v106-card.wide{grid-column:1/-1}
      .v106-card h4{margin:0 0 13px;color:var(--navy);font-size:17px}
      .v106-bars{display:grid;gap:9px}
      .v106-row{display:grid;grid-template-columns:92px 1fr 48px;gap:8px;align-items:center;font-size:12px}
      .v106-track{height:14px;background:#eee8e2;border-radius:999px;overflow:hidden}
      .v106-fill{height:100%;background:linear-gradient(90deg,var(--red),var(--red2));border-radius:999px;min-width:0}
      .v106-fill.navy{background:linear-gradient(90deg,var(--navy),#4b81aa)}
      .v106-fill.gold{background:linear-gradient(90deg,#9a6a0b,#d3a22b)}
      .v106-donut-wrap{display:grid;grid-template-columns:170px 1fr;gap:18px;align-items:center}
      .v106-donut{width:160px;height:160px;border-radius:50%;display:grid;place-items:center;position:relative}
      .v106-donut:after{content:'';width:92px;height:92px;background:#fff;border-radius:50%;position:absolute}
      .v106-donut strong{z-index:1;color:var(--navy);font-size:21px}
      .v106-legend{display:grid;gap:9px}.v106-legend div{display:flex;justify-content:space-between;gap:10px;border-bottom:1px dashed var(--line);padding-bottom:7px}
      .v106-summary{background:#fff8e8;border:1px solid #ecd79e;border-radius:15px;padding:14px;line-height:1.75;font-size:13px}
      .v106-missing{margin-top:12px;color:#7a5917;background:#fff4e5;border-radius:12px;padding:10px 12px;font-size:12px}
      @media(max-width:900px){.v106-kpis{grid-template-columns:repeat(3,1fr)}}
      @media(max-width:700px){.v106-grid{grid-template-columns:1fr}.v106-card.wide{grid-column:auto}.v106-kpis{grid-template-columns:repeat(2,1fr)}.v106-donut-wrap{grid-template-columns:1fr;justify-items:center}.v106-legend{width:100%}}
    `;
    document.head.appendChild(style);
  }

  function countsBy(arr,keyFn){
    const o={};
    arr.forEach(x=>{const k=keyFn(x); if(k)o[k]=(o[k]||0)+1;});
    return o;
  }

  function bars(data,classes=''){
    const entries=Array.isArray(data)?data:Object.entries(data);
    const max=Math.max(1,...entries.map(x=>Number(x[1])||0));
    return `<div class="v106-bars">${entries.map(([label,n])=>`
      <div class="v106-row"><span>${esc(label)}</span><div class="v106-track"><div class="v106-fill ${classes}" style="width:${(n/max)*100}%"></div></div><b>${n}명</b></div>
    `).join('')}</div>`;
  }

  function render(){
    const page=$('#statsPage');
    const D=window.APP_DATA||{};
    const members=Array.isArray(D.members)?D.members:[];
    const families=Array.isArray(D.families)?D.families:[];
    if(!page||!members.length) return false;

    addStyles();
    let root=$('#v106Dashboard');
    if(!root){
      root=document.createElement('section');
      root.id='v106Dashboard';
      root.className='v106-dashboard';
      const title=page.querySelector('.section-title');
      if(title) title.insertAdjacentElement('afterend',root);
      else page.prepend(root);
    }

    const known=members.filter(m=>Number.isFinite(ageOf(m)));
    const ages=known.map(ageOf);
    const average=ages.length?ages.reduce((a,b)=>a+b,0)/ages.length:0;
    const oldest=ages.length?Math.max(...ages):0;
    const youngest=ages.length?Math.min(...ages):0;
    const seniors=known.filter(m=>ageOf(m)>=65).length;
    const nextGen=known.filter(m=>ageOf(m)<=29).length;

    const ageBands=[
      ['0~12세',a=>a<=12],['13~19세',a=>a>=13&&a<=19],['20~29세',a=>a>=20&&a<=29],
      ['30~39세',a=>a>=30&&a<=39],['40~49세',a=>a>=40&&a<=49],['50~59세',a=>a>=50&&a<=59],
      ['60~69세',a=>a>=60&&a<=69],['70~79세',a=>a>=70&&a<=79],['80세 이상',a=>a>=80]
    ].map(([label,test])=>[label,ages.filter(test).length]);

    const gender=countsBy(members,m=>m.gender||'미기재');
    const male=gender['남']||0, female=gender['여']||0, unknownGender=members.length-male-female;
    const genderTotal=Math.max(1,members.length);
    const maleDeg=male/genderTotal*360;
    const femaleDeg=(male+female)/genderTotal*360;
    const donut=`conic-gradient(#18304a 0deg ${maleDeg}deg,#c8304a ${maleDeg}deg ${femaleDeg}deg,#d8d8d8 ${femaleDeg}deg 360deg)`;

    const familySizes=[
      ['1인 가정',families.filter(f=>f.size===1).length],
      ['2인 가정',families.filter(f=>f.size===2).length],
      ['3인 가정',families.filter(f=>f.size===3).length],
      ['4인 가정',families.filter(f=>f.size===4).length],
      ['5인 이상',families.filter(f=>f.size>=5).length]
    ];

    const positions=Object.entries(countsBy(members,m=>m.position||'미기재')).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'ko'));
    const groups=Object.entries(countsBy(members,m=>[m.group1,m.group2,m.group3,m.group4].filter(Boolean))).flatMap(()=>[]);
    const groupCounts={};
    members.forEach(m=>[m.group1,m.group2,m.group3,m.group4].filter(Boolean).forEach(g=>groupCounts[g]=(groupCounts[g]||0)+1));
    const groupEntries=Object.entries(groupCounts).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'ko'));

    const birthMonths=Array.from({length:12},(_,i)=>[`${i+1}월`,members.filter(m=>Number((String(m.birthday||'').match(/(\d{1,2})월/)||[])[1])===i+1).length]);

    root.innerHTML=`
      <div class="v106-title"><div><h3>📊 목양 통계 대시보드</h3><div class="stats-note">출생연도·성별이 포함된 군우명단 Excel을 기준으로 자동 계산합니다.</div></div><span class="v106-badge">GESMS ${VERSION}</span></div>
      <div class="v106-kpis">
        <div class="v106-kpi"><span>총 군우</span><b>${members.length}명</b></div>
        <div class="v106-kpi"><span>총 가정</span><b>${families.length}가정</b></div>
        <div class="v106-kpi"><span>평균 연령</span><b>${average.toFixed(1)}세</b></div>
        <div class="v106-kpi"><span>최고령</span><b>${oldest}세</b></div>
        <div class="v106-kpi"><span>최연소</span><b>${youngest}세</b></div>
        <div class="v106-kpi"><span>평균 가족원</span><b>${families.length?(members.length/families.length).toFixed(1):0}명</b></div>
      </div>
      <div class="v106-grid">
        <article class="v106-card wide"><h4>연령대별 군우 현황</h4>${bars(ageBands)}</article>
        <article class="v106-card"><h4>남녀 비율</h4>
          <div class="v106-donut-wrap"><div class="v106-donut" style="background:${donut}"><strong>${members.length}명</strong></div>
          <div class="v106-legend">
            <div><span>남성</span><b>${male}명 · ${pct(male,members.length)}%</b></div>
            <div><span>여성</span><b>${female}명 · ${pct(female,members.length)}%</b></div>
            ${unknownGender?`<div><span>미기재</span><b>${unknownGender}명</b></div>`:''}
          </div></div>
        </article>
        <article class="v106-card"><h4>시니어·다음세대</h4>${bars([
          ['65세 이상',seniors],['70세 이상',known.filter(m=>ageOf(m)>=70).length],['80세 이상',known.filter(m=>ageOf(m)>=80).length],
          ['0~29세',nextGen],['0~12세',known.filter(m=>ageOf(m)<=12).length],['13~19세',known.filter(m=>ageOf(m)>=13&&ageOf(m)<=19).length],['20~29세',known.filter(m=>ageOf(m)>=20&&ageOf(m)<=29).length]
        ],'gold')}</article>
        <article class="v106-card"><h4>가족 규모 분포</h4>${bars(familySizes,'navy')}</article>
        <article class="v106-card"><h4>월별 생일 분포</h4>${bars(birthMonths)}</article>
        <article class="v106-card"><h4>직분별 현황</h4>${bars(positions.slice(0,15),'navy')}</article>
        <article class="v106-card"><h4>소속별 현황</h4>${bars(groupEntries.slice(0,18),'gold')}</article>
        <article class="v106-card wide"><h4>목양 요약</h4>
          <div class="v106-summary">
            전체 군우는 <b>${members.length}명</b>, 전체 가정은 <b>${families.length}가정</b>이며 평균 가족원은 <b>${families.length?(members.length/families.length).toFixed(1):0}명</b>입니다.<br>
            출생연도가 입력된 <b>${known.length}명</b>의 평균 연령은 <b>${average.toFixed(1)}세</b>입니다.<br>
            65세 이상은 <b>${seniors}명(${pct(seniors,known.length)}%)</b>, 다음세대 0~29세는 <b>${nextGen}명(${pct(nextGen,known.length)}%)</b>입니다.<br>
            남성은 <b>${male}명</b>, 여성은 <b>${female}명</b>입니다.
          </div>
          ${known.length<members.length?`<div class="v106-missing">출생연도 미입력 ${members.length-known.length}명은 연령 통계에서 제외되었습니다.</div>`:''}
        </article>
      </div>
    `;
    return true;
  }

  let tries=0;
  function schedule(){
    if(render()) return;
    if(tries++<30) setTimeout(schedule,300);
  }
  window.addEventListener('gesms-members-loaded',()=>setTimeout(render,50));
  document.addEventListener('DOMContentLoaded',schedule);
  window.GESMS_RENDER_V106_STATS=render;
})();
