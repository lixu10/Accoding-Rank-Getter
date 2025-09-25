// ==UserScript==
// @name         BUAA Accoding Rank 一键导出（题目统计修正：仅AC + 首杀；色阶仅AC）
// @namespace    https://accoding.buaa.edu.cn/
// @version      3.1.0
// @description  在 /contest-ng/#/{id}/rank 抓取 /api/contests/{id}/rank 导出 Excel；题目统计严格以 AC 为“通过”；最快通过=首杀选手；色阶仅对 AC 的通过时间生效；无网自动走 .xls（含色阶）
// @author       you
// @license      MIT
//
// @match        https://accoding.buaa.edu.cn/*
// @run-at       document-end
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// @connect      accoding.buaa.edu.cn
// @noframes
//
// --- 优先 xlsx；不通时自动 .xls（离线） ---
// @require      https://cdn.bootcdn.net/ajax/libs/xlsx/0.19.3/xlsx.full.min.js
// @require      https://unpkg.bytedance.com/xlsx@0.19.3/dist/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js
// @require      https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.19.3/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  // ---------- 样式 ----------
  GM_addStyle(`
    .acc-rank-export { position: fixed; right: 16px; bottom: 16px; z-index: 2147483647; display: flex; flex-direction: column; gap: 8px; }
    .acc-rank-export button { background: #1f6feb; color: #fff; border: none; border-radius: 10px; padding: 9px 12px; font-size: 14px; cursor: pointer; box-shadow: 0 6px 18px rgba(0,0,0,.18); }
    .acc-rank-export button:hover{ filter: brightness(1.05); }
    .acc-rank-toast{ position: fixed; right: 16px; bottom: 96px; z-index: 2147483647; background: rgba(0,0,0,.9); color: #fff; padding: 8px 10px; border-radius: 8px; font-size: 12px; display: none; max-width: 60vw; word-break: break-all; }
  `);

  // ---------- 工具 ----------
  function waitForBody(){ return new Promise(r=>{ if(document.body) return r(); new MutationObserver((m,obs)=>{ if(document.body){ obs.disconnect(); r(); }}).observe(document.documentElement||document,{childList:true,subtree:true}); }); }
  let toastTimer=null, toastEl=null;
  function showToast(msg, dur=2400){ if(!toastEl) return; toastEl.textContent=msg; toastEl.style.display='block'; clearTimeout(toastTimer); toastTimer=setTimeout(()=>toastEl.style.display='none',dur); }
  const pad2 = n => String(n).padStart(2,'0');
  function msToHMS(ms){ if(ms==null||isNaN(ms))return ''; const s=Math.floor(ms/1000); const h=Math.floor(s/3600); const m=Math.floor((s%3600)/60); const ss=s%60; return `${pad2(h)}:${pad2(m)}:${pad2(ss)}`; }
  function median(arr){ if(!arr.length) return null; const b=[...arr].sort((a,b)=>a-b); const n=b.length; return n%2? b[(n-1)/2] : (b[n/2-1]+b[n/2])/2; }
  const sum = a => a.reduce((x,y)=>x+y,0);

  // ---------- 路由/API ----------
  function getContestIdFromUrl(){
    const h=location.hash||'';
    let m=h.match(/#\/(\d+)\/rank\b/); if(m) return m[1];
    m=h.match(/[?&](?:contestId|id)=(\d+)/i); if(m) return m[1];
    m=location.pathname.match(/contests\/(\d+)\/rank/); return m?m[1]:null;
  }
  function getRankApiUrl(id){ return `${location.origin}/api/contests/${id}/rank`; }

  // ---------- JSON 提取 ----------
  function looksLikeRankItem(o){
    if(!o||typeof o!=='object')return false;
    return !!(o.user && o.detail) || 'score' in o || 'penalty' in o;
  }
  function extractRankArray(json, depth=0){
    if(Array.isArray(json)) return json.some(looksLikeRankItem)? json : null;
    if(typeof json==='string'){
      const s=json.trim();
      if((s.startsWith('[')&&s.endsWith(']'))||(s.startsWith('{')&&s.endsWith('}'))){
        try{ return extractRankArray(JSON.parse(s), depth+1);}catch{}
      }
      const i=s.indexOf('['), j=s.lastIndexOf(']'); if(i!==-1&&j>i){ try{ return extractRankArray(JSON.parse(s.slice(i,j+1)), depth+1);}catch{} }
      return null;
    }
    if(json && typeof json==='object' && depth<8){
      const pref=['data','result','results','items','list','payload','content','records'];
      for(const k of pref) if(k in json){ const got=extractRankArray(json[k], depth+1); if(got) return got; }
      for(const k of Object.keys(json)){ const got=extractRankArray(json[k], depth+1); if(got) return got; }
    }
    return null;
  }

  // ---------- 请求 ----------
  function getJSONViaGM(url){
    return new Promise((resolve,reject)=>{
      GM_xmlhttpRequest({
        method:'GET', url,
        headers:{'Accept':'application/json,text/plain,*/*','Cache-Control':'no-cache'},
        timeout:20000,
        onload:(res)=>{
          let text=res.responseText||'';
          if(res.status<200||res.status>=300) return reject(new Error(`HTTP ${res.status}`));
          if(text.charCodeAt(0)===0xFEFF) text=text.slice(1);
          text=text.replace(/^\)\]\}',?\s*/, '');
          try{ resolve(JSON.parse(text)); }
          catch(e){
            const i=text.indexOf('['), j=text.lastIndexOf(']');
            if(i!==-1 && j>i){ try{ return resolve(JSON.parse(text.slice(i,j+1))); }catch{} }
            reject(new Error('JSON 解析失败'));
          }
        },
        onerror:()=>reject(new Error('网络错误')),
        ontimeout:()=>reject(new Error('请求超时')),
      });
    });
  }

  // ---------- 展平 Rank + 中文表头（保持不变） ----------
  function flattenRank(list){
    const probSet=new Set();
    for(const it of list){ const det=it?.detail; if(det&&typeof det==='object'){ for(const k of Object.keys(det)) probSet.add(k); } }
    const problems=Array.from(probSet).sort((a,b)=>{
      const A=/^[A-Z]$/.test(a), B=/^[A-Z]$/.test(b);
      if(A&&B) return a.localeCompare(b);
      if(A) return -1;
      if(B) return 1;
      return a.localeCompare(b);
    });

    const headers = ['排名','用户ID','昵称','学号','学校','总分','罚时(毫秒)','罚时(hh:mm:ss)'];
    for(const p of problems){
      headers.push(
        `${p} 分数`, `${p} 错题数`, `${p} 结果`,
        `${p} 通过时间(毫秒)`, `${p} 通过时间(hh:mm:ss)`, `${p} 首杀`
      );
    }

    const rows = list.map((item, idx)=>{
      const det=item?.detail||{};
      const row={
        '排名': idx+1,
        '用户ID': item?.user?.id ?? '',
        '昵称': item?.user?.nickname ?? '',
        '学号': item?.user?.student_id ?? '',
        '学校': item?.user?.school ?? '',
        '总分': item?.score ?? '',
        '罚时(毫秒)': item?.penalty ?? '',
        '罚时(hh:mm:ss)': msToHMS(item?.penalty ?? '')
      };
      for(const p of problems){
        const d=det[p]||{};
        row[`${p} 分数`] = d.score ?? '';
        row[`${p} 错题数`] = d.wrong_count ?? '';
        row[`${p} 结果`] = d.result ?? '';
        row[`${p} 通过时间(毫秒)`] = (d.result==='AC' && typeof d.accepted_time==='number') ? d.accepted_time : ''; // 仅AC保留时间
        row[`${p} 通过时间(hh:mm:ss)`] = (d.result==='AC' && typeof d.accepted_time==='number') ? msToHMS(d.accepted_time) : '';
        row[`${p} 首杀`] = d.first_blood ? '是' : '';
      }
      return row;
    });

    return { headers, rows, problems };
  }

  // ---------- 题目统计（严格按 AC） ----------
  function buildProblemStats(list, problems){
    const total = list.length;
    const stats = [];
    for(const p of problems){
      let attempted=0;         // detail里存在该题即“有记录/做过”
      let solved=0;            // 仅 result==='AC'
      let fb=0;                // 首杀标记计数（通常为1或0）
      let scoreSumAll=0;       // 全体平均分
      let scoreSumAtt=0;       // 做题者平均分
      let wrongSumAtt=0;       // 做题者平均错题
      let attCount=0;
      const acTimes=[];        // 仅 AC 的通过时间
      let fastestUser='';      // 首杀选手（如有）
      let fastestTime=null;

      for(const it of list){
        const det=it?.detail||{}; const d=det[p];
        const user = it?.user?.nickname || it?.user?.id || '';
        if(!d){ scoreSumAll += 0; continue; }

        attempted++; attCount++;
        const score = (typeof d.score==='number') ? d.score : 0;
        const wrong = (typeof d.wrong_count==='number') ? d.wrong_count : 0;
        scoreSumAll += score;
        scoreSumAtt += score;
        wrongSumAtt += wrong;

        if (d.first_blood) { fb++; if(!fastestUser){ fastestUser=user; fastestTime = (typeof d.accepted_time==='number') ? d.accepted_time : null; } }
        if (d.result === 'AC') {
          solved++;
          if (typeof d.accepted_time==='number') acTimes.push(d.accepted_time);
        }
      }

      const avgScoreAll = total? (scoreSumAll/total):0;
      const avgScoreAtt = attCount? (scoreSumAtt/attCount):0;
      const avgWrongAtt = attCount? (wrongSumAtt/attCount):0;
      const passAll = total? (solved/total):0;
      const passAtt = attempted? (solved/attempted):0;
      const avgAc = acTimes.length? (sum(acTimes)/acTimes.length) : null;
      const medAc = acTimes.length? median(acTimes) : null;

      stats.push({
        '题目': p,
        '总人数': total,
        '有记录人数': attempted,
        '通过人数(AC)': solved,
        '通过率（总）': passAll,
        '通过率（做题者）': passAtt,
        '平均得分（全体）': avgScoreAll,
        '平均得分（做题者）': avgScoreAtt,
        '平均错题（做题者）': avgWrongAtt,
        '首杀数': fb,
        '最快通过选手(首杀)': fastestUser,
        '最快通过时间(hh:mm:ss)': fastestTime!=null ? msToHMS(fastestTime) : '',
        '平均通过时间(hh:mm:ss)': avgAc!=null ? msToHMS(avgAc) : '',
        '中位通过时间(hh:mm:ss)': medAc!=null ? msToHMS(medAc) : ''
      });
    }
    return stats;
  }

  // ---------- 分数分布（总分） ----------
  function buildScoreDistribution(list){
    const scores = list.map(x=> (typeof x.score==='number'? x.score : Number(x.score)||0) );
    const max = Math.max(100, Math.ceil(Math.max(...scores,0)/20)*20);
    const buckets = [];
    for(let lo=0; lo<=max; lo+=20){
      const hi = lo+20;
      const cnt = scores.filter(s => (s>=lo && (hi===max+20? s<=max : s<hi))).length;
      buckets.push({'分数下限':lo,'分数上限':Math.min(hi,max),'人数':cnt});
    }
    return buckets;
  }

  // ---------- 保存 ----------
  function saveBlob(data, mime, filename){
    const blob = new Blob([data], {type:mime});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
  }

  // ---------- xlsx（开源 SheetJS，不写样式） ----------
  function saveXLSX({ headers, rows, problems, meta, stats, dist }){
    if (typeof XLSX==='undefined') return false;
    try{
      const wb=XLSX.utils.book_new();

      // 工作表1：榜单
      const wsRank=XLSX.utils.json_to_sheet(rows,{header:headers});
      wsRank['!cols']=headers.map(h=>({wch: Math.min(Math.max(10,String(h).length+2),28)}));
      XLSX.utils.book_append_sheet(wb, wsRank, '榜单');

      // 工作表2：题目统计
      const statsHeaders = Object.keys(stats[0]||{
        '题目':'','总人数':'','有记录人数':'','通过人数(AC)':'',
        '通过率（总）':'','通过率（做题者）':'','平均得分（全体）':'',
        '平均得分（做题者）':'','平均错题（做题者）':'','首杀数':'',
        '最快通过选手(首杀)':'','最快通过时间(hh:mm:ss)':'',
        '平均通过时间(hh:mm:ss)':'','中位通过时间(hh:mm:ss)':''
      });
      const wsStats=XLSX.utils.json_to_sheet(stats,{header:statsHeaders});
      wsStats['!cols']=statsHeaders.map(h=>({wch: Math.min(Math.max(10,String(h).length+2),24)}));
      XLSX.utils.book_append_sheet(wb, wsStats, '题目统计');

      // 工作表3：分数分布
      const wsDist=XLSX.utils.json_to_sheet(buildScoreDistribution(meta.rawList),{header:['分数下限','分数上限','人数']});
      wsDist['!cols']=[{wch:12},{wch:12},{wch:8}];
      XLSX.utils.book_append_sheet(wb, wsDist, '分数分布');

      // 工作表4：元信息
      const wsMeta=XLSX.utils.json_to_sheet([
        {键:'比赛ID', 值: meta.contestId},
        {键:'生成时间', 值: new Date().toISOString()},
        {键:'题目列表', 值: meta.problems.join(', ')},
        {键:'数据来源', 值: meta.source},
        {键:'样式说明', 值: 'xlsx（SheetJS 开源）不写单元格样式；色阶请在 Excel 里套“条件格式→色阶”。'}
      ],{header:['键','值']});
      wsMeta['!cols']=[{wch:12},{wch:80}];
      XLSX.utils.book_append_sheet(wb, wsMeta, '元信息');

      const fname=`contest_${meta.contestId}_rank_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xlsx`;
      XLSX.writeFile(wb, fname);
      return true;
    }catch(e){ console.error('xlsx 写入失败',e); return false; }
  }

  // ---------- 离线 .xls（SpreadsheetML，色阶仅对 AC） ----------
  function xmlEsc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;'); }
  function buildTableXML(sheetName, headers, rows, cellStyleFn){
    const cols=headers.map(()=> '<Column ss:AutoFitWidth="1"/>').join('');
    const headerRow = `<Row>` + headers.map(h=>`<Cell><Data ss:Type="String">${xmlEsc(h)}</Data></Cell>`).join('') + `</Row>`;
    const dataRows = rows.map(r=>{
      return `<Row>` + headers.map((h,ci)=>{
        const v=r[h];
        let styleId='';
        if (cellStyleFn) styleId = cellStyleFn(h, v, r, ci) || '';
        const t = (typeof v==='number' || (/^-?\d+(\.\d+)?$/.test(String(v)))) ? 'Number' : 'String';
        const cell = (v===''||v==null) ? '<Cell/>' : `<Cell${styleId?` ss:StyleID="${styleId}"`:''}><Data ss:Type="${t}">${xmlEsc(v)}</Data></Cell>`;
        return cell;
      }).join('') + `</Row>`;
    }).join('');
    return `
      <Worksheet ss:Name="${xmlEsc(sheetName)}">
        <Table>${cols}${headerRow}${dataRows}</Table>
      </Worksheet>
    `;
  }
  function colorHex(r,g,b){ return [r,g,b].map(x=> Math.max(0,Math.min(255,Math.round(x))).toString(16).padStart(2,'0')).join('').toUpperCase(); }

  function saveXLS_Offline({ headers, rows, problems, meta, stats }){
    // 仅对“AC 的通过时间(毫秒)”建立色阶
    const styleDefs = [];
    const probTimeRange = {};
    for(const p of problems){
      const keyTime = `${p} 通过时间(毫秒)`;
      const keyRes  = `${p} 结果`;
      let min=Infinity, max=-Infinity, has=false;
      for(const r of rows){
        if(r[keyRes]!=='AC') continue; // 只看 AC
        const t=r[keyTime];
        if(typeof t==='number'){
          has=true;
          if(t<min)min=t;
          if(t>max)max=t;
        }
      }
      if(has && min===max){ max=min+1; } // 避免除0
      if(has) probTimeRange[p]={min,max};
      // 10 档绿色色阶（越早越深）
      for(let i=0;i<10;i++){
        const base=0x30;
        const light = base + Math.floor(i*(200/9));
        const fill = colorHex(0, 160, light);
        styleDefs.push(`<Style ss:ID="P_${xmlEsc(p)}_${i}"><Interior ss:Color="#${fill}" ss:Pattern="Solid"/></Style>`);
      }
    }
    const rankCellStyleFn = (h, v, r)=>{
      const m=h.match(/^([A-Z])\s+通过时间\(毫秒\)$/);
      if(!m) return '';
      const p=m[1];
      // 非 AC 不上色
      if(r[`${p} 结果`]!=='AC') return '';
      const rng=probTimeRange[p];
      if(!rng) return '';
      if(typeof v!=='number') return '';
      const {min,max}=rng;
      const t=(v-min)/(max-min);
      const bucket=Math.max(0, Math.min(9, 9-Math.floor(t*9))); // 越早越深
      return `P_${p}_${bucket}`;
    };

    // 4 张表
    const rankXML  = buildTableXML('榜单', headers, rows, (h,v,r)=>rankCellStyleFn(h,v,r));
    const statsHdr = Object.keys(stats[0]||{
      '题目':'','总人数':'','有记录人数':'','通过人数(AC)':'',
      '通过率（总）':'','通过率（做题者）':'','平均得分（全体）':'',
      '平均得分（做题者）':'','平均错题（做题者）':'','首杀数':'',
      '最快通过选手(首杀)':'','最快通过时间(hh:mm:ss)':'',
      '平均通过时间(hh:mm:ss)':'','中位通过时间(hh:mm:ss)':''
    });
    const statsXML = buildTableXML('题目统计', statsHdr, stats);
    const distXML  = buildTableXML('分数分布', ['分数下限','分数上限','人数'], buildScoreDistribution(meta.rawList));
    const metaXML  = buildTableXML('元信息', ['键','值'], [
      {'键':'比赛ID', '值': meta.contestId},
      {'键':'生成时间', '值': new Date().toISOString()},
      {'键':'题目列表', '值': meta.problems.join(', ')},
      {'键':'数据来源', '值': meta.source},
      {'键':'样式说明', '值': '离线 .xls 含“AC 通过时间色阶”，xlsx（开源 SheetJS）不含样式'}
    ]);

    const xml = `<?xml version="1.0"?>
      <?mso-application progid="Excel.Sheet"?>
      <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:html="http://www.w3.org/TR/REC-html40">
        <Styles>
          <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/></Style>
          ${styleDefs.join('')}
        </Styles>
        ${rankXML}
        ${statsXML}
        ${distXML}
        ${metaXML}
      </Workbook>`;
    const fname=`contest_${meta.contestId}_rank_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xls`;
    saveBlob(xml, 'application/vnd.ms-excel', fname);
    return true;
  }

  // ---------- UI ----------
  let uiBox, btn, toast, lastAutoKey='';
  function wait(ms){ return new Promise(r=>setTimeout(r,ms)); }
  async function ensureUI(){
    await waitForBody();
    if(document.getElementById('acc-rank-export-root')) return;
    uiBox=document.createElement('div'); uiBox.id='acc-rank-export-root'; uiBox.className='acc-rank-export';
    btn=document.createElement('button'); btn.textContent='导出过题榜（含统计）';
    toast=document.createElement('div'); toast.className='acc-rank-toast'; toastEl=toast;
    uiBox.appendChild(btn); document.body.appendChild(uiBox); document.body.appendChild(toast);
    btn.addEventListener('click', ()=> exportCurrent(false));
  }

  // ---------- 主流程 ----------
  async function exportCurrent(auto=true){
    const id=getContestIdFromUrl();
    if(!id){ if(!auto) alert('未识别到 contestId（请确认当前是 /#/xxxx/rank 页面）'); return; }
    const key=`${id}:${auto?'auto':'manual'}`; if(auto && lastAutoKey===key) return; lastAutoKey=key;

    try{
      showToast('获取榜单数据…');
      const api=getRankApiUrl(id);
      const first=await getJSONViaGM(api);
      const arr=extractRankArray(first);
      if(!Array.isArray(arr)) throw new Error('响应不是数组（未在对象/字符串中找到 rank 列表）');
      if(arr.length===0){ showToast('数据为空'); return; }

      // 展平 + 统计（严格按 AC）
      const { headers, rows, problems } = flattenRank(arr);
      const stats = buildProblemStats(arr, problems);
      const dist  = buildScoreDistribution(arr);
      const meta  = { contestId:id, problems, source:'api', rawList:arr };

      // xlsx 优先（无样式），失败则 .xls（含色阶）
      if (saveXLSX({ headers, rows, problems, meta, stats, dist })) {
        showToast('导出完成（.xlsx；统计按 AC；色阶请在 Excel 中套条件格式）');
        return;
      }
      showToast('CDN 不通，启用离线 .xls（含 AC 色阶）…', 1800);
      saveXLS_Offline({ headers, rows, problems, meta, stats });
      showToast('导出完成（.xls 离线，含 AC 色阶高亮）');
    }catch(e){
      console.error(e);
      alert('导出失败：' + (e?.message||e));
    }
  }

  function onHashChangeMaybeExport(){
    const h=location.hash||'';
    if (/#\/\d+\/rank\b/.test(h) || /[?&](contestId|id)=\d+/.test(h)) {
      setTimeout(()=>exportCurrent(true), 500);
    }
  }

  (async function bootstrap(){
    await ensureUI();
    onHashChangeMaybeExport();
    window.addEventListener('hashchange', onHashChangeMaybeExport);
  })();
})();
