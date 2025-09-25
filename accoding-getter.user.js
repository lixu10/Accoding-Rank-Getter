// ==UserScript==
// @name         BUAA Accoding Rank 一键导出 XLS(X)（v2.4，离线 fallback 无需 CDN）
// @namespace    https://accoding.buaa.edu.cn/
// @version      2.4.0
// @description  在 /contest-ng/#/{id}/rank 页面一键抓取 /api/contests/{id}/rank 并导出；CDN 可用时导出 .xlsx，不可用时自动离线导出 .xls（SpreadsheetML）；兼容顶层字符串/包装对象/纯文本响应
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
// --- 可用就用，失败就走离线 .xls ---
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
    .acc-rank-export {
      position: fixed; right: 16px; bottom: 16px; z-index: 2147483647;
      display: flex; flex-direction: column; gap: 8px;
    }
    .acc-rank-export button {
      background: #1f6feb; color: #fff; border: none; border-radius: 10px;
      padding: 9px 12px; font-size: 14px; cursor: pointer;
      box-shadow: 0 6px 18px rgba(0,0,0,.18);
    }
    .acc-rank-export button:hover { filter: brightness(1.05); }
    .acc-rank-toast {
      position: fixed; right: 16px; bottom: 96px; z-index: 2147483647;
      background: rgba(0,0,0,.9); color: #fff; padding: 8px 10px;
      border-radius: 8px; font-size: 12px; display: none; max-width: 60vw; word-break: break-all;
    }
  `);

  // ---------- 小工具 ----------
  function waitForBody() {
    return new Promise(r => {
      if (document.body) return r();
      new MutationObserver((m,obs)=>{ if (document.body){ obs.disconnect(); r(); }})
        .observe(document.documentElement||document, {childList:true,subtree:true});
    });
  }
  let toastTimer=null, toastEl=null;
  function showToast(msg, dur=2400) {
    if (!toastEl) return;
    toastEl.textContent = msg;
    toastEl.style.display = 'block';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(()=> toastEl.style.display='none', dur);
  }
  const pad2 = n => String(n).padStart(2,'0');
  function msToHMS(ms){
    if (ms==null || isNaN(ms)) return '';
    const s = Math.floor(ms/1000);
    const h = Math.floor(s/3600);
    const m = Math.floor((s%3600)/60);
    const ss = s%60;
    return `${pad2(h)}:${pad2(m)}:${pad2(ss)}`;
  }

  // ---------- 路由 / API ----------
  function getContestIdFromUrl(){
    const h = location.hash || '';
    let m = h.match(/#\/(\d+)\/rank\b/);
    if (m) return m[1];
    m = h.match(/[?&](?:contestId|id)=(\d+)/i);
    if (m) return m[1];
    m = location.pathname.match(/contests\/(\d+)\/rank/);
    return m ? m[1] : null;
  }
  function getRankApiUrl(id){
    return `${location.origin}/api/contests/${id}/rank`;
  }

  // ---------- 响应解析 ----------
  function looksLikeRankItem(o) {
    if (!o || typeof o !== 'object') return false;
    const hasUser = o.user && typeof o.user === 'object';
    const hasDetail = o.detail && typeof o.detail === 'object';
    const maybeScore = 'score' in o;
    const maybePenalty = 'penalty' in o;
    return (hasUser && hasDetail) || (maybeScore && (hasUser || hasDetail || maybePenalty));
  }
  function extractRankArray(json, depth=0) {
    if (Array.isArray(json)) {
      return json.some(looksLikeRankItem) ? json : null;
    }
    if (typeof json === 'string') {
      const s = json.trim();
      if (s.startsWith('[') && s.endsWith(']')) {
        try { const arr = JSON.parse(s); return extractRankArray(arr, depth+1); } catch {}
      }
      if (s.startsWith('{') && s.endsWith('}')) {
        try { const obj = JSON.parse(s); return extractRankArray(obj, depth+1); } catch {}
      }
      const i = s.indexOf('['), j = s.lastIndexOf(']');
      if (i !== -1 && j > i) {
        try { const arr = JSON.parse(s.slice(i, j+1)); return extractRankArray(arr, depth+1); } catch {}
      }
      return null;
    }
    if (json && typeof json === 'object' && depth < 8) {
      const pref = ['data','result','results','items','list','payload','content','records'];
      for (const k of pref) if (k in json) {
        const got = extractRankArray(json[k], depth+1);
        if (got) return got;
      }
      for (const k of Object.keys(json)) {
        const got = extractRankArray(json[k], depth+1);
        if (got) return got;
      }
    }
    return null;
  }

  // ---------- 请求 ----------
  function getJSONViaGM(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET',
        url,
        headers: { 'Accept': 'application/json,text/plain,*/*', 'Cache-Control': 'no-cache' },
        timeout: 20000,
        onload: (res) => {
          let text = res.responseText || '';
          if (res.status < 200 || res.status >= 300) return reject(new Error(`HTTP ${res.status}`));
          if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
          text = text.replace(/^\)\]\}',?\s*/, '');
          try { resolve(JSON.parse(text)); }
          catch(e){
            const i = text.indexOf('['), j = text.lastIndexOf(']');
            if (i !== -1 && j > i) {
              try { return resolve(JSON.parse(text.slice(i, j+1))); } catch {}
            }
            reject(new Error('JSON 解析失败'));
          }
        },
        onerror: () => reject(new Error('网络错误')),
        ontimeout: () => reject(new Error('请求超时')),
      });
    });
  }

  // ---------- 整平数据为二维表 ----------
  function flattenRank(list) {
    const probSet = new Set();
    for (const it of list) {
      const det = it?.detail;
      if (det && typeof det === 'object') for (const k of Object.keys(det)) probSet.add(k);
    }
    const problems = Array.from(probSet).sort((a,b)=>{
      const A = /^[A-Z]$/.test(a), B = /^[A-Z]$/.test(b);
      if (A && B) return a.localeCompare(b);
      if (A) return -1;
      if (B) return 1;
      return a.localeCompare(b);
    });

    const headers = ['rank','user_id','nickname','student_id','school','total_score','penalty_ms','penalty_hms'];
    for (const p of problems) {
      headers.push(`${p}_score`, `${p}_wrong_count`, `${p}_result`,
                   `${p}_accepted_time_ms`, `${p}_accepted_time_hms`, `${p}_first_blood`);
    }

    const rows = list.map((item, idx) => {
      const det = item?.detail || {};
      const row = {
        rank: idx + 1,
        user_id: item?.user?.id ?? '',
        nickname: item?.user?.nickname ?? '',
        student_id: item?.user?.student_id ?? '',
        school: item?.user?.school ?? '',
        total_score: item?.score ?? '',
        penalty_ms: item?.penalty ?? '',
        penalty_hms: msToHMS(item?.penalty ?? '')
      };
      for (const p of problems) {
        const d = det[p] || {};
        row[`${p}_score`] = d.score ?? '';
        row[`${p}_wrong_count`] = d.wrong_count ?? '';
        row[`${p}_result`] = d.result ?? '';
        row[`${p}_accepted_time_ms`] = d.accepted_time ?? '';
        row[`${p}_accepted_time_hms`] = d.accepted_time != null ? msToHMS(d.accepted_time) : '';
        row[`${p}_first_blood`] = d.first_blood ? true : '';
      }
      return row;
    });

    return { headers, rows, problems };
  }

  // ---------- 文件保存 ----------
  function saveBlob(data, mime, filename){
    const blob = new Blob([data], {type: mime});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=> URL.revokeObjectURL(url), 2000);
  }

  // ---------- A. 优先用 SheetJS 生成 .xlsx ----------
  function trySaveXLSX(headers, rows, meta){
    if (typeof XLSX === 'undefined') return false;
    try {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
      ws['!cols'] = headers.map(h => ({ wch: Math.min(Math.max(10, String(h).length + 2), 30) }));
      XLSX.utils.book_append_sheet(wb, ws, 'Rank');
      const wsMeta = XLSX.utils.json_to_sheet([
        { key: 'contest_id', value: meta.contestId },
        { key: 'generated_at', value: new Date().toISOString() },
        { key: 'problems', value: meta.problems.join(', ') },
        { key: 'source', value: meta.source },
      ]);
      XLSX.utils.book_append_sheet(wb, wsMeta, 'Meta');
      const fname = `contest_${meta.contestId}_rank_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xlsx`;
      XLSX.writeFile(wb, fname);
      return true;
    } catch (e) {
      console.error('XLSX 写入失败：', e);
      return false;
    }
  }

  // ---------- B. 离线 fallback：生成 Excel 2003 XML（.xls） ----------
  function xmlEscape(s){
    return String(s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }
  function buildSpreadsheetML(sheetName, headers, rows){
    const cols = headers.map(()=> '<Column ss:AutoFitWidth="1"/>').join('');
    const headerRow = `<Row>` + headers.map(h=>`<Cell><Data ss:Type="String">${xmlEscape(h)}</Data></Cell>`).join('') + `</Row>`;
    const dataRows = rows.map(r=>{
      return `<Row>` + headers.map(h=>{
        const v = r[h];
        if (v === '' || v === null || v === undefined) return `<Cell/>`;
        const isNum = typeof v === 'number' || (/^-?\d+(\.\d+)?$/.test(String(v)));
        return `<Cell><Data ss:Type="${isNum?'Number':'String'}">${xmlEscape(v)}</Data></Cell>`;
      }).join('') + `</Row>`;
    }).join('');
    return `
      <Worksheet ss:Name="${xmlEscape(sheetName)}">
        <Table>${cols}${headerRow}${dataRows}</Table>
      </Worksheet>
    `;
  }
  function saveXLS_Offline({ headers, rows, meta }){
    const rankWS = buildSpreadsheetML('Rank', headers, rows);
    const metaHeaders = ['key','value'];
    const metaRows = [
      {key:'contest_id', value: meta.contestId},
      {key:'generated_at', value: new Date().toISOString()},
      {key:'problems', value: meta.problems.join(', ')},
      {key:'source', value: meta.source}
    ];
    const metaWS = buildSpreadsheetML('Meta', metaHeaders, metaRows);
    const xml = `<?xml version="1.0"?>
      <?mso-application progid="Excel.Sheet"?>
      <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
        xmlns:html="http://www.w3.org/TR/REC-html40">
        <Styles>
          <Style ss:ID="Default" ss:Name="Normal">
            <Alignment ss:Vertical="Center"/>
            <Borders/>
            <Font/>
            <Interior/>
            <NumberFormat/>
            <Protection/>
          </Style>
        </Styles>
        ${rankWS}
        ${metaWS}
      </Workbook>`;
    const fname = `contest_${meta.contestId}_rank_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.xls`;
    saveBlob(xml, 'application/vnd.ms-excel', fname);
    return fname;
  }

  // ---------- UI ----------
  let uiBox, btn, toast, lastAutoKey = '';
  async function ensureUI(){
    await waitForBody();
    if (document.getElementById('acc-rank-export-root')) return;

    uiBox = document.createElement('div');
    uiBox.id = 'acc-rank-export-root';
    uiBox.className = 'acc-rank-export';

    btn = document.createElement('button');
    btn.textContent = '导出过题榜（Excel）';
    btn.title = '优先导出 .xlsx；若 CDN 不通自动导出 .xls（离线）';

    toast = document.createElement('div');
    toast.className = 'acc-rank-toast';
    toastEl = toast;

    uiBox.appendChild(btn);
    document.body.appendChild(uiBox);
    document.body.appendChild(toast);

    btn.addEventListener('click', () => exportCurrent(false));
  }

  // ---------- 主流程 ----------
  async function exportCurrent(auto=true){
    const id = getContestIdFromUrl();
    if (!id) {
      if (!auto) alert('未识别到 contestId（请确认当前是 /#/xxxx/rank 页面）');
      return;
    }
    const key = `${id}:${auto?'auto':'manual'}`;
    if (auto && lastAutoKey === key) return;
    lastAutoKey = key;

    try {
      showToast('获取榜单数据…');
      const api = getRankApiUrl(id);
      const first = await getJSONViaGM(api);
      const arr = extractRankArray(first);
      if (!Array.isArray(arr)) throw new Error('响应不是数组（未在对象/字符串中找到 rank 列表）');
      if (arr.length === 0) { showToast('数据为空'); return; }

      const { headers, rows, problems } = flattenRank(arr);
      const meta = { contestId: id, problems, source: 'api' };

      // 先试 .xlsx；失败走 .xls 离线
      if (trySaveXLSX(headers, rows, meta)) {
        showToast('导出完成（.xlsx）');
        return;
      }
      showToast('CDN 不可用，已启用离线 .xls 导出…', 1800);
      saveXLS_Offline({ headers, rows, meta });
      showToast('导出完成（.xls 离线）');
    } catch (e) {
      console.error(e);
      alert('导出失败：' + (e?.message || e));
    }
  }

  function onHashChangeMaybeExport() {
    const h = location.hash || '';
    if (/#\/\d+\/rank\b/.test(h) || /[?&](contestId|id)=\d+/.test(h)) {
      setTimeout(() => exportCurrent(true), 500);
    }
  }

  (async function bootstrap(){
    await ensureUI();
    onHashChangeMaybeExport();
    window.addEventListener('hashchange', onHashChangeMaybeExport);
  })();
})();
