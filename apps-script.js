// ════════════════════════════════════════════════════════════════
//  靜華 OS × Google Sheets 同步腳本
//  貼到 Google Apps Script（script.google.com）後部署為「網頁應用程式」
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1jC4q8Uz13ZnlYiinzGarvb3eqy4RZYtm3PpY-cELPRA';
const CACHE_TAB      = '手機快取';
const PLAN_TAB       = '本週計畫';
const REVIEW_TAB     = '復盤紀錄';
const VOCAB_TAB      = '英文單字庫';
const MAIN_GID       = 974288665;
const READ_TOKEN     = 'graceos2026read';

// ── doGet：供 Claude 讀取 AAR 主紀錄 ────────────────────────────
// 呼叫格式：?token=graceos2026read&from=2026/04/28&to=2026/05/04
function doGet(e) {
  const params = e.parameter || {};
  if (params.token !== READ_TOKEN) return out({ error: 'unauthorized' });

  const from = params.from || '';
  const to   = params.to   || '';

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const main = getMainSheet(ss);
  const rows = main.getDataRange().getValues();

  const result  = [];
  let   curDate = '';

  rows.forEach(row => {
    if (isDateHeader(row)) {
      const m = row.join('').match(/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/);
      curDate = m ? m[0].replace(/-/g, '/') : String(row[0]);
      return;
    }
    if (!curDate) return;
    if (from && curDate < from) return;
    if (to   && curDate > to)   return;

    const hasContent = row.some(c => c !== '' && c !== false && c !== null);
    if (hasContent) {
      result.push({
        date:     curDate,
        attack:   row[0],
        category: row[1],
        time:     row[2],
        summary:  row[3],
        extra:    row.slice(4).filter(c => c !== '' && c !== null)
      });
    }
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: result }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── doPost：接收三種類型的資料 ───────────────────────────────────
// type: 'capture' → Grace OS 快速記錄（原有功能）
// type: 'plan'    → Claude 寫入本週計畫
// type: 'review'  → Claude 寫入週/月復盤報告
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type || 'capture';
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 英文單字庫 ──
    if (type === 'vocab') {
      const sheet = getOrCreateVocabSheet(ss);
      sheet.appendRow([
        data.date  || fmt(new Date()),  // A: 日期
        data.time  || '',               // B: 時間
        data.word  || '',               // C: 單字/句子
        data.note  || '',               // D: 中文解釋/例句
        data.tag   || '',               // E: 標籤
      ]);
      return out({ ok: true });
    }

    // ── Grace OS 快速記錄（原有邏輯不變）──
    if (type === 'capture') {
      const cache = getOrCreateCacheSheet(ss);
      const today = fmt(new Date());
      cache.appendRow([
        data.is_attack === true,
        data.category  || '',
        data.time      || '',
        data.summary   || '',
        data.date      || today,
        false
      ]);
      return out({ ok: true });
    }

    // ── 以下需要 token 驗證 ──
    if (data.token !== READ_TOKEN) return out({ error: 'unauthorized' });

    // ── 寫入本週計畫（有顏色的格式化表格）──
    if (type === 'plan') {
      writePlan(ss, data);
      return out({ ok: true });
    }

    // ── 寫入復盤紀錄 ──
    if (type === 'review') {
      const sheet = getOrCreateReviewSheet(ss);
      sheet.appendRow([
        fmt(new Date()),          // A: 寫入日期
        data.review_type || '週', // B: 類型（週/月）
        data.period      || '',   // C: 週期（例：2026/05/04-05/08）
        data.highlight   || '',   // D: 亮點
        data.pain        || '',   // E: 痛點
        data.action      || '',   // F: 防呆行動
        data.co_pct      || '',   // G: 公司目標%
        data.personal_pct|| '',   // H: 個人目標%
        data.life_pct    || '',   // I: 個人生活%
        data.misc_pct    || '',   // J: 瑣務%
        data.full_report || ''    // K: 完整報告
      ]);
      return out({ ok: true });
    }

    return out({ ok: false, error: 'unknown type' });
  } catch (err) {
    return out({ ok: false, error: err.message });
  }
}

// ── mergeToday：將今日手機快取合併至主紀錄 ──────────────────────
function mergeToday() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cache = ss.getSheetByName(CACHE_TAB);

  if (!cache || cache.getLastRow() <= 1) {
    safeAlert('手機快取是空的，沒有資料可合併。'); return;
  }

  const today    = fmt(new Date());
  const cacheAll = cache.getDataRange().getValues().slice(1);

  const toDateStr = v => {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd');
    return String(v).replace(
      /(\d{4})\/(\d{1,2})\/(\d{1,2})/,
      (_, y, m, d) => `${y}/${m.padStart(2,'0')}/${d.padStart(2,'0')}`
    );
  };
  const toMerge = cacheAll
    .map((row, i) => ({ idx: i + 2, row }))
    .filter(({ row }) => toDateStr(row[4]) === today && !row[5]);

  if (!toMerge.length) {
    safeAlert('今天（' + today + '）沒有未合併的手機快取。'); return;
  }

  const main = getMainSheet(ss);

  if (main.getLastRow() === 0) {
    appendSection(main, today, toMerge);
    markMerged(cache, toMerge);
    safeAlert('已建立今日區塊並合併 ' + toMerge.length + ' 筆 ✓');
    return;
  }

  const vals = main.getRange(1, 1, main.getLastRow(), 4).getValues();

  let headerIdx = -1;
  for (let i = 0; i < vals.length; i++) {
    if (normDate(vals[i].join('')).includes(normDate(today))) { headerIdx = i; break; }
  }

  if (headerIdx === -1) {
    appendSection(main, today, toMerge);
    markMerged(cache, toMerge);
    safeAlert('已建立今日區塊並合併 ' + toMerge.length + ' 筆 ✓');
    return;
  }

  let sectionEnd = vals.length;
  for (let i = headerIdx + 1; i < vals.length; i++) {
    if (isDateHeader(vals[i])) { sectionEnd = i; break; }
  }

  const existing = [];
  for (let i = headerIdx + 1; i < sectionEnd; i++) {
    const t = startTime(vals[i][2]);
    if (t !== null) existing.push({ rowNum: i + 1, t });
  }

  const inserts = toMerge
    .map(({ row }) => ({
      t:    startTime(row[2]) !== null ? startTime(row[2]) : 9999,
      data: [row[0], row[1], row[2], row[3]]
    }))
    .sort((a, b) => b.t - a.t);

  inserts.forEach(({ t, data }) => {
    const before   = existing.filter(e => e.t <= t);
    const afterRow = before.length
      ? before[before.length - 1].rowNum
      : headerIdx + 1;

    main.insertRowAfter(afterRow);
    main.getRange(afterRow + 1, 1, 1, 4).setValues([data]);
    existing.forEach(e => { if (e.rowNum > afterRow) e.rowNum++; });
    sectionEnd++;
  });

  markMerged(cache, toMerge);
  safeAlert('已合併 ' + toMerge.length + ' 筆到日常紀錄，並按時間排序 ✓');
}

// ── 工具函式 ────────────────────────────────────────────────────
function getOrCreateVocabSheet(ss) {
  let s = ss.getSheetByName(VOCAB_TAB);
  if (!s) {
    s = ss.insertSheet(VOCAB_TAB);
    s.appendRow(['日期', '時間', '單字/句子', '中文解釋/例句', '標籤']);
    s.setFrozenRows(1);
    s.setColumnWidth(1, 90);
    s.setColumnWidth(2, 60);
    s.setColumnWidth(3, 220);
    s.setColumnWidth(4, 300);
    s.setColumnWidth(5, 80);
  }
  return s;
}

function getOrCreateCacheSheet(ss) {
  let s = ss.getSheetByName(CACHE_TAB);
  if (!s) {
    s = ss.insertSheet(CACHE_TAB);
    s.appendRow(['進攻', '分類', '時間', '摘要+想法', '日期', '已合併']);
    s.setFrozenRows(1);
    s.setColumnWidth(1, 50);
    s.setColumnWidth(5, 90);
    s.setColumnWidth(6, 60);
  }
  return s;
}

// ── 寫入格式化週計畫 ─────────────────────────────────────────────
function writePlan(ss, data) {
  const CAT_COLORS = {
    '固定會議': '#CFE2F3', '固定行程': '#CFE2F3',
    '學習':     '#D9EAD3',
    '深度工作': '#FCE5CD',
    '無謂雜事': '#FFF2CC',
    '結尾收尾': '#EAD1DC',
    '個人行程': '#D9D2E9',
    '休息':     '#F3F3F3',
  };

  let s = ss.getSheetByName(PLAN_TAB);
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet(PLAN_TAB);

  const content = (data.content || '').replace(/\r\n/g, '\n');
  const lines   = content.split('\n');
  let   row     = 1;

  // 標題列
  const title = '本週行程表　' + (data.week || '').replace('-', ' – ');
  s.getRange(row, 1, 1, 4).merge().setValue(title);
  s.getRange(row, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  s.getRange(row, 1, 1, 4).setBackground('#C9DAF8');
  row++;

  lines.forEach(line => {
    const t = line.trim();

    // 日期標題：## 5/4（一）　密集日
    if (t.startsWith('## ')) {
      if (row > 2) row++;
      s.getRange(row, 1, 1, 4).merge().setValue(t.replace('## ', ''));
      s.getRange(row, 1).setFontWeight('bold').setFontSize(11).setFontColor('#FFFFFF');
      s.getRange(row, 1, 1, 4).setBackground('#4472C4');
      row++;
      s.getRange(row, 1, 1, 4).setValues([['時段', '內容', '類型', '工作積分（分鐘）']]);
      s.getRange(row, 1, 1, 4).setBackground('#A4C2F4').setFontWeight('bold');
      row++;
      return;
    }

    // 表格分隔線
    if (/^\|[-\s|]+\|$/.test(t)) return;

    // 表格資料列：| time | content | category |
    if (t.startsWith('|') && t.endsWith('|')) {
      const cols = t.split('|').map(c => c.trim()).filter(c => c !== '');
      if (cols.length >= 3) {
        const cat = cols[2] || '';
        const bg  = CAT_COLORS[cat] || '#FFFFFF';
        const min = calcMinutes(cols[0]);
        s.getRange(row, 1, 1, 4).setValues([[cols[0], cols[1], cat, min]]);
        s.getRange(row, 1, 1, 4).setBackground(bg);
        row++;
      }
      return;
    }

    // 本日小計
    if (t.startsWith('本日淨工作時數')) {
      s.getRange(row, 1, 1, 4).merge().setValue(t.replace(/\*\*/g, ''));
      s.getRange(row, 1).setFontWeight('bold').setFontStyle('italic');
      s.getRange(row, 1, 1, 4).setBackground('#F8F9FA');
      row++;
    }
  });

  // 欄寬與框線
  s.setColumnWidth(1, 140);
  s.setColumnWidth(2, 320);
  s.setColumnWidth(3, 100);
  s.setColumnWidth(4, 130);
  if (row > 2) {
    s.getRange(1, 1, row - 1, 4)
      .setBorder(true, true, true, true, true, true,
                 '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }
  s.setFrozenRows(1);
}

function calcMinutes(timeStr) {
  const m = timeStr.match(/(\d{1,2}):(\d{2})\s*[–\-]\s*(\d{1,2}):(\d{2})/);
  if (!m) return '';
  const start = parseInt(m[1]) * 60 + parseInt(m[2]);
  const end   = parseInt(m[3]) * 60 + parseInt(m[4]);
  return end > start ? end - start : '';
}

function getOrCreateReviewSheet(ss) {
  let s = ss.getSheetByName(REVIEW_TAB);
  if (!s) {
    s = ss.insertSheet(REVIEW_TAB);
    s.appendRow(['寫入日期', '類型', '週期', '亮點', '痛點', '防呆行動',
                 '公司目標%', '個人目標%', '個人生活%', '瑣務%', '完整報告']);
    s.setFrozenRows(1);
    s.setColumnWidths(1, 11, 120);
    s.setColumnWidth(11, 400);
  }
  return s;
}

function getMainSheet(ss) {
  return ss.getSheets().find(s => s.getSheetId() === MAIN_GID) || ss.getSheets()[0];
}

function fmt(d) {
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd');
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ui() { return SpreadsheetApp.getUi(); }

function safeAlert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) { Logger.log('[GraceOS] ' + msg); }
}

// 統一日期格式比較用（去掉補零），避免 2026/05/04 vs 2026/5/4 對不上
function normDate(str) {
  return String(str).replace(/(\d{4})[\/\-]0*(\d+)[\/\-]0*(\d+)/g, (_, y, mo, dy) => `${y}/${parseInt(mo)}/${parseInt(dy)}`);
}

function isDateHeader(row) {
  return /\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/.test(row.join(''));
}

function startTime(val) {
  const s = String(val || '').replace(/\D/g, '');
  if (s.length < 3) return null;
  return parseInt(s.substring(0, 4).padEnd(4, '0'), 10);
}

function appendSection(sheet, today, captures) {
  sheet.appendRow([today, '', '', '']);
  captures
    .slice()
    .sort((a, b) => (startTime(a.row[2]) || 0) - (startTime(b.row[2]) || 0))
    .forEach(({ row }) => sheet.appendRow([row[0], row[1], row[2], row[3]]));
}

function markMerged(cache, items) {
  items.forEach(({ idx }) => cache.getRange(idx, 6).setValue(true));
}

function diagnoseCache() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cache = ss.getSheetByName(CACHE_TAB);
  if (!cache) { ui().alert('找不到手機快取分頁'); return; }
  const today   = fmt(new Date());
  const allRows = cache.getDataRange().getValues().slice(1);
  if (!allRows.length) { ui().alert('快取是空的'); return; }
  const lines = allRows.map((row, i) => {
    const raw    = row[4];
    const parsed = (raw instanceof Date)
      ? Utilities.formatDate(raw, 'Asia/Taipei', 'yyyy/MM/dd')
      : String(raw);
    return `第${i+2}列｜原始:${raw}｜解析:${parsed}｜已合併:${row[5]}`;
  });
  ui().alert(`今天是：${today}\n\n${lines.join('\n')}`);
}

function mergeAll() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cache = ss.getSheetByName(CACHE_TAB);
  if (!cache || cache.getLastRow() <= 1) { safeAlert('手機快取是空的，沒有資料可合併。'); return; }
  const cacheAll = cache.getDataRange().getValues().slice(1);
  const toDateStr = v => {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd');
    return String(v).replace(/(\d{4})\/(\d{1,2})\/(\d{1,2})/, (_, y, m, d) => `${y}/${m.padStart(2,'0')}/${d.padStart(2,'0')}`);
  };
  const unmerged = cacheAll
    .map((row, i) => ({ idx: i + 2, row, date: toDateStr(row[4]) }))
    .filter(({ row }) => !row[5]);
  if (!unmerged.length) { safeAlert('沒有待合併的資料，全部都已合併過了。'); return; }
  const byDate = {};
  unmerged.forEach(item => { if (!byDate[item.date]) byDate[item.date] = []; byDate[item.date].push(item); });
  const main = getMainSheet(ss);
  let total = 0;
  Object.keys(byDate).sort().forEach(date => {
    const items = byDate[date];
    const lastRow = main.getLastRow();
    const vals = lastRow > 0 ? main.getRange(1, 1, lastRow, 4).getValues() : [];
    let headerIdx = -1;
    for (let i = 0; i < vals.length; i++) { if (normDate(vals[i].join('')).includes(normDate(date))) { headerIdx = i; break; } }
    if (headerIdx === -1) {
      appendSection(main, date, items);
    } else {
      let sectionEnd = vals.length;
      for (let i = headerIdx + 1; i < vals.length; i++) { if (isDateHeader(vals[i])) { sectionEnd = i; break; } }
      const existing = [];
      for (let i = headerIdx + 1; i < sectionEnd; i++) { const t = startTime(vals[i][2]); if (t !== null) existing.push({ rowNum: i + 1, t }); }
      const inserts = items.map(({ row }) => ({ t: startTime(row[2]) !== null ? startTime(row[2]) : 9999, data: [row[0], row[1], row[2], row[3]] })).sort((a, b) => b.t - a.t);
      inserts.forEach(({ t, data }) => {
        const before = existing.filter(e => e.t <= t);
        const afterRow = before.length ? before[before.length - 1].rowNum : headerIdx + 1;
        main.insertRowAfter(afterRow);
        main.getRange(afterRow + 1, 1, 1, 4).setValues([data]);
        existing.forEach(e => { if (e.rowNum > afterRow) e.rowNum++; });
        sectionEnd++;
      });
    }
    markMerged(cache, items);
    total += items.length;
  });
  safeAlert('已合併 ' + total + ' 筆紀錄到日常紀錄 ✓');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('靜華 OS')
    .addItem('合併今日手機快取', 'mergeToday')
    .addItem('合併所有未合併快取', 'mergeAll')
    .addItem('診斷手機快取', 'diagnoseCache')
    .addItem('設定每日 18:45 自動合併', 'setupDailyMerge')
    .addToUi();
}

function setupDailyMerge() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'mergeToday' || fn === 'mergeAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mergeAll')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .inTimezone('Asia/Taipei')
    .create();
  safeAlert('設定完成！每天 18:00–19:00 之間會自動合併所有未合併快取 ✓');
}
