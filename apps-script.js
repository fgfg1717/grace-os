// ════════════════════════════════════════════════════════════════
//  靜華 OS × Google Sheets 同步腳本
//  貼到 Google Apps Script（script.google.com）後部署為「網頁應用程式」
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1jC4q8Uz13ZnlYiinzGarvb3eqy4RZYtm3PpY-cELPRA';
const CACHE_TAB      = '手機快取';
const PLAN_TAB       = '本週計畫';
const REVIEW_TAB     = '復盤紀錄';
const VOCAB_TAB      = '英文單字庫';
const LEDGER_TAB     = '記帳明細';
const STOCK_TAB      = '股票紀錄';
const MAIN_GID       = 974288665;
const READ_TOKEN     = 'graceos2026read';

// ── doGet：供 Grace OS 讀取資料 ─────────────────────────────────
// AAR：?token=graceos2026read&from=2026/04/28&to=2026/05/04
// 本週計畫：?token=graceos2026read&type=plan
function doGet(e) {
  const params = e.parameter || {};
  if (params.token !== READ_TOKEN) return out({ error: 'unauthorized' });

  // ── 讀取本週計畫 ──
  if (params.type === 'plan') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_TAB);
    if (!planSheet || planSheet.getLastRow() === 0)
      return ContentService.createTextOutput(JSON.stringify({ ok: true, data: [], message: '尚未建立本週計畫' }))
        .setMimeType(ContentService.MimeType.JSON);
    const data = planSheet.getDataRange().getValues().map(r => r.map(c => String(c)));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

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

    // ── 記帳明細 ──
    if (type === 'ledger') {
      const sheet = getOrCreateLedgerSheet(ss);
      sheet.appendRow([
        data.date        || fmt(new Date()),  // A: 日期
        data.ledger_type || '支出',           // B: 支出/收入
        data.main_cat    || '',               // C: 主分類
        data.sub_cat     || '',               // D: 子分類
        data.amount      || 0,               // E: 金額
        data.notes       || '',               // F: 備註
      ]);
      return out({ ok: true });
    }

    // ── 股票紀錄 ──
    if (type === 'stock') {
      const sheet = getOrCreateStockSheet(ss);
      sheet.appendRow([
        data.date    || fmt(new Date()),  // A: 日期
        data.account || '',               // B: 個人/家庭
        data.market  || '',               // C: 台股/美股
        data.action  || '',               // D: 買入/賣出
        data.code    || '',               // E: 股票代號
        data.name    || '',               // F: 股票名稱
        data.qty     || 0,               // G: 數量
        data.price   || 0,               // H: 成交價
        data.total   || 0,               // I: 總金額
        data.notes   || '',               // J: 備註
      ]);
      return out({ ok: true });
    }

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

// ── setupTodayTemplate：在最上方插入今天的範本 ──────────────────
// 新的在最上面，複製最頂端那天的格式，清除內容後填入今天日期
function buildFreshTemplate(main, today) {
  const numCols = 7;
  main.insertRowsBefore(1, 22);
  const r = (row, col) => main.getRange(row, col);
  const rng = (row, col, rows, cols) => main.getRange(row, col, rows, cols);

  // 日期 header
  r(1,1).setValue(today).setFontSize(14).setFontWeight('bold');
  r(1,4).setValue('每日單一亮點');
  r(1,6).setValue('1% 推進問題').setFontColor('#cc0000').setFontWeight('bold');
  r(1,7).setValue('我今天到底哪裡變強了');
  rng(2,1,1,numCols).merge().setValue('上班坐到座位上的前 30 分鐘，嚴禁打開 Email 與通訊軟體。');

  // 反思欄
  r(4,2).setValue('選一件讓我有感覺、有啟發的事情');
  r(5,2).setValue('');
  r(7,2).setValue('有出門的零碎時間多，但都忘記五分鐘小抽屜');
  r(9,2).setValue('我從過程中學習或觀察到什麼事情？');
  r(10,2).setValue('');

  // AAR header
  r(12,2).setValue('AAR');
  r(13,1).setValue('進攻');
  r(13,2).setValue('分類');
  r(13,3).setValue('今天完成了什麼事情？');
  r(13,6).setValue('1% 推進問題');
  r(13,7).setValue('我今天到底哪裡變強了');

  // 8 列空白資料列（含 checkbox）
  for (let i = 14; i <= 21; i++) {
    main.getRange(i, 1).insertCheckboxes();
  }

  // 空白分隔
  r(22,1).setValue('');
  safeAlert('✓ 已建立 ' + today + ' 的全新範本（預設格式）！');
}

function setupTodayTemplate() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const main = getMainSheet(ss);
  const today = fmt(new Date());
  const lastRow = main.getLastRow();

  if (lastRow === 0) {
    buildFreshTemplate(main, today); return;
  }

  const numCols = Math.max(main.getLastColumn(), 8);
  const allVals = main.getRange(1, 1, lastRow, numCols).getValues();

  // 今天已存在？
  for (let i = 0; i < allVals.length; i++) {
    if (normDate(allVals[i].join('')).includes(normDate(today))) {
      safeAlert('今天（' + today + '）的範本已存在！'); return;
    }
  }

  // 找最上方那個 section 的列數（第 1 列到下一個日期 header 之前）
  let firstSectionEnd = allVals.length;
  for (let i = 1; i < allVals.length; i++) {
    if (isDateHeader(allVals[i])) { firstSectionEnd = i; break; }
  }
  const srcRows = firstSectionEnd; // 第一個 section 共幾列

  // ★ 安全檢查：section 大小不合理（太小或太大）代表上方區塊已損壞，
  //   直接用全新空白範本，避免把錯誤的格式往後複製
  if (srcRows < 10 || srcRows > 50) {
    buildFreshTemplate(main, today); return;
  }

  // 在最上方插入空列（srcRows 行內容 + 2 行分隔）
  main.insertRowsBefore(1, srcRows + 2);

  // 舊的第一個 section 現在往下移了 srcRows+2 列
  const oldSrcStart = srcRows + 3; // 1-indexed
  main.getRange(oldSrcStart, 1, srcRows, numCols)
      .copyTo(main.getRange(1, 1, srcRows, numCols));

  // 填入今天日期
  main.getRange(1, 1).setValue(today);

  // 清除內容保留格式：找「分類」欄位標題列
  const newVals = main.getRange(1, 1, srcRows, numCols).getValues();
  let colHeaderOffset = -1;
  for (let i = 0; i < newVals.length; i++) {
    if (newVals[i].join('').includes('分類')) { colHeaderOffset = i; break; }
  }

  if (colHeaderOffset >= 0) {
    // 清除每日重點貼區（日期列之後、欄位標題之前）
    if (colHeaderOffset > 1) {
      main.getRange(2, 1, colHeaderOffset - 1, numCols).clearContent();
    }
    // 清除 AAR 資料列（保留 checkbox 和下拉選單格式）
    const aarDataRows = srcRows - colHeaderOffset - 1;
    if (aarDataRows > 0) {
      main.getRange(colHeaderOffset + 2, 1, aarDataRows, 4).clearContent();
    }
  }

  safeAlert('✓ 已建立 ' + today + ' 的範本，已加在最上方！');
}

// ── setupMorningTrigger：設定每天早上 7 點自動建立範本 ───────────
function setupMorningTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'setupTodayTemplate') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('setupTodayTemplate')
    .timeBased().everyDays(1).atHour(7)
    .inTimezone('Asia/Taipei').create();
  Logger.log('設定完成！每天早上 7:00-8:00 會自動建立當天範本 ✓');
}

// ── mergeToday：將今日手機快取合併至 AAR 區塊（正確位置）────────
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

  const numCols = Math.max(main.getLastColumn(), 4);
  const vals    = main.getRange(1, 1, main.getLastRow(), numCols).getValues();

  // 找今日日期 header
  let headerIdx = -1;
  for (let i = 0; i < vals.length; i++) {
    if (normDate(vals[i].join('')).includes(normDate(today))) { headerIdx = i; break; }
  }

  if (headerIdx === -1) {
    // 今天的範本不存在，先建立再合併
    safeAlert('找不到今天的範本，先執行「建立今日範本」再合併。\n（或先手動複製一天的格式）');
    return;
  }

  // 找今天這個 section 的結束位置
  let sectionEnd = vals.length;
  for (let i = headerIdx + 1; i < vals.length; i++) {
    if (isDateHeader(vals[i])) { sectionEnd = i; break; }
  }

  // ★ 關鍵修正：找到欄位標題列（含「分類」），資料插在它之後
  let aarDataStart = headerIdx + 1; // 預設值：date 行之後
  for (let i = headerIdx + 1; i < sectionEnd; i++) {
    if (vals[i].join('').includes('分類')) {
      aarDataStart = i + 1; // 欄位標題的下一行才是資料區
      break;
    }
  }

  // 收集 AAR 資料區已有的有時間的列
  const existing = [];
  for (let i = aarDataStart; i < sectionEnd; i++) {
    const t = startTime(vals[i][2]); // 時間在第 C 欄（index 2）
    if (t !== null) existing.push({ rowNum: i + 1, t });
  }

  // 按時間排序插入（由大到小，確保插入順序正確）
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
      : aarDataStart; // 沒有比它早的，就插在 AAR 資料區第一列

    main.insertRowAfter(afterRow);
    main.getRange(afterRow + 1, 1, 1, 4).setValues([data]);
    existing.forEach(e => { if (e.rowNum > afterRow) e.rowNum++; });
    sectionEnd++;
  });

  markMerged(cache, toMerge);
  safeAlert('已合併 ' + toMerge.length + ' 筆到今日 AAR 區塊，並按時間排序 ✓');
}

// ── 工具函式 ────────────────────────────────────────────────────
function getOrCreateLedgerSheet(ss) {
  let s = ss.getSheetByName(LEDGER_TAB);
  if (!s) {
    s = ss.insertSheet(LEDGER_TAB);
    s.appendRow(['日期', '支出/收入', '主分類', '子分類', '金額', '備註']);
    s.setFrozenRows(1);
    s.setColumnWidths(1, 6, 100);
    s.setColumnWidth(6, 200);
  }
  return s;
}

function getOrCreateStockSheet(ss) {
  let s = ss.getSheetByName(STOCK_TAB);
  if (!s) {
    s = ss.insertSheet(STOCK_TAB);
    s.appendRow(['日期', '帳戶', '市場', '動作', '代號', '名稱', '數量', '成交價', '總金額', '備註']);
    s.setFrozenRows(1);
    s.setColumnWidths(1, 10, 90);
    s.setColumnWidth(9, 110);
    s.setColumnWidth(10, 180);
  }
  return s;
}

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
  // 先用 GID，找不到再用名稱，再找不到用第一個有資料的分頁
  const byId = ss.getSheets().find(s => s.getSheetId() === MAIN_GID);
  if (byId) return byId;
  const byName = ss.getSheets().find(s => /AAR|主紀錄/.test(s.getName()));
  if (byName) return byName;
  const withData = ss.getSheets().find(s => s.getLastRow() > 0);
  return withData || ss.getSheets()[0];
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
    ss.toast(msg, '靜華 OS', 5);
  } catch(e) { Logger.log('[GraceOS] ' + msg); }
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

  const toDateStr = v => {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy/MM/dd');
    return String(v).replace(/(\d{4})\/(\d{1,2})\/(\d{1,2})/, (_, y, m, d) => `${y}/${m.padStart(2,'0')}/${d.padStart(2,'0')}`);
  };

  const cacheAll = cache.getDataRange().getValues().slice(1);
  const unmerged = cacheAll
    .map((row, i) => ({ idx: i + 2, row, date: toDateStr(row[4]) }))
    .filter(({ row }) => !row[5]);
  if (!unmerged.length) { safeAlert('沒有待合併的資料，全部都已合併過了。'); return; }

  const byDate = {};
  unmerged.forEach(item => { if (!byDate[item.date]) byDate[item.date] = []; byDate[item.date].push(item); });

  const main    = getMainSheet(ss);
  const lastRow = main.getLastRow();

  // 表格是空的：全部 appendSection
  if (lastRow === 0) {
    let t = 0;
    Object.keys(byDate).sort().forEach(d => { appendSection(main, d, byDate[d]); markMerged(cache, byDate[d]); t += byDate[d].length; });
    safeAlert('已合併 ' + t + ' 筆紀錄 ✓');
    return;
  }

  // ★ 讀一次表格，所有日期共用（原本每個日期各讀一次，速度慢 10 倍）
  let vals = main.getRange(1, 1, lastRow, 4).getValues();

  // 建立日期 header 索引 map（0-based rowIdx）
  const headerMap = {};
  for (let i = 0; i < vals.length; i++) {
    if (isDateHeader(vals[i])) {
      const m = vals[i].join('').match(/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/);
      if (m) { const k = normDate(m[0]); if (headerMap[k] === undefined) headerMap[k] = i; }
    }
  }

  let total = 0;

  // 由舊到新處理（= sheet 由下往上）：插入不影響上方尚未處理的列號
  Object.keys(byDate).sort().forEach(date => {
    const items   = byDate[date];
    const hRowIdx = headerMap[normDate(date)];

    if (hRowIdx === undefined) {
      // 沒有範本，直接 appendSection
      appendSection(main, date, items);
      markMerged(cache, items);
      total += items.length;
      return;
    }

    // 找 section 結束位置（0-based）
    let secEnd = vals.length;
    for (let i = hRowIdx + 1; i < vals.length; i++) {
      if (isDateHeader(vals[i])) { secEnd = i; break; }
    }

    // 找「分類」標題列，確定 AAR 資料起點（0-based）
    let aarStart = hRowIdx + 1;
    for (let i = hRowIdx + 1; i < secEnd; i++) {
      if (vals[i].join('').includes('分類')) { aarStart = i + 1; break; }
    }

    // 找 AAR 資料區最後一筆有內容的列（0-based）
    let lastFilled = aarStart - 1;
    for (let i = aarStart; i < secEnd; i++) {
      if (vals[i].some(c => c !== '' && c !== false && c !== null)) lastFilled = i;
    }

    // 排序後，批次插入到最後有內容列的下方（2 次 API call 搞定一整天）
    const sorted = items.slice().sort((a, b) => (startTime(a.row[2]) || 0) - (startTime(b.row[2]) || 0));
    const count  = sorted.length;
    const insertAfter1 = lastFilled + 1; // 1-based sheet row
    const firstNew1    = lastFilled + 2; // 1-based sheet row of first new row

    main.insertRowsAfter(insertAfter1, count);
    main.getRange(firstNew1, 1, count, 4)
        .setValues(sorted.map(({ row }) => [row[0], row[1], row[2], row[3]]));

    // 同步更新 in-memory vals，讓後續日期的 secEnd / aarStart 計算正確
    const empty = Array(count).fill(null).map(() => ['', '', '', '']);
    vals.splice(lastFilled + 1, 0, ...empty);

    markMerged(cache, items);
    total += items.length;
  });

  safeAlert('已合併 ' + total + ' 筆紀錄到日常紀錄 ✓');
}

// ── fixMisplacedData：修正被插錯位置的快取資料 ─────────────────
// 之前的 bug 會把資料插進「反思區」（日期標題和分類標題之間）
// 這個函式找出這些錯位資料，移回正確的 AAR 資料區
function fixMisplacedData() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const main    = getMainSheet(ss);
  const lastRow = main.getLastRow();
  if (lastRow === 0) { safeAlert('主紀錄是空的'); return; }

  const numCols = Math.max(main.getLastColumn(), 4);
  const vals    = main.getRange(1, 1, lastRow, numCols).getValues();
  let totalMoved = 0;

  // 找出所有日期 section 的起始位置（0-based）
  const secStarts = [];
  for (let i = 0; i < vals.length; i++) {
    if (isDateHeader(vals[i])) secStarts.push(i);
  }

  // 從最底部往上處理，確保刪除/插入不影響上方的 section 位置
  for (let s = secStarts.length - 1; s >= 0; s--) {
    const secStart = secStarts[s];
    const secEnd   = s + 1 < secStarts.length ? secStarts[s + 1] : vals.length;

    // 找「分類」欄位標題列
    let catIdx = -1;
    for (let i = secStart + 1; i < secEnd; i++) {
      if (vals[i].join('').includes('分類')) { catIdx = i; break; }
    }
    if (catIdx === -1) continue; // 沒有 AAR 結構（appendSection 建立的段落），跳過

    // 找錯位資料：在日期標題和「分類」標題之間，col A 是 boolean（快取資料特徵）
    const misplaced  = [];
    const deleteRows = [];
    for (let i = secStart + 1; i < catIdx; i++) {
      const r = vals[i];
      if (r[0] === true || r[0] === false) {
        misplaced.push({ t: startTime(r[2]) || 9999, row: [r[0], r[1], r[2], r[3]] });
        deleteRows.push(i + 1); // 1-based
      }
    }
    if (!misplaced.length) continue;

    // 刪除錯位列（由下往上，避免列號位移）
    deleteRows.sort((a, b) => b - a);
    deleteRows.forEach(r => main.deleteRow(r));

    // 計算刪除後的「分類」標題列位置
    const newCatRow1 = catIdx + 1 - deleteRows.length; // 1-based
    const aarStart1  = newCatRow1 + 1;                  // 1-based，分類標題的下一列

    // 按時間排序，批次插入到 AAR 資料區
    misplaced.sort((a, b) => a.t - b.t);
    main.insertRowsAfter(newCatRow1, misplaced.length);
    main.getRange(aarStart1, 1, misplaced.length, 4).setValues(misplaced.map(m => m.row));

    totalMoved += misplaced.length;
    // 刪除 N 列 + 插入 N 列 → 上方 section 位置不變，不需要重新讀取
  }

  safeAlert('已修正 ' + totalMoved + ' 筆資料，移至正確的 AAR 欄位 ✓');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('靜華 OS')
    .addItem('① 建立今日範本（手動）', 'setupTodayTemplate')
    .addItem('② 合併今日手機快取', 'mergeToday')
    .addItem('合併所有未合併快取', 'mergeAll')
    .addItem('🔧 修正錯位資料（移回 AAR 區）', 'fixMisplacedData')
    .addSeparator()
    .addItem('⏰ 設定每早 7 點自動建立範本', 'setupMorningTrigger')
    .addItem('⏰ 設定每日 18:45 自動合併', 'setupDailyMerge')
    .addSeparator()
    .addItem('診斷手機快取', 'diagnoseCache')
    .addItem('初始化記帳明細分頁', 'initLedgerSheet')
    .addItem('初始化股票紀錄分頁', 'initStockSheet')
    .addToUi();
}

function initLedgerSheet() { getOrCreateLedgerSheet(SpreadsheetApp.openById(SPREADSHEET_ID)); safeAlert('記帳明細分頁已就緒 ✓'); }
function initStockSheet()  { getOrCreateStockSheet(SpreadsheetApp.openById(SPREADSHEET_ID));  safeAlert('股票紀錄分頁已就緒 ✓');  }

function setupDailyMerge() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'mergeToday' || fn === 'mergeAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mergeAll')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(45)
    .inTimezone('Asia/Taipei')
    .create();
  Logger.log('設定完成！每天 18:45 會自動合併所有未合併快取 ✓');
}
