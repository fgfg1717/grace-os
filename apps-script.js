// ════════════════════════════════════════════════════════════════
//  靜華 OS × Google Sheets 同步腳本
//  貼到 Google Apps Script（script.google.com）後部署為「網頁應用程式」
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1jC4q8Uz13ZnlYiinzGarvb3eqy4RZYtm3PpY-cELPRA';
const CACHE_TAB      = '手機快取';
const MAIN_GID       = 974288665;   // 主紀錄分頁的 gid

// ── doPost：接收 Grace OS 快速記錄 ─────────────────────────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const cache = getOrCreateCacheSheet(ss);
    const today = fmt(new Date());
    cache.appendRow([
      data.is_attack === true,   // A: 進攻 (true/false)
      data.category  || '',      // B: 分類
      data.time      || '',      // C: 時間 (HHMM-HHMM)
      data.summary   || '',      // D: 摘要+想法
      data.date      || today,   // E: 日期
      false                      // F: 已合併
    ]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: err.message });
  }
}

// ── mergeToday：從 Sheets 選單觸發，將今日手機快取合併至主紀錄 ──
function mergeToday() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cache = ss.getSheetByName(CACHE_TAB);

  if (!cache || cache.getLastRow() <= 1) {
    ui().alert('手機快取是空的，沒有資料可合併。'); return;
  }

  const today    = fmt(new Date());
  const cacheAll = cache.getDataRange().getValues().slice(1); // 略過標題列

  const toMerge = cacheAll
    .map((row, i) => ({ idx: i + 2, row }))          // idx 是 1-based 列號
    .filter(({ row }) => row[4] === today && !row[5]); // 今天 & 未合併

  if (!toMerge.length) {
    ui().alert('今天（' + today + '）沒有未合併的手機快取。'); return;
  }

  const main = getMainSheet(ss);

  // 主表空白 → 直接建立今日區塊
  if (main.getLastRow() === 0) {
    appendSection(main, today, toMerge);
    markMerged(cache, toMerge);
    ui().alert('已建立今日區塊並合併 ' + toMerge.length + ' 筆 ✓');
    return;
  }

  const vals = main.getRange(1, 1, main.getLastRow(), 4).getValues();

  // 找今日的區塊標題列（0-based index）
  let headerIdx = -1;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].join('').includes(today)) { headerIdx = i; break; }
  }

  // 找不到今日區塊 → 附加到最後
  if (headerIdx === -1) {
    appendSection(main, today, toMerge);
    markMerged(cache, toMerge);
    ui().alert('已建立今日區塊並合併 ' + toMerge.length + ' 筆 ✓');
    return;
  }

  // 找今日區塊結尾（下一個日期標題或表格末尾）
  let sectionEnd = vals.length; // exclusive 0-based
  for (let i = headerIdx + 1; i < vals.length; i++) {
    if (isDateHeader(vals[i])) { sectionEnd = i; break; }
  }

  // 收集今日區塊內已有時間的列（用於計算插入位置）
  const existing = [];
  for (let i = headerIdx + 1; i < sectionEnd; i++) {
    const t = startTime(vals[i][2]);
    if (t !== null) existing.push({ rowNum: i + 1, t }); // rowNum 是 1-based
  }

  // 由大到小排序後從底部插入，避免插入後列號位移錯誤
  const inserts = toMerge
    .map(({ row }) => ({
      t:    startTime(row[2]) !== null ? startTime(row[2]) : 9999,
      data: [row[0], row[1], row[2], row[3]]
    }))
    .sort((a, b) => b.t - a.t);

  inserts.forEach(({ t, data }) => {
    const before    = existing.filter(e => e.t <= t);
    const afterRow  = before.length
      ? before[before.length - 1].rowNum
      : headerIdx + 1; // 緊接在標題列後

    main.insertRowAfter(afterRow);
    main.getRange(afterRow + 1, 1, 1, 4).setValues([data]);

    // 插入後調整 existing 的列號
    existing.forEach(e => { if (e.rowNum > afterRow) e.rowNum++; });
    sectionEnd++;
  });

  markMerged(cache, toMerge);
  ui().alert('已合併 ' + toMerge.length + ' 筆到日常紀錄，並按時間排序 ✓');
}

// ── 工具函式 ────────────────────────────────────────────────────
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

// 開啟 Sheets 時自動建立選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('靜華 OS')
    .addItem('合併今日手機快取', 'mergeToday')
    .addToUi();
}
