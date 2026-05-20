# Grace-OS 專案快速索引

這是吳靜華的個人全能助理 PWA，部署於 GitHub Pages。

## 部署資訊
- **網址**：`https://fgfg1717.github.io/grace-os/`
- **本機路徑**：`C:\Users\ASUS\Downloads\Grace-agent\grace-os\`
- **更新方式**：修改 `index.html` 或 `apps-script.js` → `git add . && git commit -m "..." && git push`
- **GitHub Pages** 約 2-3 分鐘生效

## 主要檔案
- `index.html`：整個 App（CSS + HTML + JS 全部在這一個檔案）
- `apps-script.js`：Google Apps Script 程式碼（改完要重新部署才生效）
- `manifest.json`：PWA 設定

## Google Sheets & Apps Script
- **Spreadsheet ID**：`1jC4q8Uz13ZnlYiinzGarvb3eqy4RZYtm3PpY-cELPRA`
- **GAS 網址**：`https://script.google.com/macros/s/AKfycbwcFf2JuSp-z8IMiR0FQ3SCAuUHvtYnj7uKdZFUvd1mmQHo7b5W3zn0onvM-Js-giPXdQ/exec`
- **Read Token**：`graceos2026read`
- **分頁對應**：
  - 主紀錄（AAR）：快速記錄合併、每日 7 點自動建立範本、18:45 自動合併
  - 手機快取：Grace OS App 暫存
  - 本週計畫：`/週計畫` slash command 寫入
  - 記帳明細：記帳模組寫入
  - 股票紀錄：股票模組寫入

## 使用者日常流程
- **白天**：手機/網頁開 App → ⚡ 快速記錄（靈感/健康/工作/生活）
- **晚上坐電腦**：直接開 Google Sheets 填今日 AAR 反思
- **花錢**：App → 💰 記帳（不在快速記錄，兩邊資料不互通）
- **買賣股票**：App → 📈 股票追蹤
- **週覆盤**：App → 📋 讀取 AAR → 讀取資料 → 複製週覆盤格式 → 貼給 Claude
- **英文**：App → 📚 英文學習中心 → 今日 tab 刷複習卡

## App 功能模組（首頁 → 點卡片進入）
| 卡片 | 功能 | localStorage key |
|------|------|-----------------|
| ⚡ 快速記錄 | 6 分類（靈感/記帳/股票/健康/工作/生活）+ 語音 | `captures` |
| 💎 財務總覽 | 本月收支、銀行帳戶管理、股票+銀行資產快照 | `bank_accounts_v1` |
| 💰 記帳 | 支出（家用/個人/投資/定存）+ 收入（薪資/獎金等） | `ledgers` |
| 📈 股票追蹤 | 個人+家庭，台股+美股，定期定額，損益計算 | `stocks`, `stock_prices` |
| 📚 英文學習中心 | SRS 複習、TOEIC 每日新詞、查字典、課程匯入、練習 | `eng_items_v1`, `eng_lessons_v1` |
| 🎯 幕僚沙盤 | iframe 嵌入，連到獨立 App | — |
| 🗓️ 本週計畫 | 從 Sheets 讀取，今日任務高亮 | — |
| 📋 讀取 AAR | 從 Sheets 讀，一鍵複製給 Claude 週覆盤/月回顧 | — |
| 📊 週覆盤 | 自動帶入本週靈感紀錄 + 複製給 Claude 分析 | — |
| 📅 月回顧 | 輸入 → 複製給 Claude 分析 | — |

## 英文學習中心架構（SRS）
- **統一學習池**：所有來源（TOEIC 內建、查字典、課程匯入、手動）全進 `eng_items_v1`
- **SRS 間隔複習**：不會=1天、模糊=3天、會了=7天起、很熟=14天起
- **TOEIC 內建詞彙**：150+ 高頻詞，存在 `TOEIC_WORDS` 常數，每日推 5 個新詞
- **TOEIC 進度條**：掌握（reps≥3 且 lastRating≥3）/ 250 個目標（550→750）
- **課程匯入格式**：JSON，含 `courseTitle`, `classDate`, `vocabulary[]`, `idioms[]`, `sentences[]`, `grammar[]`
- **今日 tab 邏輯**：有待複習 → 先刷完複習卡；無待複習 → 顯示今日 5 個 TOEIC 新詞

## 財務模組架構
- **財務總覽**：從 `ledgers` 計算本月收支，從 `bank_accounts_v1` 抓銀行餘額，從 `stocks`+`stock_prices` 計算股票市值
- **記帳收入分類**：`INCOME_CATS = ['薪資','獎金','年終獎金','股利','利息','其他收入']`
- **股票操作類型**：買入 / 定期定額（等同買入計算）/ 賣出
- **持倉損益**：按帳戶（個人/家庭）分組顯示，含已實現損益
- **銀行帳戶**：手動新增、手動更新餘額，存 `bank_accounts_v1`

## 週覆盤靈感閉環
- 進入「週覆盤」頁，自動抓本週 `captures` 中 category='閱讀/Podcast 靈感' 的紀錄顯示
- 「複製內容 → Claude 分析」按鈕會把靈感自動帶入 prompt
- Claude 會分析哪些靈感值得下週變成行動

## 記帳分類結構
- **家用**：家庭飲食、孩子相關、家庭雜支、水電費、房租/房貸、其他家用
- **個人**：飲食、交通、娛樂、保險、買車存款、其他個人
- **投資**、**定存**（直接填金額）
- **收入**：薪資、獎金、年終獎金、股利、利息、其他收入

## 自動化設定（已完成）
- 每天早上 7:00 → 自動建立當日 AAR 範本（含格式/複選框/下拉選單），新的一天加在最上方
- 每天 18:45 → 自動合併 App 快取到 Sheets 主紀錄

## 固定觸發詞
- **「新英文講義」**：收到講義內容（PDF/文字）→ 生成課程 JSON → 告知貼到 App「匯入課程」的位置
- **「新課堂筆記」**：收 Gemini 整理的課堂內容 → 整理成複習卡片格式 → 更新進 App
- **`/週計畫`**：生成本週行程 → 寫入 Sheets
- **「週覆盤」**：自動執行以下步驟：
  1. 用 Bash 計算本週一到今天的日期（Asia/Taipei 時區，格式 YYYY/MM/DD）
  2. WebFetch GET GAS API 抓本週 AAR 資料（from=週一, to=今天）
  3. 分析資料產出報告（亮點/痛點/靈感轉行動/防呆行動/時間分配比例）
  4. WebFetch POST 寫回 GAS API（type='review', token='graceos2026read', review_type='週'）
  5. 輸出完整報告給使用者看
- **「月覆盤」**：自動執行以下步驟：
  1. 用 Bash 計算本月第一天到今天（Asia/Taipei 時區）
  2. WebFetch GET GAS API 抓本月 AAR 資料
  3. 分析產出報告（月度亮點/需改變模式/績效描述句/下月重點）
  4. WebFetch POST 寫回 GAS API（type='review', token='graceos2026read', review_type='月'）
  5. 輸出完整報告給使用者看

## 整合的獨立 App
### 幕僚沙盤
- **網址**：`https://fgfg1717.github.io/pr-practice-app/`
- **本機路徑**：`C:\Users\ASUS\Downloads\Grace-agent\pr-practice-app\`
- **AI 服務**：Groq API（`gsk_` 開頭的 key），模型 llama-3.3-70b-versatile

## Apps Script 修改注意事項
改完 `apps-script.js` 後，必須重新部署才會生效：
Google Sheets → Apps Script 編輯器 → 部署 → 管理部署 → 編輯 → 建立新版本 → 部署
