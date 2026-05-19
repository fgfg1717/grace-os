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

## App 功能模組（首頁 → 點卡片進入）
| 卡片 | 功能 | localStorage key |
|------|------|-----------------|
| ⚡ 快速記錄 | 6 分類（靈感/記帳/股票/健康/工作/生活）+ 語音 | `captures` |
| 💰 記帳 | 家用/個人/投資/定存 + 細項分類 | — |
| 📈 股票追蹤 | 個人+家庭，台股+美股，損益計算 | — |
| 📖 英文單字庫 | 隨手記、複製給 Claude 出題 | `vocabs` |
| 🎓 英文練習 App | iframe 嵌入，連到獨立 App | — |
| 🎯 幕僚沙盤 | iframe 嵌入，連到獨立 App | — |
| 🗓️ 本週計畫 | 從 Sheets 讀取，今日任務高亮 | — |
| 📋 讀取 AAR | 從 Sheets 讀，一鍵複製給 NotebookLM | — |
| 📊 週覆盤 | 輸入 → 複製給 Claude 分析 | — |
| 📅 月回顧 | 輸入 → 複製給 Claude 分析 | — |

## 整合的兩個獨立 App
### 英文練習 App
- **網址**：`https://fgfg1717.github.io/english-review-app/`
- **本機路徑**：`C:\Users\ASUS\Downloads\Grace-agent\english-review-app\`
- **課程 JSON 存放**：`C:\Users\ASUS\Downloads\Grace-agent\英文講義\`
- **JSONBin Sync Code**：`69e1929aaaba88219709eb0b`
- **已匯入課程**：b2c-01~b2c-02、grammar-09~grammar-10
- **匯入新課程**：App → ⚙️ 設定 → 匯入新課程（貼 JSON）→ 立即同步

### 幕僚沙盤
- **網址**：`https://fgfg1717.github.io/pr-practice-app/`
- **本機路徑**：`C:\Users\ASUS\Downloads\Grace-agent\pr-practice-app\`
- **AI 服務**：Groq API（`gsk_` 開頭的 key），模型 llama-3.3-70b-versatile
- **功能**：AI 動態出題 → 三維度應對（外部公關/內部治理/向上決策）→ AI 雙角色評分 + PDF 匯出

## Quick Capture 分類（對應 Sheets 欄位 B）
`閱讀/Podcast 靈感` / `財務記帳` / `股票紀錄` / `健康管理` / `工作` / `生活`

## 記帳分類結構
- **家用**：家庭飲食、孩子相關、家庭雜支、水電費、房租/房貸、其他家用
- **個人**：飲食、交通、娛樂、保險、買車存款、其他個人
- **投資**、**定存**（直接填金額）

## 自動化設定（已完成）
- 每天早上 7:00 → 自動建立當日 AAR 範本（含格式/複選框/下拉選單）
- 每天 18:45 → 自動合併 App 快取到 Sheets 主紀錄

## 固定觸發詞
- **「新英文講義」**：讀 PDF → 生成 JSON → 存到 `英文講義` 資料夾 → 告知匯入位置
- **「新課堂筆記」**：收 Gemini 整理的課堂內容 → 整理成複習卡片格式 → 更新進 App
- **`/週計畫`**：生成本週行程 → 寫入 Sheets
- **`/週復盤`**：讀取 Sheets 資料 → 輸出亮點/痛點/防呆行動
- **`/月復盤`**：彙整當月週復盤 → 月度總結

## Apps Script 修改注意事項
改完 `apps-script.js` 後，必須重新部署才會生效：
Google Sheets → Apps Script 編輯器 → 部署 → 管理部署 → 編輯 → 建立新版本 → 部署
