# 📊 Google Slides 問卷即時儀表板 — 建置說明

> 適用情境：課程中讓學員掃 QR Code 填 Google Form，教師在簡報中即時呈現回應統計與 Gemini AI 歸納。

---

## 一、事前準備

| 項目 | 說明 |
|------|------|
| Google 簡報 | 需為 `.gslides` 格式（非 .pptx），才有「擴充功能」選單 |
| Google Form | 建好問卷，取得填寫連結 |
| Google Sheet | Form 回應自動產生的試算表，取得 Sheet ID |
| Gemini API Key | 至 [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey) 免費取得 |

> ⚠️ `.pptx` 上傳後無法用 Apps Script，必須在 Google 雲端直接建立簡報，或上傳後另存為 Google 簡報格式。

---

## 二、取得 Sheet ID

Google Sheet 網址格式如下：
```
https://docs.google.com/spreadsheets/d/【這一段就是 Sheet ID】/edit
```
複製兩個 `/` 之間那段字串即可。

---

## 三、建立 Apps Script

1. 開啟 Google 簡報
2. 上方選單：**擴充功能 → Apps Script**
3. 預設會有一個空白的 `程式碼.gs`

### 3-1. 貼入 Code.gs

- 點選左側 `程式碼.gs`（或任何 `.gs` 檔案）
- **全選（Ctrl+A）→ 刪除 → 貼入** `Code.gs` 完整內容
- 按 **Ctrl+S** 儲存

### 3-2. 新增 Dashboard.html

- 點左側「檔案」旁的 **＋** → 選「HTML」
- 檔名輸入 `Dashboard`（不要加副檔名，系統自動加）
- **全選 → 刪除 → 貼入** `Dashboard.html` 完整內容
- 按 **Ctrl+S** 儲存

> ⚠️ 貼入時如出現 `SyntaxError: Unexpected token`，表示有殘留舊內容。請再次全選刪除後重貼。

---

## 四、設定 Script Properties

1. Apps Script 左側選單 → **專案設定（齒輪圖示）**
2. 下滑找到「**指令碼屬性**」→ 點「**新增屬性**」
3. 依序加入以下四組：

| 屬性名稱 | 值 |
|----------|-----|
| `FORM_URL` | Google Form 填寫連結（`https://docs.google.com/forms/d/...`） |
| `SHEET_ID` | 回應 Sheet 的 ID 字串 |
| `SHEET_NAME` | Sheet 頁籤名稱（通常是 `表單回應 1`） |
| `GEMINI_API_KEY` | 你的 Gemini API Key（`AIza...`） |

4. 按「**儲存指令碼屬性**」

> 💡 也可以之後透過簡報選單「⚙️ 初始設定」用介面填入，效果相同。

---

## 五、授權執行

1. 回到 Apps Script 編輯器
2. 上方選擇函式 `onOpen` → 按 ▶️ **執行**
3. 首次執行會跳出「需要授權」→ 點「**審查權限**」
4. 選擇你的 Google 帳號 → 點「**進階**」→「**前往（不安全）**」→「**允許**」

> 這個授權只需做一次。

---

## 六、啟動儀表板

1. 回到 Google 簡報，重新整理頁面
2. 上方選單出現「**📊 問卷互動**」
3. 點選「**🎯 開啟即時儀表板**」
4. Modal 視窗彈出，即時儀表板啟動

---

## 七、儀表板功能說明

| 區塊 | 功能 |
|------|------|
| 左側 QR Code | 自動產生問卷填寫 QR Code，學員掃描即可填寫 |
| 左側計數器 | 每 5 秒自動更新，顯示累積回應筆數 |
| 右側「Gemini AI 歸納」| 點按鈕，Gemini 分析所有題目並輸出摘要 |
| 右側「查看所有回應」| 按題目顯示統計長條圖（選擇/複選題）或文字列表（開放題） |

---

## 八、問卷題型對應顯示邏輯

| 題型 | 儀表板顯示方式 |
|------|--------------|
| 單選題 | 長條圖 + 人數 + 百分比 |
| 複選題（逗號分隔）| 自動拆分後，長條圖 + 人數 + 百分比 |
| Likert 量表 | 長條圖 |
| 開放式文字題 | 逐筆文字列表 |

### 複選題判斷邏輯

- 使用 `smartSplit()` 只在括號**外面**的逗號切分，避免誤切選項說明（如 `生成式AI (如：ChatGPT, Gemini)`）
- 使用 `normalizeOption()` 移除 `(如：...)` 和 `(請...)` 等說明文字，讓同一選項可合併計算
- 判斷條件：拆分後不重複選項數 ≤ 25，且至少一個選項出現超過一次 → 顯示長條圖

---

## 九、Gemini AI 歸納設定

- 模型：`gemini-2.5-flash`
- API 呼叫格式（官方最新）：Key 放 Header，**不放在 URL**
  ```javascript
  headers: { 'x-goog-api-key': apiKey }
  ```
- `maxOutputTokens: 8192`（10題以上必須調高，否則輸出會被截斷）
- 輸出格式：各題重點摘要 → 整體主要發現 → 值得討論的觀察 → 代表性學員回應

---

## 十、未來複用此模組的步驟（快速版）

1. 建新 Google Form + 對應 Sheet
2. 打開簡報 → 擴充功能 → Apps Script
3. 貼入 `Code.gs` 和 `Dashboard.html`（內容不變）
4. 更新 Script Properties：換掉 `FORM_URL`、`SHEET_ID`、`SHEET_NAME`
5. 重新授權一次（新簡報需要）
6. 完成 ✅

> 💡 `GEMINI_API_KEY` 同一個專案的 Key 可以繼續用，不用換。

---

## 十一、常見問題排解

| 問題 | 原因 | 解法 |
|------|------|------|
| 找不到「擴充功能」選單 | 檔案是 .pptx 格式 | 另存為 Google 簡報後再操作 |
| `SyntaxError` 語法錯誤 | 貼入時有殘留舊內容 | Ctrl+A 全刪後重貼 |
| `myFunction` 找不到 | 誤點執行預設函式 | 無害，忽略即可；選 `onOpen` 再執行 |
| Gemini 模型找不到 | 模型名稱錯誤 | 確認用 `gemini-2.5-flash` |
| Gemini 輸出被截斷 | `maxOutputTokens` 太低 | 調高至 `8192` |
| Gemini 回傳空白 | API 暫時不穩 | 稍等幾分鐘再試；可至 [status.cloud.google.com](https://status.cloud.google.com) 確認 |
| 複選題顯示成文字列表 | `isChoice` 判斷失敗 | 確認 `uniqueOptions.length <= 25` 且有重複選項 |
| Q1 沒有顯示 | `startCol` 設成 2 | 確認 `startCol = 1`（跳過時間戳欄即可） |

---

---

## 十二、發布到 GitHub 的安全注意事項

✅ **這三個檔案可以安全上傳：**
- `Code.gs` — 所有敏感資訊（API Key、Sheet ID）存在 Script Properties，**不在程式碼裡**
- `Dashboard.html` — 純前端邏輯，無任何金鑰
- 本說明文件 — 無敏感資訊

❌ **絕對不要上傳：**
- 含有真實 API Key 的截圖或文字
- 含有真實 Sheet ID 或 Form URL 的設定檔
- Google 帳號相關任何憑證

> 💡 Script Properties 只存在 Google Apps Script 伺服器端，不會隨 `.gs` 檔案一起被複製或公開，這是 Google 設計上的安全機制。

---

*建置完成日期：2026-04 ／ 適用問卷：AI工具使用調查*
