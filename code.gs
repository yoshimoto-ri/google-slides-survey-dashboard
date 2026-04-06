// =====================================================================
// Code.gs — AI工具使用調查 即時儀表板
// 掛載於 Google 簡報（Extensions > Apps Script）
// 作者：Claude for Justin | 2026-04
// =====================================================================
//
// 【使用前必填】
//   請在 Apps Script > Project Settings > Script Properties 加入：
//     FORM_URL       → 你的 Google Form 填寫連結
//     SHEET_ID       → 回應自動產生的 Google Sheets ID
//     SHEET_NAME     → Sheet 頁籤名稱（通常是「表單回應 1」）
//     GEMINI_API_KEY → Gemini API Key（至 https://aistudio.google.com/app/apikey 免費取得）
//
// =====================================================================

// ── 預設值（若 Script Properties 未設定時使用）──────────────────────
const DEFAULT_CONFIG = {
  FORM_URL:   'YOUR_GOOGLE_FORM_URL',
  SHEET_ID:   'YOUR_SHEET_ID',
  SHEET_NAME: '表單回應 1'
};

// ── 開放式問題識別關鍵字 ─────────────────────────────────────────────
const OPEN_ENDED_KEYWORDS = [
  '其他', '想法', '建議', '原因', '說明', '補充',
  '如何', '方式', '心得', '感想', '目前', '還有'
];

// ── Gemini 模型 ──────────────────────────────────────────────────────
const GEMINI_MODEL = 'gemini-2.5-flash';


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 1. 選單與初始化
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function onOpen() {
  SlidesApp.getUi()
    .createMenu('📊 問卷互動')
    .addItem('🎯 開啟即時儀表板（第8頁投放用）', 'showSurveyDashboard')
    .addSeparator()
    .addItem('⚙️  初始設定 / 設定 Gemini Key', 'showSetupDialog')
    .addToUi();
}

function showSurveyDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1100)
    .setHeight(660);
  SlidesApp.getUi().showModalDialog(html, '📊 AI工具使用調查 ── 即時儀表板');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 2. 設定管理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return {
    formUrl:   props.FORM_URL       || DEFAULT_CONFIG.FORM_URL,
    sheetId:   props.SHEET_ID       || DEFAULT_CONFIG.SHEET_ID,
    sheetName: props.SHEET_NAME     || DEFAULT_CONFIG.SHEET_NAME,
    geminiKey: props.GEMINI_API_KEY || ''
  };
}

function saveSettings(formUrl, sheetId, sheetName, geminiKey) {
  const props = PropertiesService.getScriptProperties();
  if (formUrl)   props.setProperty('FORM_URL', formUrl);
  if (sheetId)   props.setProperty('SHEET_ID', sheetId);
  if (sheetName) props.setProperty('SHEET_NAME', sheetName);
  if (geminiKey) props.setProperty('GEMINI_API_KEY', geminiKey);
  return { success: true, message: '設定已儲存！' };
}

function showSetupDialog() {
  const config = getConfig();
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 24px; color: #333; }
        h3 { margin: 0 0 16px; color: #1a73e8; }
        label { display: block; margin-top: 12px; font-size: 13px; color: #555; font-weight: 600; }
        input { width: 100%; padding: 8px 10px; border: 1px solid #ddd; border-radius: 6px;
                font-size: 14px; box-sizing: border-box; margin-top: 4px; }
        input:focus { border-color: #1a73e8; outline: none; }
        .btn { margin-top: 20px; background: #1a73e8; color: white; border: none;
               padding: 10px 24px; border-radius: 6px; cursor: pointer; font-size: 14px; }
        .btn:hover { background: #1557b0; }
        #msg { margin-top: 12px; padding: 8px; border-radius: 6px; display: none; }
        .ok { background: #e6f4ea; color: #137333; }
        .tip { font-size: 12px; color: #888; margin-top: 4px; }
        a { color: #1a73e8; }
      </style>
    </head>
    <body>
      <h3>⚙️ 問卷儀表板設定</h3>

      <label>Google Form 填寫連結</label>
      <input id="formUrl" value="${config.formUrl === 'YOUR_GOOGLE_FORM_URL' ? '' : config.formUrl}"
             placeholder="https://docs.google.com/forms/d/..." />

      <label>回應 Google Sheet ID</label>
      <input id="sheetId" value="${config.sheetId === 'YOUR_SHEET_ID' ? '' : config.sheetId}"
             placeholder="貼上 Sheet 網址中間的 ID 字串" />
      <div class="tip">例：https://docs.google.com/spreadsheets/d/<b>這一段</b>/edit</div>

      <label>Sheet 頁籤名稱</label>
      <input id="sheetName" value="${config.sheetName}" placeholder="表單回應 1" />

      <label>Gemini API Key
        <span class="tip">（<a href="https://aistudio.google.com/app/apikey" target="_blank">免費取得</a>）</span>
      </label>
      <input id="geminiKey" type="password" placeholder="AIza..." />

      <button class="btn" onclick="save()">💾 儲存設定</button>
      <div id="msg"></div>

      <script>
        function save() {
          const formUrl   = document.getElementById('formUrl').value.trim();
          const sheetId   = document.getElementById('sheetId').value.trim();
          const sheetName = document.getElementById('sheetName').value.trim();
          const geminiKey = document.getElementById('geminiKey').value.trim();
          google.script.run
            .withSuccessHandler(r => {
              const el = document.getElementById('msg');
              el.textContent = '✅ ' + r.message;
              el.className = 'ok';
              el.style.display = 'block';
            })
            .saveSettings(formUrl, sheetId, sheetName, geminiKey);
        }
      <\/script>
    </body>
    </html>
  `).setWidth(520).setHeight(420);
  SlidesApp.getUi().showModalDialog(html, '⚙️ 設定');
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 3. 回應計數 & 資料讀取
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getResponseCount() {
  try {
    const config = getConfig();
    const ss    = SpreadsheetApp.openById(config.sheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) return { count: 0, error: '找不到 Sheet：' + config.sheetName };
    const count = Math.max(0, sheet.getLastRow() - 1);
    return { count: count };
  } catch (e) {
    Logger.log('getResponseCount error: ' + e);
    return { count: 0, error: e.toString() };
  }
}

function getOpenEndedResponses() {
  try {
    const config = getConfig();
    const ss    = SpreadsheetApp.openById(config.sheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) return { error: '找不到 Sheet：' + config.sheetName };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { texts: [], count: 0, questionNames: [] };

    const headers = data[0];
    const rows    = data.slice(1);

    // 策略 1：找含關鍵字的欄位
    let openCols = [];
    headers.forEach((h, i) => {
      if (i === 0) return;
      const title = String(h).toLowerCase();
      if (OPEN_ENDED_KEYWORDS.some(kw => title.includes(kw))) {
        openCols.push(i);
      }
    });

    // 策略 2：用平均字數偵測開放式欄位
    if (openCols.length === 0) {
      for (let i = 1; i < headers.length; i++) {
        const vals = rows.map(r => String(r[i] || '')).filter(v => v.trim());
        if (vals.length === 0) continue;
        const avgLen     = vals.reduce((sum, v) => sum + v.length, 0) / vals.length;
        const uniqueCount = new Set(vals).size;
        if (avgLen > 5 && uniqueCount > Math.min(5, rows.length * 0.5)) {
          openCols.push(i);
        }
      }
    }

    // 策略 3：送全部欄位
    if (openCols.length === 0) {
      for (let i = 1; i < headers.length; i++) openCols.push(i);
    }

    Logger.log('開放式欄位偵測：' + openCols.map(i => headers[i]).join(' | '));

    const texts = rows.map((row, rowIdx) => {
      const parts = openCols.map(ci => {
        const val = String(row[ci] || '').trim();
        return val ? `${headers[ci]}：${val}` : '';
      }).filter(Boolean);
      return parts.length > 0 ? `【第 ${rowIdx + 1} 位】\n${parts.join('\n')}` : '';
    }).filter(Boolean);

    return {
      texts:         texts,
      count:         rows.length,
      questionNames: openCols.map(ci => String(headers[ci]))
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

function getAllResponses() {
  try {
    const config = getConfig();
    const ss    = SpreadsheetApp.openById(config.sheetId);
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) return { error: '找不到 Sheet' };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], rows: [], count: 0 };

    return {
      headers: data[0].map(String),
      rows:    data.slice(1).map(r => r.map(v => String(v))),
      count:   data.length - 1
    };
  } catch (e) {
    return { error: e.toString() };
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 4. Gemini AI 歸納
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function runGeminiSummary() {
  const data = getAllResponses();
  if (data.error) return { error: data.error };
  if (!data.rows || data.rows.length === 0) {
    return { error: '目前還沒有任何回應資料，請等學員填寫後再歸納！' };
  }

  // 把每題的回應整理成統計摘要，再送給 Gemini
  const headers = data.headers;
  const rows    = data.rows;
  const summary = [];

  for (let i = 1; i < headers.length; i++) { // i=0 是時間戳記，從 i=1 開始
    const q    = headers[i];
    const vals = rows.map(r => String(r[i] || '').trim()).filter(Boolean);
    if (vals.length === 0) continue;

    // 統計唯一值分佈（智慧分割：只在括號外的逗號切分，並移除選項說明文字）
    const counts = {};
    vals.forEach(v => {
      smartSplit_(String(v)).map(s => normalizeOption_(s)).filter(Boolean).forEach(s => {
        counts[s] = (counts[s] || 0) + 1;
      });
    });
    const uniqueVals  = Object.keys(counts);
    const isOpenEnded = uniqueVals.length > Math.min(8, rows.length * 0.6);

    if (isOpenEnded) {
      // 開放題：直接列出所有回應
      summary.push(`【${q}】（開放式，共 ${vals.length} 份）\n` + vals.map((v, i) => `  ${i+1}. ${v}`).join('\n'));
    } else {
      // 選擇題：列出統計分佈
      const dist = Object.entries(counts)
        .sort((a, b) => b[1] - a[1])
        .map(([k, v]) => `  ${k}：${v}人（${Math.round(v/rows.length*100)}%）`)
        .join('\n');
      summary.push(`【${q}】（選擇題，共 ${vals.length} 份）\n${dist}`);
    }
  }

  return summarizeWithGemini_(summary, rows.length);
}

function summarizeWithGemini_(summaryData, totalCount) {
  const config = getConfig();
  const apiKey = config.geminiKey;

  if (!apiKey) {
    return { error: '尚未設定 Gemini API Key。\n請至「📊 問卷互動 > ⚙️ 初始設定」填入 API Key。' };
  }

  const combinedText = summaryData.join('\n\n');

  const prompt = `以下是「AI工具使用調查」問卷的完整統計資料，共 ${totalCount} 份回應。
請用繁體中文，依照下方格式做完整分析，每個區塊都要確實填寫：

${combinedText}

---
請依以下格式輸出（每個區塊都必須完整）：

📊 **各題重點摘要**
（針對每一題，用一句話說明最重要的發現，例如：「使用頻率：超過半數每天使用」）

📌 **整體主要發現**
（條列 3~5 點，說明這份問卷最重要的洞察）

💡 **值得課堂討論的觀察**
（1~2 句，指出特別有趣或值得深思的地方）

🌟 **有代表性的學員回應**
（直接引用 1~2 句最具代表性的開放式回應原文，加引號）

🔢 本次共分析 ${totalCount} 份回應`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },   // 官方最新寫法：Key 放 Header
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.4, maxOutputTokens: 8192 }
      }),
      muteHttpExceptions: true
    });

    const rawText = resp.getContentText();
    if (!rawText || rawText.trim() === '') {
      return { error: 'Gemini API 回傳空白，請稍後再試。\n可至 https://status.cloud.google.com 確認服務狀態。' };
    }

    const json = JSON.parse(rawText);
    if (json.error) {
      return { error: 'Gemini 回傳錯誤：' + json.error.message };
    }

    const summary = json.candidates[0].content.parts[0].text;
    return { summary: summary, count: totalCount };

  } catch (e) {
    return { error: 'API 呼叫失敗：' + e.toString() };
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 5. 輔助函數：複選題智慧分割
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// 只在括號「外面」的逗號才切分
// 例："生成式AI (如：ChatGPT, Gemini), 圖像AI" → ["生成式AI (如：ChatGPT, Gemini)", "圖像AI"]
function smartSplit_(str) {
  const parts = [];
  let depth = 0;
  let current = '';
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    if (ch === '(' || ch === '（') depth++;
    else if (ch === ')' || ch === '）') depth--;
    if ((ch === ',' || ch === '，') && depth === 0) {
      const t = current.trim();
      if (t) parts.push(t);
      current = '';
    } else {
      current += ch;
    }
  }
  const t = current.trim();
  if (t) parts.push(t);
  return parts;
}

// 移除選項中的說明文字，讓同一選項可以合併計算
// "生成式AI (如：ChatGPT, Gemini, Claude)" → "生成式AI"
// "其他 (請在下題簡述)" → "其他"
function normalizeOption_(opt) {
  return opt
    .replace(/\s*[（(]如[：:][^）)]*[）)]/g, '')
    .replace(/\s*[（(]請[^）)]*[）)]/g, '')
    .trim();
}
