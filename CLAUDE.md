# AI 競賽複審評分平台 — CLAUDE.md

本文件以此 repo（高中組）為模板，說明如何快速複製建構同樣的評分平台給不同組別使用。  
目前已完整實作「高中組」版本，本文件同時附上「國小組」的差異規格，供後續建構參考。

---

## 專案概述

**用途**：供評審委員對競賽作品進行複審評分，可輸出 CSV 報告。  
**技術**：純靜態 HTML/CSS/JS 單一檔案，部署於 GitHub Pages，無需後端。  
**核心功能**：
- 左側作品清單（從 xlsx 抽取）
- 中間作品展示（PDF.js / 圖片 / Excel 試算表 / Word 報名表）
- 右側評分工作區（快速填分、逐項評分、AI 評分、CSV 輸出）
- localStorage 自動暫存，平板/手機優先設計

---

## 目錄結構（高中組 ← 此 repo）

```
repo-root/
├── index.html                  # 主程式（單一檔案完整平台）
├── CLAUDE.md                   # 本說明文件
├── .github/
│   └── workflows/
│       └── pages.yml           # GitHub Pages 自動部署
├── PDF/                        # 作品說明書 PDF（高中組直接放根目錄下）
│   ├── 1-消息便利店.pdf
│   ├── 2-別落入搶票詐騙!.pdf
│   └── ...
└── 【...】高中組評分表單.xlsx   # 初審評分資料來源
```

---

## 高中組 `index.html` 核心架構說明

### Config 區塊（JS 最上方）

```javascript
const PDF_BASE = 'PDF/';          // PDF 根目錄
const QUICK_SCORES = [14, 16, 18, 20];  // 快速填分按鈕

const CRITERIA = [                // 評分項目（5項，各 weight 總和需為 1.0）
  { key: 'ethics',    name: 'AI 倫理與工具運用',  weight: 0.2, desc: '...' },
  { key: 'media',     name: '媒體素養與查證邏輯',  weight: 0.2, desc: '...' },
  { key: 'narrative', name: '敘事創意與視覺表現',  weight: 0.2, desc: '...' },
  { key: 'value',     name: '宣導價值與實用性',    weight: 0.2, desc: '...' },
  { key: 'complete',  name: '整體完成度',          weight: 0.2, desc: '...' },
];

const WORKS = [
  {
    id: 1,
    title: '消息便利店',
    videoUrl: 'https://youtu.be/XXUG_2O6oYM',
    docUrl: 'https://drive.google.com/...',     // 備用開啟連結
    pdfFile: '1-消息便利店.pdf',                 // PDF 檔名（PDF_BASE 下）
    initial: { ethics:18, media:18, narrative:19, value:18, complete:18 }
  },
  // ...
];
```

### 中間 Viewer（高中組：2 個 Tab）

| Tab | ID | 功能 |
|-----|-----|------|
| 📄 作品說明書 | `pdf-pane` | PDF.js 渲染，支援翻頁/縮放 |
| 🎬 影片 | `video-pane` | YouTube 嵌入或外部開啟按鈕 |

### 左側作品列表（高中組）

每筆顯示：`#編號`、`作品題目`、`複審分數 / 初審分數`

---

## 複製新組別的步驟（通用）

### Step 1：建立新 GitHub Repository

```bash
# 建立新 repo，例如 AIeval-elementary
git clone https://github.com/{username}/{new-repo}.git
```

### Step 2：複製核心檔案

從本 repo 複製：
- `index.html` → 作為修改起點
- `.github/workflows/pages.yml` → 不需修改

### Step 3：從 xlsx 抽取資料

```python
import openpyxl

wb = openpyxl.load_workbook('評分表單.xlsx')
ws = wb['工作表名稱']  # 確認工作表名稱

for row in ws.iter_rows(min_row=6, max_row=200, max_col=15, values_only=True):
    if not row[0]:
        break
    print(row)

# 評分標準通常在另一個工作表
ws2 = wb['評分標準']
for row in ws2.iter_rows(min_row=1, max_row=10, max_col=3, values_only=True):
    if row[0]:
        print(f"標準: {row[0]}, 權重: {row[1]}, 說明: {row[2]}")
```

### Step 4：修改 index.html（詳見各組差異說明）

### Step 5：設定 GitHub Pages

1. Push 所有檔案到 `main` 分支
2. GitHub repo → **Settings → Pages → Source → GitHub Actions** → 儲存
3. 等待 CI 完成，網站上線於 `https://{username}.github.io/{repo-name}/`

---

## 國小組差異規格

> 本節為建構「國小組」評分系統的完整說明，  
> 介面、操作、評分核心與高中組相同，以下僅列出差異點。

---

### 差異一：目錄結構

```
repo-root/
├── index.html
├── CLAUDE.md
├── .github/workflows/pages.yml
├── 國小/
│   ├── 1-作品題目/
│   │   ├── 作品.pdf          ← 或 .jpg / .png（作品本體）
│   │   ├── 報名表.pdf        ← 或 .docx（報名表單）
│   │   ├── 對話歷程1.xlsx    ← AI 對話紀錄（必填）
│   │   └── 對話歷程2.xlsx    ← AI 對話紀錄（選填，無則不建立此檔）
│   ├── 2-作品題目/
│   │   └── ...
│   └── ...
└── 評分表單.xlsx
```

> **資料夾命名規則**：`{組別編號}-{作品題目}/`  
> **檔案命名規則**：固定使用 `作品`、`報名表`、`對話歷程1`、`對話歷程2` 作為檔名主體，副檔名依實際格式填寫。

---

### 差異二：WORKS 陣列（新增 `school`、`grade` 欄位及 4 個檔案欄位）

高中組 WORKS 只需 `pdfFile` 和 `videoUrl`；  
國小組改為 4 個獨立檔案欄位，移除 `videoUrl`：

```javascript
const FILE_BASE = '國小/';   // 取代高中組的 PDF_BASE

const WORKS = [
  {
    id: 1,
    title: '作品題目',
    school: '○○國小',          // ← 新增：學校名稱
    grade: '五年級',            // ← 新增：年級
    // 4 個檔案欄位（副檔名依實際格式填寫）
    workFile:    '作品.pdf',         // 作品（PDF/JPG/PNG 等）
    entryFile:   '報名表.docx',      // 報名表（PDF/DOCX）
    dialogFile1: '對話歷程1.xlsx',   // AI 對話歷程1（XLSX，必填）
    dialogFile2: '對話歷程2.xlsx',   // AI 對話歷程2（XLSX，null 表示無）
    initial: {                        // 初審分數（key 需與 CRITERIA 對應）
      item1: 18,
      item2: 17,
      item3: 19,
      item4: 18,
      item5: 18
    }
  },
  // ...
];
```

> 若某作品沒有第二份對話歷程，將 `dialogFile2` 設為 `null`，平台會顯示灰色佔位欄位。

---

### 差異三：左側作品列表（新增學校與年級顯示）

在 `renderWorkList()` 函式中，每筆作品卡片需顯示學校及年級。

**CSS 新增（加在既有 `.work-item .title` 之後）**：

```css
.work-item .meta {
  font-size: 0.70rem;
  color: var(--muted);
  margin-top: 2px;
}
```

**`renderWorkList()` 中的卡片 HTML 模板**：

```javascript
return `<div class="work-item${active}" onclick="selectWork(${w.id})">
  <div class="num">#${w.id}</div>
  <div class="title">${w.title}</div>
  <div class="meta">${w.school} · ${w.grade}</div>
  <div class="score-badge${done?' done':''}">複審: ${total ?? '—'} / 初審: ${calcInitial(w)}</div>
</div>`;
```

---

### 差異四：中間 Viewer（4 個 Tab 取代原本 2 個）

中間面板從 2 個 Tab（作品說明書 / 影片）改為 4 個 Tab：

| 位置 | Tab 名稱 | 格式 | 備註 |
|------|---------|------|------|
| 1 | 📄 作品 | PDF / JPG / PNG 等 | PDF.js 或 `<img>` |
| 2 | 📋 報名表 | PDF / DOCX | PDF.js 或外部開啟 |
| 3 | 📊 對話歷程1 | XLSX | SheetJS 渲染為 HTML 表格 |
| 4 | 📊 對話歷程2 | XLSX | 無檔案時顯示灰色佔位 |

#### HTML 結構

```html
<div class="viewer-tabs">
  <button class="tab-btn active"   onclick="switchTab('work')"    id="tab-work">📄 作品</button>
  <button class="tab-btn"          onclick="switchTab('entry')"   id="tab-entry">📋 報名表</button>
  <button class="tab-btn"          onclick="switchTab('dialog1')" id="tab-dialog1">📊 對話歷程1</button>
  <button class="tab-btn disabled" onclick="switchTab('dialog2')" id="tab-dialog2">📊 對話歷程2</button>
</div>

<div class="viewer-content">
  <div class="tab-pane active" id="work-pane">
    <div id="pdf-controls"><!-- 翻頁控制，圖片時隱藏 --></div>
    <div id="work-container"></div>
  </div>
  <div class="tab-pane" id="entry-pane">
    <div id="entry-container"></div>
  </div>
  <!-- Tab 3: 對話歷程1（Excel + 儲存格預覽列） -->
  <div class="tab-pane" id="dialog1-pane">
    <div id="dialog1-container" style="flex:1;overflow:auto;width:100%"></div>
    <div class="cell-preview-bar" id="dialog1-preview">
      <span class="cell-preview-label">儲存格內容檢視</span>
      <span class="cell-preview-addr" id="dialog1-preview-addr">—</span>
      <span class="cell-preview-content" id="dialog1-preview-content">點擊儲存格以檢視完整內容</span>
    </div>
  </div>

  <!-- Tab 4: 對話歷程2（Excel + 儲存格預覽列，或灰色佔位） -->
  <div class="tab-pane" id="dialog2-pane">
    <div id="dialog2-container" style="flex:1;overflow:auto;width:100%"></div>
    <div class="cell-preview-bar" id="dialog2-preview">
      <span class="cell-preview-label">儲存格內容檢視</span>
      <span class="cell-preview-addr" id="dialog2-preview-addr">—</span>
      <span class="cell-preview-content" id="dialog2-preview-content">點擊儲存格以檢視完整內容</span>
    </div>
  </div>
</div>
```

> **注意**：`dialog1-pane` 和 `dialog2-pane` 的 `.tab-pane.active` 需改為 `flex-direction: column`，  
> 才能讓表格區與預覽列上下排列：
>
> ```css
> #dialog1-pane.active, #dialog2-pane.active { flex-direction: column; }
> ```

#### CSS 新增

```css
/* Tab 停用（第四個無檔案時） */
.tab-btn.disabled {
  color: var(--border);
  cursor: not-allowed;
  opacity: 0.45;
}
.tab-btn.disabled:hover { border-bottom-color: transparent; }

/* 灰色佔位區 */
.no-file-placeholder {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: var(--muted);
  gap: 10px;
  font-size: 0.9rem;
}
.no-file-placeholder .icon { font-size: 2.5rem; opacity: 0.25; }

/* Excel 表格樣式 */
.xlsx-table { border-collapse: collapse; font-size: 0.78rem; min-width: 100%; }
.xlsx-table th, .xlsx-table td { border: 1px solid var(--border); padding: 4px 8px; white-space: nowrap; cursor: pointer; }
.xlsx-table td:hover { background: var(--hover); }
.xlsx-table td.selected { background: #1e2a50; outline: 2px solid var(--accent); outline-offset: -1px; }
.xlsx-table tr:nth-child(even) td { background: var(--score-bg); }
.xlsx-table tr:nth-child(even) td:hover { background: var(--hover); }
.xlsx-table tr:nth-child(even) td.selected { background: #1e2a50; }

/* 儲存格內容預覽列 */
.cell-preview-bar {
  flex-shrink: 0;
  background: #12142a;
  border-top: 1px solid var(--border);
  padding: 6px 10px;
  display: flex;
  align-items: flex-start;
  gap: 8px;
  font-size: 0.78rem;
  min-height: 38px;
  max-height: 100px;
}
.cell-preview-label {
  color: var(--accent);
  font-weight: 600;
  white-space: nowrap;
  padding-top: 2px;
  font-size: 0.72rem;
}
.cell-preview-addr {
  color: var(--muted);
  white-space: nowrap;
  padding-top: 2px;
  min-width: 50px;
}
.cell-preview-content {
  color: var(--text);
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-all;
  overflow-y: auto;
  flex: 1;
}
```

---

### 差異五：CDN 引入（需新增 SheetJS）

在 `<head>` 或 `<body>` 末尾的 PDF.js CDN 之後加入：

```html
<!-- PDF.js（高中組已有） -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<!-- SheetJS（國小組新增，用於 Excel 渲染） -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
```

---

### 差異六：Tab 切換與各 Tab 呈現邏輯

#### 輔助函式

```javascript
const FILE_BASE = '國小/';
const TAB_IDS = ['work', 'entry', 'dialog1', 'dialog2'];

function getFilePath(w, field) {
  if (!w[field]) return null;
  return `${FILE_BASE}${w.id}-${w.title}/${w[field]}`;
}

function getExt(filename) {
  return filename ? filename.split('.').pop().toLowerCase() : '';
}

function extOpenBtn(url, label) {
  return `<div class="ext-btn-wrap">
    <p style="color:var(--muted);font-size:0.85rem">此格式無法直接預覽</p>
    <a class="ext-link-btn" href="${url}" target="_blank">⬇ 開啟 / 下載 ${label}</a>
  </div>`;
}
```

#### `switchTab()` 函式

```javascript
function switchTab(tab) {
  // 若 dialog2 停用則阻擋切換
  if (tab === 'dialog2' && !currentWork?.dialogFile2) return;

  activeTab = tab;
  TAB_IDS.forEach(id => {
    document.getElementById('tab-' + id)?.classList.toggle('active', id === tab);
    document.getElementById(id + '-pane')?.classList.toggle('active', id === tab);
  });
  renderActiveTab();
}

function renderActiveTab() {
  if (!currentWork) return;
  switch (activeTab) {
    case 'work':    renderWorkFile();  break;
    case 'entry':   renderEntryFile(); break;
    case 'dialog1': renderExcel(1);    break;
    case 'dialog2': renderExcel(2);    break;
  }
}
```

#### Tab 1 — 作品（PDF 或圖片）

```javascript
async function renderWorkFile() {
  const w = currentWork;
  const path = getFilePath(w, 'workFile');
  const ext = getExt(w.workFile);
  const container = document.getElementById('work-container');
  const controls = document.getElementById('pdf-controls');

  if (['jpg','jpeg','png','gif','webp'].includes(ext)) {
    controls.style.display = 'none';
    container.innerHTML = `<div style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;padding:12px">
      <img src="${path}" style="max-width:100%;max-height:100%;object-fit:contain">
    </div>`;
  } else if (ext === 'pdf') {
    controls.style.display = 'flex';
    await loadPdf(path, 'work-container');  // 複用高中組 PDF 邏輯
  } else {
    controls.style.display = 'none';
    container.innerHTML = extOpenBtn(path, '作品');
  }
}
```

> `loadPdf(url, containerId)` 需從高中組的 `loadPdf()` 重構為接受參數的版本，  
> 原本寫死 `'pdf-container'` 改為傳入 `containerId`。

#### Tab 2 — 報名表（PDF 或 Word）

```javascript
async function renderEntryFile() {
  const w = currentWork;
  const path = getFilePath(w, 'entryFile');
  const ext = getExt(w.entryFile);
  const container = document.getElementById('entry-container');

  if (ext === 'pdf') {
    await loadPdf(path, 'entry-container');
  } else {
    // .docx / .doc → 外部開啟按鈕
    container.innerHTML = extOpenBtn(path, '報名表');
  }
}
```

#### Tab 3 & 4 — 對話歷程（Excel，含儲存格內容預覽）

```javascript
// 欄號轉 Excel 字母欄標（0→A, 1→B, 25→Z, 26→AA…）
function colLabel(n) {
  let s = '';
  for (n++; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(((n - 1) % 26) + 65) + s;
  return s;
}

async function renderExcel(num) {
  const field = `dialogFile${num}`;
  const containerId = `dialog${num}-container`;
  const previewAddr    = document.getElementById(`dialog${num}-preview-addr`);
  const previewContent = document.getElementById(`dialog${num}-preview-content`);
  const container = document.getElementById(containerId);
  const w = currentWork;

  // 清空預覽列
  if (previewAddr)    previewAddr.textContent    = '—';
  if (previewContent) previewContent.textContent = '點擊儲存格以檢視完整內容';

  if (!w[field]) {
    container.innerHTML = `<div class="no-file-placeholder">
      <div class="icon">📊</div>
      <div>無第 ${num} 份對話歷程</div>
    </div>`;
    // 無檔案時隱藏預覽列
    const bar = document.getElementById(`dialog${num}-preview`);
    if (bar) bar.style.display = 'none';
    return;
  }

  // 有檔案時確保預覽列可見
  const bar = document.getElementById(`dialog${num}-preview`);
  if (bar) bar.style.display = 'flex';

  const path = getFilePath(w, field);
  container.innerHTML = '<p style="color:var(--muted);padding:20px">載入中…</p>';

  try {
    const resp = await fetch(path);
    const buf  = await resp.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];

    // 保留完整原始值（不截斷）以供預覽列顯示
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    const rows = data.map((row, ri) => {
      const isHeader = ri === 0;
      const cells = row.map((cell, ci) => {
        const tag  = isHeader ? 'th' : 'td';
        const addr = `${colLabel(ci)}${ri + 1}`;        // 例如 A1、E4
        const disp = String(cell ?? '');
        // 顯示用截斷（避免欄位過寬），完整值存於 data-full
        const short = disp.length > 60 ? disp.slice(0, 60) + '…' : disp;
        return `<${tag} data-addr="${addr}" data-full="${disp.replace(/"/g,'&quot;')}">${short}</${tag}>`;
      }).join('');
      return `<tr>${cells}</tr>`;
    }).join('');

    container.innerHTML = `<div style="padding:10px;overflow:auto;width:100%;height:100%">
      <table class="xlsx-table" id="xlsx-table-${num}">${rows}</table>
    </div>`;

    // 點擊儲存格 → 更新預覽列
    document.getElementById(`xlsx-table-${num}`).addEventListener('click', e => {
      const cell = e.target.closest('td, th');
      if (!cell) return;

      // 清除上一個選取樣式
      document.querySelectorAll(`#xlsx-table-${num} .selected`).forEach(el => el.classList.remove('selected'));
      cell.classList.add('selected');

      const addr = cell.dataset.addr || '';
      const full = cell.dataset.full ?? cell.textContent;
      if (previewAddr)    previewAddr.textContent    = addr;
      if (previewContent) previewContent.textContent = full || '（空白）';
    });

  } catch(e) {
    container.innerHTML = extOpenBtn(path, `對話歷程${num}`);
  }
}
```

---

### 差異七：選擇作品時同步更新 Tab4 狀態

```javascript
function selectWork(id) {
  currentWork = WORKS.find(w => w.id === id);
  renderWorkList();
  renderScoring();

  // 更新 Tab4 停用/啟用狀態
  const tab4 = document.getElementById('tab-dialog2');
  if (tab4) {
    const has = !!currentWork.dialogFile2;
    tab4.classList.toggle('disabled', !has);
    tab4.title = has ? '' : '此作品無第二份對話歷程';
    // 若目前停在 Tab4 但新作品無檔案，切回 Tab1
    if (activeTab === 'dialog2' && !has) switchTab('work');
  }

  // 預設切回 Tab1（作品）
  if (activeTab !== 'dialog2') switchTab(activeTab);
  else renderActiveTab();
}
```

---

### 差異八：CSV 輸出（新增學校、年級欄位）

```javascript
function exportCSV() {
  const reviewer = localStorage.getItem('reviewer_name') || '未設定';
  const date = new Date().toISOString().slice(0, 10);

  const header = ['組別', '學校', '年級', '題目',
                  ...CRITERIA.map(c => c.name),
                  '複審總分', '初審總分', '評語', '評審姓名'];
  const rows = [header];
  const s = loadStorage();

  WORKS.forEach(w => {
    const d = s[w.id] || {};
    const scores = CRITERIA.map(c => d[c.key] != null ? d[c.key] : '');
    const total  = calcTotal(w, d);
    rows.push([w.id, w.school, w.grade, w.title,
               ...scores, total ?? '', calcInitial(w), d.comment || '', reviewer]);
  });

  const csv  = '﻿' + rows.map(r =>
    r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')
  ).join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = `複審評分_${reviewer}_${date}.csv`;
  a.click();
}
```

---

### 差異九：其他需修改的設定值

```javascript
// localStorage key（避免與高中組衝突）
const STORAGE_KEY = 'aieval_es_scores';

// 頁面標題
// <title>AI識詐行動 — 國小組複審評分平台</title>
// <h1>🛡 AI識詐行動 — 國小組複審評分平台</h1>

// 作品列表數量顯示
// <span>作品列表 (XX件)</span>

// loadPdf() 需重構為接受 containerId 參數：
async function loadPdf(url, containerId) {
  const container = document.getElementById(containerId);
  document.getElementById('pdf-page-info').textContent = '載入中…';
  container.innerHTML = '<p style="color:var(--muted);margin-top:40px">載入 PDF 中…</p>';
  try {
    pdfDoc = await pdfjsLib.getDocument(url).promise;
    pdfPage = 1;
    await renderPdfPage(containerId);
  } catch (e) {
    container.innerHTML = `<div style="color:#f87171;padding:20px;text-align:center">
      PDF 載入失敗：${e.message}</div>`;
  }
}

async function renderPdfPage(containerId) {
  if (!pdfDoc) return;
  if (pdfRenderTask) pdfRenderTask.cancel();
  const page = await pdfDoc.getPage(pdfPage);
  const vp   = page.getViewport({ scale: pdfScale });
  const canvas = document.createElement('canvas');
  canvas.width = vp.width; canvas.height = vp.height;
  const container = document.getElementById(containerId || 'work-container');
  container.innerHTML = '';
  container.appendChild(canvas);
  pdfRenderTask = page.render({ canvasContext: canvas.getContext('2d'), viewport: vp });
  await pdfRenderTask.promise;
  document.getElementById('pdf-page-info').textContent = `${pdfPage} / ${pdfDoc.numPages}`;
}
```

---

## 從 xlsx 抽取國小組資料

國小組 xlsx 欄位需確認實際位置，典型欄位如下（請依實際表單調整）：

| 欄 | 內容 |
|----|------|
| A | 組別編號 |
| B | 作品題目 |
| C | 學校名稱（新增） |
| D | 年級（新增） |
| E～I | 評分項目1～5 初審分數 |
| J | 作品說明書備用連結（可略） |

Python 抽取腳本範例：

```python
import openpyxl

wb = openpyxl.load_workbook('國小組評分表單.xlsx')
ws = wb['國小組']

for row in ws.iter_rows(min_row=6, max_row=200, max_col=12, values_only=True):
    if not row[0]:
        break
    print({
        'id':     row[0],
        'title':  row[1],
        'school': row[2],
        'grade':  row[3],
        'item1':  row[4],
        'item2':  row[5],
        'item3':  row[6],
        'item4':  row[7],
        'item5':  row[8],
    })
```

---

## 技術依賴

| 套件 | 版本 | 用途 | 組別 |
|------|------|------|------|
| PDF.js | 3.11.174 | PDF 渲染 | 高中組 & 國小組 |
| SheetJS (xlsx) | 0.18.5 | Excel 試算表渲染 | 國小組新增 |
| Claude API | claude-haiku-4-5-20251001 | AI 評分（選用） | 兩組共用 |

無其他框架依賴，純 HTML/CSS/JS。

---

## 快速填分設定

```javascript
const QUICK_SCORES = [14, 16, 18, 20];   // 可依組別需求修改
```

---

## 常見問題

**Q：Excel 無法顯示？**  
A：SheetJS fetch 需要 HTTPS 環境，本機直接開啟 `file://` 會遇到 CORS 限制。請 deploy 到 GitHub Pages 後測試。

**Q：Word 報名表無法預覽？**  
A：`.docx` 格式目前以「外部開啟 / 下載」按鈕處理。若需要直接預覽，可引入 `mammoth.js` 將 docx 轉為 HTML，但需額外實作。

**Q：第四個 Tab 仍可點擊？**  
A：確認 `selectWork()` 中有呼叫 `tab4.classList.toggle('disabled', !has)`，並在 `switchTab()` 開頭加入判斷：
```javascript
if (tab === 'dialog2' && !currentWork?.dialogFile2) return;
```

**Q：CSV 中學校/年級是空的？**  
A：確認 WORKS 陣列每筆都有 `school` 和 `grade` 欄位，且 `exportCSV()` 的 header 及 row 都已加入這兩欄。

**Q：兩組評分資料互相干擾？**  
A：確認國小組 `STORAGE_KEY = 'aieval_es_scores'`（高中組為 `aieval_hs_scores'`），兩者使用不同 key 存在 localStorage 中。

**Q：圖片作品顯示變形？**  
A：圖片使用 `object-fit: contain` 等比縮放，若仍有問題確認容器高度設定正確（`height: 100%` 需要父層也有明確高度）。

**Q：點擊儲存格後預覽列沒有更新？**  
A：確認 `renderExcel()` 內有正確設定 `data-addr` 與 `data-full` 屬性，並且 `addEventListener` 掛在 `#xlsx-table-${num}` 上（而非 container）。若 container 被 innerHTML 整個替換，需重新綁定事件。

**Q：儲存格內容顯示被截斷？**  
A：表格顯示欄位限制 60 字元（避免欄寬過大），完整內容儲存於 `data-full` 屬性並在下方預覽列完整顯示，不會遺失任何文字。

**Q：預覽列在無第二份對話歷程時仍顯示？**  
A：確認 `renderExcel()` 中 `!w[field]` 的分支有執行 `bar.style.display = 'none'`；有檔案時則設為 `'flex'`。
