# AI 競賽複審評分平台 — CLAUDE.md

本文件說明如何依照此專案架構，快速複製建構另一組（如高中組）的複審評分平台。

---

## 專案概述

**用途**：供評審委員對競賽作品進行複審評分，可輸出 CSV 報告。  
**技術**：純靜態 HTML/CSS/JS 單一檔案，部署於 GitHub Pages，無需後端。  
**核心功能**：
- 左側作品清單（從 xlsx 抽取）
- 中間作品展示（PDF.js 渲染說明書 + YouTube 內嵌影片）
- 右側評分工作區（快速填分、逐項評分、AI 評分、CSV 輸出）
- localStorage 自動暫存，平板/手機優先設計

---

## 目錄結構

```
repo-root/
├── index.html                  # 主程式（單一檔案完整平台）
├── CLAUDE.md                   # 本說明文件
├── .github/
│   └── workflows/
│       └── pages.yml           # GitHub Pages 自動部署
├── [組別名稱]/
│   └── PDF/                    # 作品說明書 PDF 檔案夾
│       ├── 1-作品題目.pdf
│       ├── 2-作品題目.pdf
│       └── ...
└── [評分表單].xlsx              # 初審評分資料來源（參考用）
```

---

## 複製新組別的步驟

### Step 1：建立新 GitHub Repository

1. 在 GitHub 建立新 repo（例如 `AIeval-highschool`）
2. Clone 到本機

### Step 2：準備 PDF 檔案

將所有作品說明書 PDF 放入對應資料夾：
```
高中/PDF/1-作品題目.pdf
高中/PDF/2-作品題目.pdf
...
```

> **PDF 命名規則**：`{組別編號}-{作品題目}.pdf`  
> 檔名需與 `index.html` 中 `pdfFile` 欄位完全一致（含全形符號）。

### Step 3：從 xlsx 抽取資料

用以下 Python 腳本讀取 xlsx，取得所需欄位：

```python
import openpyxl

wb = openpyxl.load_workbook('評分表單.xlsx')
ws = wb['高中組']  # 改為對應的工作表名稱

for row in ws.iter_rows(min_row=6, max_row=100, max_col=10, values_only=True):
    if not row[0]:
        break
    print(f"id:{row[0]}, title:{row[1]}, scores:C={row[2]},D={row[3]},E={row[4]},F={row[5]},G={row[6]}, doc:{row[8]}, video:{row[9]}")

# 評分標準在 '評分標準' 工作表
ws2 = wb['評分標準']
for row in ws2.iter_rows(min_row=1, max_row=10, max_col=3, values_only=True):
    if row[0]:
        print(f"標準: {row[0]}, 權重: {row[1]}, 說明: {row[2]}")
```

**xlsx 欄位對照（國中組為例，高中組請確認欄位位置）**：

| 欄 | 內容 |
|----|------|
| A (col 0) | 組別編號 |
| B (col 1) | 作品題目 |
| C (col 2) | 評分項目1 初審分數 |
| D (col 3) | 評分項目2 初審分數 |
| E (col 4) | 評分項目3 初審分數 |
| F (col 5) | 評分項目4 初審分數 |
| G (col 6) | 評分項目5 初審分數 |
| I (col 8) | 作品說明書連結（Google Doc/Drive） |
| J (col 9) | 影片連結（YouTube 或其他） |

### Step 4：修改 index.html

複製本 repo 的 `index.html`，修改以下三個區塊：

#### 4-1：評分標準（`CRITERIA` 陣列）

```javascript
const CRITERIA = [
  { key: 'item1', name: '評分項目名稱1', weight: 0.2, desc: '評分說明文字。' },
  { key: 'item2', name: '評分項目名稱2', weight: 0.2, desc: '評分說明文字。' },
  { key: 'item3', name: '評分項目名稱3', weight: 0.2, desc: '評分說明文字。' },
  { key: 'item4', name: '評分項目名稱4', weight: 0.2, desc: '評分說明文字。' },
  { key: 'item5', name: '評分項目名稱5', weight: 0.2, desc: '評分說明文字。' },
];
```

> `key` 為英文識別碼（不重複即可），`weight` 總和需為 1.0。

#### 4-2：作品資料（`WORKS` 陣列）

```javascript
const WORKS = [
  {
    id: 1,
    title: '作品題目',
    videoUrl: 'https://youtu.be/XXXXXXXXXXX',      // YouTube 連結
    docUrl: 'https://docs.google.com/...',          // 說明書連結（備用開啟）
    pdfFile: '1-作品題目.pdf',                      // PDF 檔名（需與實際檔案完全一致）
    initial: {                                       // 各項初審分數（從 xlsx 讀取）
      item1: 18,
      item2: 19,
      item3: 17,
      item4: 18,
      item5: 19
    }
  },
  // ... 其餘作品
];
```

**注意事項**：
- `pdfFile` 中的特殊字元（冒號 `：`、驚嘆號 `!` 等）需與實際檔名一致
- 非 YouTube 影片連結會自動顯示「外部開啟」按鈕
- `initial` 的 key 需與 `CRITERIA` 的 `key` 完全對應

#### 4-3：PDF 路徑（`PDF_BASE`）

```javascript
const PDF_BASE = '高中/PDF/';   // 改為新組別的資料夾路徑
```

#### 4-4：頁面標題（可選）

```html
<title>AI識詐行動 — 高中組複審評分平台</title>
```

```html
<h1>AI識詐行動 高中組複審評分平台</h1>
```

### Step 5：設定 GitHub Pages

1. 複製 `.github/workflows/pages.yml` 到新 repo（內容不需修改）
2. Push 所有檔案到 `main` 分支
3. 前往 GitHub repo → **Settings → Pages → Source → GitHub Actions** → 儲存
4. 等待 CI 完成，網站即上線於 `https://{username}.github.io/{repo-name}/`

---

## 快速填分分數設定

目前預設快速填分按鈕為 `14、16、18、20`。如需修改：

```javascript
const QUICK_SCORES = [14, 16, 18, 20];   // 可改為任意分數組合
```

---

## AI 評分功能

平台支援使用 Claude API 進行 AI 輔助評分：
- 評審在右側面板點擊「🤖 AI」按鈕
- 輸入自備的 Claude API Key（`sk-ant-...`）
- AI 會根據評分標準和初審分數給出建議分數與評語
- API Key 儲存於 localStorage，下次無需重新輸入
- 使用模型：`claude-haiku-4-5-20251001`（速度快、成本低）

---

## CSV 輸出格式

點擊「⬇ 輸出 CSV」後，下載檔案包含以下欄位：

| 組別 | 題目 | 項目1 | 項目2 | 項目3 | 項目4 | 項目5 | 複審總分 | 初審總分 | 評語 | 評審姓名 |
|------|------|-------|-------|-------|-------|-------|---------|---------|------|---------|

- 含 UTF-8 BOM，Excel 開啟中文不亂碼
- 檔名格式：`複審評分_{評審姓名}_{日期}.csv`

---

## 資料儲存說明

- 所有評分儲存於瀏覽器 **localStorage**，清除瀏覽器資料會遺失
- 建議評分完成後立即輸出 CSV 備份
- 不同瀏覽器/裝置的評分資料不會同步

---

## 常見問題

**Q：PDF 無法顯示？**  
A：平台使用 PDF.js 渲染，需要網路載入 CDN。確認 PDF 檔案已正確放入 `高中/PDF/` 並 commit 到 repo。

**Q：影片無法播放？**  
A：YouTube 連結會直接嵌入；其他平台（SharePoint、Drive）會顯示「外部開啟」按鈕。建議請參賽者上傳至 YouTube。

**Q：GitHub Pages CI 失敗？**  
A：需先在 repo Settings → Pages 手動選擇 GitHub Actions 作為 Source，才能啟用自動部署。

**Q：想新增評審姓名？**  
A：點擊右上角「👤 評審姓名」按鈕設定，會顯示於 CSV 輸出。

---

## 技術依賴

| 套件 | 版本 | 用途 |
|------|------|------|
| PDF.js | 3.11.174 | PDF 渲染（CDN） |
| Claude API | claude-haiku-4-5-20251001 | AI 評分（選用） |

無其他框架依賴，純 HTML/CSS/JS。
