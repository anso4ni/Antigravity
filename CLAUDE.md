# Claude Code 工作規範

## 必讀規則

每次開始新對話或切換思考模型時，**必須先讀取此檔案**，再進行任何操作。

---

## 專案簡介

- **名稱**：Antigravity 股票交易紀錄系統
- **框架**：Streamlit (`streamlit run astock_single.py`)
- **語言**：Python 3.12
- **GitHub**：https://github.com/anso4ni/Antigravity.git（remote: origin）

---

## 強制工作流程

### 每次更新程式碼後，必須自動推送到 GitHub

完成程式碼修改後，依序執行：

```bash
cd d:/Antigravity
git add <修改的檔案>
git commit -m "<簡短描述>"
git push origin main
```

- 不需要詢問是否要推送，**直接推送**
- commit message 使用繁體中文或英文皆可，保持簡潔

---

## 專案架構

| 檔案 | 用途 |
|------|------|
| `astock_single.py` | Streamlit 主入口 |
| `ui.py` | 所有 UI 渲染函式 |
| `config.py` | 讀寫 JSON 配置，per-user 隔離 |
| `data.py` | 市場資料抓取與計算 |
| `import_excel.py` | 從 Excel 匯入持倉/交易紀錄 |
| `google_drive.py` | Google OAuth + Drive/Sheets API |

---

## Per-User 資料隔離

- 每個 Google 帳號的資料存在獨立 JSON：`portfolio_data_{safe_email}.json`
- `config.get_config_path()` 根據 `st.session_state["google_email"]` 決定路徑
- 共用系統設定（OAuth 金鑰路徑）存在 `portfolio_config.json`

---

## Excel 匯入格式（個股持倉 sheet）

**不使用硬編碼行號**，改以關鍵字動態偵測 section：
- `col[1] == "台股"` → 進入台股 section（source="台股"）
- `col[1] == "美股"` → 進入 FT section（source="FT"）
- `col[1] == "複委託"` → 進入複委託 section（source="複委託"）
- 每行 `col[4]`（持有股數）== 0 或非數字 → 跳過（標題列、合計列）
- 各 section 股票數量不限制

讀取欄位：`col[1]`=股票代碼、`col[3]`=成本、`col[4]`=持有股數

---

## 已知重要行為

### 登入 / 登出
- 頁面重整後應保持登入狀態，直到使用者點擊「登出」
- 本地 OAuth token 存於 `google_token.json`；登入 email 快取於 `google_user_cache.json`
- auto-login 快速路徑：從 `google_user_cache.json` 讀 email（不呼叫 Google API）
- auto-login 備援路徑：快取不存在時呼叫 `authenticate_oauth()`（會建立快取）
- 登出時同時刪除 `google_token.json` 與 `google_user_cache.json`
- 登入後必須呼叫 `st.rerun()`，否則 `cfg` 仍指向匿名配置

### 資料
- 交易紀錄的 `source` 欄位為 `"FT"` / `"複委託"` / `"台股"`，用於各頁面的分類過濾
