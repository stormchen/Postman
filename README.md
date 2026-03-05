# Mantis Bug Tracker 資料撈取與 Excel 匯出工具

## 功能說明

此工具用於自動化撈取 Mantis Bug Tracker 的資料，並將篩選後的資料匯出為 Excel 檔案。

### 主要功能：
- ✓ 自動分頁撈取 Mantis API 資料
- ✓ 過濾過去 7 天內有更新的 tickets
- ✓ 動態攤平 custom_fields 欄位
- ✓ 支援 ISO 8601 時間格式與時區處理
- ✓ 完整的錯誤處理與日誌輸出
- ✓ 匯出為帶日期的 Excel 檔案

---

## 安裝與設定

### 1. 安裝必要套件

```bash
pip install requests pandas openpyxl python-dotenv
```

### 2. 設定環境變數

將 `.env.example` 複製為 `.env`，並填入實際的設定值：

```bash
# 複製範例檔案
cp .env.example .env
```

編輯 `.env` 檔案，填入你的 Mantis API 認證資訊：

```
TOKEN=your_actual_api_token_here
ACCOUNTID=your_optional_account_id_here
```

> **重要**：
> - `TOKEN` 是必填項目，請從 Mantis 帳戶設定中取得
> - `ACCOUNTID` 是可選項目，根據 Mantis 設定的要求而定
> - `.env` 檔案應保持私密，不應提交到版本控制系統

---

## 使用方式

### 執行腳本

在命令列中執行以下命令：

```bash
python mantis_exporter.py
```

腳本完成 Excel 生成後會自動發送 Email 到 Teams 頻道（透過 SMTP），前提是 `.env` 中已正確設定 SMTP 相關參數。

- SMTP_SERVER
- SMTP_PORT
- SENDER_EMAIL
- TEAMS_CHANNEL_EMAIL

若這些設定缺失，程式會輸出警告並跳過發信步驟。

### 腳本執行流程

1. **初始化**：載入 `.env` 檔案中的認證資訊
2. **撈取資料**：使用 API 分頁撈取所有 issues（Filter ID: 734）
3. **時間篩選**：只保留過去 7 天內有更新的 issues
4. **資料提取**：提取基本欄位與攤平 custom_fields
5. **Excel 匯出**：生成命名為 `Mantis_Weekly_Update_YYYYMMDD.xlsx` 的檔案

### 執行範例

```
============================================================
Mantis Bug Tracker 資料撈取與 Excel 匯出工具
============================================================

✓ Mantis Exporter 初始化完成
  API 位址：http://10.210.2.90/mantisbt/api/rest/index.php/issues
  Filter ID：734
  Page Size：300

🔄 開始撈取 Mantis 資料...
✓ 已撈取第 1 頁 (300 筆資料)
✓ 已撳取第 2 頁 (150 筆資料)
✓ 已撈取所有資料 (共 2 頁)

📊 撈取完成：共 250 筆符合條件的資料

✓ Excel 檔案已成功生成：Mantis_Weekly_Update_20260305.xlsx
  共 250 筆資料，15 個欄位
============================================================
✓ 所有流程已完成
============================================================
```

---

## 技術細節

### API 設定

- **基礎 URL**：`http://10.210.2.90/mantisbt/api/rest/index.php/issues`
- **Filter ID**：734（固定內容，用於篩選特定的 issue 集合）
- **Page Size**：300（每頁最多 300 筆資料）
- **認證方式**：Authorization Header（Bearer Token）

### 資料欄位說明

#### 基本欄位
| 欄位名稱 | 說明 | 來源 |
|---------|------|------|
| Ticket ID | Issue 編號 | `id` |
| Project | 專案名稱 | `project.name` |
| Status | 狀態 | `status.name` |
| Severity | 嚴重級別 | `severity.name` |
| Category | 類別 | `category.name` |
| Summary | 摘要 | `summary` |
| Updated At | 更新時間 | `updated_at` |

#### 動態欄位
`custom_fields` 陣列中的欄位會動態攤平，欄位名稱由 `field.name` 決定，常見的有：
- `AI_Summary`：AI 摘要
- `Processing`：處理狀態
- 其他自訂欄位...

### 時間過濾邏輯

- 讀取每筆 issue 的 `updated_at` 欄位
- 與當前時間比對，只保留過去 7 天內更新的 issues
- 支援 ISO 8601 格式（如 `2026-03-05T12:32:20+08:00`）
- 自動處理時區資訊

### 錯誤處理

腳本包含以下錯誤處理機制：

1. **環境變數驗證**：檢查必要的 TOKEN 是否存在
2. **HTTP 狀態碼檢查**：監控 API 回應狀態
3. **超時處理**：設定 30 秒超時限制
4. **JSON 解析異常**：捕捉 JSON 解析錯誤
5. **資料提取失敗**：單筆資料失敗不會中斷整個流程
6. **Excel 匯出失敗**：提供清晰的錯誤訊息

---

## 程式碼結構

### 主要類別：`MantisExporter`

```python
MantisExporter
├── __init__()                 # 初始化與環境變數載入
├── _get_headers()             # 構建 API Header
├── _parse_datetime()          # 解析 ISO 8601 時間
├── _is_within_last_7_days()   # 檢查時間篩選
├── _flatten_custom_fields()   # 攤平 custom_fields
├── _extract_issue_data()      # 提取單筆 issue 資料
├── fetch_all_issues()         # 分頁撈取所有 issues
└── export_to_excel()          # 匯出為 Excel
```

### 主程式函式：`main()`

協調整個流程，包括錯誤捕捉與提示輸出。

---

## 常見問題

### Q: 為什麼沒有生成 Excel 檔案？
**A:** 可能是因為：
- 沒有符合「過去 7 天內有更新」的 issues
- 在這種情況下，腳本會顯示提示訊息但不生成檔案

### Q: API 請求失敗，該怎麼辦？
**A:** 請檢查以下項目：
- `.env` 檔案中的 TOKEN 是否正確
- 網路連線是否正常
- Mantis 伺服器是否線上
- API 端點是否可訪問（`http://10.210.2.90/mantisbt/api/rest/index.php/issues`）

### Q: 如何修改篩選條件？
**A:** 在腳本中修改以下值：
- `self.filter_id = "734"` - 修改篩選 ID
- `self.page_size = 300` - 修改每頁資料數
- `timedelta(days=7)` - 修改時間篩選範圍（在 `_is_within_last_7_days` 方法中）

### Q: 可以匯出其他格式嗎？
**A:** 目前支援 Excel。如需其他格式（CSV、JSON 等），可修改 `export_to_excel()` 方法或創建新的匯出方法。

---

## 許可與聯絡

如有問題或建議，請聯絡開發人員。

---

**最後更新**：2026/03/05
