# AI Agent 太陽能財務分析系統 - 專案交接文件

> **專案名稱**：財模助手
> **版本**：
> **最後更新**：2026-03-26

---

## 目錄

1. [專案概述](#1-專案概述)
2. [技術架構](#2-技術架構)
3. [環境建置](#3-環境建置)
4. [專案結構](#4-專案結構)
5. [模組說明](#5-模組說明)
6. [API 端點一覽](#6-api-端點一覽)
7. [資料庫結構](#7-資料庫結構)
8. [AI Agent 運作流程](#8-ai-agent-運作流程)
9. [Tool 工具系統](#9-tool-工具系統)
10. [前端架構](#10-前端架構)
11. [Excel 檔案管理](#11-excel-檔案管理)
12. [維護指南](#12-維護指南)
13. [常見問題排除](#13-常見問題排除)
14. [注意事項與已知限制](#14-注意事項與已知限制)

---

## 1. 專案概述

本系統是一套 **太陽能光電站財務分析平台**，整合 AI 聊天代理與 Excel 財務模型，使用者可透過自然語言對話進行：

- 查詢 Excel 工作表中的財務數據（IRR、現金流量等）
- 修改 Excel 中的參數（設備成本、售電費率等）
- 執行設備成本滾算（Price Rolling）
- 計算 IRR（專案法、成本法、權益法）
- 管理多個案場（Case）

### 核心功能

| 功能 | 說明 |
|------|------|
| AI 聊天 | 透過 Ollama 本地 LLM 進行自然語言問答 |
| Excel 操作 | 讀取、寫入、查詢 Excel 工作表 |
| IRR 計算 | 三種方法的內部報酬率計算 |
| 設備成本滾算 | 四種模式（現金、比率、條件、自訂） |
| 案場管理 | 使用者可建立多個獨立案場 |
| 使用者認證 | Email 登入制 |

---

## 2. 技術架構

### 技術棧

| 項目 | 技術 |
|------|------|
| 後端框架 | Flask 2.3.3 |
| 資料庫 | SQLite（Flask-SQLAlchemy） |
| AI/LLM | Ollama（本地部署，預設模型 qwen3:4b、qwen3:14b、qwen3:32b） |
| Excel 處理 | openpyxl |
| 財務計算 | numpy-financial |
| 前端試算表 | Luckysheet（CDN 載入） |
| 公式重算 | LibreOffice（Headless 模式） |

### 系統架構圖

```
┌─────────────────────────────────────────────────────────┐
│                      前端 (Browser)                      │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐  │
│  │  Luckysheet  │  │   聊天介面    │  │  滾算對話框    │  │
│  └──────┬───────┘  └──────┬───────┘  └───────┬───────┘  │
└─────────┼─────────────────┼──────────────────┼──────────┘
          │                 │                  │
          ▼                 ▼                  ▼
┌─────────────────────────────────────────────────────────┐
│                    Flask API Layer                        │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐  │
│  │ agent_routes │  │ auth_routes  │  │ case_routes   │  │
│  └──────┬───────┘  └──────┬───────┘  └───────┬───────┘  │
└─────────┼─────────────────┼──────────────────┼──────────┘
          │                 │                  │
          ▼                 ▼                  ▼
┌──────────────────┐  ┌──────────┐  ┌──────────────────┐
│   AI Agent Core  │  │ Database │  │   Tool System    │
│  ┌────────────┐  │  │ (SQLite) │  │  ┌────────────┐  │
│  │ Ollama LLM │  │  └──────────┘  │  │ Excel Tool │  │
│  └────────────┘  │                │  │ Finance    │  │
│  ┌────────────┐  │                │  │ Price Roll │  │
│  │ Conversa-  │  │                │  └────────────┘  │
│  │ tion Mgr   │  │                └──────────────────┘
│  └────────────┘  │
└──────────────────┘
```

---

## 3. 環境建置

### 前置需求

| 軟體 | 用途 | 安裝方式 |
|------|------|----------|
| Python ≥ 3.11 | 主程式 | [python.org](https://www.python.org/) |
| uv | 套件管理 | `pip install uv` 或 `winget install astral-sh.uv` |
| Ollama | 本地 LLM 推理 | [ollama.com](https://ollama.com/) |
| LibreOffice | Excel 公式重算 | [libreoffice.org](https://www.libreoffice.org/) |

### 安裝步驟

```bash
# 1. 進入專案目錄
cd ai_agent2

# 2. 安裝所有依賴（uv 會自動建立虛擬環境 .venv）
uv sync

# 3. 下載 Ollama 模型
ollama pull qwen3:4b
# 可選：更大的模型
# ollama pull qwen3:14b
# ollama pull qwen3:32b

# 4. 確認 Ollama 服務運行中
ollama list

# 5. 啟動應用
uv run python run.py
```

### 啟動後

- Web 介面：http://localhost:5000
- API 端點：http://localhost:5000/api
- 健康檢查：http://localhost:5000/api/health

---

## 4. 專案結構

```
ai_agent2/
│
├── pyproject.toml             # 專案配置與依賴定義
├── uv.lock                    # 依賴鎖定檔
├── run.py                     # 啟動入口（Web 服務）
├── app.py                     # Flask 應用工廠
├── main.py                    # CLI 聊天入口（測試用）
│
├── api/                       # API 路由層
│   ├── agent_routes.py        #   聊天 & Excel 操作端點（主要）
│   ├── auth_routes.py         #   使用者認證（Email 登入）
│   └── case_routes.py         #   案場 CRUD
│
├── core/                      # 核心 AI Agent 邏輯
│   ├── agent.py               #   AIAgent 主類別（對話協調器）
│   ├── connection.py          #   OllamaConnection（LLM 連線）
│   └── conversation.py        #   ConversationManager（對話歷史）
│
├── config/                    # 設定檔
│   └── settings.py            #   AgentConfig 與系統提示詞
│
├── models/                    # 資料模型
│   ├── database.py            #   SQLAlchemy 初始化
│   ├── db_models.py           #   ORM 模型（User, Case, ChatMessage）
│   └── irr_models_v2.py       #   Pydantic 驗證模型（IRR 參數）
│
├── tool/                      # AI Agent 工具集
│   ├── tool_manager.py        #   工具註冊與分發中心
│   ├── excel_tool.py          #   Excel 讀寫查詢
│   ├── finance_tool.py        #   IRR 財務計算
│   ├── price_rolling_tool.py  #   滾算計算（唯讀）
│   ├── equipment_cost_tool.py #   滾算 + 寫入 Excel
│   └── equipment_cost_services.py  # 滾算模式邏輯
│
├── services/                  # 業務邏輯服務
│   ├── excel_tool.py          #   Excel 工具（輔助）
│   ├── llm_service.py         #   LLM 服務包裝
│   └── irr_calculator.py      #   IRR 計算服務
│
├── utils/                     # 工具函式
│   ├── error_handler.py       #   錯誤處理 & 重試裝飾器
│   ├── recalc.py              #   LibreOffice 公式重算
│   └── formula_evaluator.py   #   公式評估
│
├── scripts/                   # 輔助腳本
│   ├── recalc.py              #   公式重算腳本
│   └── office/
│       └── soffice.py         #   LibreOffice 整合
│
├── static/                    # 前端靜態檔案
│   ├── LLMweb.css             #   主介面樣式
│   ├── LLMweb.js              #   主前端邏輯（~2000 行）
│   ├── price_rolling.css      #   滾算對話框樣式
│   ├── price_rolling.js       #   滾算功能
│   └── result_styles.css      #   結果顯示樣式
│
├── templates/                 # HTML 模板
│   └── LLMweb.html            #   主頁面（Luckysheet + 聊天）
│
└── Excel User Data/           # 使用者 Excel 檔案（執行時產生）
    └── {email}/               #   依使用者 Email 分資料夾
        └── {case_id}_{filename}.xlsx
```

---

## 5. 模組說明

### 5.1 app.py — Flask 應用工廠

- 使用 `create_app()` 工廠函數建立 Flask 實例
- 設定 CORS、Session、資料庫
- 註冊三個 Blueprint：agent_routes、auth_routes、case_routes
- 主頁路由 `/` 回傳 `LLMweb.html`

### 5.2 core/agent.py — AIAgent 主類別

```python
AIAgent:
    chat(user_input)           # 主對話方法，包含工具呼叫迴圈
    _handle_response()         # 處理 AI 回應（最多 5 輪工具呼叫）
    _strip_thinking()          # 移除 qwen3 的 <think> 標籤
    reset()                    # 重置對話
```

- **最多執行 5 輪工具呼叫**（防止無限迴圈）
- 每輪會將工具結果回傳給 LLM，由 LLM 決定是否繼續呼叫

### 5.3 core/connection.py — Ollama 連線

- 預設連線 `http://localhost:11434`（Ollama 服務）
- 支援 `send_message()`（純聊天）和 `send_message_with_tools()`（函數呼叫）
- 內建重試機制（3 次，間隔 1 秒）

### 5.4 core/conversation.py — 對話管理

- 維護 system message + 對話歷史
- 預設保留最近 10 輪對話
- 超過長度自動裁切舊對話

### 5.5 config/settings.py — 設定

```python
AgentConfig:
    model_name = "qwen3:4b"      # 可切換 14b, 32b
    temperature = 0.1             # 低創造力（財務計算需精確）
    top_p = 0.1
    context_window = 16384        # Token 上限
    max_history_length = 10       # 對話歷史保留輪數
    system_prompt = "..."         # 繁體中文系統提示詞（~600 行）
```

**系統提示詞重點規則**：
- 必須使用繁體中文回答
- 禁止自行編造數值
- IRR 值已是百分比形式（0.0272 = 2.72%），直接加 `%` 符號，禁止再乘以 100
- Excel 修改需經使用者確認

---

## 6. API 端點一覽

### Agent 相關 (`/api`)

| 方法 | 路徑 | 說明 |
|------|------|------|
| POST | `/api/agent_chat` | 主聊天端點（AI 對話 + 工具執行） |
| GET | `/api/agent/model` | 取得當前模型名稱 |
| POST | `/api/agent/model` | 切換 LLM 模型 |
| POST | `/api/reset` | 重置對話 |
| GET | `/api/history` | 取得對話歷史 |

### Excel 相關 (`/api`)

| 方法 | 路徑 | 說明 |
|------|------|------|
| GET | `/api/get_excel_defaults` | 取得 Excel 預設參數 |
| POST | `/api/upload_excel` | 上傳 Excel 檔案 |
| GET | `/api/read_excel/{case_id}` | 讀取 Excel 資料 |
| POST | `/api/rename_excel` | 重新命名 Excel 檔案 |
| POST | `/api/delete_excel` | 刪除 Excel 檔案 |
| POST | `/api/save_excel` | 儲存 Excel 修改 |
| GET | `/api/download_excel` | 下載 Excel 檔案 |
| GET | `/api/download_template` | 下載模板 |
| GET | `/api/list_case_sheets` | 列出案場所有工作表 |
| POST | `/api/import_sheets` | 從其他案場匯入工作表 |

### 滾算相關 (`/api`)

| 方法 | 路徑 | 說明 |
|------|------|------|
| POST | `/api/calculate_price_rolling` | 計算設備成本滾算 |

### 認證相關 (`/api`)

| 方法 | 路徑 | 說明 |
|------|------|------|
| POST | `/api/auth/login` | Email 登入（自動建立帳號） |
| POST | `/api/auth/logout` | 登出 |
| GET | `/api/auth/me` | 取得當前使用者 |

### 案場相關 (`/api`)

| 方法 | 路徑 | 說明 |
|------|------|------|
| GET | `/api/cases` | 列出使用者所有案場 |
| POST | `/api/cases` | 建立新案場 |
| PUT | `/api/cases/{id}` | 更新案場 |
| DELETE | `/api/cases/{id}` | 刪除案場 |
| GET | `/api/cases/{id}/messages` | 取得案場聊天記錄 |
| POST | `/api/cases/{id}/messages` | 儲存聊天記錄 |

---

## 7. 資料庫結構

使用 SQLite，檔案位於 `app.db`（自動建立）。

### ER Diagram

```
users
├── id          (INTEGER, PK, AUTO)
├── email       (VARCHAR, UNIQUE)
└── created_at  (DATETIME)
     │
     │  1:N
     ▼
cases
├── id             (INTEGER, PK, AUTO)
├── user_id        (INTEGER, FK → users.id)
├── name           (VARCHAR)
├── site_type      (VARCHAR)
├── excel_filename (VARCHAR)
└── created_at     (DATETIME)
     │
     │  1:N
     ▼
chat_messages
├── id         (INTEGER, PK, AUTO)
├── case_id    (INTEGER, FK → cases.id)
├── role       (VARCHAR: 'user' / 'bot')
├── content    (TEXT)
└── created_at (DATETIME)
```

### 注意事項

- 使用者以 Email 唯一識別，無密碼機制
- 刪除案場會連帶刪除其聊天記錄（CASCADE）
- 資料庫在 `app.py` 啟動時自動建立

---

## 8. AI Agent 運作流程

### 聊天流程

```
使用者輸入
    │
    ▼
AIAgent.chat()
    │
    ├─ 加入對話歷史
    ├─ 注入系統提示詞（含 Excel 路徑等提醒）
    │
    ▼
OllamaConnection.send_message_with_tools()
    │
    ├─ 傳送完整對話 + 工具 Schema 給 LLM
    │
    ▼
LLM 回應
    │
    ├─ 純文字回應 → 直接回傳給使用者
    │
    └─ 包含 tool_calls → 進入工具執行迴圈
         │
         ├─ ToolManager.execute_tool(name, args)
         ├─ 將工具結果加入對話
         ├─ 重新送回 LLM（最多 5 輪）
         │
         ▼
    最終文字回應 → 回傳給使用者
```

### 模型切換

系統支援三種模型大小，可在前端即時切換：
- `qwen3:4b` — 快速，適合一般查詢
- `qwen3:14b` — 中等，較好的理解能力
- `qwen3:32b` — 最佳品質，較慢

---

## 9. Tool 工具系統

### 工具清單

| 工具名稱 | 功能 | 所在檔案 |
|----------|------|----------|
| `write_excel_cell` | 寫入 Excel 儲存格 | `tool/excel_tool.py` |
| `read_excel_cell` | 讀取 Excel 儲存格 | `tool/excel_tool.py` |
| `delete_excel_cell` | 清除 Excel 儲存格 | `tool/excel_tool.py` |
| `edit_sheet_by_field` | 以欄位名稱智慧修改 | `tool/excel_tool.py` |
| `list_excel_sheets` | 列出所有工作表 | `tool/excel_tool.py` |
| `read_sheet_by_field` | 以欄位名稱查詢資料 | `tool/excel_tool.py` |
| `query_financial_data` | 查詢財務數據 | `tool/excel_tool.py` |
| `calculate_price_rolling` | 滾算（純計算） | `tool/price_rolling_tool.py` |
| `execute_price_rolling` | 滾算 + 寫入 Excel | `tool/equipment_cost_tool.py` |

### 工具運作機制

1. `tool_manager.py` 在初始化時註冊所有工具的 Schema（JSON Schema 格式）
2. Schema 傳送給 LLM，LLM 根據使用者意圖決定呼叫哪個工具
3. `ToolManager.execute_tool(name, arguments)` 分發執行
4. 結果回傳給 LLM 組織成自然語言回覆

### 新增工具步驟

1. 在 `tool/` 下建立新檔案（如 `my_tool.py`）
2. 定義工具類別和 Schema（參考 `excel_tool.py`）
3. 在 `tool_manager.py` 中 import 並註冊
4. 在 `config/settings.py` 的系統提示詞中加入工具使用說明

### Excel 欄位模糊匹配

`edit_sheet_by_field` 和 `read_sheet_by_field` 使用 `SequenceMatcher` 做模糊匹配：
- 匹配閾值約 0.6
- 使用者說「設備成本」可匹配到「設備總成本」
- 如有多個匹配會回傳最佳結果

---

## 10. 前端架構

### 檔案

| 檔案 | 說明 | 行數 |
|------|------|------|
| `templates/LLMweb.html` | 主頁面 HTML | - |
| `static/LLMweb.js` | 主前端邏輯 | ~2000 |
| `static/LLMweb.css` | 主介面樣式 | ~1063 |
| `static/price_rolling.js` | 滾算功能 | - |
| `static/price_rolling.css` | 滾算樣式 | ~386 |
| `static/result_styles.css` | 結果顯示樣式 | ~89 |

### 頁面佈局

```
┌──────────┬──────────────────────┬──────────────┐
│          │                      │              │
│  側邊欄   │     Luckysheet       │   聊天視窗    │
│ (案場列表)│    (Excel 試算表)     │  (對話氣泡)   │
│          │                      │              │
│          │                      │              │
│          │                      ├──────────────┤
│          │                      │ 輸入框+送出   │
└──────────┴──────────────────────┴──────────────┘
     左側          中間 (flex:2)      右側 (flex:1)
```

### 前端關鍵函數（LLMweb.js）

| 函數 | 功能 |
|------|------|
| `init()` | 頁面初始化 |
| `loginWithEmail()` | Email 登入 |
| `loadCasesFromDB()` | 載入案場列表 |
| `selectCase(caseId)` | 選擇案場（載入 Excel + 聊天記錄） |
| `sendMessage()` | 送出聊天訊息 |
| `uploadExcel(file)` | 上傳 Excel 檔案 |
| `renderLuckysheet(data)` | 渲染試算表 |
| `calculatePriceRolling()` | 執行滾算計算 |
| `switchModel(modelName)` | 切換 LLM 模型 |

---

## 11. Excel 檔案管理

### 儲存位置

```
Excel User Data/
└── user@email.com/
    ├── 1_模板A.xlsx        # case_id=1 的 Excel
    ├── 2_案場B.xlsx        # case_id=2 的 Excel
    └── ...
```

### 檔案定位機制

`agent_routes.py` 中的 `_find_excel_file()` 按優先順序尋找：

1. `{case_id}_{原始檔名}` — 最精確
2. `{case_name}_{原始檔名}` — 向下相容
3. `{case_id}_*.xlsx` — 前綴搜尋
4. `{case_name}_*.xlsx` — 名稱前綴搜尋
5. 包含 `{原始檔名}` 的檔案
6. 資料夾中第一個 `.xlsx`（最後手段）

### 公式重算

openpyxl 只能讀寫值，無法執行 Excel 公式。修改 Excel 後需重算：

```python
from utils.recalc import recalc
recalc(file_path)  # 呼叫 LibreOffice headless 模式重算所有公式
```

**如果 LibreOffice 未安裝**，公式值不會更新，查詢到的 IRR 等值可能過時。

---

## 12. 維護指南

### 日常操作

```bash
# 啟動服務
uv run python run.py

# 安裝新套件
uv add 套件名稱

# 同步環境（新成員 clone 後）
uv sync

# 更新所有套件
uv lock --upgrade
uv sync

# 查看已安裝套件
uv pip list
```

### 修改系統提示詞

檔案：`config/settings.py` 中的 `DEFAULT_CONFIG.system_prompt`

這是控制 AI 行為的核心設定，修改時注意：
- 保持繁體中文
- 不要移除 IRR 格式化規則
- 不要移除 Excel 修改確認流程

### 修改 API 端點

檔案位於 `api/` 目錄：
- `agent_routes.py` — 核心聊天與 Excel 功能
- `auth_routes.py` — 認證相關
- `case_routes.py` — 案場管理

所有端點都註冊在 `/api` 前綴下。

### 新增 AI 工具

1. 在 `tool/` 建立新檔案
2. 定義工具函數與 JSON Schema
3. 在 `tool/tool_manager.py` 中註冊
4. 更新 `config/settings.py` 系統提示詞

### 修改前端

- HTML：`templates/LLMweb.html`
- JavaScript：`static/LLMweb.js`（主要邏輯）
- CSS：`static/LLMweb.css`（主要樣式）
- 修改後清除瀏覽器快取或 Hard Refresh（Ctrl+Shift+R）

### 資料庫遷移

目前未使用 Flask-Migrate。如需修改資料庫結構：

1. 修改 `models/db_models.py`
2. 刪除 `app.db`（**會遺失所有資料**）
3. 重新啟動服務，自動建立新結構

建議未來加入 Flask-Migrate 做正式的資料庫遷移管理。

---

## 13. 常見問題排除

### Ollama 連線失敗

```
ConnectionError: Cannot connect to Ollama
```

**解決**：
1. 確認 Ollama 服務已啟動：`ollama list`
2. 確認模型已下載：`ollama pull qwen3:4b`
3. 確認預設埠 11434 未被佔用

### LibreOffice 公式重算失敗

```
recalc failed: soffice not found
```

**解決**：
1. 安裝 LibreOffice
2. 確認 `soffice` 在系統 PATH 中
3. Windows 預設路徑：`C:\Program Files\LibreOffice\program\soffice.exe`

### Excel 檔案找不到

**解決**：
1. 確認 `Excel User Data/{email}/` 資料夾存在
2. 確認檔案命名格式為 `{case_id}_{filename}.xlsx`
3. 檢查 `agent_routes.py` 中的 `_find_excel_file()` 邏輯

### IRR 值顯示異常

**常見問題**：IRR 顯示 272% 而非 2.72%

**原因**：系統提示詞中有明確規則，IRR 從 Excel 讀出的值已是百分比形式（0.0272 = 2.72%），不應再乘以 100。

**解決**：檢查 `config/settings.py` 系統提示詞中的 IRR 格式化規則是否完整。

### 模組 Import 錯誤

**解決**：
1. 確認在專案根目錄啟動：`cd ai_agent2`
2. 使用 `uv run python run.py` 而非直接 `python run.py`
3. `app.py` 會自動將當前目錄加入 `sys.path`

---

## 14. 注意事項與已知限制

### 安全性

- `SECRET_KEY` 目前是寫死在 `app.py` 中（`'ai-agent-secret-key-2025'`），正式環境應改為環境變數
- CORS 設定為 `origins: "*"`（允許所有來源），正式環境應限縮
- 使用者認證僅靠 Email，無密碼驗證機制
- SQLite 不適合高併發，正式環境建議換 PostgreSQL

### 效能限制

- Ollama 本地推理速度取決於 GPU 記憶體與模型大小
- 對話歷史最多 10 輪，超過會被裁切
- 工具呼叫最多 5 輪（防止無限迴圈）
- LibreOffice 公式重算為同步呼叫，大檔案會阻塞

### Excel 限制

- openpyxl 不支援 `.xls` 格式（僅 `.xlsx`）
- 巨集（Macro）不被保留
- 複雜的 Excel 公式可能無法被 LibreOffice 正確重算
- 模糊匹配可能在欄位名稱高度相似時出錯

### 部署注意

- `run.py` 使用 Flask 內建伺服器（`debug=True`），**僅適合開發**
- 正式部署應使用 Gunicorn（Linux）或 Waitress（Windows）
- 確保 `Excel User Data/` 目錄有寫入權限
- 確保 `app.db` 所在目錄有寫入權限

---

## 附錄：專案依賴

```toml
# pyproject.toml
dependencies = [
    "flask==2.3.3",
    "flask-cors==4.0.0",
    "flask-sqlalchemy==3.1.1",
    "numpy>=2.4.3",
    "numpy-financial==1.0.0",
    "ollama>=0.6.1",
    "openpyxl==3.1.2",
    "pandas>=3.0.1",
    "pydantic>=2.12.5",
    "python-dotenv==1.0.0",
    "requests==2.31.0",
]
```

## 附錄：術語對照

| 縮寫 | 全稱 | 中文 |
|------|------|------|
| IRR | Internal Rate of Return | 內部報酬率 |
| kWp | Kilowatt Peak | 峰值千瓦 |
| FIT | Feed-in Tariff | 躉售費率 |
| PV | Photovoltaic | 太陽能光電 |
| LLM | Large Language Model | 大型語言模型 |
