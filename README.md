# 專案法 IRR 計算器 - Flask 版本

一個專業的內部報酬率 (IRR) 計算器，採用前後端分離架構，支援未來 LLM 功能擴展。

# 🏗️ 項目架構

```
web/                     # 整合後的系統目錄
├── 🚀 啟動檔案
│   ├── app.py              # Flask主應用程式
│   ├── run.py              # 應用啟動腳本
│   ├── .env                # 環境配置檔案
│   └── requirements.txt    # Python依賴清單
├── 🌐 前端檔案
│   ├── index.html          # 主要HTML頁面
│   ├── styles.css          # 所有CSS樣式
│   ├── api_client.js       # API客戶端和UI邏輯 (主要)
│   └── irr_ui.js          # UI交互邏輯 (輔助)
├── 🐍 後端檔案
│   ├── api/
│   │   ├── __init__.py
│   │   └── irr_routes.py   # Flask API路由
│   ├── models/
│   │   ├── __init__.py
│   │   └── irr_models_v2.py # Pydantic v2 數據模型
│   └── services/
│       ├── __init__.py
│       └── irr_calculator.py # 核心計算邏輯
├── 🧪 測試檔案
│   ├── test_api.py         # API功能測試
│   ├── test_import.py      # 依賴導入測試
└── 📋 文檔
    ├── README.md           # 本說明文檔
    └── __init__.py
```

# 🚀 快速開始

# 1. 安裝依賴

```bash
pip install -r requirements.txt
```

# 2. 啟動服務器

```bash
python run.py
```

# 3. 訪問應用

- **前端界面**: http://localhost:5000
- **API 文檔**: http://localhost:5000/api
- **健康檢查**: http://localhost:5000/api/irr/health

# 4. 運行測試

```bash
python test_api.py
```

# 📊 功能特色

# ✨ 前端功能
- 🎨 現代化響應式界面
- 📱 支援手機版瀏覽
- ⚡ 即時數據驗證
- 📈 詳細現金流表格
- 🎯 多種輸入模式支援

# 🔧 後端功能
- 🐍 使用 numpy-financial 優化計算
- 📝 Pydantic 數據驗證
- 🔒 完整錯誤處理
- 📊 RESTful API 設計
- 🧪 完整的測試覆蓋

# 💡 計算功能
- 💰 設備費用自動計算
- 📈 專業 IRR 演算法
- 💸 稅務計算支援
- 📊 多年期現金流分析
- ⚡ 即時結果顯示

# 🔧 API 端點

# IRR 計算
```
POST /api/irr/calculate
```

# 設備費用計算
```
POST /api/irr/equipment-cost
```

# 健康檢查
```
GET /api/irr/health
```

# 📋 輸入參數

# 設備費用參數
- **建置容量 (KW)**: 投資項目的發電容量
- **每KW價格**: 單位容量的設備價格
- **利潤率 (%)**: 業者的預期利潤率
- **開發費**: 項目開發相關費用

# 收入與支出
- **電費收入**: 每年的發電收入
- **利息費用**: 貸款利息支出
- **租金**: 場地租賃費用
- **運維費用**: 設備維護成本
- **保險費**: 投保費用
- **模組回收費**: 設備回收處理費
- **所得稅率**: 適用的稅率

# 🔮 未來擴展

此架構已預留 LLM 整合接口，未來可以輕鬆添加：
- 🤖 智能財務分析
- 💬 自然語言查詢
- 📊 AI 驅動的投資建議
- 📈 預測性現金流分析

# 🛠️ 開發說明

# 依賴說明
- **Flask**: Web 框架
- **Flask-CORS**: 跨域支援
- **numpy-financial**: 金融計算
- **pandas**: 數據處理
- **pydantic**: 數據驗證

# 開發模式
```bash
export FLASK_ENV=development
export FLASK_DEBUG=True
python run.py
```

# 生產部署
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

# 📝 使用說明

1. **設定年度範圍**: 選擇投資計畫的起始和結束年度
2. **輸入設備參數**: 填入容量、價格、利潤率等
3. **設定收入支出**: 選擇每年不同或總額攤平模式
4. **設定稅率**: 輸入適用的所得稅率
5. **計算結果**: 點擊計算按鈕獲得 IRR 和現金流分析

#   計算公式

*設備費用總額**:
```
報價總金額 = {[每KW價格 ÷ (1 - 利潤率%)] + 開發費} × 建置容量
```

*IRR 計算**:
使用 numpy-financial 的專業演算法，確保計算精度和穩定性。

# 支援

如有問題或建議，請聯繫開發團隊。