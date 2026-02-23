"""
測試所有依賴的導入
"""
try:
    print("🔍 測試基本 Python 模組...")
    import sys
    print(f"✅ Python 版本: {sys.version}")

    print("\n🔍 測試 Flask 相關...")
    import flask
    print(f"✅ Flask 版本: {flask.__version__}")

    import flask_cors
    print("✅ Flask-CORS 導入成功")

    print("\n🔍 測試數據處理...")
    import numpy as np
    print(f"✅ NumPy 版本: {np.__version__}")

    import pandas as pd
    print(f"✅ Pandas 版本: {pd.__version__}")

    import numpy_financial as npf
    print("✅ NumPy-Financial 導入成功")

    print("\n🔍 測試 Pydantic...")
    import pydantic
    print(f"✅ Pydantic 版本: {pydantic.__version__}")

    from pydantic import BaseModel, Field
    print("✅ Pydantic 基本類別導入成功")

    print("\n🔍 測試自定義模組...")
    from models.irr_models_v2 import IRRCalculationRequest
    print("✅ IRR 模型導入成功")

    from services.irr_calculator import IRRCalculatorService
    print("✅ IRR 計算服務導入成功")

    from api.irr_routes import irr_bp
    print("✅ API 路由導入成功")

    print("\n🎉 所有依賴測試通過！可以啟動應用了。")

except ImportError as e:
    print(f"❌ 導入錯誤: {e}")
    print("請執行: pip install -r requirements.txt")
except Exception as e:
    print(f"❌ 其他錯誤: {e}")
    print(f"錯誤類型: {type(e).__name__}")