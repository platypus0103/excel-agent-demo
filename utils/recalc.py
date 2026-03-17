"""
utils/recalc.py
LibreOffice 公式重算工具（跨平台版本）
在 openpyxl wb.save() 之後呼叫 recalc(path)，
LibreOffice 會執行所有公式並將結果寫入快取欄位。
"""
import subprocess
import sys
import os
import json
from pathlib import Path

_RECALC_SCRIPT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "scripts", "recalc.py"
)

def recalc(file_path: str, timeout: int = 60) -> bool:
    script = os.path.abspath(_RECALC_SCRIPT)
    if not os.path.exists(script):
        print(f"[recalc] 找不到重算腳本: {script}", file=sys.stderr)
        return False

    try:
        result = subprocess.run(
            [sys.executable, script, file_path, str(timeout)],
            capture_output=True,
            text=True,
            timeout=timeout + 15,
            # 確保子程序繼承當前工作目錄（讓 scripts/office 的相對 import 正常）
            cwd=os.path.dirname(script),
        )
        output = result.stdout.strip()
        if not output:
            err = result.stderr.strip()
            print(f"[recalc] 無輸出，returncode={result.returncode}, stderr={err}", file=sys.stderr)
            return False

        data = json.loads(output)
        if data.get("status") == "success":
            print(f"[recalc] 成功，共重算 {data.get('total_formulas', 0)} 個公式")
            return True
        elif "error" in data:
            print(f"[recalc] 錯誤: {data['error']}", file=sys.stderr)
            return False
        else:
            print(f"[recalc] 發現公式錯誤: {data.get('error_summary')}", file=sys.stderr)
            return False

    except subprocess.TimeoutExpired:
        print(f"[recalc] 超時（>{timeout}s）", file=sys.stderr)
        return False
    except Exception as e:
        print(f"[recalc] 執行失敗: {e}", file=sys.stderr)
        return False
