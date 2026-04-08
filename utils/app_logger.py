# utils/app_logger.py
"""
應用程式操作日誌模組
將所有使用者操作、API 請求、錯誤訊息寫入 logs/app.log
每天自動換一個新檔案，保留最近 30 天
"""
import logging
import os
import datetime
from logging.handlers import TimedRotatingFileHandler


class TaiwanFormatter(logging.Formatter):
    """強制使用 UTC+8（台灣時區）格式化時間，避免系統時區設定影響"""
    _TZ = datetime.timezone(datetime.timedelta(hours=8))

    def formatTime(self, record, datefmt=None):
        dt = datetime.datetime.fromtimestamp(record.created, tz=self._TZ)
        if datefmt:
            return dt.strftime(datefmt)
        return dt.isoformat()

_logger = None


def get_logger() -> logging.Logger:
    """取得全域 logger，第一次呼叫時初始化"""
    global _logger
    if _logger is not None:
        return _logger

    # 建立 logs 目錄
    log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, 'app.log')

    logger = logging.getLogger('財模助手')
    logger.setLevel(logging.DEBUG)

    # 避免重複加 handler（Flask debug mode 會重新載入模組）
    if not logger.handlers:
        # 檔案 handler：每天換檔，保留 30 天
        file_handler = TimedRotatingFileHandler(
            log_path,
            when='midnight',
            interval=1,
            backupCount=30,
            encoding='utf-8'
        )
        file_handler.suffix = '%Y-%m-%d'
        file_handler.setLevel(logging.DEBUG)

        fmt = TaiwanFormatter(
            '[%(asctime)s] [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)

        # 同時輸出到 console（方便開發期間查看）
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(fmt)
        logger.addHandler(console_handler)

    _logger = logger
    return _logger


def log_action(user_email: str, action: str, detail: str = '', status: str = 'OK'):
    """
    記錄使用者操作

    Args:
        user_email: 使用者 Email（未登入填 'anonymous'）
        action: 操作名稱，如 'agent_chat'、'upload_excel'
        detail: 操作細節，如查詢內容、檔名等
        status: 'OK' | 'ERROR' | 'WARN'
    """
    logger = get_logger()
    msg = f"[{status}] user={user_email} action={action}"
    if detail:
        msg += f" | {detail}"

    if status == 'ERROR':
        logger.error(msg)
    elif status == 'WARN':
        logger.warning(msg)
    else:
        logger.info(msg)


def log_error(user_email: str, action: str, error: Exception, detail: str = ''):
    """
    記錄例外錯誤（含 traceback）

    Args:
        user_email: 使用者 Email
        action: 發生錯誤的操作名稱
        error: 例外物件
        detail: 額外說明
    """
    import traceback
    logger = get_logger()
    tb = traceback.format_exc()
    msg = f"[ERROR] user={user_email} action={action}"
    if detail:
        msg += f" | {detail}"
    msg += f" | exception={type(error).__name__}: {error}"
    logger.error(msg)
    logger.debug(f"[TRACEBACK] {tb}")
