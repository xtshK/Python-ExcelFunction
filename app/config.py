from pathlib import Path
import platform

APP_TITLE = "CSC 4M Formatter"

def default_save_dir() -> Path:
    # 桌面/4M_formatted_result（不存在就建立）
    out = Path.home() / "Desktop" / "4M_formatted_result"
    out.mkdir(parents=True, exist_ok=True)
    return out

IS_MAC = platform.system() == "Darwin"
