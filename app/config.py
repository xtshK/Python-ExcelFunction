from pathlib import Path
import platform

APP_TITLE = "CSC 4M Formatter"

def default_save_dir() -> Path:
    if platform.system() == "Windows":
    # 桌面/4M_formatted_result（不存在就建立）
        out = Path("C:/4M formatted result")
    else:
        out = Path.home() / "4M formatted result"

    out.mkdir(parents=True, exist_ok=True)
    return out

IS_MAC = platform.system() == "Darwin"
