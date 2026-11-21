# -*- coding: utf-8 -*-
"""画面の起動スクリプト。

使い方:  python run_gui.py   (または run_gui.bat をダブルクリック)
"""

import os
import sys
import traceback

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)


def _report_startup_error():
    """起動に失敗したことを利用者に知らせる。

    ダブルクリックでの起動 (pythonw) ではエラーの表示先が無く、何も起きないように見えてしまうため、
    内容をファイルに残し、メッセージ ボックスでも知らせる。
    """
    detail = traceback.format_exc()
    log_path = os.path.join(APP_DIR, "startup_error.log")
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(detail)
    except OSError:
        log_path = "(保存できませんでした)"
    message = "起動できませんでした。\n\n" + detail.strip().splitlines()[-1][:300] + f"\n\n詳細: {log_path}"
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, "起動エラー", 0x10)
    except Exception:
        if sys.stderr:
            sys.stderr.write(message + "\n")


if __name__ == "__main__":
    try:
        from pipeline_app.gui import main
        main()
    except Exception:
        _report_startup_error()
        sys.exit(1)
