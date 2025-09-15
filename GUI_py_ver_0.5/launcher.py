import subprocess
import sys
import os

# 取得目前程式所在目錄（兼容 .py 與 exe）
if getattr(sys, 'frozen', False):  # exe 狀態
    base_dir = os.path.dirname(sys.executable)
else:  # 一般 Python 腳本
    base_dir = os.path.dirname(os.path.abspath(__file__))

script_path = os.path.join(base_dir, "QueueNumberTracker.py")

# 執行
subprocess.run(["python", script_path])
