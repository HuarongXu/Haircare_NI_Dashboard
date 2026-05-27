import sys
import os

# 把项目根目录加入 sys.path，让 tests/ 里可以 import app_config / sap_vb03 等
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
