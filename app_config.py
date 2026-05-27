"""
统一配置文件：集中管理所有可修改路径与文件名。
修改这里即可影响所有脚本。
"""

import os
from datetime import date


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 源数据目录（OneDrive）
NI_SOURCE_DIR = r"C:\Users\xu.hr\OneDrive - Procter and Gamble\General\HC IOP\1. NI Code List"

# 根输出目录
OUTPUT_DIR = BASE_DIR

# 每日数据子目录（格式：YYYYMMDD），自动创建
DATE_STR = date.today().strftime("%Y%m%d")
DAILY_DIR = os.path.join(OUTPUT_DIR, DATE_STR)
os.makedirs(DAILY_DIR, exist_ok=True)

# 公共输入/输出文件（当天目录下）
CODE_LIST_FILENAME = "code list.xlsx"
CODE_LIST_FILE = os.path.join(DAILY_DIR, CODE_LIST_FILENAME)

# SAP 导出路径（当天目录）
SAP_SAVE_PATH = DAILY_DIR
