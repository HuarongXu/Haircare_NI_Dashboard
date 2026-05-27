"""
统一配置文件：集中管理所有可修改路径与文件名。
修改这里即可影响所有脚本。
"""

import os


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 源数据目录（OneDrive）
NI_SOURCE_DIR = r"C:\Users\xu.hr\OneDrive - Procter and Gamble\General\HC IOP\1. NI Code List"

# 输出目录（默认当前项目目录）
OUTPUT_DIR = BASE_DIR

# 公共输入/输出文件
CODE_LIST_FILENAME = "code list.xlsx"
CODE_LIST_FILE = os.path.join(OUTPUT_DIR, CODE_LIST_FILENAME)

# SAP 导出路径（默认与 OUTPUT_DIR 一致）
SAP_SAVE_PATH = OUTPUT_DIR
