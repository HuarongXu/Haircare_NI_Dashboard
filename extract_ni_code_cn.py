"""
从路径A复制 GC HC NI Code List 开头的Excel文件到路径B，
提取CN sheet数据（跳过空首行），生成 PS_Planning_NI_Code_CN.xlsx
"""

import os
import shutil
import glob
import pandas as pd
from app_config import NI_SOURCE_DIR, OUTPUT_DIR

# ============ 可配置路径 ============
PATH_A = NI_SOURCE_DIR
PATH_B = OUTPUT_DIR
# ====================================

OUTPUT_CN = "PS_Planning_NI_Code_CN.xlsx"
OUTPUT_HKTW = "PS_Planning_NI_Code_HKTW.xlsx"
OUTPUT_WIP = "PS_Planning_WIPlist.xlsx"
OUTPUT_CODELIST = "code list.xlsx"
FILE_PREFIX = "GC HC NI Code List"
SHEET_CN = "CN"
SHEET_HKTW = "HKTW"
SHEET_HYPERCARE = "NI Hyper care"
SHEETS_WIP = ["WIP-follow base", "WIP-unique no sticker"]
WIP_COLS = 11  # A到K列，共11列


def find_source_file(src_dir, prefix):
    """在src_dir中查找以prefix开头的Excel文件（.xls/.xlsx）"""
    patterns = [
        os.path.join(src_dir, f"{prefix}*.xlsx"),
        os.path.join(src_dir, f"{prefix}*.xls"),
    ]
    matches = []
    for pat in patterns:
        matches.extend(glob.glob(pat))
    # 排除临时文件 ~$
    matches = [f for f in matches if not os.path.basename(f).startswith("~$")]
    if not matches:
        raise FileNotFoundError(f"在 {src_dir} 中未找到以 '{prefix}' 开头的Excel文件")
    # 如果有多个，取最新修改的
    matches.sort(key=os.path.getmtime, reverse=True)
    return matches[0]


def main():
    print(f"源路径 (PATH_A): {PATH_A}")
    print(f"目标路径 (PATH_B): {PATH_B}")

    # 1. 查找源文件
    src_file = find_source_file(PATH_A, FILE_PREFIX)
    src_name = os.path.basename(src_file)
    print(f"找到源文件: {src_name}")

    # ===== 功能1: 提取CN sheet =====
    print("正在读取CN sheet...")
    df = pd.read_excel(src_file, sheet_name=SHEET_CN, header=1)
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    print(f"CN sheet 读取完成，共 {len(df)} 行, {len(df.columns)} 列")

    # 日期列只保留日期，去掉时间部分
    date_cols = ["SOS", "最早可生产时间 （备案MD ready）", "最早开卖日 （Earliest Start ship date）"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

    output_cn = os.path.join(PATH_B, OUTPUT_CN)
    df.to_excel(output_cn, index=False, sheet_name=SHEET_CN)
    print(f"已生成: {output_cn}")

    # ===== 功能1b: 提取HKTW sheet =====
    print("\n正在读取HKTW sheet...")
    df_hktw = pd.read_excel(src_file, sheet_name=SHEET_HKTW, header=0)
    df_hktw.dropna(how="all", inplace=True)
    df_hktw.reset_index(drop=True, inplace=True)
    print(f"HKTW sheet 读取完成，共 {len(df_hktw)} 行, {len(df_hktw.columns)} 列")

    # 日期列只保留日期，去掉时间部分
    for col in date_cols:
        if col in df_hktw.columns:
            df_hktw[col] = pd.to_datetime(df_hktw[col], errors="coerce").dt.date

    output_hktw = os.path.join(PATH_B, OUTPUT_HKTW)
    df_hktw.to_excel(output_hktw, index=False, sheet_name=SHEET_HKTW)
    print(f"已生成: {output_hktw}")

    # ===== 功能2: 合并WIP sheets =====
    print("\n正在读取WIP sheets...")
    wip_frames = []
    for sheet in SHEETS_WIP:
        df_wip = pd.read_excel(src_file, sheet_name=sheet)
        df_wip = df_wip.iloc[:, :WIP_COLS]  # 只取A-K列
        df_wip.dropna(how="all", inplace=True)
        print(f"  {sheet}: {len(df_wip)} 行")
        wip_frames.append(df_wip)

    df_wip_all = pd.concat(wip_frames, ignore_index=True)
    df_wip_all.drop_duplicates(inplace=True)
    df_wip_all.reset_index(drop=True, inplace=True)
    print(f"WIP合并完成，共 {len(df_wip_all)} 行, {len(df_wip_all.columns)} 列")

    output_wip = os.path.join(PATH_B, OUTPUT_WIP)
    df_wip_all.to_excel(output_wip, index=False)
    print(f"已生成: {output_wip}")

    # ===== 功能3: 生成 code list (SOS>=今天的CN/HKTW + NI Hyper care) =====
    print("\n正在生成 code list...")
    today = pd.Timestamp.today().normalize()
    codelist_frames = []

    # CN: SOS >= 今天
    df_cn_raw = pd.read_excel(src_file, sheet_name=SHEET_CN, header=1)
    df_cn_raw.dropna(how="all", inplace=True)
    df_cn_raw["SOS_dt"] = pd.to_datetime(df_cn_raw["SOS"], errors="coerce")
    df_cn_filtered = df_cn_raw[df_cn_raw["SOS_dt"] >= today][["PI FPC", "Description", "SOS_dt"]].copy()
    df_cn_filtered.rename(columns={"Description": "Descprtion", "SOS_dt": "SOS Date"}, inplace=True)
    df_cn_filtered["SOS Date"] = df_cn_filtered["SOS Date"].dt.date
    print(f"  CN (SOS>=今天): {len(df_cn_filtered)} 行")
    codelist_frames.append(df_cn_filtered)

    # HKTW: SOS >= 今天
    df_hk_raw = pd.read_excel(src_file, sheet_name=SHEET_HKTW, header=0)
    df_hk_raw.dropna(how="all", inplace=True)
    df_hk_raw["SOS_dt"] = pd.to_datetime(df_hk_raw["SOS"], errors="coerce")
    df_hk_filtered = df_hk_raw[df_hk_raw["SOS_dt"] >= today][["PI FPC", "Descprtion", "SOS_dt"]].copy()
    df_hk_filtered.rename(columns={"SOS_dt": "SOS Date"}, inplace=True)
    df_hk_filtered["SOS Date"] = df_hk_filtered["SOS Date"].dt.date
    print(f"  HKTW (SOS>=今天): {len(df_hk_filtered)} 行")
    codelist_frames.append(df_hk_filtered)

    # NI Hyper care: 全部，不筛SOS
    df_hc = pd.read_excel(src_file, sheet_name=SHEET_HYPERCARE, header=0)
    df_hc.dropna(how="all", inplace=True)
    desc_col = "Description" if "Description" in df_hc.columns else "Descprtion"
    sos_col_hc = "SOS" if "SOS" in df_hc.columns else None
    if sos_col_hc:
        df_hc_out = df_hc[["PI FPC", desc_col, sos_col_hc]].copy()
        df_hc_out["SOS Date"] = pd.to_datetime(df_hc_out[sos_col_hc], errors="coerce").dt.date
        df_hc_out.drop(columns=[sos_col_hc], inplace=True)
    else:
        df_hc_out = df_hc[["PI FPC", desc_col]].copy()
        df_hc_out["SOS Date"] = None
    df_hc_out.rename(columns={desc_col: "Descprtion"}, inplace=True)
    print(f"  NI Hyper care: {len(df_hc_out)} 行")
    codelist_frames.append(df_hc_out)

    df_codelist = pd.concat(codelist_frames, ignore_index=True)
    df_codelist.drop_duplicates(inplace=True)
    df_codelist.reset_index(drop=True, inplace=True)
    print(f"code list 合并完成，共 {len(df_codelist)} 行")

    output_cl = os.path.join(PATH_B, OUTPUT_CODELIST)
    df_codelist.to_excel(output_cl, index=False)
    print(f"已生成: {output_cl}")

    # 尝试复制源文件到目标路径（可选，失败不影响主流程）
    dst_file = os.path.join(PATH_B, src_name)
    try:
        shutil.copy2(src_file, dst_file)
        print(f"源文件已复制到: {dst_file}")
    except OSError as e:
        print(f"[警告] 源文件复制失败（OneDrive云端文件可能未同步）: {e}")
        print("主要输出文件已成功生成，不影响使用。")


if __name__ == "__main__":
    main()
