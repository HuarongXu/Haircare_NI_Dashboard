"""
SAP GUI Automation: Run VB03 Report
- 从 code list.xlsx 读取所有 PI FPC codes
- 通过 SAP GUI Scripting 执行 VB03 T-code (Condition: z002)
- 导出报表为 Excel 并清理空行空列
"""

import os
import time
from datetime import date
import pandas as pd
import win32com.client
import win32clipboard
from app_config import SAP_SAVE_PATH, CODE_LIST_FILE

# ============ 可配置 ============
SAP_SAVE_FILENAME = f"PS_Planning_VB03_{date.today().strftime('%Y%m%d')}.xlsx"
# ================================


def get_sap_session():
    """连接已打开的SAP GUI会话"""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def set_clipboard(text):
    """将文本复制到Windows剪贴板"""
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()


def load_code_list(filepath):
    """从 code list.xlsx 读取 PI FPC 列，返回去重的字符串列表"""
    df = pd.read_excel(filepath)
    codes = df["PI FPC"].dropna().astype(int).astype(str).unique().tolist()
    return codes


def get_max_sos_date(filepath):
    """从 code list.xlsx 读取最大 SOS Date，返回 SAP 格式字符串 MM/DD/YYYY"""
    df = pd.read_excel(filepath)
    max_date = pd.to_datetime(df["SOS Date"], errors="coerce").max()
    if pd.isna(max_date):
        raise ValueError("code list 中没有有效的 SOS Date")
    return max_date.strftime("%m/%d/%Y")


def extract_vb03_subtables(src_filepath, save_path, date_str):
    """
    从原始 VB03 报表中提取两个子表并另存：
      - tob: Sale/Di/Plnt/Cl/Material/Valid From/Valid To
      - toc: Sale/Di/Ship-to/Material/Valid From/Valid To
    """
    df_raw = pd.read_excel(src_filepath, header=None)

    # 动态找所有 header 行（含 Sale 且含 Valid From/Valid To）
    header_rows = []
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.tolist()]
        if "Sale" in vals and any("Valid" in v for v in vals):
            header_rows.append(i)

    if len(header_rows) < 2:
        print(f"  [警告] 只找到 {len(header_rows)} 个表头，期望至少 2 个，跳过子表提取")
        return

    # 定位两个目标表的 header 行
    tob_hdr, toc_hdr = header_rows[0], header_rows[1]
    next_after_toc = header_rows[2] if len(header_rows) > 2 else len(df_raw)

    def _extract(hdr_row, end_row, col_map):
        """col_map: {output_name: col_index}"""
        # 取 header 行下一行到 end_row 之前的数据行
        data = df_raw.iloc[hdr_row + 1 : end_row].copy()
        # 只保留目标列
        col_indices = list(col_map.values())
        col_names = list(col_map.keys())
        data = data.iloc[:, col_indices]
        data.columns = col_names
        # 去空行，strip 字符串列
        data = data.dropna(how="all").reset_index(drop=True)
        for c in data.columns:
            if data[c].dtype == object:
                data[c] = data[c].astype(str).str.strip()
                data[c] = data[c].replace("nan", pd.NA)
        data = data.dropna(how="all").reset_index(drop=True)
        return data

    # TOB: Sale/Di/Plnt/Cl/Material/Valid From/Valid To
    tob_cols = {"Sale": 1, "Di": 2, "Plnt": 3, "Cl": 5, "Material": 7,
                "Valid From": 11, "Valid To": 13}
    df_tob = _extract(tob_hdr, toc_hdr, tob_cols)
    tob_path = os.path.join(save_path, f"PS_Planning_VB03_tob{date_str}.xlsx")
    df_tob.to_excel(tob_path, index=False)
    print(f"  TOB 子表已保存: {tob_path}  ({len(df_tob)} 行)")

    # TOC: Sale/Di/Ship-to/Material/Valid From/Valid To
    toc_cols = {"Sale": 1, "Di": 2, "Ship-to": 3, "Material": 8,
                "Valid From": 12, "Valid To": 14}
    df_toc = _extract(toc_hdr, next_after_toc, toc_cols)
    toc_path = os.path.join(save_path, f"PS_Planning_VB03_toc{date_str}.xlsx")
    df_toc.to_excel(toc_path, index=False)
    print(f"  TOC 子表已保存: {toc_path}  ({len(df_toc)} 行)")


def clean_excel(filepath):
    """清理Excel文件中的空行和空列，自动检测真正的header行"""
    df_raw = pd.read_excel(filepath, header=None)

    # 找到真正的header行：跳过全是"Unnamed"或空值的行
    header_row = 0
    for i in range(min(5, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).tolist()
        if all("unnamed" in v.lower() or v.strip() == "" or v.strip() == "nan" for v in row_vals):
            header_row = i + 1
        else:
            break

    # 用header_row行作为列名，跳过前面的垃圾行
    df = pd.read_excel(filepath, header=None, skiprows=header_row)
    df.columns = df.iloc[0].astype(str).str.strip().tolist()
    df = df.iloc[1:].reset_index(drop=True)

    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.to_excel(filepath, index=False)
    print(f"  清理后: {len(df)} 行, {len(df.columns)} 列")


def run_vb03_report():
    print("=" * 50)
    print("SAP VB03 Report Automation")
    print("=" * 50)

    # 1. 读取code list
    print(f"\n[1/5] 读取code list: {CODE_LIST_FILE}")
    codes = load_code_list(CODE_LIST_FILE)
    print(f"  共 {len(codes)} 个codes")
    max_sos = get_max_sos_date(CODE_LIST_FILE)
    print(f"  最大 SOS Date: {max_sos}")

    # 2. 连接SAP
    print("\n[2/5] 连接SAP GUI...")
    session = get_sap_session()
    print("  SAP GUI 连接成功")

    # 3. 进入T-code VB03
    print("\n[3/5] 进入 VB03 选择画面...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVB03"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    # 设置 Condition type = z002
    session.findById("wnd[0]/usr/ctxtG000-KSCHL").text = "z002"
    # 点击 Selection 变式按钮进入选择画面
    session.findById("wnd[0]/tbar[1]/btn[16]").press()
    time.sleep(1)

    # 设置 Customer = *
    session.findById("wnd[0]/usr/ctxtF005-LOW").text = "*"

    # 设置日期为 code list 中最大 SOS Date
    session.findById("wnd[0]/usr/ctxtSEL_DATE").text = max_sos
    print(f"  已设置 SEL_DATE = {max_sos}")

    # 4. 输入codes: 打开多值输入 → 剪贴板粘贴
    print("\n[4/5] 输入codes到选择画面...")
    session.findById("wnd[0]/usr/btn%_F004_%_APP_%-VALU_PUSH").press()

    codes_text = "\r\n".join(codes)
    set_clipboard(codes_text)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()  # Upload from Clipboard
    print(f"  已粘贴 {len(codes)} 个codes")

    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # 确认选择

    # 执行报表
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    print("  报表执行中...")
    time.sleep(3)

    # 5. 导出Excel
    print("\n[5/5] 导出报表...")
    output_path = os.path.join(SAP_SAVE_PATH, SAP_SAVE_FILENAME)
    # 先删除同名文件，避免SAP弹出Replace/Generate选择框
    if os.path.exists(output_path):
        os.remove(output_path)
        print(f"  已删除旧文件: {SAP_SAVE_FILENAME}")

    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()

    # 选择 Spreadsheet 格式
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # 设置保存路径和文件名
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = SAP_SAVE_PATH
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = SAP_SAVE_FILENAME
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    print(f"  报表已导出到: {output_path}")

    # 等待文件保存完成
    time.sleep(3)

    if os.path.exists(output_path):
        print(f"  文件已保存: {output_path}")
        # 提取两个子表
        print("\n正在提取子表 tob / toc...")
        extract_vb03_subtables(output_path, SAP_SAVE_PATH, date.today().strftime('%Y%m%d'))
    else:
        print(f"[警告] 文件未找到: {output_path}")

    print("\n" + "=" * 50)
    print("VB03 Report 完成!")
    print("=" * 50)


if __name__ == "__main__":
    run_vb03_report()
