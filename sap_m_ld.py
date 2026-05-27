"""
SAP GUI Automation: Run M_LD Report
- 从 code list.xlsx 读取所有 PI FPC codes
- 通过 SAP GUI Scripting 执行 M_LD T-code
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
SAP_SAVE_FILENAME = f"PS_Planning_M_LD_{date.today().strftime('%Y%m%d')}.xlsx"
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


def clean_excel(filepath):
    """清理Excel文件中的空行和空列，自动检测真正的header行"""
    df = pd.read_excel(filepath, header=None)

    # 找到真正的header行：第一行全是"Unnamed"开头则跳过
    header_row = 0
    for i in range(min(5, len(df))):
        row_vals = df.iloc[i].astype(str).tolist()
        if all(v.startswith("Unnamed") or v.strip() == "" for v in row_vals):
            header_row = i + 1
        else:
            break

    # 重新读取，以正确的行作为header
    df = pd.read_excel(filepath, header=header_row)

    # 清理列名中多余的空格
    df.columns = [str(c).strip() for c in df.columns]

    df.dropna(how="all", inplace=True)
    df.dropna(axis=1, how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.to_excel(filepath, index=False)
    print(f"  清理后: {len(df)} 行, {len(df.columns)} 列")


def run_m_ld_report():
    print("=" * 50)
    print("SAP M_LD Report Automation")
    print("=" * 50)

    # 1. 读取code list
    print(f"\n[1/5] 读取code list: {CODE_LIST_FILE}")
    codes = load_code_list(CODE_LIST_FILE)
    print(f"  共 {len(codes)} 个codes")

    # 2. 连接SAP
    print("\n[2/5] 连接SAP GUI...")
    session = get_sap_session()
    print("  SAP GUI 连接成功")

    # 3. 进入T-code M_LD 并填写选择画面
    print("\n[3/5] 进入 M_LD 选择画面...")
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nM_LD"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/ctxtRV14A-KONLI").text = "xy"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1)

    # 设置选择条件
    session.findById("wnd[0]/usr/ctxtL_1-LOW").text = "*"
    session.findById("wnd[0]/usr/ctxtKSCHL-LOW").text = "zpb0"

    # 4. 输入codes: 打开多值输入 → 剪贴板粘贴
    print("\n[4/5] 输入codes到选择画面...")
    session.findById("wnd[0]/usr/btn%_L_2_%_APP_%-VALU_PUSH").press()

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

    # 清理空行空列
    if os.path.exists(output_path):
        print("\n正在清理空行空列...")
        clean_excel(output_path)
    else:
        print(f"[警告] 文件未找到: {output_path}")

    print("\n" + "=" * 50)
    print("M_LD Report 完成!")
    print("=" * 50)


if __name__ == "__main__":
    run_m_ld_report()
