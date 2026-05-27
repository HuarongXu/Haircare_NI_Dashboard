"""
总入口：按顺序执行 3 个 SAP 报表脚本。
执行顺序: M_LD -> MM60 -> VB03
"""

import traceback

from sap_m_ld import run_m_ld_report
from sap_mm60 import run_mm60_report
from sap_vb03 import run_vb03_report


def _run_step(step_name, func):
    print("\n" + "=" * 60)
    print(f"开始执行: {step_name}")
    print("=" * 60)
    try:
        func()
        print(f"[完成] {step_name}")
        return True
    except Exception as exc:
        print(f"[失败] {step_name}: {exc}")
        traceback.print_exc()
        return False


def main():
    print("=" * 60)
    print("SAP 三报表一键执行")
    print("=" * 60)

    results = []
    results.append(("M_LD", _run_step("M_LD", run_m_ld_report)))
    results.append(("MM60", _run_step("MM60", run_mm60_report)))
    results.append(("VB03", _run_step("VB03", run_vb03_report)))

    print("\n" + "=" * 60)
    print("执行汇总")
    print("=" * 60)
    for name, ok in results:
        status = "SUCCESS" if ok else "FAILED"
        print(f"{name}: {status}")

    failed = [name for name, ok in results if not ok]
    if failed:
        print("\n存在失败步骤: " + ", ".join(failed))
    else:
        print("\n全部步骤执行成功。")


if __name__ == "__main__":
    main()
