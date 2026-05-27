"""
单元测试：覆盖不依赖 SAP GUI 的纯数据逻辑函数
"""

import os
from datetime import date

import pandas as pd
import pytest


# ─────────────────────────────────────────────
# 辅助：构造模拟 VB03 原始报表 Excel
# ─────────────────────────────────────────────
def _make_mock_vb03(tmp_path):
    """
    创建与真实 VB03 导出结构相同的最小测试 Excel：
      Row 0: 空
      Row 1: TOB header（稀疏列：1,2,3,5,7,11,13）
      Row 2: 空
      Row 3-4: TOB 数据
      Row 5: TOC header（稀疏列：1,2,3,8,12,14）
      Row 6: 空
      Row 7-8: TOC 数据
    """
    ncols = 15

    def make_row(assignments):
        r = [None] * ncols
        for idx, val in assignments:
            r[idx] = val
        return r

    blank = [None] * ncols
    rows = [
        blank,
        make_row([(1,"Sale"),(2,"Di"),(3,"Plnt"),(5,"Cl"),(7,"Material"),(11,"Valid From"),(13,"  Valid To")]),
        blank,
        make_row([(1,"CN21"),(2,"01"),(3,"A668"),(5,"15"),(7,"83912147"),(11,"04/30/2026"),(13,"  12/31/9999")]),
        make_row([(1,"CN21"),(2,"01"),(3,"A668"),(5,"15"),(7,"83915055"),(11,"06/01/2026"),(13,"  12/31/9999")]),
        make_row([(1,"Sale"),(2,"Di"),(3,"Ship-to"),(8,"Material"),(12,"Valid From"),(14,"  Valid To")]),
        blank,
        make_row([(1,"CN21"),(2,"01"),(3,"2002761186"),(8,"83912147"),(12,"04/30/2026"),(14,"  12/31/9999")]),
        make_row([(1,"CN21"),(2,"01"),(3,"2002761186"),(8,"83915055"),(12,"06/01/2026"),(14,"  12/31/9999")]),
    ]

    df = pd.DataFrame(rows)
    f = tmp_path / "PS_Planning_VB03_mock.xlsx"
    df.to_excel(f, index=False, header=False)
    return str(f)


def _make_mock_codelist(tmp_path, rows=None):
    if rows is None:
        rows = {
            "PI FPC": [83912147, 83915055, 83912147, None],
            "Descprtion": ["A", "B", "A", None],
            "SOS Date": ["2026-06-01", "2026-12-15", "2026-06-01", None],
        }
    f = tmp_path / "code list.xlsx"
    pd.DataFrame(rows).to_excel(f, index=False)
    return str(f)


# ─────────────────────────────────────────────
# app_config 路径测试
# ─────────────────────────────────────────────
class TestAppConfig:
    def test_daily_dir_is_created(self):
        import app_config
        assert os.path.isdir(app_config.DAILY_DIR)

    def test_daily_dir_contains_today(self):
        import app_config
        today = date.today().strftime("%Y%m%d")
        assert today in app_config.DAILY_DIR

    def test_code_list_file_in_daily_dir(self):
        import app_config
        assert app_config.DAILY_DIR in app_config.CODE_LIST_FILE

    def test_sap_save_path_equals_daily_dir(self):
        import app_config
        assert app_config.SAP_SAVE_PATH == app_config.DAILY_DIR


# ─────────────────────────────────────────────
# load_code_list
# ─────────────────────────────────────────────
class TestLoadCodeList:
    def test_returns_unique_strings(self, tmp_path):
        from sap_vb03 import load_code_list
        f = _make_mock_codelist(tmp_path)
        codes = load_code_list(f)
        assert "83912147" in codes
        assert "83915055" in codes
        assert len(codes) == len(set(codes)), "存在重复 code"

    def test_drops_nan(self, tmp_path):
        from sap_vb03 import load_code_list
        f = _make_mock_codelist(tmp_path)
        codes = load_code_list(f)
        assert all(c != "nan" for c in codes)

    def test_all_strings(self, tmp_path):
        from sap_vb03 import load_code_list
        f = _make_mock_codelist(tmp_path)
        codes = load_code_list(f)
        assert all(isinstance(c, str) for c in codes)


# ─────────────────────────────────────────────
# get_max_sos_date
# ─────────────────────────────────────────────
class TestGetMaxSosDate:
    def test_returns_max_date(self, tmp_path):
        from sap_vb03 import get_max_sos_date
        f = _make_mock_codelist(tmp_path)
        result = get_max_sos_date(f)
        assert result == "12/15/2026"

    def test_format_mm_dd_yyyy(self, tmp_path):
        from sap_vb03 import get_max_sos_date
        f = _make_mock_codelist(tmp_path)
        result = get_max_sos_date(f)
        parts = result.split("/")
        assert len(parts) == 3
        assert len(parts[2]) == 4, "年份应为4位"

    def test_raises_when_no_valid_date(self, tmp_path):
        from sap_vb03 import get_max_sos_date
        f = _make_mock_codelist(tmp_path, rows={
            "PI FPC": [1], "Descprtion": ["X"], "SOS Date": [None]
        })
        with pytest.raises(ValueError):
            get_max_sos_date(f)


# ─────────────────────────────────────────────
# extract_vb03_subtables
# ─────────────────────────────────────────────
class TestExtractVb03Subtables:
    def test_creates_tob_and_toc_files(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        assert os.path.exists(tmp_path / "PS_Planning_VB03_tob20260527.xlsx")
        assert os.path.exists(tmp_path / "PS_Planning_VB03_toc20260527.xlsx")

    def test_tob_columns(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_tob20260527.xlsx")
        assert list(df.columns) == ["Sale", "Di", "Plnt", "Cl", "Material", "Valid From", "Valid To"]

    def test_toc_columns(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_toc20260527.xlsx")
        assert list(df.columns) == ["Sale", "Di", "Ship-to", "Material", "Valid From", "Valid To"]

    def test_tob_no_empty_rows(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_tob20260527.xlsx")
        assert df.isnull().all(axis=1).sum() == 0, "TOB 存在空行"

    def test_toc_no_empty_rows(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_toc20260527.xlsx")
        assert df.isnull().all(axis=1).sum() == 0, "TOC 存在空行"

    def test_valid_to_stripped(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_tob20260527.xlsx")
        assert df["Valid To"].iloc[0] == "12/31/9999", "Valid To 应已去掉前导空格"

    def test_tob_row_count(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_tob20260527.xlsx")
        assert len(df) == 2

    def test_toc_row_count(self, tmp_path):
        from sap_vb03 import extract_vb03_subtables
        src = _make_mock_vb03(tmp_path)
        extract_vb03_subtables(src, str(tmp_path), "20260527")
        df = pd.read_excel(tmp_path / "PS_Planning_VB03_toc20260527.xlsx")
        assert len(df) == 2
