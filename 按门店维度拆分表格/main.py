# -*- coding: utf-8 -*-
"""
按「店铺 / 门店」列拆分 Excel（.xlsx）。
逻辑源自 excel_store_split.py，由 Excel_file 类封装。
"""

import re
from pathlib import Path

import pandas as pd

# 常见表头候选名：kind 传 "store" / "type" / "time"
COLUMN_CANDIDATES = {
    "store": (
        "店铺", "门店", "店名", "店铺名称", "门店名称", "网店", "店铺名",
        "Store", "store", "Shop", "shop", "门店编码", "店铺编码",
    ),
    "type": (
        "类型", "类别", "品类", "分类", "Type", "type", "Category", "category",
    ),
    "time": (
        "时间", "日期", "日期时间", "业务日期", "下单时间",
        "Date", "date", "Time", "time",
    ),
}


class Excel_file:
    """
    用法：
        excel = Excel_file(r"C:\data\表.xlsx")
        col = excel.guess_col() or "店铺"
        n = excel.split_by_store(col, r"C:\output")
    """

    def __init__(self, path):
        # path：xlsx 文件路径（str 或 Path）
        self.path = Path(path)
        if not self.path.is_file():
            raise FileNotFoundError(f"Excel 文件不存在：{self.path}")
        if self.path.suffix.lower() != ".xlsx":
            raise ValueError(f"仅支持 .xlsx 文件：{self.path}")
        self._headers = None
        self._df = None

    def read_headers(self):
        # 返回表头列名列表；只读第一行，结果会缓存
        if self._headers is None:
            df = pd.read_excel(self.path, engine="openpyxl", nrows=0)
            self._headers = list(df.columns)
        return self._headers

    def load(self):
        # 返回整张表的 DataFrame；结果会缓存；空表会抛 ValueError
        if self._df is None:
            self._df = pd.read_excel(self.path, engine="openpyxl")
            if self._df.empty:
                raise ValueError("表格为空，无法拆分。")
        return self._df

    def guess_col(self, kind="store"):
        # kind：列类型 store / type / time；在表头里按候选名匹配，找到返回列名，找不到返回 None
        candidates = COLUMN_CANDIDATES.get(kind, COLUMN_CANDIDATES["store"])
        names = self.read_headers()
        by_strip = {str(c).strip(): c for c in names}
        by_lower = {str(c).strip().lower(): c for c in names}
        for name in candidates:
            key = name.strip()
            if key in by_strip:
                return by_strip[key]
            if key.lower() in by_lower:
                return by_lower[key.lower()]
        return None

    def match_col(self, col_name, role="店铺列"):
        # col_name：用户选的列名；返回表里对应的真实列名；对不上抛 ValueError
        for col in self.load().columns:
            if str(col).strip() == str(col_name).strip():
                return col
        raise ValueError(f"所选「{role}」在当前表中不存在，请重新加载文件后选择列名。")

    def split_by_store(self, store_col, out_dir, name_type=None, name_time=None):
        # store_col：拆分依据的列名
        # out_dir：输出目录
        # name_type / name_time：可选，参与输出文件名前缀
        # 返回：写出的 xlsx 文件个数
        df = self.load()
        key = self.match_col(store_col)
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        t = (name_type or "").strip()
        tm = (name_time or "").strip()

        count = 0
        for val, part in df.groupby(key, dropna=False):
            store = self._safe_name(val if pd.notna(val) else "未填写店铺")
            if t and tm:
                label = f"{self._safe_name(t, 80)}-{self._safe_name(tm, 80)}-{store}"
            elif t:
                label = f"{self._safe_name(t, 80)}-{store}"
            elif tm:
                label = f"{self._safe_name(tm, 80)}-{store}"
            else:
                label = store

            path = self._next_path(out_dir, label)
            part.to_excel(path, index=False, engine="openpyxl")
            count += 1
        return count

    def _safe_name(self, value, max_len=120):
        # 把单元格内容转成 Windows 安全文件名
        text = str(value).strip()
        if not text or text.lower() == "nan":
            return "未填写店铺"
        for ch in r'\/:*?"<>|':
            text = text.replace(ch, "_")
        text = re.sub(r"\s+", " ", text).strip()
        return text[:max_len] if len(text) > max_len else text

    def _next_path(self, out_dir, base):
        # 在 out_dir 下生成不重复的 .xlsx 路径（同名则加 _2、_3…）
        first = out_dir / f"{base}.xlsx"
        if not first.exists():
            return first
        n = 2
        while True:
            p = out_dir / f"{base}_{n}.xlsx"
            if not p.exists():
                return p
            n += 1


if __name__ == "__main__":
    import sys

    config = {
        "excel_path": r"C:\RPA流程\按门店维度拆分表格\flie.xlsx",   # Excel 文件
        "out_dir": r"C:\RPA流程\按门店维度拆分表格\flie\已完成",      # 输出目录
    }

    # 命令行可覆盖：python main.py <xlsx路径> <输出目录>
    if len(sys.argv) >= 3:
        config["excel_path"] = sys.argv[1]
        config["out_dir"] = sys.argv[2]

    if not config["excel_path"].strip():
        raise ValueError("请在 config 中填写 excel_path（Excel 文件路径）。")
    if not config["out_dir"].strip():
        raise ValueError("请在 config 中填写 out_dir（输出目录）。")

    excel = Excel_file(config["excel_path"])
    store_col = excel.guess_col()
    if not store_col:
        raise ValueError("无法从表头自动识别店铺列，请检查 Excel 表头。")

    n = excel.split_by_store(store_col, config["out_dir"])
    print(f"已拆分 {n} 个文件 -> {config['out_dir']}")
