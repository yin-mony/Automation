# -*- coding: utf-8 -*-
# =============================================================================
# 模块说明：按「店铺 / 门店」列拆分 Excel（.xlsx）的纯逻辑实现。
#           不包含任何 Qt 界面代码，供 Tabellen_teilen.py 与命令行共用。
# =============================================================================

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

# -----------------------------------------------------------------------------
# 常量：常见「店铺」表头候选名；guess_store_col 按顺序取第一个命中
# -----------------------------------------------------------------------------

STORE_COLUMN_CANDIDATES: tuple[str, ...] = (
    "店铺",
    "门店",
    "店名",
    "店铺名称",
    "门店名称",
    "网店",
    "店铺名",
    "Store",
    "store",
    "Shop",
    "shop",
    "门店编码",
    "店铺编码",
)


# -----------------------------------------------------------------------------
# 作用：把单元格里的店铺名变成可安全用作 Windows 文件名的字符串（去非法字符、控长度）。
# -----------------------------------------------------------------------------
def sanitize_name(cell_value: Any, max_len: int = 120) -> str:
    text = str(cell_value).strip()
    if not text or text.lower() == "nan":
        return "未填写店铺"
    for ch in r'\/:*?"<>|':
        text = text.replace(ch, "_")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len] if len(text) > max_len else text


# -----------------------------------------------------------------------------
# 作用：只读 Excel 第一行表头，不加载数据；用于界面快速列出可选列名。
# -----------------------------------------------------------------------------
def read_headers(xlsx_path: Path) -> list[Any]:
    header_only = pd.read_excel(xlsx_path, engine="openpyxl", nrows=0)
    return list(header_only.columns)


# -----------------------------------------------------------------------------
# 作用：在表头列表里按预设候选名自动猜「哪一列是店铺列」；猜不到则返回 None。
# -----------------------------------------------------------------------------
def guess_store_col(column_names: list[Any]) -> Any | None:
    by_stripped = {str(c).strip(): c for c in column_names}
    by_lower = {str(c).strip().lower(): c for c in column_names}
    for candidate in STORE_COLUMN_CANDIDATES:
        key = candidate.strip()
        if key in by_stripped:
            return by_stripped[key]
        lower = key.lower()
        if lower in by_lower:
            return by_lower[lower]
    return None


# -----------------------------------------------------------------------------
# 作用：在用户选的列名与 DataFrame 真实列名之间对齐（忽略首尾空格）；对不上则抛错。
# -----------------------------------------------------------------------------
def match_col(df: pd.DataFrame, user_chosen_column: Any) -> Any:
    for col in df.columns:
        if str(col).strip() == str(user_chosen_column).strip():
            return col
    raise ValueError("所选「店铺列」在当前表中不存在，请重新加载文件后选择列名。")


# -----------------------------------------------------------------------------
# 作用：在输出目录里为「基础文件名」分配一个不冲突的 .xlsx 路径（必要时加 _2、_3）。
# -----------------------------------------------------------------------------
def unique_path(out_dir: Path, base: str) -> Path:
    first = out_dir / f"{base}.xlsx"
    if not first.exists():
        return first
    n = 2
    while True:
        candidate = out_dir / f"{base}_{n}.xlsx"
        if not candidate.exists():
            return candidate
        n += 1


# -----------------------------------------------------------------------------
# 作用：读取整张表，按指定列分组，每组写一个独立 xlsx；返回写出的文件个数。
# -----------------------------------------------------------------------------
def split_by_store(xlsx_path: Path, store_col: Any, out_dir: Path) -> int:
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    if df.empty:
        raise ValueError("表格为空，无法拆分。")

    key = match_col(df, store_col)
    out_dir.mkdir(parents=True, exist_ok=True)

    n = 0
    for val, part in df.groupby(key, dropna=False):
        label = sanitize_name(val if pd.notna(val) else "未填写店铺")
        path = unique_path(out_dir, label)
        part.to_excel(path, index=False, engine="openpyxl")
        n += 1
    return n
