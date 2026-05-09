# -*- coding: utf-8 -*-
# =============================================================================
# 模块说明：按「店铺 / 门店」列拆分 Excel（.xlsx）的纯逻辑实现。
#           不包含任何 Qt 界面代码，供 Tabellen_teilen.py 与命令行共用。
# =============================================================================

import re
from pathlib import Path

import pandas as pd

# -----------------------------------------------------------------------------
# 常量：常见「店铺」表头候选名；guess_store_col 按顺序取第一个命中
# -----------------------------------------------------------------------------

STORE_COLUMN_CANDIDATES = (
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

# 常见「类型」表头候选；guess_type_col 按顺序取第一个命中
TYPE_COLUMN_CANDIDATES = (
    "类型",
    "类别",
    "品类",
    "分类",
    "Type",
    "type",
    "Category",
    "category",
)

# 常见「时间」表头候选；guess_time_col 按顺序取第一个命中
TIME_COLUMN_CANDIDATES = (
    "时间",
    "日期",
    "日期时间",
    "业务日期",
    "下单时间",
    "Date",
    "date",
    "Time",
    "time",
)


# -----------------------------------------------------------------------------
# 传参：cell_value — 单元格原始值（任意类型，会转成字符串）
#       max_len — 文件名主体最大长度，默认 120
# 返回：str，可安全用于 Windows 文件名的店铺名字符串
# 作用：去非法字符、合并空白、控长度；空或 nan 则返回「未填写店铺」
# -----------------------------------------------------------------------------
def sanitize_name(cell_value, max_len=120):
    text = str(cell_value).strip()
    if not text or text.lower() == "nan":
        return "未填写店铺"
    for ch in r'\/:*?"<>|':
        text = text.replace(ch, "_")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len] if len(text) > max_len else text


# -----------------------------------------------------------------------------
# 传参：xlsx_path — Excel 文件路径（pathlib.Path）
# 返回：list，表头列名列表（与 pandas 读入后的 columns 一致）
# 作用：只读第一行表头，不加载数据行
# -----------------------------------------------------------------------------
def read_headers(xlsx_path):
    header_only = pd.read_excel(xlsx_path, engine="openpyxl", nrows=0)
    return list(header_only.columns)


# -----------------------------------------------------------------------------
# 传参：column_names — 表头列名列表（通常来自 read_headers 的返回值）
# 返回：在 column_names 里命中的那一列的「原始列名」；猜不到则返回 None
# 作用：按 STORE_COLUMN_CANDIDATES 顺序自动猜店铺列
# -----------------------------------------------------------------------------
def guess_store_col(column_names):
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


def _guess_col_from_candidates(column_names, candidates):
    by_stripped = {str(c).strip(): c for c in column_names}
    by_lower = {str(c).strip().lower(): c for c in column_names}
    for candidate in candidates:
        key = candidate.strip()
        if key in by_stripped:
            return by_stripped[key]
        lower = key.lower()
        if lower in by_lower:
            return by_lower[lower]
    return None


def guess_type_col(column_names):
    return _guess_col_from_candidates(column_names, TYPE_COLUMN_CANDIDATES)


def guess_time_col(column_names):
    return _guess_col_from_candidates(column_names, TIME_COLUMN_CANDIDATES)


# -----------------------------------------------------------------------------
# 传参：df — 已读入的整张表（pandas DataFrame）
#       user_chosen_column — 用户选择的列名（与表头比对时忽略首尾空格）
#       role — 报错信息中的列角色说明（默认「店铺列」）
# 返回：df.columns 中与 user_chosen_column 对应的那一项（真实列名）
# 作用：对齐界面选的列名与表内实际列名；对不上则抛出 ValueError
# -----------------------------------------------------------------------------
def match_col(df, user_chosen_column, role="店铺列"):
    for col in df.columns:
        if str(col).strip() == str(user_chosen_column).strip():
            return col
    raise ValueError(f"所选「{role}」在当前表中不存在，请重新加载文件后选择列名。")


# -----------------------------------------------------------------------------
# 传参：out_dir — 输出目录（pathlib.Path）
#       base — 不含扩展名的文件名主体（已由 sanitize_name 处理过）
# 返回：pathlib.Path，一个不冲突的 .xlsx 完整路径（必要时为 base_2.xlsx、base_3.xlsx…）
# 作用：避免覆盖已存在的同名文件
# -----------------------------------------------------------------------------
def unique_path(out_dir, base):
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
# 传参：xlsx_path — 待拆分的 Excel 路径（pathlib.Path）
#       store_col — 作为拆分依据的列名（用户选择或 guess_store_col 结果）
#       out_dir — 拆分结果输出目录（pathlib.Path）
#       name_type — 可选；勾选「类型」且已填写时参与文件名
#       name_time — 可选；勾选「时间」且已填写时参与文件名
# 规则：仅店铺名；
#       仅类型 →「类型-店铺名」；仅时间 →「时间-店铺名」；二者皆有 →「类型-时间-店铺名」
# 返回：int，实际写出的 xlsx 文件个数（按店铺取值分组，一组一个文件）
# 作用：读全表 → 按 store_col 分组 → 每组写一个 xlsx
# -----------------------------------------------------------------------------
def split_by_store(xlsx_path, store_col, out_dir, name_type=None, name_time=None):
    df = pd.read_excel(xlsx_path, engine="openpyxl")
    if df.empty:
        raise ValueError("表格为空，无法拆分。")

    key = match_col(df, store_col)
    t_raw = (name_type or "").strip()
    tm_raw = (name_time or "").strip()
    use_t = bool(t_raw)
    use_tm = bool(tm_raw)

    out_dir.mkdir(parents=True, exist_ok=True)

    n = 0
    for val, part in df.groupby(key, dropna=False):
        store_label = sanitize_name(val if pd.notna(val) else "未填写店铺")
        if use_t and use_tm:
            label = (
                f"{sanitize_name(t_raw, max_len=80)}-"
                f"{sanitize_name(tm_raw, max_len=80)}-{store_label}"
            )
        elif use_t:
            label = f"{sanitize_name(t_raw, max_len=80)}-{store_label}"
        elif use_tm:
            label = f"{sanitize_name(tm_raw, max_len=80)}-{store_label}"
        else:
            label = store_label
        path = unique_path(out_dir, label)
        part.to_excel(path, index=False, engine="openpyxl")
        n += 1
    return n
