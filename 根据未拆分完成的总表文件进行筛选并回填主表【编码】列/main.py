# -*- coding: utf-8 -*-
"""
主表与副表按订单号匹配，将 asin 回填到主表编码列。
逻辑源自 Filter_add.py，由 Excel_file 类封装。
"""

from pathlib import Path

import pandas as pd

# 默认列名
TOTAL_COL = "描述"
FIND_COL = "myp_order_id"
ASIN_COL = "asin"
TARGET_COL = "编码（必填）"
ALT_TARGET_COL = "编码(必填)"


class Excel_file:
    """
    用法：
        excel = Excel_file(r"C:\data\主表.xlsx")
        excel.run([r"C:\data\副表1.xlsx", r"C:\data\副表2.xlsx"])
    """

    def __init__(self, path):
        # path：主表 xlsx 路径（str 或 Path）
        self.path = Path(path)
        if not self.path.is_file():
            raise FileNotFoundError(f"主表文件不存在：{self.path}")
        self._df = None

    def load(self):
        # 返回主表 DataFrame；结果会缓存
        if self._df is None:
            df = pd.read_excel(self.path)
            df.columns = df.columns.str.strip()
            if df.empty:
                raise ValueError("主表为空，无法处理。")
            self._df = df
        return self._df

    def run(self, sub_paths, output_path=None):
        # sub_paths：副表路径，可传单个路径或路径列表
        # output_path：写回路径，留空则覆盖主表 self.path
        # 返回：dict，含回填后的 total_df、匹配统计等
        paths = self._to_sub_list(sub_paths)
        if not paths:
            raise ValueError("至少需要一个副表文件路径。")

        total_work = self.load().copy()
        per_sub = []
        sub_results = []
        match_dfs = []

        for i, sub_path in enumerate(paths):
            sub_df = pd.read_excel(sub_path)
            sub_df.columns = sub_df.columns.str.strip()
            self._check_cols(total_work, sub_df)

            matched = self._match_sub(total_work, sub_df)
            total_work, target_col = self._fill_codes(
                total_work,
                matched["total_key"],
                matched["asin_map"],
                merge_existing=(i > 0),
            )
            per_sub.append({**matched, "sub_path": str(sub_path)})
            sub_results.append(matched["sub_result"])
            match_dfs.append(matched["match_df"])

        out = Path(output_path) if output_path else self.path
        out.parent.mkdir(parents=True, exist_ok=True)
        try:
            total_work.to_excel(out, index=False)
        except PermissionError as exc:
            raise PermissionError(f"无法写入文件：{out}。请关闭 Excel 后重试。") from exc

        if len(sub_results) > 1:
            all_sub = pd.concat(sub_results, ignore_index=True)
            all_match = pd.concat(match_dfs, ignore_index=True)
        else:
            all_sub = sub_results[0]
            all_match = match_dfs[0]

        return {
            "total_df_filled": total_work,
            "target_col_used": target_col,
            "sub_result": all_sub,
            "match_df": all_match,
            "saved_path": str(out),
            "per_sub": per_sub,
        }

    def _to_sub_list(self, sub_paths):
        # 把单个路径或列表统一成 Path 列表
        if sub_paths is None:
            return []
        if isinstance(sub_paths, (list, tuple)):
            items = sub_paths
        else:
            items = [sub_paths]
        result = []
        for p in items:
            if not str(p).strip():
                continue
            path = Path(p)
            if not path.is_file():
                raise FileNotFoundError(f"副表文件不存在：{path}")
            result.append(path)
        return result

    def _check_cols(self, total_df, sub_df):
        # 校验主表、副表是否有所需列，缺列抛 KeyError
        if TOTAL_COL not in total_df.columns:
            raise KeyError(f"主表中不存在列：{TOTAL_COL}")
        if FIND_COL not in sub_df.columns:
            raise KeyError(f"副表中不存在列：{FIND_COL}")
        if ASIN_COL not in sub_df.columns:
            raise KeyError(f"副表中不存在列：{ASIN_COL}")

    def _match_sub(self, total_df, sub_df):
        # 主表描述 与 副表 myp_order_id 完全匹配，聚合 asin
        total_key = total_df[TOTAL_COL].astype("string").str.strip()
        find_key = sub_df[FIND_COL].astype("string").str.strip()

        sub_result = sub_df.copy()
        sub_result[FIND_COL] = find_key
        sub_result[ASIN_COL] = sub_result[ASIN_COL].astype("string").str.strip()
        sub_result["is_match"] = find_key.isin(set(total_key.dropna()))

        match_df = sub_result.loc[sub_result["is_match"], [FIND_COL, ASIN_COL]].copy()
        stat_df = (
            match_df.groupby(FIND_COL, dropna=False)
            .size()
            .reset_index(name="匹配数量")
            .assign(匹配状态="匹配成功")[[FIND_COL, "匹配状态", "匹配数量"]]
        )
        asin_map = (
            match_df[match_df[ASIN_COL].notna() & (match_df[ASIN_COL] != "")]
            .groupby(FIND_COL, dropna=False)[ASIN_COL]
            .apply(self._join_asins)
            .to_dict()
        )
        return {
            "total_key": total_key,
            "sub_result": sub_result,
            "match_df": match_df,
            "stat_df": stat_df,
            "asin_map": asin_map,
        }

    def _fill_codes(self, total_df, total_key, asin_map, merge_existing=False):
        # 将 asin_map 回填到主表编码列；merge_existing=True 时与已有编码合并去重
        result_df = total_df.copy()
        if TARGET_COL in result_df.columns:
            col = TARGET_COL
        elif ALT_TARGET_COL in result_df.columns:
            col = ALT_TARGET_COL
        else:
            col = TARGET_COL
            result_df[col] = pd.NA

        mapped = total_key.map(asin_map)
        if not merge_existing:
            result_df[col] = mapped
            return result_df, col

        for idx in result_df.index:
            new_v = mapped.loc[idx]
            if pd.isna(new_v) or (isinstance(new_v, str) and not str(new_v).strip()):
                continue
            old_v = result_df.at[idx, col]
            if pd.isna(old_v) or (isinstance(old_v, str) and not str(old_v).strip()):
                result_df.at[idx, col] = new_v
            else:
                result_df.at[idx, col] = self._merge_codes(old_v, new_v)
        return result_df, col

    def _join_asins(self, values, sep=","):
        # 同一订单下多个 asin 去重后用逗号拼接
        ordered = []
        for v in values:
            if pd.isna(v):
                continue
            t = str(v).strip()
            if t and t.lower() != "nan":
                ordered.append(t)
        return sep.join(list(dict.fromkeys(ordered)))

    def _merge_codes(self, existing, new_str, sep=","):
        # 多副表依次回填时，合并编码列且不去重已有片段
        if pd.isna(new_str):
            return existing
        new_s = str(new_str).strip()
        if not new_s or new_s.lower() == "nan":
            return existing

        def parts(v):
            if pd.isna(v):
                return []
            return [p.strip() for p in str(v).split(sep) if p.strip()]

        old_parts = parts(existing)
        seen = set(old_parts)
        merged = old_parts + [p for p in parts(new_s) if p not in seen]
        return sep.join(merged)


if __name__ == "__main__":
    import sys

    config = {
        "total_path": r"C:\RPA流程\根据未拆分完成的总表文件进行筛选并回填主表【编码】列\主表.xlsx",  # 主表
        "sub_paths": [  # 副表（可多个，按顺序依次匹配回填）
            r"C:\RPA流程\根据未拆分完成的总表文件进行筛选并回填主表【编码】列\副表.xlsx",
        ],
    }

    # 命令行可覆盖：python main.py <主表路径> <副表1> [副表2 ...]
    if len(sys.argv) >= 3:
        config["total_path"] = sys.argv[1]
        config["sub_paths"] = sys.argv[2:]

    if not config["total_path"].strip():
        raise ValueError("请在 config 中填写 total_path（主表路径）。")
    if not config["sub_paths"]:
        raise ValueError("请在 config 中填写 sub_paths（至少一个副表路径）。")

    excel = Excel_file(config["total_path"])
    result = excel.run(config["sub_paths"])
    sr = result["sub_result"]
    print(f"匹配成功：{int(sr['is_match'].sum())}，失败：{int((~sr['is_match']).sum())}")
    print(f"已写入：{result['saved_path']}")
