from pathlib import Path

import pandas as pd


DEFAULT_TOTAL_COL = "描述"
DEFAULT_FIND_COL = "myp_order_id"
DEFAULT_ASIN_COL = "asin"
DEFAULT_TARGET_COL = "编码（必填）"
ALT_TARGET_COL = "编码(必填)"


def _dedupe_join_asins(values, sep=","):
    # 同一匹配键下多条 asin 合并时：去重，保留首次出现顺序；跳过空值
    ordered = []
    for v in values:
        if pd.isna(v):
            continue
        t = str(v).strip()
        if not t or t.lower() == "nan":
            continue
        ordered.append(t)
    unique = list(dict.fromkeys(ordered))
    return sep.join(unique)


def _merge_code_strings(existing, new_str, sep=","):
    # 多副表依次回填时：保留已有片段顺序，仅追加新片段中尚未出现过的部分（与单表内去重规则一致）
    if pd.isna(new_str):
        return existing
    new_s = str(new_str).strip()
    if not new_s or new_s.lower() == "nan":
        return existing

    def split_codes(v):
        if pd.isna(v):
            return []
        return [p.strip() for p in str(v).split(sep) if p.strip()]

    old_parts = split_codes(existing)
    new_parts = split_codes(new_s)
    seen = set(old_parts)
    merged = old_parts + [p for p in new_parts if p not in seen]
    return sep.join(merged)


def load_data(total_path, sub_path):
    # 参数: total_path=主表路径, sub_path=副表路径
    # 返回: (total_df, sub_df) 两个 DataFrame
    total_df = pd.read_excel(total_path)
    sub_df = pd.read_excel(sub_path)
    total_df.columns = total_df.columns.str.strip()
    sub_df.columns = sub_df.columns.str.strip()
    return total_df, sub_df


def validate_columns(
    total_df,
    sub_df,
    total_col,
    find_col,
    asin_col,
):
    # 参数: total_df=主表DataFrame, sub_df=副表DataFrame, total_col/find_col/asin_col=需校验列名
    # 返回: 无返回值; 若缺少列会抛出 KeyError
    if total_col not in total_df.columns:
        raise KeyError(f"总表中不存在列: {total_col}")
    if find_col not in sub_df.columns:
        raise KeyError(f"副表中不存在列: {find_col}")
    if asin_col not in sub_df.columns:
        raise KeyError(f"副表中不存在列: {asin_col}")


def match_and_collect(
    total_df,
    sub_df,
    total_col=DEFAULT_TOTAL_COL,
    find_col=DEFAULT_FIND_COL,
    asin_col=DEFAULT_ASIN_COL,
):
    # 参数: total_df/sub_df=输入表, total_col/find_col/asin_col=匹配与取值列名
    # 返回: dict, 包含 total_key/sub_result/match_df/stat_df/asin_map
    total_key = total_df[total_col].astype("string").str.strip()
    find_key = sub_df[find_col].astype("string").str.strip()

    sub_result = sub_df.copy()
    sub_result[find_col] = find_key
    sub_result[asin_col] = sub_result[asin_col].astype("string").str.strip()

    total_key_set = set(total_key.dropna())
    sub_result["is_match"] = find_key.isin(total_key_set)

    match_df = sub_result.loc[sub_result["is_match"], [find_col, asin_col]].copy()
    stat_df = (
        match_df.groupby(find_col, dropna=False)
        .size()
        .reset_index(name="匹配数量")
        .assign(匹配状态="匹配成功")[[find_col, "匹配状态", "匹配数量"]]
    )

    asin_map = (
        match_df[match_df[asin_col].notna() & (match_df[asin_col] != "")]
        .groupby(find_col, dropna=False)[asin_col]
        .apply(_dedupe_join_asins)
        .to_dict()
    )

    return {
        "total_key": total_key,
        "sub_result": sub_result,
        "match_df": match_df,
        "stat_df": stat_df,
        "asin_map": asin_map,
    }


def fill_target_column(
    total_df,
    total_key,
    asin_map,
    target_col=DEFAULT_TARGET_COL,
    merge_existing=False,
):
    # 参数: total_df=主表DataFrame, total_key=主表匹配键Series, asin_map=订单到asin映射, target_col=回填列
    # 参数: merge_existing=False 时与原先一致（整列按 map 写入）；True 时在已有编码格上合并本轮 map，不重复追加
    # 返回: (回填后的主表 DataFrame, 实际写入列名)
    result_df = total_df.copy()
    if target_col in result_df.columns:
        target_col_used = target_col
    elif ALT_TARGET_COL in result_df.columns:
        target_col_used = ALT_TARGET_COL
    else:
        target_col_used = target_col
        result_df[target_col_used] = pd.NA

    mapped = total_key.map(asin_map)
    if not merge_existing:
        result_df[target_col_used] = mapped
        return result_df, target_col_used

    if target_col_used not in result_df.columns:
        result_df[target_col_used] = pd.NA
    for idx in result_df.index:
        new_v = mapped.loc[idx]
        if pd.isna(new_v) or (isinstance(new_v, str) and not str(new_v).strip()):
            continue
        old_v = result_df.at[idx, target_col_used]
        if pd.isna(old_v) or (isinstance(old_v, str) and not str(old_v).strip()):
            result_df.at[idx, target_col_used] = new_v
        else:
            result_df.at[idx, target_col_used] = _merge_code_strings(old_v, new_v)
    return result_df, target_col_used


def print_match_summary(
    sub_result,
    match_df,
    stat_df,
    total_df_filled,
    find_col=DEFAULT_FIND_COL,
    asin_col=DEFAULT_ASIN_COL,
    total_col=DEFAULT_TOTAL_COL,
    target_col=DEFAULT_TARGET_COL,
):
    # 参数: sub_result/match_df/stat_df/total_df_filled=流程结果DataFrame, 其余为列名
    # 返回: 无返回值; 仅打印结果摘要与预览
    print(f"总匹配行数: {len(sub_result)}")
    print(f"匹配成功: {int(sub_result['is_match'].sum())}")
    print(f"匹配失败: {int((~sub_result['is_match']).sum())}")
    print(sub_result[[find_col, "is_match"]].head(10))

    print("\n完全匹配结果（myp_order_id + 匹配状态 + 匹配数量）：")
    print(stat_df if not stat_df.empty else "无匹配成功数据")

    print("\n完全匹配结果（myp_order_id + 匹配数量 + asin值）：")
    if match_df.empty:
        print("无匹配成功数据")
    else:
        grouped = match_df.groupby(find_col, dropna=False)
        for order_id, group in grouped:
            print(f"myp_order_id: {order_id} | 匹配数量: {len(group)}")
            for asin_value in group[asin_col].tolist():
                print(f"asin: {asin_value}")

    print("\n主表回填预览（描述 + 编码（必填））：")
    print(total_df_filled[[total_col, target_col]].head(10))


def _normalize_sub_paths(sub_path):
    # 统一为 Path 列表：支持单个路径或路径列表/元组
    if sub_path is None:
        return []
    if isinstance(sub_path, (list, tuple)):
        return [Path(p) for p in sub_path if str(p).strip()]
    return [Path(sub_path)]


def run_pipeline(
    total_path,
    sub_path,
    total_col=DEFAULT_TOTAL_COL,
    find_col=DEFAULT_FIND_COL,
    asin_col=DEFAULT_ASIN_COL,
    target_col=DEFAULT_TARGET_COL,
    print_summary=True,
    save_result=False,
    output_path=None,
):
    # 参数: total_path=主表路径; sub_path=单个副表路径，或多个副表路径组成的 list/tuple（按顺序依次匹配回填）
    # 参数: 列名参数=流程字段配置, print_summary=是否打印摘要
    # 参数: save_result=是否写回文件, output_path=输出文件路径(None时覆盖主表文件)
    # 返回: dict, 包含匹配结果和回填后的 total_df_filled 与写入路径
    sub_paths = _normalize_sub_paths(sub_path)
    if not sub_paths:
        raise ValueError("至少需要一个副表文件路径。")

    for sp in sub_paths:
        if not Path(sp).exists():
            raise FileNotFoundError(f"副表文件不存在: {sp}")

    total_df = pd.read_excel(total_path)
    total_df.columns = total_df.columns.str.strip()

    total_work = total_df.copy()
    per_sub_match_results = []
    sub_result_frames = []
    match_dfs = []
    stat_dfs = []

    for i, sp in enumerate(sub_paths):
        sub_df = pd.read_excel(sp)
        sub_df.columns = sub_df.columns.str.strip()
        validate_columns(
            total_df=total_work,
            sub_df=sub_df,
            total_col=total_col,
            find_col=find_col,
            asin_col=asin_col,
        )

        match_result = match_and_collect(
            total_df=total_work,
            sub_df=sub_df,
            total_col=total_col,
            find_col=find_col,
            asin_col=asin_col,
        )

        total_work, target_col_used = fill_target_column(
            total_df=total_work,
            total_key=match_result["total_key"],
            asin_map=match_result["asin_map"],
            target_col=target_col,
            merge_existing=(i > 0),
        )

        entry = {
            **match_result,
            "sub_path": str(sp),
            "total_df_filled_step": total_work.copy(),
        }
        per_sub_match_results.append(entry)

        sr = match_result["sub_result"].copy()
        sr["副表文件"] = Path(sp).name
        sub_result_frames.append(sr)
        match_dfs.append(match_result["match_df"])
        stat_dfs.append(match_result["stat_df"])

        if print_summary:
            print(f"\n========== 副表 {i + 1}/{len(sub_paths)}: {sp} ==========")
            print_match_summary(
                sub_result=match_result["sub_result"],
                match_df=match_result["match_df"],
                stat_df=match_result["stat_df"],
                total_df_filled=total_work,
                find_col=find_col,
                asin_col=asin_col,
                total_col=total_col,
                target_col=target_col_used,
            )

    combined_sub_result = (
        pd.concat(sub_result_frames, ignore_index=True) if len(sub_result_frames) > 1 else sub_result_frames[0]
    )
    combined_match_df = pd.concat(match_dfs, ignore_index=True) if len(match_dfs) > 1 else match_dfs[0]
    combined_stat_df = pd.concat(stat_dfs, ignore_index=True) if len(stat_dfs) > 1 else stat_dfs[0]

    last_mr = per_sub_match_results[-1]
    result = {
        "total_key": last_mr["total_key"],
        "sub_result": combined_sub_result,
        "match_df": combined_match_df,
        "stat_df": combined_stat_df,
        "asin_map": last_mr["asin_map"],
        "total_df_filled": total_work,
        "target_col_used": target_col_used,
        "sub_paths": [str(p) for p in sub_paths],
        "per_sub_match_results": per_sub_match_results,
    }

    if save_result:
        save_path = Path(output_path) if output_path else Path(total_path)
        save_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            result["total_df_filled"].to_excel(save_path, index=False)
        except PermissionError as exc:
            raise PermissionError(f"无法写入文件: {save_path}。请关闭该Excel文件后重试。") from exc
        result["saved_path"] = str(save_path)
        print(f"\n已写入主表文件: {save_path}")
    return result


def run_interactive():
    # 参数: 无
    # 返回: run_pipeline 的结果 dict
    total_input = input("请输入主表路径（必填）: ").strip()
    print("请输入副表路径（必填）。多个文件请逐行输入，输入空行结束：")
    sub_lines = []
    while True:
        line = input().strip()
        if not line:
            break
        sub_lines.append(line)
    if not total_input or not sub_lines:
        raise ValueError("主表路径和至少一个副表路径都必须输入，不能留空。")

    total_path = Path(total_input)
    if not total_path.exists():
        raise FileNotFoundError(f"主表文件不存在: {total_path}")

    sub_paths = []
    for line in sub_lines:
        p = Path(line)
        if not p.exists():
            raise FileNotFoundError(f"副表文件不存在: {p}")
        sub_paths.append(p)

    output_input = input("请输入输出主表路径（留空则覆盖原主表）: ").strip()
    output_path = Path(output_input) if output_input else None

    return run_pipeline(
        total_path=total_path,
        sub_path=sub_paths,
        save_result=True,
        output_path=output_path,
    )


if __name__ == "__main__":
    run_interactive()