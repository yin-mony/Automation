"""
测试草稿文件（专用）：
1) 用于快速验证流程和结果；
2) 逻辑可运行，但不作为正式入口；
3) Filter_add.py 才是主逻辑实现文件。
"""
import pandas as pd
from pathlib import Path


base_dir = Path(__file__).parent
# 读取总表文件
total_df = pd.read_excel(base_dir / '总表A.xlsx')
# 读取需要与总表文件查找匹配的表格文件
find_df = pd.read_excel(base_dir / '店铺C表.xlsx')

# 统一列名，避免列名前后空格导致找不到列
total_df.columns = total_df.columns.str.strip()
find_df.columns = find_df.columns.str.strip()

# 匹配字段列名（两个表分别对应）
total_col = '描述'
find_col = 'myp_order_id'
asin_col = 'asin'
target_col = '编码（必填）'

# 列存在性校验
if total_col not in total_df.columns:
    raise KeyError(f"总表中不存在列: {total_col}")
if find_col not in find_df.columns:
    raise KeyError(f"查找表中不存在列: {find_col}")
if asin_col not in find_df.columns:
    raise KeyError(f"查找表中不存在列: {asin_col}")

# 全量完全匹配前先做基础清洗：
# 1) 转为 pandas string 类型，避免 NaN 被转成字符串 "nan"
# 2) 去除前后空格，避免格式差异造成误判
total_key = total_df[total_col].astype('string').str.strip()
find_key = find_df[find_col].astype('string').str.strip()
find_df[asin_col] = find_df[asin_col].astype('string').str.strip()

# 使用 set 提升大数据量下的匹配效率（完全匹配）
total_key_set = set(total_key.dropna())
find_df['is_match'] = find_key.isin(total_key_set)

# 仅做结果展示
print(f"总匹配行数: {len(find_df)}")
print(f"匹配成功: {int(find_df['is_match'].sum())}")
print(f"匹配失败: {int((~find_df['is_match']).sum())}")
print(find_df[[find_col, 'is_match']].head(10))

# 完全匹配成功后，依次展示：myp_order_id、匹配成功数量、asin字段值
match_df = find_df.loc[find_df['is_match'], [find_col, asin_col]].copy()
match_df[find_col] = find_key[find_df['is_match']].values

print("\n完全匹配结果（myp_order_id + 匹配数量 + asin值）：")
if match_df.empty:
    print("无匹配成功数据")
else:
    grouped = match_df.groupby(find_col, dropna=False)
    for order_id, group in grouped:
        match_count = len(group)
        asin_values = [str(v) for v in group[asin_col].tolist()]
        print(f"myp_order_id: {order_id} | 匹配数量: {match_count}")
        for asin_value in asin_values:
            print(f"asin: {asin_value}")

# 将符合匹配条件的 asin 回填至主表的“编码（必填）”列，多个 asin 用英文逗号拼接
asin_map = (
    match_df[match_df[asin_col].notna() & (match_df[asin_col] != "")]
    .groupby(find_col, dropna=False)[asin_col]
    .apply(lambda s: ",".join(s.astype(str).tolist()))
    .to_dict()
)

if target_col not in total_df.columns:
    total_df[target_col] = pd.NA

total_df[target_col] = total_key.map(asin_map)

print("\n主表回填预览（描述 + 编码（必填））：")
print(total_df[[total_col, target_col]].head(10))