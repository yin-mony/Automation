import os
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from difflib import SequenceMatcher
import pdfplumber

# 1. 设置已知文件夹路径（可由 Qt 界面通过环境变量 TARGET_FOLDER 传入）
target_folder = Path(os.environ.get("TARGET_FOLDER", r"C:\Users\admin\Desktop\赛狐请款单下载\请款单批量下载20260528"))

# 2. 打开文件夹
print(f"正在打开文件夹: {target_folder}")
os.startfile(target_folder)

# 3. 创建根目录下的 Excel 统一存放文件夹
excel_output_folder = target_folder / "Excel转换结果"
excel_output_folder.mkdir(exist_ok=True)
print(f"[INFO] Excel 文件将统一保存在: {excel_output_folder}")

# 4. 遍历当前文件夹下的所有子文件夹
print(f"\n正在遍历 {target_folder} 下的所有子文件夹...")
print("=" * 50)

# 获取当前文件夹下所有的子文件夹
subfolders = [f for f in target_folder.iterdir() if f.is_dir()]

if not subfolders:
    print("[ERROR] 当前文件夹下没有子文件夹")
    exit()


# 5. 定义 PDF 转 Excel 函数（保留表头，去除空行）
def pdf_to_excel(pdf_path, output_path):
    """将 PDF 文件转换为 Excel 文件，保留表头，去除空行"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []

            # 遍历所有页面，提取表格
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                for table_num, table in enumerate(tables):
                    if table:  # 如果表格不为空
                        # 将表格转换为 DataFrame
                        df = pd.DataFrame(table)

                        # ✅ 去除完全为空的行（所有列都是空值）
                        df = df.dropna(how='all')

                        # ✅ 去除所有列都是空字符串的行
                        df = df[(df != '').any(axis=1)]

                        # ✅ 如果第一行是空行（所有列都是空值或空字符串），去除它
                        if not df.empty:
                            # 检查第一行是否全部为空
                            first_row_empty = df.iloc[0].isna().all() or (df.iloc[0] == '').all()
                            if first_row_empty:
                                # 删除第一行空行
                                df = df.iloc[1:]
                                # 重置索引
                                df = df.reset_index(drop=True)

                        # ✅ 处理表头：将第一行设为表头
                        if not df.empty:
                            # 将第一行设为列名
                            df.columns = df.iloc[0]
                            # 删除第一行（因为已经变成列名了）
                            df = df.iloc[1:]
                            # 重置索引
                            df = df.reset_index(drop=True)

                        # ✅ 再次去除可能产生的空行
                        df = df.dropna(how='all')
                        df = df[(df != '').any(axis=1)]

                        all_tables.append(df)

            if all_tables:
                # 合并所有表格
                combined_df = pd.concat(all_tables, ignore_index=True)
                # 保存为 Excel
                combined_df.to_excel(output_path, index=False)
                return True
            else:
                # 如果没有表格，提取文本保存
                text_data = []
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if line.strip():
                                text_data.append(line.strip())

                if text_data:
                    df = pd.DataFrame(text_data, columns=["内容"])
                    df.to_excel(output_path, index=False)
                    return True
                else:
                    return False

    except Exception as e:
        print(f"   [ERROR] 转换失败: {e}")
        return False


# 6. 遍历所有子文件夹，转换 PDF 文件
print(f"\n[INFO] 开始转换 PDF 文件为 Excel...")
print("=" * 50)

converted_count = 0
failed_count = 0

for idx, folder in enumerate(subfolders, 1):
    print(f"\n[INFO] 子文件夹 {idx}: {folder.name}")

    # 查找该子文件夹中的所有 PDF 文件
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        print("   没有 PDF 文件")
        continue

    for pdf_file in pdf_files:
        print(f"   [FILE] 处理文件: {pdf_file.name}")

        # 生成输出文件名
        output_filename = f"{folder.name}_{pdf_file.stem}.xlsx"
        output_file = excel_output_folder / output_filename

        # 转换
        success = pdf_to_excel(pdf_file, output_file)

        if success:
            print(f"   [OK] 转换成功: {output_filename}")
            converted_count += 1
        else:
            print(f"   [ERROR] 转换失败: {pdf_file.name}")
            failed_count += 1

# 7. 输出统计结果
print("\n" + "=" * 50)
print(f"[INFO] 转换统计:")
print(f"   [OK] 成功转换: {converted_count} 个文件")
print(f"   [ERROR] 转换失败: {failed_count} 个文件")
print(f"[INFO] 所有 Excel 文件保存在: {excel_output_folder}")

# 8. 打开 Excel 输出文件夹
print(f"\n正在打开输出文件夹...")
os.startfile(excel_output_folder)

# ============================================
# 9. 汇总所有 Excel 文件中的字段数据（按行展开）
# ============================================
print("\n" + "=" * 50)
print("[INFO] 开始汇总所有 Excel 文件中的字段数据...")

# 定义需要提取的列名
target_columns = ["采购单号", "品名/SKU", "采购单价"]

# 存储所有提取的数据（按行展开）
all_extracted_data = []

# 遍历 Excel 输出文件夹中的所有 Excel 文件
excel_files = list(excel_output_folder.glob("*.xlsx"))

if not excel_files:
    print("[ERROR] 没有找到任何 Excel 文件")
else:
    print(f"[INFO] 共找到 {len(excel_files)} 个 Excel 文件")

    for excel_file in excel_files:
        print(f"   [FILE] 处理文件: {excel_file.name}")

        try:
            # 读取 Excel 文件
            df = pd.read_excel(excel_file)

            # 查找并提取目标列的数据
            extracted_data = {}

            for target_col in target_columns:
                # 在 DataFrame 的列名中查找包含目标列名的列
                matching_cols = [col for col in df.columns if target_col in str(col)]

                if matching_cols:
                    col_name = matching_cols[0]
                    # 提取该列的所有非空值
                    values = df[col_name].dropna().tolist()
                    # 去除空字符串和只包含空格的值
                    values = [str(v).strip() for v in values if str(v).strip()]
                    extracted_data[target_col] = values
                else:
                    extracted_data[target_col] = []

            # 确定最大行数（取三个列中数据最多的行数）
            max_rows = max(len(extracted_data[col]) for col in target_columns) if any(extracted_data.values()) else 0

            if max_rows > 0:
                # 按行展开数据
                for i in range(max_rows):
                    row_data = {
                        "源文件名": excel_file.name,
                        "行号": i + 1
                    }
                    for col in target_columns:
                        if i < len(extracted_data[col]):
                            row_data[col] = extracted_data[col][i]
                        else:
                            row_data[col] = ""
                    all_extracted_data.append(row_data)
            else:
                # 如果没有数据，至少记录文件名
                row_data = {
                    "源文件名": excel_file.name,
                    "行号": 1,
                    "采购单号": "",
                    "品名/SKU": "",
                    "采购单价": ""
                }
                all_extracted_data.append(row_data)

        except Exception as e:
            print(f"   [ERROR] 读取 Excel 文件失败: {e}")

# 10. 保存汇总数据到新的 Excel 文件
print("\n" + "=" * 50)
print("[INFO] 保存汇总数据...")

if all_extracted_data:
    # 创建汇总 DataFrame
    summary_df = pd.DataFrame(all_extracted_data)

    # 11. 拆分“品名/SKU”为两列（优先使用该列）
    print("[INFO] 开始拆分“品名/SKU”列...")

    split_source_col = ""
    for col in summary_df.columns:
        col_text = str(col).lower()
        if "品名/sku" in col_text or "/sku" in col_text:
            split_source_col = col
            break

    fallback_name_col = ""
    fallback_sku_col = ""
    for col in summary_df.columns:
        col_text = str(col).lower()
        if (not fallback_name_col) and ("品名" in str(col)) and ("/" not in str(col)):
            fallback_name_col = col
        if (not fallback_sku_col) and ("sku" in col_text) and (col != split_source_col):
            fallback_sku_col = col

    split_names = []
    split_skus = []

    for _, row in summary_df.iterrows():
        source_value = row.get(split_source_col, "") if split_source_col else ""
        source_text = str(source_value).strip() if pd.notna(source_value) else ""

        if source_text:
            lines = [line.strip() for line in re.split(r"[\r\n]+", source_text) if line.strip()]

            if len(lines) <= 1:
                name_value = re.sub(r"\s+", " ", source_text).strip()
                sku_value = ""
            else:
                best_score = -1
                best_similarity = 0
                best_name = ""
                best_sku = ""

                for split_index in range(1, len(lines)):
                    name_part = " ".join(lines[:split_index]).strip()
                    sku_part = "".join(lines[split_index:]).strip()
                    if (not name_part) or (not sku_part):
                        continue

                    norm_name = re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "", name_part).lower()
                    norm_sku = re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "", sku_part).lower()
                    similarity = SequenceMatcher(None, norm_name, norm_sku).ratio()
                    balance = 1 - abs(split_index - len(lines) / 2) / len(lines)
                    score = similarity * 0.75 + balance * 0.25

                    if score > best_score:
                        best_score = score
                        best_similarity = similarity
                        best_name = name_part
                        best_sku = sku_part

                if (best_score >= 0) and ((best_similarity >= 0.45) or (len(lines) >= 4)):
                    name_value = re.sub(r"\s+", " ", best_name).strip()
                    sku_value = best_sku
                else:
                    name_value = re.sub(r"\s+", " ", " ".join(lines[:-1])).strip()
                    sku_value = lines[-1].strip()
        else:
            fallback_name = row.get(fallback_name_col, "") if fallback_name_col else ""
            fallback_sku = row.get(fallback_sku_col, "") if fallback_sku_col else ""
            name_value = str(fallback_name).strip() if pd.notna(fallback_name) else ""
            sku_value = str(fallback_sku).strip() if pd.notna(fallback_sku) else ""

        # 增强规则：在“原拆分结果”基础上再归类
        # 含中文的部分归“品名”，不含中文的部分归“SKU”
        if source_text:
            line_parts = [line.strip() for line in re.split(r"[\r\n]+", source_text) if line.strip()]
            chinese_parts = [part for part in line_parts if re.search(r"[\u4e00-\u9fff]", part)]
            non_chinese_parts = [part for part in line_parts if not re.search(r"[\u4e00-\u9fff]", part)]

            # 以原拆分结果为主，不覆盖成“纯中文”；仅补充中文信息到品名
            if chinese_parts:
                chinese_text = re.sub(r"\s+", " ", " ".join(chinese_parts)).strip()
                if chinese_text:
                    if name_value:
                        if chinese_text not in name_value:
                            name_value = re.sub(r"\s+", " ", f"{name_value} {chinese_text}").strip()
                    else:
                        name_value = chinese_text

            # SKU 优先沿用原拆分结果；为空时再从非中文行提取
            sku_from_split = "".join(re.findall(r"[A-Za-z0-9]+", str(sku_value)))
            sku_from_non_chinese = "".join(re.findall(r"[A-Za-z0-9]+", "".join(non_chinese_parts)))
            if sku_from_split:
                sku_value = sku_from_split
            else:
                sku_value = sku_from_non_chinese
        else:
            if re.search(r"[\u4e00-\u9fff]", str(sku_value)) and (not re.search(r"[\u4e00-\u9fff]", str(name_value))):
                name_value, sku_value = sku_value, name_value
            sku_value = "".join(re.findall(r"[A-Za-z0-9]+", str(sku_value)))

        split_names.append(name_value)
        split_skus.append(sku_value)

    summary_df["品名"] = split_names
    summary_df["SKU"] = split_skus
    if split_source_col and split_source_col in summary_df.columns:
        summary_df = summary_df.drop(columns=[split_source_col])
    print("[OK] “品名/SKU”拆分完成")

    # 12. 生成汇总文件名（包含日期）
    today = datetime.now().strftime("%Y%m%d")
    summary_file = target_folder / f"请款单汇总_{today}.xlsx"

    # 13. 保存汇总文件（若被占用则自动改名）
    try:
        summary_df.to_excel(summary_file, index=False)
    except PermissionError:
        summary_file = target_folder / f"请款单汇总_{today}_{datetime.now().strftime('%H%M%S')}.xlsx"
        summary_df.to_excel(summary_file, index=False)

    print(f"[OK] 汇总完成！")
    print(f"[INFO] 汇总文件保存路径: {summary_file}")
    print(f"[INFO] 共汇总 {len(all_extracted_data)} 条数据记录")
    print(f"[INFO] 涉及 {len(excel_files)} 个 Excel 文件")

    # 打开汇总文件
    os.startfile(summary_file)
else:
    print("[ERROR] 没有提取到任何数据，无法生成汇总文件")
