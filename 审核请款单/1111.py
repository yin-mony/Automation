from pathlib import Path

target_folder = Path(r"C:\Users\admin\Downloads")
matched_files = list(target_folder.glob("请款单*.zip"))

if matched_files:
    # 直接找修改时间最新的文件
    latest_file = max(matched_files, key=lambda f: f.stat().st_mtime)
    print(f"✅ 最新的文件是: {latest_file}")
else:
    print("未找到符合条件的文件。")