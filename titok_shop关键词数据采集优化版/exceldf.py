import pandas as pd
from pathlib import Path


class ExcelDF:
    def __init__(self, config):
        self.path = config['file_path']
        self.sheet_name = '关键词表格'
        self.keyword_col = '关键词'

    # 读取配置里的 Excel，从「关键词表格」工作表提取「关键词」列
    def read_excel(self):
        data = []

        excel_path = Path(self.path)
        if not excel_path.exists():
            raise FileNotFoundError(f'未找到 Excel 文件: {excel_path}')

        df = pd.read_excel(excel_path, sheet_name=self.sheet_name)
        if self.keyword_col not in df.columns:
            raise ValueError(f'工作表「{self.sheet_name}」中未找到「{self.keyword_col}」列')

        seen = set()
        for _, row in df.iterrows():
            keyword = row[self.keyword_col]
            if pd.isna(keyword):
                continue
            keyword = str(keyword).strip()
            if not keyword or keyword in seen:
                continue
            data.append(keyword)
            seen.add(keyword)

        print(data)
        return data


if __name__ == '__main__':
    config = {
        'file_path': r'C:\Users\admin\Desktop\tiktok-竞品信息抓取.xlsx',
    }
    exceldf = ExcelDF(config)
    exceldf.read_excel()
