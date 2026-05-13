from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import os

def process_excel_file(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    
    headers = []
    data = []
    
    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            headers = list(row)
        else:
            data.append(list(row))
    
    myp_order_id_col = headers.index('myp_order_id')
    msku_col = headers.index('msku')
    charged_amount_col = headers.index('charged_amount')
    
    order_groups = {}
    for row in data:
        order_id = row[myp_order_id_col]
        if order_id not in order_groups:
            order_groups[order_id] = []
        order_groups[order_id].append(row)

    def order_has_multiple_msku(rows):
        return len({r[msku_col] for r in rows}) > 1

    kept_order_ids = {
        oid for oid, rows in order_groups.items() if order_has_multiple_msku(rows)
    }
    filtered_data = [row for row in data if row[myp_order_id_col] in kept_order_ids]

    new_wb = Workbook()
    new_ws1 = new_wb.active
    new_ws1.title = 'Sheet1'
    new_ws1.append(headers)
    for row in filtered_data:
        new_ws1.append(row)
    
    new_ws2 = new_wb.create_sheet(title='sheet2')
    
    header_row = ['myp_order_id', 'msku', 'charged_amount']
    new_ws2.append(header_row)
    
    header_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
    total_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    total_font = Font(bold=True)
    right_align = Alignment(horizontal='right')
    
    for col in range(1, len(header_row) + 1):
        new_ws2.cell(row=1, column=col).fill = header_fill
    
    overall_total = 0
    for order_id, rows in order_groups.items():
        if order_id not in kept_order_ids:
            continue
        first_row = True
        for row in rows:
            if first_row:
                new_ws2.append([row[myp_order_id_col], row[msku_col], row[charged_amount_col]])
                first_row = False
            else:
                new_ws2.append(['', row[msku_col], row[charged_amount_col]])
            overall_total += row[charged_amount_col] if row[charged_amount_col] else 0
    
    new_ws2.append(['', '', ''])
    new_ws2.append(['总计', '', overall_total])
    total_row = new_ws2.max_row
    for col in range(1, 4):
        cell = new_ws2.cell(row=total_row, column=col)
        cell.fill = total_fill
        cell.font = total_font
        if col == 3:
            cell.alignment = right_align
    
    new_wb.save(file_path)
    
    return len(filtered_data)

def process_multiple_files(file_paths):
    results = {}
    for file_path in file_paths:
        if os.path.exists(file_path):
            try:
                count = process_excel_file(file_path)
                results[file_path] = count
                print(f"处理完成: {file_path}")
                print(f"记录数: {count}")
            except Exception as e:
                print(f"处理失败 {file_path}: {str(e)}")
                results[file_path] = None
        else:
            print(f"文件不存在: {file_path}")
            results[file_path] = None
    return results