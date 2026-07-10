# -*- coding: utf-8 -*-
"""从现有日报 Excel 提取表头+备注生成空白模板，保留原样式"""
import openpyxl, os, shutil

# 从最新的 Excel 复制
base = r"C:\Users\Administrator\Desktop\报告-2026年\日报-2026-7月"
src = None
for f in sorted(os.listdir(base), reverse=True):
    if f.startswith('日报表格-胡志伟~~') and f.endswith('.xlsx'):
        src = os.path.join(base, f)
        break

if src is None:
    print('错误：找不到源文件')
    exit(1)

print(f'源文件: {src}')
out_dir = os.path.dirname(os.path.abspath(__file__))
out_path = os.path.join(out_dir, '日报模板.xlsx')
shutil.copy(src, out_path)

wb = openpyxl.load_workbook(out_path)
ws = wb[wb.sheetnames[0]]

# 找到备注行
notes_start = 0
for row in range(1, ws.max_row + 1):
    a_val = ws.cell(row=row, column=1).value
    if a_val and '填充说明' in str(a_val):
        notes_start = row
        break

# 找到最后一条数据行（备注行之前最后一个有序号的行）
last_data = 0
for row in range(1, notes_start):
    c_val = ws.cell(row=row, column=3).value
    if c_val is not None and str(c_val).strip().isdigit():
        last_data = row

print(f'最后数据行: {last_data}, 备注起始: {notes_start}')

# 删除所有数据行（从第2行到备注行之前）
if last_data >= 2:
    for _ in range(2, notes_start):
        ws.delete_rows(2)
    print(f'已清除数据行 2-{last_data}')

# 删除 A 列合并
for mr in list(ws.merged_cells.ranges):
    if mr.min_col == 1:
        ws.unmerge_cells(str(mr))

# 清空序号列（C列）从第2行开始，留 3 行
for r in range(2, 5):
    for c in range(1, 15):
        ws.cell(row=r, column=c).value = None
    ws.row_dimensions[r].height = 25

ws.freeze_panes = 'A2'
wb.save(out_path)
print(f'已生成: {out_path}')
