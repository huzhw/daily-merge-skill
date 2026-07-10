# -*- coding: utf-8 -*-
"""月初脚本：从模板创建当月首个日报，序号从 1 开始"""
import openpyxl, os, shutil
from datetime import datetime
from copy import copy
from openpyxl.styles import Alignment

DESKTOP = r"C:\Users\Administrator\Desktop"
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
YEAR = TODAY.strftime("%Y")
MONTH = str(TODAY.month)
MM = TODAY.strftime("%m")
DD = TODAY.strftime("%d")

REPO_MAP = {
    'lanxum-amisp': '档案V6', 'lanxum-amisp-java': '档案V6', 'lanxum-amisp-react': '档案V6',
    'workingpaper-v5.5': '中信底稿v5',
}

TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'templates', '日报模板.xlsx')
REPORT_DIR = os.path.join(DESKTOP, f"报告-{YEAR}年", f"日报-{YEAR}-{MONTH}月")
MD_FILE = os.path.join(REPORT_DIR, f"日报需求记录-{YEAR}-{MM}-{DD}.md")
XLSX_FILE = os.path.join(REPORT_DIR, f"日报表格-胡志伟~~{MM}-{DD}.xlsx")

def next_workday(d):
    d = d + __import__('datetime').timedelta(days=1)
    while d.weekday() >= 5:
        d = d + __import__('datetime').timedelta(days=1)
    return d

def parse_md(filepath):
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"md 文件不存在: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    tasks = []
    in_table = False
    for line in content.split('\n'):
        line = line.strip()
        if line.startswith('| 序号 |'): in_table = True; continue
        if line.startswith('|------|'): continue
        if not in_table or not line.startswith('|'): continue
        if line.startswith('| >') or '空行为模板' in line: break
        cells = [c.strip() for c in line.split('|')[1:-1]]
        if len(cells) < 8: continue
        seq, date, repo, desc, modules, status, human_h, ai_h, note = cells[:9]
        is_done = status in ('已完成', '100%')
        if is_done and seq.isdigit():
            pct = status.replace('%', '') if '%' in status else '100'
            tasks.append({'repo':repo,'desc':desc,'human_h':float(human_h),'ai_h':float(ai_h),'note':note,'pct':f'{pct}%'})
    return tasks

def calc_g(hours, prev_date, remaining):
    total = remaining + hours
    d = prev_date
    while total > 8:
        total -= 8
        d = next_workday(d)
    return d, total

def copy_style(src, dst):
    """复制字号字体但不复制颜色（表头是白色）"""
    if src.has_style:
        from openpyxl.styles import Font
        sf = src.font
        dst.font = Font(name=sf.name, size=sf.size, bold=sf.bold, italic=sf.italic)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)

def main():
    tasks = parse_md(MD_FILE)
    if not tasks:
        print("没有已完成的任务，跳过。")
        return
    print(f"找到 {len(tasks)} 条已完成任务")

    # 从模板起新表
    os.makedirs(REPORT_DIR, exist_ok=True)
    shutil.copy(TEMPLATE, XLSX_FILE)
    wb = openpyxl.load_workbook(XLSX_FILE)
    ws = wb[wb.sheetnames[0]]

    # 拆掉模板里残留的 A 列合并
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col == 1:
            try: ws.unmerge_cells(str(mr))
            except: pass

    # 从第 2 行开始写
    insert_pos = 2
    current_g = TODAY
    current_remaining = 0
    prev_repo = None
    seq = 0

    for task in tasks:
        ws.insert_rows(insert_pos)
        ws.cell(row=insert_pos, column=1).value = TODAY
        ws.cell(row=insert_pos, column=1).number_format = 'yyyy/m/d;@'

        repo = task['repo']
        display = REPO_MAP.get(repo, repo)
        if repo != prev_repo:
            ws.cell(row=insert_pos, column=2).value = display
            prev_repo = repo

        seq += 1
        ws.cell(row=insert_pos, column=3).value = seq
        ws.cell(row=insert_pos, column=4).value = task['desc']
        ws.cell(row=insert_pos, column=5).value = task['pct']

        for c in (6, 10):
            cell = ws.cell(row=insert_pos, column=c)
            cell.value = TODAY
            cell.number_format = 'yyyy/m/d;@'

        current_g, current_remaining = calc_g(task['human_h'], current_g, current_remaining)
        g_cell = ws.cell(row=insert_pos, column=7)
        g_cell.value = current_g
        g_cell.number_format = 'yyyy/m/d;@'

        ws.cell(row=insert_pos, column=8).value = task['human_h']
        ws.cell(row=insert_pos, column=11).value = task['ai_h']
        ws.cell(row=insert_pos, column=14).value = task['note']

        # 样式参考第 1 行表头
        for c in range(1, 15):
            copy_style(ws.cell(row=1, column=c), ws.cell(row=insert_pos, column=c))
        for c in (4, 14):
            ws.cell(row=insert_pos, column=c).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # 行高
        lines = max(1, len(task['desc']) / 35)
        ws.row_dimensions[insert_pos].height = max(30, lines * 15)

        print(f"  新增行 {insert_pos}: [{display}] #{seq} G={current_g.strftime('%m-%d')} H={task['human_h']}h K={task['ai_h']}h")
        insert_pos += 1

    # A 列合并
    if len(tasks) > 1:
        ws.merge_cells(start_row=2, start_column=1, end_row=insert_pos-1, end_column=1)
    for r in range(2, insert_pos):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='center')

    # 列宽
    cw = {'A':13,'B':23,'C':15,'D':60,'E':16,'F':16,'G':16,'H':14,'I':24,'J':13,'K':15,'L':12,'M':42,'N':60}
    for k, v in cw.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = 'A2'

    wb.save(XLSX_FILE)
    print(f"\n已保存: {XLSX_FILE}")

if __name__ == '__main__':
    main()
