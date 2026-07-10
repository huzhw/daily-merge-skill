# -*- coding: utf-8 -*-
"""月初脚本：从模板复制，直接往空行写数据，不插行"""
import openpyxl, os, shutil
from datetime import datetime, timedelta
from copy import copy
from openpyxl.styles import Font, Alignment

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
    d = d + timedelta(days=1)
    while d.weekday() >= 5: d = d + timedelta(days=1)
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
            tasks.append({'repo':repo,'desc':desc,'human_h':float(human_h),'ai_h':float(ai_h),'note':note,'pct':pct+'%'})
    return tasks

def calc_g(hours, prev_date, remaining):
    total = remaining + hours
    d = prev_date
    while total > 8: total -= 8; d = next_workday(d)
    return d, total

def find_notes(ws):
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=1).value and '填充说明' in str(ws.cell(row=row, column=1).value):
            return row
    return ws.max_row + 3

def main():
    tasks = parse_md(MD_FILE)
    if not tasks:
        print("没有已完成的任务，跳过。")
        return
    print(f"找到 {len(tasks)} 条已完成任务")

    os.makedirs(REPORT_DIR, exist_ok=True)
    shutil.copy(TEMPLATE, XLSX_FILE)
    wb = openpyxl.load_workbook(XLSX_FILE)
    ws = wb[wb.sheetnames[0]]

    notes_row = find_notes(ws)
    # 第2行到备注前全是空行，直接写，不插行
    insert_pos = 2
    current_g = TODAY
    current_remaining = 0
    prev_repo = None
    seq = 0

    data_font = Font(name='宋体', size=10)
    wrap_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    for task in tasks:
        # 如果空行不够，在备注前补插
        if insert_pos >= notes_row:
            ws.insert_rows(notes_row)
            notes_row += 1

        r = insert_pos
        ws.cell(row=r, column=1).value = TODAY
        ws.cell(row=r, column=1).number_format = 'yyyy/m/d;@'
        ws.cell(row=r, column=1).font = data_font

        repo = task['repo']
        display = REPO_MAP.get(repo, repo)
        if repo != prev_repo:
            ws.cell(row=r, column=2).value = display
            ws.cell(row=r, column=2).font = data_font
            prev_repo = repo

        seq += 1
        ws.cell(row=r, column=3).value = seq
        ws.cell(row=r, column=3).font = data_font
        ws.cell(row=r, column=4).value = task['desc']
        ws.cell(row=r, column=4).font = data_font
        ws.cell(row=r, column=4).alignment = wrap_align
        ws.cell(row=r, column=5).value = task['pct']
        ws.cell(row=r, column=5).font = data_font

        for c in (6, 10):
            cell = ws.cell(row=r, column=c)
            cell.value = TODAY
            cell.number_format = 'yyyy/m/d;@'
            cell.font = data_font

        current_g, current_remaining = calc_g(task['human_h'], current_g, current_remaining)
        g_cell = ws.cell(row=r, column=7)
        g_cell.value = current_g
        g_cell.number_format = 'yyyy/m/d;@'
        g_cell.font = data_font

        ws.cell(row=r, column=8).value = task['human_h']
        ws.cell(row=r, column=8).font = data_font
        ws.cell(row=r, column=11).value = task['ai_h']
        ws.cell(row=r, column=11).font = data_font
        ws.cell(row=r, column=14).value = task['note']
        ws.cell(row=r, column=14).font = data_font
        ws.cell(row=r, column=14).alignment = wrap_align

        lines = max(1, len(task['desc']) / 35)
        ws.row_dimensions[r].height = max(30, lines * 15)

        print(f"  新增行 {r}: [{display}] #{seq} G={current_g.strftime('%m-%d')} H={task['human_h']}h K={task['ai_h']}h")
        insert_pos += 1

    # A 列合并当天行
    if len(tasks) > 1:
        ws.merge_cells(start_row=2, start_column=1, end_row=insert_pos-1, end_column=1)
    for r in range(2, insert_pos):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='center')

    # 列宽
    cw = {'A':13,'B':23,'C':15,'D':60,'E':16,'F':16,'G':16,'H':14,'I':24,'J':13,'K':15,'L':12,'M':42,'N':60}
    for k, v in cw.items(): ws.column_dimensions[k].width = v
    ws.freeze_panes = 'A2'
    wb.save(XLSX_FILE)
    print(f"\n已保存: {XLSX_FILE}")

if __name__ == '__main__':
    main()
