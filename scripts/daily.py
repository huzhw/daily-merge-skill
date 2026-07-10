# -*- coding: utf-8 -*-
"""每日脚本：读取昨天 Excel，追加当天已完成需求"""
import openpyxl, os, shutil
from datetime import datetime, timedelta
from copy import copy
from openpyxl.styles import Alignment

DESKTOP = r"C:\Users\Administrator\Desktop"
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
YESTERDAY = TODAY - timedelta(days=1)
YEAR = TODAY.strftime("%Y")
MONTH = str(TODAY.month)
MM = TODAY.strftime("%m")
DD = TODAY.strftime("%d")
Y_MM = YESTERDAY.strftime("%m")
Y_DD = YESTERDAY.strftime("%d")

REPO_MAP = {
    'lanxum-amisp': '档案V6', 'lanxum-amisp-java': '档案V6', 'lanxum-amisp-react': '档案V6',
    'workingpaper-v5.5': '中信底稿v5',
}

REPORT_DIR = os.path.join(DESKTOP, f"报告-{YEAR}年", f"日报-{YEAR}-{MONTH}月")
MD_FILE = os.path.join(REPORT_DIR, f"日报需求记录-{YEAR}-{MM}-{DD}.md")
XLSX_FILE = os.path.join(REPORT_DIR, f"日报表格-胡志伟~~{MM}-{DD}.xlsx")
YESTERDAY_XLSX = os.path.join(REPORT_DIR, f"日报表格-胡志伟~~{Y_MM}-{Y_DD}.xlsx")

def next_workday(d):
    d = d + timedelta(days=1)
    while d.weekday() >= 5:
        d = d + timedelta(days=1)
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

def get_last_info(ws):
    last_g = TODAY
    remaining = 0
    last_row = 0
    for row in range(1, ws.max_row + 1):
        a = ws.cell(row=row, column=1).value
        if a and '填充说明' in str(a): break
        c = ws.cell(row=row, column=3).value
        if c is not None and str(c).strip().isdigit():
            last_row = row
    if last_row > 0:
        gv = ws.cell(row=last_row, column=7).value
        if gv:
            if isinstance(gv, datetime): last_g = gv
            else:
                try: last_g = datetime.strptime(str(gv)[:10], '%Y-%m-%d')
                except: pass
        day_hours = 0
        for row in range(last_row, 0, -1):
            rg = ws.cell(row=row, column=7).value
            rd = None
            if rg:
                if isinstance(rg, datetime): rd = rg.date()
                else:
                    try: rd = datetime.strptime(str(rg)[:10], '%Y-%m-%d').date()
                    except: pass
            if rd != last_g.date(): break
            hv = ws.cell(row=row, column=8).value
            if hv:
                try: day_hours += float(hv)
                except: pass
        remaining = day_hours % 8
    return last_g, remaining, last_row

def find_notes(ws):
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=1).value and '填充说明' in str(ws.cell(row=row, column=1).value):
            return row
    return ws.max_row + 3

def calc_g(hours, prev_date, remaining):
    total = remaining + hours
    d = prev_date
    while total > 8:
        total -= 8
        d = next_workday(d)
    return d, total

def copy_style(src, dst):
    if src.has_style:
        # 复制字号字体但不复制颜色（表头是白色）
        from openpyxl.styles import Font
        sf = src.font
        dst.font = Font(name=sf.name, size=sf.size, bold=sf.bold, italic=sf.italic)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)

def remove_today_rows(ws):
    today_merged = set()
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col == 1 and mr.max_col == 1:
            tv = ws.cell(row=mr.min_row, column=1).value
            if tv:
                if isinstance(tv, datetime): ad = tv.date()
                else:
                    try: ad = datetime.strptime(str(tv)[:10], '%Y-%m-%d').date()
                    except: continue
                if ad == TODAY.date():
                    for r in range(mr.min_row, mr.max_row + 1): today_merged.add(r)
                    try: ws.unmerge_cells(str(mr))
                    except: pass
    rows_to_del = set()
    notes = find_notes(ws)
    for row in range(1, notes):
        av = ws.cell(row=row, column=1).value
        if av:
            if isinstance(av, datetime): ad = av.date()
            else:
                try: ad = datetime.strptime(str(av)[:10], '%Y-%m-%d').date()
                except: continue
            if ad == TODAY.date(): rows_to_del.add(row)
    rows_to_del |= today_merged
    for row in sorted(rows_to_del, reverse=True):
        ws.delete_rows(row)
    return len(rows_to_del)

def main():
    tasks = parse_md(MD_FILE)
    if not tasks:
        print("没有已完成的任务，跳过。")
        return
    print(f"找到 {len(tasks)} 条已完成任务")

    # 找昨天的 Excel
    if not os.path.exists(YESTERDAY_XLSX):
        print(f"错误：找不到昨天的文件 {YESTERDAY_XLSX}")
        return
    shutil.copy(YESTERDAY_XLSX, XLSX_FILE)

    wb = openpyxl.load_workbook(XLSX_FILE)
    ws = wb[wb.sheetnames[0]]

    # 去重
    removed = remove_today_rows(ws)
    if removed:
        print(f"已清理当天 {removed} 行旧数据")

    last_g, remaining, last_data_row = get_last_info(ws)
    print(f"上一个任务 G={last_g.strftime('%Y-%m-%d')}, 当天剩余 {remaining}h")

    # 找插入位置
    notes_row = find_notes(ws)
    gap = notes_row - last_data_row - 1
    if len(tasks) > gap:
        need = len(tasks) - gap
        for _ in range(need):
            ws.insert_rows(notes_row)
        print(f"空行不足，补插 {need} 行")

    insert_pos = last_data_row + 1
    current_g = last_g
    current_remaining = remaining
    prev_repo = None
    max_seq = 0
    for row in range(1, insert_pos):
        cv = ws.cell(row=row, column=3).value
        if cv is not None and str(cv).strip().isdigit():
            max_seq = max(max_seq, int(str(cv).strip()))

    for row in range(last_data_row, 0, -1):
        bv = ws.cell(row=row, column=2).value
        if bv: prev_repo = str(bv).strip(); break

    for task in tasks:
        ws.insert_rows(insert_pos)
        ws.cell(row=insert_pos, column=1).value = TODAY
        ws.cell(row=insert_pos, column=1).number_format = 'yyyy/m/d;@'

        repo = task['repo']
        display = REPO_MAP.get(repo, repo)
        if repo != prev_repo:
            ws.cell(row=insert_pos, column=2).value = display
            prev_repo = repo

        max_seq += 1
        ws.cell(row=insert_pos, column=3).value = max_seq
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

        for c in range(1, 15):
            copy_style(ws.cell(row=1, column=c), ws.cell(row=insert_pos, column=c))
        for c in (4, 14):
            ws.cell(row=insert_pos, column=c).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

        lines = max(1, len(task['desc']) / 35)
        ws.row_dimensions[insert_pos].height = max(30, lines * 15)

        print(f"  新增行 {insert_pos}: [{display}] #{max_seq} G={current_g.strftime('%m-%d')} H={task['human_h']}h K={task['ai_h']}h")
        insert_pos += 1

    # A 列合并当天行
    start = insert_pos - len(tasks)
    end = insert_pos - 1
    if len(tasks) > 1:
        ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
    for r in range(start, end + 1):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='center')

    # 数据与备注间距 ≥2
    notes_now = find_notes(ws)
    g = notes_now - insert_pos
    if g < 2:
        need = 2 - g
        for _ in range(need):
            ws.insert_rows(notes_now)
        print(f"补空行 {need} 行")

    # 列宽
    cw = {'A':13,'B':23,'C':15,'D':60,'E':16,'F':18,'G':18,'H':14,'I':32,'J':13,'K':15,'L':14,'M':25,'N':60}
    for k, v in cw.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = 'A2'

    wb.save(XLSX_FILE)
    print(f"\n已保存: {XLSX_FILE}")

if __name__ == '__main__':
    main()
