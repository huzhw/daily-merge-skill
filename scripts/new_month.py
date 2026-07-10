# -*- coding: utf-8 -*-
import openpyxl, os, shutil
from datetime import datetime, timedelta

DESKTOP = r"C:\Users\Administrator\Desktop"
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
Y = TODAY.strftime("%Y"); M = str(TODAY.month); MM = TODAY.strftime("%m"); DD = TODAY.strftime("%d")
REPO_MAP = {'lanxum-amisp':'档案V6','lanxum-amisp-java':'档案V6','lanxum-amisp-react':'档案V6','workingpaper-v5.5':'中信底稿v5'}
TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'templates', '日报模板.xlsx')
REPORT_DIR = os.path.join(DESKTOP, f"报告-{Y}年", f"日报-{Y}-{M}月")
MD = os.path.join(REPORT_DIR, f"日报需求记录-{Y}-{MM}-{DD}.md")
XLSX = os.path.join(REPORT_DIR, f"日报表格-胡志伟~~{MM}-{DD}.xlsx")

def nwd(d):
    d=d+timedelta(days=1)
    while d.weekday()>=5:d=d+timedelta(days=1)
    return d

def parse():
    tasks=[]
    with open(MD,encoding='utf-8') as f: c=f.read()
    for l in c.split('\n'):
        l=l.strip()
        if not l.startswith('|') or '---' in l or '序号' in l or '空行' in l: continue
        p=[x.strip() for x in l.split('|')[1:-1]]
        if len(p)<8: continue
        if p[5] in ('已完成','100%') and p[0].isdigit():
            tasks.append({'r':p[2],'d':p[3],'h':float(p[6]),'a':float(p[7]),'n':p[8],'pct':p[5].replace('%','')+'%'})
    return tasks

def main():
    ts=parse()
    if not ts: print('无任务'); return
    os.makedirs(REPORT_DIR,exist_ok=True)
    shutil.copy(TEMPLATE,XLSX)
    wb=openpyxl.load_workbook(XLSX)
    ws=wb[wb.sheetnames[0]]

    row=2; gd=TODAY; gr=0; repo=None; seq=0
    for t in ts:
        ws.cell(row=row,column=1,value=TODAY).number_format='yyyy/m/d;@'
        d=t['r']; dn=REPO_MAP.get(d,d)
        if d!=repo: ws.cell(row=row,column=2,value=dn); repo=d
        seq+=1; ws.cell(row=row,column=3,value=seq)
        ws.cell(row=row,column=4,value=t['d'])
        ws.cell(row=row,column=5,value=t['pct'])
        for c in(6,10): ws.cell(row=row,column=c,value=TODAY).number_format='yyyy/m/d;@'
        gr2=gr+t['h']; gd2=gd
        while gr2>8: gr2-=8; gd2=nwd(gd2)
        ws.cell(row=row,column=7,value=gd2).number_format='yyyy/m/d;@'
        gd=gd2; gr=gr2
        ws.cell(row=row,column=8,value=t['h'])
        ws.cell(row=row,column=11,value=t['a'])
        ws.cell(row=row,column=14,value=t['n'])
        print(f'Row{row}: {dn} #{seq} G={gd2.strftime("%m-%d")}')
        row+=1

    ws.freeze_panes='A2'
    cw={'A':13,'B':23,'C':15,'D':60,'E':16,'F':16,'G':16,'H':14,'I':24,'J':13,'K':15,'L':12,'M':42,'N':60}
    for k,v in cw.items(): ws.column_dimensions[k].width=v
    wb.save(XLSX)
    print(f'Done: {XLSX}')

if __name__=='__main__':main()