# -*- coding: utf-8 -*-
import openpyxl, os
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DESKTOP = r"C:\Users\Administrator\Desktop"
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
Y=TODAY.strftime("%Y");M=str(TODAY.month);MM=TODAY.strftime("%m");DD=TODAY.strftime("%d")
MAP={'lanxum-amisp':'档案V6','lanxum-amisp-java':'档案V6','lanxum-amisp-react':'档案V6','workingpaper-v5.5':'中信底稿v5'}
RD=os.path.join(DESKTOP,f"报告-{Y}年",f"日报-{Y}-{M}月")
MD=os.path.join(RD,f"日报需求记录-{Y}-{MM}-{DD}.md")
XL=os.path.join(RD,f"日报表格-胡志伟~~{MM}-{DD}.xlsx")

def nwd(d):
    d+=timedelta(days=1)
    while d.weekday()>=5:d+=timedelta(days=1)
    return d

def parse():
    ts=[]
    with open(MD,encoding='utf-8') as f:c=f.read()
    for l in c.split('\n'):
        l=l.strip()
        if not l.startswith('|') or '---' in l or '序号' in l or '空行' in l:continue
        p=[x.strip() for x in l.split('|')[1:-1]]
        if len(p)<8:continue
        if p[5] in ('已完成','100%') and p[0].isdigit():
            ts.append({'r':p[2],'d':p[3],'h':float(p[6]),'a':float(p[7]),'n':p[8],'p':p[5].replace('%','')+'%'})
    return ts

def main():
    ts=parse()
    if not ts:print('无任务');return
    os.makedirs(RD,exist_ok=True)

    wb=openpyxl.Workbook()
    ws=wb.active
    ws.title='日报'

    # 表头
    hd=['*日期','*项目名称','*序号\n（全局唯一）','*任务描述','*完成百分比',
        '*任务创建时间\n（录入后勿修改）','*预计完成时间\n（录入后勿修改）',
        '*预计工作人时\n','调整预计完成时间\n（由于优先级问题调整时填写）',
        '实际完成时间','*实际工作人时','插队序号\n若有插队必填','延期/调整原因','备注/说明']
    hfont=Font(name='宋体',size=10,bold=True,color='FFFFFF')
    hfill=PatternFill(start_color='4472C4',end_color='4472C4',fill_type='solid')
    halign=Alignment(horizontal='center',vertical='center',wrap_text=True)
    border=Border(left=Side('thin'),right=Side('thin'),top=Side('thin'),bottom=Side('thin'))

    for i,h in enumerate(hd,1):
        c=ws.cell(row=1,column=i,value=h)
        c.font=hfont;c.fill=hfill;c.alignment=halign;c.border=border
    ws.row_dimensions[1].height=45

    # 备注
    notes=['填充说明：\n填写时例子请删除',
           '1、带*号的字段必填。',
           '2、前一天已完成的任务不需要重新出现在当天的记录中（若之前的任务完成后有bug需要修改，需要新建一条记录，复制之前任务的描述后面增加【bug】字样，并在插队序号中填写之前的任务序号，主序号依然递增）',
           '3、如果遇到更优先级的任务阻断了当前任务，需要新增阻断任务的信息，在插队序号中写入被影响的任务序号，多个任务逗号隔开，并插入/修改被影响任务的【调整预计完成时间】字段',
           '4、工时采用小时单位，按一天8小时计算，其中实际工时每天更新一直到任务完成100%，任务完成的当天做最后一次记录更新后，第二天不用重复写入该记录。',
           '5、此表在一个自然月内持续迭代，下一个自然月后重新起一个新表。',
           '6、备注按需填写，延迟原因需要简要写明。',
           '7、每天日报发送到公邮：kpi-daily-report@lscjz.com；每人每天一封。',
           '8、开会和沟通也当作任务记录时长']
    nr=len(ts)+5  # 数据后空几行
    for i,n in enumerate(notes):
        if i==0:ws.cell(row=nr+i,column=1,value=n)
        else:ws.cell(row=nr+i,column=2,value=n)

    # 写数据
    row=2;gd=TODAY;gr=0;repo=None;seq=0
    dfont=Font(name='宋体',size=10)
    for t in ts:
        ws.cell(row=row,column=1,value=TODAY).number_format='yyyy/m/d;@';ws.cell(row=row,column=1).font=dfont;ws.cell(row=row,column=1).border=border
        d=t['r'];dn=MAP.get(d,d)
        if d!=repo:ws.cell(row=row,column=2,value=dn);repo=d
        ws.cell(row=row,column=2).font=dfont;ws.cell(row=row,column=2).border=border
        seq+=1;ws.cell(row=row,column=3,value=seq);ws.cell(row=row,column=3).font=dfont;ws.cell(row=row,column=3).border=border
        ws.cell(row=row,column=4,value=t['d']);ws.cell(row=row,column=4).font=dfont;ws.cell(row=row,column=4).border=border
        ws.cell(row=row,column=5,value=t['p']);ws.cell(row=row,column=5).font=dfont;ws.cell(row=row,column=5).border=border
        for c in(6,10):
            ws.cell(row=row,column=c,value=TODAY).number_format='yyyy/m/d;@';ws.cell(row=row,column=c).font=dfont;ws.cell(row=row,column=c).border=border
        gr2=gr+t['h'];gd2=gd
        while gr2>8:gr2-=8;gd2=nwd(gd2)
        ws.cell(row=row,column=7,value=gd2).number_format='yyyy/m/d;@';ws.cell(row=row,column=7).font=dfont;ws.cell(row=row,column=7).border=border
        gd=gd2;gr=gr2
        for c,hv in[(8,t['h']),(11,t['a']),(14,t['n'])]:
            ws.cell(row=row,column=c,value=hv);ws.cell(row=row,column=c).font=dfont;ws.cell(row=row,column=c).border=border
        ws.row_dimensions[row].height=max(30,(len(t['d'])/35)*15)
        print(f'Row{row}: {dn} #{seq}')
        row+=1

    if len(ts)>1:
        ws.merge_cells(start_row=2,start_column=1,end_row=row-1,end_column=1)
    cw={'A':13,'B':23,'C':15,'D':60,'E':16,'F':16,'G':16,'H':14,'I':24,'J':13,'K':15,'L':12,'M':42,'N':60}
    for k,v in cw.items():ws.column_dimensions[k].width=v
    ws.freeze_panes='A2'
    wb.save(XL)
    print(f'Done: {XL}')

if __name__=='__main__':main()