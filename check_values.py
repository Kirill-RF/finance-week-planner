# -*- coding: utf-8 -*-
import datetime as dt, ast
FN="План на фин неделю.xlsx"
vals={}
for line in open("values.txt"):
    k,v=line.rstrip("\n").split("\t",1)
    vals[k]=ast.literal_eval(v)
def getv(sheet,coord):
    sheet=sheet.replace('Ё','Е').replace('ё','е')
    return vals.get(f"'[{FN}]{sheet.upper()}'!{coord}".upper())

YEAR=2026
jan1=dt.date(YEAR,1,1)
start=jan1-dt.timedelta(days=(jan1.weekday()-3)%7)
weeks=[]; s=start
while s<=dt.date(YEAR,12,31):
    e=s+dt.timedelta(days=6); weeks.append((s,e)); s=e+dt.timedelta(days=1)
ND=len(weeks)*7
wd_in_month={m:sum(1 for i in range(ND) if (start+dt.timedelta(days=i)).month==m and (start+dt.timedelta(days=i)).weekday()<5) for m in range(1,13)}
plans=[5_000_000,4_800_000,5_200_000,5_500_000,5_300_000,6_000_000,6_200_000,6_000_000,5_800_000,6_500_000,7_000_000,8_000_000]
def expected_week(k):
    a,b=weeks[k-1]; tot=0.0
    for i in range(7):
        d=a+dt.timedelta(days=i)
        if d.weekday()<5: tot+=plans[d.month-1]/wd_in_month[d.month]
    return tot
WK_FIRST=20; DAY_FIRST=5
DAY_LAST=DAY_FIRST+12*ND-1
H2=DAY_LAST+3; WK_TBL=H2+1
fails=[]
# 1 рабочие дни месяца
for m in range(1,13):
    got=getv("Рабочие дни","B%d"%(4+m))
    if got is None or abs(got-wd_in_month[m])>1e-9: fails.append(("wd month",m,got,wd_in_month[m]))
# 2 рабочих дней в неделе
for k in range(1,len(weeks)+1):
    got=getv("Рабочие дни","E%d"%(WK_FIRST+k-1))
    exp=sum(1 for i in range(7) if (weeks[k-1][0]+dt.timedelta(days=i)).weekday()<5)
    if got is None or abs(got-exp)>1e-9: fails.append(("wd week",k,got,exp))
# 3 план недели Иванов (блок2 D) + 4 стык E+F=D, проверка OK
for k in range(1,len(weeks)+1):
    r=WK_TBL+k-1
    d=getv("Расчёт","D%d"%r); e=getv("Расчёт","E%d"%r); f=getv("Расчёт","F%d"%r); chk=getv("Расчёт","G%d"%r)
    exp=expected_week(k)
    if d is None or abs(d-exp)>1.0: fails.append(("week plan",k,d,round(exp,2)))
    if d is None or e is None or f is None or abs(e+f-d)>1.0 or chk!="OK": fails.append(("joint",k,d,e,f,chk))
# 5 лист "План на фин. неделю": таблица1 строка 10 (Иванов) = ожидаемый план
from openpyxl.utils import get_column_letter
for k in range(1,len(weeks)+1):
    L=get_column_letter(1+k)
    got=getv("План на фин. неделю",f"{L}10")
    exp=expected_week(k)
    if got is None or abs(got-exp)>1.0: fails.append(("wsW t1",k,got,round(exp,2)))
base=len(weeks)+3
n_act=getv("План на фин. неделю",f"{get_column_letter(base)}8")
act=1+int((dt.date(2026,10,22)-start).days//7)
if n_act!=act: fails.append(("active num",n_act,act))
p_act=getv("План на фин. неделю",f"{get_column_letter(base+1)}11")
exp=expected_week(act)
if p_act is None or abs(p_act-exp)>1.0: fails.append(("active plan",p_act,round(exp,2)))
per=getv("План на фин. неделю",f"{get_column_letter(base+1)}8")
a,b=weeks[act-1]
expper=f"{a.day:02d}.{a.month:02d} — {b.day:02d}.{b.month:02d}"
if per!=expper: fails.append(("period str",per,expper))
# годовая сумма у Иванова = сумма месячных планов
tot_year=sum(getv("План на фин. неделю",f"{get_column_letter(1+k)}10") for k in range(1,len(weeks)+1))
if abs(tot_year-sum(plans))>100: fails.append(("year sum",tot_year,sum(plans)))
# ИТОГО по менеджерам в таблице1: сумма столбца = блок2 всех менеджеров? проверим только итог недели act: J... строка TOTR
TOTR=10+12
got_tot=getv("План на фин. неделю",f"{get_column_letter(1+act)}{TOTR}")
exp_tot=sum(expected_week(act)*(1.0) for _ in range(1)) # для всех менеджеров планы одинаковы => 12*exp
if got_tot is None or abs(got_tot-12*exp)>12: fails.append(("col total",got_tot,12*exp))
print("active week:",act,"period:",expper)
print("sample expected week1:",round(expected_week(1),2),"| joint week example:")
# найти неделю со стыком и показать цифры
for k in range(1,len(weeks)+1):
    a,b=weeks[k-1]
    if a.month!=b.month:
        r=WK_TBL+k-1
        print(f"  стык: нед {k} {a}->{b}: итого={getv('Расчёт','D%d'%r)}, месA={getv('Расчёт','E%d'%r)}, месB={getv('Расчёт','F%d'%r)}, chk={getv('Расчёт','G%d'%r)}")
        break
print("FAILS:",len(fails))
for x in fails[:15]: print(x)
