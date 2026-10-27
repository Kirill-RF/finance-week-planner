# -*- coding: utf-8 -*-
"""Независимая проверка значений из values.json (результат расчёта модели Excel)."""
import json, datetime as dt, re

YEAR=2026
MANAGERS=["Иванов","Петрова","Сидоров","Козлова","Новиков","Морозова",
          "Волков","Алексеева","Лебедев","Семенова","Егоров","Павлова"]
BASE_PLAN=[5_000_000,4_800_000,5_200_000,5_500_000,5_300_000,6_000_000,
           6_200_000,6_000_000,5_800_000,6_500_000,7_000_000,8_000_000]

vals=json.load(open("values.json",encoding="utf-8"))
def get(sheet,cell):
    for k,v in vals.items():
        if f"'[ПЛАН НА ФИН НЕДЕЛЮ.XLSX]{sheet.upper()}'!{cell}"==k.upper().replace('"',""):
            return v
    return None

def parse_num(s):
    s=s.strip()
    m=re.search(r"=\[\[([^\]]+)\]\]",s)   # формат "<Ranges>(...=[[22.0]])"
    if m: s=m.group(1).split(",")[0].strip()
    else:
        m=re.match(r"\[\[([^\]]+)\]\]",s)
        if m: s=m.group(1).split(",")[0].strip()
    try: return float(s)
    except: return s

# --- эталонное значение (python) ---
jan1=dt.date(YEAR,1,1); start=jan1-dt.timedelta(days=(jan1.weekday()-3)%7)
weeks=[]; s=start
while s<=dt.date(YEAR,12,31):
    e=s+dt.timedelta(days=6); weeks.append((s,e)); s=e+dt.timedelta(days=1)
wdm={m:sum(1 for a,b in weeks if a.month==m and False) for m in range(1,13)}  # не нужно
# рабочие дни месяца по календарю пн-пт в диапазоне недель
cal=[]; d=start
while d<=dt.date(YEAR,12,31):
    cal.append(d); d+=dt.timedelta(days=1)
MONTH_WD={m: sum(1 for x in cal if x.month==m and x.weekday()<5) for m in range(1,13)}

def day_plan(date, plan_month):  # план дня для даты в месяце m
    m=date.month
    if date.weekday()>=5: return 0.0
    return BASE_PLAN[m-1]/MONTH_WD[m]

def week_plan(k):
    a,b=weeks[k-1]; tot=0.0
    d=a
    while d<=b:
        tot+=day_plan(d,None); d+=dt.timedelta(days=1)
    return tot

errors=[]; checked=0
# 1) рабочие дни месяца: лист "Рабочие дни" B5:B16
for i,m in enumerate(range(1,13)):
    v=get("РАБОЧИЕ ДНИ",f"B{5+i}")
    if v is None: errors.append(f"нет значения Рабочие дни!B{5+i}"); continue
    fv=float(parse_num(v))
    exp=MONTH_WD[m]
    checked+=1
    if abs(fv-exp)>1e-9: errors.append(f"Раб.дни {m}: файл={fv}, ожидание={exp}")

# 2) рабочие дни недели: E20..E72
for k,(a,b) in enumerate(weeks,1):
    v=get("РАБОЧИЕ ДНИ",f"E{20+k-1}")
    fv=float(parse_num(v))
    exp=sum(1 for j in range(7) if (a+dt.timedelta(days=j)).weekday()<5)
    checked+=1
    if abs(fv-exp)>1e-9: errors.append(f"Раб.дни недели {k}: файл={fv}, ожидание={exp}")

# 3) план недели Иванов (строки Расчёт блок2) — первые 5 и стыковые недели
# блок2 начинается с H2 = DAY_LAST+3 +1; посчитаем строки как в генераторе
ND=len(weeks)*7; DAY_FIRST=5
DAY_LAST=DAY_FIRST+ND-1
H2=DAY_LAST+3; WK_TBL=H2+1
ivanov_first=WK_TBL
test_weeks=[1,2,44,45,46,52,53]  # начало года, стык янв/фев, окт/ноя, конец года
for k in test_weeks:
    v=get("РАСЧЁТ",f"D{ivanov_first+k-1}")
    if v is None: errors.append(f"нет Расчёт!D{ivanov_first+k-1}"); continue
    fv=float(parse_num(v)); exp=week_plan(k)
    checked+=1
    if abs(fv-exp)>max(1e-6,abs(exp)*1e-9): errors.append(f"Иванов нед {k}: файл={fv:.2f}, ожидание={exp:.2f}")

# 4) стык месяцев: неделя, где чт=30.10.2026? найдём неделю со сменой месяца внутри
cross=None
for k,(a,b) in enumerate(weeks,1):
    ms={x.month for x in [a+dt.timedelta(days=j) for j in range(7)]}
    if len(ms)>1: cross=k; break
if cross:
    fa=get("РАСЧЁТ",f"E{ivanov_first+cross-1}")  # мес A
    fb=get("РАСЧЁТ",f"F{ivanov_first+cross-1}")  # мес B
    chk=get("РАСЧЁТ",f"G{ivanov_first+cross-1}") # OK?
    va=float(parse_num(fa)); vb=float(parse_num(fb))
    a,b=weeks[cross-1]
    daysA=[a+dt.timedelta(days=j) for j in range(7) if (a+dt.timedelta(days=j)).month==a.month]
    daysB=[a+dt.timedelta(days=j) for j in range(7) if (a+dt.timedelta(days=j)).month!=a.month]
    expA=sum(day_plan(x,None) for x in daysA); expB=sum(day_plan(x,None) for x in daysB)
    checked+=3
    if abs(va-expA)>1e-6: errors.append(f"стык A: {va} vs {expA}")
    if abs(vb-expB)>1e-6: errors.append(f"стык B: {vb} vs {expB}")
    print("Стыковая неделя №%d (%s–%s): часть A=%.2f, часть B=%.2f, проверка=%s"%(cross,a,b,va,vb,str(chk)))

# 5) активная дата 22.10.2026 -> номер недели и планы в Таблице 2
act_date=dt.date(2026,10,22)
exp_k=next(k for k,(a,b) in enumerate(weeks,1) if a<=act_date<=b)
base=len(weeks)+3  # 56 -> 'BD'
from openpyxl.utils import get_column_letter
LB=get_column_letter(base)
v=get("ПЛАН НА ФИН. НЕДЕЛЮ",f"{LB}8")
print("Активная неделя: файл=",str(v)," ожидание=",exp_k)
checked+=1
if abs(float(parse_num(v))-exp_k)>1e-9: errors.append("номер активной недели неверный")
LP=get_column_letter(base+1)
for i,name in enumerate(MANAGERS):
    v=get("ПЛАН НА ФИН. НЕДЕЛЮ",f"{LP}{11+i}")
    fv=float(parse_num(v))
    # план того же менеджера на ту же неделю из блока 2 Расчёта
    first=WK_TBL+i*len(weeks)
    wv=float(parse_num(get("РАСЧЁТ",f"D{first+exp_k-1}")))
    checked+=1
    if abs(fv-wv)>1e-6: errors.append(f"Табл2 {name}: {fv} vs {wv}")

# 6) ИТОГО за год по менеджеру = сумма его месячных планов (с точностью до дней вне месяцев)
tot_expected_year=sum(BASE_PLAN)  # для Иванова при полном годе в календаре
# проверим сумму строки Иванова по всем неделям Таблицы 1
sm=0.0
for k in range(1,len(weeks)+1):
    L=get_column_letter(1+k)
    sm+=float(parse_num(get("ПЛАН НА ФИН. НЕДЕЛЮ",f"{L}10")))
checked+=1
print("Сумма плана Иванова по всем неделям: %.2f (годовой план 5+4.8+...+8 млн = %d)"%(sm,tot_expected_year))
if abs(sm-tot_expected_year)>1.0: errors.append(f"итог года Иванова: {sm} vs {tot_expected_year}")

# 7) ошибок вида #REF/#NAME/#DIV в расчётных ячейках нет
badcells=[k for k,v in vals.items() if any(e in str(v) for e in ("#REF","#NAME","#DIV","#VALUE","#N/A"))]
if badcells:
    errors.append(f"{len(badcells)} ячеек с ошибками, примеры: {badcells[:5]}")

print(f"\nПроверено значений: {checked}")
if errors:
    print("ОШИБКИ:"); [print(" -",e) for e in errors]
else:
    print("ВСЁ В ПОРЯДКЕ ✔")
