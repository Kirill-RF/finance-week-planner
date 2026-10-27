# -*- coding: utf-8 -*-
"""
Сборка файла "План на фин неделю.xlsx" (фин. неделя: четверг–среда).
Редактируемые параметры: YEAR, MANAGERS, BASE_PLAN. Запуск: python3 build_workbook.py
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

YEAR = 2026
MONTHS_RU = ["Январь","Февраль","Март","Апрель","Май","Июнь",
             "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
WD_RU = ["Понедельник","Вторник","Среда","Четверг","Пятница","Суббота","Воскресенье"]
MANAGERS = ["Иванов","Петрова","Сидоров","Козлова","Новиков","Морозова",
            "Волков","Алексеева","Лебедев","Семенова","Егоров","Павлова"]
BASE_PLAN = [5_000_000, 4_800_000, 5_200_000, 5_500_000, 5_300_000, 6_000_000,
             6_200_000, 6_000_000, 5_800_000, 6_500_000, 7_000_000, 8_000_000]
OUT = "План на фин неделю.xlsx"

# ---------- границы финансовых недель (чт..ср) ----------
jan1 = dt.date(YEAR,1,1)
start = jan1 - dt.timedelta(days=(jan1.weekday()-3) % 7)   # первый четверг года или предыдущий
dec31 = dt.date(YEAR,12,31)
end = dec31 + dt.timedelta(days=(2-dec31.weekday()) % 7)   # ближайшая среда >= 31.12 (полное покрытие года)
weeks=[]; s=start
while s <= end:
    e=s+dt.timedelta(days=6); weeks.append((s,e)); s=e+dt.timedelta(days=1)
ND=len(weeks)*7                                  # дней в календаре (полные недели)
WEEKS_PER_MONTH={m: sum(1 for a,b in weeks if a.month==m) for m in range(1,13)}

# ---------- именованные диапазоны (надёжно работают в Excel/LibreOffice/Google Sheets) ----------
from openpyxl.workbook.defined_name import DefinedName

# ================= Лист "Рабочие дни" =================
F_TITLE=Font(bold=True,size=14); F_HDR=Font(bold=True,color="FFFFFF"); F_BOLD=Font(bold=True)
F_NOTE=Font(italic=True,color="808080")
FILL_HDR=PatternFill("solid",fgColor="4472C4"); FILL_IN=PatternFill("solid",fgColor="FFF2CC")
FILL_CALC=PatternFill("solid",fgColor="E2EFDA"); FILL_ALT=PatternFill("solid",fgColor="D9E1F2")
thin=Side(style="thin",color="999999"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
CENTER=Alignment(horizontal="center",vertical="center",wrap_text=True)
def st(c,font=None,fill=None,align=None,border=True,fmt=None):
    if font:c.font=font
    if fill:c.fill=fill
    if align:c.alignment=align
    if border:c.border=BORDER
    if fmt:c.number_format=fmt

wb=Workbook()

# ================= Лист "Рабочие дни" =================
wsR=wb.active; wsR.title="Рабочие дни"
wsR["A1"]="РАБОЧИЕ ДНИ И ФИНАНСОВЫЕ НЕДЕЛИ"; wsR["A1"].font=F_TITLE
wsR["A2"]="ЖЁЛТЫЕ ячейки — ввод вручную. Остальное считается формулами."; wsR["A2"].font=F_NOTE
st(wsR.cell(4,1,"Месяц"),F_HDR,FILL_HDR,CENTER)
st(wsR.cell(4,2,"Рабочих дней в месяце (ВАША СТРОКА ВВОДА)"),F_HDR,FILL_HDR,CENTER)
st(wsR.cell(4,3,"Недель в месяце"),F_HDR,FILL_HDR,CENTER)
MON2ROW={}
for i,m in enumerate(range(1,13),start=5):
    MON2ROW[m]=i
    st(wsR.cell(i,1,f"{MONTHS_RU[m-1]} {YEAR}"))
    c=wsR.cell(i,2,None); st(c,fill=FILL_IN,align=CENTER)      # <-- ручной ввод
    st(wsR.cell(i,3,WEEKS_PER_MONTH[m]),align=CENTER)
wsR["E5"]="По умолчанию в «Календаре» рабочими считаются пн–пт. Праздники вычитайте из числа рабочих дней вручную."
wsR["E5"].font=F_NOTE

HR=19
wsR.cell(HR-1,1,"Границы финансовых недель (четверг — среда)").font=F_BOLD
for j,h in enumerate(["№ фин. недели","Начало (чт)","Конец (ср)","Месяц открытия","Рабочих дней в неделе"],1):
    st(wsR.cell(HR,j,h),F_HDR,FILL_HDR,CENTER)
WK_FIRST=HR+1
for k,(a,b) in enumerate(weeks,1):
    r=WK_FIRST+k-1
    st(wsR.cell(r,1,k),align=CENTER)
    wsR.cell(r,2,a); st(wsR.cell(r,2),fmt="dd.mm.yyyy",align=CENTER)
    wsR.cell(r,3,b); st(wsR.cell(r,3),fmt="dd.mm.yyyy",align=CENTER)
    mm = a.month if a.year == YEAR else (a + dt.timedelta(days=3)).month  # хвостовая неделя за годом -> месяц её рабочих дней (декабрь)
    st(wsR.cell(r,4,f"{MONTHS_RU[mm-1]} {YEAR}"),align=CENTER)
    st(wsR.cell(r,5,None),fill=FILL_CALC,align=CENTER)
WK_LAST=WK_FIRST+len(weeks)-1
wsR.column_dimensions["A"].width=16
for col,w in zip("BCDE",[13,13,17,22]): wsR.column_dimensions[col].width=w

# ================= Лист "Календарь" =================
CAL_FIRST=2
wsC=wb.create_sheet("Календарь")
for j,h in enumerate(["Дата","Месяц №","День недели","Рабочий день (1/0)"],1):
    st(wsC.cell(1,j,h),F_HDR,FILL_HDR,CENTER)
d=start; r=CAL_FIRST
while d<=dt.date(YEAR,12,31):
    wsC.cell(r,1,d).number_format="dd.mm.yyyy"
    wsC.cell(r,2,d.month); wsC.cell(r,3,WD_RU[d.weekday()])
    wsC.cell(r,4,1 if d.weekday()<5 else 0)
    if d.weekday()>=5:
        for j in range(1,5): wsC.cell(r,j).fill=PatternFill("solid",fgColor="FCE4E4")
    d+=dt.timedelta(days=1); r+=1
CAL_LAST=r-1
for col,w in zip("ABCD",[13,9,15,18]): wsC.column_dimensions[col].width=w
wsC.freeze_panes="A2"

# формулы листа "Рабочие дни" (через именованные диапазоны)
for m,i in MON2ROW.items():
    wsR.cell(i,2).value=f"=SUMPRODUCT((Кал_Месяц={m})*(Кал_Раб=1))"
for k in range(len(weeks)):
    rw=WK_FIRST+k
    wsR.cell(rw,5).value=f"=SUMPRODUCT((Кал_Даты>=$B{rw})*(Кал_Даты<=$C{rw})*(Кал_Раб=1))"

# ================= Лист "Планы менеджеров" =================
wsP=wb.create_sheet("Планы менеджеров",0)
wsP["A1"]="МЕСЯЧНЫЕ ПЛАНЫ МЕНЕДЖЕРОВ"; wsP["A1"].font=F_TITLE
wsP["A2"]=("Редактируйте жёлтые ячейки: план на календарный месяц для каждого менеджера. "
           "Распределение по фин. неделям и стыки месяцев считаются автоматически.")
wsP["A2"].font=F_NOTE
st(wsP.cell(4,1,"Менеджер"),F_HDR,FILL_HDR,CENTER)
for j,m in enumerate(range(1,13)):
    st(wsP.cell(4,2+j,f"{MONTHS_RU[m-1]} {YEAR}"),F_HDR,FILL_HDR,CENTER)
MGR_ROW={}
for i,name in enumerate(MANAGERS):
    r=5+i; MGR_ROW[name]=r
    st(wsP.cell(r,1,name),F_BOLD)
    for j,p in enumerate(BASE_PLAN):
        st(wsP.cell(r,2+j,p),fill=FILL_IN,fmt="#,##0")
RT=5+len(MANAGERS)
st(wsP.cell(RT,1,"ИТОГО"),F_BOLD,FILL_ALT)
for j in range(12):
    L=get_column_letter(2+j)
    st(wsP.cell(RT,2+j,f"=SUM({L}5:{L}{RT-1})"),F_BOLD,FILL_ALT,fmt="#,##0")
wsP.column_dimensions["A"].width=18
for j in range(2,14): wsP.column_dimensions[get_column_letter(j)].width=13
wsP.freeze_panes="B5"

# ================= Лист "Расчёт" =================
wsX=wb.create_sheet("Расчёт")
wsX["A1"]="СЛУЖЕБНЫЙ РАСЧЁТ (не редактировать)"; wsX["A1"].font=F_TITLE
wsX["A2"]=("Блок 1: план дня = план месяца / число рабочих дней месяца (выходные = 0). "
           "Блок 2: план недели = сумма планов её дней; дни недели, попавшие на разные месяцы, "
           "считаются по плану каждого месяца (стык месяцев учтён).")
wsX["A2"].font=F_NOTE
H1=4
for j,h in enumerate(["Менеджер","Строка плана","Дата","Месяц даты","Рабочий?","№ фин. недели","План на день"],1):
    st(wsX.cell(H1,j,h),F_HDR,FILL_HDR,CENTER)
DAY_FIRST=H1+1
rp=DAY_FIRST
for name in MANAGERS:
    mr=MGR_ROW[name]
    for idx in range(ND):
        r=rp+idx; cr=CAL_FIRST+idx
        wsX.cell(r,1,name)
        wsX.cell(r,2,mr)
        wsX.cell(r,3,f"='Календарь'!A{cr}").number_format="dd.mm.yyyy"
        wsX.cell(r,4,f"=MONTH(C{r})")
        wsX.cell(r,5,f"='Календарь'!D{cr}")
        wsX.cell(r,6,f"=INT((C{r}-Нед_Начало)/7)+1")
        # План на день = план месяца этой даты / число рабочих дней месяца; в выходной = 0
        wsX.cell(r,7,
            f"=IF(E{r}=0,0,INDEX('Планы менеджеров'!$B$5:$M${RT-1},B{r},D{r})"
            f"/INDEX(Мес_Раб,D{r}))")
    rp+=ND
DAY_LAST=rp-1

H2=DAY_LAST+3
wsX.cell(H2-1,1,"Блок 2. План по финансовым неделям").font=F_BOLD
for j,h in enumerate(["Менеджер","№ недели","Рабочих дней в неделе","План недели (итого)",
                      "Из них план мес. A","Из них план мес. B (стык)","Проверка A+B=итого"],1):
    st(wsX.cell(H2,j,h),F_HDR,FILL_HDR,CENTER)
WK_TBL=H2+1
W2ROWS={}; rp=WK_TBL
for name in MANAGERS:
    first=rp
    for k in range(1,len(weeks)+1):
        r=rp+k-1
        df=DAY_FIRST+(k-1)*7; dl=df+6
        wsX.cell(r,1,name)
        wsX.cell(r,2,k)
        wsX.cell(r,3,f"=INDEX(Нед_Раб,{k})")
        wsX.cell(r,4,f"=SUM(G{df}:G{dl})")
        wsX.cell(r,5,f"=SUMIFS($G${df}:$G${dl},$D${df}:$D${dl},INDEX($D${df}:$D${dl},1))")
        wsX.cell(r,6,f"=SUMIFS($G${df}:$G${dl},$D${df}:$D${dl},INDEX($D${df}:$D${dl},7))")
        wsX.cell(r,7,f"=IF(ROUND(E{r}+F{r}-D{r},2)=0,\"OK\",\"ОШИБКА\")")
    W2ROWS[name]=(first,rp+len(weeks)-1)
    rp+=len(weeks)
WK2_LAST=rp-1
for r in range(WK_TBL,WK2_LAST+1):
    for j in range(1,8):
        c=wsX.cell(r,j); st(c,align=None if j==1 else CENTER)
        if j in (4,5,6): c.number_format="#,##0"
        if j in (3,4,5,6,7): c.fill=FILL_CALC
wsX.column_dimensions["A"].width=16
for col,w in zip("BCDEFG",[10,12,15,14,16,14]): wsX.column_dimensions[col].width=w
wsX.freeze_panes=f"A{DAY_FIRST}"

# ---------- именованные диапазоны (используются в формулах вместо прямых ссылок) ----------
for nm, ref in {
    "Кал_Даты":   f"Календарь!$A${CAL_FIRST}:$A${CAL_LAST}",
    "Кал_Месяц":  f"Календарь!$B${CAL_FIRST}:$B${CAL_LAST}",
    "Кал_Раб":    f"Календарь!$D${CAL_FIRST}:$D${CAL_LAST}",
    "Нед_Начало": f"'Рабочие дни'!$B${WK_FIRST}:$B${WK_LAST}",
    "Нед_Конец":  f"'Рабочие дни'!$C${WK_FIRST}:$C${WK_LAST}",
    "Нед_Раб":    f"'Рабочие дни'!$E${WK_FIRST}:$E${WK_LAST}",
    "Мес_Раб":    f"'Рабочие дни'!$B$5:$B$16",
}.items():
    wb.defined_names.add(DefinedName(nm, attr_text=ref))

# ================= Лист "План на фин. неделю" =================
wsW=wb.create_sheet("План на фин. неделю",0)
wsW["A1"]="ПЛАН НА ФИНАНСОВУЮ НЕДЕЛЮ (четверг — среда)"; wsW["A1"].font=F_TITLE
wsW["A2"]="Сотрудник находит свою фамилию и активную неделю — это его план на неделю."
wsW["A2"].font=F_NOTE
st(wsW.cell(4,1,"Активная дата:"),F_BOLD,border=False)
st(wsW.cell(4,2,dt.date(YEAR,10,22)),fill=FILL_IN,fmt="dd.mm.yyyy",align=CENTER)
wsW.cell(4,3,"← поставьте любую дату (хоть =СЕГОДНЯ()). Активная неделя подсветится в таблице и появится справа.").font=F_NOTE

wsW["A6"]="ТАБЛИЦА 1. Все недели года"; wsW["A6"].font=F_BOLD
st(wsW.cell(7,1,"№ недели"),F_HDR,FILL_HDR,CENTER)
st(wsW.cell(8,1,"Период"),F_HDR,FILL_HDR,CENTER)
st(wsW.cell(9,1,"Рабочих дней"),F_HDR,FILL_HDR,CENTER)
for k,(a,b) in enumerate(weeks,1):
    L=get_column_letter(1+k)
    st(wsW.cell(7,1+k,k),align=CENTER)
    st(wsW.cell(8,1+k,f"{a.day:02d}.{a.month:02d} — {b.day:02d}.{b.month:02d}"),align=CENTER)
    st(wsW.cell(9,1+k,f"=INDEX(Нед_Раб,{k})"),align=CENTER,fill=FILL_CALC)
    wsW.column_dimensions[L].width=13
wsW.column_dimensions["A"].width=16
for i,name in enumerate(MANAGERS):
    r=10+i
    st(wsW.cell(r,1,name),F_BOLD)
    f,l=W2ROWS[name]
    for k in range(1,len(weeks)+1):
        st(wsW.cell(r,1+k,f"='Расчёт'!D{f+k-1}"),fmt="#,##0",align=CENTER)
TOTR=10+len(MANAGERS)
st(wsW.cell(TOTR,1,"ИТОГО"),F_BOLD,FILL_ALT)
for k in range(1,len(weeks)+1):
    L=get_column_letter(1+k)
    st(wsW.cell(TOTR,1+k,f"=SUM({L}10:{L}{TOTR-1})"),F_BOLD,FILL_ALT,fmt="#,##0",align=CENTER)
wsW.freeze_panes="B10"

base=len(weeks)+3          # пустой столбец-разделитель +1
LB=get_column_letter(base); LP=get_column_letter(base+1)
wsW.cell(6,base,"ТАБЛИЦА 2. Активная неделя").font=F_BOLD
st(wsW.cell(7,base,"№ активной недели"),F_HDR,FILL_HDR,CENTER)
st(wsW.cell(7,base+1,"Период"),F_HDR,FILL_HDR,CENTER)
st(wsW.cell(8,base,f"=IFERROR(MATCH($B$4,Нед_Начало,1),MATCH($B$4,Нед_Конец,1))"),align=CENTER,fill=FILL_CALC)
st(wsW.cell(8,base+1,f'=TEXT(INDEX(Нед_Начало,${LB}$8),"dd.mm.yyyy")&" — "&TEXT(INDEX(Нед_Конец,${LB}$8),"dd.mm.yyyy")'),
   align=CENTER,fill=FILL_CALC)
st(wsW.cell(10,base,"Менеджер"),F_HDR,FILL_HDR,CENTER)
st(wsW.cell(10,base+1,"План на активную неделю"),F_HDR,FILL_HDR,CENTER)
for i,name in enumerate(MANAGERS):
    r=11+i
    st(wsW.cell(r,base,name),F_BOLD)
    f,l=W2ROWS[name]
    st(wsW.cell(r,base+1,f"=INDEX('Расчёт'!$D${f}:$D${l},${LB}$8)"),fmt="#,##0",align=CENTER,fill=FILL_CALC)
wsW.column_dimensions[LB].width=18; wsW.column_dimensions[LP].width=22
wsW.column_dimensions[get_column_letter(base-1)].width=3

lastL=get_column_letter(1+len(weeks))
wsW.conditional_formatting.add(f"B7:{lastL}{TOTR}",
    FormulaRule(formula=[f"B$7=${LB}$8"],fill=PatternFill("solid",bgColor="FFE699")))

wb.save(OUT)
print("saved:",OUT,"| недель:",len(weeks),"| дней календаря:",ND)
