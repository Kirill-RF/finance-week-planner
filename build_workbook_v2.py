# -*- coding: utf-8 -*-
"""
Сборка файла "План на фин неделю (обновлённый).xlsx" (фин. неделя: четверг–среда).

Отличия от оригинала (по требованиям):
 1) Расчёт начинается с сентября 2026 и продолжается дальше (сен.2026 ... дек.2027).
 2) Имена менеджеров и месячные планы взяты из файла "План недели - дополнение.xlsx".
 3) ТАБЛИЦА 2 ("Активная неделя") перенесена на вкладку "Планы менеджеров";
    все её формулы работают (исправлены ссылки и обработка дат вне диапазона недель).
 4) Добавлена строка "Рабочих дней месяца": план дня = план месяца / рабочих дней месяца,
    поэтому количество рабочих дней реально участвует в расчёте (в дополнении стояло
    усреднённое 22 — теперь считается по календарю автоматически).

Редактируемые параметры: START_YM, N_MONTHS, PLAN_FILE. Запуск: python3 build_workbook_v2.py
"""
import calendar as _cal
import datetime as dt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

START_YM = (2026, 9)          # первый месяц расчёта — сентябрь 2026
N_MONTHS = 16                 # сентябрь 2026 ... декабрь 2027 (месяцев в плане дополнения)
PLAN_FILE = "План недели - дополнение.xlsx"
OUT = "План на фин неделю (обновлённый).xlsx"

MONTHS_RU = ["Январь","Февраль","Март","Апрель","Май","Июнь",
             "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"]
WD_RU = ["Понедельник","Вторник","Среда","Четверг","Пятница","Суббота","Воскресенье"]

# ---------- список месяцев расчёта ----------
months = []
y, m = START_YM
for _ in range(N_MONTHS):
    months.append((y, m))
    m += 1
    if m == 13:
        y, m = y + 1, 1

# ---------- менеджеры и планы из файла-дополнения ----------
src = load_workbook(PLAN_FILE, data_only=True)["Лист1"]
MANAGERS = []                 # имена менеджеров (как в дополнении)
PLANS = {}                    # name -> {(year, month): план}
for r in range(4, src.max_row + 1):
    name = src.cell(r, 1).value
    if not name:
        continue
    name = str(name).strip()
    plans = {}
    for j, (yy, mm) in enumerate(months):
        v = src.cell(r, 2 + j).value
        if isinstance(v, (int, float)):
            plans[(yy, mm)] = float(v)
    MANAGERS.append(name)
    PLANS[name] = plans
NM = len(MANAGERS)

# ---------- границы финансовых недель (чт..ср), полное покрытие периода ----------
fy, fm = months[0]
ly, lm = months[-1]
first_day = dt.date(fy, fm, 1)
start = first_day - dt.timedelta(days=(first_day.weekday() - 3) % 7)   # первый четверг или предыдущий
last_dom = _cal.monthrange(ly, lm)[1]
end = dt.date(ly, lm, last_dom)
end = end + dt.timedelta(days=(2 - end.weekday()) % 7)                 # ближайшая среда >= конца периода
weeks = []; s = start
while s <= end:
    e = s + dt.timedelta(days=6); weeks.append((s, e)); s = e + dt.timedelta(days=1)
NW = len(weeks)
ND = NW * 7                                                            # дней в календаре (полные недели)

# недель в месяце (справочно; месяц недели = месяц её четверга)
WEEKS_PER_MONTH = {}
for (yy, mm) in months:
    WEEKS_PER_MONTH[(yy, mm)] = sum(1 for a, b in weeks if (a.year, a.month) == (yy, mm))

# ---------- стили ----------
F_TITLE = Font(bold=True, size=14); F_HDR = Font(bold=True, color="FFFFFF"); F_BOLD = Font(bold=True)
F_NOTE = Font(italic=True, color="808080")
FILL_HDR = PatternFill("solid", fgColor="4472C4"); FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_CALC = PatternFill("solid", fgColor="E2EFDA"); FILL_ALT = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="999999"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
def st(c, font=None, fill=None, align=None, border=True, fmt=None):
    if font: c.font = font
    if fill: c.fill = fill
    if align: c.alignment = align
    if border: c.border = BORDER
    if fmt: c.number_format = fmt

wb = Workbook()

# ================= Лист "Рабочие дни" =================
wsR = wb.active; wsR.title = "Рабочие дни"
wsR["A1"] = "РАБОЧИЕ ДНИ И ФИНАНСОВЫЕ НЕДЕЛИ"; wsR["A1"].font = F_TITLE
wsR["A2"] = "ЖЁЛТЫЕ ячейки — ввод вручную. Остальное считается формулами."; wsR["A2"].font = F_NOTE
st(wsR.cell(4, 1, "Месяц"), F_HDR, FILL_HDR, CENTER)
st(wsR.cell(4, 2, "Рабочих дней в месяце (формула; праздники правьте в «Календаре»)"), F_HDR, FILL_HDR, CENTER)
st(wsR.cell(4, 3, "Недель в месяце"), F_HDR, FILL_HDR, CENTER)
st(wsR.cell(4, 6, "Ключ месяца"), F_HDR, FILL_HDR, CENTER)
MON2ROW = {}
# хвостовые месяцы недель на границах периода (например авг.2026 и янв.2028) — план по ним = 0
tail_months = []
for a, b in [weeks[0], weeks[-1]]:
    for d_ in (a, b):
        if (d_.year, d_.month) not in months and (d_.year, d_.month) not in tail_months:
            tail_months.append((d_.year, d_.month))
ALL_KEYS = months + sorted(tail_months)
for i, (yy, mm) in enumerate(ALL_KEYS, start=5):
    MON2ROW[(yy, mm)] = i
    st(wsR.cell(i, 1, f"{MONTHS_RU[mm-1]} {yy}" + (" (хвост недели)" if (yy, mm) not in months else "")))
    st(wsR.cell(i, 2, None), fill=FILL_IN, align=CENTER)               # SUMPRODUCT по ключу месяца
    st(wsR.cell(i, 3, WEEKS_PER_MONTH.get((yy, mm), 0)), align=CENTER)
    st(wsR.cell(i, 6, int(f"{yy}{mm:02d}")), align=CENTER)             # ключ месяца для MATCH
MR_FIRST = 5; MR_LAST = 4 + len(ALL_KEYS)
wsR.cell(MR_LAST + 2, 1,
    "По умолчанию рабочими считаются пн–пт. В «Календаре» обнулите признак «Рабочий день» "
    "для праздничных дней — все планы пересчитаются автоматически.").font = F_NOTE

HR = MR_LAST + 4
wsR.cell(HR - 1, 1, "Границы финансовых недель (четверг — среда)").font = F_BOLD
for j, h in enumerate(["№ фин. недели", "Начало (чт)", "Конец (ср)", "Месяц открытия", "Рабочих дней в неделе"], 1):
    st(wsR.cell(HR, j, h), F_HDR, FILL_HDR, CENTER)
WK_FIRST = HR + 1
for k, (a, b) in enumerate(weeks, 1):
    r = WK_FIRST + k - 1
    st(wsR.cell(r, 1, k), align=CENTER)
    wsR.cell(r, 2, a); st(wsR.cell(r, 2), fmt="dd.mm.yyyy", align=CENTER)
    wsR.cell(r, 3, b); st(wsR.cell(r, 3), fmt="dd.mm.yyyy", align=CENTER)
    st(wsR.cell(r, 4, f"{MONTHS_RU[a.month-1]} {a.year}"), align=CENTER)
    st(wsR.cell(r, 5, None), fill=FILL_CALC, align=CENTER)
WK_LAST = WK_FIRST + NW - 1
wsR.column_dimensions["A"].width = 16
for col, w in zip("BCDEF", [13, 13, 17, 22, 12]): wsR.column_dimensions[col].width = w

# ================= Лист "Календарь" =================
CAL_FIRST = 2
wsC = wb.create_sheet("Календарь")
for j, h in enumerate(["Дата", "Месяц (ключ ГГГГММ)", "День недели", "Рабочий день (1/0)"], 1):
    st(wsC.cell(1, j, h), F_HDR, FILL_HDR, CENTER)
d = start; r = CAL_FIRST
while d <= weeks[-1][1]:
    wsC.cell(r, 1, d).number_format = "dd.mm.yyyy"
    wsC.cell(r, 2, int(f"{d.year}{d.month:02d}"))
    wsC.cell(r, 3, WD_RU[d.weekday()])
    wsC.cell(r, 4, 1 if d.weekday() < 5 else 0)
    if d.weekday() >= 5:
        for j in range(1, 5): wsC.cell(r, j).fill = PatternFill("solid", fgColor="FCE4E4")
    d += dt.timedelta(days=1); r += 1
CAL_LAST = r - 1
for col, w in zip("ABCD", [13, 12, 15, 18]): wsC.column_dimensions[col].width = w
wsC.freeze_panes = "A2"

# формулы листа "Рабочие дни"
for (yy, mm), i in MON2ROW.items():
    wsR.cell(i, 2).value = f"=SUMPRODUCT((Кал_Месяц={int(f'{yy}{mm:02d}')}).*(Кал_Раб=1))"
for k in range(NW):
    rw = WK_FIRST + k
    wsR.cell(rw, 5).value = f"=SUMPRODUCT((Кал_Даты>=$B{rw}).*(Кал_Даты<=$C{rw}).*(Кал_Раб=1))"

# ================= Лист "Планы менеджеров" =================
wsP = wb.create_sheet("Планы менеджеров")
wsP["A1"] = "МЕСЯЧНЫЕ ПЛАНЫ МЕНЕДЖЕРОВ + ТАБЛИЦА 2 «АКТИВНАЯ НЕДЕЛЯ»"; wsP["A1"].font = F_TITLE
wsP["A2"] = ("Жёлтые ячейки — редактируемые: планы менеджеров (из файла «План недели - дополнение»), "
             "рабочие дни месяца, активная дата. ТАБЛИЦА 2 «Активная неделя» перенесена сюда с листа "
             "«План на фин. неделю». Распределение по фин. неделям и стыки месяцев считаются автоматически.")
wsP["A2"].font = F_NOTE
c=wsP.cell(3, 1, "Ключ месяца →"); c.font=F_NOTE
for j, (yy, mm) in enumerate(months):
    c=wsP.cell(3, 2 + j, int(f"{yy}{mm:02d}")); c.font=F_NOTE; c.alignment=CENTER
st(wsP.cell(4, 1, "Менеджер"), F_HDR, FILL_HDR, CENTER)
for j, (yy, mm) in enumerate(months):
    st(wsP.cell(4, 2 + j, f"{MONTHS_RU[mm-1]} {yy}"), F_HDR, FILL_HDR, CENTER)
LAST_PL = get_column_letter(1 + len(months))       # последний столбец месяцев (P при 16 мес.)
st(wsP.cell(5, 1, "Рабочих дней месяца"), F_BOLD, FILL_ALT)
for j, (yy, mm) in enumerate(months):
    st(wsP.cell(5, 2 + j, f"='Рабочие дни'!$B${MON2ROW[(yy, mm)]}"), F_BOLD, FILL_IN, CENTER)
MGR_ROW = {}
for i, name in enumerate(MANAGERS):
    r = 6 + i; MGR_ROW[name] = r
    st(wsP.cell(r, 1, name), F_BOLD)
    for j, (yy, mm) in enumerate(months):
        st(wsP.cell(r, 2 + j, PLANS[name].get((yy, mm))), fill=FILL_IN, fmt="#,##0")
RT = 6 + NM
st(wsP.cell(RT, 1, "ИТОГО"), F_BOLD, FILL_ALT)
for j in range(len(months)):
    L = get_column_letter(2 + j)
    st(wsP.cell(RT, 2 + j, f"=SUM({L}6:{L}{RT-1})"), F_BOLD, FILL_ALT, fmt="#,##0")
wsP.column_dimensions["A"].width = 24
for j in range(2, 2 + len(months)): wsP.column_dimensions[get_column_letter(j)].width = 13
wsP.freeze_panes = "B6"

# ---- ТАБЛИЦА 2. Активная неделя (перенесена на эту вкладку) ----
TB = 2 + len(months) + 1                            # пустой столбец-разделитель перед таблицей
LB = get_column_letter(TB); LP = get_column_letter(TB + 1)
wsP.cell(4, TB, "ТАБЛИЦА 2. АКТИВНАЯ НЕДЕЛЯ").font = F_BOLD
st(wsP.cell(5, TB, "Активная дата:"), F_BOLD, border=False)
# по умолчанию = сегодня; если сегодня вне периода расчёта — последняя известная дата плана
wsP.cell(5, TB + 1).value = dt.datetime(2026, 10, 22)   # активная дата (значение); при желании замените на =СЕГОДНЯ()
wsP.cell(5, TB + 1).number_format = "dd.mm.yyyy"
wsP.cell(5, TB + 2, "← поставьте любую дату (можно =СЕГОДНЯ()). Активная неделя подсветится "
                    "на листе «План на фин. неделю».").font = F_NOTE
ADATE = f"${LP}$5"
st(wsP.cell(7, TB, "№ активной недели"), F_HDR, FILL_HDR, CENTER)
st(wsP.cell(7, TB+1, "Период"), F_HDR, FILL_HDR, CENTER)
# исправлено: корректный поиск номера недели даже для даты раньше 1-й / позже последней недели
st(wsP.cell(8, TB,
    f"=IF({ADATE}<'Рабочие дни'!$B${WK_FIRST},'Рабочие дни'!$A${WK_FIRST},"
    f"IF({ADATE}>'Рабочие дни'!$C${WK_LAST},'Рабочие дни'!$A${WK_LAST},"
    f"MATCH({ADATE},'Рабочие дни'!$B${WK_FIRST}:$B${WK_LAST},1)+1"
    f"-IF(COUNTIF('Рабочие дни'!$B${WK_FIRST}:$B${WK_LAST},{ADATE})>0,0,"
    f"(MATCH({ADATE},'Рабочие дни'!$B${WK_FIRST}:$B${WK_LAST},1)"
    f"-MATCH({ADATE},'Рабочие дни'!$C${WK_FIRST}:$C${WK_LAST},1))))"),
    align=CENTER, fill=FILL_CALC)
st(wsP.cell(8, TB+1,
    f'=TEXT(INDEX(\'Рабочие дни\'!$B${WK_FIRST}:$B${WK_LAST},${LB}$8),"dd.mm.yyyy")&" — "&'
    f'TEXT(INDEX(\'Рабочие дни\'!$C${WK_FIRST}:$C${WK_LAST},${LB}$8),"dd.mm.yyyy")'),
    align=CENTER, fill=FILL_CALC)
st(wsP.cell(10, TB, "Менеджер"), F_HDR, FILL_HDR, CENTER)
st(wsP.cell(10, TB+1, "План на активную неделю"), F_HDR, FILL_HDR, CENTER)
for i, name in enumerate(MANAGERS):
    st(wsP.cell(11 + i, TB, name), F_BOLD)
ACTIVE_FIRST = 11; ACTIVE_LAST = 11 + NM - 1
ACTIVE_TOT = ACTIVE_LAST + 1
wsP.column_dimensions[LB].width = 26; wsP.column_dimensions[LP].width = 22
wsP.column_dimensions[get_column_letter(TB - 1)].width = 3

# ================= Лист "Расчёт" =================
wsX = wb.create_sheet("Расчёт")
wsX["A1"] = "СЛУЖЕБНЫЙ РАСЧЁТ (не редактировать)"; wsX["A1"].font = F_TITLE
wsX["A2"] = ("Блок 1: план дня = план месяца / число рабочих дней месяца (выходные = 0). "
             "Блок 2: план недели = сумма планов её дней; дни недели, попавшие на разные месяцы, "
             "считаются по плану каждого месяца (стык месяцев учтён).")
wsX["A2"].font = F_NOTE
H1 = 4
for j, h in enumerate(["Менеджер", "Строка плана", "Дата", "Месяц (ключ)", "Рабочий?",
                       "№ фин. недели", "План месяца", "План на день"], 1):
    st(wsX.cell(H1, j, h), F_HDR, FILL_HDR, CENTER)
DAY_FIRST = H1 + 1
rp = DAY_FIRST
for name in MANAGERS:
    mr = MGR_ROW[name]
    for idx in range(ND):
        r = rp + idx; cr = CAL_FIRST + idx
        wsX.cell(r, 1, name)
        wsX.cell(r, 2, mr)
        wsX.cell(r, 3, f"='Календарь'!A{cr}").number_format = "dd.mm.yyyy"
        wsX.cell(r, 4, f"='Календарь'!B{cr}")
        wsX.cell(r, 5, f"='Календарь'!D{cr}")
        wsX.cell(r, 6, f"=INT((C{r}-Нед_Начало)/7)+1")
        # План месяца этого менеджера на месяц даты
        wsX.cell(r, 7,
            f"=IFNA(INDEX('Планы менеджеров'!$B$6:${LAST_PL}${RT-1},"
            f"MATCH(A{r},'Планы менеджеров'!$A$6:$A${RT-1},0),"
            f"MATCH(D{r},'Планы менеджеров'!$B$3:${LAST_PL}$3,0)),0)")
        # План на день = план месяца / рабочие дни месяца; в выходной = 0
        wsX.cell(r, 8, f"=IF(E{r}=0,0,G{r}/INDEX(Мес_Раб,MATCH(D{r},Мес_Ключ,0)))")
    rp += ND
DAY_LAST = rp - 1

H2 = DAY_LAST + 3
wsX.cell(H2 - 1, 1, "Блок 2. План по финансовым неделям").font = F_BOLD
for j, h in enumerate(["Менеджер", "№ недели", "Рабочих дней в неделе", "План недели (итого)",
                       "Из них план мес. A", "Из них план мес. B (стык)", "Проверка A+B=итого"], 1):
    st(wsX.cell(H2, j, h), F_HDR, FILL_HDR, CENTER)
WK_TBL = H2 + 1
W2ROWS = {}; rp = WK_TBL
for name in MANAGERS:
    first = rp
    for k in range(1, NW + 1):
        r = rp + k - 1
        df = DAY_FIRST + (k - 1) * 7; dl = df + 6
        wsX.cell(r, 1, name)
        wsX.cell(r, 2, k)
        wsX.cell(r, 3, f"=INDEX(Нед_Раб,{k})")
        wsX.cell(r, 4, f"=SUM(H{df}:H{dl})")
        # «Из них» — вклад каждого месяца в неделю. Формула через D-диапазон той же недели:
        # SUMIFS по столбцу H не работает (Excel запрещает пересечение диапазонов разной ширины).
        wsX.cell(r, 5, f"=SUMPRODUCT(($D${df}:$D${dl}=INDEX($D${df}:$D${dl},1))*($E${df}:$E${dl}))")
        wsX.cell(r, 6, f"=SUMPRODUCT(($D${df}:$D${dl}=INDEX($D${df}:$D${dl},7))*($E${df}:$E${dl}))")
        wsX.cell(r, 7, f"=IF(ROUND(E{r}+F{r}-D{r},2)=0,\"OK\",\"ОШИБКА\")")
    W2ROWS[name] = (first, rp + NW - 1)
    rp += NW
WK2_LAST = rp - 1
for r in range(WK_TBL, WK2_LAST + 1):
    for j in range(1, 8):
        c = wsX.cell(r, j); st(c, align=None if j == 1 else CENTER)
        if j in (4, 5, 6): c.number_format = "#,##0"
        if j in (3, 4, 5, 6, 7): c.fill = FILL_CALC
wsX.column_dimensions["A"].width = 26
for col, w in zip("BCDEFGH", [10, 12, 15, 14, 16, 14, 14]): wsX.column_dimensions[col].width = w
wsX.freeze_panes = f"A{DAY_FIRST}"

# ---- формулы ТАБЛИЦЫ 2 (колонка «План на активную неделю») ----
for i, name in enumerate(MANAGERS):
    r = ACTIVE_FIRST + i
    f, l = W2ROWS[name]
    st(wsP.cell(r, TB + 1, f"=INDEX('Расчёт'!$D${f}:$D${l},${LB}$8)"), fmt="#,##0", align=CENTER, fill=FILL_CALC)
st(wsP.cell(ACTIVE_TOT, TB, "ИТОГО"), F_BOLD, FILL_ALT)
st(wsP.cell(ACTIVE_TOT, TB + 1, f"=SUM({LP}{ACTIVE_FIRST}:{LP}{ACTIVE_LAST})"),
   F_BOLD, FILL_ALT, fmt="#,##0", align=CENTER)

# ---------- именованные диапазоны ----------
for nm, ref in {
    "Кал_Даты":   f"Календарь!$A${CAL_FIRST}:$A${CAL_LAST}",
    "Кал_Месяц":  f"Календарь!$B${CAL_FIRST}:$B${CAL_LAST}",
    "Кал_Раб":    f"Календарь!$D${CAL_FIRST}:$D${CAL_LAST}",
    "Нед_Начало": f"'Рабочие дни'!$B${WK_FIRST}:$B${WK_LAST}",
    "Нед_Конец":  f"'Рабочие дни'!$C${WK_FIRST}:$C${WK_LAST}",
    "Нед_Раб":    f"'Рабочие дни'!$E${WK_FIRST}:$E${WK_LAST}",
    "Мес_Раб":    f"'Рабочие дни'!$B${MR_FIRST}:$B${MR_LAST}",
    "Мес_Ключ":   f"'Рабочие дни'!$F${MR_FIRST}:$F${MR_LAST}",
}.items():
    wb.defined_names.add(DefinedName(nm, attr_text=ref))

# ================= Лист "План на фин. неделю" =================
wsW = wb.create_sheet("План на фин. неделю", 0)
wsW["A1"] = "ПЛАН НА ФИНАНСОВУЮ НЕДЕЛЮ (четверг — среда)"; wsW["A1"].font = F_TITLE
wsW["A2"] = ("Сотрудник находит свою фамилию и активную неделю — это его план на неделю. "
             "ТАБЛИЦА 2 «Активная неделя» перенесена на вкладку «Планы менеджеров».")
wsW["A2"].font = F_NOTE
st(wsW.cell(4, 1, "Активная дата:"), F_BOLD, border=False)
st(wsW.cell(4, 2, f"='Планы менеджеров'!{LP}5"), fill=FILL_CALC, fmt="dd.mm.yyyy", align=CENTER)
wsW.cell(4, 3, "← задаётся на вкладке «Планы менеджеров» (ТАБЛИЦА 2). Активная неделя подсвечивается ниже.").font = F_NOTE

wsW["A6"] = (f"ТАБЛИЦА 1. Все недели расчётного периода "
             f"({MONTHS_RU[months[0][1]-1]} {months[0][0]} — {MONTHS_RU[months[-1][1]-1]} {months[-1][0]})")
wsW["A6"].font = F_BOLD
st(wsW.cell(7, 1, "№ недели"), F_HDR, FILL_HDR, CENTER)
st(wsW.cell(8, 1, "Период"), F_HDR, FILL_HDR, CENTER)
st(wsW.cell(9, 1, "Рабочих дней"), F_HDR, FILL_HDR, CENTER)
for k, (a, b) in enumerate(weeks, 1):
    L = get_column_letter(1 + k)
    st(wsW.cell(7, 1 + k, k), align=CENTER)
    st(wsW.cell(8, 1 + k, f"{a.day:02d}.{a.month:02d}.{a.year} — {b.day:02d}.{b.month:02d}.{b.year}"), align=CENTER)
    st(wsW.cell(9, 1 + k, f"=INDEX(Нед_Раб,{k})"), align=CENTER, fill=FILL_CALC)
    wsW.column_dimensions[L].width = 15
wsW.column_dimensions["A"].width = 26
for i, name in enumerate(MANAGERS):
    r = 10 + i
    st(wsW.cell(r, 1, name), F_BOLD)
    f, l = W2ROWS[name]
    for k in range(1, NW + 1):
        st(wsW.cell(r, 1 + k, f"='Расчёт'!D{f + k - 1}"), fmt="#,##0", align=CENTER)
TOTR = 10 + NM
st(wsW.cell(TOTR, 1, "ИТОГО"), F_BOLD, FILL_ALT)
for k in range(1, NW + 1):
    L = get_column_letter(1 + k)
    st(wsW.cell(TOTR, 1 + k, f"=SUM({L}10:{L}{TOTR-1})"), F_BOLD, FILL_ALT, fmt="#,##0", align=CENTER)
wsW.freeze_panes = "B10"

# подсветка активной недели (номер недели берём из ТАБЛИЦЫ 2 на листе «Планы менеджеров»)
lastL = get_column_letter(1 + NW)
wsW.conditional_formatting.add(f"B8:{lastL}{TOTR}",
    FormulaRule(formula=[f"B$7='Планы менеджеров'!${LB}$8"], fill=PatternFill("solid", bgColor="FFE699")))

wb.save(OUT)
print("saved:", OUT, "| недель:", NW, "| дней календаря:", ND, "| менеджеров:", NM)
