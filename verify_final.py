# -*- coding: utf-8 -*-
"""Финальная независимая проверка файла «План на фин неделю.xlsx».

Сверяет значения, рассчитанные моделью Excel (дамп values.json из пакета formulas),
с эталонным расчётом на чистом Python по данным самих листов файла.

Запуск:
    python3 dump_values.py     # пересобрать values.json из xlsx
    python3 verify_final.py    # сверить; при 0 ошибок печатает ОК
"""
import json, re, datetime as dt
import openpyxl

FN = "План на фин неделю.xlsx"
wb = openpyxl.load_workbook(FN)
vals = {k.upper().replace('"', ''): v
        for k, v in json.load(open("values.json", encoding="utf-8")).items()}

def get(sheet, cell):
    """Значение ячейки по дампу вычислений Excel (float/str/None)."""
    v = vals.get(f"'[{FN.upper()}]{sheet.upper()}'!{cell}")
    if v is None:
        return None
    m = re.search(r"\[\[([^\]]+)\]\]", str(v))
    s = m.group(1).split(",")[0] if m else str(v)
    try:
        return float(s)
    except ValueError:
        return s

# --- данные листа «Календарь» (даты и флаги рабочих дней — статические ячейки) ---
wsK = wb["Календарь"]
cal = []
r = 2
while isinstance(wsK.cell(r, 1).value, dt.datetime):
    cal.append((wsK.cell(r, 1).value.date(), wsK.cell(r, 4).value))
    r += 1
all_days = [d for d, _ in cal]
n = len(all_days)
wd_in_month = {m: sum(1 for d, f in cal if d.month == m and f == 1) for m in range(1, 13)}
ND = (n + 6) // 7                      # число финансовых недель
weeks = [(all_days[i], all_days[min(i + 6, n - 1)]) for i in range(0, n, 7)]

# --- планы менеджеров (лист «Планы менеджеров», строки 5..16, столбцы B..M) ---
wsP = wb["Планы менеджеров"]
mgrs = [wsP.cell(rr, 1).value for rr in range(5, 17)]
real_plans = {name: [get("ПЛАНЫ МЕНЕДЖЕРОВ", f"{chr(66 + c)}{rr}") for c in range(12)]
              for rr, name in zip(range(5, 17), mgrs)}

errors = []
def chk(tag, got, exp, tol=1e-6):
    if got is None or isinstance(got, str) or abs(got - exp) > tol:
        errors.append((tag, got, exp))

# 1) Рабочие дни месяца и недели (лист «Рабочие дни»)
for m in range(1, 13):
    chk(f"Раб.дни мес {m}", get("РАБОЧИЕ ДНИ", f"B{4 + m}"), wd_in_month[m])
for k in range(ND):
    exp = sum(1 for d, f in cal[k * 7:k * 7 + 7] if f == 1)
    chk(f"Раб.дни нед {k + 1}", get("РАБОЧИЕ ДНИ", f"E{20 + k}"), exp)

# 2) Блок 1 листа «Расчёт»: план на каждый день для каждого менеджера (все ячейки)
W0 = 5
for mi, mgr in enumerate(mgrs):
    for i in range(n):
        d, f = cal[i]
        rr = W0 + mi * n + i
        exp = 0.0 if f != 1 else real_plans[mgr][d.month - 1] / wd_in_month[d.month]
        chk(f"G{rr} {mgr} {d}", get("РАСЧЁТ", f"G{rr}"), exp)

# 3) Блок 2 листа «Расчёт»: план недели, стык месяцев (E+F=D -> H=OK), годовая сумма
B2 = 4460
for mi, mgr in enumerate(mgrs):
    for k in range(ND):
        rr = B2 + mi * ND + k
        exp = 0.0
        for j in range(7):
            idx = k * 7 + j
            if idx < n and cal[idx][1] == 1:
                dd = all_days[idx]
                exp += real_plans[mgr][dd.month - 1] / wd_in_month[dd.month]
        chk(f"D{rr} {mgr} нед{k + 1}", get("РАСЧЁТ", f"D{rr}"), exp)
        if get("РАСЧЁТ", f"H{rr}") != "OK":
            errors.append((f"H{rr} {mgr} нед{k + 1}", get("РАСЧЁТ", f"H{rr}"), "OK"))
    ann = sum(get("РАСЧЁТ", f"D{B2 + mi * ND + k}") for k in range(ND))
    chk(f"Годовая {mgr}", ann, sum(real_plans[mgr]), tol=1.0)

# 4) Лист «План на фин. неделю»: активная неделя по активной дате (B4)
act = wb["План на фин. неделю"].cell(4, 2).value
act = act.date() if isinstance(act, dt.datetime) else act
exp_active = next(k for k, (a, b) in enumerate(weeks, 1) if a <= act <= b)
chk("Активная неделя", get("ПЛАН НА ФИН. НЕДЕЛЮ", "C4"), exp_active)

print(f"Проверено: {n} дней × {len(mgrs)} менеджеров, {ND} недель")
print("ОШИБОК:", len(errors))
for e in errors[:20]:
    print(" -", e)
raise SystemExit(0 if not errors else 1)
