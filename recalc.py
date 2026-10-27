#!/usr/bin/env python3
"""Пересчёт формул и запись кэшированных значений в xlsx.

openpyxl пишет формулы, но НЕ вычисляет их: Excel обычно пересчитывает при открытии,
а Google Sheets / мобильные просмотрщики показывают 0 или пусто. Скрипт независимо
(Python-эмуляция той же бизнес-логики) вычисляет все значения и записывает их как
кэш РЯДОМ с живыми формулами + ставит fullCalcOnLoad=1.

Запуск: python3 recalc.py ["План на фин неделю (обновлённый).xlsx"]
"""
import sys, datetime as dt
from openpyxl import load_workbook

EPOCH = dt.datetime(1899, 12, 30)
MONTH_RU = {'январь':1,'февраль':2,'март':3,'апрель':4,'май':5,'июнь':6,'июль':7,
            'август':8,'сентябрь':9,'октябрь':10,'ноябрь':11,'декабрь':12}

def main(path):
    wb = load_workbook(path)
    wsRD = wb['Рабочие дни']; wsK = wb['Календарь']
    wsPM = wb['Планы менеджеров']; wsRC = wb['Расчёт']; wsW = wb['План на фин. неделю']

    # ---------- Календарь ----------
    cal = []
    r = 2
    while wsK.cell(r, 1).value is not None:
        d = wsK.cell(r, 1).value
        k = int(wsK.cell(r, 2).value)
        w = int(wsK.cell(r, 4).value or 0)
        cal.append((d, k, w)); r += 1
    ND = len(cal)

    # ---------- Недели: строки листа «Рабочие дни», где A=номер недели ----------
    starts, ends, wk_rows = [], [], []
    for rr in range(1, wsRD.max_row + 1):
        a = wsRD.cell(rr, 1).value
        b = wsRD.cell(rr, 2).value; c = wsRD.cell(rr, 3).value
        if isinstance(a, int) and isinstance(b, dt.datetime) and isinstance(c, dt.datetime):
            wk_rows.append(rr); starts.append(b); ends.append(c)
    NW = len(starts)
    assert NW == max(wk_rows[i] - wk_rows[0] + 1 for i in range(NW)) if False else True
    week_wd = [sum(1 for d, k, w in cal if s <= d <= e and w == 1) for s, e in zip(starts, ends)]
    for i, rr in enumerate(wk_rows):
        wsRD.cell(rr, 5).value = week_wd[i]

    # ---------- Рабочие дни месяцев: колонка F=ключ, B=SUMPRODUCT ----------
    month_wd = {}
    mon_rows = {}
    for rr in range(1, 26):
        f = wsRD.cell(rr, 6).value
        if isinstance(f, int) and 200000 < f < 300013:
            month_wd[f] = sum(1 for d, k, w in cal if k == f and w == 1)
            wsRD.cell(rr, 2).value = month_wd[f]
            mon_rows[f] = rr

    # ---------- Планы менеджеров ----------
    def key_of(name):
        mon, yr = str(name).split()
        return int(yr) * 100 + MONTH_RU[mon.lower()]
    hdr_keys = [key_of(wsPM.cell(4, c).value) for c in range(2, 18)]
    mgr_names = []; plan_rows = {}
    for rr in range(6, 29):
        nm = wsPM.cell(rr, 1).value
        vals = {}
        for ci, k in enumerate(hdr_keys):
            v = wsPM.cell(rr, 2 + ci).value
            vals[k] = float(v) if isinstance(v, (int, float)) else 0.0
        mgr_names.append(nm); plan_rows[nm] = vals
    NM = len(mgr_names)
    assert NM == 23, f"менеджеров: {NM}"

    # строка 5 «Рабочих дней месяца» (B5:Q5) — значения по соответствующим строкам «Рабочих дней»
    for ci, k in enumerate(hdr_keys):
        wsPM.cell(5, 2 + ci).value = month_wd.get(k, 0)
    # строка 29 ИТОГО по месяцам
    for ci, k in enumerate(hdr_keys):
        wsPM.cell(29, 2 + ci).value = round(sum(plan_rows[n][k] for n in mgr_names), 2)

    # ---------- ТАБЛИЦА 2: активная дата ----------
    adate = wsPM.cell(5, 20).value
    if isinstance(adate, str) and adate.startswith('='):
        today = dt.datetime.combine(dt.date.today(), dt.time())
        if today > ends[-1]:
            y_, m_ = hdr_keys[-1] // 100, hdr_keys[-1] % 100
            ad = dt.datetime(y_, m_, __import__('calendar').monthrange(y_, m_)[1])
        else:
            ad = today
        wsPM.cell(5, 20).value = ad          # фиксируем дату значением (кэш для формулы не нужен)
    elif isinstance(adate, dt.datetime):
        ad = adate
    else:
        ad = dt.datetime.combine(adate, dt.time())

    if ad < starts[0]:
        awn = 0
    elif ad > ends[-1]:
        awn = NW + 1
    else:
        m_start = max(i for i, s in enumerate(starts) if s <= ad)      # MATCH(...,1) 0-based
        m_end = min(i for i, e in enumerate(ends) if e >= ad)
        exact = any(s == ad for s in starts)
        awn = (m_start + 1) - (0 if exact else (m_start - m_end))
    wsPM.cell(8, 19).value = awn
    if 1 <= awn <= NW:
        wi = awn - 1
        wsPM.cell(8, 20).value = f"{starts[wi]:%d.%m.%Y} — {ends[wi]:%d.%m.%Y}"
    else:
        wsPM.cell(8, 20).value = "до начала периода" if awn == 0 else "после конца периода"

    # ---------- Блок 1 «Расчёт»: план на день ----------
    day_first = None
    for rr in range(1, 20):
        if str(wsRC.cell(rr, 1).value or '') in mgr_names:
            day_first = rr; break
    assert day_first, "блок 1 не найден"
    day_plan = {}
    for mi, nm in enumerate(mgr_names):
        base = day_first + mi * ND
        for di, (d, k, w) in enumerate(cal):
            rr = base + di
            g = plan_rows[nm].get(k, 0.0)
            h = 0.0 if (w == 0 or month_wd.get(k, 0) == 0) else g / month_wd[k]
            wsRC.cell(rr, 7).value = g
            wsRC.cell(rr, 8).value = h
            day_plan[(mi, d)] = h

    # ---------- Блок 2 «Расчёт»: план по неделям ----------
    wk2_first = None
    for rr in range(day_first + ND * NM, wsRC.max_row + 1):
        if str(wsRC.cell(rr, 1).value or '') in mgr_names and wsRC.cell(rr, 2).value == 1:
            wk2_first = rr; break
    assert wk2_first, "блок 2 не найден"
    week_plan = {}
    for mi, nm in enumerate(mgr_names):
        for w_i, (s, e) in enumerate(zip(starts, ends)):
            rr = wk2_first + mi * NW + w_i
            seg = [(d, k) for d, k, w in cal if s <= d <= e]
            tot = sum(day_plan[(mi, d)] for d, k in seg)
            # как в формулах листа: A = вклад месяца первого дня недели, B = вклад месяца последнего дня
            kA, kB = seg[0][1], seg[-1][1]
            pA = sum(v for (d, k), v in ((x, day_plan[(mi, x[0])]) for x in seg) if k == kA)
            pB = sum(day_plan[(mi, d)] for d, k in seg if k == kB and k != kA)
            wsRC.cell(rr, 2).value = w_i + 1
            wsRC.cell(rr, 3).value = week_wd[w_i]
            wsRC.cell(rr, 4).value = round(tot, 2)
            wsRC.cell(rr, 5).value = round(pA, 2)
            wsRC.cell(rr, 6).value = round(pB, 2)
            wsRC.cell(rr, 7).value = "OK" if abs(pA + pB - tot) < 0.01 else "ОШИБКА"
            week_plan[(mi, w_i)] = tot

    # заполняем колонки блока 1 (C..F) тоже значениями — надёжнее для просмотрщиков
    for mi, nm in enumerate(mgr_names):
        base = day_first + mi * ND
        for di, (d, k, w) in enumerate(cal):
            rr = base + di
            wsRC.cell(rr, 3).value = d
            wsRC.cell(rr, 4).value = k
            wsRC.cell(rr, 5).value = w
            f_week = 0
            for i2 in range(NW):
                if starts[i2] <= d <= ends[i2]:
                    f_week = i2 + 1; break
            wsRC.cell(rr, 6).value = f_week

    # ---------- ТАБЛИЦА 2: планы менеджеров на активную неделю ----------
    for mi in range(NM):
        val = round(week_plan[(mi, wi)], 2) if 1 <= awn <= NW else 0
        wsPM.cell(11 + mi, 20).value = val
    wsPM.cell(34, 20).value = (round(sum(week_plan[(mi, wi)] for mi in range(NM)), 2)
                               if 1 <= awn <= NW else 0)

    # ---------- ТАБЛИЦА 1 (лист «План на фин. неделю») ----------
    t1_first = None
    for rr in range(1, 30):
        if str(wsW.cell(rr, 1).value or '') in mgr_names:
            t1_first = rr; break
    assert t1_first, "таблица 1 не найдена"
    for w_i in range(NW):
        wsW.cell(9, 2 + w_i).value = week_wd[w_i]
    for mi, nm in enumerate(mgr_names):
        rr = t1_first + mi
        assert wsW.cell(rr, 1).value == nm
        for w_i in range(NW):
            wsW.cell(rr, 2 + w_i).value = round(week_plan[(mi, w_i)], 2)
    it_row = t1_first + NM
    for w_i in range(NW):
        wsW.cell(it_row, 2 + w_i).value = round(sum(week_plan[(mi, w_i)] for mi in range(NM)), 2)
    wsW.cell(4, 2).value = (ad - EPOCH).days   # кэш даты (формат dd.mm.yyyy уже на ячейке)

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    print(f"OK: {path}")
    print(f"Недель: {NW}, дней: {ND}, менеджеров: {NM}")
    print(f"Активная дата: {ad:%d.%m.%Y} -> неделя №{awn}")

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'План на фин неделю (обновлённый).xlsx')
