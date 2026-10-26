# -*- coding: utf-8 -*-
"""Пересчёт всех формул файла «План на фин неделю.xlsx» независимым движком Excel-формул
(pip install formulas) и дамп значений в values.json для последующей сверки verify_final.py.

Запуск: python3 dump_values.py
"""
import json
import formulas

FN = "План на фин неделю.xlsx"

xl = formulas.ExcelModel().loads(FN).finish()
sol = xl.calculate()

out = {}
for k, v in sol.items():
    try:
        val = getattr(v, "value", None)
        if val is None:
            continue
        out[k.upper()] = str(val[0, 0]) if hasattr(val, "shape") else str(val)
    except Exception:
        pass

json.dump(out, open("values.json", "w", encoding="utf-8"), ensure_ascii=False)
print("dumped", len(out), "values")
