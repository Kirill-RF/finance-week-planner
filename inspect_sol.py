import datetime as dt, logging, pickle
logging.disable(logging.CRITICAL)
import formulas as F
FN="План на фин неделю.xlsx"
sol=F.ExcelModel().loads(FN).finish().calculate()
with open("sol.pkl","wb") as f: pickle.dump({k:(getattr(v,'value',v)) for k,v in sol.items()},f)
keys=list(sol.keys())
print(len(keys))
# показать ключи для листа План на фин. неделю строки 10
cnt=0
for k in keys:
    if "ПЛАН НА ФИН" in k.upper() and "!B10" in k.upper():
        v=sol[k]
        try: print(k,"=>",v.value[0,0])
        except Exception: print(k,"=>",v)
        cnt+=1
        if cnt>3: break
