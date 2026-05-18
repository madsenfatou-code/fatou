import os, sys
sys.path.insert(0, '/home/user/fatou')

from pptx_helpers import make_day
from days_data_01 import DAYS as DAYS1
from days_data_02 import DAYS as DAYS2
from days_data_03 import DAYS as DAYS3

ALL_DAYS = DAYS1 + DAYS2 + DAYS3

out_dir = "/home/user/fatou/modules_pptx"
os.makedirs(out_dir, exist_ok=True)

errors = []
for d in ALL_DAYS:
    n = d['day_num']
    safe = d['title'][:35].replace(' ','_').replace('/','').replace("'","")
    filename = f"Jour{n:02d}_{d['day_name']}_{safe}.pptx"
    path = os.path.join(out_dir, filename)
    try:
        prs = make_day(d)
        prs.save(path)
        print(f"  ✅  Jour {n:02d} — {d['title'][:50]}")
    except Exception as e:
        errors.append((n, str(e)))
        print(f"  ❌  Jour {n:02d} — ERREUR : {e}")

print(f"\n{'='*50}")
print(f"Généré : {len(ALL_DAYS) - len(errors)} / {len(ALL_DAYS)} fichiers")
if errors:
    print("Erreurs :")
    for n, e in errors:
        print(f"  Jour {n}: {e}")
print(f"Dossier : {out_dir}")
