"""Check the Excel report against the SQL control totals.
   python test/verify.py 2026-10
Fails (exit 1) if any number in the report differs from what the database says."""
import json, sys
from pathlib import Path
from openpyxl import load_workbook

root = Path(__file__).resolve().parent.parent
period = sys.argv[1] if len(sys.argv) > 1 else "2026-10"
control = json.loads((root / "data" / f"control-{period}.json").read_text())
cur = control[period]
prev = next(v for k, v in control.items() if k != period)
wb = load_workbook(root / "out" / f"Sales-Report-{period}.xlsx")
s = wb["Summary"]
fails = []


def check(name, got, want, tol=0.005):
    ok = got is not None and abs(float(got) - float(want)) <= tol
    print(f"{'OK  ' if ok else 'FAIL'} {name}: report {got} / SQL {want}")
    if not ok:
        fails.append(name)


check("revenue", s["B5"].value, cur["revenue"])
check("gross margin", s["B6"].value, cur["margin"])
check("margin %", s["B7"].value, cur["margin"] / cur["revenue"], 1e-9)
check("orders", s["B8"].value, cur["orders"])
check("average order value", s["B9"].value, cur["revenue"] / cur["orders"], 0.005)
check("previous revenue", s["C5"].value, prev["revenue"])
check("previous orders", s["C8"].value, prev["orders"])
check("revenue change", s["D5"].value, (cur["revenue"] - prev["revenue"]) / prev["revenue"], 1e-9)
check("regions add up to revenue", sum(s.cell(r, 7).value for r in range(5, 9)), cur["revenue"], 0.01)
data_rows = wb["Data"].max_row - 1
check("data lines", data_rows, cur["lines"], 0)
top = wb["Top customers"]
names = [top.cell(r, 1).value for r in range(5, top.max_row + 1)
         if top.cell(r, 1).value not in (None, "Grand Total")]
check("top customers listed", len(names), 10, 0)
revs = [top.cell(r, 2).value for r in range(5, 5 + len(names))]
print("OK  " if revs == sorted(revs, reverse=True) else "FAIL", "top customers sorted by revenue:", names[:3], "...")
if revs != sorted(revs, reverse=True):
    fails.append("top sort")
pdf = root / "out" / f"Sales-Report-{period}.pdf"
check("pdf written (more than 10 kB)", int(pdf.exists() and pdf.stat().st_size > 10000), 1, 0)
print("\nALL CHECKS PASSED" if not fails else f"\n{len(fails)} FAILED: {fails}")
sys.exit(1 if fails else 0)
