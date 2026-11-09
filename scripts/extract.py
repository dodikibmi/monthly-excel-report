"""Run the monthly SQL and write the extract the Excel macro reads.
   python scripts/extract.py 2026-10
-> data/extract-2026-10.csv, data/extract-2026-09.csv (previous month, for the comparison)
   and data/control-2026-10.json (control totals for the test)."""
import csv, json, sqlite3, sys
from pathlib import Path

root = Path(__file__).resolve().parent.parent
period = sys.argv[1] if len(sys.argv) > 1 else "2026-10"


def bounds(p):
    y, m = map(int, p.split("-"))
    ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
    return f"{y:04d}-{m:02d}-01", f"{ny:04d}-{nm:02d}-01"


def previous(p):
    y, m = map(int, p.split("-"))
    return f"{y - 1:04d}-12" if m == 1 else f"{y:04d}-{m - 1:02d}"


con = sqlite3.connect(root / "data" / "sales.db")
sql = (root / "sql" / "monthly_extract.sql").read_text()
checks = (root / "sql" / "checks.sql").read_text()
control = {}
for p in (period, previous(period)):
    s, e = bounds(p)
    cur = con.execute(sql, {"start": s, "end": e})
    out = root / "data" / f"extract-{p}.csv"
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([c[0] for c in cur.description])
        rows = cur.fetchall()
        w.writerows(rows)
    orders, revenue, margin = con.execute(checks, {"start": s, "end": e}).fetchone()
    control[p] = {"lines": len(rows), "orders": orders, "revenue": revenue, "margin": margin}
    print(f"{out.name}: {len(rows)} lines, {orders} orders, revenue {revenue:,.2f}, margin {margin:,.2f}")
(root / "data" / f"control-{period}.json").write_text(json.dumps(control, indent=2))
