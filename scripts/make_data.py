"""Create data/sales.db with fictitious, reproducible data for Sep and Oct 2026 (seeded)."""
import random, sqlite3
from datetime import date, timedelta
from pathlib import Path

root = Path(__file__).resolve().parent.parent
db = root / "data" / "sales.db"
db.unlink(missing_ok=True)
con = sqlite3.connect(db)
con.executescript((root / "sql" / "schema.sql").read_text())
rnd = random.Random(2026)

regions = ["North", "South", "East", "West"]
segments = ["Retail", "Wholesale", "Online"]
first = ["Blue", "Cedar", "Summit", "Harbor", "Golden", "River", "Oak", "Silver", "Prime", "Metro"]
last = ["Traders", "Mart", "Supply Co", "Stores", "Outlet"]
customers = []
for i in range(40):
    cid = f"C{i + 1:03d}"
    customers.append((cid, f"{first[i % 10]} {last[i // 10 + (i % 2)]}", regions[i % 4], segments[i % 3]))
con.executemany("INSERT INTO customers VALUES (?,?,?,?)", customers)

products = [
    ("OF-100", "Copy paper A4 (box)", "Office", 18.0), ("OF-110", "Ballpoint pens (50)", "Office", 6.5),
    ("OF-120", "Lever arch files (10)", "Office", 11.0), ("CL-200", "Floor cleaner 5L", "Cleaning", 9.0),
    ("CL-210", "Hand soap refill 5L", "Cleaning", 7.5), ("CL-220", "Paper towels (24)", "Cleaning", 14.0),
    ("PK-300", "Shipping boxes (25)", "Packaging", 12.0), ("PK-310", "Packing tape (36)", "Packaging", 21.0),
    ("PK-320", "Bubble wrap roll", "Packaging", 16.0), ("PA-400", "Coffee beans 1kg", "Pantry", 13.0),
    ("PA-410", "Tea bags (500)", "Pantry", 10.0), ("PA-420", "Bottled water (24)", "Pantry", 4.0),
]
con.executemany("INSERT INTO products VALUES (?,?,?,?)", products)

n = 0
for month_start, count, growth in ((date(2026, 9, 1), 210, 1.0), (date(2026, 10, 1), 230, 1.06)):
    days = 30 if month_start.month == 9 else 31
    for _ in range(count):
        n += 1
        order_no = f"SO-{26000 + n}"
        d = month_start + timedelta(days=rnd.randrange(days))
        status = rnd.choices(["INVOICED", "CANCELLED", "RETURNED"], [92, 5, 3])[0]
        con.execute("INSERT INTO orders VALUES (?,?,?,?)", (order_no, d.isoformat(), rnd.choice(customers)[0], status))
        for line in range(1, rnd.randint(1, 4) + 1):
            sku, _, _, cost = rnd.choice(products)
            price = round(cost * rnd.uniform(1.25, 1.6) * growth, 2)
            con.execute("INSERT INTO order_lines VALUES (?,?,?,?,?,?)",
                        (order_no, line, sku, rnd.randint(1, 30), price, rnd.choice([0, 0, 0, 5, 10])))
con.commit()
print("orders:", con.execute("select count(*) from orders").fetchone()[0],
      "lines:", con.execute("select count(*) from order_lines").fetchone()[0])
