"""Generate a deliberately messy orders CSV so the cleaner has something real to fix.

Deterministic (fixed seed) and dependency-free. Every defect class is injected on
purpose and counted, so the audit report can be checked against this file.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(20261126)

FIRST = ["ana", "Bo", "chris", "DANA", "eli", "fran", "gita", "Hugo", "ines", "jo"]
LAST = ["liu", "patel", "smith", "oksana", "nagy", "kim", "dubois", "rossi", "costa", "wang"]
REGIONS = ["north", "South", "east ", "WEST", " central"]
TIERS = ["basic", "pro", "enterprise", " PRO", "Basic "]

start = date(2026, 1, 1)
rows = []
for i in range(1, 481):
    first = random.choice(FIRST)
    last = random.choice(LAST)
    email = f"{first}.{last}@example.com".lower()
    d = start + timedelta(days=random.randint(0, 320))
    amount = round(random.uniform(9, 4800), 2)
    region = random.choice(REGIONS)
    tier = random.choice(TIERS)
    rows.append({
        "order_id": f"SO-{10000 + i}",
        "customer_name": f"{first} {last}",
        "email": email,
        "order_date": d.isoformat(),
        "amount": f"{amount:.2f}",
        "region": region,
        "tier": tier,
        "notes": "",
        "_iso": d.isoformat(),
        "_amt": amount,
    })

defects = {k: 0 for k in [
    "dup_order", "email_case", "date_dmy", "date_us", "date_long",
    "amount_currency", "amount_thousands", "blank_amount", "name_padded",
    "region_padded", "tier_case", "missing_email",
]}

# 1) duplicate whole rows (same order_id, re-submitted twice)
for _ in range(34):
    dup = dict(random.choice(rows))
    rows.append(dup)
    defects["dup_order"] += 1

# 2) email case / whitespace
for _ in range(45):
    r = random.choice(rows)
    r["email"] = " " + r["email"].upper() + " "
    defects["email_case"] += 1

# 3) three extra date formats
for _ in range(40):
    r = random.choice(rows)
    y, m, dd = r["_iso"].split("-")
    r["order_date"] = f"{dd}/{m}/{y}"
    defects["date_dmy"] += 1
for _ in range(25):
    r = random.choice(rows)
    y, m, dd = r["_iso"].split("-")
    r["order_date"] = f"{m}/{dd}/{y}"
    defects["date_us"] += 1
for _ in range(15):
    r = random.choice(rows)
    y, m, dd = r["_iso"].split("-")
    r["order_date"] = f"{dd} {['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][int(m)-1]} {y}"
    defects["date_long"] += 1

# 4) currency symbols and thousands separators
for _ in range(60):
    r = random.choice(rows)
    r["amount"] = f" ${r['_amt']:,.2f} "
    defects["amount_currency"] += 1
for _ in range(20):
    r = random.choice(rows)
    r["amount"] = f"{r['_amt']:,.2f}"
    defects["amount_thousands"] += 1
for _ in range(18):
    r = random.choice(rows)
    r["amount"] = ""
    defects["blank_amount"] += 1

# 5) padded / mixed-case categoricals
for _ in range(38):
    r = random.choice(rows)
    r["customer_name"] = "  " + r["customer_name"] + "   "
    defects["name_padded"] += 1
for _ in range(30):
    r = random.choice(rows)
    r["region"] = r["region"].upper()
    defects["region_padded"] += 1
for _ in range(30):
    r = random.choice(rows)
    r["tier"] = r["tier"].lower()
    defects["tier_case"] += 1

# 6) truly missing email (unrecoverable -> must be flagged, not invented)
for _ in range(12):
    r = random.choice(rows)
    r["email"] = ""
    defects["missing_email"] += 1

random.shuffle(rows)

out = Path(__file__).resolve().parent.parent / "samples" / "messy_orders.csv"
out.parent.mkdir(parents=True, exist_ok=True)
with out.open("w", newline="", encoding="utf-8") as fh:
    fields = [k for k in rows[0].keys() if not k.startswith("_")]
    w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)

import json
print(json.dumps({"rows_written": len(rows), "injected_defects": defects}, indent=1))
