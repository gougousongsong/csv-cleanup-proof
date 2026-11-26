"""Invariant tests for the cleaner. Plain asserts, no framework needed:  python scripts/test_clean.py"""
from __future__ import annotations

import csv
import json
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "samples" / "messy_orders.csv"
OUT = ROOT / "out" / "clean_orders.xlsx"
REPORT = ROOT / "out" / "report.json"


def main() -> int:
    rep = json.loads(REPORT.read_text())
    wb = load_workbook(OUT, read_only=True)
    rows = list(wb["cleaned"].iter_rows(values_only=True))
    head = rows[0]
    body = rows[1:]
    rec = [dict(zip(head, r)) for r in body]
    review = list(wb["needs_review"].iter_rows(values_only=True))
    review_body = review[1:]

    # 1) no duplicate order ids survive
    ids = [r["order_id"] for r in rec]
    assert len(ids) == len(set(ids)), "duplicate order_id left in cleaned sheet"

    # 2) nothing is silently lost: read = kept + dropped
    assert rep["rows_read"] == rep["rows_kept"] + rep["duplicates_dropped"], "row accounting does not balance"

    # 3) every flagged row is present in needs_review, and vice versa
    flagged = {r["order_id"] for r in rec if r["needs_review"] == "yes"}
    in_review = {r[0] for r in review_body}
    assert flagged == in_review, f"flag/review mismatch: only-flag={flagged - in_review} only-review={in_review - flagged}"
    assert len(flagged) == rep["rows_flagged"], "rows_flagged does not match the sheet"

    # 4) amounts are numeric or empty, never text leftovers
    for r in rec:
        a = r["amount"]
        assert a is None or isinstance(a, (int, float)), f"non-numeric amount in {r['order_id']}: {a!r}"

    # 5) money reconciles with the audit sheet
    total = round(sum(r["amount"] for r in rec if r["amount"] is not None), 2)
    assert abs(total - rep["amount_total_all_parsed"]) < 0.01, f"total {total} != report {rep['amount_total_all_parsed']}"

    # 6) dates are ISO and parseable
    from datetime import date
    for r in rec:
        if r["order_date"]:
            date.fromisoformat(r["order_date"])

    # 7) missing values were not invented: count of blanks equals the reported missing counts
    blank_amount = sum(1 for r in rec if r["amount"] is None)
    assert blank_amount == rep["transformations"]["amount_missing"], "blank amounts != reported missing"
    blank_email = sum(1 for r in rec if not r["email"])
    assert blank_email == rep["transformations"]["email_missing"], "blank emails != reported missing"

    # 8) source file really is messier than the output (proves work was done)
    with SRC.open(newline="", encoding="utf-8") as fh:
        raw = list(csv.DictReader(fh))
    assert len(raw) > len(rec), "output is not smaller than input"

    print(json.dumps({
        "tests": 8,
        "passed": 8,
        "rows_in": len(raw),
        "rows_out": len(rec),
        "flagged": len(flagged),
        "amount_total": total,
    }, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
