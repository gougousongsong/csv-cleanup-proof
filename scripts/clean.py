"""Clean a messy orders export into an auditable Excel workbook.

Design rules that matter to a buyer:
  * every transformation is counted and written to the audit sheet;
  * anything ambiguous is flagged in needs_review with the reason, never guessed silently;
  * missing values are never invented - they are marked and reported;
  * the run is deterministic and re-runnable, so the buyer can regress it.

Usage:
    python scripts/make_sample.py
    python scripts/clean.py samples/messy_orders.csv out/clean_orders.xlsx

Requires: openpyxl
"""
from __future__ import annotations

import csv
import json
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}
HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FLAG_FILL = PatternFill("solid", fgColor="FDE68A")


def parse_amount(raw: str) -> tuple[float | None, str | None]:
    """Return (value, problem). Handles $, commas, spaces, parentheses negatives."""
    if raw is None:
        return None, "missing"
    s = str(raw).strip()
    if not s:
        return None, "missing"
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace("USD", "").replace("\u00a0", "").replace(" ", "").replace(",", "")
    if not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return None, f"unparseable:{raw!r}"
    v = float(s)
    return (-v if neg else v), None


def parse_date(raw: str) -> tuple[date | None, str]:
    """Return (date, confidence). confidence in {iso, dmy, mdy, month_word, ambiguous, bad}."""
    s = str(raw or "").strip()
    if not s:
        return None, "bad"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        y, m, d = (int(x) for x in s.split("-"))
        return date(y, m, d), "iso"
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 12 and b <= 12:
            return date(y, b, a), "dmy"
        if b > 12 and a <= 12:
            return date(y, a, b), "mdy"
        if a <= 12 and b <= 12:
            # both readings valid -> do not pretend to know; default to day-first
            return date(y, b, a), "ambiguous"
        return None, "bad"
    m = re.fullmatch(r"(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})", s)
    if m and m.group(2).title() in MONTHS:
        return date(int(m.group(3)), MONTHS[m.group(2).title()], int(m.group(1))), "month_word"
    return None, "bad"


def title_case_name(raw: str) -> str:
    return " ".join(w.capitalize() if w.isupper() or w.islower() else w
                    for w in str(raw or "").strip().split())


def run(src: Path, dst: Path) -> dict:
    with src.open(newline="", encoding="utf-8") as fh:
        raw_rows = list(csv.DictReader(fh))

    audit: Counter = Counter()
    review: list[dict] = []
    cleaned: list[dict] = []
    seen_orders: dict[str, int] = {}

    for i, r in enumerate(raw_rows, start=2):  # +2 = spreadsheet line number
        row_issues: list[str] = []

        oid = (r.get("order_id") or "").strip()
        if oid in seen_orders:
            audit["duplicate_order_id_dropped"] += 1
            continue
        seen_orders[oid] = i

        # email: trim + lowercase, keep original if it changes
        raw_email = r.get("email") or ""
        email = raw_email.strip().lower()
        if raw_email != email and raw_email.strip():
            audit["email_normalised"] += 1
        if not email:
            row_issues.append("email_missing")
            audit["email_missing"] += 1
        elif not re.fullmatch(r"[^@\s]+@[^@\s]+\.[A-Za-z]{2,}", email):
            row_issues.append("email_malformed")
            audit["email_malformed"] += 1

        # name
        raw_name = r.get("customer_name") or ""
        name = title_case_name(raw_name)
        if raw_name != name:
            audit["name_normalised"] += 1

        # date
        raw_date = r.get("order_date") or ""
        d, conf = parse_date(raw_date)
        if d is None:
            row_issues.append("date_unparseable")
            audit["date_unparseable"] += 1
        else:
            if conf != "iso":
                audit["date_reformatted"] += 1
            if conf == "ambiguous":
                row_issues.append("date_ambiguous_day_month_first_assumed")
                audit["date_ambiguous"] += 1

        # amount
        amount, problem = parse_amount(r.get("amount"))
        if problem == "missing":
            row_issues.append("amount_missing")
            audit["amount_missing"] += 1
        elif problem:
            row_issues.append(problem)
            audit["amount_unparseable"] += 1
        elif str(r.get("amount")).strip() != f"{amount:.2f}":
            audit["amount_normalised"] += 1

        # categoricals
        region_raw = r.get("region") or ""
        region = region_raw.strip().title()
        if region_raw != region:
            audit["region_normalised"] += 1
        tier_raw = r.get("tier") or ""
        tier = tier_raw.strip().title()
        if tier_raw != tier:
            audit["tier_normalised"] += 1

        out = {
            "order_id": oid,
            "customer_name": name,
            "email": email,
            "order_date": d.isoformat() if d else "",
            "date_confidence": conf,
            "amount": round(amount, 2) if amount is not None else "",
            "region": region,
            "tier": tier,
            "source_row": i,
            "needs_review": "yes" if row_issues else "",
            "issues": "; ".join(row_issues),
        }
        cleaned.append(out)
        if row_issues:
            review.append(out)

    keep = [c for c in cleaned if not c["needs_review"]]
    total = sum(c["amount"] for c in cleaned if c["amount"] != "")
    total_clean = sum(c["amount"] for c in keep if c["amount"] != "")

    report = {
        "source_file": src.name,
        "rows_read": len(raw_rows),
        "rows_kept": len(cleaned),
        "duplicates_dropped": audit["duplicate_order_id_dropped"],
        "rows_clean": len(keep),
        "rows_flagged": len(review),
        "flag_rate_pct": round(100 * len(review) / max(len(cleaned), 1), 1),
        "amount_total_all_parsed": round(total, 2),
        "amount_total_clean_rows": round(total_clean, 2),
        "transformations": dict(sorted(audit.items())),
        "flag_reasons": dict(Counter(
            reason for c in review for reason in c["issues"].split("; ") if reason)),
    }

    # ---- workbook ----
    wb = Workbook()
    ws = wb.active
    ws.title = "cleaned"
    cols = ["order_id", "customer_name", "email", "order_date", "date_confidence",
            "amount", "region", "tier", "source_row", "needs_review", "issues"]
    ws.append(cols)
    for c in ws[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for row in cleaned:
        ws.append([row[c] for c in cols])
        if row["needs_review"]:
            for cell in ws[ws.max_row]:
                cell.fill = FLAG_FILL
    widths = {"order_id": 12, "customer_name": 20, "email": 30, "order_date": 12,
              "date_confidence": 15, "amount": 11, "region": 11, "tier": 11,
              "source_row": 11, "needs_review": 13, "issues": 44}
    for idx, name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths[name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("needs_review")
    ws2.append(cols)
    for c in ws2[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for row in review:
        ws2.append([row[c] for c in cols])
    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions

    ws3 = wb.create_sheet("audit")
    ws3.append(["metric", "value"])
    for c in ws3[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for k, v in report.items():
        if isinstance(v, dict):
            for sub, n in v.items():
                ws3.append([f"{k}.{sub}", n])
        else:
            ws3.append([k, v])
    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 22

    ws4 = wb.create_sheet("rules")
    ws4.append(["rule", "action", "why"])
    for c in ws4[1]:
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    for rule, action, why in [
        ("duplicate order_id", "keep first occurrence, drop later ones", "resubmitted exports double-count revenue"),
        ("email", "trim + lowercase", "same mailbox, one customer"),
        ("customer_name", "trim + title case", "report consistency"),
        ("order_date", "reformat to ISO; day-first when day/month both <=12", "one sortable format"),
        ("ambiguous date", "kept, flagged needs_review", "05/06 could be May 6 or Jun 5 - not guessable from data"),
        ("amount", "strip $/commas/spaces, parentheses = negative", "numeric sums"),
        ("missing amount/email", "left blank, flagged", "never invent a value the buyer did not send"),
        ("region / tier", "trim + title case", "group-by consistency"),
    ]:
        ws4.append([rule, action, why])
    ws4.column_dimensions["A"].width = 24
    ws4.column_dimensions["B"].width = 50
    ws4.column_dimensions["C"].width = 56
    for row in ws4.iter_rows(min_row=2):
        row[2].alignment = Alignment(wrap_text=True, vertical="top")

    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)
    report["workbook"] = str(dst)
    return report


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("samples/messy_orders.csv")
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("out/clean_orders.xlsx")
    rep = run(src, dst)
    print(json.dumps(rep, indent=1, ensure_ascii=False))
