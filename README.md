# CSV / Excel cleanup with an audit trail

Messy order exports in, a clean workbook out — **plus a sheet that proves what changed**.
This is a working demo, not a mockup: every number below comes from `scripts/clean.py`
run against `samples/messy_orders.csv` in this repository.

## Why an audit trail

Anyone can delete duplicate rows. The expensive part is answering two questions
afterwards:

1. *What exactly did you change to my data?* → the `audit` sheet counts every rule.
2. *Where did you guess?* → the `needs_review` sheet lists **every** row whose value
   could not be determined from the source, with the reason. Missing amounts and
   ambiguous dates are left blank and flagged, never filled in with a plausible value.

## Measured output (514 input rows → 480 cleaned rows)

```
rows_read                          514
rows_kept                          480
duplicates_dropped                  34   (re-submitted exports double-count revenue)
rows_clean                         430
rows_flagged                        50   (10.4% - listed with a reason each)
amount_total_all_parsed     1,135,974.74
amount_total_clean_rows     1,051,120.88
```

Transformations applied, with counts:

```
duplicate_order_id_dropped   34    keep first occurrence, drop the rest
date_reformatted             68    3 extra formats -> ISO 8601
date_ambiguous               24   05/06 could be May 6 or Jun 5 -> flagged, day-first recorded
email_normalised             40    trim + lowercase (same mailbox, one customer)
email_missing                10    left blank, flagged
amount_normalised            66    "$1,197.00 " and "1,197.00" -> 1197.0
amount_missing               17    left blank, flagged (never invented)
name_normalised             480    trim + title case
region_normalised           400    trim + title case
tier_normalised             480    trim + title case
```

## What you get

`out/clean_orders.xlsx`, four sheets:

- **cleaned** — one row per order, ISO dates, numeric amounts, normalised text,
  `source_row` pointing back at the original line, flagged rows highlighted.
- **needs_review** — only the rows a human must decide on, with `issues` per row.
- **audit** — every metric above as `key / value`, so nothing is claimed without a count.
- **rules** — each rule, its action, and why it exists.

## Run it

```bash
pip install openpyxl
python scripts/make_sample.py                                  # regenerate the messy input
python scripts/clean.py samples/messy_orders.csv out/clean_orders.xlsx
python scripts/test_clean.py                                   # 8 invariant checks
```

`scripts/test_clean.py` asserts, among others: no duplicate `order_id` survives;
`rows_read == rows_kept + duplicates_dropped`; every flagged row appears in
`needs_review` and vice versa; amounts are numeric or empty; the money total in the
workbook equals the total in the report; blank counts equal reported missing counts.

```
{"tests": 8, "passed": 8, "rows_in": 514, "rows_out": 480, "flagged": 50,
 "amount_total": 1135974.74}
```

## Adapting it to your file

`scripts/clean.py` is plain Python with no framework. A new client file usually means
editing three places: the column names in `run()`, the amount/date patterns in
`parse_amount()` / `parse_date()`, and the `rules` sheet text. The audit and
needs-review machinery stays as is — that is the part buyers keep asking for.

Deliverables are reproducible: same input, same workbook, and the audit sheet tells you
whether a re-run changed anything.
