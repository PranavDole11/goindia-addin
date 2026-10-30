"""Regenerates src/taskpane/metric-registry.json from Screener_Filter_IndAS.xlsx.

Run from the repo root after the Screener sheet changes:

    python tools/build-metric-registry.py

Why this exists
---------------
Query Data's metric dropdown used to list EVERY column name of EVERY catalog table, so one concept
showed up several times over: "Revenue", "revenue" and "Revenue (Q)" were three separate entries
the user had to choose between, and picking the wrong one silently queried a different table. The
Screener sheet is the company's own canonical registry of metrics, so it — not the raw database
columns — is what the dropdown and the query builder are standardised on.

How the sheet is read
---------------------
`Type` is the authoritative granularity signal, NOT `table_type`: a row can read
`Q | Long term Borrowings (Q) | BalanceSheetIND`, i.e. a quarterly metric whose table_type names an
annual table. So Q/QC always mean sc_quater_data; Y/C mean whatever table_type maps to.

`matching_parameter` means two different things depending on `Type`:
  * Y/Q  - ALIAS COLUMNS for the same line item, because the same concept is stored under different
           column names for different kinds of company ("Revenue" for most, "Interest Income" for a
           bank). These become a COALESCE at query time.
  * C/QC - the BASE metric the growth derives from ("Revenue Growth" -> "Revenue"). NOT an alias of
           the growth column itself, so it must never be COALESCEd in.

`future_estimate` names the consensus column backing a metric for periods that have no actuals yet.

Nothing here is trusted blindly at runtime: every column name below is only a CANDIDATE. The add-in
intersects these against the real columns it resolves from INFORMATION_SCHEMA, so an alias that
isn't a real column is dropped rather than being emitted into SQL (see resolveRegistryMetric).
"""

import collections
import json
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required: pip install openpyxl")

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE = os.path.join(REPO, "Screener_Filter_IndAS.xlsx")
TARGET = os.path.join(REPO, "src", "taskpane", "metric-registry.json")

QUARTERLY_TYPES = {"Q", "QC"}
GROWTH_TYPES = {"C", "QC"}
# table_type -> the real table it means. Only consulted for ANNUAL rows; a quarterly row is always
# sc_quater_data regardless of what table_type says (see the module docstring).
TABLE_BY_TYPE = {
    "ProfitLossIND": "sc_year_data",
    "BalanceSheetIND": "sc_year_data",
    "CashFlowIND": "sc_year_data",
    "QProfitLossIND": "sc_quater_data",
    "ratios": "ratios_ind",
    "basic_info": "basic_info",
    "technical_analysis": "technical_analysis",
    "workingcap_ana": "workingcap_ana",
}
# Same-metric-different-table ties resolve here: the annual actuals table wins, which is the rule
# for cases like "Net Debt to EBITDA" existing on both ratios_ind and sc_year_data.
TABLE_PRIORITY = ["sc_year_data", "sc_quater_data", "ratios_ind", "workingcap_ana", "basic_info",
                  "technical_analysis"]

strip_q = lambda n: re.sub(r"\s*\(Q\)\s*$", "", str(n)).strip()
norm = lambda n: re.sub(r"[\s_]+", " ", strip_q(n).lower()).strip()


def cell(v):
    """The sheet stores every value wrapped in literal single quotes ('Revenue')."""
    if v is None:
        return ""
    s = str(v).strip()
    if len(s) >= 2 and s[0] == "'" and s[-1] == "'":
        s = s[1:-1]
    return s.strip()


def dedupe(columns):
    seen, out = set(), []
    for c in columns:
        if c and c.lower() not in seen:
            seen.add(c.lower())
            out.append(c)
    return out


def read_rows():
    book = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    sheet = book["filters"]
    rows = list(sheet.iter_rows(values_only=True))
    book.close()
    header = [h for h in rows[0] if h]
    out = []
    for raw in rows[1:]:
        rec = {header[i]: cell(raw[i]) for i in range(len(header))}
        if rec.get("filter_name"):
            out.append(rec)
    return out


def table_for(rec):
    """None means 'not a plain column read' — Custom rows are formulas ({EBITDA (Q)} - {Depreciation
    (Q)}) and Holding rows are a different shape entirely. Both are left out of the registry and
    keep whatever behaviour the raw catalog already gave them."""
    if rec["Type"] in QUARTERLY_TYPES:
        return None if rec["table_type"] in ("Holding", "Custom") else "sc_quater_data"
    if rec["table_type"] == "Custom":
        return None
    return TABLE_BY_TYPE.get(rec["table_type"])


def better_table(a, b):
    rank = lambda t: TABLE_PRIORITY.index(t) if t in TABLE_PRIORITY else len(TABLE_PRIORITY)
    return a if rank(a) <= rank(b) else b


def build():
    records = read_rows()
    registry, skipped = {}, collections.Counter()

    for rec in records:
        table = table_for(rec)
        if not table:
            skipped[f"{rec['table_type']}/{rec['Type']}"] += 1
            continue
        quarterly = rec["Type"] in QUARTERLY_TYPES
        growth = rec["Type"] in GROWTH_TYPES
        canonical = strip_q(rec["filter_name"]) if quarterly else rec["filter_name"]
        entry = registry.setdefault(canonical, {"name": canonical, "unit": rec["Unit"], "growth": growth})

        aliases = [] if growth else [a.strip() for a in rec["matching_parameter"].split(",") if a.strip()]
        if quarterly:
            # The sheet lists quarterly aliases WITHOUT the "(Q)" suffix the real column carries, so
            # offer both spellings; the add-in keeps only whichever actually exists.
            candidates = [rec["filter_name"]] + [a + " (Q)" for a in aliases] + aliases
        else:
            candidates = [rec["filter_name"]] + aliases

        slot = "quarterly" if quarterly else "annual"
        if slot in entry:  # two sheet rows for one canonical+granularity — union, don't overwrite
            entry[slot]["columns"] = dedupe(entry[slot]["columns"] + candidates)
            entry[slot]["prefer"] = better_table(entry[slot]["prefer"], table)
        else:
            entry[slot] = {"prefer": table, "columns": dedupe(candidates)}
        if rec.get("future_estimate"):
            entry["estimate"] = rec["future_estimate"]

    # A quarterly-only canonical whose name is really an ALIAS of an annual one is the same concept
    # ("Profit after tax (Q)" is PAT), so fold it in rather than leaving a near-duplicate in the list.
    alias_owner = {}
    for entry in registry.values():
        if "annual" in entry:
            for column in entry["annual"]["columns"]:
                alias_owner.setdefault(norm(column), entry["name"])
    merged = []
    for name in list(registry):
        entry = registry[name]
        if "annual" in entry or "quarterly" not in entry:
            continue
        owner = alias_owner.get(norm(name))
        if not owner or owner == name:
            continue
        target = registry[owner]
        if "quarterly" in target:
            target["quarterly"]["columns"] = dedupe(target["quarterly"]["columns"] + entry["quarterly"]["columns"])
        else:
            target["quarterly"] = entry["quarterly"]
        merged.append((name, owner))
        del registry[name]

    metrics = sorted(registry.values(), key=lambda e: e["name"].lower())
    return metrics, skipped, merged


def main():
    metrics, skipped, merged = build()
    payload = {
        "_comment": ("Generated by tools/build-metric-registry.py from Screener_Filter_IndAS.xlsx "
                     "- do not hand-edit. Column names are CANDIDATES; the add-in keeps only those "
                     "that exist in the live schema."),
        "metrics": metrics,
    }
    with open(TARGET, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=1, ensure_ascii=False)
        handle.write("\n")

    both = [m for m in metrics if "annual" in m and "quarterly" in m]
    print(f"wrote {TARGET}")
    print(f"  canonical metrics      : {len(metrics)}")
    print(f"  both granularities     : {len(both)}")
    print(f"  annual only            : {len([m for m in metrics if 'quarterly' not in m])}")
    print(f"  quarterly only         : {len([m for m in metrics if 'annual' not in m])}")
    print(f"  with future_estimate   : {len([m for m in metrics if 'estimate' in m])}")
    print(f"  growth metrics         : {len([m for m in metrics if m['growth']])}")
    print(f"  merged quarterly orphans: {len(merged)}")
    print(f"  skipped (formula/holding rows): {dict(skipped)}")


if __name__ == "__main__":
    main()
