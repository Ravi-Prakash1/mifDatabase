#!/usr/bin/env python3
"""
enrich_data.py — fill missing prospect / exhibition fields via online search.

Fill-only, never overwrite:
  * Only EMPTY cells in the allow-listed fields are touched.
  * A row whose "Estimated vs Verified Flag" says "Verified" is never modified.
  * Every filled cell gets a source URL in "Source Hyperlink(s)", is marked
    "Estimated", stamped with today's date, and logged to the "Fill Audit Log"
    sheet (Company, Field, New Value, Method, Source, Date).

Enrichment uses Claude with the server-side web_search tool. If no
ANTHROPIC_API_KEY is set (or --dry-run), the script only reports which
companies/fields it WOULD enrich and writes nothing.

Per-run caps keep cost/time bounded; a full backfill happens over many days.

Usage:
    python scripts/enrich_data.py --dry-run --limit 3
    python scripts/enrich_data.py --limit 10                 # prospects only
    python scripts/enrich_data.py --limit 10 --exhibitions-limit 5
    python scripts/enrich_data.py --which exhibitions --exhibitions-limit 8
"""

import os
import re
import sys

import openpyxl

import mif_common as mc


# ─── ARGS ─────────────────────────────────────────────────────────────────────
def parse_args(argv):
    opts = {"dry_run": False, "limit": 8, "exhibitions_limit": 0, "which": "prospects"}
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--dry-run":
            opts["dry_run"] = True; i += 1
        elif a == "--limit" and i + 1 < len(argv):
            opts["limit"] = int(argv[i + 1]); i += 2
        elif a == "--exhibitions-limit" and i + 1 < len(argv):
            opts["exhibitions_limit"] = int(argv[i + 1]); i += 2
        elif a == "--which" and i + 1 < len(argv):
            opts["which"] = argv[i + 1]; i += 2
        else:
            i += 1
    # A bare --exhibitions-limit implies we also want exhibitions.
    if opts["exhibitions_limit"] and opts["which"] == "prospects":
        opts["which"] = "both"
    return opts


# ─── SHARED HELPERS ───────────────────────────────────────────────────────────
def header_map(ws, header_row=1):
    """Map column header text -> 1-based column index for an openpyxl sheet."""
    return {str(c.value).strip(): c.column
            for c in ws[header_row] if c.value is not None}


def append_source(existing, url):
    """Append a source URL to a 'Source Hyperlink(s)' cell without duplicating."""
    existing = "" if mc.is_empty(existing) else str(existing).strip()
    if not url or url in existing:
        return existing
    return (existing + "\n" + url).strip() if existing else url


def ensure_audit_sheet(wb):
    if mc.AUDIT_SHEET in wb.sheetnames:
        return wb[mc.AUDIT_SHEET]
    ws = wb.create_sheet(mc.AUDIT_SHEET)
    ws.append(mc.AUDIT_COLS)
    return ws


def audit(ws_audit, company, field, value, source):
    ws_audit.append([company, field, value,
                     "Claude web search (auto-enrich)", source, mc.today_iso()])


# ─── CLAUDE ENRICHMENT CALL ───────────────────────────────────────────────────
def enrich_company(name, context, fields):
    """Ask Claude to find values for `fields` (list of column names) for one
    company. Returns {field: {"value":..., "source":...}} for confirmed fields."""
    field_lines = "\n".join(f"  - {f}" for f in fields)
    ctx = "; ".join(f"{k}={v}" for k, v in context.items() if not mc.is_empty(v))
    prompt = f"""Find verifiable, current facts about this Indian company and fill ONLY
the fields you can confirm from a citable web source. Company: "{name}".
Known context: {ctx or "(none)"}.

Fields needed (leave out any you cannot confirm):
{field_lines}

Use web_search. Return ONLY a ```json object mapping each confirmed field name to
{{"value": "<concise value>", "source": "<url you verified it from>"}}. Do not
guess; if unsure, omit the field. Money in Rs. Cr where relevant. Example:
```json
{{"Company Website": {{"value": "example.com", "source": "https://example.com"}}}}
```"""
    try:
        data = mc.claude_json(prompt, use_web_search=True, max_tokens=3000)
    except mc.ClaudeUnavailable as e:
        raise
    except Exception as e:
        mc.log(f"  enrich call failed for {name}: {e}")
        return {}
    if not isinstance(data, dict):
        return {}
    # Keep only entries with both a value and a source.
    out = {}
    for f, v in data.items():
        if f in fields and isinstance(v, dict) and v.get("value") and v.get("source"):
            out[f] = {"value": str(v["value"]).strip(), "source": str(v["source"]).strip()}
    return out


# ─── PROSPECTS ────────────────────────────────────────────────────────────────
def select_prospect_rows(ws, hmap, limit):
    """Rank non-verified rows by how many allow-listed fields are empty."""
    flag_col = hmap.get(mc.COL_EST_FLAG)
    name_col = hmap[mc.COL_COMPANY_NAME]
    enrich_cols = [(f, hmap[f]) for f in mc.ENRICHABLE_PROSPECT_FIELDS if f in hmap]

    ranked = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if not mc.looks_like_company(name):
            continue
        if flag_col and mc.is_verified(ws.cell(r, flag_col).value):
            continue
        missing = [f for f, c in enrich_cols if mc.is_empty(ws.cell(r, c).value)]
        if missing:
            ranked.append((len(missing), r, str(name).strip(), missing))
    ranked.sort(reverse=True)
    return ranked[:limit]


def context_for(ws, hmap, row):
    out = {}
    for f in ("CIN", "GSTIN", "Company Website", "HQ Address", "Primary Segment",
              "Products", "Region"):
        c = hmap.get(f)
        if c:
            out[f] = ws.cell(row, c).value
    return out


def enrich_prospects(opts):
    wb = openpyxl.load_workbook(mc.PROSPECTS_FILE)
    ws = wb[mc.PROSPECTS_SHEET]
    hmap = header_map(ws)
    targets = select_prospect_rows(ws, hmap, opts["limit"])
    mc.log(f"Prospects: {len(targets)} companies selected for enrichment "
           f"(most-empty first, cap {opts['limit']}).")

    if opts["dry_run"]:
        for miss, r, name, fields in targets:
            print(f"  • {name} (row {r}) — {miss} empty fields: {', '.join(fields)}")
        print("\nDry-run: no Claude calls, no writes.")
        return 0

    ws_audit = ensure_audit_sheet(wb)
    src_col = hmap.get(mc.COL_SOURCE_LINKS)
    conf_col = hmap.get(mc.COL_DATA_CONFIDENCE)
    flag_col = hmap.get(mc.COL_EST_FLAG)
    date_col = hmap.get(mc.COL_LAST_VERIFIED)
    method_col = hmap.get(mc.COL_VERIF_METHOD)

    filled_total = 0
    for _, r, name, fields in targets:
        try:
            found = enrich_company(name, context_for(ws, hmap, r), fields)
        except mc.ClaudeUnavailable as e:
            mc.log(f"Claude unavailable ({e}) — stopping prospect enrichment.")
            break
        if not found:
            mc.log(f"  {name}: nothing verifiable found")
            continue
        row_filled = 0
        for field, payload in found.items():
            col = hmap.get(field)
            if not col or not mc.is_empty(ws.cell(r, col).value):
                continue                                   # never overwrite
            ws.cell(r, col).value = payload["value"]
            if src_col:
                ws.cell(r, src_col).value = append_source(
                    ws.cell(r, src_col).value, payload["source"])
            audit(ws_audit, name, field, payload["value"], payload["source"])
            row_filled += 1
        if row_filled:
            # Only stamp provenance columns that are currently empty (don't
            # clobber a human's existing note).
            if conf_col and mc.is_empty(ws.cell(r, conf_col).value):
                ws.cell(r, conf_col).value = "Low"
            if flag_col and mc.is_empty(ws.cell(r, flag_col).value):
                ws.cell(r, flag_col).value = "Estimated"
            if date_col and mc.is_empty(ws.cell(r, date_col).value):
                ws.cell(r, date_col).value = mc.today_iso()
            if method_col and mc.is_empty(ws.cell(r, method_col).value):
                ws.cell(r, method_col).value = "Auto web search"
            filled_total += row_filled
            mc.log(f"  {name}: filled {row_filled} field(s)")

    if filled_total:
        wb.save(mc.PROSPECTS_FILE)
        mc.log(f"Prospects: {filled_total} cells filled and saved.")
    else:
        mc.log("Prospects: no cells filled (nothing saved).")
    return filled_total


# ─── EXHIBITIONS ──────────────────────────────────────────────────────────────
# The parsed columns of interest (0-based index within the MASTER INDEX sheet)
# map to these human labels; enrichment only fills the sparse informational ones.
EXH_FIELDS = {
    4:  "Confirmed Dates",
    6:  "Next Edition",
    14: "Registration Website",
    15: "Organiser Contact",
    16: "Venue & Google Maps",
}


def find_exh_sheet_and_header(wb):
    sheet = None
    for s in wb.sheetnames:
        u = s.upper()
        if "MASTER INDEX" in u and "COMPLETE" in u:
            sheet = s; break
    if sheet is None:
        for s in wb.sheetnames:
            if "MASTER INDEX" in s.upper():
                sheet = s; break
    if sheet is None:
        return None, None
    ws = wb[sheet]
    hdr = None
    for i in range(1, min(9, ws.max_row + 1)):
        vals = [str(c.value) for c in ws[i]]
        if any("Exhibition / Conference" in v for v in vals):
            hdr = i; break
    return ws, (hdr or 3)


def enrich_exhibitions(opts):
    if not os.path.exists(mc.EXHIBITIONS_FILE):
        mc.log("Exhibitions file not found — skipping.")
        return 0
    wb = openpyxl.load_workbook(mc.EXHIBITIONS_FILE)
    ws, hdr = find_exh_sheet_and_header(wb)
    if ws is None:
        mc.log("Exhibitions master index sheet not found — skipping.")
        return 0

    name_col = 2  # column B = "Exhibition / Conference"
    targets = []
    for r in range(hdr + 1, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if not mc.looks_like_company(name):
            continue
        missing = [(idx, lbl) for idx, lbl in EXH_FIELDS.items()
                   if mc.is_empty(ws.cell(r, idx + 1).value)]
        if missing:
            targets.append((len(missing), r, str(name).split("\n")[0].strip(), missing))
    targets.sort(reverse=True)
    targets = targets[: opts["exhibitions_limit"]]
    mc.log(f"Exhibitions: {len(targets)} events selected (cap {opts['exhibitions_limit']}).")

    if opts["dry_run"]:
        for miss, r, name, fields in targets:
            print(f"  • {name} (row {r}) — missing: {', '.join(l for _, l in fields)}")
        print("\nDry-run: no Claude calls, no writes.")
        return 0

    ws_audit = ensure_audit_sheet(wb)
    country_col = 4  # column D = "Country / City" for context
    filled_total = 0
    for _, r, name, fields in targets:
        ctx = {"Location": ws.cell(r, country_col).value}
        want = [lbl for _, lbl in fields]
        try:
            found = enrich_company(name + " (trade exhibition / conference)", ctx, want)
        except mc.ClaudeUnavailable as e:
            mc.log(f"Claude unavailable ({e}) — stopping exhibition enrichment.")
            break
        if not found:
            continue
        label_to_idx = {lbl: idx for idx, lbl in EXH_FIELDS.items()}
        row_filled = 0
        for label, payload in found.items():
            idx = label_to_idx.get(label)
            if idx is None or not mc.is_empty(ws.cell(r, idx + 1).value):
                continue
            ws.cell(r, idx + 1).value = payload["value"]
            audit(ws_audit, name, label, payload["value"], payload["source"])
            row_filled += 1
        if row_filled:
            filled_total += row_filled
            mc.log(f"  {name}: filled {row_filled} field(s)")

    if filled_total:
        wb.save(mc.EXHIBITIONS_FILE)
        mc.log(f"Exhibitions: {filled_total} cells filled and saved.")
    return filled_total


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    opts = parse_args(sys.argv[1:])
    mc.log(f"enrich_data — which={opts['which']} limit={opts['limit']} "
           f"exhibitions_limit={opts['exhibitions_limit']} dry_run={opts['dry_run']}")

    total = 0
    if opts["which"] in ("prospects", "both"):
        total += enrich_prospects(opts)
    if opts["which"] in ("exhibitions", "both") and opts["exhibitions_limit"]:
        total += enrich_exhibitions(opts)

    mc.log(f"Done. {total} cell(s) filled." if not opts["dry_run"] else "Dry-run complete.")


if __name__ == "__main__":
    main()
