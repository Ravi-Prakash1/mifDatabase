#!/usr/bin/env python3
"""
mif_context.py — the MIF business framework, in one editable place.

Encodes who MIF's customers are, the end-use segment taxonomy, and how the BD
funnel ranks companies. Both enrich_data.py (priority order) and
discover_companies.py (segment classification + relevance) import from here.

Most of this is DERIVED from the live Master Database at runtime (so it stays
fresh as the sheet grows); the segment keyword map is the one hand-curated part
you may want to extend.
"""

import re
import pandas as pd

import mif_common as mc

# Relationship-status values that mark an existing customer / partner (not a
# cold prospect). Matched case-insensitively as substrings.
CUSTOMER_STATUSES = ("existing customer", "active oem", "strategic partner")

# Competitor end-use segments — mirror generate_portal_v7_News.py so a
# discovered competitor is treated consistently.
COMPETITOR_SEGMENTS = {
    "Pipes & Tubes", "Machine Makers", "Roll Forming Competitors", "Rolls & Dies",
}

# Canonical segment -> keyword set. Used to (a) classify a news item / newly
# discovered company into a segment, and (b) build discovery queries. Keep the
# canonical key exactly as it should appear in the sheet's "Segment" column.
SEGMENT_KEYWORDS = {
    "Auto Components":          ("auto component", "auto part", "ancillary", "tier-1 supplier"),
    "Auto - Passenger (LMV)":   ("passenger vehicle", "car maker", "pv oem", "hatchback", "suv"),
    "Auto - Commercial (LCV_MCV)": ("commercial vehicle", "lcv", "mcv", "truck maker", "cv oem"),
    "Bus Body":                 ("bus body", "bus builder", "coach builder", "e-bus", "bus maker"),
    "Trucks & Trailers":        ("trailer", "tipper", "truck body", "haulage"),
    "Railways":                 ("railway", "rail coach", "metro coach", "wagon", "rolling stock"),
    "Construction Equipment":   ("construction equipment", "excavator", "backhoe", "earthmoving", "jcb"),
    "Agricultural Machinery":   ("tractor", "agricultural machinery", "farm equipment", "harvester"),
    "Elevator & Escalator":     ("elevator", "escalator", "lift manufacturer"),
    "Mounting Structures":      ("module mounting", "solar mounting", "solar structure", "racking"),
    "Wind Energy":              ("wind turbine", "wind energy", "windmill", "nacelle"),
    "Power Transmission":       ("transmission tower", "power transmission", "substation structure"),
    "HVAC":                     ("hvac", "air handling", "ducting", "chiller", "ahu"),
    "Home Appliances":          ("home appliance", "white goods", "refrigerator", "washing machine"),
    "Steel Furniture":          ("steel furniture", "office furniture", "modular furniture"),
    "Storage Technology":       ("storage rack", "warehouse racking", "pallet rack", "mezzanine rack"),
    "Mezzanine Floor":          ("mezzanine", "mezzanine floor"),
    "Formwork & Scaffolding":   ("formwork", "scaffolding", "shuttering"),
    "Windows & Facades":        ("facade", "curtain wall", "window system", "fenestration"),
    "Fencing":                  ("fencing", "crash barrier", "guard rail"),
    "Road Safety":              ("crash barrier", "road safety", "guardrail", "highway barrier"),
    "Greenhouse":               ("greenhouse", "polyhouse", "protected cultivation"),
    "Food Processing":          ("food processing", "cold storage", "dairy equipment"),
    "Textile Machinery":        ("textile machinery", "loom", "spinning machine"),
    "Cranes & Hoists":          ("crane", "hoist", "eot crane", "gantry"),
    "ISO Freight Containers":   ("freight container", "iso container", "shipping container"),
    "Cleaning Equipment":       ("cleaning equipment", "sweeping machine"),
    "Shipbuilding":             ("shipyard", "shipbuilding", "vessel fabrication"),
    "Sun Shading":              ("sun shading", "louvre", "pergola"),
    "Interior Fittings":        ("interior fitting", "partition system", "false ceiling"),
    "Metal Hooks & Hardware":   ("display hook", "metal hardware", "fastener"),
    "Pipes & Tubes":            ("pipe maker", "tube maker", "steel tube", "erw pipe", "gi pipe"),
    "Machine Makers":           ("roll forming machine", "rollforming machine", "press brake maker"),
}


# ─── CUSTOMERS (derived from the live sheet) ──────────────────────────────────
def load_customers(prospects_file=None):
    """Return a DataFrame of the customer/partner rows (Existing Customer /
    Active OEM / Strategic Partner)."""
    path = prospects_file or mc.PROSPECTS_FILE
    df = pd.read_excel(path, sheet_name=mc.PROSPECTS_SHEET)
    pattern = "|".join(re.escape(s) for s in CUSTOMER_STATUSES)
    mask = df["Relationship Status"].astype(str).str.contains(pattern, case=False, na=False)
    return df[mask]


def customer_names(prospects_file=None):
    try:
        return sorted({str(n).strip() for n in load_customers(prospects_file)["Company Name"]
                       if not mc.is_empty(n)})
    except Exception as e:
        mc.log(f"customer_names: {e}")
        return []


# ─── SEGMENT NORMALISATION & CLASSIFICATION ───────────────────────────────────
def normalize_segment(seg):
    """Fold punctuation/plural variants so 'Auto — Passenger (LMV)' and
    'Auto - Passenger (LMV)' collapse to one canonical string."""
    if mc.is_empty(seg):
        return ""
    s = re.sub(r"[—–]", "-", str(seg)).strip()
    s = re.sub(r"\s+", " ", s)
    # Map a few known aliases onto the canonical keys.
    aliases = {
        "Elevators & Escalators": "Elevator & Escalator",
        "Greenhouse Structures": "Greenhouse",
        "Bus Body Building": "Bus Body",
        "Auto - Commercial (LCV/MCV)": "Auto - Commercial (LCV_MCV)",
    }
    return aliases.get(s, s)


def classify_segment(text):
    """Return the best-matching canonical segment for a blob of text, or '' if
    none of the keyword sets match."""
    t = (text or "").lower()
    best, best_hits = "", 0
    for seg, kws in SEGMENT_KEYWORDS.items():
        hits = sum(1 for k in kws if k in t)
        if hits > best_hits:
            best, best_hits = seg, hits
    return best


# ─── BD VALUE RANKING ─────────────────────────────────────────────────────────
def bd_score(row):
    """Rank a company row by business-development value (higher = enrich first).
    `row` is a dict-like mapping of column -> value."""
    def g(k):
        return str(row.get(k, "") or "").lower()

    rel, ctier, ptier = g("Relationship Status"), g("Competitor Tier"), g("Priority Tier")
    if any(k in rel for k in CUSTOMER_STATUSES):
        return 100
    if "t1 captive" in ctier or "t1 oem" in ctier or "tier 1" in ptier or "tier-1" in ptier:
        return 80
    if "active prospect" in ctier or "active prospect" in rel:
        return 60
    if "t2 conversion" in ctier or "tier 2" in ptier or "tier-2" in ptier:
        return 40
    if "pure prospect" in ctier:
        return 20
    return 10


if __name__ == "__main__":
    print("Customers/partners:", len(customer_names()))
    for n in customer_names():
        print("  -", n)
    print("\nSegments with keyword rules:", len(SEGMENT_KEYWORDS))
