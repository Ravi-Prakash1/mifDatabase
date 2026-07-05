#!/usr/bin/env python3
"""
MIF MARKET RESEARCH — Portal Generator (v4, dual-login)
=======================================================
By Ravi Prakash.

Reads the MIF Master Database Excel (MERGED, 56 columns) and generates ONE
self-contained offline HTML portal with TWO separate, password-gated views:

  * Password  "MIF2026"      -> Market view: customers + BD prospects.
                               Competitor-class companies are NOT shown here.
  * Password  "Confidential" -> Competitor Intelligence view: ONLY
                               Competitors, Machine Makers, Pipe & Tube makers,
                               Rolls & Dies makers, and Roll-Forming competitors.

The split is decided by the "Segment" column (Pipes & Tubes / Machine Makers /
Roll Forming Competitors / Rolls & Dies) plus any row flagged "Competitor" in
Relationship Status or Competitor Tier. Everything else is the Market view.

Company logos are pulled live from the "Company Website" column (favicon service).
If a mark can't load it hides itself cleanly (no broken-image icon, no error).

Usage:
    python generate_portal_v4.py
    python generate_portal_v4.py --input MyFile.xlsx --output Portal.html
    python generate_portal_v4.py --mif-password Alpha --comp-password Beta
    python generate_portal_v4.py --accent "#4f8ff7"

Requires:  pip install pandas openpyxl
"""

import pandas as pd
import json
import hashlib
import re
import sys
import os
from datetime import datetime

# ─── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE     = "MIF_Master_Database_MERGED_04July2026.xlsx"
OUTPUT_FILE    = "MIF_Market_Research_Portal.html"
PASSWORD_MIF   = "MIF2026"       # -> customers + prospects (competitors hidden)
PASSWORD_COMP  = "Confidential"  # -> competitors / machine / pipe&tube / rolls&dies only
ACCENT         = "#ff8a3d"       # molten copper. Try "#4f8ff7", "#35d08a", "#c9a24b"
BRAND          = "MIF MARKET RESEARCH"
BYLINE         = "BY RAVI PRAKASH"

# Segments that belong to the Competitor Intelligence (Confidential) view.
COMPETITOR_SEGMENTS = {
    "Pipes & Tubes", "Machine Makers", "Roll Forming Competitors", "Rolls & Dies",
}

# Columns that denote a company's OWN classification. If the whole word
# "competitor" appears in ANY of these, the company is treated as a competitor
# and routed to the Confidential view (this is what catches every
# "Tier 1 Competitor", "Tier 2 Competitor", "Direct Competitor", etc.).
DESIGNATION_FIELDS = [
    "Priority Tier", "Competitor Tier", "Relationship Status", "Primary Segment",
]
# Free-text notes are NOT swept in full (a note like "faces competitor X" or
# "can displace competitor" must NOT hide a genuine prospect). But a note that
# LEADS with a competitor designation ("COMPETITOR — they make…",
# "Direct roll forming competitor…") does count.
_LEAD_COMPETITOR = re.compile(r'^[\s\W]*(direct\s+|partial\s+)?(roll[\s-]*forming\s+)?competitor\b', re.I)
_WORD_COMPETITOR = re.compile(r'\bcompetitor\b', re.I)
# ---------------------------------------------------------------------------
# ─────────────────────────────────────────────────────────────────────────────

def parse_args():
    global PASSWORD_MIF, PASSWORD_COMP, EXCEL_FILE, OUTPUT_FILE, ACCENT, BRAND
    args = sys.argv[1:]; i = 0
    while i < len(args):
        a = args[i]
        if a == '--mif-password' and i+1 < len(args):   PASSWORD_MIF = args[i+1]; i += 2
        elif a == '--comp-password' and i+1 < len(args): PASSWORD_COMP = args[i+1]; i += 2
        elif a == '--input' and i+1 < len(args):         EXCEL_FILE = args[i+1]; i += 2
        elif a == '--output' and i+1 < len(args):        OUTPUT_FILE = args[i+1]; i += 2
        elif a == '--accent' and i+1 < len(args):        ACCENT = args[i+1]; i += 2
        elif a == '--brand' and i+1 < len(args):         BRAND = args[i+1]; i += 2
        else: i += 1

def sha256(pwd):
    return hashlib.sha256(pwd.encode('utf-8')).hexdigest()

def clean(v):
    if v is None: return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'none', 'nat', '<na>') else s

def is_competitor_side(c):
    seg  = c.get('Segment', '').strip()
    if seg in COMPETITOR_SEGMENTS:
        return True
    # Any designation column that carries the "competitor" keyword.
    for f in DESIGNATION_FIELDS:
        if _WORD_COMPETITOR.search(str(c.get(f, ''))):
            return True
    # Notes / entry route that LEAD with a competitor designation.
    for f in ('BD Notes', 'Entry Route'):
        if _LEAD_COMPETITOR.match(str(c.get(f, '')).strip()):
            return True
    return False

def mif_category(c):
    """Customer vs Prospect classification for the Market view."""
    rel = c.get('Relationship Status', '').lower()
    if 'existing customer' in rel or 'active oem' in rel:
        return 'Customer'
    return 'Prospect'

def read_companies():
    print(f"Reading: {EXCEL_FILE}")
    xl = pd.ExcelFile(EXCEL_FILE)
    sheet = ('Master Database' if 'Master Database' in xl.sheet_names
             else ('Combined Master' if 'Combined Master' in xl.sheet_names else xl.sheet_names[0]))
    print(f"Sheet: {sheet}")
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=0)
    cols = [str(c) for c in df.columns]

    def valid_name(x):
        if pd.isna(x): return False
        s = str(x).strip()
        return bool(s) and s.lower() not in ('nan', 'none', 'nat', '<na>', 'company name') and not s.startswith('\u25bc')

    rows = df[df['Company Name'].apply(valid_name)].copy()
    mif, comp = [], []
    for _, row in rows.iterrows():
        c = {col: clean(row[col]) for col in cols}
        if is_competitor_side(c):
            c['_cat'] = 'Competitor'
            comp.append(c)
        else:
            c['_cat'] = mif_category(c)
            mif.append(c)
    print(f"Market view (MIF2026):        {len(mif)} companies "
          f"({sum(1 for c in mif if c['_cat']=='Customer')} customers, "
          f"{sum(1 for c in mif if c['_cat']=='Prospect')} prospects)")
    print(f"Competitor view (Confidential): {len(comp)} companies")
    return mif, comp, cols

def generate(mif, comp, cols, ts):
    repl = {
        '__DATA_MIF__':  json.dumps(mif,  ensure_ascii=False, separators=(',', ':')),
        '__DATA_COMP__': json.dumps(comp, ensure_ascii=False, separators=(',', ':')),
        '__COLS__':      json.dumps(cols, ensure_ascii=False),
        '__TS__':        ts,
        '__PH_MIF__':    sha256(PASSWORD_MIF),
        '__PH_COMP__':   sha256(PASSWORD_COMP),
        '__ACCENT__':    ACCENT,
        '__BRAND__':     BRAND,
        '__BYLINE__':    BYLINE,
    }
    html = HTML_TEMPLATE
    for k, v in repl.items():
        html = html.replace(k, v)
    return html


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MIF Market Research</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --accent:__ACCENT__;
  --bg:#080b12; --panel:#0a0f1c; --card:#0c1220; --card2:#0e1524;
  --border:#182236; --border2:#1c2740;
  --t1:#eef3fb; --t2:#93a1bd; --t3:#5f6e8c; --t4:#465370;
  --steel:#4f8ff7; --green:#35d08a; --red:#ff6274; --amber:#ffc24d; --purple:#9d7bff;
  --sid:250px; --head:60px;
  --sans:'IBM Plex Sans',system-ui,'Segoe UI',sans-serif;
  --disp:'Space Grotesk','IBM Plex Sans',system-ui,sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,Consolas,monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);background:var(--bg);color:var(--t1);font-size:14px;overflow-x:hidden}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#233046;border-radius:8px}
::-webkit-scrollbar-thumb:hover{background:#334463}
input::placeholder{color:#54617d}
button{font-family:inherit}
a{color:var(--steel)}
@keyframes fade{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes glow{0%,100%{opacity:.45}50%{opacity:.8}}

/* LOGIN */
#login{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;
  background:radial-gradient(1200px 700px at 50% -10%,#141f36 0%,#0a0f1c 55%,#060911 100%);overflow:hidden}
.login-glow{position:absolute;width:520px;height:520px;border-radius:50%;
  background:radial-gradient(circle,color-mix(in srgb,var(--accent) 14%,transparent) 0%,transparent 70%);
  filter:blur(24px);animation:glow 6s ease-in-out infinite;pointer-events:none}
.login-box{position:relative;width:min(92vw,430px);background:linear-gradient(180deg,#0f1626,#0c1120);
  border:1px solid #1d2942;border-radius:18px;padding:44px 42px 26px;
  box-shadow:0 40px 90px rgba(0,0,0,.55),inset 0 1px 0 rgba(255,255,255,.04);animation:fade .6s ease both}
.login-brand{display:flex;align-items:center;gap:12px;margin-bottom:26px}
.glyph{width:38px;height:38px;position:relative;flex:none}
.glyph:before{content:"";position:absolute;inset:0;border-radius:10px;
  background:linear-gradient(135deg,var(--accent),#c85a1e);transform:rotate(45deg)}
.glyph:after{content:"";position:absolute;inset:11px;border-radius:3px;background:#0c1120;transform:rotate(45deg)}
.login-name{font-family:var(--disp);font-weight:700;font-size:16px;letter-spacing:.5px;color:var(--t1);line-height:1}
.login-tag{font-size:9.5px;letter-spacing:2px;color:var(--accent);font-weight:600;margin-top:4px}
.login-h{font-family:var(--disp);font-size:22px;font-weight:600;color:var(--t1);letter-spacing:-.3px}
.login-sub{font-size:12.5px;color:#7382a0;margin:6px 0 24px;line-height:1.5}
.login-lbl{font-size:10px;letter-spacing:1.5px;color:#61708e;font-weight:600;text-transform:uppercase}
#pwd{width:100%;margin-top:8px;background:#0a0f1c;border:1.5px solid #223052;border-radius:10px;
  padding:13px 15px;color:var(--t1);font-size:15px;font-family:var(--mono);letter-spacing:3px;outline:none;transition:border .2s}
#pwd:focus{border-color:var(--accent)}
.login-btn{width:100%;margin-top:16px;background:linear-gradient(135deg,var(--accent),#e06a24);
  color:#160a02;border:none;border-radius:10px;padding:14px;font-size:14px;font-weight:700;
  font-family:var(--disp);letter-spacing:.4px;cursor:pointer;transition:transform .15s,box-shadow .2s;
  box-shadow:0 10px 30px color-mix(in srgb,var(--accent) 25%,transparent)}
.login-btn:hover{transform:translateY(-1px)}
.login-err{min-height:18px;margin-top:12px;font-size:12.5px;color:#ff6274;text-align:center;font-weight:500}
.login-conf{margin-top:14px;padding-top:16px;border-top:1px solid #182135;display:flex;align-items:center;
  gap:7px;justify-content:center;font-size:10.5px;color:#54617d;letter-spacing:.3px}
.login-conf b{color:var(--accent);font-weight:400}
.login-copy{margin-top:10px;font-size:9px;color:#3d4761;text-align:center;line-height:1.6;letter-spacing:.2px}

/* SHELL */
#app{display:none}
#header{position:fixed;top:0;left:0;right:0;height:var(--head);z-index:100;display:flex;align-items:center;
  gap:18px;padding:0 22px;background:rgba(10,15,26,.82);backdrop-filter:blur(14px);border-bottom:1px solid #172035}
.hbrand{display:flex;align-items:center;gap:11px;flex:none}
.glyph-sm{width:30px;height:30px;position:relative}
.glyph-sm:before{content:"";position:absolute;inset:0;border-radius:8px;
  background:linear-gradient(135deg,var(--accent),#c85a1e);transform:rotate(45deg)}
.glyph-sm:after{content:"";position:absolute;inset:9px;border-radius:2px;background:#0a0f1c;transform:rotate(45deg)}
.hname{font-family:var(--disp);font-weight:700;font-size:15px;letter-spacing:.4px}
.hchip{font-size:9px;letter-spacing:2px;font-weight:700;border-radius:5px;padding:2px 7px;
  color:var(--accent);border:1px solid color-mix(in srgb,var(--accent) 40%,transparent)}
.hchip.comp{color:var(--red);border-color:color-mix(in srgb,var(--red) 45%,transparent)}
.hsearch{position:relative;flex:1;max-width:460px}
.hsearch input{width:100%;background:#0d1424;border:1px solid var(--border2);border-radius:9px;
  padding:9px 14px 9px 34px;color:var(--t1);font-size:13px;outline:none;transition:border .2s}
.hsearch input:focus{border-color:color-mix(in srgb,var(--accent) 40%,transparent)}
.hsearch .ic{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:#54617d;font-size:14px}
.hsearch .kbd{position:absolute;right:11px;top:50%;transform:translateY(-50%);font-size:10px;color:#4a566f;
  font-family:var(--mono);border:1px solid #24304a;border-radius:4px;padding:1px 5px}
.hright{margin-left:auto;display:flex;align-items:center;gap:16px;flex:none}
.hts{display:flex;flex-direction:column;align-items:flex-end;line-height:1.2}
.hts .k{font-size:9px;letter-spacing:1.5px;color:#4a566f;font-weight:600}
.hts .v{font-size:11.5px;color:var(--t2);font-family:var(--mono)}
.hlogout{background:#131c30;border:1px solid #223052;color:var(--t2);padding:7px 14px;border-radius:8px;
  font-size:12px;font-weight:600;cursor:pointer;transition:all .18s}
.hlogout:hover{border-color:#ff6274;color:#ff6274}
.menu-btn{display:none;background:none;border:none;color:var(--t1);font-size:22px;cursor:pointer}

#sidebar{position:fixed;top:var(--head);left:0;width:var(--sid);height:calc(100vh - var(--head));
  overflow-y:auto;background:var(--panel);border-right:1px solid #141d31;z-index:90;padding:14px 0 30px;transition:transform .25s}
.sid-sec{padding:0 14px;margin-top:14px}
.sid-sec:first-child{margin-top:0}
.sid-lbl{font-size:9px;letter-spacing:2px;color:var(--t4);font-weight:700;padding:8px 8px 6px}
.sid-item{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;cursor:pointer;
  font-size:13px;font-weight:500;margin-bottom:2px;transition:all .15s;color:var(--t2);border-left:2px solid transparent}
.sid-item:hover{background:#111a2d;color:var(--t1)}
.sid-item.active{background:color-mix(in srgb,var(--accent) 11%,transparent);color:var(--t1);border-left-color:var(--accent)}
.sid-item .ico{width:16px;text-align:center;font-size:14px}
.sid-item .dot{width:7px;height:7px;border-radius:50%;flex:none}
.sid-item .nm{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sid-badge{font-family:var(--mono);font-size:10.5px;color:var(--t2);background:#131c30;border-radius:20px;padding:2px 8px}
.sid-item.active .sid-badge{background:color-mix(in srgb,var(--accent) 13%,transparent);color:var(--accent)}
.sid-seg{font-size:12px;padding:7px 12px}
.sid-seg .sc{font-family:var(--mono);font-size:10px;color:#5f6e8c}

#main{margin-left:var(--sid);padding-top:var(--head);min-height:100vh}

/* DASHBOARD */
#view-dash{padding:30px 34px 40px;animation:fade .5s ease both}
.dash-head{display:flex;align-items:flex-end;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:26px}
.eyebrow{font-size:10px;letter-spacing:2.5px;color:var(--accent);font-weight:700;margin-bottom:7px}
.dash-title{font-family:var(--disp);font-size:28px;font-weight:600;letter-spacing:-.6px;color:var(--t1)}
.dash-sub{font-size:13px;color:#7382a0;margin-top:6px;max-width:640px}
.synced{display:flex;align-items:center;gap:8px;background:#0d1424;border:1px solid var(--border2);border-radius:10px;padding:9px 14px;font-size:11.5px;color:var(--t2)}
.synced .live{width:8px;height:8px;border-radius:50%;background:#35d08a;box-shadow:0 0 10px #35d08a;animation:glow 2s infinite}
.synced .v{color:var(--t1);font-family:var(--mono)}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px;margin-bottom:26px}
.kpi{position:relative;overflow:hidden;background:linear-gradient(160deg,#0f1728,#0c1220);
  border:1px solid var(--border);border-radius:14px;padding:18px 20px;cursor:pointer;transition:all .2s}
.kpi:hover{transform:translateY(-2px)}
.kpi .halo{position:absolute;top:-30px;right:-30px;width:90px;height:90px;border-radius:50%}
.kpi .top{display:flex;align-items:center;gap:8px;margin-bottom:12px}
.kpi .sq{width:9px;height:9px;border-radius:3px}
.kpi .lbl{font-size:10.5px;letter-spacing:1px;color:#7382a0;font-weight:600;text-transform:uppercase}
.kpi .num{font-family:var(--disp);font-size:36px;font-weight:700;letter-spacing:-1px;color:var(--t1);line-height:1}
.kpi .sub{font-size:11.5px;color:#5f6e8c;margin-top:6px}
.charts{display:grid;grid-template-columns:1.5fr 1fr;gap:16px;margin-bottom:16px}
.panel-card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:22px}
.pc-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px}
.pc-title{font-family:var(--disp);font-size:15px;font-weight:600;color:var(--t1)}
.pc-meta{font-size:10.5px;color:#5f6e8c;font-family:var(--mono)}
.bar-row{display:flex;align-items:center;gap:12px;margin-bottom:9px}
.bar-row.clk{cursor:pointer}
.bar-row.clk:hover{opacity:.82}
.bar-lbl{width:150px;font-size:12px;color:#a7b4cd;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-lbl.sm{width:120px;font-size:11.5px}
.bar-lbl.xs{width:80px;font-size:11.5px}
.bar-track{flex:1;height:9px;background:#131c30;border-radius:6px;overflow:hidden}
.bar-track.sm{height:8px}
.bar-fill{height:100%;border-radius:6px;width:0;transition:width 1.1s cubic-bezier(.2,.8,.2,1)}
.bar-num{width:40px;text-align:right;font-family:var(--mono);font-size:12px;color:var(--t1);font-weight:600}
.col-2{display:flex;flex-direction:column;gap:16px}
.strip{background:linear-gradient(120deg,#0f1728,#0c1220);border:1px solid var(--border);border-radius:14px;padding:22px 24px}
.strip-row{display:flex;gap:10px;flex-wrap:wrap}
.strip-cell{flex:1;min-width:150px;background:#0a0f1c;border:1px solid #1a2439;border-radius:10px;padding:14px 16px}
.strip-cell .n{font-family:var(--mono);font-size:26px;font-weight:600;color:var(--t1)}
.strip-cell .l{font-size:11.5px;color:#8493b0;margin-top:4px}

/* LIST */
#view-list{display:none;animation:fade .35s ease both}
.toolbar{position:sticky;top:var(--head);z-index:80;background:rgba(10,15,26,.9);backdrop-filter:blur(12px);
  border-bottom:1px solid #141d31;padding:16px 34px;display:flex;flex-wrap:wrap;gap:12px;align-items:center}
.tb-title{font-family:var(--disp);font-size:17px;font-weight:600;color:var(--t1);margin-right:auto}
.sel{background:#0d1424;border:1px solid var(--border2);border-radius:8px;padding:8px 12px;color:#a7b4cd;font-size:12px;cursor:pointer;outline:none}
.results{padding:14px 34px 6px;display:flex;align-items:center;justify-content:space-between}
.results .c{font-size:12px;color:#7382a0}
.results .c b{color:var(--t1);font-weight:600;font-family:var(--mono)}
.results .n{font-size:11px;color:#5f6e8c}
.company-list{padding:8px 34px 20px;display:flex;flex-direction:column;gap:10px}
.card{position:relative;background:linear-gradient(120deg,#0e1524,#0b111e);border:1px solid #172035;
  border-radius:13px;padding:16px 20px 16px 22px;cursor:pointer;transition:all .16s;overflow:hidden}
.card:hover{border-color:color-mix(in srgb,var(--accent) 40%,transparent);transform:translateY(-2px);box-shadow:0 12px 30px rgba(0,0,0,.35)}
.card .fitbar{position:absolute;left:0;top:0;bottom:0;width:3px}
.card .row{display:flex;align-items:flex-start;gap:14px}
.card .num{font-family:var(--mono);font-size:11px;color:var(--t4);padding-top:3px;min-width:26px}
.card .body{flex:1;min-width:0}
.card .n1{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.card .cname{font-family:var(--disp);font-size:15.5px;font-weight:600;color:var(--t1)}
.card .rev{font-family:var(--mono);font-size:12px;color:var(--accent);
  background:color-mix(in srgb,var(--accent) 12%,transparent);border-radius:6px;padding:2px 8px}
.card .tags{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
.card .notes{font-size:12px;color:#6f7d99;margin-top:9px;line-height:1.5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.card .chev{align-self:center;color:#3a4560;font-size:18px;flex:none}
.tag{font-size:10.5px;font-weight:600;border-radius:6px;padding:3px 9px;white-space:nowrap}
.pager{display:flex;justify-content:center;align-items:center;gap:8px;padding:10px 20px 44px;flex-wrap:wrap}
.pg{min-width:36px;padding:8px 12px;border-radius:8px;border:1px solid #223052;background:#0d1424;
  color:#a7b4cd;font-size:12.5px;font-family:var(--mono);cursor:pointer;transition:all .15s}
.pg:hover{border-color:var(--accent);color:var(--accent)}
.pg.active{background:var(--accent);color:#160a02;border-color:var(--accent)}
.pg-info{font-size:12px;color:#5f6e8c;margin-left:6px}
.no-res{text-align:center;padding:80px 20px;color:#5f6e8c}
.no-res .ic{font-size:44px;margin-bottom:14px;opacity:.5}
.no-res .m{font-family:var(--disp);font-size:16px;font-weight:600;color:var(--t2);margin-bottom:6px}

/* DETAIL DRAWER */
#overlay{display:none;position:fixed;inset:0;z-index:200;background:rgba(4,7,13,.55);backdrop-filter:blur(3px)}
#drawer{position:fixed;top:0;right:0;bottom:0;width:min(96vw,720px);z-index:210;background:var(--panel);
  border-left:1px solid #1a2439;box-shadow:-20px 0 60px rgba(0,0,0,.5);overflow-y:auto;transform:translateX(100%);
  transition:transform .3s cubic-bezier(.2,.8,.2,1)}
#drawer.open{transform:none}
.dp-head{position:sticky;top:0;z-index:5;background:linear-gradient(135deg,#111c33,#0c1425);
  border-bottom:1px solid #1a2439;padding:24px 28px}
.dp-close{position:absolute;top:20px;right:22px;width:32px;height:32px;border-radius:9px;background:#182338;
  border:1px solid #26344f;color:#a7b4cd;font-size:15px;cursor:pointer;transition:all .18s}
.dp-close:hover{background:#ff6274;color:#fff;border-color:#ff6274}
.dp-eyebrow{font-size:10px;letter-spacing:2px;color:var(--accent);font-weight:700;margin-bottom:8px}
.dp-name{font-family:var(--disp);font-size:23px;font-weight:600;letter-spacing:-.4px;color:var(--t1);padding-right:40px;line-height:1.2}
.dp-tags{display:flex;flex-wrap:wrap;gap:6px;margin-top:12px}
.dp-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:#141d31;border-bottom:1px solid #1a2439}
.dp-stat{background:#0b111e;padding:16px 14px}
.dp-stat .k{font-size:9px;letter-spacing:1px;color:#5f6e8c;font-weight:600;text-transform:uppercase}
.dp-stat .v{font-family:var(--disp);font-size:19px;font-weight:600;margin-top:5px;color:var(--t1)}
.dp-body{padding:24px 28px 60px}
.dp-logo{display:flex;align-items:center;gap:16px;background:#0e1524;border:1px solid var(--border);
  border-radius:12px;padding:14px 18px;margin-bottom:24px}
.dp-logo-tile{width:64px;height:64px;flex:none;border-radius:12px;background:#fff;display:flex;
  align-items:center;justify-content:center;padding:8px;overflow:hidden}
.dp-logo-tile img{max-width:100%;max-height:100%;object-fit:contain}
.dp-logo .k{font-size:9px;letter-spacing:1.5px;color:#5f6e8c;font-weight:700}
.dp-logo .nm{font-family:var(--disp);font-size:15px;font-weight:600;color:var(--t1);margin-top:4px}
.dp-logo .dm{font-size:11px;color:#5f6e8c;font-family:var(--mono);margin-top:2px}
.dp-sec{margin-bottom:26px}
.dp-sec-head{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.dp-sec-bar{width:5px;height:16px;border-radius:3px}
.dp-sec-title{font-family:var(--disp);font-size:12px;letter-spacing:1.5px;color:var(--t2);font-weight:600;text-transform:uppercase}
.dp-sec-line{flex:1;height:1px;background:#141d31}
.dp-grid{display:grid;grid-template-columns:1fr 1fr;gap:1px;background:#131c30;border-radius:10px;overflow:hidden}
.dp-field{background:#0b111e;padding:12px 15px}
.dp-field.full{grid-column:1/-1}
.dp-key{font-size:10px;letter-spacing:.5px;color:#5f6e8c;font-weight:600;margin-bottom:4px;text-transform:uppercase}
.dp-val{font-size:13px;color:#c3cee2;line-height:1.5;word-break:break-word}
.dp-val a{text-decoration:none}
.hl-H{color:#4fe0a0;font-weight:700}
.hl-M{color:#ffc24d;font-weight:700}
.hl-L{color:#ff8a97;font-weight:700}

@media(max-width:900px){
  :root{--sid:0px}
  #sidebar{transform:translateX(-100%);width:262px}
  #sidebar.open{transform:none;box-shadow:4px 0 24px rgba(0,0,0,.5)}
  .menu-btn{display:block}
  #main{margin-left:0}
  .hsearch{display:none}
  .charts{grid-template-columns:1fr}
  .dp-stats{grid-template-columns:repeat(2,1fr)}
  #drawer{width:100%}
}
@media(max-width:560px){
  .hts{display:none}
  .kpi-grid{grid-template-columns:1fr 1fr}
  .dp-grid{grid-template-columns:1fr}
  .toolbar,.results,.company-list,#view-dash{padding-left:16px;padding-right:16px}
}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="login">
  <div class="login-glow"></div>
  <div class="login-box">
    <div class="login-brand">
      <div class="glyph"></div>
      <div>
        <div class="login-name">__BRAND__</div>
        <div class="login-tag">__BYLINE__</div>
      </div>
    </div>
    <div class="login-h">Secure Sign-In</div>
    <div class="login-sub">Mother India Forming &middot; Intelligence Terminal</div>
    <div class="login-lbl">Access Code</div>
    <input type="password" id="pwd" placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;" autocomplete="current-password" autofocus>
    <button class="login-btn" onclick="tryLogin()">Sign In &rarr;</button>
    <div class="login-err" id="login-err"></div>
    <div class="login-conf"><b>&#9670;</b> Confidential &middot; Authorised Access Only</div>
    <div class="login-copy">&copy; By Ravi Prakash. Proprietary Software. Unauthorized copying, distribution, or modification is prohibited | 2026.</div>
  </div>
</div>

<!-- APP -->
<div id="app">
  <div id="header">
    <button class="menu-btn" onclick="toggleSidebar()">&#9776;</button>
    <div class="hbrand">
      <div class="glyph-sm"></div>
      <div class="hname">__BRAND__</div>
      <span class="hchip" id="mode-chip">By Ravi</span>
    </div>
    <div class="hsearch">
      <span class="ic">&#8981;</span>
      <input type="search" id="search-box" placeholder="Search companies, CIN, products, notes&hellip;" oninput="onSearch()">
      <span class="kbd">&#8984;K</span>
    </div>
    <div class="hright">
      <div class="hts"><div class="k">DATA AS OF</div><div class="v" id="hts"></div></div>
      <button class="hlogout" onclick="logout()">Logout</button>
    </div>
  </div>

  <div id="sidebar">
    <div class="sid-sec">
      <div class="sid-lbl">NAVIGATION</div>
      <div class="sid-item active" data-nav="dashboard" onclick="showDash(this)"><span class="ico">&#9638;</span><span class="nm">Dashboard</span></div>
      <div class="sid-item" data-nav="all" onclick="filterView('all',this)"><span class="ico">&#9632;</span><span class="nm">All Companies</span><span class="sid-badge" id="sb-all"></span></div>
    </div>
    <div class="sid-sec">
      <div class="sid-lbl" id="portfolio-lbl">PORTFOLIO</div>
      <div id="portfolio-list"></div>
    </div>
    <div class="sid-sec">
      <div class="sid-lbl">BY SEGMENT</div>
      <div id="seg-list"></div>
    </div>
  </div>

  <div id="main">
    <div id="view-dash">
      <div class="dash-head">
        <div>
          <div class="eyebrow" id="dash-eyebrow">MARKET OVERVIEW</div>
          <div class="dash-title" id="dash-title">Market Research Dashboard</div>
          <div class="dash-sub" id="dash-sub"></div>
        </div>
        <div class="synced"><span class="live"></span>Dataset synced &middot; <span class="v" id="synced-ts"></span></div>
      </div>
      <div class="kpi-grid" id="kpi-grid"></div>
      <div class="charts">
        <div class="panel-card">
          <div class="pc-head"><div class="pc-title">Coverage by Segment</div><div class="pc-meta" id="seg-meta"></div></div>
          <div id="seg-bars"></div>
        </div>
        <div class="col-2">
          <div class="panel-card"><div class="pc-title" style="margin-bottom:16px">Priority / Competitor Tiers</div><div id="tier-bars"></div></div>
          <div class="panel-card"><div class="pc-title" style="margin-bottom:16px">Regional Spread</div><div id="reg-bars"></div></div>
        </div>
      </div>
      <div class="strip">
        <div class="pc-head"><div class="pc-title" id="strip-title">Portfolio Split</div><div class="pc-meta" id="strip-meta"></div></div>
        <div class="strip-row" id="strip-row"></div>
      </div>
    </div>

    <div id="view-list">
      <div class="toolbar">
        <div class="tb-title" id="view-label">All Companies</div>
        <select class="sel" id="f-region" onchange="applyFilters()"><option value="">All Regions</option></select>
        <select class="sel" id="f-fit" onchange="applyFilters()">
          <option value="">All Strategic Fit</option><option value="H">High Fit</option><option value="M">Medium Fit</option><option value="L">Low Fit</option>
        </select>
        <select class="sel" id="f-conf" onchange="applyFilters()">
          <option value="">All Confidence</option><option value="High">High Confidence</option><option value="Medium">Medium Confidence</option><option value="Low">Low Confidence</option>
        </select>
      </div>
      <div class="results">
        <div class="c" id="result-count"></div>
        <div class="n">Data generated from master database</div>
      </div>
      <div class="company-list" id="company-list"></div>
      <div class="pager" id="pager"></div>
    </div>
  </div>
</div>

<div id="overlay" onclick="closeDetail()"></div>
<div id="drawer">
  <div class="dp-head">
    <button class="dp-close" onclick="closeDetail()">&#x2715;</button>
    <div class="dp-eyebrow" id="dp-eyebrow">COMPANY DOSSIER</div>
    <div class="dp-name" id="dp-name"></div>
    <div class="dp-tags" id="dp-tags"></div>
  </div>
  <div class="dp-stats" id="dp-stats"></div>
  <div class="dp-body" id="dp-body"></div>
</div>

<script>
// ── DATA ──────────────────────────────────────────────────────────
const DATASETS = {mif: __DATA_MIF__, comp: __DATA_COMP__};
const COLUMNS = __COLS__;
const TIMESTAMP = "__TS__";
const PH_MIF = "__PH_MIF__";
const PH_COMP = "__PH_COMP__";

let MODE = 'mif';          // set at login
let DATA = [];             // active dataset
let state = {cat:'all', label:'All Companies', search:'', region:'', fit:'', conf:'', page:1, perPage:50, filtered:[]};

const MODE_CFG = {
  mif: {
    chip:'By Ravi', chipClass:'', eyebrow:'MARKET OVERVIEW',
    title:'Market Research Dashboard',
    sub:'Steel-consuming manufacturing landscape — active customers & conversion targets.',
    portfolioLabel:'PORTFOLIO',
    portfolio:[
      {cat:'Customer', label:'Active Customers', dot:'var(--green)'},
      {cat:'Prospect', label:'BD Prospects', dot:'var(--accent)'},
    ],
    stripTitle:'Portfolio Split',
  },
  comp: {
    chip:'COMPETITOR INTEL', chipClass:'comp', eyebrow:'COMPETITOR INTELLIGENCE',
    title:'Competitor Intelligence Dashboard',
    sub:'Competitors, machine makers, pipe & tube producers, rolls & dies and roll-forming rivals.',
    portfolioLabel:'COMPETITOR CLASSES',
    portfolio:[
      {seg:'Roll Forming Competitors', label:'Roll-Forming Rivals', dot:'var(--red)'},
      {seg:'Pipes & Tubes', label:'Pipe & Tube Makers', dot:'var(--steel)'},
      {seg:'Machine Makers', label:'Machine Manufacturers', dot:'var(--purple)'},
      {seg:'Rolls & Dies', label:'Rolls & Dies Makers', dot:'var(--amber)'},
    ],
    stripTitle:'Competitor Classes',
  }
};

// ── UTIL ──────────────────────────────────────────────────────────
function esc(s){return s==null?'':String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function fmt(n){return Number(n).toLocaleString('en-IN');}
function firstChar(v){return (v||'').trim().toUpperCase().charAt(0);}
function count(pred){return DATA.filter(pred).length;}
function segCount(s){return DATA.filter(c=>(c['Segment']||'')===s).length;}

function logoDomain(c){
  const w=(c['Company Website']||'').trim(); let dom='';
  if(w){const m=w.match(/([a-z0-9-]+(?:\.[a-z0-9-]+)+)/i); if(m) dom=m[1].replace(/^www\./i,'');}
  if(!dom){
    const emails=[c['Sales Contact'],c['Director Contact'],c['Proc. Contact']].filter(Boolean).join(' ');
    const em=emails.match(/[a-z0-9._%+-]+@([a-z0-9.-]+\.[a-z]{2,})/i);
    if(em){const d=em[1].toLowerCase(); if(!/gmail|yahoo|hotmail|outlook|rediff|live\.com|ymail|protonmail/.test(d)) dom=d;}
  }
  return dom;
}

// ── LOGIN ─────────────────────────────────────────────────────────
async function sha256(str){
  const buf=await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
async function tryLogin(){
  const pwd=document.getElementById('pwd').value;
  if(!pwd){setErr('Please enter your access code'); return;}
  const h=await sha256(pwd);
  if(h===PH_MIF)       { MODE='mif';  launch(); }
  else if(h===PH_COMP) { MODE='comp'; launch(); }
  else { setErr('Incorrect access code. Try again.'); document.getElementById('pwd').value=''; document.getElementById('pwd').focus(); }
}
function launch(){
  DATA = DATASETS[MODE];
  document.getElementById('login').style.display='none';
  document.getElementById('app').style.display='block';
  initApp();
}
function setErr(m){const e=document.getElementById('login-err'); e.textContent=m; setTimeout(()=>e.textContent='',3000);}
document.getElementById('pwd').addEventListener('keydown',e=>{if(e.key==='Enter')tryLogin();});
function logout(){if(confirm('Log out of MIF Market Research?'))location.reload();}

// ── INIT ──────────────────────────────────────────────────────────
function initApp(){
  const cfg=MODE_CFG[MODE];
  const chip=document.getElementById('mode-chip');
  chip.textContent=cfg.chip; chip.className='hchip '+cfg.chipClass;
  document.getElementById('hts').textContent=TIMESTAMP;
  document.getElementById('synced-ts').textContent=TIMESTAMP;
  document.getElementById('dash-eyebrow').textContent=cfg.eyebrow;
  document.getElementById('dash-title').textContent=cfg.title;
  document.getElementById('dash-sub').textContent=cfg.sub;
  document.getElementById('strip-title').textContent=cfg.stripTitle;
  document.getElementById('sb-all').textContent=fmt(DATA.length);

  // Regions dropdown
  const regs=[...new Set(DATA.map(c=>(c['Region']||'').trim()).filter(Boolean))].sort();
  const rf=document.getElementById('f-region');
  regs.forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;rf.appendChild(o);});

  // Portfolio nav
  document.getElementById('portfolio-lbl').textContent=cfg.portfolioLabel;
  const pl=document.getElementById('portfolio-list'); pl.innerHTML='';
  cfg.portfolio.forEach(item=>{
    const n = item.cat ? count(c=>(c._cat||'').indexOf(item.cat)>=0) : segCount(item.seg);
    const key = item.cat ? item.cat : ('seg:'+item.seg);
    const d=document.createElement('div');
    d.className='sid-item';
    d.innerHTML=`<span class="dot" style="background:${item.dot}"></span><span class="nm">${esc(item.label)}</span><span class="sid-badge">${fmt(n)}</span>`;
    d.onclick=()=>filterView(key,d);
    pl.appendChild(d);
  });

  // Segment nav
  const segCounts={};
  DATA.forEach(c=>{const s=(c['Segment']||'').trim(); if(s) segCounts[s]=(segCounts[s]||0)+1;});
  const sl=document.getElementById('seg-list'); sl.innerHTML='';
  Object.entries(segCounts).sort((a,b)=>b[1]-a[1]).forEach(([s,v])=>{
    const d=document.createElement('div');
    d.className='sid-item sid-seg';
    d.innerHTML=`<span class="nm">${esc(s)}</span><span class="sc">${v}</span>`;
    d.onclick=()=>filterView('seg:'+s,d);
    sl.appendChild(d);
  });

  buildDashboard();
}

// ── TIER NORMALISATION ────────────────────────────────────────────
function normTier(k){
  const s=(k||'').toLowerCase();
  if(!s.trim()) return 'Unclassified';
  if(/t1 captive/.test(s)) return 'T1 Captive';
  if(/t1 oem/.test(s)) return 'T1 OEM';
  if(/t1 active|t1 prospect|active prospect|highest priority|walk-in|priority/.test(s)) return 'T1 Active Prospect';
  if(/t1 psu|defence|shipyard/.test(s)) return 'T1 PSU / Defence';
  if(/^t1|tier 1|tier-1/.test(s)) return 'T1 Other';
  if(/t2 conversion/.test(s)) return 'T2 Conversion Possible';
  if(/^t2|tier 2/.test(s)) return 'T2 Other';
  if(/pure prospect/.test(s)) return 'Pure Prospect';
  if(/t3 small|^t3/.test(s)) return 'T3 Small';
  if(/competitor/.test(s)) return 'Competitor';
  if(/channel partner|franchise|fabricator|strategic partner|supplier/.test(s)) return 'Channel / Partner';
  if(/not applicable|aluminium|exited|exclude|watch|monitor/.test(s)) return 'Not Applicable';
  return 'Other';
}

// ── DASHBOARD ─────────────────────────────────────────────────────
const KPI_COLORS={accent:'var(--accent)',steel:'#4f8ff7',green:'#35d08a',red:'#ff6274',amber:'#ffc24d',purple:'#9d7bff'};
function buildDashboard(){
  const segCounts={};
  DATA.forEach(c=>{const s=(c['Segment']||'').trim(); if(s) segCounts[s]=(segCounts[s]||0)+1;});
  const nSeg=Object.keys(segCounts).length;
  const highConf=count(c=>c['Data Confidence']==='High');

  let kpis;
  if(MODE==='comp'){
    kpis=[
      {v:DATA.length, l:'Competitor Entities', s:`across ${nSeg} classes`, c:'red', go:()=>filterView('all',navEl('all'))},
      {v:segCount('Roll Forming Competitors'), l:'Roll-Forming Rivals', s:'direct competition', c:'accent', go:()=>filterView('seg:Roll Forming Competitors',null)},
      {v:segCount('Pipes & Tubes'), l:'Pipe &amp; Tube Makers', s:'tube / hollow section', c:'steel', go:()=>filterView('seg:Pipes & Tubes',null)},
      {v:segCount('Machine Makers'), l:'Machine Manufacturers', s:'equipment builders', c:'purple', go:()=>filterView('seg:Machine Makers',null)},
      {v:segCount('Rolls & Dies'), l:'Rolls &amp; Dies Makers', s:'tooling suppliers', c:'amber', go:()=>filterView('seg:Rolls & Dies',null)},
      {v:highConf, l:'High-Confidence', s:'verified records', c:'green', go:()=>filterView('all',navEl('all'))},
    ];
  } else {
    const cust=count(c=>(c._cat||'').indexOf('Customer')>=0);
    const prosp=count(c=>(c._cat||'').indexOf('Prospect')>=0);
    const hiFit=count(c=>firstChar(c['Strategic Fit H/M/L'])==='H');
    kpis=[
      {v:DATA.length, l:'Companies Tracked', s:`across ${nSeg} segments`, c:'accent', go:()=>filterView('all',navEl('all'))},
      {v:prosp, l:'BD Prospects', s:'active pipeline', c:'amber', go:()=>filterView('Prospect',null)},
      {v:cust, l:'Active Customers', s:'live accounts', c:'green', go:()=>filterView('Customer',null)},
      {v:hiFit, l:'High Strategic Fit', s:'top-priority targets', c:'red', go:()=>{document.getElementById('f-fit').value='H';filterView('all',navEl('all'));document.getElementById('f-fit').value='H';applyFilters();}},
      {v:highConf, l:'High-Confidence', s:'verified records', c:'steel', go:()=>filterView('all',navEl('all'))},
      {v:nSeg, l:'Market Segments', s:'tracked verticals', c:'purple', go:()=>showDash(navEl('dashboard'))},
    ];
  }
  const kg=document.getElementById('kpi-grid');
  kg.innerHTML=kpis.map((k,i)=>`
    <div class="kpi" data-i="${i}">
      <div class="halo" style="background:radial-gradient(circle,color-mix(in srgb,${KPI_COLORS[k.c]} 22%,transparent),transparent 70%)"></div>
      <div class="top"><span class="sq" style="background:${KPI_COLORS[k.c]}"></span><span class="lbl">${k.l}</span></div>
      <div class="num">0</div><div class="sub">${k.s}</div>
    </div>`).join('');
  [...kg.querySelectorAll('.kpi')].forEach((el,i)=>{el.onclick=kpis[i].go; animateCount(el.querySelector('.num'), kpis[i].v);});

  // Segment bars (top 12)
  const segE=Object.entries(segCounts).sort((a,b)=>b[1]-a[1]).slice(0,12);
  const segMax=segE[0]?segE[0][1]:1, segTop=segE.reduce((a,[,v])=>a+v,0);
  document.getElementById('seg-meta').textContent=`TOP ${segE.length} · ${fmt(segTop)} companies`;
  document.getElementById('seg-bars').innerHTML=segE.map(([k,v])=>`
    <div class="bar-row clk" onclick="filterView('seg:${esc(k).replace(/'/g,"\\'")}',null)">
      <div class="bar-lbl" title="${esc(k)}">${esc(k)}</div>
      <div class="bar-track"><div class="bar-fill" data-w="${Math.round(v/segMax*100)}" style="background:linear-gradient(90deg,var(--accent),#ffb066)"></div></div>
      <div class="bar-num">${v}</div></div>`).join('');

  // Tier bars (normalised)
  const tierAgg={};
  DATA.forEach(c=>{const n=normTier(c['Competitor Tier']||c['Priority Tier']); tierAgg[n]=(tierAgg[n]||0)+1;});
  delete tierAgg['Unclassified'];
  const tierE=Object.entries(tierAgg).sort((a,b)=>b[1]-a[1]).slice(0,7);
  const tierMax=tierE[0]?tierE[0][1]:1;
  document.getElementById('tier-bars').innerHTML=tierE.map(([k,v])=>`
    <div class="bar-row"><div class="bar-lbl sm" title="${esc(k)}">${esc(k)}</div>
      <div class="bar-track sm"><div class="bar-fill" data-w="${Math.round(v/tierMax*100)}" style="background:linear-gradient(90deg,#4f8ff7,#7cb0ff)"></div></div>
      <div class="bar-num">${fmt(v)}</div></div>`).join('') || '<div style="font-size:12px;color:#5f6e8c">No tier data</div>';

  // Region bars
  const regAgg={};
  DATA.forEach(c=>{const r=(c['Region']||'').trim().toUpperCase(); if(r) regAgg[r]=(regAgg[r]||0)+1;});
  const regE=Object.entries(regAgg).sort((a,b)=>b[1]-a[1]).slice(0,6);
  const regMax=regE[0]?regE[0][1]:1;
  document.getElementById('reg-bars').innerHTML=regE.map(([k,v])=>`
    <div class="bar-row"><div class="bar-lbl xs">${esc(k)}</div>
      <div class="bar-track sm"><div class="bar-fill" data-w="${Math.round(v/regMax*100)}" style="background:linear-gradient(90deg,#35d08a,#5fe3a8)"></div></div>
      <div class="bar-num">${fmt(v)}</div></div>`).join('') || '<div style="font-size:12px;color:#5f6e8c">No region data</div>';

  // Bottom strip
  let cells;
  if(MODE==='comp'){
    document.getElementById('strip-meta').textContent='by classification';
    cells=[
      {n:segCount('Roll Forming Competitors'),l:'Roll-forming rivals',c:'var(--red)'},
      {n:segCount('Pipes & Tubes'),l:'Pipe & tube makers',c:'var(--steel)'},
      {n:segCount('Machine Makers'),l:'Machine manufacturers',c:'var(--purple)'},
      {n:segCount('Rolls & Dies'),l:'Rolls & dies makers',c:'var(--amber)'},
    ];
  } else {
    document.getElementById('strip-meta').textContent='market composition';
    cells=[
      {n:DATA.length,l:'Total universe',c:'var(--steel)'},
      {n:count(c=>(c._cat||'').indexOf('Prospect')>=0),l:'BD prospects',c:'var(--accent)'},
      {n:count(c=>firstChar(c['Strategic Fit H/M/L'])==='H'),l:'High strategic fit',c:'var(--red)'},
      {n:count(c=>(c._cat||'').indexOf('Customer')>=0),l:'Active customers',c:'var(--green)'},
    ];
  }
  document.getElementById('strip-row').innerHTML=cells.map(p=>
    `<div class="strip-cell" style="border-left:3px solid ${p.c}"><div class="n">${fmt(p.n)}</div><div class="l">${p.l}</div></div>`).join('');

  requestAnimationFrame(()=>requestAnimationFrame(()=>{
    document.querySelectorAll('.bar-fill').forEach(b=>{b.style.width=b.dataset.w+'%';});
  }));
  setTimeout(()=>{document.querySelectorAll('.bar-fill').forEach(b=>{b.style.width=b.dataset.w+'%';});},120);
}
function animateCount(el,target){
  const dur=1200,t0=performance.now();
  function tick(now){const p=Math.min(1,(now-t0)/dur);const e=1-Math.pow(1-p,3);el.textContent=fmt(Math.round(target*e));if(p<1)requestAnimationFrame(tick);}
  requestAnimationFrame(tick);
  setTimeout(()=>{el.textContent=fmt(target);},1300);
}
function navEl(nav){return document.querySelector('.sid-item[data-nav="'+nav+'"]');}

// ── VIEW MGMT ─────────────────────────────────────────────────────
function setActive(el){document.querySelectorAll('.sid-item').forEach(i=>i.classList.remove('active')); if(el)el.classList.add('active');}
function showDash(el){document.getElementById('view-dash').style.display='block';document.getElementById('view-list').style.display='none';setActive(el||navEl('dashboard'));closeSidebar();closeDetail();}
function showList(el){document.getElementById('view-dash').style.display='none';document.getElementById('view-list').style.display='block';setActive(el);closeSidebar();closeDetail();}
function filterView(cat,el){
  state.cat=cat; state.page=1; state.search=''; document.getElementById('search-box').value='';
  state.region=''; document.getElementById('f-region').value='';
  state.fit=''; document.getElementById('f-fit').value='';
  state.conf=''; document.getElementById('f-conf').value='';
  let label='All Companies';
  if(cat==='Customer')label='Active Customers';
  else if(cat==='Prospect')label='BD Prospects';
  else if(cat.indexOf('seg:')===0)label=cat.slice(4);
  state.label=label;
  document.getElementById('view-label').textContent=label;
  if(!el){
    if(cat==='all')el=navEl('all');
    else el=[...document.querySelectorAll('.sid-item')].find(i=>i.querySelector('.nm')&&i.querySelector('.nm').textContent.trim()===label);
  }
  applyFilters(); showList(el);
}

// ── FILTER / SEARCH ───────────────────────────────────────────────
let searchTimer;
function onSearch(){clearTimeout(searchTimer);searchTimer=setTimeout(()=>{
  state.search=document.getElementById('search-box').value.toLowerCase(); state.page=1;
  if(document.getElementById('view-list').style.display==='none'){state.cat='all';state.label='All Companies';document.getElementById('view-label').textContent='All Companies';showList(navEl('all'));}
  renderList();
},200);}
function applyFilters(){
  state.region=document.getElementById('f-region').value;
  state.fit=document.getElementById('f-fit').value;
  state.conf=document.getElementById('f-conf').value;
  state.page=1; renderList();
}
const SEARCH_FIELDS=['Company Name','Segment','Products','BD Notes','Key Products for MIF','HQ Address','Plant Locations','Director Name','Procurement Head','Sales / BD Head','Primary Segment','CIN','GSTIN','Revenue Band','Customer OEMs Served','Parent / Group','Priority Tier','Competitor Tier','Entry Route','Current Steel Suppliers'];
function matches(c){
  const cat=state.cat;
  if(cat==='Customer'&&(c._cat||'').indexOf('Customer')<0)return false;
  if(cat==='Prospect'&&(c._cat||'').indexOf('Prospect')<0)return false;
  if(cat.indexOf('seg:')===0&&c['Segment']!==cat.slice(4))return false;
  if(state.region&&(c['Region']||'').toUpperCase().indexOf(state.region.toUpperCase())<0)return false;
  if(state.fit&&firstChar(c['Strategic Fit H/M/L'])!==state.fit)return false;
  if(state.conf&&c['Data Confidence']!==state.conf)return false;
  if(state.search){const q=state.search;return SEARCH_FIELDS.some(f=>c[f]&&String(c[f]).toLowerCase().indexOf(q)>=0);}
  return true;
}

// ── TAGS ──────────────────────────────────────────────────────────
function fitMeta(c){
  const f=firstChar(c['Strategic Fit H/M/L']);
  if(f==='H')return{bar:'#35d08a',text:'High Fit',bg:'#12281f',fg:'#4fe0a0'};
  if(f==='M')return{bar:'var(--accent)',text:'Medium Fit',bg:'#2a2113',fg:'#ffc24d'};
  if(f==='L')return{bar:'#4a566f',text:'Low Fit',bg:'#1a2133',fg:'#8493b0'};
  return{bar:'#4a566f',text:'',bg:'#1a2133',fg:'#8493b0'};
}
function confMeta(c){
  const v=c['Data Confidence']||'';
  if(v==='High')return{text:'High Confidence',bg:'#12281f',fg:'#4fe0a0'};
  if(v==='Medium')return{text:'Med Confidence',bg:'#2a2113',fg:'#ffc24d'};
  if(v==='Low')return{text:'Low Confidence',bg:'#2a1518',fg:'#ff8a97'};
  return null;
}
function tagHTML(text,bg,fg){return `<span class="tag" style="background:${bg};color:${fg}">${esc(text)}</span>`;}
function money(v){v=(v||'').trim();if(!v)return '';return /[a-z₹]/i.test(v)?v:'₹'+v+' Cr';}

// ── LIST ──────────────────────────────────────────────────────────
function renderList(){
  state.filtered=DATA.filter(matches);
  const total=state.filtered.length;
  const pages=Math.ceil(total/state.perPage);
  const start=(state.page-1)*state.perPage;
  const slice=state.filtered.slice(start,start+state.perPage);
  document.getElementById('result-count').innerHTML= total===0?'No companies found':
    `<b>${fmt(total)}</b> ${total===1?'company':'companies'} &middot; showing ${start+1}\u2013${Math.min(start+state.perPage,total)}`;
  const cl=document.getElementById('company-list');
  if(total===0){
    cl.innerHTML=`<div class="no-res"><div class="ic">&#8981;</div><div class="m">No companies match</div><div>Try adjusting your filters or search terms.</div></div>`;
    document.getElementById('pager').innerHTML=''; return;
  }
  cl.innerHTML=slice.map((c,i)=>{
    const idx=DATA.indexOf(c), fm=fitMeta(c), cm=confMeta(c);
    const rev=money(c['Revenue (Rs. Cr)'])||(c['Revenue Band']||'');
    const notes=c['BD Notes']||c['Key Products for MIF']||'';
    const seg=c['Segment']||c['Primary Segment']||'', region=c['Region']||'';
    let tags='';
    if(seg)tags+=tagHTML(seg,'#132133','#7fb0ff');
    if(region)tags+=tagHTML(region,'#0f2218','#4fe0a0');
    const tier=c['Priority Tier']||c['Competitor Tier']||'';
    if(tier)tags+=tagHTML(tier.length>34?tier.slice(0,32)+'\u2026':tier,'color-mix(in srgb,var(--accent) 12%,transparent)','var(--accent)');
    if(fm.text)tags+=tagHTML(fm.text,fm.bg,fm.fg);
    if(cm)tags+=tagHTML(cm.text,cm.bg,cm.fg);
    return `<div class="card" onclick="openDetail(${idx})">
      <div class="fitbar" style="background:${fm.bar}"></div>
      <div class="row">
        <div class="num">${String(start+i+1).padStart(2,'0')}</div>
        <div class="body">
          <div class="n1"><div class="cname">${esc(c['Company Name']||'\u2014')}</div>${rev?`<span class="rev">${esc(rev)}</span>`:''}</div>
          <div class="tags">${tags}</div>
          ${notes?`<div class="notes">${esc(notes.substring(0,180))}${notes.length>180?'\u2026':''}</div>`:''}
        </div>
        <div class="chev">&rsaquo;</div>
      </div>
    </div>`;
  }).join('');
  const pg=document.getElementById('pager');
  if(pages<=1){pg.innerHTML='';return;}
  const maxB=7,half=Math.floor(maxB/2);
  let s2=Math.max(1,state.page-half),e2=Math.min(pages,s2+maxB-1);
  if(e2-s2<maxB-1)s2=Math.max(1,e2-maxB+1);
  let h=state.page>1?`<button class="pg" onclick="goPage(${state.page-1})">&lsaquo;</button>`:'';
  for(let p=s2;p<=e2;p++)h+=`<button class="pg ${p===state.page?'active':''}" onclick="goPage(${p})">${p}</button>`;
  h+=state.page<pages?`<button class="pg" onclick="goPage(${state.page+1})">&rsaquo;</button>`:'';
  h+=`<span class="pg-info">Page ${state.page} of ${pages}</span>`;
  pg.innerHTML=h;
}
function goPage(p){state.page=p;renderList();window.scrollTo(0,0);}

// ── DETAIL DOSSIER ────────────────────────────────────────────────
function openDetail(idx){
  const c=DATA[idx], fm=fitMeta(c), cm=confMeta(c);
  document.getElementById('dp-eyebrow').textContent = MODE==='comp' ? 'COMPETITOR DOSSIER' : 'COMPANY DOSSIER';
  document.getElementById('dp-name').textContent=c['Company Name']||'\u2014';

  let tags='';
  if(c['Segment'])tags+=tagHTML(c['Segment'],'#132133','#7fb0ff');
  if(c['Region'])tags+=tagHTML(c['Region'],'#0f2218','#4fe0a0');
  const tier=c['Priority Tier']||c['Competitor Tier']||'';
  if(tier)tags+=tagHTML(tier.length>40?tier.slice(0,38)+'\u2026':tier,'color-mix(in srgb,var(--accent) 12%,transparent)','var(--accent)');
  if(fm.text)tags+=tagHTML(fm.text,fm.bg,fm.fg);
  if(cm)tags+=tagHTML(cm.text,cm.bg,cm.fg);
  document.getElementById('dp-tags').innerHTML=tags;

  const stats=[
    ['Revenue', money(c['Revenue (Rs. Cr)'])||'\u2014', 'var(--t1)'],
    ['Employees', c['Employees']||'\u2014', 'var(--t1)'],
    ['Plants', c['Plant Count']||'\u2014', 'var(--t1)'],
    ['Strategic Fit', firstChar(c['Strategic Fit H/M/L'])||'\u2014', fm.fg],
  ];
  document.getElementById('dp-stats').innerHTML=stats.map(([k,v,col])=>
    `<div class="dp-stat"><div class="k">${k}</div><div class="v" style="color:${col}">${esc(v)}</div></div>`).join('');

  // Logo mark from Company Website (old inline-onerror hide method)
  const dom=logoDomain(c);
  let logoHTML='';
  if(dom){
    logoHTML=`<div class="dp-logo">
      <div class="dp-logo-tile"><img src="https://icons.duckduckgo.com/ip3/${esc(dom)}.ico" alt="${esc(dom)}" referrerpolicy="no-referrer" onerror="this.closest('.dp-logo').style.display='none'"></div>
      <div><div class="k">COMPANY MARK</div><div class="nm">${esc(c['Company Name']||'')}</div><div class="dm">${esc(dom)}</div></div>
    </div>`;
  }

  const sections=[
    {t:'Company Identity',a:'#4f8ff7',f:[
      ['Legal Entity Name',c['Company Name']],['CIN',c['CIN']],['GSTIN',c['GSTIN']],
      ['Legal Form',c['Legal Form']],['Year Established',c['Year Est.']],['Parent / Group',c['Parent / Group']],
      ['Company Website',c['Company Website'],1],
    ]},
    {t:'Financial Profile',a:'#35d08a',f:[
      ['Revenue (Rs. Cr)',money(c['Revenue (Rs. Cr)'])],['Revenue Band',c['Revenue Band']],
      ['Revenue Source',c['Revenue Source']],['EBITDA or PAT (Rs. Cr)',money(c['EBITDA or PAT (Rs. Cr)'])],
      ['Growth / 3-yr CAGR',c['Revenue Growth YoY / 3-yr CAGR']],['Capex (Last 24m)',c['Capex Announced (Last 24 months)']],
    ]},
    {t:'Operations',a:'#9d7bff',f:[
      ['Employees',c['Employees']],['Employee Source',c['Employee Source']],['No. of Plants',c['Plant Count']],
      ['Plant Locations',c['Plant Locations'],1],['HQ Address',c['HQ Address'],1],
    ]},
    {t:'Key Contacts',a:'var(--accent)',f:[
      ['Director / Promoter',c['Director Name']],['Director Contact',c['Director Contact']],
      ['Procurement Head',c['Procurement Head']],['Proc. Contact',c['Proc. Contact']],
      ['Sales / BD Head',c['Sales / BD Head']],['Sales Contact',c['Sales Contact']],
    ]},
    {t:'Products &amp; Steel Buying',a:'#4f8ff7',f:[
      ['Primary Segment',c['Primary Segment']],['Products Made',c['Products'],1],
      ['Coil-to-Component?',c['Coil-to-Component']],['In-house Roll Forming?',c['In-house Roll Forming']],
      ['HSN Codes',c['HSN Codes']],['Steel Types',c['Steel Types']],['Total Volume Produced',c['Total Volume Produced']],
      ['Roll-Form Source',c['Roll-Form Source']],['Current Steel Suppliers',c['Current Steel Suppliers'],1],
      ['Annual Steel Tonnage (TPA)',c['Approximate Annual Steel Tonnage (TPA)']],['Buying Pattern',c['Buying Pattern']],
    ]},
    {t:'MIF Business Intelligence',a:'var(--accent)',f:[
      ['Strategic Fit (H/M/L)',c['Strategic Fit H/M/L'],0,'fit'],['Volume Potential (L/M/H)',c['Volume L/M/H']],
      ['Priority Tier',c['Priority Tier']],['Relationship Status',c['Relationship Status']],
      ['Customer OEMs Served',c['Customer OEMs Served'],1],['IATF Certified?',c['IATF Certified?']],
      ['MIF Proximity (km est.)',c['MIF Proximity (km est.)']],['Key Products for MIF',c['Key Products for MIF'],1],
      ['BD Notes',c['BD Notes'],1],['Entry Route',c['Entry Route'],1],['Action Owner',c['Action Owner']],
    ]},
    {t:'Competitor Analysis',a:'#ff6274',f:[
      ['Competitor Tier',c['Competitor Tier'],1],['Switching Difficulty',c['Switching Difficulty']],
    ]},
    {t:'Data Quality',a:'#5f6e8c',f:[
      ['Data Confidence',c['Data Confidence'],0,'conf'],['Last Verified Date',c['Last Verified Date']],
      ['Verification Method',c['Verification Method']],['Estimated vs Verified',c['Estimated vs Verified Flag'],1],
      ['Source Hyperlink(s)',c['Source Hyperlink(s)'],1],
    ]},
  ];

  let html=logoHTML;
  sections.forEach(sec=>{
    const rows=sec.f.filter(f=>f[1]&&String(f[1]).trim());
    if(!rows.length)return;
    html+=`<div class="dp-sec"><div class="dp-sec-head"><span class="dp-sec-bar" style="background:${sec.a}"></span><span class="dp-sec-title">${sec.t}</span><span class="dp-sec-line"></span></div><div class="dp-grid">`;
    rows.forEach(([k,v,full,type])=>{
      const cls=full?'dp-field full':'dp-field';
      let val;
      if(type==='fit'){const ch=firstChar(v);val=`<div class="dp-val hl-${ch}">${esc(v)}</div>`;}
      else if(type==='conf'){const ch=v==='High'?'H':v==='Medium'?'M':'L';val=`<div class="dp-val hl-${ch}">${esc(v)}</div>`;}
      else if(String(v).trim().indexOf('http')===0){const u=esc(String(v).trim());val=`<div class="dp-val"><a href="${u}" target="_blank" rel="noopener">${u.replace(/^https?:\/\/(www\.)?/,'')} \u2197</a></div>`;}
      else val=`<div class="dp-val">${esc(v)}</div>`;
      html+=`<div class="${cls}"><div class="dp-key">${k}</div>${val}</div>`;
    });
    html+=`</div></div>`;
  });
  document.getElementById('dp-body').innerHTML=html;
  document.getElementById('overlay').style.display='block';
  document.getElementById('drawer').classList.add('open');
}
function closeDetail(){document.getElementById('drawer').classList.remove('open');document.getElementById('overlay').style.display='none';}

// ── SIDEBAR (mobile) ──────────────────────────────────────────────
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDetail();});
</script>
</body>
</html>"""

# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    parse_args()
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found — {EXCEL_FILE}")
        print("Place this script in the same folder as the Excel file, or pass --input.")
        sys.exit(1)

    ts = datetime.now().strftime("%d %b %Y \u00b7 %I:%M %p")
    mif, comp, cols = read_companies()
    html = generate(mif, comp, cols, ts)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT_FILE) // 1024
    print(f"\n\u2713 Portal generated: {OUTPUT_FILE} ({size_kb} KB)")
    print(f"\u2713 Market login  ({PASSWORD_MIF}):  {len(mif)} companies")
    print(f"\u2713 Competitor login ({PASSWORD_COMP}): {len(comp)} companies")
    print(f"\u2713 Accent:    {ACCENT}")
    print(f"\u2713 Timestamp: {ts}")
    print(f"\nOpen '{OUTPUT_FILE}' in any browser. Fonts & company marks load from the web when online;")
    print("everything else works fully offline.")

if __name__ == '__main__':
    main()
