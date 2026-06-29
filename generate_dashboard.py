#!/usr/bin/env python3
"""
MIF Intelligence Portal Generator
===================================
Reads MIF_Master_Database Excel file → generates a single offline HTML portal.

Usage:
    python generate_dashboard.py
    python generate_dashboard.py --password YourPassword
    python generate_dashboard.py --input MyFile.xlsx --output MyPortal.html

Default password: MIF2026
"""

import pandas as pd
import json
import hashlib
import sys
import os
from datetime import datetime

# ─── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE  = "MIF_Master_Database_29June2026.xlsx"
OUTPUT_FILE = "MIF_Intelligence_Portal.html"
PASSWORD    = "MIF2026"
# ─────────────────────────────────────────────────────────────────────────────

def parse_args():
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == '--password' and i+1 < len(args):
            global PASSWORD; PASSWORD = args[i+1]; i+=2
        elif args[i] == '--input' and i+1 < len(args):
            global EXCEL_FILE; EXCEL_FILE = args[i+1]; i+=2
        elif args[i] == '--output' and i+1 < len(args):
            global OUTPUT_FILE; OUTPUT_FILE = args[i+1]; i+=2
        else:
            i+=1

def get_hash(pwd):
    return hashlib.sha256(pwd.encode('utf-8')).hexdigest()

def clean(v):
    if v is None: return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan','none','nat','<na>') else s

def read_companies():
    print(f"Reading: {EXCEL_FILE}")
    df = pd.read_excel(EXCEL_FILE, sheet_name='Combined Master', header=0)
    cols = list(df.columns)

    def has_company_name(x):
        if pd.isna(x): return False
        s = str(x).strip()
        return bool(s) and s.lower() not in ('nan','none','nat','<na>','▼  phase 1 — closed segments','▼  phase 2 — emerging segments')

    rows = df[df['Company Name'].apply(has_company_name)].copy()
    companies = []
    for _, row in rows.iterrows():
        c = {col: clean(row[col]) for col in cols}
        # Derive category for filtering
        rel  = c.get('Relationship Status','').lower()
        tier = c.get('Competitor Tier','').lower()
        cats = []
        if any(x in rel for x in ['existing customer','active oem','existing','customer (cnh)']):
            cats.append('Customer')
        if any(x in tier for x in ['t1 captive','t1 oem','competitor']):
            cats.append('Competitor')
        if any(x in tier for x in ['pure prospect','t2 conversion','t1 prospect','t1 active','channel partner']):
            cats.append('Prospect')
        if not cats: cats.append('Other')
        c['_cat'] = ','.join(cats)
        companies.append(c)
    print(f"Found {len(companies)} companies")
    return companies, cols

def generate(companies, cols, pw_hash, ts):
    segs = sorted({c['Segment'] for c in companies
                   if c.get('Segment') and '▼' not in c['Segment'] and 'PHASE' not in c['Segment'] and c['Segment'].strip()})
    regions = sorted({c['Region'] for c in companies if c.get('Region') and c['Region'].strip()})

    data_js  = json.dumps(companies, ensure_ascii=False, separators=(',',':'))
    segs_js  = json.dumps(segs, ensure_ascii=False)
    regs_js  = json.dumps(regions, ensure_ascii=False)
    cols_js  = json.dumps(cols, ensure_ascii=False)

    total  = len(companies)
    p1     = sum(1 for c in companies if 'Phase 1' in c.get('Phase / Source',''))
    p2     = sum(1 for c in companies if 'Phase 2' in c.get('Phase / Source',''))
    hi_pri = sum(1 for c in companies if c.get('Data Confidence','') == 'High')
    cust   = sum(1 for c in companies if 'Customer' in c.get('_cat',''))
    comp_c = sum(1 for c in companies if 'Competitor' in c.get('_cat',''))
    prosp  = sum(1 for c in companies if 'Prospect' in c.get('_cat',''))

    seg_counts = {}
    for c in companies:
        s = c.get('Segment','')
        if s and '▼' not in s and 'PHASE' not in s and s.strip():
            seg_counts[s] = seg_counts.get(s, 0) + 1

    tier_counts = {}
    for c in companies:
        t = c.get('Competitor Tier','').strip()
        if t: tier_counts[t] = tier_counts.get(t,0)+1

    reg_counts = {}
    for c in companies:
        r = c.get('Region','').strip().upper()
        if r: reg_counts[r] = reg_counts.get(r,0)+1
    # merge case variants
    merged_reg = {}
    for k,v in reg_counts.items():
        merged_reg[k] = merged_reg.get(k,0)+v

    seg_counts_js  = json.dumps(seg_counts, ensure_ascii=False)
    tier_counts_js = json.dumps(tier_counts, ensure_ascii=False)
    reg_counts_js  = json.dumps(merged_reg, ensure_ascii=False)

    return HTML_TEMPLATE.format(
        DATA=data_js, SEGS=segs_js, REGS=regs_js, COLS=cols_js,
        TS=ts, PH=pw_hash,
        TOTAL=total, P1=p1, P2=p2, HI=hi_pri,
        CUST=cust, COMP=comp_c, PROSP=prosp,
        SEG_COUNTS=seg_counts_js, TIER_COUNTS=tier_counts_js, REG_COUNTS=reg_counts_js
    )


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>MIF Intelligence Portal</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --blue:#1a3c6e;--blue2:#1e4a8a;--amber:#e67e22;--amber2:#f39c12;
  --bg:#f0f2f5;--card:#ffffff;--border:#d1d5db;--text:#1a1a2e;
  --text2:#4b5563;--text3:#9ca3af;--green:#16a34a;--red:#dc2626;
  --orange:#d97706;--sid:240px;--head:56px;
}}
body{{font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:var(--text);background:var(--bg);overflow-x:hidden}}
/* ── LOGIN ─────────────────────────────── */
#login-screen{{
  position:fixed;inset:0;background:var(--blue);display:flex;
  align-items:center;justify-content:center;z-index:9999;
  background:linear-gradient(135deg,#0d1f3c 0%,#1a3c6e 50%,#0d2e52 100%);
}}
.login-box{{
  background:#fff;border-radius:12px;padding:40px 48px;width:100%;max-width:420px;
  box-shadow:0 25px 60px rgba(0,0,0,.5);text-align:center;
}}
.login-logo{{color:var(--blue);font-size:22px;font-weight:800;letter-spacing:.5px;margin-bottom:4px}}
.login-sub{{color:var(--text3);font-size:12px;margin-bottom:28px;letter-spacing:.3px}}
.login-box input{{
  width:100%;border:2px solid var(--border);border-radius:8px;
  padding:12px 16px;font-size:15px;outline:none;margin-bottom:14px;
  transition:border .2s;
}}
.login-box input:focus{{border-color:var(--blue)}}
.login-btn{{
  width:100%;background:var(--blue);color:#fff;border:none;border-radius:8px;
  padding:13px;font-size:15px;font-weight:700;cursor:pointer;transition:background .2s;
  letter-spacing:.3px;
}}
.login-btn:hover{{background:var(--blue2)}}
.login-err{{color:var(--red);font-size:13px;margin-top:10px;min-height:20px}}
.login-conf{{color:var(--text3);font-size:11px;margin-top:18px}}
/* ── APP SHELL ─────────────────────────── */
#app{{display:none;min-height:100vh}}
/* HEADER */
#header{{
  position:fixed;top:0;left:0;right:0;height:var(--head);
  background:var(--blue);color:#fff;display:flex;align-items:center;
  padding:0 20px;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.3);gap:16px;
}}
.hdr-logo{{font-weight:800;font-size:16px;letter-spacing:.3px;white-space:nowrap}}
.hdr-sub{{font-size:12px;opacity:.7;display:none}}
.hdr-ts{{font-size:11px;opacity:.6;margin-left:auto;white-space:nowrap}}
.hdr-logout{{
  background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
  color:#fff;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:12px;
  white-space:nowrap;transition:background .2s;
}}
.hdr-logout:hover{{background:rgba(255,255,255,.22)}}
.menu-btn{{
  display:none;background:none;border:none;color:#fff;font-size:20px;
  cursor:pointer;padding:4px;
}}
/* SIDEBAR */
#sidebar{{
  position:fixed;top:var(--head);left:0;width:var(--sid);
  height:calc(100vh - var(--head));overflow-y:auto;
  background:#12294d;z-index:90;transition:transform .25s;
}}
#sidebar::-webkit-scrollbar{{width:4px}}
#sidebar::-webkit-scrollbar-thumb{{background:#2a4a7a}}
.sid-section{{padding:10px 0 4px 0;}}
.sid-label{{
  font-size:10px;font-weight:700;color:rgba(255,255,255,.35);
  padding:0 16px 4px;letter-spacing:1.2px;text-transform:uppercase;
}}
.sid-item{{
  display:flex;align-items:center;gap:8px;
  padding:9px 16px;color:rgba(255,255,255,.75);cursor:pointer;
  font-size:13px;transition:all .15s;border-left:3px solid transparent;
}}
.sid-item:hover{{background:rgba(255,255,255,.06);color:#fff}}
.sid-item.active{{
  background:rgba(230,126,34,.15);color:var(--amber2);
  border-left-color:var(--amber);font-weight:600;
}}
.sid-badge{{
  margin-left:auto;background:rgba(255,255,255,.12);
  border-radius:10px;padding:2px 7px;font-size:11px;
}}
.sid-item.active .sid-badge{{background:rgba(230,126,34,.25);color:var(--amber2)}}
/* MAIN CONTENT */
#main{{
  margin-left:var(--sid);padding-top:var(--head);min-height:100vh;
}}
/* DASHBOARD */
#view-dashboard{{padding:24px}}
.dash-title{{font-size:20px;font-weight:700;color:var(--blue);margin-bottom:20px}}
.stat-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:14px;margin-bottom:28px}}
.stat-card{{
  background:var(--card);border-radius:10px;padding:18px 20px;
  box-shadow:0 1px 4px rgba(0,0,0,.08);border-top:3px solid var(--blue);
}}
.stat-num{{font-size:32px;font-weight:800;color:var(--blue);line-height:1}}
.stat-lbl{{font-size:12px;color:var(--text2);margin-top:4px}}
.stat-card.amber{{border-top-color:var(--amber)}}
.stat-card.amber .stat-num{{color:var(--amber)}}
.stat-card.green{{border-top-color:var(--green)}}
.stat-card.green .stat-num{{color:var(--green)}}
.stat-card.red{{border-top-color:var(--red)}}
.stat-card.red .stat-num{{color:var(--red)}}
.chart-row{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}}
.chart-card{{background:var(--card);border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
.chart-title{{font-size:13px;font-weight:700;color:var(--blue);margin-bottom:14px;border-bottom:1px solid var(--border);padding-bottom:8px}}
.bar-row{{display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:12px}}
.bar-label{{width:180px;min-width:120px;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:11px}}
.bar-track{{flex:1;background:#e5e7eb;border-radius:4px;height:14px;overflow:hidden}}
.bar-fill{{height:100%;border-radius:4px;background:var(--blue);transition:width .3s}}
.bar-num{{width:40px;text-align:right;font-weight:600;color:var(--text);font-size:11px}}
/* LIST VIEW */
#view-list{{padding:0}}
.list-toolbar{{
  background:var(--card);border-bottom:1px solid var(--border);
  padding:14px 20px;position:sticky;top:var(--head);z-index:80;
  display:flex;flex-wrap:wrap;gap:10px;align-items:center;
}}
.search-wrap{{position:relative;flex:1;min-width:200px}}
.search-wrap input{{
  width:100%;border:1.5px solid var(--border);border-radius:8px;
  padding:9px 14px 9px 36px;font-size:13px;outline:none;transition:border .2s;
}}
.search-wrap input:focus{{border-color:var(--blue)}}
.search-icon{{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--text3);font-size:14px}}
.filter-select{{
  border:1.5px solid var(--border);border-radius:8px;padding:9px 12px;
  font-size:12px;color:var(--text2);outline:none;background:#fff;cursor:pointer;
}}
.filter-select:focus{{border-color:var(--blue)}}
.results-bar{{
  padding:10px 20px;font-size:12px;color:var(--text3);background:var(--bg);
  display:flex;align-items:center;justify-content:space-between;
}}
.result-count{{font-weight:600;color:var(--text2)}}
/* Company Cards */
.company-list{{padding:12px 20px;display:flex;flex-direction:column;gap:8px}}
.company-card{{
  background:var(--card);border-radius:10px;padding:14px 18px;
  cursor:pointer;transition:all .15s;border:1.5px solid transparent;
  box-shadow:0 1px 3px rgba(0,0,0,.06);
}}
.company-card:hover{{border-color:var(--blue);box-shadow:0 3px 12px rgba(26,60,110,.12);transform:translateY(-1px)}}
.cc-row1{{display:flex;align-items:flex-start;gap:10px;margin-bottom:6px}}
.cc-name{{font-weight:700;font-size:14px;color:var(--blue);flex:1;line-height:1.3}}
.cc-num{{font-size:11px;color:var(--text3);min-width:24px;padding-top:2px}}
.cc-row2{{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px}}
.cc-row3{{display:flex;flex-wrap:wrap;gap:6px;align-items:center}}
.tag{{
  display:inline-block;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600;
}}
.tag-seg{{background:#dbeafe;color:#1d4ed8}}
.tag-region{{background:#f0fdf4;color:#15803d}}
.tag-tier{{background:#fef3c7;color:#92400e}}
.tag-phase{{background:#ede9fe;color:#6d28d9}}
.tag-conf-H{{background:#dcfce7;color:#166534}}
.tag-conf-M{{background:#fef9c3;color:#854d0e}}
.tag-conf-L{{background:#fee2e2;color:#991b1b}}
.tag-fit-H{{background:#0ea5e9;color:#fff}}
.tag-fit-M{{background:#f59e0b;color:#fff}}
.tag-fit-L{{background:#94a3b8;color:#fff}}
.cc-rev{{font-size:12px;color:var(--text2);margin-left:auto}}
.cc-notes{{font-size:11px;color:var(--text3);line-height:1.4;margin-top:4px;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
/* DETAIL PANEL */
#detail-overlay{{
  display:none;position:fixed;inset:0;background:rgba(0,0,0,.3);z-index:200;
}}
#detail-panel{{
  position:fixed;top:var(--head);right:-100%;bottom:0;
  width:clamp(320px,62%,900px);background:#fff;z-index:210;
  overflow-y:auto;transition:right .25s ease;box-shadow:-5px 0 30px rgba(0,0,0,.15);
}}
#detail-panel.open{{right:0}}
#detail-panel::-webkit-scrollbar{{width:5px}}
#detail-panel::-webkit-scrollbar-thumb{{background:var(--border)}}
.dp-header{{
  background:var(--blue);color:#fff;padding:20px 24px;
  position:sticky;top:0;z-index:5;
}}
.dp-close{{
  position:absolute;top:16px;right:20px;background:rgba(255,255,255,.15);
  border:none;color:#fff;width:30px;height:30px;border-radius:50%;
  cursor:pointer;font-size:18px;line-height:30px;text-align:center;
  transition:background .2s;
}}
.dp-close:hover{{background:rgba(255,255,255,.3)}}
.dp-name{{font-size:18px;font-weight:800;margin-bottom:6px;padding-right:40px;line-height:1.3}}
.dp-tags{{display:flex;flex-wrap:wrap;gap:5px;margin-top:8px}}
.dp-body{{padding:20px 24px}}
.dp-section{{margin-bottom:20px}}
.dp-sec-title{{
  font-size:11px;font-weight:800;color:var(--text3);text-transform:uppercase;
  letter-spacing:1px;padding-bottom:8px;border-bottom:1.5px solid var(--border);
  margin-bottom:12px;
}}
.dp-grid{{display:grid;grid-template-columns:1fr 1fr;gap:0}}
.dp-field{{padding:7px 0;border-bottom:1px solid #f3f4f6}}
.dp-field:last-child,.dp-field:nth-last-child(2){{border-bottom:none}}
.dp-key{{font-size:11px;color:var(--text3);margin-bottom:2px;font-weight:600}}
.dp-val{{font-size:13px;color:var(--text);line-height:1.4;word-break:break-word}}
.dp-val.em{{color:var(--text3);font-style:italic}}
.dp-full{{grid-column:1/-1}}
.highlight-H{{color:var(--green);font-weight:700}}
.highlight-M{{color:var(--orange);font-weight:700}}
.highlight-L{{color:var(--red);font-weight:600}}
/* PAGINATION */
.pagination{{
  display:flex;justify-content:center;align-items:center;
  gap:8px;padding:20px;flex-wrap:wrap;
}}
.pg-btn{{
  padding:7px 13px;border:1.5px solid var(--border);border-radius:6px;
  background:#fff;cursor:pointer;font-size:12px;color:var(--text2);transition:.15s;
}}
.pg-btn:hover{{border-color:var(--blue);color:var(--blue)}}
.pg-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.pg-info{{font-size:12px;color:var(--text3)}}
/* NO RESULTS */
.no-results{{
  text-align:center;padding:60px 20px;color:var(--text3);
}}
.no-results-icon{{font-size:48px;margin-bottom:12px}}
.no-results-msg{{font-size:15px;font-weight:600;color:var(--text2);margin-bottom:6px}}
/* RESPONSIVE */
@media(max-width:768px){{
  :root{{--sid:0px}}
  #sidebar{{transform:translateX(-100%);width:260px}}
  #sidebar.open{{transform:none;box-shadow:4px 0 20px rgba(0,0,0,.3)}}
  .menu-btn{{display:block}}
  #main{{margin-left:0}}
  .hdr-sub{{display:none}}
  .hdr-ts{{display:none}}
  .chart-row{{grid-template-columns:1fr}}
  .dp-grid{{grid-template-columns:1fr}}
  #detail-panel{{width:100%;left:-100%;right:auto;top:var(--head)}}
  #detail-panel.open{{left:0}}
  .bar-label{{width:110px}}
  .stat-grid{{grid-template-columns:repeat(2,1fr)}}
  .dp-full{{grid-column:1}}
}}
@media(max-width:480px){{
  .login-box{{padding:28px 24px;margin:16px}}
  .stat-grid{{grid-template-columns:1fr 1fr}}
  .list-toolbar{{padding:10px 12px}}
  .company-list{{padding:8px 12px}}
  .dp-body{{padding:16px}}
  #view-dashboard{{padding:16px}}
}}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="login-screen">
  <div class="login-box">
    <div class="login-logo">&#9679; MIF INTELLIGENCE</div>
    <div class="login-sub">Mother India Forming · Market Intelligence Portal</div>
    <input type="password" id="pwd-input" placeholder="Enter access password" autocomplete="current-password">
    <button class="login-btn" onclick="tryLogin()">Access Portal</button>
    <div class="login-err" id="login-err"></div>
    <div class="login-conf">&#128274; Confidential · Authorised Access Only</div>
     <div class="login-conf login-copyright"> &copy; By Ravi Prakash. Proprietary Software. Unauthorized copying, distribution, or modification is prohibited | 2026. </div>
  </div>
</div>

<!-- APP -->
<div id="app">
  <!-- Header -->
  <div id="header">
    <button class="menu-btn" onclick="toggleSidebar()">&#9776;</button>
    <div class="hdr-logo">&#9679; MIF INTELLIGENCE PORTAL</div>
    <div class="hdr-sub">Market Intelligence · Confidential</div>
    <div class="hdr-ts" id="hdr-ts"></div>
    <button class="hdr-logout" onclick="logout()">Logout</button>
  </div>

  <!-- Sidebar -->
  <div id="sidebar">
    <div class="sid-section">
      <div class="sid-label">Navigation</div>
      <div class="sid-item active" onclick="showView('dashboard',this)" data-view="dashboard">
        &#9632; Dashboard
      </div>
      <div class="sid-item" onclick="filterView('all',this)" data-view="list" data-cat="all">
        &#9632; All Companies <span class="sid-badge" id="sb-all"></span>
      </div>
    </div>
    <div class="sid-section">
      <div class="sid-label">By Category</div>
      <div class="sid-item" onclick="filterView('Customer',this)" data-view="list" data-cat="Customer">
        &#9632; Customers <span class="sid-badge" id="sb-cust"></span>
      </div>
      <div class="sid-item" onclick="filterView('Competitor',this)" data-view="list" data-cat="Competitor">
        &#9632; Competitors <span class="sid-badge" id="sb-comp"></span>
      </div>
      <div class="sid-item" onclick="filterView('Prospect',this)" data-view="list" data-cat="Prospect">
        &#9632; BD Prospects <span class="sid-badge" id="sb-prosp"></span>
      </div>
    </div>
    <div class="sid-section">
      <div class="sid-label">By Phase</div>
      <div class="sid-item" onclick="filterView('phase1',this)" data-view="list" data-cat="phase1">
        &#9632; Phase 1 — Closed Segs <span class="sid-badge" id="sb-p1"></span>
      </div>
      <div class="sid-item" onclick="filterView('phase2',this)" data-view="list" data-cat="phase2">
        &#9632; Phase 2 — Emerging <span class="sid-badge" id="sb-p2"></span>
      </div>
    </div>
    <div class="sid-section">
      <div class="sid-label">By Segment</div>
      <div id="seg-list"></div>
    </div>
  </div>

  <!-- Main -->
  <div id="main">
    <!-- Dashboard View -->
    <div id="view-dashboard">
      <div class="dash-title">Market Intelligence Dashboard</div>
      <div class="stat-grid" id="stat-grid"></div>
      <div class="chart-row" id="chart-row"></div>
    </div>

    <!-- List View -->
    <div id="view-list" style="display:none">
      <div class="list-toolbar">
        <div class="search-wrap">
          <span class="search-icon">&#128269;</span>
          <input type="search" id="search-box" placeholder="Search by name, CIN, products, BD notes..." oninput="onSearch()">
        </div>
        <select class="filter-select" id="f-region" onchange="applyFilters()">
          <option value="">All Regions</option>
        </select>
        <select class="filter-select" id="f-tier" onchange="applyFilters()">
          <option value="">All Tiers</option>
          <option value="T1 Captive">T1 Captive</option>
          <option value="T1 OEM">T1 OEM</option>
          <option value="T2 Conversion Possible">T2 Conversion</option>
          <option value="Pure Prospect">Pure Prospect</option>
          <option value="T3 Small">T3 Small</option>
        </select>
        <select class="filter-select" id="f-fit" onchange="applyFilters()">
          <option value="">All Fit</option>
          <option value="H">High Fit</option>
          <option value="M">Medium Fit</option>
          <option value="L">Low Fit</option>
        </select>
        <select class="filter-select" id="f-conf" onchange="applyFilters()">
          <option value="">All Confidence</option>
          <option value="High">High Confidence</option>
          <option value="Medium">Medium Confidence</option>
          <option value="Low">Low Confidence</option>
        </select>
      </div>
      <div class="results-bar">
        <span class="result-count" id="result-count"></span>
        <span id="view-label" style="color:var(--blue);font-weight:600;font-size:12px"></span>
      </div>
      <div class="company-list" id="company-list"></div>
      <div class="pagination" id="pagination"></div>
    </div>
  </div>
</div>

<!-- Detail Overlay -->
<div id="detail-overlay" onclick="closeDetail()"></div>
<div id="detail-panel">
  <div class="dp-header">
    <button class="dp-close" onclick="closeDetail()">&#x2715;</button>
    <div class="dp-name" id="dp-name"></div>
    <div class="dp-tags" id="dp-tags"></div>
  </div>
  <div class="dp-body" id="dp-body"></div>
</div>

<script>
// ── DATA ──────────────────────────────────────────────────────────
const DATA = {DATA};
const SEGMENTS = {SEGS};
const REGIONS = {REGS};
const COLUMNS = {COLS};
const TIMESTAMP = "{TS}";
const PASS_HASH = "{PH}";
const SEG_COUNTS = {SEG_COUNTS};
const TIER_COUNTS = {TIER_COUNTS};
const REG_COUNTS = {REG_COUNTS};
const TOTALS = {{total:{TOTAL},p1:{P1},p2:{P2},hi:{HI},cust:{CUST},comp:{COMP},prosp:{PROSP}}};

// ── STATE ─────────────────────────────────────────────────────────
let state = {{
  cat: 'all', search: '', region: '', tier: '', fit: '', conf: '',
  page: 1, perPage: 50, filtered: []
}};

// ── LOGIN ─────────────────────────────────────────────────────────
async function sha256(str) {{
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}}
async function tryLogin() {{
  const pwd = document.getElementById('pwd-input').value;
  if (!pwd) {{ setErr('Please enter password'); return; }}
  const h = await sha256(pwd);
  if (h === PASS_HASH) {{
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app').style.display = 'block';
    initApp();
  }} else {{
    setErr('Incorrect password. Please try again.');
    document.getElementById('pwd-input').value = '';
    document.getElementById('pwd-input').focus();
  }}
}}
function setErr(msg) {{
  document.getElementById('login-err').textContent = msg;
  setTimeout(()=>document.getElementById('login-err').textContent='', 3000);
}}
document.getElementById('pwd-input').addEventListener('keydown', e => {{ if(e.key==='Enter') tryLogin(); }});

function logout() {{
  if(confirm('Log out of MIF Intelligence Portal?')) location.reload();
}}

// ── INIT ──────────────────────────────────────────────────────────
function initApp() {{
  document.getElementById('hdr-ts').textContent = 'Data as of: '+TIMESTAMP;
  document.getElementById('sb-all').textContent = TOTALS.total;
  document.getElementById('sb-cust').textContent = TOTALS.cust;
  document.getElementById('sb-comp').textContent = TOTALS.comp;
  document.getElementById('sb-prosp').textContent = TOTALS.prosp;
  document.getElementById('sb-p1').textContent = TOTALS.p1;
  document.getElementById('sb-p2').textContent = TOTALS.p2;

  // Region filter options
  const rf = document.getElementById('f-region');
  REGIONS.forEach(r => {{
    if(r&&r.trim()) {{ const o=document.createElement('option'); o.value=r; o.textContent=r; rf.appendChild(o); }}
  }});

  // Segment sidebar
  const sl = document.getElementById('seg-list');
  SEGMENTS.forEach(s => {{
    const d = document.createElement('div');
    d.className = 'sid-item';
    const cnt = SEG_COUNTS[s]||0;
    d.innerHTML = `&#9632; ${{s}} <span class="sid-badge">${{cnt}}</span>`;
    d.onclick = () => filterView('seg:'+s, d);
    d.dataset.view = 'list';
    sl.appendChild(d);
  }});

  buildDashboard();
}}

// ── DASHBOARD ────────────────────────────────────────────────────
function buildDashboard() {{
  const sg = document.getElementById('stat-grid');
  const stats = [
    {{num:TOTALS.total, lbl:'Total Companies', cls:''}},
    {{num:TOTALS.p1,    lbl:'Phase 1 — Closed Segments', cls:''}},
    {{num:TOTALS.p2,    lbl:'Phase 2 — Emerging Segments', cls:'amber'}},
    {{num:TOTALS.cust,  lbl:'Active Customers', cls:'green'}},
    {{num:TOTALS.comp,  lbl:'Identified Competitors', cls:'red'}},
    {{num:TOTALS.prosp, lbl:'BD Prospects', cls:'amber'}},
    {{num:TOTALS.hi,    lbl:'High Data Confidence', cls:'green'}},
    {{num:SEGMENTS.length, lbl:'Market Segments Covered', cls:''}},
  ];
  sg.innerHTML = stats.map(s=>`
    <div class="stat-card ${{s.cls}}" onclick="filterView('all',null);showView('list',null);" style="cursor:pointer">
      <div class="stat-num">${{s.num}}</div>
      <div class="stat-lbl">${{s.lbl}}</div>
    </div>`).join('');

  const cr = document.getElementById('chart-row');

  // Segment chart
  const segEntries = Object.entries(SEG_COUNTS).sort((a,b)=>b[1]-a[1]).slice(0,18);
  const segMax = segEntries[0]?.[1]||1;
  cr.innerHTML += `<div class="chart-card">
    <div class="chart-title">Companies by Segment</div>
    ${{segEntries.map(([k,v])=>`
      <div class="bar-row">
        <div class="bar-label" title="${{k}}">${{k.replace('P1 — ','').replace('P2 — ','')}}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${{Math.round(v/segMax*100)}}%"></div></div>
        <div class="bar-num">${{v}}</div>
      </div>`).join('')}}
  </div>`;

  // Competitor tier + Region
  const tierOrder = ['T1 Captive','T1 OEM','T1 Active Prospect','T2 Conversion Possible','Pure Prospect','T3 Small'];
  const tierEntries = Object.entries(TIER_COUNTS).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const tierMax = tierEntries[0]?.[1]||1;

  const regEntries = Object.entries(REG_COUNTS).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const regMax = regEntries[0]?.[1]||1;

  cr.innerHTML += `<div class="chart-card">
    <div class="chart-title">Competitor Tier Breakdown</div>
    ${{tierEntries.map(([k,v])=>`
      <div class="bar-row">
        <div class="bar-label" title="${{k}}">${{k}}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${{Math.round(v/tierMax*100)}}%;background:var(--amber)"></div></div>
        <div class="bar-num">${{v}}</div>
      </div>`).join('')}}
    <div class="chart-title" style="margin-top:20px">Companies by Region</div>
    ${{regEntries.map(([k,v])=>`
      <div class="bar-row">
        <div class="bar-label">${{k}}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${{Math.round(v/regMax*100)}}%;background:var(--green)"></div></div>
        <div class="bar-num">${{v}}</div>
      </div>`).join('')}}
  </div>`;
}}

// ── VIEW MANAGEMENT ──────────────────────────────────────────────
function showView(name, el) {{
  document.getElementById('view-dashboard').style.display = name==='dashboard' ? 'block' : 'none';
  document.getElementById('view-list').style.display     = name==='list'      ? 'block' : 'none';
  document.querySelectorAll('.sid-item').forEach(i=>i.classList.remove('active'));
  if(el) el.classList.add('active');
  closeSidebar();
  closeDetail();
}}

function filterView(cat, el) {{
  state.cat = cat; state.page = 1;
  state.search = ''; document.getElementById('search-box').value='';
  state.region=''; document.getElementById('f-region').value='';
  state.tier='';   document.getElementById('f-tier').value='';
  state.fit='';    document.getElementById('f-fit').value='';
  state.conf='';   document.getElementById('f-conf').value='';

  let label = 'All Companies';
  if(cat==='Customer') label='Customers';
  else if(cat==='Competitor') label='Competitors';
  else if(cat==='Prospect') label='BD Prospects';
  else if(cat==='phase1') label='Phase 1 Companies';
  else if(cat==='phase2') label='Phase 2 Companies';
  else if(cat.startsWith('seg:')) label=cat.replace('seg:','');
  document.getElementById('view-label').textContent = label;

  applyFilters();
  showView('list', el);
}}

// ── FILTERING & SEARCH ───────────────────────────────────────────
let searchTimer;
function onSearch() {{
  clearTimeout(searchTimer);
  searchTimer = setTimeout(()=>{{ state.search=document.getElementById('search-box').value.toLowerCase(); state.page=1; renderList(); }}, 200);
}}

function applyFilters() {{
  state.region = document.getElementById('f-region').value;
  state.tier   = document.getElementById('f-tier').value;
  state.fit    = document.getElementById('f-fit').value;
  state.conf   = document.getElementById('f-conf').value;
  state.page   = 1;
  renderList();
}}

const SEARCH_FIELDS = ['Company Name','Segment','Products','BD Notes','Key Products for MIF',
  'HQ Address','Plant Locations','Director Name','Procurement Head','Sales / BD Head',
  'Primary Segment','CIN','GSTIN','Revenue Band','Customer OEMs Served','Parent / Group',
  'Priority Tier','Competitor Tier','Entry Route','Current Steel Suppliers'];

function matchesFilters(c) {{
  const cat = state.cat;
  // Category filter
  if(cat==='Customer'  && !c._cat.includes('Customer'))   return false;
  if(cat==='Competitor'&& !c._cat.includes('Competitor')) return false;
  if(cat==='Prospect'  && !c._cat.includes('Prospect'))   return false;
  if(cat==='phase1' && !c['Phase / Source'].includes('Phase 1')) return false;
  if(cat==='phase2' && !c['Phase / Source'].includes('Phase 2')) return false;
  if(cat.startsWith('seg:') && c['Segment'] !== cat.replace('seg:','')) return false;
  // Dropdown filters
  if(state.region && !c['Region'].toUpperCase().includes(state.region.toUpperCase())) return false;
  if(state.tier   && !c['Competitor Tier'].includes(state.tier))   return false;
  if(state.fit    && c['Strategic Fit H/M/L'].toUpperCase().trim().charAt(0) !== state.fit) return false;
  if(state.conf   && c['Data Confidence'] !== state.conf) return false;
  // Text search
  if(state.search) {{
    const q = state.search;
    return SEARCH_FIELDS.some(f => c[f] && c[f].toLowerCase().includes(q));
  }}
  return true;
}}

function renderList() {{
  state.filtered = DATA.filter(matchesFilters);
  const total = state.filtered.length;
  const pages = Math.ceil(total/state.perPage);
  const start = (state.page-1)*state.perPage;
  const slice = state.filtered.slice(start, start+state.perPage);

  document.getElementById('result-count').textContent =
    total===0 ? 'No companies found' :
    `Showing ${{start+1}}–${{Math.min(start+state.perPage,total)}} of ${{total}} companies`;

  const cl = document.getElementById('company-list');
  if(total===0) {{
    cl.innerHTML=`<div class="no-results">
      <div class="no-results-icon">&#128269;</div>
      <div class="no-results-msg">No companies found</div>
      <div>Try adjusting your filters or search terms</div>
    </div>`;
    document.getElementById('pagination').innerHTML='';
    return;
  }}

  cl.innerHTML = slice.map((c,i)=>{{
    const idx = DATA.indexOf(c);
    const fit = (c['Strategic Fit H/M/L']||'').trim().toUpperCase().charAt(0);
    const fitLabel = fit==='H'?'High Fit':fit==='M'?'Medium Fit':fit==='L'?'Low Fit':'';
    const conf = (c['Data Confidence']||'').trim().charAt(0);
    const rev = c['Revenue Band']||c['Revenue (Rs. Cr)']||'';
    const notes = c['BD Notes']||c['Key Products for MIF']||'';
    const phase = c['Phase / Source'].includes('Phase 1')?'P1':'P2';
    const seg = c['Segment']||c['Primary Segment']||'';
    const region = c['Region']||'';
    return `<div class="company-card" onclick="openDetail(${{idx}})">
      <div class="cc-row1">
        <div class="cc-num">${{start+i+1}}</div>
        <div class="cc-name">${{esc(c['Company Name']||'—')}}</div>
        ${{rev?`<div class="cc-rev">${{esc(rev)}}</div>`:''}}
      </div>
      <div class="cc-row2">
        ${{seg?`<span class="tag tag-seg">${{esc(seg)}}</span>`:''}}
        ${{region?`<span class="tag tag-region">${{esc(region)}}</span>`:''}}
        ${{c['Competitor Tier']?`<span class="tag tag-tier">${{esc(c['Competitor Tier'])}}</span>`:''}}
        <span class="tag tag-phase">${{phase}}</span>
        ${{fitLabel?`<span class="tag tag-fit-${{fit}}">${{fitLabel}}</span>`:''}}
        ${{conf?`<span class="tag tag-conf-${{conf}}">${{conf==='H'?'High Confidence':conf==='M'?'Medium Confidence':'Low Confidence'}}</span>`:''}}
      </div>
      ${{notes?`<div class="cc-notes">${{esc(notes.substring(0,160))}}${{notes.length>160?'…':''}}</div>`:''}}
    </div>`;
  }}).join('');

  // Pagination
  const pg = document.getElementById('pagination');
  if(pages<=1){{ pg.innerHTML=''; return; }}
  const maxBtns=7, half=Math.floor(maxBtns/2);
  let start2=Math.max(1,state.page-half), end2=Math.min(pages,start2+maxBtns-1);
  if(end2-start2<maxBtns-1) start2=Math.max(1,end2-maxBtns+1);
  let phtml = state.page>1?`<button class="pg-btn" onclick="goPage(${{state.page-1}})">&#8249; Prev</button>`:'';
  for(let p=start2;p<=end2;p++) phtml+=`<button class="pg-btn ${{p===state.page?'active':''}}" onclick="goPage(${{p}})">${{p}}</button>`;
  phtml += state.page<pages?`<button class="pg-btn" onclick="goPage(${{state.page+1}})">Next &#8250;</button>`:'';
  phtml += `<span class="pg-info">Page ${{state.page}} of ${{pages}}</span>`;
  pg.innerHTML=phtml;
}}

function goPage(p) {{ state.page=p; renderList(); document.getElementById('main').scrollTop=0; window.scrollTo(0,56); }}

// ── DETAIL PANEL ─────────────────────────────────────────────────
function openDetail(idx) {{
  const c = DATA[idx];
  const panel = document.getElementById('detail-panel');

  document.getElementById('dp-name').textContent = c['Company Name']||'—';

  // Tags
  const tags = [];
  if(c['Segment']) tags.push(`<span class="tag tag-seg">${{esc(c['Segment'])}}</span>`);
  if(c['Region']) tags.push(`<span class="tag tag-region">${{esc(c['Region'])}}</span>`);
  if(c['Competitor Tier']) tags.push(`<span class="tag tag-tier">${{esc(c['Competitor Tier'])}}</span>`);
  if(c['Phase / Source']) tags.push(`<span class="tag tag-phase">${{esc(c['Phase / Source'])}}</span>`);
  if(c['Data Confidence']) tags.push(`<span class="tag tag-conf-${{(c['Data Confidence']||'').charAt(0)}}">${{esc(c['Data Confidence'])}} Confidence</span>`);
  document.getElementById('dp-tags').innerHTML = tags.join('');

  // Body sections
  const sections = [
    {{ title:'Company Identity', fields:[
      ['Legal Entity Name', c['Company Name']],
      ['CIN', c['CIN']],
      ['GSTIN', c['GSTIN']],
      ['Legal Form', c['Legal Form']],
      ['Year Established', c['Year Est.']],
      ['Parent / Group', c['Parent / Group']],
    ]}},
    {{ title:'Financials', fields:[
      ['Revenue (Rs. Cr)', c['Revenue (Rs. Cr)']],
      ['Revenue Band', c['Revenue Band']],
      ['Revenue Source', c['Revenue Source']],
      ['EBITDA or PAT (Rs. Cr)', c['EBITDA or PAT (Rs. Cr)']],
      ['Revenue Growth / 3-yr CAGR', c['Revenue Growth YoY / 3-yr CAGR']],
      ['Capex Announced (Last 24m)', c['Capex Announced (Last 24 months)']],
    ]}},
    {{ title:'Operations', fields:[
      ['Employees', c['Employees']],
      ['Employee Source', c['Employee Source']],
      ['No. of Plants', c['Plant Count']],
      ['Plant Locations', c['Plant Locations'], true],
      ['HQ Address', c['HQ Address'], true],
    ]}},
    {{ title:'Key Contacts', fields:[
      ['Director / Promoter', c['Director Name']],
      ['Director Contact', c['Director Contact']],
      ['Procurement Head', c['Procurement Head']],
      ['Proc. Contact', c['Proc. Contact']],
      ['Sales / BD Head', c['Sales / BD Head']],
      ['Sales Contact', c['Sales Contact']],
    ]}},
    {{ title:'Products & Steel Buying', fields:[
      ['Primary Segment', c['Primary Segment']],
      ['Products Made', c['Products'], true],
      ['Coil-to-Component?', c['Coil-to-Component']],
      ['In-house Roll Forming?', c['In-house Roll Forming']],
      ['HSN Codes', c['HSN Codes']],
      ['Steel Types Used', c['Steel Types']],
      ['Total Volume Produced', c['Total Volume Produced']],
      ['Roll-Form Source', c['Roll-Form Source']],
      ['Current Steel Suppliers', c['Current Steel Suppliers'], true],
      ['Annual Steel Tonnage (TPA)', c['Approximate Annual Steel Tonnage (TPA)']],
      ['Buying Pattern', c['Buying Pattern']],
    ]}},
    {{ title:'MIF Business Development Intelligence', fields:[
      ['Strategic Fit (H/M/L)', c['Strategic Fit H/M/L'], false, true],
      ['Volume Potential (L/M/H)', c['Volume L/M/H']],
      ['Priority Tier', c['Priority Tier']],
      ['Relationship Status', c['Relationship Status']],
      ['Customer OEMs Served', c['Customer OEMs Served'], true],
      ['IATF Certified?', c['IATF Certified?']],
      ['MIF Proximity (km est.)', c['MIF Proximity (km est.)']],
      ['Key Products for MIF', c['Key Products for MIF'], true],
      ['BD Notes', c['BD Notes'], true],
      ['Entry Route', c['Entry Route'], true],
      # ['Action Owner', c['Action Owner']],
    ]}},
    {{ title:'Competitor Analysis', fields:[
      ['Competitor Tier', c['Competitor Tier']],
      ['Switching Difficulty', c['Switching Difficulty']],
      ['Vendor Reg. Status (MIF)', c['Vendor Registration Status with MIF']],
    ]}},
    {{ title:'Data Quality', fields:[
      ['Data Confidence', c['Data Confidence'], false, true],
      ['Last Verified Date', c['Last Verified Date']],
      ['Verification Method', c['Verification Method']],
      ['Estimated vs Verified', c['Estimated vs Verified Flag'], true],
      ['Source Hyperlink(s)', c['Source Hyperlink(s)'], true],
    ]}},
  ];

  let html = '';
  sections.forEach(sec => {{
    const rows = sec.fields.filter(f => f[1] && f[1].trim());
    if(!rows.length) return;
    html += `<div class="dp-section"><div class="dp-sec-title">${{sec.title}}</div><div class="dp-grid">`;
    rows.forEach(([k,v,full,highlight]) => {{
      const cls = full?'dp-full':'';
      let valHtml = '';
      if(highlight) {{
        const ch = (v||'').trim().toUpperCase().charAt(0);
        valHtml = `<div class="dp-val highlight-${{ch==='H'?'H':ch==='M'?'M':ch==='L'?'L':''}}">${{esc(v)}}</div>`;
      }} else {{
        valHtml = `<div class="dp-val">${{esc(v||'—')}}</div>`;
      }}
      html += `<div class="dp-field ${{cls}}"><div class="dp-key">${{k}}</div>${{valHtml}}</div>`;
    }});
    html += `</div></div>`;
  }});

  document.getElementById('dp-body').innerHTML = html;
  document.getElementById('detail-overlay').style.display = 'block';
  panel.classList.add('open');
}}

function closeDetail() {{
  document.getElementById('detail-panel').classList.remove('open');
  document.getElementById('detail-overlay').style.display = 'none';
}}

// ── SIDEBAR TOGGLE (mobile) ──────────────────────────────────────
function toggleSidebar() {{
  document.getElementById('sidebar').classList.toggle('open');
}}
function closeSidebar() {{
  document.getElementById('sidebar').classList.remove('open');
}}

// ── UTILS ─────────────────────────────────────────────────────────
function esc(s) {{
  if(!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

// ── KEYBOARD ──────────────────────────────────────────────────────
document.addEventListener('keydown', e => {{
  if(e.key==='Escape') closeDetail();
}});
</script>
<div style="margin-left:var(--sid);padding:14px 20px;text-align:center;font-size:11px;color:var(--text3);border-top:1px solid var(--border);background:var(--card);">
 &copy; 2026 Ravi Prakash &bull; Proprietary Software. Unauthorized copying, distribution, or modification is prohibited &bull; Internal Use Only &bull; 2026  Version 1.0-beta &bull; Build 20260629.
</div>
</body>
</html>"""

# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    parse_args()
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found — {EXCEL_FILE}")
        print("Please place this script in the same folder as the Excel file.")
        sys.exit(1)

    pw_hash = get_hash(PASSWORD)
    ts = datetime.now().strftime("%d %b %Y, %I:%M %p")

    companies, cols = read_companies()
    html = generate(companies, cols, pw_hash, ts)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT_FILE) // 1024
    print(f"\n✓ Portal generated: {OUTPUT_FILE} ({size_kb} KB)")
    print(f"✓ Companies: {len(companies)}")
    print(f"✓ Password:  {PASSWORD}")
    print(f"✓ Timestamp: {ts}")
    print(f"\nOpen '{OUTPUT_FILE}' in any browser — no internet needed.")

if __name__ == '__main__':
    main()
