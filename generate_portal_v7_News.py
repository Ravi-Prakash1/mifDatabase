#!/usr/bin/env python3
"""
MIF MARKET RESEARCH — Portal Generator (v5, dual-login) — LIGHT THEME
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
import base64
from datetime import datetime

# ─── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE     = "MIF_Prospects_Master_Database.xlsx"
OUTPUT_FILE    = f"MIF_Market_Research_Portal_{datetime.now().strftime('%d%B%Y')}.html"
LOGO_MARK_FILE = "mif_mark.png"       # clean M-in-oval mark -> brandmark + favicon
NEWS_FILE      = "MIF_News_Intelligence.xlsx"   # rolling daily/weekly briefing feed
EXHIBITIONS_FILE = "MIF_Global_Exhibitions_FY_2026_27.xlsx"
PASSWORD_MIF   = "MIF2026"       # -> customers + prospects (competitors hidden)
PASSWORD_COMP  = "Confidential"  # -> competitors / machine / pipe&tube / rolls&dies only
ACCENT         = "#e07a2e"       # molten copper (deepened for light theme). Try "#2f6fe0", "#12a05f", "#b08417"
BRAND          = "MIF Market Research"
BYLINE         = "By Ravi Prakash"

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
    global PASSWORD_MIF, PASSWORD_COMP, EXCEL_FILE, OUTPUT_FILE, ACCENT, BRAND, EXHIBITIONS_FILE, NEWS_FILE
    args = sys.argv[1:]; i = 0
    while i < len(args):
        a = args[i]
        if a == '--mif-password' and i+1 < len(args):   PASSWORD_MIF = args[i+1]; i += 2
        elif a == '--comp-password' and i+1 < len(args): PASSWORD_COMP = args[i+1]; i += 2
        elif a == '--input' and i+1 < len(args):         EXCEL_FILE = args[i+1]; i += 2
        elif a == '--output' and i+1 < len(args):        OUTPUT_FILE = args[i+1]; i += 2
        elif a == '--accent' and i+1 < len(args):        ACCENT = args[i+1]; i += 2
        elif a == '--brand' and i+1 < len(args):         BRAND = args[i+1]; i += 2
        elif a == '--exhibitions' and i+1 < len(args):   EXHIBITIONS_FILE = args[i+1]; i += 2
        elif a == '--news' and i+1 < len(args):          NEWS_FILE = args[i+1]; i += 2
        else: i += 1

def sha256(pwd):
    return hashlib.sha256(pwd.encode('utf-8')).hexdigest()

def img_data_uri(path):
    """Read an image file and return a base64 data URI so the portal stays
    fully self-contained/offline (no external logo requests)."""
    try:
        with open(path, 'rb') as fh:
            data = fh.read()
        ext = os.path.splitext(path)[1].lower().lstrip('.')
        mime = {'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg',
                'svg':'image/svg+xml','gif':'image/gif','webp':'image/webp'}.get(ext, 'image/png')
        return 'data:' + mime + ';base64,' + base64.b64encode(data).decode('ascii')
    except Exception as e:
        print('Logo load failed (' + str(path) + '):', e)
        return ''

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


# ─── EXHIBITIONS READER ───────────────────────────────────────────────────────
_FLAG_RE  = re.compile('[\U0001F1E6-\U0001F1FF]')
_MONTHS   = {'Jan':0,'Feb':1,'Mar':2,'Apr':3,'May':4,'Jun':5,'Jul':6,
             'Aug':7,'Sep':8,'Oct':9,'Nov':10,'Dec':11}
_MONTH_RE = re.compile(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b')

def _parse_months(s):
    s = str(s or '')
    ym = re.search(r'\b(20\d{2})\b', s)
    year = int(ym.group(1)) if ym else None
    out, seen = [], set()
    for mm in _MONTH_RE.findall(s):
        if mm in seen:
            continue
        seen.add(mm)
        out.append({'m': _MONTHS[mm], 'y': year})
    return out

def _split_priority(raw):
    raw   = str(raw or '')
    lines = raw.split('\n')
    first = lines[0].strip()
    note  = ' '.join(l.strip() for l in lines[1:]).strip()
    stars = first.count('\u2b50')
    if re.search(r'MUST ATTEND|PLAN BOTH TRIPS', first): flag = 'MUST_ATTEND'
    elif re.search(r'AVOID', first):                     flag = 'AVOID'
    elif re.search(r'ATTEND', first):                    flag = 'ATTEND'
    elif re.search(r'MONITOR', first):                   flag = 'WATCH'
    else:                                                flag = 'WATCH'
    return flag, stars, note

def _split_loc(raw):
    s = _FLAG_RE.sub('', str(raw or '')).strip()
    parts = s.split('\u2014')  # em dash
    country = parts[0].split('\n')[0].strip()
    city = ''
    if len(parts) > 1:
        city = parts[1].split('\n')[0].strip()
    else:
        nl = parts[0].split('\n')
        if len(nl) > 1:
            city = nl[1].strip()
    country = country.split('/')[0].split('+')[0].strip()
    return country, city

def read_exhibitions():
    if not os.path.exists(EXHIBITIONS_FILE):
        print(f"WARNING: Exhibitions file not found — {EXHIBITIONS_FILE}. Portal built with no exhibitions.")
        return []
    print(f"Reading exhibitions: {EXHIBITIONS_FILE}")
    xl = pd.ExcelFile(EXHIBITIONS_FILE)
    sheet = None
    for s in xl.sheet_names:
        u = s.upper()
        if 'MASTER INDEX' in u and 'COMPLETE' in u:
            sheet = s; break
    if sheet is None:
        for s in xl.sheet_names:
            if 'MASTER INDEX' in s.upper():
                sheet = s; break
    if sheet is None:
        sheet = xl.sheet_names[0]
    print(f"Exhibitions sheet: {sheet}")
    df = pd.read_excel(EXHIBITIONS_FILE, sheet_name=sheet, header=None)

    hdr = None
    for i in range(min(8, len(df))):
        vals = [str(x) for x in df.iloc[i].tolist()]
        if any('Exhibition / Conference' in v for v in vals):
            hdr = i; break
    if hdr is None:
        hdr = 2

    exh = []
    for i in range(hdr + 1, len(df)):
        row = df.iloc[i].tolist()
        def cell(idx):
            if idx >= len(row):
                return ''
            v = row[idx]
            if v is None:
                return ''
            try:
                if isinstance(v, float) and pd.isna(v):
                    return ''
            except Exception:
                pass
            return str(v).strip()
        c0        = cell(0)
        name_full = cell(1)
        if c0.startswith('----'):
            continue
        if not name_full:
            continue
        name_lines = name_full.split('\n')
        name       = name_lines[0].strip()
        name_note  = ' '.join(l.strip() for l in name_lines[1:]).strip()
        country, city = _split_loc(cell(3))
        flag, stars, pnote = _split_priority(cell(19))
        exh.append({
            'name': name, 'nameNote': name_note, 'segment': cell(2),
            'country': country, 'city': city, 'locationRaw': cell(3),
            'dateRaw': cell(4), 'months': _parse_months(cell(4)),
            'frequency': cell(5), 'nextEdition': cell(6), 'type': cell(7),
            'cost': cell(8), 'leadTime': cell(9), 'scale': cell(10),
            'productFit': cell(11), 'keyCustomers': cell(12), 'competitors': cell(13),
            'website': cell(14), 'organiser': cell(15), 'venue': cell(16),
            'visitorReq': cell(17), 'visa': cell(18),
            'priority': flag, 'stars': stars, 'priorityNote': pnote,
            'source': cell(20), 'status': cell(21),
        })
    print(f"Exhibitions loaded: {len(exh)}")
    return exh


def _news_yn(v):
    return str(v).strip().lower() in ('y', 'yes', 'true', '1', 'x', '\u2713')

def _news_move(v):
    s = str(v).strip().lower()
    if s.startswith('up') or 'ris' in s or 'firm' in s or 'gain' in s or '\u2191' in s: return 'up'
    if s.startswith('down') or 'soft' in s or 'fall' in s or 'fell' in s or 'weak' in s or '\u2193' in s: return 'down'
    return 'flat'

def _news_status(v):
    s = str(v).strip().lower()
    if 'red' in s or 'critical' in s or s == 'act' or 'risk' in s: return 'red'
    if 'green' in s or 'positive' in s or 'no action' in s or 'opportun' in s: return 'green'
    if 'yellow' in s or 'watch' in s or 'amber' in s: return 'yellow'
    return 'neutral'

def parse_news(path):
    """Read the rolling News Intelligence Excel (sheets: Briefing, Indicators,
    Trend, News) into the structure the portal renders. Missing file -> empty
    feed (portal simply hides the ticker / news card)."""
    news = {'latest': '', 'updated': '', 'trend': {}, 'editions': []}
    if not path or not os.path.exists(path):
        print(f"News file not found ({path}) - portal will render without a news feed.")
        return news
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print("Could not open news file:", e); return news

    def sheet(name):
        for nm in xl.sheet_names:
            if nm.strip().lower() == name.lower():
                return pd.read_excel(xl, sheet_name=nm).fillna('')
        return pd.DataFrame()

    def dkey(d):
        try: return pd.to_datetime(d)
        except Exception: return pd.Timestamp.min

    def ndate(d):
        try: return pd.to_datetime(d).strftime('%Y-%m-%d')
        except Exception: return str(d).strip()

    def ldate(d):
        try: return pd.to_datetime(d).strftime('%d %b %Y')
        except Exception: return str(d).strip()

    def col(row, *names):
        for n in names:
            for k in row.index:
                if str(k).strip().lower() == n.lower():
                    return clean(row[k])
        return ''

    bdf, idf, tdf, ndf = sheet('Briefing'), sheet('Indicators'), sheet('Trend'), sheet('News')
    editions, order = {}, []

    def ed(dk, d):
        if dk not in editions:
            editions[dk] = {'date': dk, 'dateLabel': ldate(d), 'type': 'Daily',
                            'indicators': [], 'note': '', 'stories': []}
            order.append(dk)
        return editions[dk]

    for _, r in bdf.iterrows():
        d = col(r, 'Date')
        if not d: continue
        e = ed(ndate(d), d)
        t = col(r, 'Type');  note = col(r, 'Input Cost Note', 'Note', 'InputCostNote')
        if t: e['type'] = t
        if note: e['note'] = note

    for _, r in idf.iterrows():
        d, name = col(r, 'Date'), col(r, 'Indicator', 'Name')
        if not d or not name: continue
        ed(ndate(d), d)['indicators'].append({
            'name': name, 'value': col(r, 'Value'), 'unit': col(r, 'Unit'),
            'move': _news_move(col(r, 'Movement', 'Move')),
            'driver': col(r, 'Driver', 'Key Driver'), 'source': col(r, 'Source')})

    for _, r in ndf.iterrows():
        d, hl = col(r, 'Date'), col(r, 'Headline')
        if not d or not hl: continue
        ed(ndate(d), d)['stories'].append({
            'section': col(r, 'Section') or 'Update',
            'segment': col(r, 'Segment') or '\u2014',
            'status': _news_status(col(r, 'Status')),
            'company': col(r, 'Company') or '\u2014',
            'headline': hl,
            'what': col(r, 'What Happened', 'What', 'Summary'),
            'why': col(r, 'Why It Matters', 'Why', 'Why It Matters to MIF'),
            'source': col(r, 'Source'),
            'importance': (col(r, 'Importance') or 'normal').lower(),
            'marquee': _news_yn(col(r, 'Marquee')),
            'dashboard': _news_yn(col(r, 'Dashboard'))})

    order_sorted = sorted(set(order), key=dkey, reverse=True)
    news['editions'] = [editions[d] for d in order_sorted]
    if news['editions']:
        news['latest'] = news['editions'][0]['date']
        news['updated'] = news['editions'][0]['dateLabel']

    def _tlabel(d):
        try: return pd.to_datetime(d).strftime('%d %b')
        except Exception: return str(d).strip()

    def _series(colnames, title, meta, color, kfmt):
        pts = []
        for _, r in tdf.iterrows():
            d = col(r, 'Date'); v = ''
            for cn in colnames:
                v = col(r, cn)
                if v != '': break
            if not d or v == '': continue
            try: fv = float(str(v).replace('\u20b9', '').replace(',', '').strip())
            except Exception: continue
            pts.append((d, fv))
        pts.sort(key=lambda x: dkey(x[0]))
        if len(pts) < 2: return None
        def disp(fv): return ('%.1fk' % (fv / 1000)) if kfmt else ('%.2f' % fv)
        return {'title': title, 'meta': meta, 'color': color,
                'points': [{'label': _tlabel(d), 'v': fv, 'disp': disp(fv)} for d, fv in pts]}

    trends = []
    st_ = _series(['HR Coil', 'HR Coil (Steel)', 'Steel'], 'HR Coil (Steel) \u2014 Trend', '\u20b9 per ton', '#e07a2e', True)
    if st_: trends.append(st_)
    inr_ = _series(['USD/INR', 'USDINR', 'USD INR'], 'USD/INR \u2014 Trend', '\u20b9 per USD', '#2f6fe0', False)
    if inr_: trends.append(inr_)
    news['trends'] = trends
    return news


def generate(mif, comp, cols, ts, exh, news):
    repl = {
        '__DATA_MIF__':  json.dumps(mif,  ensure_ascii=False, separators=(',', ':')),
        '__DATA_COMP__': json.dumps(comp, ensure_ascii=False, separators=(',', ':')),
        '__COLS__':      json.dumps(cols, ensure_ascii=False),
        '__DATA_EXH__':  json.dumps(exh, ensure_ascii=False, separators=(',', ':')),
        '__DATA_NEWS__': json.dumps(news, ensure_ascii=False, separators=(',', ':')),
        '__TS__':        ts,
        '__PH_MIF__':    sha256(PASSWORD_MIF),
        '__PH_COMP__':   sha256(PASSWORD_COMP),
        '__ACCENT__':    ACCENT,
        '__BRAND__':     BRAND,
        '__BYLINE__':    BYLINE,
        '__LOGO_MARK__': img_data_uri(LOGO_MARK_FILE),
    }
    html = HTML_TEMPLATE
    for k, v in repl.items():
        html = html.replace(k, v)
    return html


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<link rel="icon" type="image/png" href="__LOGO_MARK__">
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__BRAND__ · __BYLINE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --accent:__ACCENT__;
  --bg:#f5f7fa; --panel:#ffffff; --card:#ffffff; --card2:#f7f9fc;
  --border:#e6eaf0; --border2:#e0e5ec;
  --t1:#1b2333; --t2:#5b6678; --t3:#6b7688; --t4:#96a0b0;
  --steel:#2f6fe0; --green:#12a05f; --red:#e0455e; --amber:#c58a10; --purple:#7a5cf0;
  --sid:230px; --head:60px; --tick:0px;
  --sans:'IBM Plex Sans',system-ui,'Segoe UI',sans-serif;
  --disp:'Space Grotesk','IBM Plex Sans',system-ui,sans-serif;
  --mono:'IBM Plex Mono',ui-monospace,Consolas,monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--sans);background:var(--bg);color:var(--t1);font-size:14px;overflow-x:hidden}
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#cdd5df;border-radius:8px}
::-webkit-scrollbar-thumb:hover{background:#b4bfcc}
input::placeholder{color:#98a2b2}
button{font-family:inherit}
a{color:var(--steel)}
@keyframes fade{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
@keyframes glow{0%,100%{opacity:.5}50%{opacity:.9}}
@keyframes shimmer{0%{background-position:-260px 0}100%{background-position:260px 0}}

.brandmark{position:relative;flex:none}
.brandmark img{display:block;object-fit:contain;border-radius:8px;background:#fff}
.glyph{position:relative}
.glyph:before{content:"";position:absolute;inset:0;border-radius:22%;
  background:linear-gradient(135deg,var(--accent),#c85a1e);transform:rotate(45deg)}
.glyph:after{content:"";position:absolute;inset:29%;border-radius:12%;background:#fff;transform:rotate(45deg)}

/* WELCOME */
#welcome{position:fixed;inset:0;z-index:10000;display:flex;flex-direction:column;align-items:center;justify-content:center;
  text-align:center;padding:40px;overflow:hidden;
  background:radial-gradient(1200px 720px at 50% -8%,#ffffff 0%,#eef2f7 52%,#e4eaf2 100%)}
.wl-glow{position:absolute;width:640px;height:640px;border-radius:50%;top:-180px;
  background:radial-gradient(circle,color-mix(in srgb,var(--accent) 16%,transparent) 0%,transparent 70%);
  filter:blur(30px);animation:glow 7s ease-in-out infinite;pointer-events:none}
.wl-inner{position:relative;animation:fade .8s ease both;max-width:760px}
.wl-mark{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;margin-bottom:24px}
.wl-mark img{width:80px;height:80px;flex:none;object-fit:contain;background:#fff;border-radius:18px;padding:7px;border:1px solid #e6eaf0;box-shadow:0 12px 32px rgba(20,30,50,.12)}
.wl-eyebrow{font-family:var(--disp);font-size:12.5px;font-weight:700;letter-spacing:3px;color:#7a8598;text-transform:uppercase}
.wl-title{font-family:var(--disp);font-size:44px;font-weight:600;letter-spacing:-1px;color:var(--t1);line-height:1.08}
.wl-title .brand{color:var(--accent)}
.wl-sub{font-family:var(--disp);font-size:21px;font-weight:500;color:var(--t2);margin-top:14px;letter-spacing:-.2px}
.wl-desc{font-size:14.5px;color:#6b7688;margin:20px auto 0;max-width:520px;line-height:1.6}
.wl-cta{margin-top:38px;display:inline-flex;align-items:center;gap:12px;cursor:pointer;
  background:linear-gradient(135deg,var(--accent),#e06a24);color:#fff;border:none;
  border-radius:13px;padding:17px 34px;font-family:var(--disp);font-size:16px;font-weight:600;letter-spacing:.2px;
  box-shadow:0 16px 40px color-mix(in srgb,var(--accent) 34%,transparent);transition:transform .16s,box-shadow .2s}
.wl-cta:hover{transform:translateY(-2px);box-shadow:0 22px 52px color-mix(in srgb,var(--accent) 42%,transparent)}
.wl-cta .arw{font-size:19px;transition:transform .2s}
.wl-cta:hover .arw{transform:translateX(4px)}
.wl-foot{position:fixed;bottom:0;left:0;width:100%;padding:14px 20px;font-size:10px;color:#a7b0be;letter-spacing:.3px;text-align:center;line-height:1.6;background:linear-gradient(to top,#e4eaf2,rgba(228,234,242,0));z-index:1}

/* LOGIN */
#login{position:fixed;inset:0;z-index:9999;display:none;align-items:center;justify-content:center;
  background:radial-gradient(1200px 700px at 50% -10%,#ffffff 0%,#eef2f7 55%,#e6ecf3 100%);overflow:hidden}
.login-glow{position:absolute;width:520px;height:520px;border-radius:50%;
  background:radial-gradient(circle,color-mix(in srgb,var(--accent) 20%,transparent) 0%,transparent 70%);
  filter:blur(22px);animation:glow 6s ease-in-out infinite;pointer-events:none}
.login-box{position:relative;width:min(92vw,430px);background:#ffffff;
  border:1px solid #e6eaf0;border-radius:18px;padding:40px 42px 26px;
  box-shadow:0 30px 80px rgba(25,40,65,.14),inset 0 1px 0 rgba(255,255,255,.6);animation:fade .5s ease both}
.login-brand{display:flex;align-items:center;gap:12px;margin-bottom:24px}
.login-brand .brandmark,.login-brand .glyph{width:40px;height:40px}
.login-brand img{width:40px;height:40px}
.login-name{font-family:var(--disp);font-weight:700;font-size:16px;letter-spacing:.5px;color:var(--t1);line-height:1}
.login-tag{font-size:9.5px;letter-spacing:2px;color:var(--accent);font-weight:600;margin-top:5px}
.login-h{font-family:var(--disp);font-size:22px;font-weight:600;color:var(--t1);letter-spacing:-.3px}
.login-sub{font-size:12.5px;color:#6b7688;margin:6px 0 22px;line-height:1.5}
.login-lbl{font-size:10px;letter-spacing:1.5px;color:#7a8598;font-weight:600;text-transform:uppercase;margin-top:14px}
.login-inp{width:100%;margin-top:7px;background:#f5f7fa;border:1.5px solid #d5dbe4;border-radius:10px;
  padding:12px 15px;color:var(--t1);font-size:14.5px;font-family:var(--sans);outline:none;transition:border .2s}
#pwd{font-family:var(--mono);letter-spacing:3px}
.login-inp:focus{border-color:var(--accent)}
.login-btn{width:100%;margin-top:20px;background:linear-gradient(135deg,var(--accent),#e06a24);
  color:#fff;border:none;border-radius:10px;padding:14px;font-size:14px;font-weight:700;
  font-family:var(--disp);letter-spacing:.4px;cursor:pointer;transition:transform .15s,box-shadow .2s;
  box-shadow:0 10px 26px color-mix(in srgb,var(--accent) 33%,transparent)}
.login-btn:hover{transform:translateY(-1px)}
.login-back{display:inline-flex;align-items:center;gap:6px;background:none;border:none;color:#8791a3;
  font-size:12px;cursor:pointer;margin-top:16px;padding:4px}
.login-back:hover{color:var(--accent)}
.login-err{min-height:18px;margin-top:12px;font-size:12.5px;color:#dc3a50;text-align:center;font-weight:500}
.login-copy{margin-top:16px;padding-top:14px;border-top:1px solid #eaedf2;font-size:9px;color:#aab3c0;text-align:center;line-height:1.6;letter-spacing:.2px}

/* SHELL */
#app{display:none}
#header{position:fixed;top:0;left:0;right:0;height:var(--head);z-index:100;display:flex;align-items:center;
  gap:18px;padding:0 22px;background:rgba(255,255,255,.9);backdrop-filter:blur(14px);border-bottom:1px solid #e6eaf0;box-shadow:0 1px 3px rgba(20,30,50,.03)}
.hbrand{display:flex;align-items:center;gap:11px;flex:none;cursor:pointer}
.hbrand .brandmark,.hbrand .glyph{width:32px;height:32px}
.hbrand img{width:32px;height:32px}
.hname{font-family:var(--disp);font-weight:700;font-size:15px;letter-spacing:.4px}
.htag{font-size:8.5px;letter-spacing:2px;color:var(--accent);font-weight:600;margin-top:2px}
.hchip{font-size:9px;letter-spacing:1.5px;font-weight:700;border-radius:5px;padding:3px 8px;
  color:var(--accent);border:1px solid color-mix(in srgb,var(--accent) 45%,transparent)}
.hchip.comp{color:var(--red);border-color:color-mix(in srgb,var(--red) 45%,transparent)}
.hsearch{position:relative;flex:1;max-width:440px;margin-left:12px}
.hsearch input{width:100%;background:#f5f7fa;border:1px solid var(--border2);border-radius:9px;
  padding:9px 14px 9px 36px;color:var(--t1);font-size:13px;outline:none;transition:border .2s}
.hsearch input:focus{border-color:color-mix(in srgb,var(--accent) 55%,transparent)}
.hsearch .ic{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:#98a2b2;font-size:14px}
.hright{margin-left:auto;display:flex;align-items:center;gap:14px;flex:none}
.hts{display:flex;flex-direction:column;align-items:flex-end;line-height:1.2}
.hts .k{font-size:9px;letter-spacing:1.5px;color:#96a0b0;font-weight:600}
.hts .v{font-size:11.5px;color:var(--t2);font-family:var(--mono)}
.hlogout{background:#f0f3f7;border:1px solid #dce1e9;color:var(--t2);padding:7px 14px;border-radius:8px;
  font-size:12px;font-weight:600;cursor:pointer;transition:all .18s}
.hlogout:hover{border-color:#e0455e;color:#e0455e}
.menu-btn{display:none;background:none;border:none;color:var(--t1);font-size:22px;cursor:pointer}

#sidebar{position:fixed;top:calc(var(--head) + var(--tick));left:0;width:var(--sid);height:calc(100vh - var(--head) - var(--tick));
  overflow-y:auto;background:var(--panel);border-right:1px solid #e6eaf0;z-index:90;padding:18px 0 30px;transition:transform .25s}
.sid-sec{padding:0 14px;margin-top:18px}
.sid-sec:first-child{margin-top:0}
.sid-lbl{font-size:9px;letter-spacing:2px;color:var(--t4);font-weight:700;padding:8px 8px 8px}
.sid-item{display:flex;align-items:center;gap:12px;padding:11px 13px;border-radius:9px;cursor:pointer;
  font-size:13.5px;font-weight:500;margin-bottom:3px;transition:all .15s;color:var(--t2);border-left:2px solid transparent}
.sid-item:hover{background:#f2f5f9;color:var(--t1)}
.sid-item.active{background:color-mix(in srgb,var(--accent) 12%,transparent);color:var(--t1);border-left-color:var(--accent);font-weight:600}
.sid-item .ico{width:18px;text-align:center;font-size:15px;color:var(--t4)}
.sid-item.active .ico{color:var(--accent)}
.sid-item .nm{flex:1}
.sid-hint{padding:14px 16px;margin:18px 14px 0;background:#f7f9fc;border:1px solid var(--border);border-radius:11px}
.sid-hint .h{font-family:var(--disp);font-size:12px;font-weight:600;color:var(--t1);margin-bottom:5px}
.sid-hint .p{font-size:11px;color:#8791a3;line-height:1.5}

#main{margin-left:var(--sid);padding-top:calc(var(--head) + var(--tick));min-height:100vh}

/* DASHBOARD */
#view-dash{display:none;padding:30px 34px 40px;animation:fade .5s ease both}
.dash-head{display:flex;align-items:flex-end;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:26px}
.eyebrow{font-size:10px;letter-spacing:2.5px;color:var(--accent);font-weight:700;margin-bottom:7px}
.dash-title{font-family:var(--disp);font-size:28px;font-weight:600;letter-spacing:-.6px;color:var(--t1)}
.dash-sub{font-size:13px;color:#6b7688;margin-top:6px;max-width:640px}
.synced{display:flex;align-items:center;gap:8px;background:#ffffff;border:1px solid var(--border);border-radius:10px;padding:9px 14px;font-size:11.5px;color:var(--t2);box-shadow:0 1px 2px rgba(20,30,50,.04)}
.synced .live{width:8px;height:8px;border-radius:50%;background:#16a06a;box-shadow:0 0 10px rgba(22,160,106,.5);animation:glow 2s infinite}
.synced .v{color:var(--t1);font-family:var(--mono)}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:14px;margin-bottom:26px}
.kpi{position:relative;overflow:hidden;background:#ffffff;border:1px solid var(--border);border-radius:14px;padding:18px 20px;cursor:pointer;transition:all .2s;box-shadow:0 1px 2px rgba(20,30,50,.04)}
.kpi:hover{transform:translateY(-2px);box-shadow:0 10px 26px rgba(20,30,50,.08)}
.kpi .halo{position:absolute;top:-30px;right:-30px;width:90px;height:90px;border-radius:50%}
.kpi .top{display:flex;align-items:center;gap:8px;margin-bottom:12px}
.kpi .sq{width:9px;height:9px;border-radius:3px}
.kpi .lbl{font-size:10.5px;letter-spacing:1px;color:#6b7688;font-weight:600;text-transform:uppercase}
.kpi .num{font-family:var(--disp);font-size:36px;font-weight:700;letter-spacing:-1px;color:var(--t1);line-height:1}
.kpi .sub{font-size:11.5px;color:#8791a3;margin-top:6px}
.charts{display:grid;grid-template-columns:1.35fr 1fr;gap:16px;margin-bottom:16px;align-items:start}
.panel-card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:22px;box-shadow:0 1px 2px rgba(20,30,50,.04)}
.pc-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px}
.pc-title{font-family:var(--disp);font-size:15px;font-weight:600;color:var(--t1)}
.pc-meta{font-size:10.5px;color:#8791a3;font-family:var(--mono)}
.bar-row{display:flex;align-items:center;gap:12px;margin-bottom:9px}
.bar-row.clk{cursor:pointer}
.bar-row.clk:hover{opacity:.82}
.bar-lbl{width:150px;font-size:12px;color:#4c5769;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar-lbl.sm{width:120px;font-size:11.5px}
.bar-lbl.xs{width:80px;font-size:11.5px}
.bar-track{flex:1;height:9px;background:#eef1f6;border-radius:6px;overflow:hidden}
.bar-track.sm{height:8px}
.bar-fill{height:100%;border-radius:6px;width:0;transition:width 1.1s cubic-bezier(.2,.8,.2,1)}
.bar-num{width:40px;text-align:right;font-family:var(--mono);font-size:12px;color:var(--t1);font-weight:600}
.col-2{display:flex;flex-direction:column;gap:16px}
.strip{background:#ffffff;border:1px solid var(--border);border-radius:14px;padding:22px 24px;box-shadow:0 1px 2px rgba(20,30,50,.04)}
.strip-row{display:flex;gap:10px;flex-wrap:wrap}
.strip-cell{flex:1;min-width:150px;background:#f7f9fc;border:1px solid var(--border);border-radius:10px;padding:14px 16px}
.strip-cell .n{font-family:var(--mono);font-size:26px;font-weight:600;color:var(--t1)}
.strip-cell .l{font-size:11.5px;color:#6b7688;margin-top:4px}
.note-card{display:flex;gap:16px;align-items:flex-start;background:linear-gradient(180deg,#fff,#fbfcfe);border:1px solid var(--border);border-radius:14px;padding:20px 24px;box-shadow:0 1px 2px rgba(20,30,50,.04)}
.note-ic{width:30px;height:30px;border-radius:8px;flex:none;background:color-mix(in srgb,var(--accent) 14%,#fff);color:var(--accent);display:flex;align-items:center;justify-content:center;font-family:var(--disp);font-weight:700;font-style:italic;font-size:15px}
.note-h{font-family:var(--disp);font-size:14px;font-weight:600;color:var(--t1);margin-bottom:5px}
.note-p{font-size:12.5px;color:#6b7688;line-height:1.6;max-width:760px}
.dash-sec-lbl{display:flex;align-items:center;gap:14px;font-family:var(--disp);font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--t3);margin:28px 0 14px}
.dash-sec-lbl:after{content:"";flex:1;height:1px;background:var(--border)}
.dash-sec-lbl .n{width:20px;height:20px;border-radius:6px;background:color-mix(in srgb,var(--accent) 14%,#fff);color:var(--accent);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800;flex:none}
.app-footer{padding:6px 34px 42px;margin-top:14px}
.app-footer .note-card{max-width:1100px;margin:0 auto}
.af-copy{text-align:center;font-size:10px;color:#aab3c0;letter-spacing:.2px;line-height:1.6;margin-top:16px}

/* ── NEWS / TICKER ───────────────────────────────────────────── */
#ticker{position:fixed;top:var(--head);left:0;right:0;height:34px;z-index:95;background:#141a26;overflow:hidden;display:none;align-items:center}
#ticker.on{display:flex}
#ticker:before{content:"MARKET";position:absolute;left:0;top:0;height:100%;display:flex;align-items:center;padding:0 14px;background:var(--accent);color:#fff;font-family:var(--disp);font-size:10px;font-weight:700;letter-spacing:1.5px;z-index:2}
.tick-track{display:inline-flex;align-items:center;white-space:nowrap;will-change:transform;animation:tickerscroll 70s linear infinite;padding-left:80px}
#ticker:hover .tick-track{animation-play-state:paused}
.tk-item{display:inline-flex;align-items:center;gap:7px;padding:0 20px;font-size:12px;color:#c7cfdb;border-right:1px solid rgba(255,255,255,.09)}
.tk-item b{color:#fff;font-weight:600}
.tk-item .val{font-family:var(--mono);color:#e7ecf3}
.tk-item .arr{font-size:10px;font-family:var(--mono)}
.tk-item .up{color:#37d67a}
.tk-item .down{color:#ff6b7d}
.tk-item .flat{color:#9aa6b6}
.tk-dot{width:8px;height:8px;border-radius:50%;flex:none}
.tk-dot.red{background:#ff5a6e;box-shadow:0 0 6px rgba(255,90,110,.7)}
.tk-dot.green{background:#37d67a;box-shadow:0 0 6px rgba(55,214,122,.7)}
.tk-dot.amber{background:#f5b642;box-shadow:0 0 6px rgba(245,182,66,.7)}
@keyframes tickerscroll{from{transform:translateX(0)}to{transform:translateX(-50%)}}
.newsdash-ind{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0 6px}
.ndi{background:var(--card2);border:1px solid var(--border);border-radius:10px;padding:10px 12px}
.ndi .l{font-size:9.5px;letter-spacing:.5px;color:#8791a3;font-weight:700;text-transform:uppercase}
.ndi .v{font-family:var(--mono);font-size:14px;font-weight:600;color:var(--t1);margin-top:3px;display:flex;align-items:center;gap:6px}
.nd-list{display:flex;flex-direction:column;gap:2px;margin-top:8px}
.nd-row{display:flex;align-items:flex-start;gap:11px;padding:11px 8px;border-radius:9px;cursor:pointer;transition:background .14s}
.nd-row:hover{background:#f5f7fa}
.nd-dot{width:9px;height:9px;border-radius:50%;flex:none;margin-top:5px}
.nd-row .nm{font-size:13px;font-weight:600;color:var(--t1);line-height:1.35}
.nd-row .mt{font-size:11.5px;color:#7b8698;margin-top:2px;line-height:1.45}
.nd-row .tag{font-size:9.5px;font-weight:700;letter-spacing:.4px;color:#8791a3;text-transform:uppercase;margin-bottom:2px}
.nd-more{display:inline-flex;align-items:center;gap:6px;margin-top:10px;font-size:12px;font-weight:700;color:var(--accent);cursor:pointer}
.news-ind-strip{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin:22px 0 18px}
.nis{background:#fff;border:1px solid var(--border);border-radius:14px;padding:15px 16px}
.nis .l{font-size:10px;letter-spacing:.5px;color:#8791a3;font-weight:700;text-transform:uppercase}
.nis .v{font-family:var(--mono);font-size:16px;font-weight:600;color:var(--t1);margin-top:6px;line-height:1.15}
.nis .u{font-size:10px;color:#96a0b0}
.nis .mv{display:inline-flex;align-items:center;gap:5px;margin-top:9px;font-size:10.5px;font-weight:700;border-radius:20px;padding:3px 9px}
.nis .mv.up{background:#e6f7ee;color:#0f8a56}
.nis .mv.down{background:#fdeaec;color:#c93350}
.nis .mv.flat{background:#eef1f6;color:#5b6678}
.nis .dr{font-size:10.5px;color:#8791a3;margin-top:9px;line-height:1.45}
.news-trend-card{margin-bottom:20px}
.nt-svg{width:100%;height:160px;display:block;margin-top:6px}
.news-ed{margin-bottom:28px}
.news-ed-head{display:flex;align-items:center;gap:12px;margin:6px 0 4px}
.news-ed-date{font-family:var(--disp);font-size:20px;font-weight:600;color:var(--t1)}
.news-ed-badge{font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;background:color-mix(in srgb,var(--accent) 12%,#fff);color:var(--accent);border-radius:7px;padding:4px 10px}
.news-ed-line{height:1px;background:var(--border);flex:1}
.news-note{display:flex;gap:12px;background:linear-gradient(180deg,#fff,#fbfcfe);border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:10px;padding:14px 16px;margin:12px 0 18px}
.news-note .h{font-size:10px;letter-spacing:1px;color:var(--accent);font-weight:700;text-transform:uppercase;margin-bottom:5px}
.news-note .p{font-size:12.5px;color:#5b6678;line-height:1.6}
.news-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.ncard{background:#fff;border:1px solid var(--border);border-radius:14px;overflow:hidden;display:flex;flex-direction:column}
.ncard .top{display:flex;align-items:center;gap:9px;padding:12px 15px;border-bottom:1px solid var(--border)}
.ncard .sbar{width:4px;align-self:stretch;border-radius:4px;flex:none}
.ncard .seg{font-family:var(--disp);font-size:13.5px;font-weight:600;color:var(--t1)}
.ncard .co{font-size:11px;color:#8791a3;margin-left:auto;font-weight:600;text-align:right;max-width:44%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ncard .stat{font-size:9px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;border-radius:6px;padding:3px 7px;flex:none}
.ncard .bd{padding:13px 15px 15px}
.ncard .hl{font-family:var(--disp);font-size:14px;font-weight:600;color:var(--t1);line-height:1.35;margin-bottom:9px}
.ncard .k{font-size:9px;letter-spacing:.6px;color:#98a2b2;font-weight:700;text-transform:uppercase;margin:9px 0 3px}
.ncard .t{font-size:12px;color:#4a5567;line-height:1.55}
.ncard .why{background:var(--card2);border-radius:8px;padding:9px 11px;margin-top:10px}
.ncard .src{font-size:10px;color:#a7b0be;margin-top:10px;font-style:italic}
.st-red{background:#fdeaec;color:#c93350}.st-yellow{background:#fdf3e0;color:#b5790a}.st-green{background:#e6f7ee;color:#0f8a56}.st-neutral{background:#eef1f6;color:#5b6678}
.sb-red{background:#e0455e}.sb-yellow{background:#c58a10}.sb-green{background:#12a05f}.sb-neutral{background:#96a0b0}
@media(max-width:900px){.news-ind-strip{grid-template-columns:1fr 1fr}.news-grid{grid-template-columns:1fr}.newsdash-ind{grid-template-columns:1fr}}

/* GENERIC PAGE HEADER (search sub-pages) */
.page{display:none;animation:fade .4s ease both}
.page-wrap{max-width:1180px;margin:0 auto;padding:36px 34px 60px}
.crumb{display:flex;align-items:center;gap:8px;font-size:12px;color:#8791a3;margin-bottom:18px}
.crumb .lk{cursor:pointer;color:var(--t2);font-weight:500}
.crumb .lk:hover{color:var(--accent)}
.crumb .sep{color:#c3ccd8}
.page-eyebrow{font-size:10.5px;letter-spacing:2.5px;color:var(--accent);font-weight:700;margin-bottom:9px}
.page-title{font-family:var(--disp);font-size:30px;font-weight:600;letter-spacing:-.6px;color:var(--t1)}
.page-sub{font-size:13.5px;color:#6b7688;margin-top:8px;max-width:620px;line-height:1.55}

/* SEARCH HUB */
.hub-wrap{min-height:calc(100vh - var(--head));display:flex;flex-direction:column;align-items:center;justify-content:center;padding:60px 34px}
.hub-mark{width:56px;height:56px;margin-bottom:24px}
.hub-mark img{width:56px;height:56px;border-radius:12px}
.hub-eyebrow{font-size:11px;letter-spacing:3px;color:var(--accent);font-weight:700;margin-bottom:14px}
.hub-title{font-family:var(--disp);font-size:34px;font-weight:600;letter-spacing:-.8px;color:var(--t1);text-align:center}
.hub-sub{font-size:14px;color:#6b7688;margin-top:12px;text-align:center;max-width:520px;line-height:1.6}
.hub-grid{display:flex;flex-direction:column;gap:16px;margin-top:40px;width:100%;max-width:640px}
.hub-btn{position:relative;overflow:hidden;border:none;border-radius:20px;
  padding:24px 26px;cursor:pointer;transition:transform .18s,box-shadow .2s;text-align:left;
  display:flex;align-items:flex-start;gap:20px;color:#fff;box-shadow:0 12px 30px rgba(20,30,50,.10)}
.hub-btn:hover{transform:translateY(-3px);box-shadow:0 22px 46px rgba(20,30,50,.18)}
.hub-btn .halo{position:absolute;top:-56px;right:-30px;width:190px;height:190px;border-radius:50%;background:rgba(255,255,255,.10);pointer-events:none}
.hub-btn .ic{width:60px;height:60px;border-radius:16px;flex:none;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,.20);color:#fff;position:relative}
.hub-btn .ic svg{width:30px;height:30px}
.hub-btn .body{position:relative;flex:1;min-width:0;display:flex;flex-direction:column}
.hub-btn .t{font-family:var(--disp);font-size:21px;font-weight:600;color:#fff;letter-spacing:-.3px}
.hub-btn .d{font-size:13px;color:rgba(255,255,255,.86);margin-top:6px;line-height:1.5}
.hub-btn .go{margin-top:16px;font-size:12.5px;font-weight:600;color:#fff;display:flex;align-items:center;gap:6px}
.hub-btn .go .arw{transition:transform .2s}
.hub-btn:hover .go .arw{transform:translateX(4px)}
.hub-btn.c-region{background:linear-gradient(135deg,#5b6bd6,#4055c2)}
.hub-btn.c-segments{background:linear-gradient(135deg,#e9863b,#d76d20)}
.hub-btn.c-company{background:linear-gradient(135deg,#8265f2,#6a4bdf)}
.hub-btn.c-exh{background:linear-gradient(135deg,#17a866,#0d8850)}
.hub-btn .cardart{position:absolute;right:-4px;top:50%;transform:translateY(-50%);width:138px;height:138px;opacity:.17;pointer-events:none;color:#fff}
.hub-btn .cardart svg{width:100%;height:100%;display:block}
.hub-btn.c-region .cardart{opacity:.24;right:14px;width:118px;height:118px}
.hub-exh-sec{width:100%;max-width:900px;margin-top:26px}
.hub-exh-eyebrow{font-size:10.5px;letter-spacing:2.5px;color:var(--accent);font-weight:700;margin-bottom:12px;text-align:left}
.hub-exh{position:relative;overflow:hidden;width:100%;display:flex;align-items:center;gap:22px;background:linear-gradient(120deg,#fff,#f7f9fc);border:1px solid var(--border);border-radius:18px;padding:26px 30px;cursor:pointer;transition:all .2s;text-align:left;box-shadow:0 2px 8px rgba(20,30,50,.04)}
.hub-exh:hover{transform:translateY(-3px);border-color:color-mix(in srgb,var(--green) 45%,transparent);box-shadow:0 18px 40px rgba(20,30,50,.1)}
.hub-exh .hx-ic{width:56px;height:56px;border-radius:14px;flex:none;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--green) 13%,#fff);color:var(--green)}
.hub-exh .hx-ic svg{width:28px;height:28px}
.hub-exh .hx-txt{flex:1;min-width:0}
.hub-exh .hx-t{font-family:var(--disp);font-size:20px;font-weight:600;color:var(--t1);letter-spacing:-.3px}
.hub-exh .hx-d{font-size:12.5px;color:#7b8698;margin-top:5px;line-height:1.5}
.hub-exh .hx-go{flex:none;font-size:12.5px;font-weight:600;color:var(--green);display:flex;align-items:center;gap:6px}
.hub-exh .hx-go .arw{transition:transform .2s}
.hub-exh:hover .hx-go .arw{transform:translateX(4px)}
@media(max-width:700px){.hub-exh{flex-wrap:wrap;gap:14px}.hub-exh .hx-go{width:100%}}
.hub-exh-sec+.hub-exh-sec{margin-top:16px}
.hub-exh.news .hx-ic{background:color-mix(in srgb,var(--accent) 13%,#fff);color:var(--accent)}
.hub-exh.news .hx-go{color:var(--accent)}
.hub-exh.news:hover{border-color:color-mix(in srgb,var(--accent) 45%,transparent)}
.hub-feat{position:relative;overflow:hidden;width:100%;display:flex;align-items:center;gap:20px;background:#fff;border:1px solid var(--border);border-radius:20px;padding:22px 24px;cursor:pointer;text-align:left;transition:transform .18s,box-shadow .2s,border-color .2s;box-shadow:0 8px 24px rgba(20,30,50,.05)}
.hub-feat:hover{transform:translateY(-3px);box-shadow:0 20px 44px rgba(20,30,50,.12)}
.hub-feat .feat-art{position:absolute;right:150px;top:50%;transform:translateY(-50%);width:140px;height:110px;opacity:.09;pointer-events:none}
.hub-feat .feat-art svg{width:100%;height:100%;display:block}
.hub-feat .feat-med{width:60px;height:60px;border-radius:16px;flex:none;display:flex;align-items:center;justify-content:center;color:#fff;box-shadow:0 8px 20px rgba(20,30,50,.14)}
.hub-feat .feat-med svg{width:30px;height:30px}
.hub-feat .feat-txt{flex:1;min-width:0;position:relative}
.hub-feat .feat-t{font-family:var(--disp);font-size:20px;font-weight:600;color:var(--t1);letter-spacing:-.3px}
.hub-feat .feat-d{font-size:12.5px;color:#7b8698;margin-top:5px;line-height:1.5;max-width:520px}
.hub-feat .feat-cta{flex:none;position:relative;display:inline-flex;align-items:center;gap:7px;font-size:12.5px;font-weight:700;color:#fff;border-radius:11px;padding:11px 18px;white-space:nowrap}
.hub-feat .feat-cta .arw{transition:transform .2s}
.hub-feat:hover .feat-cta .arw{transform:translateX(4px)}
.feat-exh{border-color:color-mix(in srgb,var(--green) 24%,var(--border))}
.feat-exh .feat-med,.feat-exh .feat-cta{background:linear-gradient(135deg,#17a866,#0d8850)}
.feat-exh .feat-art{color:var(--green)}
.feat-exh:hover{border-color:color-mix(in srgb,var(--green) 55%,transparent)}
.feat-news{border-color:color-mix(in srgb,var(--accent) 24%,var(--border))}
.feat-news .feat-med,.feat-news .feat-cta{background:linear-gradient(135deg,#e9863b,#d76d20)}
.feat-news .feat-art{color:var(--accent)}
.feat-news:hover{border-color:color-mix(in srgb,var(--accent) 55%,transparent)}
@media(max-width:640px){.hub-feat{flex-wrap:wrap;gap:14px}.hub-feat .feat-art{display:none}.hub-feat .feat-cta{width:100%;justify-content:center}}

/* REGION MAP */
.region-layout{display:grid;grid-template-columns:1.05fr .95fr;gap:30px;align-items:center;margin-top:14px}
.map-frame{position:relative;background:#fff;border:1px solid var(--border);border-radius:20px;padding:22px;box-shadow:0 2px 10px rgba(20,30,50,.05)}
.map-inner{position:relative;width:100%;max-width:440px;margin:0 auto}
.map-inner svg{width:100%;height:auto;display:block;overflow:visible}
.rzone-path{cursor:pointer;stroke:#ffffff;stroke-width:.6;stroke-linejoin:round;transition:filter .16s,opacity .16s}
.rzone-path:hover{filter:brightness(1.07) saturate(1.18)}
.rzone-lbl{pointer-events:none;font-family:var(--disp);font-weight:700;fill:#ffffff;text-anchor:middle;paint-order:stroke;stroke:rgba(22,32,52,.32);stroke-width:2.6px;stroke-linejoin:round}
.rzone-lbl .zn{font-size:15px}
.rzone-lbl .zc{font-size:12px;font-weight:600}
.map-legend{display:flex;flex-wrap:wrap;gap:6px;justify-content:center;margin-top:16px}
.map-legend .lg{display:flex;align-items:center;gap:7px;font-size:11.5px;color:var(--t2);cursor:pointer;padding:5px 10px;border-radius:8px;transition:background .15s;font-weight:500}
.map-legend .lg:hover{background:#f2f5f9;color:var(--t1)}
.map-legend .lg .sw{width:11px;height:11px;border-radius:3px;flex:none}
.map-legend .lg .lc{font-family:var(--mono);font-size:11px;color:#96a0b0}
/* regional pie */
.pie-wrap{display:flex;align-items:center;gap:20px;flex-wrap:wrap}
.pie-svg{flex:none}
.pie-slice{cursor:pointer;transition:opacity .16s}
.pie-slice:hover{opacity:.82}
.pie-legend{flex:1;min-width:140px;display:flex;flex-direction:column;gap:6px}
.pie-lg{display:flex;align-items:center;gap:10px;cursor:pointer;padding:6px 8px;border-radius:8px;transition:background .15s}
.pie-lg:hover{background:#f2f5f9}
.pie-lg .sw{width:10px;height:10px;border-radius:3px;flex:none}
.pie-lg .nm{font-size:12.5px;color:#4c5769;flex:1}
.pie-lg .vl{font-family:var(--mono);font-size:12px;font-weight:600;color:var(--t1)}
.pie-lg .pc{font-family:var(--mono);font-size:10.5px;color:#96a0b0;width:36px;text-align:right}
/* coverage by segment (market-grouped) */
.seg-cov{display:grid;grid-template-columns:1fr 1fr;gap:2px 26px;max-height:430px;overflow-y:auto;padding-right:6px;margin:-2px}
.seg-cov-grp{grid-column:1/-1;display:flex;align-items:center;gap:9px;margin-top:13px;padding:0 2px 6px;border-bottom:1px solid var(--border)}
.seg-cov-grp:first-child{margin-top:2px}
.seg-cov-grp .mk-dot{width:9px;height:9px;border-radius:3px;flex:none}
.seg-cov-grp .mk-nm{font-family:var(--disp);font-size:11.5px;font-weight:700;letter-spacing:.4px;color:var(--t1);text-transform:uppercase}
.seg-cov-grp .mk-tot{margin-left:auto;font-family:var(--mono);font-size:11px;color:#96a0b0}
.seg-cov-row{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:8px;cursor:pointer;transition:background .14s}
.seg-cov-row:hover{background:#f5f7fa}
.seg-cov-row .nm{flex:1;font-size:12.5px;color:#3a4557;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.seg-cov-row .ct{font-family:var(--mono);font-size:12px;font-weight:600;color:var(--t1);background:#eff2f6;border-radius:20px;padding:2px 9px;min-width:30px;text-align:center;flex:none}
.seg-cov-row:hover .ct{background:color-mix(in srgb,var(--accent) 14%,transparent);color:var(--accent)}
/* sub-page back button */
.page-back{display:inline-flex;align-items:center;gap:7px;background:#fff;border:1px solid var(--border2);color:var(--t2);border-radius:9px;padding:8px 15px;font-size:12.5px;font-weight:600;cursor:pointer;transition:all .15s;margin-bottom:16px}
.page-back:hover{border-color:var(--accent);color:var(--accent)}
.region-side .rs-item{display:flex;align-items:center;gap:14px;padding:14px 16px;background:#fff;border:1px solid var(--border);
  border-radius:12px;margin-bottom:10px;cursor:pointer;transition:all .16s}
.region-side .rs-item:hover{transform:translateX(3px);border-color:color-mix(in srgb,var(--accent) 45%,transparent);box-shadow:0 8px 20px rgba(20,30,50,.07)}
.region-side .rs-sq{width:38px;height:38px;border-radius:10px;flex:none;display:flex;align-items:center;justify-content:center;
  font-family:var(--disp);font-weight:700;font-size:15px;color:#fff}
.region-side .rs-nm{font-family:var(--disp);font-size:15px;font-weight:600;color:var(--t1)}
.region-side .rs-mt{font-size:11.5px;color:#8791a3;margin-top:2px}
.region-side .rs-ct{margin-left:auto;font-family:var(--mono);font-size:15px;font-weight:600;color:var(--t2)}
.region-side .rs-arw{color:#c3ccd8;font-size:18px}

/* SEGMENTS FRAMEWORK */
.mkt-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-top:14px}
.mkt-card{background:#fff;border:1px solid var(--border);border-radius:16px;overflow:hidden;box-shadow:0 2px 8px rgba(20,30,50,.04);display:flex;flex-direction:column}
.mkt-top{display:flex;align-items:center;gap:12px;padding:18px 20px;border-bottom:1px solid var(--border)}
.mkt-ic{width:40px;height:40px;border-radius:11px;flex:none;display:flex;align-items:center;justify-content:center}
.mkt-ic svg{width:22px;height:22px}
.mkt-nm{font-family:var(--disp);font-size:16px;font-weight:600;color:var(--t1);letter-spacing:-.2px}
.mkt-ct{font-size:11px;color:#8791a3;margin-top:2px;font-family:var(--mono)}
.mkt-body{padding:12px;display:flex;flex-direction:column;gap:4px;flex:1}
.seg-chip{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:9px;cursor:pointer;transition:all .14s}
.seg-chip:hover{background:#f5f7fa}
.seg-chip.void{cursor:default;opacity:.42}
.seg-chip.void:hover{background:transparent}
.seg-chip .sn{flex:1;font-size:13px;color:#3a4557;font-weight:500}
.seg-chip .sc{font-family:var(--mono);font-size:11px;color:var(--t2);background:#eff2f6;border-radius:20px;padding:2px 8px;min-width:26px;text-align:center}
.seg-chip.void .sc{background:transparent;color:#c3ccd8}
.seg-chip:hover .sc{background:color-mix(in srgb,var(--accent) 14%,transparent);color:var(--accent)}
.seg-sub{padding:2px 12px 6px 40px;display:flex;flex-direction:column;gap:2px}
.seg-sub .ss{font-size:11.5px;color:#8791a3;display:flex;align-items:center;gap:7px}
.seg-sub .ss:before{content:"";width:4px;height:4px;border-radius:50%;background:#c3ccd8;flex:none}

/* COMPANY SEARCH */
.cs-wrap{max-width:760px;margin:0 auto;padding-top:8px}
.cs-bar{position:relative;margin-top:8px}
.cs-bar input{width:100%;background:#fff;border:1.5px solid var(--border2);border-radius:16px;
  padding:20px 22px 20px 58px;font-size:17px;color:var(--t1);outline:none;font-family:var(--sans);
  box-shadow:0 6px 26px rgba(20,30,50,.07);transition:border .2s,box-shadow .2s}
.cs-bar input:focus{border-color:color-mix(in srgb,var(--accent) 55%,transparent);box-shadow:0 10px 34px rgba(20,30,50,.1)}
.cs-bar .ic{position:absolute;left:22px;top:50%;transform:translateY(-50%);color:var(--accent);font-size:20px}
.cs-hints{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:18px}
.cs-hint{font-size:11.5px;color:#7b8698;background:#fff;border:1px solid var(--border);border-radius:20px;padding:5px 13px;cursor:pointer;transition:all .15s}
.cs-hint:hover{border-color:var(--accent);color:var(--accent)}
.cs-res{margin-top:26px;display:flex;flex-direction:column;gap:8px}
.cs-count{font-size:12px;color:#8791a3;margin-bottom:4px}
.cs-count b{color:var(--t1);font-family:var(--mono)}
.cs-row{display:flex;align-items:center;gap:14px;padding:13px 16px;background:#fff;border:1px solid var(--border);border-radius:11px;cursor:pointer;transition:all .14s}
.cs-row:hover{border-color:color-mix(in srgb,var(--accent) 45%,transparent);transform:translateX(3px);box-shadow:0 8px 20px rgba(20,30,50,.06)}
.cs-row .cn{font-family:var(--disp);font-size:14.5px;font-weight:600;color:var(--t1)}
.cs-row .cm{font-size:11.5px;color:#8791a3;margin-top:2px}
.cs-row .ct{margin-left:auto;font-size:10.5px;font-weight:600;border-radius:6px;padding:3px 9px;background:#eaf1fd;color:#2a63cf;white-space:nowrap}
.cs-more{text-align:center;margin-top:14px}
.cs-more button{background:none;border:1px solid var(--border2);border-radius:10px;padding:11px 22px;font-size:13px;font-weight:600;color:var(--t2);cursor:pointer;transition:all .15s}
.cs-more button:hover{border-color:var(--accent);color:var(--accent)}

/* LIST */
#view-list{display:none;animation:fade .35s ease both}
.toolbar{position:sticky;top:calc(var(--head) + var(--tick));z-index:80;background:rgba(255,255,255,.92);backdrop-filter:blur(12px);
  border-bottom:1px solid var(--border);padding:14px 34px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.tb-title{font-family:var(--disp);font-size:17px;font-weight:600;color:var(--t1);margin-right:auto;display:flex;align-items:center;gap:10px}
.tb-back{background:#f0f3f7;border:1px solid #dce1e9;color:var(--t2);width:30px;height:30px;border-radius:8px;cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;transition:all .15s}
.tb-back:hover{border-color:var(--accent);color:var(--accent)}
.fgroup{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.sel{background:#f5f7fa;border:1px solid var(--border2);border-radius:8px;padding:8px 10px;color:#4c5769;font-size:12px;cursor:pointer;outline:none;max-width:180px}
.sel:focus{border-color:color-mix(in srgb,var(--accent) 55%,transparent)}
.sel.sort{background:#fff;border-color:color-mix(in srgb,var(--accent) 30%,transparent);color:var(--t1);font-weight:600}
.sortwrap{display:flex;align-items:center;gap:6px}
.sortwrap .lb{font-size:10px;letter-spacing:1px;color:#96a0b0;font-weight:700;text-transform:uppercase}
.results{padding:13px 34px 4px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.results .c{font-size:12px;color:#6b7688}
.results .c b{color:var(--t1);font-weight:600;font-family:var(--mono)}
.results .clr{font-size:11.5px;color:var(--accent);cursor:pointer;font-weight:600}
.company-list{padding:8px 34px 16px;display:flex;flex-direction:column;gap:7px}
.card{position:relative;display:flex;align-items:center;gap:14px;background:#fff;border:1px solid var(--border);
  border-radius:11px;padding:13px 18px;cursor:pointer;transition:all .14s;box-shadow:0 1px 2px rgba(20,30,50,.03)}
.card:hover{border-color:color-mix(in srgb,var(--accent) 55%,transparent);transform:translateX(3px);box-shadow:0 8px 22px rgba(20,30,50,.08)}
.card .num{font-family:var(--mono);font-size:11px;color:var(--t4);min-width:30px}
.card .cname{font-family:var(--disp);font-size:14.5px;font-weight:600;color:var(--t1);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.card .seg-tag{font-size:11px;font-weight:600;border-radius:7px;padding:5px 12px;white-space:nowrap;background:#eaf1fd;color:#2a63cf;flex:none}
.card .chev{color:#c0c8d4;font-size:19px;flex:none}
.pager{display:flex;justify-content:center;align-items:center;gap:8px;padding:16px 20px 50px;flex-wrap:wrap}
.pg{min-width:36px;padding:8px 12px;border-radius:8px;border:1px solid #dce1e9;background:#fff;
  color:#4c5769;font-size:12.5px;font-family:var(--mono);cursor:pointer;transition:all .15s}
.pg:hover{border-color:var(--accent);color:var(--accent)}
.pg.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.pg-info{font-size:12px;color:#8791a3;margin-left:6px}
.no-res{text-align:center;padding:70px 20px;color:#8791a3}
.no-res .ic{font-size:42px;margin-bottom:12px;opacity:.5}
.no-res .m{font-family:var(--disp);font-size:16px;font-weight:600;color:var(--t2);margin-bottom:6px}

/* DETAIL FULL PAGE */
#view-detail{display:none;animation:fade .3s ease both;height:calc(100vh - var(--head));display:none;flex-direction:column;padding:16px 24px 20px}
#view-detail.on{display:flex}
.det-top{display:flex;align-items:center;gap:16px;padding:4px 4px 14px;flex:none}
.det-back{background:#f0f3f7;border:1px solid #dce1e9;color:var(--t2);width:38px;height:38px;border-radius:10px;cursor:pointer;font-size:17px;display:flex;align-items:center;justify-content:center;transition:all .15s;flex:none}
.det-back:hover{border-color:var(--accent);color:var(--accent)}
.det-logo{width:46px;height:46px;border-radius:11px;background:#fff;border:1px solid var(--border);display:flex;align-items:center;justify-content:center;padding:7px;flex:none;overflow:hidden}
.det-logo img{max-width:100%;max-height:100%;object-fit:contain}
.det-head-txt{min-width:0;flex:1}
.det-eyebrow{font-size:9.5px;letter-spacing:2px;color:var(--accent);font-weight:700;margin-bottom:4px}
.det-name{font-family:var(--disp);font-size:23px;font-weight:600;letter-spacing:-.4px;color:var(--t1);line-height:1.15;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.det-tags{display:flex;gap:6px;flex:none;flex-wrap:wrap;justify-content:flex-end;max-width:40%}
.dtag{font-size:10.5px;font-weight:600;border-radius:7px;padding:4px 10px;white-space:nowrap}
.det-grid{flex:1;min-height:0;display:grid;grid-template-columns:1fr 1fr;grid-template-rows:minmax(0,1.15fr) minmax(0,.9fr) minmax(0,1fr);gap:14px}
.dcard{background:#fff;border:1px solid var(--border);border-radius:14px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 1px 3px rgba(20,30,50,.04);min-height:0}
.dcard.wide{grid-column:1/-1}
.dcard-head{display:flex;align-items:center;gap:10px;padding:13px 18px;border-bottom:1px solid var(--border);flex:none}
.dcard-bar{width:4px;height:16px;border-radius:3px;flex:none}
.dcard-title{font-family:var(--disp);font-size:13px;font-weight:600;color:var(--t1);letter-spacing:.2px}
.dcard-badge{margin-left:auto;font-size:9.5px;letter-spacing:1px;color:#a7b0be;font-weight:700;text-transform:uppercase}
.dcard-body{padding:6px 8px;overflow-y:auto;flex:1}
.dsub{font-size:9px;letter-spacing:1.5px;color:var(--accent);font-weight:700;text-transform:uppercase;padding:10px 10px 4px}
.dfields{display:grid;grid-template-columns:1fr 1fr;gap:1px;background:#eef1f6;border-radius:9px;overflow:hidden;margin:4px 8px 8px}
.dfields.three{grid-template-columns:1fr 1fr 1fr}
.dfield{background:#fff;padding:9px 12px;min-width:0}
.dfield.full{grid-column:1/-1}
.dkey{font-size:9px;letter-spacing:.4px;color:#98a2b2;font-weight:600;margin-bottom:3px;text-transform:uppercase}
.dval{font-size:12.5px;color:#3a4557;line-height:1.45;word-break:break-word}
.dval a{text-decoration:none}
.dval.big{font-family:var(--disp);font-size:15px;font-weight:600;color:var(--t1)}
.hl-H{color:#0f8a56;font-weight:700}
.hl-M{color:#b5790a;font-weight:700}
.hl-L{color:#d23b52;font-weight:700}
.contact-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:12px}
.ccard{background:#f7f9fc;border:1px solid var(--border);border-radius:11px;padding:14px 16px}
.ccard .role{font-size:9.5px;letter-spacing:1px;color:var(--accent);font-weight:700;text-transform:uppercase;margin-bottom:7px}
.ccard .who{font-family:var(--disp);font-size:14px;font-weight:600;color:var(--t1);line-height:1.3}
.ccard .cc{font-size:12px;color:#4c5769;margin-top:6px;font-family:var(--mono);word-break:break-all}
.ccard .cc.muted{color:#a7b0be;font-family:var(--sans);font-style:italic}


/* EXHIBITIONS */
.exh-tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:18px}
.exh-search{position:relative;flex:1;min-width:200px;max-width:320px}
.exh-search input{width:100%;background:#f5f7fa;border:1px solid var(--border2);border-radius:9px;padding:9px 14px 9px 34px;font-size:12.5px;color:var(--t1);outline:none;font-family:var(--sans)}
.exh-search input:focus{border-color:color-mix(in srgb,var(--accent) 55%,transparent)}
.exh-search .ic{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:#98a2b2;font-size:13px}
.exh-chips{display:flex;flex-wrap:wrap;gap:8px}
.exh-chip{display:flex;align-items:center;gap:7px;padding:8px 14px;border-radius:20px;border:1.5px solid var(--border2);background:#fff;font-size:12px;font-weight:600;color:var(--t2);cursor:pointer;transition:all .15s}
.exh-chip:hover{border-color:color-mix(in srgb,var(--accent) 40%,transparent)}
.exh-chip .dot{width:8px;height:8px;border-radius:50%;flex:none}
.exh-chip .ct{font-family:var(--mono);font-size:11px;opacity:.8}
.exh-chip.active{color:#fff;border-color:transparent}
.exh-seg-sel{background:#f5f7fa;border:1px solid var(--border2);border-radius:8px;padding:8px 12px;font-size:12px;color:#4c5769;cursor:pointer;outline:none;font-family:var(--sans)}
.exh-regions{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:20px}
.exh-region-card{display:flex;align-items:center;gap:10px;background:#fff;border:1.5px solid var(--border);border-radius:12px;padding:11px 15px;cursor:pointer;transition:all .15s;min-width:118px}
.exh-region-card:hover{border-color:color-mix(in srgb,var(--accent) 45%,transparent);transform:translateY(-2px);box-shadow:0 8px 20px rgba(20,30,50,.07)}
.exh-region-card.active{background:color-mix(in srgb,var(--accent) 10%,#fff);border-color:var(--accent)}
.exh-region-card .rc-ic{width:34px;height:34px;border-radius:9px;flex:none;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--steel) 12%,#fff);color:var(--steel);font-family:var(--disp);font-weight:700;font-size:12px}
.exh-region-card.active .rc-ic{background:color-mix(in srgb,var(--accent) 16%,#fff);color:var(--accent)}
.exh-region-card .rc-nm{font-family:var(--disp);font-size:13px;font-weight:600;color:var(--t1);line-height:1.15}
.exh-region-card .rc-ct{font-family:var(--mono);font-size:10.5px;color:#8791a3;margin-top:2px}
.exh-map-frame{position:relative;background:#eef2f7;border-radius:14px;overflow:hidden;border:1px solid var(--border)}
.exh-map-inner{position:relative;width:100%;aspect-ratio:2/1;background:linear-gradient(160deg,#eef4fb,#e6edf5)}
.exh-world{position:absolute;inset:0;width:100%;height:100%;display:block}
.exh-world path{fill:#d4dde7;stroke:#c3cedb;stroke-width:.25;stroke-linejoin:round}
.exh-pins{position:absolute;inset:0}
.exh-continent{position:absolute;transform:translate(-50%,-50%);font-family:var(--disp);font-size:11px;font-weight:700;letter-spacing:2px;color:rgba(60,75,95,.38);text-transform:uppercase;pointer-events:none;white-space:nowrap}
.exh-clabel{position:absolute;transform:translate(-50%,-50%);font-size:8.5px;font-weight:600;color:rgba(50,62,82,.8);pointer-events:none;white-space:nowrap;text-shadow:0 0 3px #fff,0 1px 2px #fff}
.exh-pin{position:absolute;transform:translate(-50%,-50%);width:13px;height:13px;border-radius:50%;border:2px solid #fff;
  cursor:pointer;box-shadow:0 2px 6px rgba(20,30,50,.35);transition:transform .15s;padding:0}
.exh-pin:hover{transform:translate(-50%,-50%) scale(1.4);z-index:5}
.exh-pin .cnt{position:absolute;top:-6px;right:-7px;background:#1b2333;color:#fff;font-size:8.5px;font-weight:700;
  border-radius:20px;min-width:14px;height:14px;display:flex;align-items:center;justify-content:center;padding:0 3px;font-family:var(--mono)}
.exh-pin-pop{position:absolute;z-index:20;background:#fff;border:1px solid var(--border);border-radius:12px;
  box-shadow:0 16px 40px rgba(20,30,50,.18);padding:8px;min-width:220px;max-width:280px;display:none}
.exh-pin-pop .hd{font-size:10px;letter-spacing:1px;color:#96a0b0;font-weight:700;text-transform:uppercase;padding:4px 8px 8px}
.exh-pin-pop .row{display:flex;align-items:center;gap:8px;padding:8px;border-radius:8px;cursor:pointer;transition:background .14s}
.exh-pin-pop .row:hover{background:#f5f7fa}
.exh-pin-pop .row .nm{flex:1;font-size:12px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.exh-map-legend{display:flex;flex-wrap:wrap;gap:14px;margin-top:14px;padding:0 2px}
.exh-map-legend .lg{display:flex;align-items:center;gap:7px;font-size:11.5px;color:var(--t2);font-weight:500}
.exh-map-legend .lg .sw{width:10px;height:10px;border-radius:50%;flex:none}
.exh-list{display:flex;flex-direction:column;gap:8px;max-height:640px;overflow-y:auto;padding-right:4px}
.exh-card{display:flex;align-items:center;gap:14px;background:#fff;border:1px solid var(--border);border-radius:12px;
  padding:14px 18px;cursor:pointer;transition:all .14s}
.exh-card:hover{border-color:color-mix(in srgb,var(--accent) 50%,transparent);transform:translateX(3px);box-shadow:0 8px 22px rgba(20,30,50,.07)}
.exh-card .body{flex:1;min-width:0}
.exh-card .nm{font-family:var(--disp);font-size:14.5px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.exh-card .mt{margin-top:5px;font-size:11.5px;color:#8791a3;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.exh-card .seg-tag{font-size:10.5px;font-weight:600;border-radius:7px;padding:4px 10px;white-space:nowrap;background:#eff2f6;color:#5b6678;flex:none}
.exh-card .chev{color:#c0c8d4;font-size:19px;flex:none}
.pri-badge{font-size:10px;font-weight:700;letter-spacing:.3px;border-radius:7px;padding:4px 9px;white-space:nowrap;flex:none;text-transform:uppercase}
.exh-empty{text-align:center;padding:50px 20px;color:#8791a3;font-size:13px}
.exh-up-list{display:flex;flex-direction:column;gap:2px;max-height:320px;overflow-y:auto;margin:-2px}
.exh-up-row{display:flex;align-items:center;gap:12px;padding:11px 10px;border-radius:9px;cursor:pointer;transition:background .14s}
.exh-up-row:hover{background:#f5f7fa}
.exh-up-row .datebox{width:48px;height:48px;border-radius:10px;background:color-mix(in srgb,var(--accent) 12%,#fff);border:1px solid color-mix(in srgb,var(--accent) 30%,transparent);display:flex;flex-direction:column;align-items:center;justify-content:center;flex:none}
.exh-up-row .datebox .mo{font-size:10px;letter-spacing:.5px;font-weight:800;color:var(--accent);text-transform:uppercase}
.exh-up-row .datebox .yr{font-size:9.5px;color:var(--accent);opacity:.8;font-family:var(--mono)}
.exh-card .datebox{width:54px;height:54px;border-radius:11px;background:color-mix(in srgb,var(--accent) 12%,#fff);border:1px solid color-mix(in srgb,var(--accent) 30%,transparent);display:flex;flex-direction:column;align-items:center;justify-content:center;flex:none}
.exh-card .datebox .mo{font-size:12px;letter-spacing:.5px;font-weight:800;color:var(--accent);text-transform:uppercase;line-height:1.1}
.exh-card .datebox .yr{font-size:10.5px;color:var(--accent);opacity:.82;font-family:var(--mono);margin-top:1px}
.exh-up-row .body{flex:1;min-width:0}
.exh-up-row .nm{font-size:13px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.exh-up-row .mt{font-size:11px;color:#8791a3;margin-top:2px}
.exh-up-empty{padding:30px 10px;text-align:center;color:#8791a3;font-size:12.5px}
#view-exhdetail{display:none;animation:fade .3s ease both;height:calc(100vh - var(--head));flex-direction:column;padding:16px 24px 20px}
#view-exhdetail.on{display:flex}

@media(max-width:1000px){
  :root{--sid:0px}
  #sidebar{transform:translateX(-100%);width:250px}
  #sidebar.open{transform:none;box-shadow:4px 0 24px rgba(20,30,50,.15)}
  .menu-btn{display:block}
  #main{margin-left:0}
  .hsearch{display:none}
  .charts{grid-template-columns:1fr}
  .hub-grid{grid-template-columns:1fr}
  .mkt-grid{grid-template-columns:1fr 1fr}
  .region-layout{grid-template-columns:1fr}
  #view-detail{height:auto}
  .det-grid{grid-template-columns:1fr;grid-template-rows:none}
  .det-tags{max-width:none}
  .contact-grid{grid-template-columns:1fr}
}
@media(max-width:600px){
  .hts{display:none}
  .kpi-grid{grid-template-columns:1fr 1fr}
  .mkt-grid{grid-template-columns:1fr}
  .dfields,.dfields.three{grid-template-columns:1fr}
  .wl-title{font-size:32px}
  .hub-title{font-size:26px}
  .toolbar,.results,.company-list,#view-dash{padding-left:16px;padding-right:16px}
}
</style>
</head>
<body>

<!-- WELCOME -->
<div id="welcome">
  <div class="wl-glow"></div>
  <div class="wl-inner">
    <div class="wl-mark">
      <img src="__LOGO_MARK__" alt="Mother India Forming Pvt Ltd">
      <div class="wl-eyebrow">MOTHER INDIA FORMING PVT LTD</div>
    </div>
    <div class="wl-title">Welcome to the <span class="brand">Market Intelligence</span> Portal</div>
    <div class="wl-sub">Where the steel-forming market comes into focus.</div>
    <div class="wl-desc">A live intelligence workspace mapping manufacturers, prospects and competitors across every segment, region and revenue band that matters to MIF.</div>
    <button class="wl-cta" onclick="enterLogin()">Step Inside the Portal <span class="arw">&rarr;</span></button>
  </div>
  <div class="wl-foot">&copy; By Ravi Prakash. Proprietary Software. Unauthorized copying, distribution, or modification is prohibited | 2026.</div>
</div>

<!-- LOGIN -->
<div id="login">
  <div class="login-glow"></div>
  <div class="login-box">
    <div class="login-brand">
      <div class="brandmark"><span class="glyph"></span><img src="__LOGO_MARK__" alt="MIF" style="display:none;position:absolute;inset:0;width:100%;height:100%;object-fit:contain" onload="this.style.display='block';this.previousElementSibling.style.display='none'"></div>
      <div>
        <div class="login-name">__BRAND__</div>
        <div class="login-tag">__BYLINE__</div>
      </div>
    </div>
    <div class="login-h">Secure Sign-In</div>
    <div class="login-sub">Mother India Forming &middot; Authorised access only</div>
    <div class="login-lbl">Username</div>
    <input type="text" id="usr" class="login-inp" placeholder="Enter username" autocomplete="username" autocapitalize="none" spellcheck="false">
    <div class="login-lbl">Password</div>
    <input type="password" id="pwd" class="login-inp" placeholder="Enter password" autocomplete="current-password">
    <button class="login-btn" onclick="tryLogin()">Sign In &rarr;</button>
    <div class="login-err" id="login-err"></div>
    <div style="text-align:center"><button class="login-back" onclick="backWelcome()">&larr; Back</button></div>
    <div class="login-copy">&copy; By Ravi Prakash. Proprietary Software. Unauthorized copying, distribution, or modification is prohibited | 2026.</div>
  </div>
</div>

<!-- APP -->
<div id="app">
  <div id="header">
    <button class="menu-btn" onclick="toggleSidebar()">&#9776;</button>
    <div class="hbrand" onclick="goSearch()">
      <div class="brandmark"><span class="glyph"></span><img src="__LOGO_MARK__" alt="MIF" style="display:none;position:absolute;inset:0;width:100%;height:100%;object-fit:contain" onload="this.style.display='block';this.previousElementSibling.style.display='none'"></div>
      <div>
        <div class="hname">__BRAND__</div>
        <div class="htag">__BYLINE__</div>
      </div>
      <span class="hchip" id="mode-chip">MARKET</span>
    </div>
    <div class="hsearch">
      <span class="ic">&#8981;</span>
      <input type="search" id="search-box" placeholder="Quick search &mdash; company, CIN, tier, product&hellip;" oninput="onQuickSearch()">
    </div>
    <div class="hright">
      <div class="hts"><div class="k">DATA AS OF</div><div class="v" id="hts"></div></div>
      <button class="hlogout" onclick="logout()">Logout</button>
    </div>
  </div>

  <div id="ticker"><div class="tick-track" id="tick-track"></div></div>

  <div id="sidebar">
    <div class="sid-sec">
      <div class="sid-lbl">NAVIGATION</div>
      <div class="sid-item active" data-nav="dashboard" onclick="goDash()"><span class="ico">&#9638;</span><span class="nm">Dashboard</span></div>
      <div class="sid-item" data-nav="search" onclick="goSearch()"><span class="ico">&#8981;</span><span class="nm">Search</span></div>
      <div class="sid-item" data-nav="exh" onclick="goExhibitions()"><span class="ico">&#8982;</span><span class="nm">Exhibitions</span></div>
      <div class="sid-item" data-nav="news" onclick="goNews()"><span class="ico">&#9636;</span><span class="nm">News Update</span></div>
    </div>
    <div class="sid-hint">
      <div class="h">Explore the market</div>
      <div class="p">Use <b>Search</b> to browse companies by region, by segment framework, or by name &amp; CIN.</div>
    </div>
  </div>

  <div id="main">

    <!-- DASHBOARD -->
    <div id="view-dash">
      <div class="dash-head">
        <div>
          <div class="eyebrow" id="dash-eyebrow">MARKET OVERVIEW</div>
          <div class="dash-title" id="dash-title">Market Research Dashboard</div>
          <div class="dash-sub" id="dash-sub"></div>
        </div>
        <div class="synced"><span class="live"></span>Dataset synced &middot; <span class="v" id="synced-ts"></span></div>
      </div>
      <div class="dash-sec-lbl"><span class="n">1</span>Market News &amp; Customer Intelligence</div>
      <div class="panel-card" id="news-dash-card" style="margin-bottom:4px"></div>

      <div class="dash-sec-lbl"><span class="n">2</span>Market &amp; Company Coverage</div>
      <div class="kpi-grid" id="kpi-grid"></div>
      <div class="charts">
        <div class="panel-card">
          <div class="pc-head"><div class="pc-title">Coverage by Segment</div><div class="pc-meta" id="seg-meta"></div></div>
          <div class="seg-cov" id="seg-cov"></div>
        </div>
        <div class="panel-card">
          <div class="pc-head"><div class="pc-title">Regional Spread</div><div class="pc-meta" id="reg-meta"></div></div>
          <div class="pie-wrap"><div class="pie-svg" id="reg-pie"></div><div class="pie-legend" id="reg-legend"></div></div>
        </div>
      </div>

      <div class="dash-sec-lbl"><span class="n">3</span>Exhibitions</div>
      <div class="panel-card" id="exh-upcoming-panel" style="margin-bottom:16px">
        <div class="pc-head"><div class="pc-title">Upcoming Exhibitions</div><div class="pc-meta" id="exh-up-meta"></div></div>
        <div class="exh-up-list" id="exh-up-list"></div>
      </div>
    </div>

    <!-- EXHIBITIONS -->
    <div id="view-exh" class="page" data-screen-label="Exhibitions">
      <div class="page-wrap">
        <div class="page-eyebrow">GLOBAL EXHIBITIONS</div>
        <div class="page-title">Exhibition Intelligence</div>
        <div class="page-sub">Every trade show, expo and conference MIF is tracking worldwide — mapped, prioritised and ready to plan around.</div>

        <div class="exh-tools">
          <div class="exh-search"><span class="ic">&#8981;</span><input type="search" id="exh-search-input" placeholder="Search exhibitions, segment, country…" oninput="applyExhFilters()" autocomplete="off"></div>
          <div class="exh-chips" id="exh-chips"></div>
          <select class="exh-seg-sel" id="exh-seg-sel" onchange="applyExhFilters()"><option value="">All Segments</option></select>
        </div>
        <div class="exh-regions" id="exh-regions"></div>

        <div class="panel-card" style="margin-bottom:20px">
          <div class="pc-head"><div class="pc-title">Map View</div><div class="pc-meta" id="exh-map-meta"></div></div>
          <div class="exh-map-frame">
            <div class="exh-map-inner">
              <svg class="exh-world" id="exh-world" viewBox="0 0 360 180" preserveAspectRatio="none"></svg>
              <div class="exh-pins" id="exh-pins"></div>
            </div>
          </div>
          <div class="exh-map-legend" id="exh-map-legend"></div>
        </div>

        <div class="panel-card">
          <div class="pc-head"><div class="pc-title">List of Exhibitions</div><div class="pc-meta" id="exh-list-meta"></div></div>
          <div class="exh-list" id="exh-list"></div>
        </div>
      </div>
    </div>

    <!-- SEARCH HUB -->
    <div id="view-search" class="page">
      <div class="hub-wrap">
        <div class="hub-mark"><div class="brandmark"><span class="glyph" style="width:56px;height:56px"></span><img src="__LOGO_MARK__" alt="MIF" style="display:none;position:absolute;inset:0;width:100%;height:100%;object-fit:contain" onload="this.style.display='block';this.previousElementSibling.style.display='none'"></div></div>
        <div class="hub-eyebrow" id="hub-eyebrow">MARKET INTELLIGENCE</div>
        <div class="hub-title" id="hub-title">Search Prospects by</div>
        <div class="hub-sub">Pick a lens to explore the database &mdash; jump in by geography, by our segment framework, or search a company directly.</div>
        <div class="hub-grid">
          <button class="hub-btn c-region" onclick="goRegion()">
            <div class="halo"></div>
            <div class="cardart" id="region-art"></div>
            <div class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M12 21s-6.5-5.6-6.5-10.2A6.5 6.5 0 0 1 12 4.3a6.5 6.5 0 0 1 6.5 6.5C18.5 15.4 12 21 12 21Z"/><circle cx="12" cy="10.6" r="2.4"/></svg></div>
            <div class="body">
              <div class="t">Explore Regions</div>
              <div class="d">Interactive map across the five zones of India &mdash; open any region&rsquo;s companies.</div>
              <div class="go">Open region map <span class="arw">&rarr;</span></div>
            </div>
          </button>
          <button class="hub-btn c-segments" onclick="goSegments()">
            <div class="halo"></div>
            <div class="cardart"><svg viewBox="0 0 120 120" fill="none"><rect x="12" y="14" width="42" height="26" rx="6" stroke="#fff" stroke-width="2.5"/><rect x="66" y="14" width="42" height="26" rx="6" stroke="#fff" stroke-width="2.5"/><rect x="12" y="50" width="42" height="26" rx="6" fill="#fff"/><rect x="66" y="50" width="42" height="26" rx="6" stroke="#fff" stroke-width="2.5"/><rect x="12" y="86" width="42" height="22" rx="6" stroke="#fff" stroke-width="2.5"/><rect x="66" y="86" width="42" height="22" rx="6" stroke="#fff" stroke-width="2.5"/></svg></div>
            <div class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3.5" y="3.5" width="7" height="7" rx="1.5"/><rect x="13.5" y="3.5" width="7" height="7" rx="1.5"/><rect x="3.5" y="13.5" width="7" height="7" rx="1.5"/><rect x="13.5" y="13.5" width="7" height="7" rx="1.5"/></svg></div>
            <div class="body">
              <div class="t">Browse Segments</div>
              <div class="d">The full classification framework &mdash; markets, segments and sub-segments.</div>
              <div class="go">Explore segments <span class="arw">&rarr;</span></div>
            </div>
          </button>
          <button class="hub-btn c-company" onclick="goCompSearch()">
            <div class="halo"></div>
            <div class="cardart"><svg viewBox="0 0 120 120" fill="none" stroke="#fff" stroke-width="2.5"><rect x="14" y="20" width="92" height="13" rx="4"/><rect x="14" y="42" width="92" height="13" rx="4"/><rect x="14" y="64" width="54" height="13" rx="4"/><circle cx="80" cy="82" r="17"/><line x1="92" y1="94" x2="106" y2="108" stroke-width="4" stroke-linecap="round"/></svg></div>
            <div class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="10.5" cy="10.5" r="6.5"/><path d="m20 20-4.6-4.6"/></svg></div>
            <div class="body">
              <div class="t">Search Companies</div>
              <div class="d">Master search across name, CIN, customer, tier and every indexed field.</div>
              <div class="go">Start searching <span class="arw">&rarr;</span></div>
            </div>
          </button>
        </div>
        <div class="hub-exh-sec">
          <div class="hub-exh-eyebrow">GLOBAL INTELLIGENCE</div>
          <button class="hub-feat feat-exh" onclick="goExhibitions()">
            <div class="feat-art"><svg viewBox="0 0 120 120" fill="none" stroke="currentColor" stroke-width="3"><circle cx="60" cy="60" r="46"/><ellipse cx="60" cy="60" rx="20" ry="46"/><line x1="14" y1="60" x2="106" y2="60"/><line x1="24" y1="36" x2="96" y2="36"/><line x1="24" y1="84" x2="96" y2="84"/></svg></div>
            <div class="feat-med"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><circle cx="12" cy="12" r="8.5"/><path d="M3.5 12h17M12 3.5c2.6 2.6 2.6 14.4 0 17M12 3.5c-2.6 2.6-2.6 14.4 0 17"/></svg></div>
            <div class="feat-txt"><div class="feat-t">Global Exhibitions</div><div class="feat-d">Every trade show, expo and conference MIF is tracking worldwide &mdash; mapped, prioritised and ready to plan around.</div></div>
            <div class="feat-cta">Open exhibitions <span class="arw">&rarr;</span></div>
          </button>
        </div>
        <div class="hub-exh-sec">
          <div class="hub-exh-eyebrow" style="color:var(--accent)">MARKET INTELLIGENCE</div>
          <button class="hub-feat feat-news" onclick="goNews()">
            <div class="feat-art"><svg viewBox="0 0 120 120" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="8,84 34,58 56,68 82,34 112,18"/><polyline points="92,18 112,18 112,38"/></svg></div>
            <div class="feat-med"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 5h13v14H5a1 1 0 0 1-1-1V5Z"/><path d="M17 8h3v9a2 2 0 0 1-2 2"/><path d="M7 9h7M7 12.5h7M7 16h4"/></svg></div>
            <div class="feat-txt"><div class="feat-t">News Update</div><div class="feat-d">Daily &amp; bi-weekly market briefings &mdash; steel &amp; scrap prices, currency, supply-chain status and customer health flags.</div></div>
            <div class="feat-cta">Open news <span class="arw">&rarr;</span></div>
          </button>
        </div>
      </div>
    </div>

    <!-- NEWS -->
    <div id="view-news" class="page" data-screen-label="News">
      <div class="page-wrap">
        <div class="page-eyebrow" style="color:var(--accent)">MARKET INTELLIGENCE</div>
        <div class="page-title">News Update</div>
        <div class="page-sub">Daily &amp; bi-weekly market briefings &mdash; steel &amp; scrap pricing, currency, supply-chain status, customer health flags and roll-forming opportunities. Newest on top.</div>
        <div id="news-body">
          <div class="news-ind-strip" id="news-indicators"></div>
          <div id="news-trend"></div>
          <div id="news-feed"></div>
          <div id="news-empty" style="display:none;text-align:center;padding:60px 20px;color:#8791a3">No briefings loaded yet.</div>
        </div>
      </div>
    </div>

    <!-- REGION -->
    <div id="view-region" class="page">
      <div class="page-wrap">
        <div class="crumb"><span class="lk" onclick="goSearch()">Search</span><span class="sep">/</span><span>Region</span></div>
        <button class="page-back" onclick="goSearch()">&larr; Back to Search</button>
        <div class="page-eyebrow">GEOGRAPHIC LENS</div>
        <div class="page-title">Search by Region</div>
        <div class="page-sub">Select a zone on the map of India to open every company operating there. Companies with a pan-India footprint appear under each region.</div>
        <div class="region-layout">
          <div class="map-frame">
            <div class="map-inner" id="india-map"></div>
              <div class="map-legend" id="map-legend"></div>
          </div>
          <div class="region-side" id="region-side"></div>
        </div>
      </div>
    </div>

    <!-- SEGMENTS -->
    <div id="view-segments" class="page">
      <div class="page-wrap">
        <div class="crumb"><span class="lk" onclick="goSearch()">Search</span><span class="sep">/</span><span>Segments</span></div>
        <button class="page-back" onclick="goSearch()">&larr; Back to Search</button>
        <div class="page-eyebrow">CLASSIFICATION FRAMEWORK</div>
        <div class="page-title">Search by Segment</div>
        <div class="page-sub">The MIF market taxonomy &mdash; six markets, each broken into the segments and sub-segments we track. Click any segment to open its companies.</div>
        <div class="mkt-grid" id="mkt-grid"></div>
      </div>
    </div>

    <!-- COMPANY SEARCH -->
    <div id="view-compsearch" class="page">
      <div class="page-wrap">
        <div class="crumb"><span class="lk" onclick="goSearch()">Search</span><span class="sep">/</span><span>Company</span></div>
        <button class="page-back" onclick="goSearch()" style="display:block;margin:0 auto 16px">&larr; Back to Search</button>
        <div class="cs-wrap">
          <div class="page-eyebrow" style="text-align:center">MASTER SEARCH</div>
          <div class="page-title" style="text-align:center">Find a Company</div>
          <div class="page-sub" style="text-align:center;margin:8px auto 24px">Search across company name, CIN, customer, tier, products and every indexed field.</div>
          <div class="cs-bar">
            <span class="ic">&#8981;</span>
            <input type="search" id="cs-input" placeholder="Search company, CIN, customer name, tier, product&hellip;" oninput="onCompSearch()" autocomplete="off">
          </div>
          <div class="cs-hints" id="cs-hints"></div>
          <div class="cs-res" id="cs-res"></div>
        </div>
      </div>
    </div>

    <!-- LIST -->
    <div id="view-list">
      <div class="toolbar">
        <div class="tb-title"><button class="tb-back" onclick="listBack()" title="Back">&larr;</button><span id="view-label">All Companies</span></div>
        <div class="fgroup">
          <select class="sel" id="f-seg" onchange="applyFilters()"><option value="">All Segments</option></select>
          <select class="sel" id="f-region" onchange="applyFilters()"><option value="">All Regions</option></select>
          <select class="sel" id="f-rev" onchange="applyFilters()"><option value="">All Revenue</option></select>
          <select class="sel" id="f-tier" onchange="applyFilters()"><option value="">All Tiers</option></select>
          <select class="sel" id="f-pv" onchange="applyFilters()"><option value="">All Volume (PV)</option><option value="H">High Volume</option><option value="M">Medium Volume</option><option value="L">Low Volume</option></select>
          <div class="sortwrap"><span class="lb">Sort</span>
            <select class="sel sort" id="f-sort" onchange="applyFilters()">
              <option value="az">A &ndash; Z</option>
              <option value="rev">Revenue (high&rarr;low)</option>
              <option value="tier">Tier</option>
            </select>
          </div>
        </div>
      </div>
      <div class="results">
        <div class="c" id="result-count"></div>
        <div class="clr" id="clear-filters" onclick="clearFilters()" style="display:none">Clear filters</div>
      </div>
      <div class="company-list" id="company-list"></div>
      <div class="pager" id="pager"></div>
    </div>

    <!-- DETAIL -->
    <div id="view-detail">
      <div class="det-top">
        <button class="det-back" onclick="closeDetail()" title="Back">&larr;</button>
        <div class="det-logo" id="det-logo" style="display:none"><img id="det-logo-img" src="" alt=""></div>
        <div class="det-head-txt">
          <div class="det-eyebrow" id="det-eyebrow">COMPANY DOSSIER</div>
          <div class="det-name" id="det-name"></div>
        </div>
        <div class="det-tags" id="det-tags"></div>
      </div>
      <div class="det-grid">
        <div class="dcard">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--steel)"></span><span class="dcard-title">Company Information</span><span class="dcard-badge">Identity &middot; Operations</span></div>
          <div class="dcard-body" id="dc-info"></div>
        </div>
        <div class="dcard">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--green)"></span><span class="dcard-title">Financial Profile</span><span class="dcard-badge">Revenue &middot; Scale</span></div>
          <div class="dcard-body" id="dc-fin"></div>
        </div>
        <div class="dcard wide">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--accent)"></span><span class="dcard-title">Contact Information</span><span class="dcard-badge">Key People</span></div>
          <div class="dcard-body" id="dc-contact"></div>
        </div>
        <div class="dcard wide">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--purple)"></span><span class="dcard-title">Company Intelligence</span><span class="dcard-badge">Products &middot; MIF Fit &middot; Data</span></div>
          <div class="dcard-body" id="dc-intel"></div>
        </div>
      </div>
    </div>
    <!-- EXHIBITION DETAIL -->
    <div id="view-exhdetail" data-screen-label="Exhibition Detail">
      <div class="det-top">
        <button class="det-back" onclick="closeExhDetail()" title="Back">&larr;</button>
        <div class="det-head-txt">
          <div class="det-eyebrow">EXHIBITION DOSSIER</div>
          <div class="det-name" id="exdet-name"></div>
        </div>
        <div class="det-tags" id="exdet-tags"></div>
      </div>
      <div class="det-grid">
        <div class="dcard">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--steel)"></span><span class="dcard-title">Exhibition Overview</span><span class="dcard-badge">Where &middot; When</span></div>
          <div class="dcard-body" id="exdc-overview"></div>
        </div>
        <div class="dcard">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--green)"></span><span class="dcard-title">MIF Priority &amp; Fit</span><span class="dcard-badge">Business Case</span></div>
          <div class="dcard-body" id="exdc-fit"></div>
        </div>
        <div class="dcard wide">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--accent)"></span><span class="dcard-title">Logistics &amp; Registration</span><span class="dcard-badge">Cost &middot; Visa &middot; Contact</span></div>
          <div class="dcard-body" id="exdc-logistics"></div>
        </div>
        <div class="dcard wide">
          <div class="dcard-head"><span class="dcard-bar" style="background:var(--purple)"></span><span class="dcard-title">Data Source</span><span class="dcard-badge">Verification</span></div>
          <div class="dcard-body" id="exdc-source"></div>
        </div>
      </div>
    </div>

    <div class="app-footer" id="app-footer">
      <div class="note-card">
        <div class="note-ic">i</div>
        <div class="note-body">
          <div class="note-h">About this intelligence</div>
          <div class="note-p">All information in this portal is compiled from MIF&rsquo;s own primary market research. For a deeper briefing on any company, segment or exhibition featured here, please reach out to <b>Ravi Prakash</b> &mdash; R&amp;D / Business Intelligence, MIF.</div>
        </div>
      </div>
      <div class="af-copy">&copy; By Ravi Prakash. Proprietary Software. Unauthorized copying, distribution, or modification is prohibited | 2026.</div>
    </div>

  </div>
</div>

<script>
// ── DATA ──────────────────────────────────────────────────────────
const DATASETS = {mif: __DATA_MIF__, comp: __DATA_COMP__};
const COLUMNS = __COLS__;
const TIMESTAMP = "__TS__";
const PH_MIF = "__PH_MIF__";
const PH_COMP = "__PH_COMP__";
const USER_MIF = "admin";
const USER_COMP = "competitor";

// ── EXHIBITIONS DATA ──────────────────────────────────────────────
const EXHIBITIONS = __DATA_EXH__;
const NEWS = __DATA_NEWS__;
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const PRI_CFG = {
  MUST_ATTEND:{label:'Must Attend', fg:'#0f8a56', bg:'#e6f6ee'},
  ATTEND:{label:'Attend', fg:'#2a63cf', bg:'#eaf1fd'},
  WATCH:{label:'Watch', fg:'#b5790a', bg:'#fdf3e0'},
  AVOID:{label:'Avoid', fg:'#c93350', bg:'#fdeaec'},
};
const CITY_COORDS = {
  "shanghai":[31.23,121.47],"dubai":[25.20,55.27],"new delhi":[28.61,77.21],"delhi":[28.61,77.21],
  "bengaluru":[12.97,77.59],"bangalore":[12.97,77.59],"hannover":[52.37,9.73],"mumbai":[19.08,72.88],
  "munich":[48.14,11.58],"frankfurt":[50.11,8.68],"paris":[48.86,2.35],"las vegas":[36.17,-115.14],
  "birmingham":[52.48,-1.90],"stuttgart":[48.78,9.18],"milan":[45.46,9.19],"brussels":[50.85,4.35],
  "chicago":[41.88,-87.63],"düsseldorf":[51.23,6.78],"dusseldorf":[51.23,6.78],"chennai":[13.08,80.27],
  "bangkok":[13.76,100.50],"nuremberg":[49.45,11.08],"istanbul":[41.01,28.98],"amsterdam":[52.37,4.90],
  "tokyo":[35.68,139.69],"berlin":[52.52,13.40],"hyderabad":[17.39,78.49],"hamburg":[53.55,9.99],
  "rotterdam":[51.92,4.48],"london":[51.51,-0.13],"cologne":[50.94,6.96],"greater noida":[28.47,77.50],
  "beijing":[39.90,116.41],"bologna":[44.49,11.34],"madrid":[40.42,-3.70],"pune":[18.52,73.86],
  "guangzhou":[23.13,113.26],"são paulo":[-23.55,-46.63],"sao paulo":[-23.55,-46.63],"gandhinagar":[23.22,72.65],
  "decatur":[39.84,-88.95],"geneva":[46.20,6.15],"los angeles":[34.05,-118.24],"detroit":[42.33,-83.05],
  "singapore":[1.35,103.82],"seoul":[37.57,126.98],"buenos aires":[-34.60,-58.38],"lyon":[45.76,4.84],
  "karlsruhe":[49.01,8.40],"verona":[45.44,10.99],"peterborough":[52.57,-0.24],"indianapolis":[39.77,-86.16],
  "gdansk":[54.35,18.65],"turin":[45.07,7.69],"athens":[37.98,23.73],
  "charlotte":[35.23,-80.84],"shenzhen":[22.54,114.06],"coventry":[52.41,-1.51],"orlando":[28.54,-81.38],
  "newark":[40.74,-74.17],"atlanta":[33.75,-84.39],"ahmedabad":[23.02,72.57],"surat":[21.17,72.83],
  "dhaka":[23.81,90.41],"houston":[29.76,-95.37],"dallas":[32.78,-96.80],"taipei":[25.03,121.56],
  "bilbao":[43.26,-2.93],"glasgow":[55.86,-4.25],"new orleans":[29.95,-90.07],"columbus":[39.96,-83.00],
  "bari":[41.12,16.87],"zwingenberg":[49.73,8.62],"monterrey":[25.67,-100.31],"cairo":[30.04,31.24],
  "johannesburg":[-26.20,28.05],"kolkata":[22.57,88.36],"lillestrom":[59.95,11.05],
  "riyadh":[24.71,46.68],"long beach":[33.77,-118.19],"abu dhabi":[24.45,54.38],"mexico city":[19.43,-99.13],
  "sharjah":[25.35,55.42],"ranchi":[23.34,85.31],"jamshedpur":[22.80,86.18],"vijayanagar":[15.18,76.40],
  "bellary":[15.14,76.92],"goa":[15.30,74.12],"chandigarh":[30.73,76.78]
};
const COUNTRY_COORDS = {
  "Argentina":[-34.60,-58.38],"Bangladesh":[23.81,90.41],"Belgium":[50.85,4.35],"Brazil":[-23.55,-46.63],
  "China":[39.90,116.41],"Egypt":[30.04,31.24],"Europe":[50.0,10.0],"Finland":[60.17,24.94],
  "France":[48.86,2.35],"Germany":[50.94,6.96],"Greece":[37.98,23.73],"Hong Kong":[22.32,114.17],
  "India":[28.61,77.21],"Italy":[45.46,9.19],"Japan":[35.68,139.69],"Mexico":[19.43,-99.13],
  "Netherlands":[52.37,4.90],"Norway":[59.91,10.75],"Poland":[52.23,21.01],"Saudi Arabia":[24.71,46.68],
  "Singapore":[1.35,103.82],"South Africa":[-26.20,28.05],"South Korea":[37.57,126.98],"Spain":[40.42,-3.70],
  "Switzerland":[46.20,6.15],"Taiwan":[25.03,121.56],"Thailand":[13.76,100.50],"Turkey":[41.01,28.98],
  "UAE":[25.20,55.27],"UK":[51.51,-0.13],"USA":[39.83,-98.58]
};
let exhState = {priority:'', search:'', seg:'', continent:''};
const CONTINENTS=[{name:'North America',lat:44,lon:-100},{name:'South America',lat:-14,lon:-60},{name:'Europe',lat:56,lon:12},{name:'Africa',lat:3,lon:22},{name:'Asia',lat:47,lon:95},{name:'Oceania',lat:-25,lon:134}];
const CONTINENT_OF={'China':'Asia','Hong Kong':'Asia','Japan':'Asia','Singapore':'Asia','South Korea':'Asia','Taiwan':'Asia','Thailand':'Asia','Bangladesh':'Asia','UAE':'Asia','Saudi Arabia':'Asia','India':'Asia','Belgium':'Europe','Finland':'Europe','France':'Europe','Germany':'Europe','Greece':'Europe','Italy':'Europe','Netherlands':'Europe','Norway':'Europe','Poland':'Europe','Spain':'Europe','Switzerland':'Europe','UK':'Europe','Turkey':'Europe','Europe':'Europe','USA':'North America','Mexico':'North America','Argentina':'South America','Brazil':'South America','Egypt':'Africa','South Africa':'Africa'};
function continentOf(country){return CONTINENT_OF[country]||'Other';}
const WORLD_LAND=[[[-168,66],[-150,71],[-125,70],[-100,72],[-82,73],[-62,66],[-56,53],[-60,47],[-67,45],[-70,41],[-75,35],[-81,25],[-90,29],[-97,26],[-105,22],[-108,25],[-114,31],[-120,34],[-124,40],[-124,48],[-130,54],[-140,60],[-152,58],[-165,60]],[[-45,60],[-22,70],[-20,76],[-35,83],[-55,82],[-58,72],[-50,64]],[[-78,8],[-72,11],[-62,10],[-50,0],[-44,-2],[-35,-6],[-39,-14],[-48,-25],[-55,-34],[-62,-40],[-66,-45],[-71,-52],[-75,-50],[-73,-42],[-71,-33],[-70,-18],[-76,-14],[-81,-6],[-80,0]],[[-16,15],[-16,21],[-10,30],[-2,35],[10,37],[22,32],[32,31],[34,28],[36,22],[43,12],[51,12],[44,2],[42,-4],[40,-15],[35,-24],[28,-33],[20,-35],[15,-28],[12,-16],[9,-1],[5,4],[-8,4]],[[44,-16],[50,-15],[50,-25],[45,-25]],[[-10,36],[-9,44],[-2,43],[-2,48],[2,51],[-5,58],[2,60],[8,58],[12,55],[10,64],[24,71],[30,70],[40,66],[55,62],[60,55],[50,46],[42,44],[40,47],[30,46],[28,41],[26,38],[22,40],[15,44],[12,46],[18,42],[15,40],[8,44],[3,43]],[[40,47],[50,52],[60,55],[75,55],[90,55],[105,55],[120,53],[132,50],[130,44],[122,40],[122,34],[120,30],[118,24],[110,21],[108,15],[106,10],[104,1],[100,7],[98,10],[92,21],[89,22],[87,20],[82,10],[78,8],[74,15],[71,20],[68,24],[62,25],[58,24],[57,25],[52,20],[48,24],[45,29],[42,30],[40,36],[38,40]],[[35,30],[43,30],[48,30],[57,25],[60,22],[56,17],[52,15],[45,12],[43,17],[40,22],[35,28]],[[130,31],[136,34],[141,40],[141,45],[138,42],[133,34]],[[-6,50],[-3,54],[-5,58],[-8,56],[-6,51]],[[95,5],[105,6],[119,5],[120,0],[118,-4],[105,-6],[98,-3],[95,1]],[[120,22],[122,25],[122,22],[121,20]],[[114,-22],[122,-18],[130,-12],[137,-11],[143,-12],[147,-20],[153,-28],[150,-37],[141,-38],[131,-32],[123,-34],[115,-34],[113,-26]],[[173,-41],[175,-37],[178,-38],[174,-46],[168,-46],[170,-41]]];
function renderWorldBase(){
  const svg=document.getElementById('exh-world');if(!svg)return;
  const toPath=poly=>'M'+poly.map(p=>((p[0]+180).toFixed(1)+' '+(90-p[1]).toFixed(1))).join('L')+'Z';
  svg.innerHTML=WORLD_LAND.map(poly=>'<path d="'+toPath(poly)+'"/>').join('');
}

let MODE = 'mif';
let DATA = [];
let state = {ctx:'all', label:'All Companies', search:'', seg:'', region:'', rev:'', tier:'', pv:'', sort:'az', page:1, perPage:50, filtered:[], lastView:'search'};

const ZONES = ['North','South','East','West','Central'];
const ZONE_MAP_FILL={North:'#6b9bea',West:'#eb8f4d',Central:'#e0be4a',East:'#9b86ee',South:'#3fb583'};
const MAP={"vb":[620,692],"cen":{"North":[181,168],"West":[168,380],"Central":[240,320],"East":[432,286],"South":[212,498]},"paths":{"North":"M211.1 182.4L217.7 183.8L220.8 182.8L220.7 185.1L222.6 185.6L223.5 187.8L222.8 188.7L218.2 189.3L218.9 193.7L218.3 196.6L216.0 197.1L214.0 195.1ZM156.1 186.0L157.9 193.2L159.2 193.4L160.2 195.3L156.7 199.5L154.2 199.4L151.3 196.5L150.8 198.1L149.1 198.4L149.2 200.4L147.0 201.3L146.6 202.5L143.7 202.8L140.5 209.6L142.2 211.7L140.6 212.4L140.1 214.5L138.2 215.1L132.1 214.4L129.0 218.6L129.2 220.3L125.5 217.8L123.0 218.8L122.1 212.2L124.7 211.4L124.8 213.2L126.3 213.3L127.4 211.8L129.7 211.3L129.9 207.1L132.5 202.4L130.1 201.5L129.1 199.6L126.2 199.6L125.5 198.6L126.0 195.1L127.2 193.8L126.6 191.8L129.0 191.0L128.3 189.4L133.8 186.9L135.0 188.5L140.8 187.7L141.9 185.9L150.7 187.4L151.2 185.9L153.1 184.7L153.8 186.3ZM212.1 187.1L214.0 192.8L208.8 190.8L207.5 192.7L206.5 192.2L204.2 194.4L202.9 194.4L200.0 191.5L210.1 187.1ZM193.2 189.6L194.7 191.0L195.6 193.2L194.7 194.2L195.8 195.3L192.1 197.9L190.4 195.5L188.5 194.8L185.9 195.7L185.1 194.6L187.8 191.2L187.2 188.9L193.0 187.7ZM185.9 188.7L187.2 188.9L187.8 191.2L185.1 194.6L185.9 195.7L181.4 198.5L180.5 197.5L173.4 196.4L175.1 193.4L177.1 193.1L176.2 191.1L177.2 190.2L179.5 190.4L180.0 189.0L183.3 189.6ZM245.2 187.6L245.4 190.0L243.9 192.5L247.3 195.6L247.0 197.2L243.5 205.2L241.4 206.8L241.2 204.4L239.7 203.2L232.9 202.1L232.0 198.4L230.1 197.5L235.1 194.1L235.7 192.1L238.4 190.0L239.7 186.7ZM245.6 201.3L247.3 195.6L243.9 192.5L245.4 190.0L245.2 187.6L248.1 186.9L251.0 190.6L252.5 190.0L253.3 188.0L254.4 188.3L261.4 195.3L259.4 197.6L259.5 199.6L256.6 200.3L252.2 198.5L251.6 204.7L249.6 202.7L247.6 203.5ZM214.0 192.8L214.0 195.1L216.0 197.1L216.2 200.9L219.7 204.6L218.3 205.2L216.8 203.8L213.4 205.4L210.2 204.2L209.5 205.4L206.2 203.1L203.9 204.2L203.0 200.7L204.0 199.9L201.8 197.4L202.9 194.4L204.2 194.4L206.5 192.2L207.5 192.7L208.8 190.8ZM200.4 192.2L202.9 194.4L201.8 197.4L204.0 199.9L203.0 200.7L203.9 204.2L198.5 205.2L199.1 198.7L194.7 194.2L195.6 193.2ZM259.5 199.6L259.4 197.6L261.4 195.3L258.6 192.8L261.6 192.9L263.6 191.4L274.2 197.9L273.9 200.6L277.2 204.8L277.2 207.0L279.5 209.2L279.1 214.4L276.9 212.0L272.9 211.1L272.2 209.4L269.4 211.9L265.5 210.8L264.7 212.8L262.9 211.5L258.5 214.4L257.3 212.7L253.2 212.6L252.5 209.4L257.3 205.1ZM146.6 202.5L147.0 201.3L149.2 200.4L149.1 198.4L150.8 198.1L151.3 196.5L154.2 199.4L156.7 199.5L160.2 195.3L165.0 198.2L168.4 202.1L166.7 208.7L162.3 211.0L160.6 214.5L157.5 213.7L155.1 215.3L152.1 213.5L152.1 211.2L148.1 205.5L148.8 204.0ZM181.4 198.5L182.4 197.2L188.5 194.8L190.9 196.0L192.1 197.9L191.7 199.9L193.1 201.1L192.5 202.4L188.9 201.5L188.0 199.5L185.4 201.9L181.6 201.1ZM195.8 195.3L199.1 198.7L199.0 202.4L193.1 201.1L191.7 199.9L192.1 197.9ZM165.8 208.8L166.7 208.7L168.4 202.1L165.0 198.2L165.2 197.3L166.9 197.0L167.6 198.3L169.6 196.9L173.3 197.3L175.0 199.8L173.8 203.1L174.8 204.3L174.0 206.8L172.7 205.6L170.6 207.2L172.0 210.9L171.0 211.7L166.6 210.5ZM173.4 196.4L180.5 197.5L181.6 201.1L185.4 201.9L181.8 205.1L181.4 206.8L178.6 207.7L178.7 206.4L173.8 203.1L175.0 199.8ZM230.1 197.5L232.0 198.4L232.9 202.1L239.7 203.2L241.2 204.4L241.4 206.8L239.9 210.7L238.3 211.6L238.2 214.9L234.6 213.7L234.2 211.8L229.0 210.6L228.4 209.1L222.4 206.7L221.8 205.7L223.4 199.9L228.0 196.4ZM80.0 209.1L83.2 208.2L85.5 209.9L86.0 211.9L84.6 213.0L87.1 214.5L89.7 214.1L87.7 218.9L86.2 219.4L83.6 215.9L78.5 219.9L78.7 222.2L81.7 226.7L83.4 227.5L83.4 231.4L84.7 232.5L83.7 234.4L84.4 236.1L83.4 237.2L84.6 239.0L83.1 241.3L84.2 242.4L83.9 244.8L80.3 246.0L79.1 244.2L78.1 244.9L74.8 243.5L73.5 241.2L71.7 242.3L70.6 241.4L68.3 243.2L67.4 242.1L65.9 242.7L64.4 245.7L63.0 245.4L61.6 246.4L62.4 249.6L60.1 251.0L54.1 248.3L47.9 249.5L43.2 249.2L43.9 240.2L41.4 239.1L35.8 239.1L29.8 235.8L29.9 229.7L31.5 225.7L40.8 217.1L43.0 211.6L48.1 206.9L52.7 206.9L54.5 208.9L54.7 211.1L56.3 213.5L59.4 213.8L65.7 210.9L75.5 209.9L79.9 208.2ZM279.1 214.4L279.5 209.2L277.2 207.0L277.2 204.8L273.9 200.6L274.2 197.9L277.6 199.0L279.9 202.6L282.1 203.2L283.2 205.2L288.8 208.2L286.7 212.3L287.8 213.4L287.5 216.9L292.2 218.3L293.8 220.0L290.3 223.9L288.8 222.5L286.7 226.3L284.0 228.0L282.3 225.9L281.8 222.3L278.7 216.5ZM185.4 201.9L188.0 199.5L188.9 201.5L192.5 202.4L191.0 204.5L191.7 208.3L195.7 210.6L194.4 211.7L189.3 211.2L187.9 215.0L186.1 213.2L187.2 210.7L187.8 204.0ZM241.4 206.8L243.5 205.2L245.6 201.3L247.6 203.5L249.6 202.7L251.6 204.7L252.2 198.5L256.6 200.3L259.5 199.6L257.3 205.1L252.5 209.4L253.2 212.6L249.0 215.1L246.9 213.7L245.3 215.4L246.3 219.8L243.2 216.3L238.2 214.9L238.3 211.6L239.9 210.7ZM140.1 214.5L140.6 212.4L142.2 211.7L140.5 209.6L143.7 202.8L146.6 202.5L148.8 204.0L148.1 205.5L152.1 211.2L152.1 213.5L155.1 215.3L157.5 213.7L160.6 214.5L162.3 211.0L165.8 208.8L166.6 210.5L168.9 210.8L168.7 212.4L166.0 215.8L165.5 219.7L162.9 221.4L161.7 220.8L158.8 223.1L157.2 222.3L156.1 223.5L154.4 223.6L152.7 226.7L150.2 227.1L147.5 225.0L149.5 222.7L146.8 221.9L143.9 222.4L143.1 219.7L140.2 218.5L141.0 217.1ZM199.0 202.4L198.5 205.2L199.9 207.4L199.6 208.8L195.7 210.6L191.7 208.3L191.0 204.5L193.1 201.1ZM172.0 210.9L170.6 207.2L172.7 205.6L174.0 206.8L174.8 204.3L173.8 203.1L178.7 206.4L178.6 207.7L181.4 206.8L181.8 205.1L185.3 202.1L187.8 204.0L187.2 210.7L186.1 213.2L187.9 215.0L187.0 216.2L189.2 218.2L188.6 220.9L189.9 221.0L191.5 222.6L189.5 226.3L187.5 226.7L185.6 224.7L184.7 226.9L181.6 227.3L180.9 226.2L178.9 226.0L176.9 228.2L173.2 228.5L172.8 225.7L171.0 225.7L169.7 224.6L172.3 219.1L171.3 215.9L172.4 215.6ZM203.9 204.2L206.2 203.1L207.8 204.9L209.5 205.4L210.2 204.2L213.4 205.4L216.8 203.8L218.3 205.2L219.7 204.6L222.4 206.7L219.6 211.5L218.2 211.0L217.4 212.9L214.7 214.8L212.4 212.3L210.1 212.5L210.1 214.7L209.0 216.4L207.4 216.7L205.3 214.0L205.6 211.4L204.6 209.3L202.0 207.9L199.6 208.8L199.9 207.4L198.5 205.2ZM222.4 206.7L228.4 209.1L229.0 210.6L234.2 211.8L234.6 213.7L233.5 215.5L235.3 216.7L232.3 217.2L230.7 215.2L229.7 216.8L228.3 216.6L224.3 214.3L224.1 213.0L222.3 212.4L220.4 213.0L219.6 211.5ZM194.4 211.7L202.0 207.9L204.6 209.3L205.6 211.4L205.3 214.0L207.4 216.7L207.2 222.2L209.2 222.5L208.1 223.9L206.4 223.2L202.1 224.4L195.6 217.9L195.9 215.1ZM168.9 210.8L171.0 211.7L172.0 210.9L172.4 215.6L171.3 215.9L172.3 219.1L169.7 224.6L171.0 225.7L172.8 225.7L173.2 228.5L171.8 232.0L172.7 233.7L170.3 236.9L171.1 238.8L170.5 240.9L163.3 240.9L161.1 242.6L158.1 242.5L153.8 241.6L153.0 240.0L150.3 241.0L149.6 242.2L147.4 240.1L147.4 237.8L145.3 237.2L144.2 233.7L146.3 231.2L149.5 231.5L150.7 228.1L153.1 227.9L153.0 225.2L154.4 223.6L156.1 223.5L157.2 222.3L158.8 223.1L161.7 220.8L162.9 221.4L165.5 219.7L166.0 215.8L168.7 212.4ZM293.8 220.0L292.2 218.3L287.5 216.9L287.8 213.4L286.7 212.3L288.8 208.2L291.8 210.4L293.4 208.9L295.7 208.9L298.6 210.7L294.7 216.7L295.8 220.8ZM83.9 244.8L84.2 242.4L83.1 241.3L84.6 239.0L83.4 237.2L84.4 236.1L83.7 234.4L84.7 232.5L83.4 231.4L83.4 227.5L81.7 226.7L78.7 222.2L78.5 219.9L83.6 215.9L86.2 219.4L94.3 215.8L96.4 217.7L100.0 216.6L100.0 219.3L103.3 221.2L105.3 220.2L106.4 220.8L106.3 231.3L108.0 233.0L116.3 233.1L118.1 235.5L117.5 236.4L119.0 240.5L122.1 243.0L120.9 244.6L120.9 250.3L117.6 251.2L117.0 249.5L113.5 250.8L112.2 249.1L110.2 252.1L107.8 251.2L104.8 252.4L102.8 255.1L100.2 255.9L95.7 253.4L97.2 251.9L97.2 248.5L95.1 246.4L92.3 245.9L91.3 247.4L88.2 248.3ZM258.5 214.4L262.9 211.5L264.7 212.8L265.5 210.8L269.4 211.9L272.2 209.4L272.9 211.1L276.9 212.0L279.1 214.4L278.7 216.5L281.8 222.3L280.7 224.1L278.5 222.3L276.9 224.4L275.5 223.7L274.5 225.0L272.4 224.5L271.6 226.3L269.4 227.6L268.8 225.4L262.9 223.8L261.5 220.8L258.8 218.0ZM187.9 215.0L189.3 211.2L194.4 211.7L195.9 215.1L195.6 217.9L202.1 224.4L202.3 226.1L200.1 227.4L202.4 230.6L196.8 233.8L198.0 235.8L195.8 236.5L188.9 231.3L187.5 226.7L189.5 226.3L191.5 222.6L189.9 221.0L188.6 220.9L189.2 218.2L187.0 216.2ZM207.4 216.7L209.0 216.4L210.1 214.7L210.1 212.5L212.4 212.3L214.7 214.8L217.4 212.9L218.2 211.0L220.4 213.0L220.5 214.7L218.4 216.7L216.6 216.7L213.9 218.5L213.8 220.7L212.4 221.6L207.2 222.2ZM122.2 217.7L123.0 218.8L125.5 217.8L129.2 220.3L129.0 218.6L132.1 214.4L138.2 215.1L140.1 214.5L141.0 217.1L140.2 218.5L143.1 219.7L143.9 222.4L146.8 221.9L149.5 222.7L147.5 225.0L150.2 227.1L152.7 226.7L153.1 227.9L150.7 228.1L149.5 231.5L146.3 231.2L145.5 232.4L145.6 231.2L142.9 230.4L141.3 234.8L139.7 235.1L140.2 236.8L139.1 239.2L137.4 238.4L134.7 241.7L129.9 242.3L128.3 243.4L122.1 243.0L119.0 240.5L117.5 236.4L118.1 235.5L116.3 233.1L108.0 233.0L106.3 231.3L105.9 224.9L106.5 222.9L108.5 222.3L109.2 224.0L111.3 224.0L114.0 221.8L115.2 218.7L120.6 218.6ZM220.4 213.0L222.3 212.4L224.1 213.0L224.3 214.3L228.3 216.6L229.7 216.8L230.7 215.2L232.3 217.2L235.3 216.7L236.5 218.2L236.1 222.5L228.4 219.2L227.2 221.7L224.8 222.1L223.6 219.2L222.2 218.3L220.0 221.1L217.8 221.7L213.8 220.7L213.9 218.5L216.6 216.7L218.4 216.7L220.5 214.7ZM253.2 212.6L257.3 212.7L258.5 214.4L258.8 218.0L261.5 220.8L262.9 223.8L268.8 225.4L269.4 227.6L265.6 228.2L259.6 231.7L258.4 229.4L255.7 230.9L254.4 230.1L252.8 231.7L250.4 226.7L246.0 224.9L245.3 221.3L246.4 218.4L245.3 215.4L246.9 213.7L249.0 215.1ZM234.6 213.7L243.2 216.3L246.3 219.8L245.3 221.3L246.0 224.9L244.9 226.2L239.1 225.5L237.4 224.4L236.1 222.5L236.5 218.2L233.5 215.5ZM215.7 221.1L220.0 221.1L222.2 218.3L223.6 219.2L224.9 225.5L226.8 226.6L225.7 228.8L226.7 230.9L223.8 232.7L216.6 227.7L214.4 227.8L213.7 225.6ZM318.5 227.3L313.4 229.9L311.0 228.9L310.4 227.4L306.1 228.1L303.9 226.5L305.1 225.1L304.7 221.8L306.5 220.3L309.4 220.8L309.7 218.8L313.9 218.5L316.2 219.7L319.4 219.6L321.6 222.3L318.9 224.8ZM224.8 222.1L227.2 221.7L228.4 219.2L236.1 222.5L239.1 225.5L238.8 230.2L237.6 230.9L234.5 229.8L226.7 230.9L225.7 228.8L226.8 226.6L224.9 225.5ZM333.7 222.3L334.6 223.7L332.7 224.5L330.8 231.1L328.6 231.0L327.0 232.5L324.7 230.9L322.3 230.7L323.7 228.1L318.5 227.3L318.9 224.8L324.1 220.5L323.8 219.0L328.4 219.3ZM202.1 224.4L206.4 223.2L208.1 223.9L209.2 222.5L213.8 220.7L215.7 221.1L213.7 225.6L214.4 227.8L216.6 227.7L223.8 232.7L225.2 231.8L227.6 234.3L226.1 235.2L221.1 235.3L217.9 234.1L217.2 232.9L214.3 233.8L215.3 231.8L212.7 231.1L211.9 232.3L209.3 232.5L204.5 231.5L203.7 232.8L199.5 234.1L198.0 235.8L196.8 233.8L202.4 230.6L200.1 227.4L202.3 226.1ZM284.0 228.0L286.7 226.3L288.8 222.5L290.3 223.9L293.8 220.0L300.5 224.3L299.1 226.1L301.1 228.4L303.2 228.0L303.9 226.5L306.1 228.1L306.1 230.8L304.6 232.1L300.9 230.5L299.1 234.3L298.3 233.7L293.7 234.9L288.9 233.0L286.1 229.2ZM281.8 222.3L282.3 225.9L288.9 233.0L286.7 234.0L284.4 240.4L281.2 239.3L280.4 240.4L276.5 239.4L277.7 236.6L272.5 228.7L273.3 226.7L271.6 226.3L272.4 224.5L274.5 225.0L275.5 223.7L276.9 224.4L278.5 222.3L280.7 224.1ZM173.2 228.5L176.9 228.2L178.9 226.0L180.9 226.2L181.6 227.3L184.7 226.9L185.6 224.7L187.5 226.7L188.9 231.3L190.4 233.1L187.4 232.9L186.3 230.6L184.1 230.3L183.6 232.4L177.3 236.4L175.8 239.9L177.3 241.4L175.9 243.3L170.5 240.9L171.1 238.8L170.3 236.9L172.7 233.7L171.8 232.0ZM334.6 223.7L336.4 226.0L335.8 228.0L337.6 228.9L337.9 232.5L342.1 234.4L343.6 237.2L345.4 238.4L342.5 239.1L339.0 238.4L337.7 240.2L333.2 235.4L328.7 237.7L328.6 235.6L326.5 233.3L328.6 231.0L330.8 231.1L332.7 224.5ZM246.0 224.9L250.4 226.7L252.2 231.0L250.7 231.3L248.6 235.1L243.4 232.4L241.2 232.7L240.0 231.3L237.6 230.9L238.8 230.2L239.1 225.5L244.9 226.2ZM269.4 227.6L271.6 226.3L273.3 226.7L272.5 228.7L277.7 236.6L276.5 239.4L274.1 241.2L269.7 237.6L268.7 238.0L267.1 236.8L267.2 234.5L263.5 229.6L265.6 228.2ZM306.1 228.1L310.4 227.4L311.0 228.9L313.4 229.9L312.7 231.3L313.0 235.0L315.0 236.6L311.6 239.8L309.3 240.1L308.3 239.0L302.6 237.9L299.1 234.3L300.9 230.5L304.6 232.1L306.1 230.8ZM318.0 227.7L323.7 228.1L322.3 230.7L324.7 230.9L327.0 232.5L326.5 233.3L328.6 235.6L328.7 237.7L328.1 241.0L325.9 241.5L326.1 243.3L328.6 245.5L328.7 247.8L326.7 246.5L324.3 246.5L320.3 243.8L316.9 243.4L318.7 240.5L317.5 234.4L320.0 232.7L320.2 230.4ZM313.4 229.9L318.0 227.7L320.2 230.4L320.0 232.7L317.5 234.4L318.7 240.5L316.9 243.4L315.8 241.9L311.6 239.8L315.0 236.6L313.0 235.0L312.7 231.3ZM190.4 233.1L194.6 235.0L195.8 236.5L196.4 238.6L193.2 240.7L193.8 242.9L195.4 242.9L196.3 244.4L185.3 252.0L182.1 248.7L179.4 247.9L180.0 245.1L187.8 240.2L188.0 239.0L186.6 237.2L183.3 239.8L177.3 236.4L183.6 232.4L184.1 230.3L186.3 230.6L187.4 232.9ZM252.8 231.7L254.4 230.1L255.7 230.9L258.4 229.4L259.6 231.7L263.5 229.6L264.6 230.3L265.2 232.7L267.2 234.5L267.1 236.8L268.7 238.0L269.7 237.6L273.1 240.6L274.1 244.2L266.2 250.0L264.4 248.7L262.9 244.6L258.5 240.4L257.1 237.1L254.4 235.0L254.9 233.2ZM225.2 231.8L234.5 229.8L237.6 230.9L237.5 237.7L235.4 239.9L236.4 241.8L235.3 242.8L233.6 242.7L230.5 239.6L230.9 237.5L228.0 236.9L226.1 235.2L227.6 234.3ZM195.8 236.5L198.0 235.8L199.5 234.1L203.7 232.8L204.5 231.5L209.3 232.5L211.9 232.3L212.7 231.1L215.3 231.8L214.3 233.8L211.9 234.6L211.5 237.4L209.6 236.9L207.5 237.7L205.8 240.2L196.3 244.4L195.4 242.9L193.8 242.9L193.2 240.7L196.4 238.6ZM252.2 231.0L254.9 233.2L254.4 235.0L257.1 237.1L258.5 240.4L262.9 244.6L264.4 247.5L258.8 250.2L257.3 251.7L256.9 254.5L253.7 252.8L254.3 251.7L251.0 251.0L252.1 246.8L255.6 244.9L254.4 242.2L255.6 241.3L254.7 239.7L252.4 238.9L252.0 236.9L249.7 234.1L250.7 231.3ZM238.1 231.6L243.4 232.4L246.7 234.7L245.9 237.8L244.3 238.9L243.4 242.1L241.5 244.4L240.8 242.9L237.6 243.7L235.3 242.8L236.4 241.8L235.4 239.9L237.5 237.7ZM43.2 249.2L47.9 249.5L54.1 248.3L60.1 251.0L62.4 249.6L61.6 246.4L63.0 245.4L64.4 245.7L65.9 242.7L67.4 242.1L68.3 243.2L70.6 241.4L71.7 242.3L73.5 241.2L74.8 243.5L78.1 244.9L79.1 244.2L80.3 246.0L83.9 244.8L88.2 248.3L91.3 247.4L92.3 245.9L95.1 246.4L97.2 248.5L97.2 251.9L95.7 253.4L100.2 255.9L96.9 261.6L95.0 261.7L94.4 263.5L91.5 265.6L89.7 265.6L89.3 266.7L83.2 264.8L80.8 267.7L81.4 268.4L78.1 273.8L77.0 274.0L75.8 275.9L74.8 274.8L68.6 277.0L68.3 279.2L65.4 283.6L63.7 283.0L60.1 277.0L59.0 272.2L54.3 266.5L54.5 260.2L53.1 259.3L48.4 260.2L45.9 259.3L42.3 254.2L42.0 250.9ZM288.9 233.0L293.7 234.9L298.3 233.7L302.6 237.9L304.3 237.9L303.3 240.3L301.4 239.6L299.0 243.4L296.3 241.9L287.9 240.4L286.7 238.2L284.9 238.1L286.7 234.0ZM246.7 234.7L248.6 235.1L249.7 234.1L252.0 236.9L252.4 238.9L254.7 239.7L255.6 241.3L254.4 242.2L255.6 244.9L252.1 246.8L250.1 250.7L247.3 250.3L246.9 248.6L242.4 246.6L241.5 244.4L243.4 242.1L244.3 238.9L245.9 237.8ZM170.9 241.7L175.9 243.3L177.3 241.4L175.8 239.9L177.3 236.4L183.3 239.8L186.6 237.2L188.0 239.0L187.8 240.2L180.0 245.1L179.4 247.9L182.1 248.7L185.3 252.0L183.9 254.0L177.3 256.4L176.6 257.9L174.8 258.7L173.8 257.0L171.7 256.0L172.3 252.3L169.2 252.8L169.1 250.4L167.1 250.1L167.0 247.7L169.4 247.3ZM328.7 237.7L333.2 235.4L337.7 240.2L334.8 241.3L334.6 242.6L339.9 244.2L340.3 246.7L338.4 247.8L337.8 250.5L334.6 250.0L333.1 247.5L331.1 246.7L328.7 247.8L328.6 245.5L326.1 243.3L325.9 241.5L328.1 241.0ZM149.6 242.2L150.3 241.0L153.0 240.0L153.8 241.6L158.1 242.5L161.1 242.6L163.3 240.9L170.5 240.9L169.4 247.3L167.0 247.7L167.1 250.1L169.1 250.4L169.2 252.8L172.3 252.3L171.7 256.0L173.8 257.0L170.2 257.7L168.2 257.1L168.3 255.9L164.3 255.6L162.3 256.9L162.0 258.3L159.9 258.5L159.1 260.1L154.5 257.9L152.8 258.1L151.3 254.2L153.5 252.5L150.0 250.1L148.4 246.2ZM293.0 241.2L302.3 244.1L302.7 245.9L308.6 246.9L308.6 249.3L306.5 248.5L303.5 250.5L303.4 252.6L295.3 251.4L295.3 249.4L291.1 246.3L291.7 243.9L293.2 243.2ZM100.2 255.9L102.8 255.1L104.8 252.4L107.8 251.2L110.2 252.1L112.2 249.1L113.5 250.8L117.0 249.5L117.6 251.2L120.9 250.3L120.9 244.6L122.1 243.0L128.3 243.4L129.9 242.3L131.0 242.5L131.3 245.3L130.2 246.2L132.7 248.4L128.7 252.0L127.2 256.1L126.0 256.7L124.8 255.8L119.4 266.6L116.6 269.1L116.3 271.2L112.4 273.6L111.0 275.3L110.7 277.7L106.8 281.3L106.9 279.1L103.6 277.4L106.6 273.4L103.7 273.3L102.7 272.3L102.4 270.1L105.7 264.9L101.4 258.8L99.0 257.2ZM304.3 237.9L308.3 239.0L309.3 240.1L311.6 239.8L315.8 241.9L318.2 244.0L317.7 246.3L316.8 246.4L314.0 243.4L312.9 244.0L313.0 247.2L311.7 248.5L308.6 246.9L302.7 245.9L302.3 244.1L299.0 243.4L301.4 239.6L303.3 240.3ZM276.5 239.4L280.4 240.4L281.7 245.6L280.6 246.4L281.0 248.5L279.6 252.4L282.5 254.3L280.1 255.7L280.2 257.1L277.7 255.8L277.7 254.1L274.4 252.8L273.4 251.1L270.4 251.5L266.2 250.0L274.1 244.2L273.1 240.6L274.1 241.2ZM233.6 242.7L237.6 243.7L240.8 242.9L242.4 246.6L246.9 248.6L247.3 250.3L250.1 250.7L251.0 252.2L249.1 254.2L249.2 255.5L243.7 255.1L242.4 256.8L241.3 256.0L237.6 258.1L231.6 255.8L229.4 253.9L229.7 252.1L231.0 250.9L230.6 248.3L233.7 245.0ZM318.2 244.0L320.3 243.8L325.3 246.5L324.3 250.3L322.2 251.7L321.7 254.0L322.8 256.9L320.0 258.3L319.4 260.1L316.8 260.7L310.2 256.5L310.6 255.1L308.6 246.9L311.7 248.5L313.0 247.2L312.9 244.0L314.0 243.4L316.8 246.4L317.7 246.3ZM325.3 246.5L328.7 247.8L331.1 246.7L333.1 247.5L333.6 248.6L332.2 250.3L329.8 251.1L329.7 252.5L332.1 253.3L332.3 255.0L330.7 256.1L329.1 255.1L322.8 256.9L321.7 254.0L322.2 251.7L324.3 250.3ZM264.4 247.5L264.4 248.7L266.2 250.0L270.4 251.5L273.4 251.1L274.4 252.8L277.7 254.1L277.7 255.8L280.2 257.1L277.9 259.3L277.5 262.0L278.6 263.1L277.2 265.4L277.1 264.4L273.5 263.1L271.6 259.8L269.7 259.0L267.2 260.0L264.0 259.2L263.0 257.9L263.6 255.7L256.9 254.5L257.3 251.7L258.8 250.2ZM333.6 248.6L334.6 250.0L337.8 250.5L342.9 254.0L344.1 253.7L349.8 257.5L349.3 258.9L348.4 260.1L343.7 258.6L343.3 260.4L340.9 260.3L339.6 258.9L338.2 259.7L338.3 261.1L335.1 262.9L335.6 258.5L330.7 256.1L332.3 255.0L332.1 253.3L329.7 252.5L329.8 251.1L332.2 250.3ZM303.4 252.6L303.5 250.5L306.5 248.5L308.6 249.3L310.6 255.1L310.2 256.5L316.8 260.7L316.9 263.1L315.9 262.4L314.2 263.5L312.8 262.7L309.9 263.1L308.9 266.6L307.4 266.6L303.6 263.8L299.0 262.1L297.4 260.2L298.9 256.4L300.2 256.4L301.7 255.0L302.1 253.2ZM152.8 258.1L154.5 257.9L155.8 259.1L153.7 264.6L152.1 266.4L155.5 269.6L155.5 272.1L154.6 273.2L155.6 274.5L150.1 274.8L150.4 273.5L144.2 271.6L141.7 272.1L140.0 270.9L137.9 272.4L133.1 271.5L130.5 274.1L128.9 274.3L125.9 270.2L125.3 267.7L127.6 260.8L127.2 259.1L129.0 258.7L128.7 256.2L133.1 253.6L146.4 256.0L148.3 259.7L149.4 259.1L151.4 259.9ZM65.4 283.6L68.3 279.2L68.6 277.0L74.8 274.8L75.8 275.9L77.0 274.0L78.1 273.8L81.4 268.4L80.8 267.7L83.2 264.8L89.3 266.7L89.7 265.6L91.5 265.6L94.4 263.5L95.0 261.7L96.9 261.6L99.0 257.2L104.9 262.9L105.7 264.9L102.4 270.1L100.1 269.1L96.3 275.2L92.7 278.5L92.9 281.1L90.4 280.9L89.2 283.0L89.6 284.0L86.0 284.2L83.8 282.2L75.3 283.8L74.3 283.0L71.5 282.9L67.5 284.4ZM173.8 257.0L174.8 258.7L173.5 259.7L173.3 263.1L171.6 264.5L170.5 267.4L161.4 271.6L161.8 273.1L159.6 275.3L155.6 274.5L154.6 273.2L155.5 272.1L155.5 269.6L152.1 266.4L153.7 264.6L155.8 259.1L159.1 260.1L159.9 258.5L162.0 258.3L162.3 256.9L164.3 255.6L168.3 255.9L168.2 257.1L170.2 257.7ZM228.3 254.8L229.4 253.9L235.2 257.6L239.2 257.6L238.3 260.9L239.1 263.7L236.3 269.4L237.6 272.2L233.5 273.1L232.1 270.5L229.3 270.9L228.2 272.0L229.9 267.7L228.2 266.8L229.5 264.4L228.3 263.8L225.1 265.2L226.0 267.5L223.5 266.8L220.4 269.3L221.4 270.7L219.0 272.8L216.9 268.1L216.8 265.7L217.9 265.5L219.1 262.5L223.5 262.8L226.7 261.5L225.7 259.2ZM174.8 258.7L176.6 257.9L177.3 256.4L178.8 256.5L178.1 263.6L179.8 266.3L174.1 266.9L171.4 272.7L173.6 276.5L175.6 275.9L177.2 277.1L176.6 279.5L174.8 278.5L168.6 281.5L168.8 285.4L165.6 286.0L162.7 281.0L163.3 278.5L158.4 275.0L157.4 275.3L158.3 274.5L159.6 275.3L161.8 273.1L161.4 271.6L170.5 267.4L171.6 264.5L173.3 263.1L173.5 259.7ZM259.2 255.1L263.6 255.7L263.0 257.9L264.0 259.2L267.2 260.0L269.7 259.0L271.6 259.8L273.6 264.1L266.6 269.5L267.0 274.1L264.5 273.5L262.5 274.7L262.0 276.1L260.4 274.0L259.1 275.3L257.5 275.0L261.0 271.6L260.5 270.0L258.5 269.1L258.6 266.7L254.1 267.6L258.7 258.4L256.8 257.4ZM316.8 260.7L319.4 260.1L320.0 258.3L322.8 256.9L329.1 255.1L335.6 258.5L335.1 262.9L333.3 265.7L326.8 268.5L326.7 266.6L324.0 264.3L322.5 265.0L320.6 263.6L316.9 263.1ZM277.2 265.4L278.6 263.1L277.5 262.0L277.9 259.3L280.2 257.1L281.4 258.0L281.0 259.7L283.0 261.5L286.0 262.0L287.3 264.1L286.4 267.6L282.7 269.7L281.6 268.4L276.3 266.6ZM297.4 260.2L301.2 263.4L300.6 268.7L299.0 269.3L298.4 270.5L299.9 271.4L301.1 274.0L296.2 279.8L293.0 279.1L292.0 275.9L287.2 274.2L286.3 271.9L284.5 271.3L284.7 268.9L286.4 267.6L287.3 264.1L286.0 262.0L289.0 260.3L290.0 258.5L291.8 259.3L294.7 258.6L295.6 260.6ZM308.9 266.6L309.9 263.1L312.8 262.7L314.2 263.5L315.9 262.4L319.1 263.6L318.1 265.3L319.4 266.1L318.4 268.3L315.9 268.4L315.5 270.9L310.8 271.0L310.7 269.4L308.4 269.8ZM102.4 270.1L102.7 272.3L103.7 273.3L106.6 273.4L103.6 277.4L106.9 279.1L106.7 282.9L102.8 290.2L98.0 290.0L97.0 287.8L93.6 287.0L93.1 288.9L90.0 285.5L89.2 283.0L90.4 280.9L92.9 281.1L92.7 278.5L96.3 275.2L100.1 269.1ZM273.5 263.1L277.1 264.4L276.3 266.6L281.6 268.4L282.7 269.7L284.7 268.9L283.6 272.2L279.0 271.8L278.7 274.4L275.6 277.8L271.9 276.4L270.1 277.2L269.8 271.3L267.2 272.9L267.0 274.1L266.6 269.5L273.6 264.1ZM319.1 263.6L322.5 265.0L324.0 264.3L326.7 266.6L326.8 268.5L322.8 271.4L323.8 280.6L321.6 282.0L319.4 282.0L317.3 280.6L316.7 276.6L316.7 274.2L318.2 274.2L318.8 271.3L316.1 270.0L315.9 268.4L318.4 268.3L319.4 266.1L318.1 265.3ZM301.2 263.4L308.9 266.6L308.4 269.8L305.2 270.3L300.7 269.5L299.9 271.4L298.4 270.5L299.0 269.3L300.6 268.7ZM179.8 266.3L181.4 267.6L185.4 268.0L187.2 269.1L190.2 267.9L193.1 268.3L193.4 266.9L195.0 265.7L196.2 266.2L196.7 272.9L185.8 274.7L185.5 276.5L187.3 278.6L184.3 280.0L185.3 281.4L187.3 281.1L189.1 282.3L189.9 285.7L187.9 287.9L185.5 285.9L184.5 286.2L185.1 288.5L180.1 288.7L178.6 286.8L178.4 283.3L177.3 283.4L176.6 279.5L177.2 277.1L175.6 275.9L173.6 276.5L171.4 272.7L174.1 266.9ZM106.8 281.3L110.7 277.7L111.0 275.3L112.4 273.6L114.0 273.2L113.7 275.0L116.6 278.5L118.3 278.7L117.9 280.8L120.1 280.6L124.8 277.4L126.4 278.5L127.9 278.3L127.1 281.7L128.5 282.2L130.6 285.3L130.3 287.4L131.4 287.9L132.7 291.9L134.0 292.5L132.4 294.6L129.9 293.7L129.6 296.1L130.6 299.6L129.5 300.5L128.1 298.2L126.4 299.2L123.1 298.8L121.5 299.9L116.1 299.3L113.5 302.1L111.5 302.6L112.5 300.1L112.3 297.5L110.4 297.2L108.9 298.4L108.3 296.6L105.1 294.0L106.9 290.3L105.3 289.4L105.6 287.1L103.8 287.5L106.7 282.9ZM296.2 279.8L301.1 274.0L299.9 271.4L300.7 269.5L305.2 270.3L310.7 269.4L310.8 271.0L315.5 270.9L316.1 270.0L318.8 271.3L318.2 274.2L316.7 274.2L316.7 276.6L314.4 279.3L311.3 279.0L308.2 280.6L305.8 280.1L305.3 283.0L303.1 283.5L303.1 284.9L300.5 284.6L298.8 280.3ZM219.0 272.8L221.4 270.7L223.1 274.0L223.0 276.6L225.9 278.8L225.5 284.7L228.3 283.8L229.7 285.9L228.9 287.9L230.4 288.3L230.1 290.5L226.6 294.3L220.2 289.5L217.8 292.4L216.5 290.9L217.4 289.9L215.2 288.1L215.4 283.3L213.2 278.4L216.9 275.3L216.7 273.7ZM165.6 286.0L168.8 285.4L168.6 281.5L174.8 278.5L176.6 279.5L177.3 283.4L178.4 283.3L178.6 286.8L180.1 288.7L185.1 288.5L184.8 290.2L186.8 293.6L185.9 295.4L184.2 295.7L181.5 294.0L180.9 292.3L179.3 293.6L174.0 292.8L169.7 294.0L169.9 296.4L168.1 296.9L165.9 300.8L162.9 301.4L160.7 304.2L155.8 300.3L157.1 297.8L160.9 299.2L163.8 297.1L162.0 295.1L163.3 292.8L161.5 289.4L162.7 288.1L165.1 288.3ZM316.7 276.6L317.3 280.6L319.4 282.0L321.6 282.0L323.8 280.6L325.8 281.6L326.5 285.9L323.9 287.1L324.1 292.3L319.8 300.2L317.3 301.4L314.3 301.4L308.3 295.5L310.0 294.6L310.6 290.0L309.4 289.7L311.1 285.3L310.6 283.8L309.2 283.9L308.4 282.5L305.3 283.0L305.8 280.1L308.2 280.6L311.3 279.0L314.4 279.3ZM111.5 302.6L113.5 302.1L116.1 299.3L121.5 299.9L123.1 298.8L126.4 299.2L128.1 298.2L129.5 300.5L130.6 299.6L132.9 300.6L128.0 303.4L127.5 307.3L126.4 307.5L124.0 312.6L122.6 313.6L121.3 311.5L117.0 311.0L117.6 307.1L115.8 306.3L114.8 307.4ZM124.0 312.6L126.4 307.5L127.5 307.3L128.0 303.4L132.9 300.6L135.1 300.1L136.8 302.9L135.9 306.7L140.7 307.2L141.2 308.9L137.0 311.6L136.1 313.9L140.4 316.2L135.6 319.3L133.1 318.7L131.7 319.9L130.1 317.0L127.6 317.2L127.8 315.2ZM280.4 240.4L281.2 239.3L284.4 240.4L284.9 238.1L286.7 238.2L287.9 240.4L293.0 241.2L293.2 243.2L291.7 243.9L291.1 246.3L295.3 249.4L295.3 251.4L289.3 251.2L285.8 249.5L285.1 253.3L282.5 254.3L279.6 252.4L281.0 248.5L280.6 246.4L281.7 245.6ZM203.0 190.1L200.0 191.5L200.4 192.2L195.6 193.2L192.8 189.1L193.1 188.2L196.7 188.2L197.9 186.0L202.9 187.2ZM216.0 197.1L218.3 196.6L218.9 193.7L218.2 189.3L222.8 188.7L224.0 189.7L223.2 191.9L225.5 192.4L228.0 196.4L223.4 199.9L221.8 205.7L216.2 200.9ZM176.2 191.1L177.1 193.1L175.1 193.4L173.3 197.3L169.6 196.9L167.6 198.3L166.9 197.0L165.2 197.3L163.4 195.5L164.0 193.8L170.1 191.5L172.2 191.7L173.4 188.7L175.4 189.3ZM123.0 260.0L124.8 255.8L126.8 258.0L125.1 258.7L124.6 260.4ZM145.5 232.4L144.2 233.7L145.3 237.2L147.4 237.8L147.4 240.1L149.6 242.2L148.4 246.2L150.0 250.1L153.5 252.5L151.3 254.2L152.8 258.1L151.4 259.9L149.4 259.1L148.3 259.7L146.4 256.0L134.0 254.4L131.7 252.8L130.5 254.3L130.9 255.4L129.7 256.3L127.2 256.1L128.7 252.0L132.7 248.4L130.2 246.2L131.3 245.3L131.0 242.5L134.7 241.7L137.4 238.4L139.1 239.2L140.2 236.8L139.7 235.1L141.3 234.8L142.9 230.4L145.6 231.2ZM126.0 256.7L128.7 256.2L129.0 258.7L127.2 259.1L127.6 260.8L125.3 267.7L125.9 270.2L128.9 274.3L131.6 274.9L127.9 278.3L126.4 278.5L124.8 277.4L120.1 280.6L117.9 280.8L118.3 278.7L116.6 278.5L113.7 275.0L114.0 273.2L116.3 271.2L116.6 269.1L119.4 266.6L122.0 260.8L123.0 260.0L124.6 260.4L125.1 258.7L126.8 258.0ZM133.1 253.6L130.9 255.4L130.5 254.3L131.7 252.8ZM236.7 268.2L239.1 263.7L244.3 264.5L247.4 262.1L253.5 262.2L256.5 263.8L254.9 265.1L254.1 267.6L252.6 267.8L248.8 271.2L248.5 273.3L242.8 271.7L241.1 273.8L239.3 273.2L240.0 269.8L238.2 268.1ZM153.5 275.2L155.6 274.5L158.4 275.0L163.3 278.5L162.7 281.0L158.9 282.8L149.9 281.3ZM130.5 274.1L133.1 271.5L137.9 272.4L140.0 270.9L141.7 272.1L144.2 271.6L150.4 273.5L148.5 276.4L148.5 278.4L146.9 279.1L145.2 278.8L144.0 277.2L142.4 277.9L143.1 280.2L145.6 281.0L145.9 282.6L143.7 283.9L141.9 281.6L141.5 285.0L139.7 287.4L137.5 286.8L135.9 291.9L134.0 292.5L132.7 291.9L131.4 287.9L130.3 287.4L130.6 285.3L128.5 282.2L127.1 281.7L129.0 276.6L131.6 274.9ZM218.0 103.9L216.2 102.8L218.2 101.7L217.8 98.6L214.2 100.7L212.6 100.3L209.9 101.7L207.8 100.1L208.2 99.3L204.1 93.7L201.2 94.3L198.3 96.3L197.0 95.7L195.6 97.2L194.2 95.9L195.1 95.4L194.3 93.1L195.6 91.1L197.1 91.1L198.8 84.2L197.8 81.8L194.8 82.0L192.8 78.5L187.1 76.7L180.9 72.3L179.4 73.3L178.1 71.5L179.5 69.2L181.8 68.4L181.4 65.2L182.5 62.8L180.5 61.6L179.0 59.1L179.1 57.2L176.1 53.8L170.1 55.1L166.7 53.5L165.0 57.8L162.8 58.8L158.8 58.8L154.2 62.0L152.8 60.3L153.6 58.1L151.1 55.9L140.4 54.9L136.8 52.7L131.7 52.3L136.9 50.3L137.3 47.2L135.2 44.9L130.2 46.6L120.6 41.8L119.2 42.4L118.5 39.2L120.5 36.9L120.3 35.4L117.2 34.4L112.3 34.5L106.6 31.0L107.0 28.2L105.6 27.3L97.2 28.4L93.8 26.5L95.1 24.0L93.8 22.9L95.1 18.6L99.0 17.5L101.2 15.7L101.2 14.5L103.5 14.1L106.4 8.9L113.0 7.6L118.5 9.0L120.8 8.3L122.0 6.6L119.0 5.6L118.3 3.8L122.3 4.2L126.1 6.2L127.9 5.4L128.5 4.0L130.4 4.4L134.0 2.0L136.1 2.1L139.9 0.0L140.4 1.3L142.7 1.3L144.4 3.9L146.4 2.3L148.1 3.1L150.5 0.9L153.3 0.6L156.1 6.1L159.1 7.2L161.9 11.2L168.9 13.8L172.6 16.7L174.7 16.4L177.4 19.8L180.8 20.4L184.2 23.5L182.5 25.7L184.4 28.1L187.0 29.5L190.8 29.2L191.9 30.7L197.5 32.5L198.2 34.3L197.6 35.6L201.4 36.7L205.7 35.7L207.4 36.7L208.8 34.4L210.4 34.0L212.5 35.1L213.8 32.1L215.3 30.9L217.7 30.6L218.5 29.5L220.6 30.0L223.6 27.8L229.9 27.0L233.5 28.3L234.8 27.2L235.3 25.1L238.7 24.7L239.7 26.9L249.1 29.6L251.9 28.3L254.2 31.9L258.0 33.5L258.8 36.7L255.7 43.6L255.6 47.3L253.4 51.3L253.3 53.9L248.9 54.5L247.1 55.9L247.5 59.2L241.5 59.4L243.2 63.8L241.4 65.5L241.2 67.1L238.2 69.4L231.2 69.0L228.7 70.8L233.0 79.0L228.9 79.9L230.5 85.4L240.3 87.2L238.7 91.1L239.7 93.9L244.0 99.0L240.2 103.7L237.5 102.3L233.6 105.0L233.1 107.1L231.4 107.0L230.3 108.1L227.7 105.8L225.7 100.9L223.6 100.9ZM154.2 62.0L158.8 58.8L162.8 58.8L165.0 57.8L166.7 53.5L170.1 55.1L176.1 53.8L179.1 57.2L179.0 59.1L180.5 61.6L182.5 62.8L181.4 65.2L181.8 68.4L179.5 69.2L178.1 71.5L179.4 73.3L180.9 72.3L187.1 76.7L192.8 78.5L194.8 82.0L197.8 81.8L198.8 84.2L198.5 86.8L197.3 88.5L197.1 91.1L195.6 91.1L194.3 93.1L195.1 95.4L194.2 95.9L191.6 93.6L187.0 92.4L177.9 81.3L173.8 80.0L172.8 76.1L170.8 74.9L167.7 70.3L164.4 70.7L160.6 66.6L154.8 63.7ZM131.7 52.3L137.6 53.4L136.5 57.6L132.2 62.3L128.8 62.9L127.9 64.7L124.8 64.3L123.6 62.6L119.8 62.2L120.8 60.5L122.6 60.0L124.4 54.6L128.2 54.6ZM137.6 53.4L140.4 54.9L151.1 55.9L153.6 58.1L152.8 60.3L154.0 61.4L149.2 62.3L146.4 60.3L143.9 60.0L143.8 61.6L139.0 62.9L139.4 66.2L136.5 64.7L135.8 63.4L136.4 61.1L134.0 60.3L136.5 57.6ZM124.8 64.3L127.9 64.7L128.8 62.9L132.2 62.3L134.0 60.3L136.4 61.1L135.8 63.4L139.6 67.1L135.0 70.1L133.7 72.3L130.1 69.7L123.9 69.8L122.7 68.7L125.0 66.0ZM154.0 61.4L154.8 63.7L156.6 64.9L154.0 65.5L151.8 64.4L148.3 66.4L144.6 65.6L142.9 67.4L138.8 65.0L139.0 62.9L143.8 61.6L143.9 60.0L146.4 60.3L149.2 62.3ZM156.6 64.9L158.7 65.7L157.1 69.1L155.5 69.3L154.4 72.4L157.1 76.7L157.3 79.3L155.7 81.7L156.1 83.9L154.2 84.7L149.3 80.3L148.3 78.6L148.6 75.6L146.7 75.2L150.9 71.2L150.4 68.6L148.3 66.4L151.8 64.4L154.0 65.5ZM158.7 65.7L160.6 66.6L164.4 70.7L167.7 70.3L170.8 74.9L172.8 76.1L173.8 80.0L177.9 81.3L183.6 87.3L183.1 88.9L180.9 89.3L179.3 88.2L175.9 88.8L171.9 92.5L170.3 89.4L167.2 91.0L158.9 87.2L156.1 83.9L155.7 81.7L157.3 79.3L157.1 76.7L154.4 72.4L155.5 69.3L157.1 69.1ZM139.4 66.2L142.9 67.4L144.6 65.6L148.3 66.4L149.2 68.3L147.7 67.5L144.3 69.7L139.6 67.1ZM133.7 72.3L135.0 70.1L139.6 67.1L144.3 69.7L142.2 73.3L136.4 77.5L134.5 75.9ZM149.2 68.3L150.4 68.6L150.9 71.2L150.1 72.4L147.0 74.5L140.0 74.7L142.2 73.3L144.3 69.7L147.7 67.5ZM122.7 68.7L123.9 69.8L130.1 69.7L133.7 72.3L134.5 75.9L136.4 77.5L135.8 79.4L137.0 81.4L136.1 82.4L133.8 80.1L129.3 80.6L129.7 81.9L128.4 83.6L128.0 80.7L123.9 78.8L121.7 79.5L118.6 78.8L115.6 75.3L116.2 72.9L114.0 66.9L117.5 68.8ZM147.0 74.5L145.0 76.2L145.3 77.5L143.3 78.8L139.3 77.0L136.9 77.2L140.0 74.7ZM146.8 75.2L148.6 75.6L148.3 78.6L149.3 80.3L146.6 81.7L141.7 82.4L139.9 81.1L138.1 81.7L135.8 79.4L136.4 77.5L139.3 77.0L143.3 78.8L145.3 77.5L145.0 76.2ZM128.4 83.6L129.7 81.9L129.3 80.6L133.8 80.1L136.1 82.4L138.1 81.7L139.5 87.5L137.0 92.4L132.0 94.1L131.7 92.4L128.4 91.3L128.0 89.8L124.9 88.1L125.0 87.1L127.7 86.0ZM149.3 80.3L154.2 84.7L153.6 91.4L146.8 88.3L146.4 85.9L147.3 84.9L146.3 83.5L146.6 81.7ZM138.1 81.7L139.9 81.1L141.7 82.4L146.6 81.7L146.3 83.5L147.3 84.9L146.4 85.9L146.8 88.3L148.0 88.8L146.7 90.8L147.8 91.5L144.7 95.6L140.5 91.7L138.0 93.0L137.0 92.4L139.5 87.5ZM154.2 84.7L156.1 83.9L158.9 87.2L167.2 91.0L170.3 89.4L171.9 92.5L169.6 92.9L166.3 95.9L164.1 94.6L160.5 96.2L160.0 94.8L157.9 93.8L155.6 94.0L153.6 91.4ZM194.2 95.9L195.6 97.2L197.0 95.7L198.3 96.3L201.2 94.3L204.1 93.7L208.2 99.3L207.8 100.1L209.9 101.7L212.6 100.3L214.2 100.7L217.8 98.6L218.2 101.7L216.2 102.8L220.9 106.4L219.8 109.5L222.5 111.0L222.1 112.3L223.8 114.1L219.1 114.5L218.5 117.3L213.8 117.4L210.2 121.3L206.7 120.8L204.6 117.1L205.3 116.1L204.1 113.8L198.3 110.7L197.3 111.5L195.1 108.5L191.6 106.4L188.0 106.6L185.6 105.3L185.6 103.7L182.4 103.6L179.8 101.3L177.1 101.0L175.8 99.2L175.3 97.2L177.6 97.6L181.6 96.5L183.9 94.1L182.9 90.7L183.6 87.3L187.0 92.4L191.6 93.6ZM171.9 92.5L175.9 88.8L179.3 88.2L180.9 89.3L183.1 88.9L183.9 94.1L181.6 96.5L177.6 97.6L175.3 97.2L177.1 101.0L179.8 101.3L182.4 103.6L185.6 103.7L185.6 105.3L183.1 107.7L183.4 111.1L181.3 111.8L176.3 109.9L171.9 106.9L170.7 109.0L167.9 109.7L167.7 108.2L163.8 104.2L165.5 101.4L165.6 98.8L163.0 95.7L164.1 94.6L166.3 95.9L169.6 92.9ZM132.0 94.1L137.0 92.4L138.0 93.0L140.5 91.7L144.7 95.6L145.6 94.7L149.6 98.4L148.0 99.5L146.2 98.9L143.9 101.1L144.8 102.5L143.3 104.7L139.3 104.6L138.4 101.9L138.8 97.9L135.8 98.9ZM160.5 96.2L163.0 95.7L165.6 98.8L165.5 101.4L163.8 104.2L156.9 108.8L150.4 106.3L152.1 101.2L154.8 100.4L156.5 98.4L159.9 98.3ZM154.1 111.7L158.5 112.6L158.8 113.7L157.3 117.2L158.1 120.4L154.6 125.1L147.6 120.2L145.4 119.7L143.7 117.5L144.2 114.9L149.8 114.3ZM185.6 105.3L190.0 106.8L187.0 114.4L183.6 114.4L181.3 115.5L182.0 117.4L180.8 118.8L179.0 118.5L176.4 120.6L173.6 120.9L173.8 122.2L169.6 120.5L168.1 119.0L166.4 119.1L165.3 117.0L158.7 114.0L160.1 112.7L159.6 110.1L162.4 109.3L164.5 106.8L167.7 108.2L167.9 109.7L170.7 109.0L171.9 106.9L176.3 109.9L181.3 111.8L183.4 111.1L183.1 107.7ZM190.0 106.8L191.6 106.4L195.1 108.5L197.3 111.5L198.3 110.7L204.1 113.8L205.3 116.1L204.6 117.1L206.7 120.8L202.1 123.5L203.4 127.2L201.7 128.4L201.7 129.8L199.4 130.7L196.5 130.4L194.7 126.1L194.8 123.4L192.9 122.1L192.2 119.5L189.6 119.1L187.0 114.4ZM156.1 123.9L158.1 120.4L157.3 117.2L158.7 114.0L165.3 117.0L167.6 124.3L171.5 131.8L173.4 131.7L174.2 133.6L172.2 133.1L171.7 135.6L170.2 135.3L167.1 132.8L164.9 132.4L164.0 130.2L162.7 130.3L163.0 128.3L160.2 124.8ZM144.2 114.9L143.7 117.5L145.4 119.7L147.6 120.2L154.6 125.1L152.2 127.5L142.8 126.6L141.7 125.6L137.8 127.3L135.5 121.8L136.7 121.4L137.0 119.5L139.0 117.6ZM180.8 118.8L182.0 117.4L181.3 115.5L183.6 114.4L187.0 114.4L189.6 119.1L192.2 119.5L192.9 122.1L194.8 123.4L194.7 126.1L196.5 130.4L194.9 132.5L191.7 133.4L185.8 130.7L184.2 129.1L182.5 123.5L180.7 121.6ZM206.7 120.8L210.2 121.3L213.8 117.4L218.5 117.3L219.1 114.5L223.8 114.1L226.1 116.0L224.5 121.0L227.5 124.8L225.1 126.4L226.4 128.6L225.4 130.3L226.3 131.7L228.2 132.2L230.9 136.2L228.3 136.4L226.7 134.1L222.5 133.5L219.8 134.1L215.7 131.2L213.0 130.8L211.7 129.1L208.8 129.5L206.2 127.1L205.6 123.5L204.2 122.4ZM166.4 119.1L168.1 119.0L169.6 120.5L173.8 122.2L175.7 124.8L177.0 127.6L173.4 131.7L171.5 131.8L167.6 124.3ZM196.5 130.4L199.4 130.7L201.7 129.8L201.7 128.4L203.4 127.2L202.1 123.5L204.2 122.4L205.6 123.5L206.2 127.1L208.8 129.5L211.7 129.1L213.0 130.8L215.7 131.2L216.2 132.2L206.7 136.2L205.5 142.0L203.3 144.1L202.4 143.1L200.0 143.8L198.8 143.3L196.4 138.2L193.6 140.5L189.4 136.1L189.1 132.5L191.7 133.4L194.9 132.5ZM158.7 124.9L160.2 124.8L163.0 128.3L162.7 130.3L160.5 131.5L160.3 132.5L165.1 135.1L166.2 138.4L158.6 139.4L153.7 138.8L147.7 135.3L151.8 135.3L151.9 134.4L155.8 133.3L156.5 129.9L155.4 127.4L157.1 127.3ZM137.8 127.3L141.7 125.6L142.8 126.6L152.2 127.5L148.3 131.4L147.0 134.3L140.6 136.0L139.9 137.0L139.5 135.7L135.7 135.7L136.6 130.3L138.7 128.2ZM174.2 133.6L173.4 131.7L175.5 129.3L177.8 132.5L179.7 132.4L180.9 133.9L180.4 138.7L182.7 140.1L180.1 141.2L178.5 144.0L176.6 144.2L175.0 141.6L174.0 141.5L173.7 138.9L177.7 139.2ZM216.2 132.2L217.4 132.1L219.8 134.1L222.5 133.5L226.7 134.1L228.3 136.4L230.9 136.2L228.2 132.2L232.3 128.3L233.7 128.8L234.3 130.4L235.7 131.1L235.4 132.7L239.4 137.0L236.9 139.4L235.5 143.0L236.5 145.0L231.4 143.0L230.8 141.9L228.2 141.5L227.5 144.7L224.8 144.7L222.3 146.2L221.1 149.3L217.9 150.0L215.4 147.5L213.9 148.3L212.9 147.2L209.8 147.4L211.2 144.6L209.0 141.2L208.8 139.2L206.5 139.7L206.0 138.0L206.7 136.2ZM185.8 130.7L189.1 132.5L189.4 136.1L193.6 140.5L188.9 144.4L186.6 141.0L185.6 141.6L180.4 138.7L180.9 133.9L185.0 133.4L186.9 131.7ZM164.9 132.4L171.7 135.6L172.2 133.1L174.2 133.6L177.7 139.2L166.2 138.4L165.1 135.1L163.0 133.7ZM134.4 162.7L122.7 162.1L122.9 160.3L124.4 158.8L124.1 155.3L122.3 153.3L124.1 151.8L124.1 150.4L129.5 145.7L133.7 147.5L132.7 148.1L132.5 152.3L130.5 155.9L132.5 156.0ZM149.2 136.6L153.7 138.8L154.4 143.1L154.0 147.2L154.7 148.2L153.7 149.8L149.9 149.8L149.6 148.3L146.5 150.3L144.9 148.5L145.1 146.0L144.1 144.7L145.0 141.1L147.7 139.9L148.1 137.6ZM153.7 138.8L158.6 139.4L161.9 138.4L173.1 138.3L174.0 141.5L173.2 142.6L173.6 146.0L170.4 146.5L169.0 148.7L167.4 147.4L165.9 148.5L164.7 146.1L160.5 148.0L154.7 148.2L154.0 147.2L154.4 143.1ZM236.5 145.0L235.5 143.0L236.9 139.4L239.4 137.0L243.5 140.2L249.0 139.3L249.6 140.5L254.1 143.5L251.6 145.8L252.8 147.9L251.8 149.2L252.7 151.4L251.6 154.2L249.2 155.3L246.9 158.5L247.9 161.0L243.7 161.0L239.7 163.0L235.6 162.1L232.4 156.5L234.1 154.0L234.3 151.3L235.9 148.7L237.8 147.9L238.1 146.7ZM203.3 144.1L205.5 142.0L206.0 138.0L206.5 139.7L208.8 139.2L209.0 141.2L211.2 144.6L209.0 150.4L211.8 150.6L214.3 152.2L212.4 153.7L215.4 155.3L214.5 157.3L216.0 158.9L213.9 162.3L211.0 160.5L208.7 156.0L203.7 154.0L202.3 152.3L200.6 152.3L205.7 150.0ZM193.6 140.5L196.4 138.2L198.8 143.3L202.4 143.1L203.3 144.1L205.7 150.0L200.7 152.8L199.5 151.4L197.0 151.7L192.8 150.4L191.0 148.6L192.0 147.8L191.9 145.9L188.9 144.4ZM178.5 144.0L180.1 141.2L182.7 140.1L183.7 140.9L185.1 142.5L185.0 144.3L182.3 144.1L182.9 145.7L186.8 147.6L186.5 151.8L183.1 151.7L183.5 150.0L182.8 148.4L178.5 145.3ZM144.1 144.7L145.1 146.0L144.9 148.5L146.5 150.3L146.0 152.2L144.4 153.1L140.9 152.5L139.5 149.6L139.9 148.7L137.0 146.3L135.5 146.3L134.9 144.4L138.1 142.4L141.2 142.9ZM174.0 141.5L175.0 141.6L176.6 144.2L178.5 144.0L178.5 145.3L180.7 147.4L176.4 151.4L172.1 148.5L169.0 149.2L170.4 146.5L173.6 146.0L173.2 142.6ZM209.8 147.4L212.9 147.2L213.9 148.3L215.4 147.5L217.9 150.0L221.1 149.3L222.3 146.2L224.8 144.7L227.5 144.7L228.2 141.5L230.8 141.9L231.4 143.0L228.8 145.7L227.8 148.5L227.3 153.3L228.6 155.9L223.3 156.8L220.4 160.2L218.3 158.4L216.0 158.9L214.5 157.3L215.4 155.3L212.4 153.7L214.3 152.2L211.8 150.6L209.0 150.4ZM184.8 145.9L182.9 145.7L182.3 144.1L185.0 144.3ZM134.4 162.7L132.5 156.0L130.5 155.9L132.5 152.3L132.7 148.1L135.5 146.3L137.0 146.3L139.9 148.7L139.5 149.6L140.9 152.5L142.3 152.5L140.2 154.4L139.8 162.1L138.6 163.5ZM231.4 143.0L236.5 145.0L238.1 146.7L234.3 151.3L234.1 154.0L232.4 156.5L230.4 157.4L227.3 153.3L228.8 145.7ZM251.6 154.2L252.7 151.4L251.8 149.2L252.8 147.9L251.6 145.8L255.6 143.1L257.2 144.5L256.4 148.2L263.4 151.3L264.8 150.7L274.0 156.7L271.6 157.5L266.2 162.5L264.6 162.5L263.6 164.9L259.7 167.4L260.6 170.7L257.0 174.2L251.9 171.5L250.9 168.8L249.6 168.1L251.6 166.9L253.1 163.7L255.3 163.7L253.3 161.1L254.4 156.3ZM160.5 148.0L164.7 146.1L165.9 148.5L167.4 147.4L169.0 148.7L169.1 151.1L167.0 153.7L170.8 154.2L170.7 156.0L168.3 157.2L169.4 159.1L166.9 163.6L169.3 165.8L168.3 167.3L166.1 167.7L162.5 165.5L161.3 165.7L162.3 163.4L160.4 163.6L160.6 161.2L157.8 157.0L161.5 155.3L159.9 149.1ZM152.3 149.7L153.7 149.8L155.9 147.7L158.3 148.5L160.5 148.0L159.9 149.1L161.5 155.3L157.8 157.0L158.3 158.2L153.9 156.9L153.1 152.9L151.4 150.9ZM146.5 150.3L149.6 148.3L149.9 149.8L152.3 149.7L151.4 150.9L153.9 155.5L153.3 158.5L151.4 159.1L151.1 162.4L149.9 162.2L150.1 165.0L146.0 164.5L144.5 162.6L142.1 161.7L139.8 162.1L140.2 154.4L142.3 152.5L144.4 153.1L146.0 152.2ZM180.7 147.4L182.8 148.4L183.5 150.0L182.3 154.1L179.1 156.0L180.8 158.1L179.4 159.2L177.1 159.2L175.9 158.2L172.6 159.5L170.9 163.2L171.8 165.1L169.3 165.8L166.9 163.6L169.4 159.1L168.3 157.2L170.7 156.0L170.8 154.2L167.0 153.7L169.1 151.1L169.0 149.2L172.1 148.5L176.4 151.4ZM192.8 150.4L197.0 151.7L199.5 151.4L200.7 154.5L193.2 162.9L190.9 161.7L189.8 158.6L192.2 154.1L191.9 151.5ZM200.6 152.3L202.3 152.3L203.7 154.0L208.7 156.0L205.4 159.3L203.6 164.1L205.8 168.9L204.1 170.4L191.6 168.6L190.9 167.1L193.2 162.9L200.7 154.5ZM124.2 157.1L124.4 158.8L122.9 160.3L122.7 162.1L131.2 162.5L130.3 163.9L130.9 165.1L128.8 167.2L121.3 169.3L124.1 171.4L125.3 177.8L126.8 178.4L128.4 176.9L129.5 178.8L129.2 180.9L125.8 184.1L125.4 185.8L122.3 185.7L120.6 184.2L116.8 183.5L114.7 184.5L110.8 184.0L109.7 185.8L108.1 186.1L108.1 187.8L103.9 190.6L100.7 190.6L98.4 189.3L96.6 186.8L102.6 183.7L109.7 171.3L112.1 162.7L120.8 160.0ZM158.3 158.2L160.6 161.2L160.4 163.6L162.3 163.4L161.3 165.7L158.3 167.4L155.8 165.9L153.4 168.4L152.5 171.5L149.6 169.1L151.2 167.2L149.9 162.2L151.1 162.4L151.4 159.1L153.3 158.5L153.9 156.9ZM180.7 158.1L184.5 156.5L187.6 157.1L189.8 158.6L190.9 161.7L190.4 162.6L188.1 162.4L185.9 164.2L177.1 163.5L176.2 160.8L177.1 159.2L179.4 159.2ZM240.7 162.0L247.9 161.0L246.9 158.5L249.2 155.3L251.6 154.2L254.4 156.3L253.3 161.1L255.3 163.7L253.1 163.7L251.6 166.9L249.6 168.1L246.8 168.3L245.5 166.8L245.7 165.5L241.5 164.4ZM208.7 156.0L211.0 160.5L213.9 162.3L216.7 166.1L210.0 171.0L208.0 168.0L205.8 168.9L203.6 164.1L205.4 159.3ZM232.4 156.5L235.6 162.1L232.5 164.0L233.3 167.9L229.2 174.3L222.6 171.6L220.4 167.8L216.7 166.1L213.9 162.3L216.0 158.9L218.3 158.4L220.4 160.2L224.7 156.1L228.6 155.9L230.4 157.4ZM171.8 165.1L170.9 163.2L172.6 159.5L175.9 158.2L177.1 159.2L176.2 160.8L177.1 163.5L181.2 163.5L182.7 164.4L181.7 168.8L178.7 170.3L177.9 172.5L172.3 170.8L170.8 168.6ZM136.0 162.8L138.6 163.5L139.8 162.1L142.1 161.7L144.5 162.6L146.0 164.5L150.1 165.0L151.2 167.2L149.6 169.1L152.5 171.5L152.4 174.3L150.7 175.2L151.1 178.4L147.8 179.1L147.1 177.8L145.2 177.9L144.7 176.1L142.9 175.1L141.6 176.1L137.6 176.0L137.3 173.9L138.1 172.2L137.1 171.4L137.9 167.1L135.1 167.3L135.2 165.7L136.8 164.5ZM156.1 186.0L153.8 186.3L153.1 184.7L151.2 185.9L150.7 187.4L141.9 185.9L140.8 187.7L135.0 188.5L133.8 186.9L130.6 189.0L126.1 189.4L127.7 187.1L125.8 184.1L129.2 180.9L129.5 178.8L128.4 176.9L126.8 178.4L125.3 177.8L124.1 171.4L121.3 169.3L128.8 167.2L130.9 165.1L130.3 163.9L131.2 162.5L136.0 162.8L136.8 164.5L135.2 165.7L135.1 167.3L137.9 167.1L137.1 171.4L138.1 172.2L137.3 173.9L137.6 176.0L141.6 176.1L142.9 175.1L144.7 176.1L145.2 177.9L147.1 177.8L147.8 179.1L153.4 177.8L154.6 178.4L154.4 182.9L157.2 185.0ZM190.9 161.7L193.2 162.9L190.9 167.1L191.6 168.6L190.6 169.9L191.2 173.3L188.6 174.1L184.6 172.6L183.3 173.8L178.5 173.7L177.9 172.5L178.7 170.3L181.7 168.8L182.7 164.4L185.9 164.2L188.1 162.4L190.4 162.6ZM235.6 162.1L237.6 163.1L239.7 163.0L240.7 162.0L241.5 164.4L245.7 165.5L245.5 166.8L246.8 168.3L250.9 168.8L251.9 171.5L253.5 172.4L248.7 174.5L245.3 173.5L243.9 171.6L241.9 171.8L240.9 173.0L237.8 170.3L235.4 172.4L232.0 170.9L233.3 167.9L232.5 164.0ZM166.1 167.7L171.8 165.1L170.8 168.6L172.3 170.8L177.9 172.5L178.5 173.7L183.3 173.8L182.8 175.5L181.0 175.9L181.5 177.9L179.1 177.9L177.3 182.2L173.6 182.3L173.6 180.0L172.0 179.3L171.3 175.0L170.0 175.4L166.5 171.5ZM161.3 165.7L162.5 165.5L166.1 167.7L166.5 171.5L164.2 171.8L159.4 174.8L159.1 175.8L155.1 175.9L153.4 177.8L151.1 178.4L150.7 175.2L152.4 174.3L153.4 168.4L155.8 165.9L158.3 167.4ZM220.8 182.8L217.7 183.8L211.1 182.4L212.2 180.4L212.0 178.5L209.3 172.3L213.9 167.5L216.7 166.1L220.4 167.8L222.6 171.6L229.2 174.3L225.7 178.3L224.1 178.8L223.3 181.5ZM198.7 169.8L204.1 170.4L208.0 168.0L210.0 171.0L209.3 172.3L210.1 175.5L212.0 178.5L209.4 178.5L205.2 180.1L200.8 178.3L199.1 179.9L194.3 178.2L198.2 174.0ZM157.2 185.0L154.4 182.9L154.6 178.4L153.4 177.8L155.1 175.9L159.1 175.8L159.4 174.8L164.2 171.8L166.5 171.5L170.0 175.4L171.3 175.0L172.0 179.3L173.6 180.0L173.6 182.3L172.1 184.5L171.3 183.1L166.6 182.9L163.7 186.3L160.5 186.3L159.4 184.8ZM229.2 174.3L230.9 170.9L235.4 172.4L237.8 170.3L240.9 173.0L241.9 171.8L243.9 171.6L245.3 173.5L248.7 174.5L249.6 175.9L248.8 178.0L250.4 179.4L249.3 180.7L251.4 181.5L251.3 183.1L249.7 184.5L239.9 182.4L236.6 178.8L227.6 176.4ZM183.3 173.8L184.6 172.6L188.6 174.1L191.2 173.3L191.4 180.5L184.4 179.4L183.7 177.9L181.5 177.9L181.0 175.9L182.8 175.5ZM253.5 172.4L257.0 174.2L258.7 177.4L257.4 180.0L255.0 182.0L253.5 185.7L253.2 184.2L251.3 183.1L251.4 181.5L249.3 180.7L250.4 179.4L248.8 178.0L249.6 175.9L248.7 174.5ZM245.2 187.6L239.7 186.7L239.4 185.4L233.8 182.8L234.0 181.7L231.5 180.4L229.2 181.6L228.1 179.6L225.7 178.3L227.6 176.4L229.2 176.5L236.6 178.8L239.9 182.4L249.7 184.5L251.3 183.1L253.2 184.2L253.5 185.7L252.5 190.0L251.0 190.6L248.1 186.9ZM193.0 187.7L185.9 188.7L182.8 184.2L177.3 182.2L177.7 180.3L179.2 179.2L179.1 177.9L183.7 177.9L184.4 179.4L191.4 180.5L193.3 186.2ZM193.1 188.2L193.3 186.2L191.4 180.5L191.7 177.5L199.1 179.9L197.5 181.9L198.5 185.1L196.7 188.2ZM211.1 182.4L212.1 187.1L210.1 187.1L203.0 190.1L202.9 187.2L197.9 186.0L198.5 185.1L197.5 181.9L200.8 178.3L205.2 180.1L209.4 178.5L212.0 178.5L212.2 180.4ZM222.8 188.7L223.5 187.8L222.6 185.6L220.7 185.1L220.8 182.8L223.3 181.5L224.1 178.8L225.7 178.3L228.1 179.6L229.2 181.6L230.2 181.3L228.7 184.4L230.8 189.7L227.8 192.1L229.9 194.6L229.5 196.6L226.5 194.9L225.5 192.4L223.2 191.9L224.0 189.7ZM128.3 189.4L129.0 191.0L126.6 191.8L127.2 193.8L126.0 195.1L125.5 198.6L126.2 199.6L129.1 199.6L130.1 201.5L132.5 202.4L129.9 207.1L129.7 211.3L127.4 211.8L126.3 213.3L124.8 213.2L124.7 211.4L122.1 212.2L122.2 217.7L120.6 218.6L115.2 218.7L114.0 221.8L111.3 224.0L109.2 224.0L108.5 222.3L106.5 222.9L106.4 220.8L105.3 220.2L103.3 221.2L100.0 219.3L100.0 216.6L96.4 217.7L94.3 215.8L87.7 218.9L89.7 214.1L87.1 214.5L84.6 213.0L86.0 211.9L85.5 209.9L83.2 208.2L80.0 209.1L81.0 204.4L86.9 198.1L88.9 191.9L91.1 189.3L96.6 186.8L100.7 190.6L103.9 190.6L108.1 187.8L108.1 186.1L109.7 185.8L110.8 184.0L114.7 184.5L116.8 183.5L120.6 184.2L122.3 185.7L125.4 185.8L125.8 184.1L127.7 187.1L126.1 189.4ZM239.7 186.7L238.4 190.0L235.7 192.1L235.1 194.1L230.1 197.5L229.9 194.6L227.8 192.1L230.8 189.7L228.7 184.4L230.2 181.3L231.5 180.4L234.0 181.7L233.8 182.8L239.4 185.4ZM160.2 195.3L159.2 193.4L157.9 193.2L156.1 186.0L157.2 185.0L159.4 184.8L160.5 186.3L163.7 186.3L166.6 182.9L171.3 183.1L172.2 186.3L174.7 186.7L172.2 191.7L170.1 191.5L164.0 193.8L163.4 195.5L165.2 197.3L165.0 198.2ZM185.9 188.7L183.3 189.6L180.0 189.0L179.5 190.4L177.2 190.2L176.2 191.1L174.1 188.6L174.7 186.7L172.2 186.3L172.1 184.5L173.6 182.3L181.1 183.1ZM191.6 168.6L195.5 169.6L197.7 169.0L198.7 169.8L198.2 174.0L194.3 178.2L191.7 177.5L191.1 176.3L190.6 169.9ZM147.7 135.3L149.2 136.6L148.1 137.6L147.7 139.9L145.0 141.1L144.1 144.7L141.2 142.9L138.1 142.4L134.9 144.4L135.5 146.3L133.7 147.5L129.5 145.7L132.6 141.1L140.6 136.0ZM163.8 104.2L165.8 106.9L159.6 110.1L160.1 112.7L158.8 113.7L158.5 112.6L154.1 111.7L153.3 107.9L156.9 108.8ZM148.0 88.8L151.0 89.6L155.6 94.0L160.0 94.8L159.9 98.3L156.5 98.4L154.8 100.4L152.1 101.2L149.5 99.4L148.0 99.5L149.6 98.4L145.6 94.7L147.8 91.5L146.7 90.8ZM143.7 104.0L144.8 102.5L143.9 101.1L146.2 98.9L149.5 99.4L152.1 101.2L150.4 106.3L148.3 105.3L144.9 105.7ZM162.7 130.3L164.0 130.2L164.9 132.4L163.0 133.7L160.3 132.5L160.5 131.5ZM154.6 125.1L156.1 123.9L158.7 124.9L157.1 127.3L155.4 127.4L156.5 129.9L155.8 133.3L151.9 134.4L151.8 135.3L145.3 135.2L147.0 134.3L148.3 131.4ZM183.7 140.9L185.6 141.6L186.6 141.0L188.9 144.4L191.9 145.9L192.0 147.8L186.4 151.0L186.8 147.6L184.8 145.9L185.1 142.5ZM191.0 148.6L192.8 150.4L191.9 151.5L192.2 154.1L189.8 158.6L187.6 157.1L184.5 156.5L180.8 158.1L179.1 156.0L182.3 154.1L183.1 151.7L186.5 151.8L186.4 151.0ZM295.8 220.8L294.7 216.7L298.6 210.7L303.8 214.5L309.2 213.9L310.2 216.6L309.4 220.8L306.5 220.3L304.7 221.8L305.1 225.1L303.2 228.0L301.1 228.4L299.1 226.1L300.5 224.3ZM251.1 249.4L251.0 251.0L254.3 251.7L253.7 252.8L259.2 255.1L256.8 257.4L258.7 258.4L256.5 263.8L253.5 262.2L247.4 262.1L244.3 264.5L239.1 263.7L238.3 260.9L239.2 257.6L241.3 256.0L242.4 256.8L243.7 255.1L249.2 255.5L249.1 254.2L251.0 252.2L250.1 250.7ZM173.8 122.2L173.6 120.9L176.4 120.6L179.0 118.5L180.8 118.8L180.7 121.6L182.5 123.5L182.7 125.4L180.3 125.5L180.2 129.2L177.0 127.6ZM130.6 299.6L129.6 296.1L129.9 293.7L132.4 294.6L134.0 292.5L135.9 291.9L137.5 286.8L142.1 287.7L143.0 288.7L140.9 292.9L144.1 293.6L143.5 294.5L145.5 297.5L144.0 301.3L144.4 306.0L142.8 308.6L141.2 308.9L140.7 307.2L138.7 306.5L135.9 306.7L136.8 302.9L135.1 300.1L132.9 300.6ZM295.3 251.4L303.4 252.6L302.1 253.2L301.7 255.0L300.2 256.4L298.9 256.4L297.4 260.2L295.6 260.6L294.7 258.6L291.8 259.3L290.0 258.5L289.0 260.3L286.0 262.0L283.0 261.5L281.0 259.7L281.4 258.0L280.1 255.7L285.1 253.3L285.8 249.5L289.3 251.2ZM182.7 125.4L184.2 129.1L186.9 131.7L185.0 133.4L180.9 133.9L179.7 132.4L177.8 132.5L175.5 129.3L177.0 127.6L180.2 129.2L180.3 125.5ZM128.0 45.2L130.2 46.6L135.2 44.9L137.3 47.2L136.9 50.3L134.6 50.8L128.2 54.6L124.4 54.6L122.6 60.0L120.8 60.5L119.8 62.2L123.6 62.6L124.8 64.3L125.0 66.0L122.7 68.7L117.5 68.8L114.0 66.9L114.1 65.2L112.0 61.7L113.2 57.1L116.8 56.8L119.6 52.3L124.1 51.0L126.4 48.8L126.3 46.3ZM116.4 76.8L118.6 78.8L121.7 79.5L123.9 78.8L128.0 80.7L127.7 86.0L125.0 87.1L124.9 88.1L128.0 89.8L128.4 91.3L131.7 92.4L134.3 97.8L133.0 98.6L124.0 93.3L117.0 91.0L117.6 88.3L115.7 85.2L116.9 79.9L115.6 78.9ZM225.7 178.3L228.1 179.6L229.2 181.6L231.5 180.4L234.0 181.7L233.8 182.8L239.4 185.4L239.7 186.7L243.1 187.8L248.1 186.9L251.0 190.6L252.5 190.0L253.3 188.0L257.8 190.9L258.6 192.8L261.6 192.9L263.6 191.4L266.0 192.5L267.0 194.2L277.6 199.0L279.9 202.6L282.1 203.2L283.2 205.2L291.8 210.4L293.4 208.9L295.7 208.9L303.8 214.5L309.2 213.9L310.2 216.6L309.7 218.8L313.9 218.5L316.2 219.7L319.4 219.6L321.6 222.3L322.7 222.3L324.1 220.5L323.8 219.0L328.4 219.3L333.7 222.3L336.4 226.0L335.8 228.0L337.6 228.9L337.9 232.5L342.1 234.4L343.6 237.2L345.4 238.4L342.5 239.1L339.0 238.4L337.7 240.2L334.8 241.3L334.6 242.6L339.9 244.2L340.3 246.7L338.4 247.8L337.8 250.5L349.8 257.5L348.4 260.1L343.7 258.6L343.3 260.4L340.9 260.3L339.6 258.9L333.3 265.7L326.8 268.5L322.8 271.4L323.8 280.6L325.8 281.6L326.5 285.9L323.9 287.1L324.1 292.3L319.8 300.2L317.3 301.4L314.3 301.4L308.3 295.5L310.0 294.6L310.6 290.0L309.4 289.7L311.1 285.3L310.6 283.8L309.2 283.9L308.4 282.5L303.1 283.5L303.1 284.9L300.5 284.6L298.8 280.3L293.0 279.1L292.0 275.9L287.2 274.2L286.3 271.9L284.5 271.3L281.4 272.6L279.0 271.8L278.7 274.4L275.6 277.8L271.9 276.4L270.1 277.2L269.8 271.3L267.2 272.9L267.0 274.1L264.5 273.5L262.5 274.7L262.0 276.1L260.4 274.0L259.1 275.3L257.5 275.0L261.0 271.6L260.5 270.0L258.5 269.1L258.6 266.7L256.6 266.5L252.6 267.8L248.8 271.2L248.5 273.3L242.8 271.7L241.1 273.8L239.3 273.2L240.0 269.8L238.2 268.1L236.7 268.2L236.3 269.4L237.6 272.2L233.5 273.1L232.1 270.5L229.3 270.9L228.2 272.0L229.9 267.7L228.2 266.8L229.5 264.4L228.3 263.8L225.1 265.2L226.0 267.5L223.5 266.8L220.4 269.3L223.1 274.0L223.0 276.6L225.9 278.8L225.5 284.7L228.3 283.8L229.7 285.9L228.9 287.9L230.4 288.3L230.1 290.5L226.6 294.3L220.2 289.5L217.8 292.4L216.5 290.9L217.4 289.9L215.2 288.1L215.4 283.3L213.2 278.4L216.9 275.3L216.7 273.7L219.0 272.8L216.9 268.1L216.8 265.7L217.9 265.5L219.1 262.5L223.5 262.8L226.7 261.5L225.7 259.2L231.0 250.9L230.6 248.3L233.7 245.0L233.6 242.7L230.5 239.6L230.9 237.5L228.0 236.9L226.1 235.2L221.1 235.3L217.9 234.1L217.2 232.9L214.3 233.8L215.3 231.8L212.7 231.1L211.9 232.3L209.3 232.5L204.5 231.5L203.7 232.8L199.5 234.1L198.0 235.8L196.8 233.8L202.4 230.6L200.1 227.4L202.3 226.1L202.1 224.4L195.6 217.9L195.9 215.1L194.4 211.7L199.6 208.8L199.9 207.4L198.5 205.2L199.1 198.7L194.7 194.2L195.6 193.2L192.8 189.1L193.3 186.2L191.4 180.5L190.6 169.9L191.6 168.6L190.9 167.1L193.2 162.9L200.7 154.5L200.6 152.3L202.3 152.3L203.7 154.0L208.7 156.0L205.4 159.3L203.6 164.1L205.8 168.9L208.0 168.0L210.0 171.0L216.7 166.1L220.4 167.8L222.6 171.6L229.2 174.3ZM165.8 208.8L166.6 210.5L171.0 211.7L172.0 210.9L170.6 207.2L172.7 205.6L174.0 206.8L174.8 204.3L173.8 203.1L178.7 206.4L178.6 207.7L181.4 206.8L181.8 205.1L185.3 202.1L187.8 204.0L187.2 210.7L186.1 213.2L187.9 215.0L189.3 211.2L194.4 211.7L195.9 215.1L195.6 217.9L202.1 224.4L202.3 226.1L200.1 227.4L202.4 230.6L196.8 233.8L198.0 235.8L199.5 234.1L203.7 232.8L204.5 231.5L209.3 232.5L211.9 232.3L212.7 231.1L215.3 231.8L214.3 233.8L211.9 234.6L211.5 237.4L209.6 236.9L207.5 237.7L205.8 240.2L199.8 243.4L197.7 243.6L194.1 246.7L186.3 250.7L183.9 254.0L178.8 256.5L178.1 263.6L181.4 267.6L185.4 268.0L187.2 269.1L190.2 267.9L193.1 268.3L193.4 266.9L195.0 265.7L196.2 266.2L197.1 270.5L196.7 272.9L195.0 273.7L192.1 273.1L190.2 274.3L185.8 274.7L185.5 276.5L187.3 278.6L184.3 280.0L185.3 281.4L187.3 281.1L189.1 282.3L189.9 285.7L187.9 287.9L185.5 285.9L184.5 286.2L184.8 290.2L186.8 293.6L185.9 295.4L184.2 295.7L181.5 294.0L180.9 292.3L179.3 293.6L174.0 292.8L169.7 294.0L169.9 296.4L168.1 296.9L165.9 300.8L162.9 301.4L160.7 304.2L155.8 300.3L157.1 297.8L160.9 299.2L163.8 297.1L162.0 295.1L163.3 292.8L161.5 289.4L162.7 288.1L165.1 288.3L165.6 286.0L162.7 281.0L158.9 282.8L149.9 281.3L153.5 275.2L150.1 274.8L148.5 276.4L148.5 278.4L146.9 279.1L145.2 278.8L144.0 277.2L142.4 277.9L143.1 280.2L145.6 281.0L145.9 282.6L143.7 283.9L141.9 281.6L141.5 285.0L139.7 287.4L142.1 287.7L143.0 288.7L140.9 292.9L144.1 293.6L143.5 294.5L145.5 297.5L144.0 301.3L144.8 303.8L142.8 308.6L137.0 311.6L136.1 313.9L140.4 316.2L135.6 319.3L133.1 318.7L131.7 319.9L130.1 317.0L127.6 317.2L127.8 315.2L124.0 312.6L122.6 313.6L121.3 311.5L117.0 311.0L117.6 307.1L115.8 306.3L114.8 307.4L111.5 302.6L112.5 300.1L112.3 297.5L110.4 297.2L108.9 298.4L108.3 296.6L105.1 294.0L106.9 290.3L105.3 289.4L105.6 287.1L103.8 287.5L102.8 290.2L98.0 290.0L97.0 287.8L93.6 287.0L93.1 288.9L89.6 284.0L86.0 284.2L83.8 282.2L75.3 283.8L71.5 282.9L67.5 284.4L63.7 283.0L60.1 277.0L59.0 272.2L54.3 266.5L54.5 260.2L53.1 259.3L48.4 260.2L45.9 259.3L42.3 254.2L42.0 250.9L43.9 247.1L43.9 240.2L41.4 239.1L35.8 239.1L29.8 235.8L29.3 234.4L31.5 225.7L40.8 217.1L43.0 211.6L48.1 206.9L52.7 206.9L54.5 208.9L54.7 211.1L56.3 213.5L59.4 213.8L65.7 210.9L75.5 209.9L79.9 208.2L80.0 209.1L81.0 204.4L86.9 198.1L88.9 191.9L91.1 189.3L102.6 183.7L109.7 171.3L112.1 162.7L120.8 160.0L124.2 157.1L124.4 158.8L122.9 160.3L122.7 162.1L136.0 162.8L136.8 164.5L135.2 165.7L135.1 167.3L137.9 167.1L137.1 171.4L138.1 172.2L137.3 173.9L137.6 176.0L141.6 176.1L142.9 175.1L144.7 176.1L145.2 177.9L147.1 177.8L147.8 179.1L153.4 177.8L154.6 178.4L154.4 182.9L157.2 185.0L156.1 186.0L157.9 193.2L165.0 198.2L168.4 202.1L166.7 208.7ZM193.2 189.6L194.7 191.0L195.6 193.2L194.7 194.2L195.8 195.3L192.1 197.9L190.4 195.5L188.5 194.8L185.9 195.7L185.1 194.6L187.8 191.2L187.2 188.9L193.0 187.7ZM193.0 187.7L187.2 188.9L187.8 191.2L185.1 194.6L185.9 195.7L188.5 194.8L190.4 195.5L192.1 197.9L195.8 195.3L199.1 198.7L198.5 205.2L199.9 207.4L199.6 208.8L194.4 211.7L189.3 211.2L187.9 215.0L186.1 213.2L187.2 210.7L187.8 204.0L185.3 202.1L181.8 205.1L181.4 206.8L178.6 207.7L178.7 206.4L173.8 203.1L174.8 204.3L174.0 206.8L172.7 205.6L170.6 207.2L172.0 210.9L171.0 211.7L166.6 210.5L165.8 208.8L166.7 208.7L168.4 202.1L165.0 198.2L157.9 193.2L156.1 186.0L157.2 185.0L154.4 182.9L154.6 178.4L153.4 177.8L147.8 179.1L147.1 177.8L145.2 177.9L144.7 176.1L142.9 175.1L141.6 176.1L137.6 176.0L137.3 173.9L138.1 172.2L137.1 171.4L137.9 167.1L135.1 167.3L135.2 165.7L136.8 164.5L136.0 162.8L138.6 163.5L139.8 162.1L142.1 161.7L144.5 162.6L146.0 164.5L150.1 165.0L151.2 167.2L149.6 169.1L152.5 171.5L153.4 168.4L155.8 165.9L158.3 167.4L162.5 165.5L166.1 167.7L168.3 167.3L169.3 165.8L171.8 165.1L170.9 163.2L172.6 159.5L175.9 158.2L177.1 159.2L179.4 159.2L180.8 158.1L179.1 156.0L182.3 154.1L183.1 151.7L186.5 151.8L186.8 147.6L184.8 145.9L185.1 142.5L183.7 140.9L185.6 141.6L186.6 141.0L188.9 144.4L191.9 145.9L192.0 147.8L191.0 148.6L192.8 150.4L197.0 151.7L199.5 151.4L200.7 152.8L200.7 154.5L193.2 162.9L190.9 167.1L191.6 168.6L190.6 169.9L191.4 180.5L193.3 186.2ZM154.2 62.0L152.8 60.3L153.6 58.1L151.1 55.9L140.4 54.9L136.8 52.7L131.7 52.3L136.9 50.3L137.3 47.2L135.2 44.9L130.2 46.6L120.6 41.8L119.2 42.4L118.5 39.2L120.5 36.9L120.3 35.4L117.2 34.4L112.3 34.5L106.6 31.0L107.0 28.2L105.6 27.3L97.2 28.4L93.8 26.5L95.1 24.0L93.8 22.9L95.1 18.6L99.0 17.5L101.2 15.7L101.2 14.5L103.5 14.1L106.4 8.9L113.0 7.6L118.5 9.0L120.8 8.3L122.0 6.6L119.0 5.6L118.3 3.8L122.3 4.2L126.1 6.2L127.9 5.4L128.5 4.0L130.4 4.4L134.0 2.0L136.1 2.1L139.9 0.0L140.4 1.3L142.7 1.3L144.4 3.9L146.4 2.3L148.1 3.1L150.5 0.9L153.3 0.6L156.1 6.1L159.1 7.2L161.9 11.2L168.9 13.8L172.6 16.7L174.7 16.4L177.4 19.8L180.8 20.4L184.2 23.5L182.5 25.7L184.4 28.1L187.0 29.5L190.8 29.2L191.9 30.7L197.5 32.5L198.2 34.3L197.6 35.6L201.4 36.7L205.7 35.7L207.4 36.7L208.8 34.4L210.4 34.0L212.5 35.1L213.8 32.1L218.5 29.5L220.6 30.0L223.6 27.8L229.9 27.0L233.5 28.3L234.8 27.2L235.3 25.1L238.7 24.7L239.7 26.9L249.1 29.6L251.9 28.3L254.2 31.9L258.0 33.5L258.8 36.7L255.7 43.6L255.6 47.3L253.4 51.3L253.3 53.9L248.9 54.5L247.1 55.9L247.5 59.2L241.5 59.4L243.2 63.8L241.4 65.5L241.2 67.1L238.2 69.4L231.2 69.0L228.7 70.8L233.0 79.0L228.9 79.9L230.5 85.4L240.3 87.2L238.7 91.1L239.7 93.9L244.0 99.0L240.2 103.7L237.5 102.3L233.6 105.0L233.1 107.1L231.4 107.0L230.3 108.1L227.7 105.8L225.7 100.9L223.6 100.9L218.0 103.9L216.2 102.8L218.2 101.7L217.8 98.6L214.2 100.7L212.6 100.3L209.9 101.7L207.8 100.1L208.2 99.3L204.1 93.7L201.2 94.3L198.3 96.3L197.0 95.7L195.6 97.2L191.6 93.6L187.0 92.4L177.9 81.3L173.8 80.0L172.8 76.1L170.8 74.9L167.7 70.3L164.4 70.7L160.6 66.6L154.8 63.7ZM154.0 61.4L154.8 63.7L160.6 66.6L164.4 70.7L167.7 70.3L170.8 74.9L172.8 76.1L173.8 80.0L177.9 81.3L183.6 87.3L183.1 88.9L180.9 89.3L179.3 88.2L175.9 88.8L171.9 92.5L169.6 92.9L166.3 95.9L164.1 94.6L163.0 95.7L165.6 98.8L165.5 101.4L160.6 106.9L158.7 107.2L156.9 108.8L148.3 105.3L144.9 105.7L143.7 104.0L143.3 104.7L139.3 104.6L138.4 101.9L138.8 97.9L135.8 98.9L134.3 97.8L133.0 98.6L124.0 93.3L117.0 91.0L117.6 88.3L115.7 85.2L116.9 79.9L115.6 78.9L116.2 72.9L114.8 70.6L114.1 65.2L112.0 61.7L113.2 57.1L116.8 56.8L119.6 52.3L124.1 51.0L126.4 48.8L126.3 46.3L128.0 45.2L130.2 46.6L135.2 44.9L137.3 47.2L136.9 50.3L131.7 52.3L136.8 52.7L140.4 54.9L151.1 55.9L153.6 58.1L152.8 60.3ZM216.2 132.2L206.7 136.2L205.5 142.0L203.3 144.1L205.7 150.0L200.7 152.8L199.5 151.4L197.0 151.7L192.8 150.4L191.0 148.6L192.0 147.8L191.9 145.9L188.9 144.4L186.6 141.0L185.6 141.6L180.4 138.7L180.9 133.9L179.7 132.4L177.8 132.5L175.5 129.3L173.4 131.7L171.5 131.8L167.6 124.3L165.3 117.0L158.7 114.0L160.1 112.7L159.6 110.1L165.8 106.9L163.8 104.2L165.5 101.4L165.6 98.8L163.0 95.7L164.1 94.6L166.3 95.9L169.6 92.9L171.9 92.5L175.9 88.8L179.3 88.2L180.9 89.3L183.1 88.9L183.6 87.3L187.0 92.4L191.6 93.6L195.6 97.2L197.0 95.7L198.3 96.3L201.2 94.3L204.1 93.7L208.2 99.3L207.8 100.1L209.9 101.7L212.6 100.3L214.2 100.7L217.8 98.6L218.2 101.7L216.2 102.8L220.9 106.4L219.8 109.5L222.5 111.0L222.1 112.3L226.1 116.0L224.5 121.0L227.5 124.8L225.1 126.4L226.4 128.6L225.4 130.3L226.3 131.7L228.2 132.2L230.9 136.2L228.3 136.4L226.7 134.1L222.5 133.5L219.8 134.1L217.4 132.1ZM163.8 104.2L165.8 106.9L159.6 110.1L160.1 112.7L158.7 114.0L165.3 117.0L167.6 124.3L171.5 131.8L173.4 131.7L175.5 129.3L177.8 132.5L179.7 132.4L180.9 133.9L180.4 138.7L185.1 142.5L185.0 144.3L182.3 144.1L182.9 145.7L186.8 147.6L186.5 151.8L183.1 151.7L182.3 154.1L179.1 156.0L180.8 158.1L179.4 159.2L177.1 159.2L175.9 158.2L172.6 159.5L170.9 163.2L171.8 165.1L169.3 165.8L168.3 167.3L166.1 167.7L162.5 165.5L158.3 167.4L155.8 165.9L153.4 168.4L152.5 171.5L149.6 169.1L151.2 167.2L150.1 165.0L146.0 164.5L144.5 162.6L142.1 161.7L139.8 162.1L138.6 163.5L122.7 162.1L122.9 160.3L124.4 158.8L124.1 155.3L122.3 153.3L124.1 151.8L124.1 150.4L129.5 145.7L132.6 141.1L139.9 137.0L139.5 135.7L135.7 135.7L136.6 130.3L138.7 128.2L135.5 121.8L136.7 121.4L137.0 119.5L144.2 114.9L149.8 114.3L154.1 111.7L153.3 107.9L156.9 108.8ZM254.1 143.5L255.6 143.1L257.2 144.5L256.4 148.2L263.4 151.3L264.8 150.7L274.0 156.7L271.6 157.5L266.2 162.5L264.6 162.5L263.6 164.9L259.7 167.4L260.6 170.7L257.0 174.2L258.7 177.4L257.4 180.0L255.0 182.0L252.5 190.0L251.0 190.6L248.1 186.9L243.1 187.8L239.7 186.7L239.4 185.4L236.2 184.5L233.8 182.8L234.0 181.7L231.5 180.4L229.2 181.6L228.1 179.6L225.7 178.3L229.2 174.3L222.6 171.6L220.4 167.8L216.7 166.1L210.0 171.0L208.0 168.0L205.8 168.9L203.6 164.1L205.4 159.3L208.7 156.0L203.7 154.0L202.3 152.3L200.6 152.3L205.7 150.0L203.3 144.1L205.5 142.0L206.7 136.2L208.9 134.7L217.4 132.1L219.8 134.1L222.5 133.5L226.7 134.1L228.3 136.4L230.9 136.2L228.2 132.2L232.3 128.3L235.7 131.1L235.4 132.7L237.0 133.7L237.5 135.6L243.5 140.2L249.0 139.3L249.6 140.5ZM184.8 145.9L182.9 145.7L182.3 144.1L185.0 144.3Z","West":"M265.9 359.3L264.9 361.7L261.7 362.9L261.2 364.7L261.7 368.5L263.5 368.4L263.8 370.9L258.9 371.0L257.0 374.7L253.1 373.4L251.5 374.2L251.8 372.6L249.0 371.8L250.6 367.9L253.4 367.9L253.3 362.3L248.2 360.6L247.6 358.3L249.4 357.1L250.2 354.9L253.3 354.2L254.5 352.9L257.6 352.8L260.9 356.8L265.1 359.5ZM250.2 354.9L249.4 357.1L247.6 358.3L248.2 360.6L253.3 362.3L253.4 367.9L250.6 367.9L249.0 371.8L251.8 372.6L251.5 374.2L244.7 374.5L241.7 373.4L244.6 367.7L243.4 365.2L241.2 363.7L242.4 363.0L242.9 360.6L240.8 355.7L246.4 353.1ZM166.9 358.5L169.2 358.4L170.6 361.1L170.8 365.1L175.3 365.1L174.6 368.0L170.7 366.0L168.9 369.0L169.6 369.3L166.9 372.3L167.4 374.9L165.8 376.6L163.5 377.0L162.0 375.7L157.4 375.9L157.2 377.0L154.4 377.8L154.3 376.5L151.5 376.6L150.9 378.4L148.3 380.3L148.7 381.7L145.9 382.0L144.7 383.5L141.9 379.5L142.1 375.3L144.6 374.7L145.8 372.9L145.3 369.1L143.1 365.5L143.8 363.5L146.8 362.5L148.8 360.5L148.6 359.0L150.5 357.5ZM214.9 359.3L216.6 361.2L219.0 361.5L229.5 372.2L233.8 374.2L235.4 376.4L233.5 379.3L226.9 380.6L226.5 382.9L225.2 383.0L220.3 377.7L218.3 378.0L215.1 376.5L216.2 372.9L215.0 369.5L212.6 367.9L212.6 363.7L210.7 361.1L211.2 358.8ZM175.3 365.1L177.6 362.5L180.4 362.7L181.3 360.5L182.5 360.3L183.8 362.5L184.1 366.8L182.7 374.7L183.8 375.9L183.5 377.9L182.2 378.3L183.0 382.0L184.6 383.3L183.0 385.0L183.3 386.8L181.4 386.8L180.1 391.4L180.8 392.7L178.6 393.1L177.7 391.5L175.9 391.5L173.5 392.7L169.3 393.0L166.5 389.7L169.2 387.5L170.2 382.8L169.0 381.3L167.9 382.7L166.5 381.3L166.0 378.2L168.0 377.0L166.9 372.3L169.6 369.3L168.9 369.0L170.7 366.0L174.6 368.0ZM99.0 386.7L98.2 384.1L101.3 379.2L101.7 376.6L100.6 372.6L104.7 373.7L105.3 376.0L109.2 376.5L112.3 375.5L113.7 376.4L112.3 380.9L112.5 385.0L110.1 384.9L108.2 386.9L106.6 385.1L107.3 383.1L104.6 382.3L102.3 383.0L102.9 384.9L101.0 384.5ZM184.6 383.3L183.0 382.0L182.2 378.3L183.5 377.9L183.8 375.9L182.7 374.7L184.1 366.8L183.8 362.5L191.9 361.3L191.1 370.0L199.6 370.3L201.3 371.4L201.3 373.3L195.6 376.6L195.6 378.7L193.3 378.9L192.3 380.6L188.7 380.3ZM142.1 375.3L141.9 379.5L144.7 383.5L143.9 385.6L142.2 384.0L140.0 386.6L140.5 387.8L139.1 390.9L131.6 390.0L130.7 391.8L132.6 392.6L132.8 394.2L128.4 395.8L127.4 397.9L123.3 397.2L121.5 395.6L120.2 397.9L116.8 399.1L113.3 395.1L112.3 389.3L109.0 386.9L110.1 384.9L112.5 385.0L112.3 380.9L113.7 376.4L112.3 375.5L113.8 373.3L117.5 376.8L119.5 376.8L121.5 375.5L122.2 373.0L123.1 373.2L123.7 370.5L128.2 370.0L134.1 371.4L136.3 370.6L139.0 372.0L138.7 373.9ZM251.5 374.2L253.1 373.4L257.0 374.7L258.9 371.0L263.8 370.9L264.2 374.3L262.9 376.5L265.0 376.3L264.4 380.7L265.2 382.1L260.3 384.1L261.1 386.6L262.6 386.4L263.4 387.4L263.1 391.1L261.8 393.7L263.5 393.8L267.0 398.5L270.9 400.6L270.0 404.2L267.8 405.9L264.9 405.4L264.1 403.4L262.1 404.8L258.9 409.2L257.7 413.8L259.3 416.0L257.8 418.9L254.4 419.8L250.1 416.5L251.1 415.6L250.6 411.7L248.9 410.2L250.4 408.9L251.4 403.4L247.8 399.3L246.4 393.9L247.7 392.1L247.3 389.5L248.8 389.0L250.1 390.4L251.7 388.7L250.3 377.0L250.5 374.3ZM180.8 392.7L180.1 391.4L181.4 386.8L183.3 386.8L183.0 385.0L184.6 383.3L188.7 380.3L192.3 380.6L193.3 378.9L195.6 378.7L195.6 376.6L201.3 373.3L202.7 377.4L200.5 380.6L201.3 383.7L202.5 384.1L202.6 387.4L200.7 389.0L199.5 387.5L196.3 388.6L194.4 391.8L192.9 390.3L190.6 391.2L189.2 389.5L183.6 391.2L182.9 392.6ZM241.7 373.4L244.7 374.5L250.5 374.3L251.7 388.7L250.1 390.4L248.8 389.0L247.3 389.5L247.7 392.1L246.4 393.9L247.8 399.3L246.5 398.6L240.7 401.1L238.5 399.4L236.1 399.0L234.8 402.0L232.8 400.4L229.4 400.0L229.8 397.5L227.8 396.8L227.4 395.1L228.7 395.7L230.1 394.1L232.3 393.8L233.0 389.5L229.1 383.9L226.8 384.0L226.9 380.6L233.5 379.3L235.4 376.4L237.5 376.3L238.2 374.8ZM215.4 373.9L215.1 376.5L218.3 378.0L220.3 377.7L225.2 383.0L229.1 383.9L233.0 389.5L232.3 393.8L230.1 394.1L228.7 395.7L221.8 393.8L218.7 393.9L216.1 391.7L208.6 391.8L206.4 394.3L206.7 395.5L211.5 398.0L213.4 397.7L213.6 399.8L211.0 402.1L207.6 400.3L205.3 402.2L202.1 399.8L201.2 397.7L198.1 396.9L196.0 392.9L194.4 391.8L196.3 388.6L199.5 387.5L200.7 389.0L202.6 387.4L202.5 384.1L201.3 383.7L200.5 380.6L202.7 377.4L202.2 376.8L204.9 377.4L209.0 376.3L210.4 374.5L212.3 374.9ZM166.7 377.6L166.0 378.2L166.5 381.3L167.9 382.7L169.0 381.3L170.2 382.8L169.2 387.5L166.5 389.7L169.3 393.0L177.7 391.5L178.6 393.1L177.5 394.3L176.6 400.2L175.2 400.2L171.6 406.3L169.4 405.5L166.0 405.9L165.4 403.6L162.0 404.3L158.9 403.7L159.9 399.5L158.7 398.8L159.9 391.8L159.7 389.5L158.6 389.5L159.7 385.5L162.3 379.4L163.8 379.3L164.7 376.9ZM116.8 399.1L120.2 397.9L121.5 395.6L123.3 397.2L127.4 397.9L128.4 395.8L132.8 394.2L132.6 392.6L130.7 391.8L131.6 390.0L139.1 390.9L137.4 394.7L140.1 394.7L142.9 397.8L149.3 399.1L153.0 401.0L153.8 403.4L156.1 402.8L157.3 403.9L157.1 406.2L154.6 408.0L155.6 410.6L152.6 410.3L152.1 411.7L150.0 410.4L149.6 411.7L144.8 409.5L142.3 414.0L149.4 420.7L150.5 418.5L151.9 418.7L153.0 417.0L158.4 419.4L154.8 423.2L151.2 423.0L148.3 424.1L148.0 425.5L142.5 427.7L142.0 426.6L136.5 423.7L137.1 422.8L128.6 411.8L129.8 409.7L129.5 407.9L125.5 406.8L124.9 404.3L123.5 404.8L118.4 402.6L118.8 401.5ZM178.6 393.1L182.9 392.6L183.6 391.2L189.2 389.5L190.6 391.2L192.9 390.3L196.0 392.9L198.1 396.9L198.3 401.3L196.4 404.8L195.0 404.4L192.2 407.3L191.7 410.6L190.2 410.0L187.9 405.3L186.0 404.4L185.2 400.9L186.5 397.9L184.0 397.2L178.1 393.7ZM198.1 396.9L201.2 397.7L202.1 399.8L205.3 402.2L207.6 400.3L211.0 402.1L213.6 399.8L213.4 397.7L211.5 398.0L206.7 395.5L206.4 394.3L208.6 391.8L216.1 391.7L217.3 394.7L216.8 396.2L215.4 396.7L216.1 401.5L214.0 402.7L213.3 407.1L208.4 404.7L206.7 405.6L206.6 408.4L205.3 410.8L206.2 414.6L208.5 416.1L206.2 417.0L205.1 419.8L203.9 420.1L204.0 422.6L201.0 422.8L199.5 426.8L200.1 428.7L198.2 429.4L197.1 428.4L197.1 426.4L195.1 425.2L194.0 425.7L194.7 422.9L194.1 421.4L191.6 420.6L191.6 419.3L189.4 419.8L189.5 417.8L187.5 414.0L188.9 414.1L190.2 410.0L191.7 410.6L192.2 407.3L195.0 404.4L196.4 404.8L198.3 401.3ZM171.6 406.3L175.2 400.2L176.6 400.2L178.1 393.7L186.5 397.9L185.2 400.9L186.0 404.4L187.9 405.3L190.2 410.0L188.9 414.1L187.5 414.0L188.2 416.4L185.7 416.4L182.5 418.2L179.6 415.4L179.6 413.8L177.3 413.6L174.9 408.9L174.9 407.2ZM120.2 403.7L123.5 404.8L124.9 404.3L125.5 406.8L129.5 407.9L129.8 409.7L128.6 411.8L137.1 422.8L136.5 423.7L142.0 426.6L142.7 429.7L145.0 429.3L146.5 431.0L148.4 430.8L147.1 434.8L148.4 436.7L147.2 437.5L143.2 436.4L138.7 433.8L134.6 434.3L132.4 433.1L131.2 433.8L122.9 431.7L121.9 434.6L117.6 434.7L116.3 432.9L117.5 431.4L112.6 428.4L112.4 423.9L110.6 422.6L111.9 420.4L111.5 417.2L115.0 411.8L115.0 408.5L117.1 407.9ZM158.9 403.7L162.0 404.3L165.4 403.6L166.0 405.9L169.4 405.5L174.9 407.2L174.9 408.9L177.3 413.6L179.6 413.8L179.6 415.4L182.8 419.4L173.0 422.6L168.4 422.3L165.9 420.7L158.4 419.4L153.0 417.0L151.9 418.7L150.5 418.5L149.4 420.7L142.3 414.0L144.8 409.5L149.6 411.7L150.0 410.4L152.1 411.7L152.6 410.3L155.6 410.6L154.6 408.0L157.1 406.2L157.3 403.9L156.1 402.8ZM101.5 412.9L101.8 413.7L100.3 415.6L99.5 414.7L100.3 411.4L99.2 409.6L99.2 406.5L101.7 406.9L103.3 409.0L102.8 412.8ZM188.2 416.4L189.5 417.8L189.4 419.8L191.6 419.3L191.6 420.6L194.1 421.4L194.7 422.9L193.6 426.9L189.8 431.9L187.4 431.0L186.7 437.3L183.6 438.1L181.5 437.2L181.6 435.4L176.2 434.6L173.1 432.8L173.9 427.8L172.9 425.7L173.0 422.6L182.8 419.4L182.5 418.2L185.7 416.4ZM152.7 423.5L154.8 423.2L158.4 419.4L165.9 420.7L168.4 422.3L173.0 422.6L172.9 425.7L173.9 427.8L173.1 432.8L176.2 434.6L181.6 435.4L181.5 437.2L184.0 439.3L182.4 440.3L181.7 442.6L178.3 440.8L177.6 443.1L176.3 443.5L172.9 441.6L170.7 442.2L170.0 440.6L167.3 441.0L167.6 439.3L164.6 438.9L163.8 435.8L165.3 435.2L166.8 436.0L167.6 434.4L165.9 432.5L165.0 427.2L161.0 426.3L157.3 432.3L155.3 430.7L152.5 425.4ZM142.5 427.7L148.0 425.5L148.3 424.1L151.2 423.0L152.7 423.5L152.5 425.4L155.3 430.7L157.3 432.3L161.0 426.3L165.0 427.2L165.9 432.5L167.6 434.4L166.8 436.0L165.3 435.2L163.8 435.8L164.6 438.9L167.6 439.3L167.3 441.0L170.0 440.6L170.7 442.2L172.9 441.6L176.3 443.5L174.3 444.5L175.3 451.0L172.1 449.8L165.8 450.8L165.0 449.0L163.3 449.6L159.5 447.3L158.2 449.4L158.9 451.2L157.0 452.7L155.0 452.1L151.1 453.2L147.7 452.5L144.4 454.4L143.6 453.1L146.8 448.6L144.7 446.5L145.7 445.7L143.3 441.6L139.5 439.8L138.4 438.2L139.9 434.7L147.2 437.5L148.4 436.7L147.1 434.8L148.4 430.8L146.5 431.0L145.0 429.3L142.7 429.7ZM117.6 434.7L121.9 434.6L122.9 431.7L131.2 433.8L132.4 433.1L134.6 434.3L138.7 433.8L139.9 434.7L138.4 438.2L139.5 439.8L143.3 441.6L144.1 443.8L143.1 446.5L142.0 447.8L139.6 447.2L139.4 449.1L130.9 448.6L130.3 451.5L130.9 453.5L130.1 454.9L125.5 456.1L121.4 453.0L118.4 452.2L117.8 449.2L119.5 445.5L118.3 442.5L115.8 441.6L116.5 440.0L115.8 438.2ZM115.8 438.2L116.5 440.0L115.8 441.6L118.3 442.5L119.5 445.5L117.8 449.2L118.7 454.3L121.3 457.6L118.7 462.0L120.9 463.4L121.8 466.2L113.6 469.7L111.1 468.9L109.6 460.6L109.8 455.5L108.4 452.1L108.9 450.8L107.6 449.4L107.5 444.8L104.7 435.8L106.1 435.8L107.4 434.0L109.8 433.7L111.6 437.0L114.4 438.7ZM144.1 443.8L145.7 445.7L144.7 446.5L146.8 448.6L143.6 453.1L144.4 454.4L147.7 452.5L151.1 453.2L155.0 452.1L157.0 452.7L158.9 451.2L159.9 451.9L159.3 454.2L160.3 455.5L160.1 458.6L152.0 459.2L151.7 461.3L149.9 461.7L147.9 459.2L144.0 461.3L144.3 463.3L139.6 464.6L134.9 461.2L133.5 461.7L131.1 460.5L126.3 460.4L121.7 454.7L118.7 454.3L118.4 452.2L121.4 453.0L125.5 456.1L130.1 454.9L130.9 453.5L130.3 451.5L130.9 448.6L139.4 449.1L139.6 447.2L142.0 447.8ZM121.8 466.2L120.9 463.4L118.7 462.0L121.3 457.6L118.7 454.3L121.7 454.7L126.3 460.4L131.1 460.5L133.5 461.7L134.9 461.2L139.6 464.6L138.2 467.7L137.0 468.3L136.4 466.5L131.6 468.3L131.6 473.5L132.8 475.0L135.1 475.3L135.1 477.6L132.8 479.2L134.7 480.1L132.2 486.4L129.9 486.8L127.2 485.0L126.9 480.8L124.9 480.1L124.7 478.9L123.0 477.9L123.6 475.4L121.3 473.1L122.5 470.7L121.3 468.4ZM111.1 468.9L113.6 469.7L121.8 466.2L121.3 468.4L122.5 470.7L121.3 473.1L123.6 475.4L123.0 477.9L124.7 478.9L124.9 480.1L126.9 480.8L127.2 485.0L129.3 486.0L127.5 487.4L127.4 488.9L124.7 489.8L123.6 486.9L118.2 487.2L116.2 483.0L114.5 482.4ZM127.4 488.9L129.9 488.5L130.8 491.9L128.7 493.8L126.1 494.5L124.1 491.5L123.5 493.9L119.9 492.5L118.2 487.2L123.6 486.9L124.7 489.8ZM123.5 493.9L124.1 491.5L126.1 494.5L130.2 492.6L130.7 494.9L131.7 495.4L131.8 497.3L130.4 498.3L131.6 499.5L130.7 501.6L131.2 502.9L129.2 505.5L126.7 506.1L124.5 502.5L122.4 495.7ZM114.0 396.0L118.8 401.5L118.4 402.6L120.2 403.7L117.1 407.9L115.0 408.5L114.9 409.6L111.2 409.5L109.0 411.1L107.4 410.0L104.7 410.3L104.7 412.1L102.8 412.8L103.3 409.0L101.7 406.9L99.2 406.5L99.4 405.6L101.6 405.9L102.8 403.4L108.2 399.7L109.8 396.7ZM69.7 344.5L71.3 345.7L71.0 347.0L73.9 351.2L71.5 355.9L74.4 356.8L74.6 359.4L72.6 360.3L72.5 362.6L74.6 364.3L75.7 367.0L65.7 371.5L65.2 369.6L66.0 367.9L64.0 367.7L64.2 365.3L61.5 363.1L58.3 363.1L57.2 361.5L60.2 359.8L59.8 356.3L57.8 356.4L58.5 353.2L57.0 351.7L57.2 350.0L60.9 350.3L62.4 349.6L62.0 346.9L64.1 345.2L68.4 343.3ZM123.7 370.5L123.1 373.2L122.2 373.0L121.5 375.5L119.5 376.8L117.5 376.8L113.8 373.3L114.1 368.8L115.7 366.6L119.6 366.7L122.5 368.4ZM102.9 384.9L102.3 383.0L104.6 382.3L107.3 383.1L106.6 385.1L108.2 386.9L106.6 387.7L103.0 386.8ZM108.2 386.9L112.3 389.3L114.0 396.0L109.8 396.7L108.2 399.7L102.8 403.4L101.6 405.9L99.4 405.6L96.4 393.4L96.6 391.2L98.1 390.3L97.5 388.0L101.0 384.5L102.9 384.9L103.0 386.8L106.6 387.7ZM115.6 336.5L114.4 339.5L114.4 344.9L108.1 348.1L105.0 346.8L104.8 344.6L101.7 342.5L100.9 339.2L102.6 337.4L105.0 336.9L107.0 330.6L109.4 328.8L109.5 326.2L111.1 326.0L112.6 329.0L112.6 334.3L113.6 335.8ZM100.9 339.2L101.7 342.5L104.8 344.6L105.0 346.8L108.1 348.1L110.2 346.6L114.3 351.3L113.3 354.8L109.0 355.6L102.5 355.2L101.7 357.1L99.2 357.3L95.5 353.6L97.4 352.1L94.0 350.1L95.7 348.4L93.3 345.3L94.6 339.9L96.3 339.1L98.5 340.1ZM36.3 348.1L36.4 350.0L40.0 348.4L43.3 351.7L42.8 351.3L40.0 353.7L40.1 356.0L39.0 357.8L39.8 361.7L26.9 347.9L27.4 346.6L28.5 347.4L30.9 345.4L33.1 345.7L32.9 347.8ZM110.2 346.6L114.4 344.9L115.0 342.9L117.4 342.4L118.6 344.6L121.8 345.2L123.6 346.9L121.0 347.7L121.2 349.5L122.4 350.6L120.9 352.2L121.4 355.0L124.5 355.0L124.3 356.3L118.5 357.8L118.0 356.1L114.2 355.8L113.3 354.8L114.3 351.3ZM121.7 366.9L117.8 363.3L121.0 362.9L121.1 360.7L123.7 360.0L124.0 357.9L125.9 356.7L131.7 355.5L130.7 354.1L128.4 353.9L121.4 355.0L120.9 352.2L122.4 350.6L121.2 349.5L121.0 347.7L128.0 345.2L130.9 345.6L134.1 343.4L136.0 346.1L135.6 350.3L137.2 351.7L141.3 352.8L140.8 355.9L139.5 357.5L135.5 357.4L135.1 359.8L132.0 361.7L127.4 362.0L123.8 365.0L123.8 366.1ZM181.3 360.5L180.4 359.3L183.8 356.3L183.8 354.1L185.3 352.9L186.3 353.2L188.4 351.1L192.8 351.1L194.3 349.6L198.5 349.3L199.9 350.8L200.6 354.8L198.0 354.2L198.7 358.3L205.4 357.5L207.1 358.1L216.8 353.4L218.5 353.8L219.0 357.7L214.9 359.3L211.2 358.8L210.7 361.1L212.6 363.7L212.6 367.9L215.0 369.5L216.2 372.9L215.4 373.9L212.3 374.9L210.4 374.5L209.0 376.3L204.9 377.4L202.2 376.8L201.3 371.4L199.6 370.3L191.1 370.0L191.9 361.3L183.8 362.5L182.5 360.3ZM98.1 356.0L99.2 357.3L101.7 357.1L102.5 355.2L109.0 355.6L113.3 354.8L114.2 355.8L118.0 356.1L118.5 357.8L113.3 358.1L111.7 361.7L107.9 365.1L108.2 367.8L110.4 369.6L110.1 371.1L106.4 369.2L102.7 365.3L101.8 365.9L95.7 364.6L96.0 362.1L95.2 359.6ZM150.5 357.5L148.6 359.0L148.8 360.5L146.8 362.5L143.8 363.5L143.1 365.5L145.3 369.1L145.8 372.9L144.6 374.7L142.1 375.3L138.7 373.9L139.0 372.0L136.3 370.6L134.1 371.4L128.2 370.0L123.7 370.5L121.7 366.9L123.8 366.1L123.8 365.0L127.4 362.0L132.0 361.7L135.1 359.8L135.5 357.4L139.5 357.5L140.8 355.9L141.3 352.8L143.2 352.4L147.2 353.9L148.3 356.3ZM120.8 366.7L122.5 368.4L119.6 366.7L115.7 366.6L114.1 368.8L114.2 371.0L111.8 370.9L109.8 368.4L108.2 367.8L107.9 365.1L111.7 361.7L113.3 358.1L116.2 358.5L120.7 357.5L124.3 356.3L124.5 354.9L130.7 354.1L131.7 355.5L125.9 356.7L124.0 357.9L123.7 360.0L121.1 360.7L121.0 362.9L117.8 363.3ZM242.1 354.5L240.8 355.7L240.8 357.1L242.9 360.6L242.4 363.0L241.2 363.7L243.4 365.2L244.6 367.7L241.7 373.4L238.2 374.8L237.5 376.3L235.4 376.4L233.8 374.2L229.5 372.2L219.0 361.5L216.6 361.2L214.9 359.3L219.0 357.7L218.7 355.5L220.4 354.8L221.9 355.8L228.6 355.4L229.0 353.4L235.5 352.1L235.7 350.5L241.2 351.5ZM73.9 351.2L71.0 347.0L71.3 345.7L69.7 344.5L72.0 342.8L70.6 338.1L72.6 336.2L79.6 335.6L79.1 338.0L83.8 338.4L85.3 339.5L83.8 341.7L80.9 341.1L79.7 342.7L76.9 343.8L75.8 347.0L74.5 347.8ZM115.0 342.9L114.4 339.5L115.6 336.5L117.6 336.4L119.2 337.5L119.2 332.3L120.7 331.5L122.3 332.6L125.2 331.3L126.1 331.5L127.4 334.5L129.5 334.0L128.9 336.7L126.4 335.9L126.0 337.2L127.5 339.2L126.9 343.7L128.0 345.2L123.6 346.9L121.8 345.2L118.6 344.6L117.4 342.4ZM58.3 363.1L61.5 363.1L64.2 365.3L64.0 367.7L66.0 367.9L65.2 369.6L65.7 371.5L62.4 373.0L59.7 372.8L57.7 373.9L49.7 370.3L45.3 367.0L48.7 365.7L52.8 360.5L56.0 361.4L56.0 363.8ZM56.7 351.8L49.5 351.8L46.6 353.4L43.3 351.7L40.9 348.8L44.8 346.2L47.2 342.9L51.8 343.0L52.5 338.6L51.2 338.5L50.2 337.0L50.1 333.3L52.6 332.9L51.9 331.3L52.5 330.3L54.9 334.1L57.0 333.2L61.0 334.6L64.7 340.3L66.4 337.7L70.6 338.1L71.7 339.8L72.0 342.8L69.7 344.5L68.4 343.3L64.1 345.2L62.0 346.9L62.4 349.6L57.2 350.0ZM90.1 348.7L88.7 344.8L90.5 343.4L91.8 345.7ZM75.7 367.0L74.6 364.3L72.5 362.6L72.6 360.3L74.6 359.4L74.4 356.8L71.5 355.9L72.9 354.1L74.5 347.8L75.8 347.0L76.9 343.8L79.7 342.7L80.9 341.1L83.8 341.7L86.1 344.2L87.7 344.2L86.9 346.5L87.7 348.1L86.7 349.8L89.1 352.6L84.8 360.0L84.9 362.3ZM55.9 352.2L57.0 351.7L58.5 353.2L57.8 356.4L59.8 356.3L60.2 359.8L57.2 361.5L58.3 363.1L56.0 363.8L56.0 361.4L52.8 360.5L48.7 365.7L45.3 367.0L39.8 361.7L39.0 357.8L40.1 356.0L40.0 353.7L42.8 351.3L46.6 353.4L49.5 351.8ZM41.8 350.0L40.0 348.4L36.4 350.0L36.3 348.1L38.3 347.7L38.4 345.7L34.5 337.0L35.0 334.5L39.7 331.7L42.9 331.4L44.9 329.0L44.8 327.7L48.7 323.4L49.7 324.8L50.7 322.8L52.5 323.1L53.9 324.4L53.9 327.9L51.9 331.3L52.6 332.9L50.1 333.3L50.2 337.0L51.2 338.5L52.5 338.6L52.5 341.8L51.8 343.0L47.2 342.9L44.8 346.2L40.9 348.8ZM36.3 348.1L32.9 347.8L33.1 345.7L30.9 345.4L28.5 347.4L27.4 346.6L26.5 347.7L18.3 338.7L17.7 337.0L19.2 334.0L22.5 336.1L22.8 339.2L26.3 336.9L35.0 334.5L34.5 337.0L38.4 345.7L38.3 347.7ZM144.7 383.5L145.9 382.0L148.7 381.7L148.3 380.3L150.9 378.4L151.5 376.6L154.3 376.5L154.4 377.8L157.2 377.0L157.4 375.9L162.0 375.7L163.5 377.0L165.8 376.6L167.4 374.9L168.0 377.0L164.7 376.9L163.8 379.3L162.3 379.4L159.7 385.5L158.6 389.5L159.7 389.5L159.9 391.8L158.7 398.8L159.9 399.5L158.9 403.7L156.1 402.8L153.8 403.4L153.0 401.0L149.3 399.1L142.9 397.8L140.1 394.7L137.4 394.7L140.5 387.8L140.0 386.6L142.2 384.0L143.9 385.6ZM114.9 409.6L115.0 411.8L111.5 417.2L111.9 420.4L110.6 422.6L112.4 423.9L112.6 428.4L117.5 431.4L116.3 432.9L117.6 434.7L115.8 438.2L114.4 438.7L111.6 437.0L109.8 433.7L107.4 434.0L106.1 435.8L104.7 435.8L103.1 432.5L103.1 429.8L104.3 429.3L101.9 427.2L101.5 424.8L102.3 423.0L100.7 419.4L100.3 415.6L101.8 413.7L101.5 412.9L104.7 412.1L104.7 410.3L107.4 410.0L109.0 411.1L111.2 409.5ZM112.3 375.5L109.2 376.5L105.3 376.0L104.7 373.7L100.6 372.6L97.8 365.8L102.7 365.3L106.4 369.2L110.1 371.1L110.4 369.6L111.8 370.9L114.2 371.0ZM101.3 379.2L100.1 381.3L100.3 379.3ZM59.7 372.8L61.4 372.8L60.9 373.7ZM67.5 284.4L71.5 282.9L75.3 283.8L83.8 282.2L86.0 284.2L89.6 284.0L93.1 288.9L93.6 287.0L97.0 287.8L98.0 290.0L102.8 290.2L104.3 292.2L102.1 297.5L100.9 297.8L99.1 296.1L96.4 296.5L96.3 298.6L94.9 299.9L90.4 299.1L89.7 297.6L87.3 296.8L86.9 295.8L81.7 298.6L79.6 301.9L76.5 301.4L75.9 299.1L68.1 298.8L67.6 291.5L69.3 290.8L68.1 289.2L68.8 286.6L66.9 286.2ZM68.1 298.8L75.9 299.1L76.5 301.4L79.6 301.9L81.7 298.6L86.9 295.8L87.3 296.8L89.7 297.6L90.4 299.1L92.8 299.9L90.6 301.4L89.5 301.1L88.7 307.1L86.4 306.6L84.1 307.9L82.8 306.9L81.9 309.2L77.7 310.5L72.7 310.3L69.3 306.4L66.1 306.2L66.0 304.6L63.1 307.6L62.7 302.1L63.5 300.6ZM92.8 299.9L94.9 299.9L96.3 298.6L96.4 296.5L99.1 296.1L100.9 297.8L99.3 302.2L100.2 308.8L99.4 309.8L95.8 308.5L93.2 312.5L93.6 313.2L91.3 314.8L91.6 318.0L88.5 319.4L86.5 318.0L86.2 315.0L88.9 311.8L83.7 311.1L81.9 309.2L82.8 306.9L84.1 307.9L86.4 306.6L88.7 307.1L89.5 301.1L90.6 301.4ZM89.7 318.8L91.6 318.0L91.3 314.8L93.6 313.2L93.2 312.5L95.8 308.5L99.4 309.8L99.4 311.5L98.1 312.0L102.6 315.5L104.1 314.9L103.3 316.2L103.9 318.9L100.3 321.3L98.8 319.0L92.8 318.0L92.1 319.9ZM83.8 341.7L85.3 339.5L83.8 338.4L79.1 338.0L81.2 333.7L80.7 332.0L83.3 332.5L84.3 334.1L86.0 333.1L85.7 330.7L84.0 330.6L83.7 325.9L82.3 324.4L82.9 319.9L80.8 319.7L80.6 317.2L82.3 316.4L79.9 314.9L79.8 312.5L81.6 312.4L82.3 310.1L83.7 311.1L88.9 311.8L86.2 315.0L86.5 318.0L88.5 319.4L89.7 318.8L92.1 319.9L92.8 318.0L95.4 318.3L98.8 319.0L100.3 321.3L99.3 323.3L96.2 323.6L93.6 325.9L93.3 330.1L91.6 331.0L91.8 332.4L90.5 333.3L90.5 336.9L87.7 344.2L86.1 344.2ZM115.6 336.5L113.6 335.8L112.6 334.3L112.6 329.0L111.1 326.0L112.9 321.2L116.1 321.5L118.0 318.9L119.4 321.0L120.8 318.7L123.1 320.9L120.7 331.5L119.2 332.3L119.2 337.5L117.6 336.4ZM132.9 329.6L130.5 329.4L128.1 332.2L125.2 331.3L122.3 332.6L120.7 331.5L124.5 313.5L127.8 315.2L127.6 317.2L130.1 317.0L133.2 323.6L134.6 323.3L134.7 325.4ZM111.7 300.6L111.5 302.6L114.8 307.4L115.8 306.3L117.6 307.1L117.1 311.8L114.8 314.7L113.3 314.8L110.9 316.7L111.7 317.3L109.5 319.9L108.0 317.9L106.3 317.0L104.0 317.6L103.3 316.2L105.9 313.8L108.2 307.3L106.4 305.1L106.7 303.5L108.5 301.7L109.8 302.2ZM117.0 311.0L121.3 311.5L122.6 313.6L124.0 312.6L124.8 314.6L124.0 318.6L122.8 320.1L120.8 318.7L119.4 321.0L118.0 318.9L116.1 321.5L112.9 321.2L112.7 322.5L111.1 323.0L108.8 320.8L110.6 319.4L111.7 317.3L110.9 316.7L113.3 314.8L114.8 314.7ZM102.8 290.2L103.8 287.5L105.6 287.1L105.3 289.4L106.9 290.3L105.1 294.0L108.3 296.6L108.9 298.4L110.4 297.2L112.3 297.5L112.5 300.1L109.8 302.2L108.5 301.7L106.7 303.5L106.4 305.1L108.2 307.3L105.9 313.8L102.6 315.5L98.1 312.0L99.4 311.5L100.2 308.8L99.3 302.2L99.9 299.6L102.1 297.5L104.3 292.2ZM63.7 283.0L67.5 284.4L66.9 286.2L68.8 286.6L68.1 289.2L69.3 290.8L67.6 291.5L68.1 298.8L63.5 300.6L62.7 302.1L63.1 307.6L66.0 304.6L66.1 306.2L69.3 306.4L72.7 310.3L76.0 317.1L75.6 318.0L69.6 317.9L64.0 316.5L54.8 318.5L55.4 316.9L46.8 316.3L41.9 323.1L32.8 325.9L27.9 325.5L21.8 324.1L11.3 317.4L9.1 314.8L9.5 313.1L6.3 311.6L6.7 310.1L4.8 309.9L4.4 308.0L0.9 307.5L0.0 305.7L1.5 305.3L2.3 301.3L5.1 299.2L13.8 299.1L14.0 291.7L15.8 292.9L17.9 291.5L19.1 293.3L21.0 292.2L23.1 293.0L25.7 292.0L31.6 291.7L34.5 294.5L40.8 294.5L42.5 291.7L52.3 288.8L52.3 292.7L57.3 293.4L61.1 290.1L62.4 290.3L63.5 288.5L61.3 288.3L61.1 284.8ZM79.6 335.6L72.6 336.2L70.6 338.1L66.4 337.7L64.7 340.3L61.0 334.6L63.1 334.9L64.8 332.5L63.4 331.2L64.8 328.5L64.0 325.9L66.8 324.9L68.7 319.0L67.9 317.3L75.6 318.0L76.0 317.1L72.7 310.3L79.6 310.5L81.9 309.2L81.6 312.4L79.8 312.5L79.9 314.9L82.3 316.4L80.6 317.2L80.8 319.7L82.9 319.9L82.3 324.4L83.7 325.9L84.0 330.6L85.7 330.7L86.0 333.1L84.3 334.1L83.3 332.5L80.7 332.0L81.2 333.7ZM111.1 326.0L109.5 326.2L109.4 328.8L108.4 329.7L105.8 327.4L103.8 327.6L102.0 330.5L98.2 329.3L96.4 332.6L94.8 332.5L93.3 330.1L93.6 325.9L96.2 323.6L99.3 323.3L102.1 319.4L103.9 318.9L104.0 317.6L106.3 317.0L108.0 317.9L109.5 319.9L108.8 320.8L111.1 323.0L112.7 322.5ZM108.4 329.7L107.0 330.6L105.0 336.9L96.5 337.7L93.5 336.7L92.3 338.2L90.5 336.9L90.5 333.3L91.8 332.4L91.6 331.0L93.3 330.1L94.8 332.5L96.4 332.6L98.2 329.3L102.0 330.5L103.8 327.6L105.8 327.4ZM61.0 334.6L57.0 333.2L54.9 334.1L52.5 330.3L53.9 327.9L53.9 324.4L50.7 322.8L52.1 320.6L57.5 317.5L60.6 317.7L64.0 316.5L67.9 317.3L68.7 319.0L66.8 324.9L64.0 325.9L64.8 328.5L63.4 331.2L64.8 332.5L63.1 334.9ZM247.8 399.3L246.5 398.6L240.7 401.1L238.5 399.4L236.1 399.0L234.8 402.0L232.8 400.4L229.4 400.0L229.8 397.5L227.8 396.8L227.4 395.1L218.7 393.9L216.1 391.7L217.3 394.7L216.8 396.2L215.4 396.7L216.1 401.5L214.0 402.7L213.3 407.1L208.4 404.7L206.7 405.6L206.6 408.4L205.3 410.8L206.2 414.6L208.5 416.1L206.2 417.0L205.1 419.8L203.9 420.1L204.0 422.6L201.0 422.8L199.5 426.8L200.1 428.7L198.2 429.4L197.1 428.4L197.1 426.4L195.1 425.2L189.8 431.9L187.4 431.0L186.7 437.3L182.9 437.6L184.0 439.3L182.4 440.3L181.7 442.6L178.3 440.8L177.6 443.1L174.3 444.5L175.3 451.0L172.1 449.8L165.8 450.8L165.0 449.0L163.3 449.6L159.5 447.3L158.2 449.4L159.9 451.9L159.3 454.2L160.3 455.5L160.1 458.6L152.0 459.2L151.7 461.3L149.9 461.7L147.9 459.2L144.0 461.3L144.3 463.3L139.6 464.6L138.2 467.7L137.0 468.3L136.4 466.5L131.6 468.3L131.6 473.5L132.8 475.0L135.1 475.3L135.1 477.6L132.8 479.2L134.7 480.1L132.2 486.4L129.9 486.8L129.3 486.0L127.5 487.4L127.4 488.9L124.7 489.8L123.6 486.9L118.2 487.2L116.2 483.0L114.5 482.4L111.5 472.7L109.6 460.6L109.8 455.5L108.4 452.1L108.9 450.8L107.6 449.4L107.5 444.8L103.1 432.5L103.1 429.8L104.3 429.3L101.9 427.2L101.5 424.8L102.3 423.0L99.5 414.7L100.3 411.4L99.2 409.6L99.4 405.6L96.4 393.4L96.6 391.2L98.1 390.3L97.5 388.0L101.0 384.5L102.9 384.9L103.0 386.8L106.6 387.7L109.0 386.9L110.1 384.9L112.5 385.0L112.3 380.9L113.7 376.4L112.3 375.5L113.8 373.3L117.5 376.8L119.5 376.8L121.5 375.5L122.2 373.0L123.1 373.2L123.7 370.5L122.5 368.4L117.8 363.3L121.0 362.9L121.1 360.7L123.7 360.0L124.0 357.9L125.9 356.7L131.7 355.5L130.7 354.1L128.4 353.9L121.4 355.0L120.9 352.2L122.4 350.6L121.2 349.5L121.0 347.7L128.0 345.2L130.9 345.6L134.1 343.4L136.0 346.1L135.6 350.3L137.2 351.7L141.3 352.8L143.2 352.4L147.2 353.9L148.3 356.3L150.5 357.5L169.2 358.4L170.6 361.1L170.8 365.1L175.3 365.1L177.6 362.5L180.4 362.7L181.3 360.5L180.4 359.3L183.8 356.3L183.8 354.1L185.3 352.9L186.3 353.2L188.4 351.1L192.8 351.1L194.3 349.6L198.5 349.3L199.9 350.8L200.6 354.8L198.0 354.2L198.7 358.3L205.4 357.5L207.1 358.1L216.8 353.4L218.5 353.8L218.7 355.5L220.4 354.8L221.9 355.8L228.6 355.4L229.0 353.4L235.5 352.1L235.7 350.5L241.2 351.5L242.1 354.5L246.4 353.1L250.2 354.9L253.3 354.2L254.5 352.9L257.6 352.8L260.9 356.8L265.1 359.5L265.9 359.3L264.9 361.7L261.7 362.9L261.2 364.7L261.7 368.5L263.5 368.4L264.2 374.3L262.9 376.5L265.0 376.3L264.4 380.7L265.2 382.1L260.3 384.1L261.1 386.6L262.6 386.4L263.4 387.4L263.1 391.1L261.8 393.7L263.5 393.8L267.0 398.5L270.9 400.6L270.0 404.2L267.8 405.9L264.9 405.4L264.1 403.4L262.1 404.8L258.9 409.2L257.7 413.8L259.3 416.0L257.8 418.9L254.4 419.8L250.1 416.5L251.1 415.6L250.6 411.7L248.9 410.2L250.4 408.9L251.4 403.4ZM98.1 356.0L95.5 353.6L97.4 352.1L94.0 350.1L95.7 348.4L93.3 345.3L94.6 339.9L96.3 339.1L98.5 340.1L102.6 337.4L96.5 337.7L93.5 336.7L92.3 338.2L90.5 336.9L86.9 346.5L87.7 348.1L86.7 349.8L89.1 352.6L84.8 360.0L84.9 362.3L72.2 368.1L70.6 369.9L62.4 373.0L59.7 372.8L57.7 373.9L49.7 370.3L42.1 364.3L34.0 354.7L26.9 347.9L27.4 346.6L26.5 347.7L23.6 345.0L17.7 337.0L19.2 334.0L22.5 336.1L22.8 339.2L26.3 336.9L37.3 333.5L39.7 331.7L42.9 331.4L44.9 329.0L44.8 327.7L48.7 323.4L49.7 324.8L55.4 316.9L46.8 316.3L41.9 323.1L32.8 325.9L21.8 324.1L11.3 317.4L9.1 314.8L9.5 313.1L6.3 311.6L6.7 310.1L4.8 309.9L4.4 308.0L0.9 307.5L0.0 305.7L1.5 305.3L2.3 301.3L5.1 299.2L13.8 299.1L14.0 291.7L15.8 292.9L17.9 291.5L19.1 293.3L21.0 292.2L23.1 293.0L25.7 292.0L31.6 291.7L34.5 294.5L40.8 294.5L42.5 291.7L52.3 288.8L52.3 292.7L57.3 293.4L61.1 290.1L62.4 290.3L63.5 288.5L61.3 288.3L61.1 284.8L63.7 283.0L67.5 284.4L71.5 282.9L75.3 283.8L83.8 282.2L86.0 284.2L89.6 284.0L93.1 288.9L93.6 287.0L97.0 287.8L98.0 290.0L102.8 290.2L103.8 287.5L105.6 287.1L105.3 289.4L106.9 290.3L105.1 294.0L108.3 296.6L108.9 298.4L110.4 297.2L112.3 297.5L112.5 300.1L111.5 302.6L114.8 307.4L115.8 306.3L117.6 307.1L117.0 311.0L119.7 310.9L121.3 311.5L122.6 313.6L124.0 312.6L127.8 315.2L127.6 317.2L130.1 317.0L133.2 323.6L134.6 323.3L134.7 325.4L132.9 329.6L130.5 329.4L128.1 332.2L126.1 331.5L127.4 334.5L129.5 334.0L128.9 336.7L126.4 335.9L126.0 337.2L127.5 339.2L126.9 343.7L128.0 345.2L121.0 347.7L121.2 349.5L122.4 350.6L120.9 352.2L121.4 355.0L128.4 353.9L130.7 354.1L131.7 355.5L125.9 356.7L124.0 357.9L123.7 360.0L121.1 360.7L121.0 362.9L117.8 363.3L122.5 368.4L123.7 370.5L123.1 373.2L122.2 373.0L121.5 375.5L119.5 376.8L117.5 376.8L113.8 373.3L112.3 375.5L113.7 376.4L112.3 380.9L112.5 385.0L110.1 384.9L108.2 386.9L106.6 385.1L107.3 383.1L104.6 382.3L102.3 383.0L102.9 384.9L101.0 384.5L99.0 386.7L98.2 384.1L101.3 379.2L101.7 376.6L97.8 365.8L98.4 365.4L95.7 364.6L96.0 362.1L95.2 359.6ZM90.1 348.7L88.7 344.8L90.5 343.4L91.8 345.7ZM123.5 493.9L119.9 492.5L118.2 487.2L123.6 486.9L124.7 489.8L129.9 488.5L131.8 497.3L130.4 498.3L131.6 499.5L130.7 501.6L131.2 502.9L129.2 505.5L126.7 506.1L124.5 502.5L122.4 495.7ZM102.9 384.9L102.3 383.0L104.6 382.3L107.3 383.1L106.6 385.1L108.2 386.9L106.6 387.7L103.0 386.8ZM101.3 379.2L100.1 381.3L100.3 379.3ZM59.7 372.8L61.4 372.8L60.9 373.7Z","Central":"M183.8 354.1L183.8 356.3L180.4 359.3L181.3 360.5L180.4 362.7L177.6 362.5L175.3 365.1L170.8 365.1L170.6 361.1L169.2 358.4L166.9 358.5L167.4 356.2L166.2 354.4L168.9 353.8L171.0 354.5L175.2 353.2ZM321.1 358.3L319.5 363.7L317.1 364.3L314.7 362.9L307.8 363.3L307.8 364.9L304.0 370.8L294.0 365.0L294.8 361.6L298.3 358.2L302.5 360.9L304.2 359.4L309.0 359.3L308.9 355.8L312.8 354.2L316.7 356.9ZM266.1 396.5L263.8 395.2L263.5 393.8L261.8 393.7L264.3 385.7L265.7 384.5L267.0 385.3L267.8 383.0L269.5 382.6L270.1 380.9L272.9 379.2L276.2 381.1L280.0 376.8L287.7 378.7L289.9 382.2L289.5 384.4L290.3 385.9L287.5 386.7L285.5 385.2L282.6 386.3L281.3 388.6L279.2 389.2L281.7 392.4L281.0 394.7L276.4 395.3L274.9 396.4L269.8 390.9ZM270.0 404.2L270.9 400.6L267.0 398.5L266.1 396.5L269.8 390.9L274.9 396.4L276.4 395.3L281.0 394.7L281.0 395.8L283.4 396.8L282.1 401.6L283.2 401.7L283.0 403.9L275.8 406.7L272.1 407.1L272.1 405.4ZM275.8 406.7L279.4 404.9L281.0 405.5L280.3 407.2L281.6 408.1L283.7 407.9L283.5 410.6L284.3 411.7L282.9 413.7L284.8 414.5L286.4 416.8L288.1 416.2L289.5 417.7L287.3 420.4L286.4 418.6L285.1 418.4L284.6 421.5L283.4 421.1L280.9 425.5L276.9 423.0L276.5 419.6L278.1 413.3L277.5 408.3L276.1 408.4ZM294.8 400.9L296.4 400.7L296.9 402.7L298.2 402.7L297.7 407.9L299.4 414.4L298.0 414.7L297.6 417.3L296.2 418.5L289.3 421.8L287.3 420.4L289.5 417.7L288.1 416.2L286.4 416.8L284.8 414.5L282.9 413.7L284.3 411.7L283.5 410.6L283.7 407.9L281.6 408.1L280.3 407.2L281.0 405.5L284.4 407.4L287.2 404.2L289.9 404.5L289.8 402.2L291.4 399.9ZM281.0 394.7L281.7 392.4L279.2 389.2L281.3 388.6L282.6 386.3L285.5 385.2L287.5 386.7L290.3 385.9L291.6 388.5L291.5 391.9L293.6 392.9L293.9 394.2L295.6 394.5L294.8 400.9L291.4 399.9L289.8 402.2L289.9 404.5L287.2 404.2L284.4 407.4L280.3 405.2L283.0 403.9L283.2 401.7L282.1 401.6L283.4 396.8L281.0 395.8ZM294.0 365.0L297.7 366.5L298.4 368.3L301.4 369.7L302.2 375.3L301.3 376.9L303.5 379.8L302.7 388.4L309.4 389.9L309.0 393.6L306.8 394.8L305.5 392.5L303.7 391.7L301.6 393.4L299.9 390.3L294.6 388.6L296.4 387.3L298.1 383.8L296.5 383.0L291.4 367.7ZM298.3 358.2L294.8 361.6L294.0 365.0L291.4 367.7L290.3 368.1L288.9 366.6L286.0 366.4L285.4 364.1L286.3 363.2L284.6 361.2L287.8 353.2L292.2 353.2L292.8 355.6ZM284.8 359.3L284.6 361.2L286.3 363.2L285.4 364.1L285.6 368.2L281.9 370.3L282.0 366.7L276.4 365.1L277.5 364.5L277.3 361.8L276.2 362.0L277.4 355.1L280.6 354.5ZM290.3 385.9L289.5 384.4L289.9 382.2L287.7 378.7L282.7 378.0L281.8 377.1L283.6 375.8L284.0 373.4L283.1 369.8L285.6 368.2L286.0 366.4L288.9 366.6L290.3 368.1L291.4 367.7L296.5 383.0L298.1 383.8L296.4 387.3L294.6 388.6L293.1 387.3L291.6 388.5ZM281.8 377.1L280.0 376.8L276.2 381.1L272.9 379.2L272.2 375.3L269.2 374.0L269.2 372.3L271.4 371.8L272.1 366.8L275.6 366.8L276.4 365.1L282.0 366.7L281.9 370.3L283.1 369.8L284.0 373.4L283.6 375.8ZM251.1 324.6L253.3 326.7L257.8 326.2L258.0 327.5L255.2 330.4L252.7 330.2L254.0 336.4L254.0 338.4L252.6 339.9L253.3 342.3L250.5 342.5L250.3 344.6L249.0 344.8L247.0 348.4L245.4 348.6L245.0 351.0L241.5 353.3L241.2 351.5L235.7 350.5L237.9 345.1L239.2 344.1L236.8 336.6L238.0 333.0L237.7 330.0L239.5 327.6L241.1 327.4L244.1 322.4L246.6 323.0L247.4 324.7ZM140.6 331.3L138.6 337.8L136.0 338.7L134.9 341.5L135.5 343.0L130.9 345.6L128.0 345.2L126.9 343.7L127.5 339.2L126.0 337.2L126.4 335.9L128.9 336.7L129.5 334.0L127.4 334.5L126.1 331.5L128.1 332.2L130.5 329.4L132.9 329.6L134.2 330.8ZM237.7 330.0L238.0 333.0L236.8 336.6L239.2 344.1L237.9 345.1L235.5 352.1L230.2 352.8L229.0 353.4L228.6 355.4L221.9 355.8L220.4 354.8L218.7 355.5L218.5 353.8L216.8 353.4L219.3 350.5L219.5 348.3L221.2 346.9L220.5 345.5L215.4 342.0L215.3 334.9L217.6 334.5L219.5 335.5L221.4 334.5L221.2 331.8L220.1 330.5L220.6 328.9L222.5 329.2L224.2 328.0L224.9 329.3L229.0 328.7L231.0 329.4L231.6 326.7L235.8 325.7L235.9 327.6ZM196.7 339.0L198.4 340.3L198.1 342.1L199.1 343.4L194.3 344.9L193.1 346.2L191.3 344.2L186.0 344.4L184.7 340.5L185.7 337.6L183.8 336.2L184.0 334.3L193.1 331.1ZM168.9 353.8L166.2 354.4L167.4 356.2L166.9 358.5L158.9 358.0L158.4 356.0L156.6 355.8L155.9 354.2L156.7 352.9L154.4 351.3L152.2 352.4L150.7 351.0L151.9 350.0L154.8 342.7L152.8 341.0L156.5 340.5L157.9 335.5L159.1 336.3L167.1 334.6L169.0 331.8L169.4 333.3L171.7 334.5L171.7 336.6L169.0 339.8L172.0 342.4L170.2 344.6L170.8 346.7L169.4 347.4ZM283.1 332.9L284.8 336.2L282.7 346.0L276.7 350.8L273.1 350.3L271.6 348.1L269.8 347.8L269.5 346.4L271.1 341.5L273.2 342.7L274.0 338.6L275.6 337.4L275.4 334.1L277.2 332.9L278.3 333.6L280.0 332.1ZM183.8 354.1L175.2 353.2L171.0 354.5L168.9 353.8L169.4 347.4L170.8 346.7L170.2 344.6L172.0 342.4L169.0 339.8L171.7 336.6L174.1 337.2L177.8 334.7L180.0 334.8L181.7 337.3L183.8 336.2L185.7 337.6L184.7 340.5L186.0 344.4L191.3 344.2L193.1 346.2L191.9 347.1L191.6 349.3L186.3 353.2L185.3 352.9ZM265.9 359.3L265.1 359.5L260.9 356.8L257.6 352.8L254.5 352.9L253.3 354.2L250.2 354.9L246.4 353.1L242.1 354.5L241.5 353.3L245.0 351.0L245.4 348.6L247.0 348.4L249.0 344.8L250.3 344.6L250.5 342.5L253.3 342.3L252.6 339.9L254.0 338.4L254.0 336.4L260.4 337.1L266.8 339.4L266.8 337.9L270.2 335.4L271.3 337.7L273.2 337.6L274.0 338.6L273.2 342.7L271.1 341.5L269.2 349.5L267.9 349.5L267.0 351.8L267.5 356.1ZM158.9 358.0L150.5 357.5L148.3 356.3L147.2 353.9L143.2 352.4L141.3 352.8L137.2 351.7L135.6 350.3L136.0 346.1L134.1 343.4L152.8 341.0L154.8 342.7L151.9 350.0L150.7 351.0L152.2 352.4L154.4 351.3L156.7 352.9L155.9 354.2L156.6 355.8L158.4 356.0ZM314.1 341.3L315.7 342.3L315.5 344.8L317.3 344.6L319.4 345.7L319.2 347.5L321.0 350.5L312.7 351.0L302.4 349.5L300.7 346.2L301.3 341.6L303.2 340.9L303.9 339.2L306.8 338.1L308.3 339.8L308.4 342.2L312.5 340.1ZM287.8 353.2L286.7 354.4L286.6 357.7L284.8 359.3L280.6 354.5L277.4 355.1L276.7 350.8L282.7 346.0L283.7 344.0L289.0 345.0L292.9 348.3L291.0 349.5L287.9 349.8ZM263.8 370.9L263.5 368.4L261.7 368.5L261.7 362.9L264.9 361.7L267.5 356.1L267.0 351.8L267.9 349.5L269.2 349.5L269.8 347.8L271.6 348.1L273.1 350.3L276.7 350.8L277.4 355.1L276.2 362.0L277.3 361.8L277.5 364.5L275.6 366.8L272.1 366.8L271.4 371.8L269.2 372.3L269.2 374.0L272.2 375.3L272.9 379.2L270.1 380.9L269.5 382.6L267.8 383.0L267.0 385.3L265.7 384.5L263.4 387.4L262.6 386.4L261.1 386.6L260.3 384.1L265.2 382.1L264.4 380.7L265.0 376.3L262.9 376.5L264.2 374.3ZM314.0 355.5L312.8 354.2L308.9 355.8L309.0 359.3L304.2 359.4L302.5 360.9L292.8 355.6L292.2 353.2L287.8 353.2L287.9 349.8L291.0 349.5L295.0 346.7L298.5 349.5L302.6 350.6L303.7 349.9L312.7 351.0L314.3 350.4L315.1 352.1ZM283.7 344.0L284.8 336.2L283.1 332.9L286.2 331.8L288.5 328.8L288.9 330.6L293.7 335.8L292.1 337.1L289.2 337.4L289.3 338.5L294.0 344.2L295.0 346.7L292.9 348.3L289.0 345.0ZM215.3 334.9L215.4 342.0L220.5 345.5L221.2 346.9L219.5 348.3L219.3 350.5L216.8 353.4L207.1 358.1L205.4 357.5L198.7 358.3L198.0 354.2L200.6 354.8L199.9 350.8L198.5 349.3L194.3 349.6L192.8 351.1L189.7 350.6L191.6 349.3L191.9 347.1L194.3 344.9L199.1 343.4L198.1 342.1L198.4 340.3L196.7 339.0L198.0 339.1L200.0 336.0L202.9 336.7L206.5 335.6L210.3 336.5ZM257.8 418.9L259.3 416.0L257.7 413.8L258.9 409.2L264.1 403.4L264.9 405.4L267.8 405.9L270.0 404.2L272.1 405.4L272.1 407.1L275.8 406.7L276.1 408.4L277.5 408.3L278.1 413.3L276.5 419.6L276.9 423.0L273.8 423.1L273.5 424.7L271.3 426.6L273.9 427.5L274.7 431.0L270.2 432.1L270.0 430.6L267.9 431.5L266.8 425.3L265.4 423.4L261.5 420.9L259.8 421.4ZM321.1 358.3L314.0 355.5L313.8 354.4L315.1 352.1L314.3 350.4L321.0 350.5L319.2 347.5L319.4 345.7L317.3 344.6L315.5 344.8L315.7 342.3L314.1 341.3L316.1 340.2L317.8 334.4L317.8 328.7L318.8 326.4L320.0 326.1L322.9 327.4L324.1 329.4L323.9 331.4L329.2 335.1L331.0 335.4L332.4 337.3L328.7 339.4L326.8 343.0L327.9 348.0L325.2 351.0L325.7 352.2L323.5 352.7L322.6 355.4L323.7 359.0ZM303.9 339.2L303.2 340.9L301.3 341.6L300.7 346.2L302.4 349.5L303.7 349.9L302.6 350.6L298.5 349.5L295.0 346.7L294.0 344.2L289.3 338.5L289.2 337.4L292.1 337.1L293.7 335.8L288.9 330.6L288.5 328.8L289.9 326.6L289.5 324.1L291.3 323.8L293.2 321.6L293.0 319.4L296.9 318.9L298.8 322.5L297.1 324.6L297.9 328.7L299.3 329.2L297.2 333.7L298.6 336.9L301.7 336.6ZM292.1 420.3L292.4 422.3L289.0 427.2L287.0 427.4L284.3 429.9L283.3 435.8L281.5 437.8L281.5 439.5L278.7 439.6L276.6 438.6L273.4 440.0L272.6 431.5L274.7 431.0L273.9 427.5L271.3 426.6L273.5 424.7L273.8 423.1L276.9 423.0L280.9 425.5L283.4 421.1L284.6 421.5L285.1 418.4L286.4 418.6L289.3 421.8ZM214.3 233.8L217.2 232.9L217.9 234.1L221.1 235.3L220.0 237.5L217.3 239.7L214.1 240.7L215.8 243.5L215.0 245.3L212.8 245.0L209.8 246.5L205.7 247.0L202.6 253.5L203.1 255.0L202.0 255.1L202.1 253.6L197.4 250.0L195.0 249.5L193.8 250.8L190.9 247.9L194.1 246.7L197.7 243.6L199.8 243.4L205.8 240.2L207.5 237.7L209.6 236.9L211.5 237.4L211.9 234.6ZM221.1 235.3L226.1 235.2L228.0 236.9L230.9 237.5L230.5 239.6L233.6 242.7L233.7 245.0L230.6 248.3L231.0 250.9L228.3 254.8L226.6 254.0L225.9 252.4L227.4 250.2L226.3 247.7L224.3 247.8L222.9 249.2L218.3 246.0L215.0 245.3L215.8 243.5L214.1 240.7L217.3 239.7ZM203.1 255.0L202.6 253.5L205.7 247.0L215.7 244.8L222.9 249.2L221.9 252.6L214.1 258.8L210.7 259.4ZM185.3 252.0L190.9 247.9L193.8 250.8L195.0 249.5L197.4 250.0L202.1 253.6L202.0 255.1L199.4 256.2L197.4 255.2L197.1 257.1L195.4 257.6L194.3 259.2L194.8 263.4L192.2 263.9L190.2 267.9L187.2 269.1L185.4 268.0L181.4 267.6L178.1 263.6L178.8 256.5L183.9 254.0ZM202.0 255.1L210.7 259.4L215.8 257.7L214.8 259.1L219.1 262.5L217.9 265.5L216.8 265.7L217.8 270.9L216.8 270.5L215.8 272.7L217.5 273.3L216.7 273.7L216.9 275.3L213.2 278.4L213.6 279.2L209.0 278.2L209.9 276.6L206.0 275.6L204.0 278.9L201.0 277.3L200.4 278.5L198.5 278.4L198.0 277.2L196.2 277.3L195.0 273.7L196.7 272.9L197.1 270.5L196.2 266.2L195.0 265.7L193.4 266.9L193.1 268.3L190.2 267.9L192.2 263.9L194.8 263.4L194.3 259.2L195.4 257.6L197.1 257.1L197.4 255.2L199.4 256.2ZM228.2 272.0L229.3 270.9L232.1 270.5L233.5 273.1L237.6 272.2L238.0 277.8L237.2 282.7L235.3 283.3L234.9 284.4L232.7 284.0L230.4 288.3L228.9 287.9L229.7 285.9L228.3 283.8L225.5 284.7L225.9 278.8L223.0 276.6L223.1 274.0L227.4 273.4ZM150.1 274.8L153.5 275.2L149.9 281.3L155.7 282.9L157.7 282.4L158.6 286.5L153.8 291.2L150.4 291.4L146.6 290.0L144.1 293.6L140.9 292.9L143.0 288.7L142.1 287.7L139.7 287.4L141.5 285.0L141.9 281.6L143.7 283.9L145.9 282.6L145.6 281.0L143.1 280.2L142.4 277.9L144.0 277.2L145.2 278.8L146.9 279.1L148.5 278.4L148.5 276.4ZM284.5 271.3L286.3 271.9L287.2 274.2L292.0 275.9L293.0 279.1L298.8 280.3L300.7 284.9L289.8 286.3L279.5 291.1L279.2 289.4L276.4 287.1L276.4 282.3L275.0 280.7L275.6 277.8L278.7 274.4L279.0 271.8L281.4 272.6ZM262.0 276.1L262.5 274.7L264.5 273.5L267.0 274.1L267.2 272.9L269.8 271.3L270.1 277.2L271.9 276.4L275.6 277.8L275.0 280.7L276.4 282.3L276.4 287.1L281.3 292.4L280.4 294.1L279.1 294.1L277.0 295.9L275.9 295.3L272.2 296.9L270.3 296.6L268.1 298.6L267.5 295.4L259.5 299.1L265.3 294.4L261.7 288.4L261.9 284.4L259.9 282.9L264.7 279.5L264.1 277.9L262.4 277.6ZM195.0 273.7L196.2 277.3L200.4 278.5L199.4 279.6L200.2 281.0L198.6 284.0L198.5 288.5L202.2 291.3L200.1 292.8L197.9 293.0L195.2 291.8L195.7 294.3L193.9 296.9L195.0 298.2L195.1 300.9L193.2 300.9L191.9 299.6L192.3 296.6L188.9 296.5L187.8 297.4L185.9 295.4L186.8 293.6L184.8 290.2L184.5 286.2L185.5 285.9L187.9 287.9L189.9 285.7L189.1 282.3L187.3 281.1L185.3 281.4L184.3 280.0L187.3 278.6L185.5 276.5L185.8 274.7L190.2 274.3L192.1 273.1ZM200.4 278.5L201.0 277.3L204.0 278.9L206.0 275.6L209.9 276.6L209.0 278.2L213.6 279.2L215.4 283.3L215.2 288.1L213.0 289.4L211.3 292.5L209.9 293.0L205.0 291.2L202.2 291.3L198.5 288.5L198.6 284.0L200.2 281.0L199.4 279.6ZM162.7 281.0L165.6 286.0L165.1 288.3L162.7 288.1L161.5 289.4L163.3 292.8L162.0 295.1L163.8 297.1L160.9 299.2L157.1 297.8L155.8 300.3L153.2 302.0L147.2 302.4L144.8 303.8L144.0 301.3L145.5 297.5L143.5 294.5L146.6 290.0L150.4 291.4L153.8 291.2L158.6 286.5L157.7 282.4L158.9 282.8ZM300.5 284.6L303.1 284.9L303.1 283.5L308.4 282.5L309.2 283.9L310.6 283.8L311.1 285.3L309.4 289.7L310.6 290.0L310.0 294.6L308.3 295.5L311.5 299.3L304.8 303.3L299.8 302.5L298.7 301.8L298.5 299.6L293.6 298.1L291.7 295.5L293.0 293.0L294.3 293.0L297.8 290.2L300.4 287.0ZM279.5 291.1L289.8 286.3L300.7 284.9L300.4 287.0L297.8 290.2L294.3 293.0L293.0 293.0L291.7 295.5L293.6 298.1L298.5 299.6L298.7 301.8L296.0 302.6L292.6 301.4L290.2 302.8L288.4 302.1L287.2 300.1L285.9 301.0L284.6 300.4L283.3 293.1ZM215.2 288.1L217.4 289.9L216.5 290.9L217.8 292.4L220.2 289.5L226.6 294.3L230.6 288.9L236.1 289.0L235.7 290.8L231.9 294.5L234.7 298.2L232.7 300.3L235.7 302.5L237.1 305.8L236.8 310.3L234.6 311.5L235.6 314.6L234.8 316.1L232.7 316.8L226.7 315.8L224.6 313.2L222.9 313.0L222.9 310.9L225.2 309.1L220.1 307.9L220.3 306.7L218.6 304.8L217.3 306.4L215.8 305.8L214.9 304.0L215.5 297.8L213.5 298.2L212.0 297.4L212.4 295.6L211.3 292.5L213.0 289.4ZM234.8 316.1L235.6 314.6L234.6 311.5L236.8 310.3L237.1 305.8L235.7 302.5L232.7 300.3L234.7 298.2L233.8 296.1L238.8 294.0L239.5 290.4L242.7 288.7L244.2 289.7L247.6 289.9L248.9 291.6L247.0 297.5L248.6 297.9L249.0 302.3L248.3 305.4L250.6 305.7L250.4 307.6L245.0 313.5L238.6 317.4L237.2 315.9ZM189.4 310.6L185.3 318.8L183.6 317.3L178.7 318.3L178.1 317.1L174.9 315.9L169.8 315.7L171.1 315.1L171.7 312.5L171.7 310.1L170.5 308.2L174.4 306.9L176.4 306.8L176.9 308.8L178.4 308.1L181.9 310.1L186.6 309.5ZM195.1 300.9L195.0 298.2L193.9 296.9L195.7 294.3L195.2 291.8L197.9 293.0L200.1 292.8L202.2 291.3L205.0 291.2L209.9 293.0L211.3 292.5L212.4 295.6L212.0 297.4L213.5 298.2L215.5 297.8L214.9 304.0L215.8 305.8L214.3 309.2L212.4 309.4L211.4 311.0L205.6 310.8L203.4 308.4L201.6 310.6L199.4 309.7L199.1 308.3L202.1 305.0L200.5 303.7L198.4 304.0L196.4 301.1ZM185.9 295.4L187.8 297.4L188.9 296.5L192.3 296.6L191.9 299.6L193.2 300.9L193.1 302.7L190.4 310.0L181.9 310.1L178.4 308.1L176.9 308.8L175.3 302.7L171.3 296.4L171.7 293.6L174.0 292.8L179.3 293.6L180.9 292.3L181.5 294.0L184.2 295.7ZM273.1 296.6L280.4 294.1L281.3 292.4L283.3 293.1L284.6 300.4L285.9 301.0L287.8 304.8L286.1 306.1L286.0 309.5L289.1 308.7L290.4 309.5L292.5 309.0L293.9 311.8L292.4 314.4L290.9 314.5L289.1 316.4L286.0 316.7L284.0 319.4L282.4 319.9L279.8 318.9L280.0 317.9L278.4 316.4L280.5 313.2L279.4 308.5L278.3 308.0L278.3 304.7L276.9 303.0L277.8 301.7L275.8 299.1L274.6 299.2ZM259.5 299.1L267.5 295.4L268.1 298.6L270.3 296.6L272.2 296.9L269.8 299.2L271.3 303.2L271.1 305.8L268.2 305.9L266.5 304.5L265.1 307.7L263.8 308.3L263.4 312.2L258.9 314.1L256.9 308.1L251.3 308.1L250.4 307.6L250.6 305.7L248.3 305.4L249.0 302.3L251.6 300.8L253.0 298.6L258.4 300.5ZM272.2 296.9L273.1 296.6L274.6 299.2L275.8 299.1L277.8 301.7L276.9 303.0L278.3 304.7L278.3 308.0L279.4 308.5L280.5 313.2L279.1 316.2L278.4 316.4L276.8 314.0L275.7 316.4L274.5 316.8L271.4 314.8L270.2 315.9L267.3 315.3L267.8 313.7L266.8 312.7L264.1 313.2L263.4 312.2L263.8 308.3L265.1 307.7L266.5 304.5L268.2 305.9L271.1 305.8L271.3 303.2L269.8 299.2ZM155.8 300.3L160.7 304.2L159.2 306.2L157.1 306.0L156.0 309.1L154.4 307.5L150.3 309.0L149.5 312.1L151.6 315.1L151.1 319.3L141.4 317.8L140.4 316.2L136.1 313.9L137.0 311.6L142.8 308.6L144.8 303.8L147.2 302.4L153.2 302.0ZM193.2 300.9L196.4 301.1L198.4 304.0L200.5 303.7L202.1 305.0L199.1 308.3L198.8 311.7L201.6 314.0L200.3 318.1L196.4 319.6L194.1 319.1L192.9 314.7L192.0 314.3L194.5 311.5L194.3 307.5L191.9 305.9ZM298.8 322.5L296.9 318.9L298.5 314.1L296.4 312.1L293.9 311.8L292.5 309.0L290.4 309.5L289.1 308.7L286.0 309.5L286.1 306.1L287.8 304.8L285.9 301.0L287.2 300.1L288.4 302.1L290.2 302.8L292.6 301.4L296.0 302.6L298.7 301.8L304.8 303.3L307.0 305.8L306.6 308.8L307.8 309.8L308.0 312.5L310.1 313.2L310.0 315.1L306.5 318.3L307.1 320.7L306.4 322.1L302.4 321.2ZM160.7 304.2L161.5 304.0L164.7 308.5L169.7 308.9L170.5 308.2L171.7 310.1L171.1 315.1L169.2 316.6L169.0 320.1L167.1 315.2L166.1 317.7L166.8 320.3L165.1 321.2L164.6 320.3L160.3 320.9L159.3 319.4L158.0 319.3L156.1 321.6L155.4 324.3L153.9 324.2L152.5 322.8L152.8 318.8L151.2 318.1L151.6 315.1L149.5 312.1L150.3 309.0L154.4 307.5L156.0 309.1L157.1 306.0L159.2 306.2ZM215.8 305.8L217.3 306.4L218.6 304.8L220.3 306.7L220.1 307.9L225.2 309.1L222.9 310.9L222.9 313.0L224.6 313.2L226.7 315.8L225.1 320.2L219.5 321.1L210.2 325.7L208.7 321.9L201.9 324.9L197.6 322.8L195.6 320.3L200.3 318.1L201.6 314.0L198.8 311.7L199.4 309.7L201.6 310.6L203.4 308.4L205.6 310.8L211.4 311.0L212.4 309.4L214.3 309.2ZM193.1 331.1L191.2 331.4L188.5 324.0L184.5 324.4L179.4 326.5L176.9 323.0L177.1 319.7L178.7 318.3L183.6 317.3L185.3 318.8L191.9 305.9L194.3 307.5L194.5 311.5L192.0 314.3L192.9 314.7L194.1 319.1L196.4 319.6L195.6 320.3L197.6 322.8L201.9 324.9L208.7 321.9L210.2 325.7L206.3 324.1L204.1 326.7L201.6 327.1L200.6 328.6L195.4 331.2ZM251.1 324.6L247.4 324.7L246.6 323.0L244.1 322.4L243.6 319.0L241.6 318.0L240.4 319.0L238.1 318.8L241.7 315.1L245.0 313.5L250.4 307.6L256.9 308.1L258.9 314.1L263.4 312.2L264.1 313.2L264.2 314.8L262.6 316.6L263.1 319.3L261.6 318.0L258.3 319.0L257.2 317.8L251.8 320.6ZM191.2 331.4L184.0 334.3L183.8 336.2L181.7 337.3L180.0 334.8L177.8 334.7L174.1 337.2L171.7 336.6L171.7 334.5L169.4 333.3L169.0 331.8L169.8 330.2L172.4 329.4L169.4 328.1L165.1 321.2L166.8 320.3L166.1 317.7L167.1 315.2L169.0 320.1L169.8 315.7L174.9 315.9L178.1 317.1L178.7 318.3L177.1 319.7L176.9 323.0L179.4 326.5L184.5 324.4L188.5 324.0ZM296.9 318.9L293.0 319.4L293.2 321.6L291.3 323.8L289.5 324.1L289.9 326.6L288.6 328.6L287.7 326.2L282.7 326.5L282.2 325.0L277.3 323.2L277.4 321.2L278.7 319.6L275.7 316.4L276.8 314.0L280.0 317.9L279.8 318.9L282.4 319.9L284.0 319.4L286.0 316.7L289.1 316.4L290.9 314.5L292.4 314.4L293.9 311.8L296.4 312.1L298.5 314.1ZM140.6 331.3L134.2 330.8L132.9 329.6L134.7 325.4L134.6 323.3L133.2 323.6L131.7 319.9L133.1 318.7L135.6 319.3L140.4 316.2L141.4 317.8L146.2 318.7L146.1 323.2L143.8 327.0L141.0 326.9L141.4 330.8ZM283.1 332.9L280.0 332.1L278.3 333.6L275.6 330.5L273.4 330.6L270.3 326.8L270.3 324.7L267.4 325.5L263.0 324.2L263.9 321.4L262.6 316.6L264.2 314.8L264.1 313.2L266.8 312.7L267.8 313.7L267.3 315.3L270.2 315.9L271.4 314.8L274.5 316.8L275.7 316.4L278.7 319.6L277.4 321.2L277.3 323.2L282.2 325.0L282.7 326.5L287.7 326.2L288.5 328.8L286.2 331.8ZM244.1 322.4L241.1 327.4L239.5 327.6L237.7 330.0L235.9 327.6L235.8 325.7L231.6 326.7L231.0 329.4L229.0 328.7L224.9 329.3L224.4 326.7L221.8 324.4L221.7 322.2L219.0 321.6L225.1 320.2L226.7 315.8L232.7 316.8L237.2 315.9L238.6 317.4L238.1 318.8L240.4 319.0L241.6 318.0L243.6 319.0ZM135.5 343.0L134.9 341.5L136.0 338.7L138.6 337.8L141.4 330.8L141.0 326.9L143.8 327.0L146.1 323.2L146.2 318.7L151.1 319.3L151.2 318.1L152.8 318.8L152.5 322.8L153.9 324.2L155.4 324.3L155.4 326.4L156.5 326.6L157.3 329.7L159.9 329.5L160.8 330.5L157.8 332.3L157.4 338.7L156.5 340.5L141.4 342.3L139.3 343.2ZM336.8 332.0L337.6 333.4L336.8 335.6L332.4 337.3L331.0 335.4L329.2 335.1L323.9 331.4L324.1 329.4L328.7 327.8L329.1 325.2L326.1 322.3L327.0 320.5L326.0 319.7L328.6 316.2L332.2 315.5L335.0 317.4L337.4 317.3L340.4 321.7L344.5 321.7L345.0 322.9L341.6 328.3L338.4 329.5ZM169.0 331.8L167.1 334.6L159.1 336.3L157.9 335.5L157.8 332.3L160.8 330.5L159.9 329.5L157.3 329.7L156.5 326.6L155.4 326.4L156.1 321.6L158.0 319.3L159.3 319.4L160.3 320.9L164.6 320.3L169.4 328.1L172.4 329.4L169.8 330.2ZM254.0 336.4L252.7 330.2L255.2 330.4L258.0 327.5L257.8 326.2L253.3 326.7L251.1 324.6L251.8 320.6L257.2 317.8L258.3 319.0L261.6 318.0L263.9 321.4L263.0 324.2L267.4 325.5L270.3 324.7L270.3 326.8L273.4 330.6L275.6 330.5L277.2 332.9L275.4 334.1L275.6 337.4L274.0 338.6L273.2 337.6L271.3 337.7L270.2 335.4L266.8 337.9L266.8 339.4L260.4 337.1ZM224.2 328.0L222.5 329.2L220.6 328.9L220.1 330.5L221.2 331.8L221.4 334.5L219.5 335.5L217.6 334.5L213.5 335.0L210.3 336.5L206.5 335.6L202.9 336.7L200.0 336.0L198.0 339.1L196.7 339.0L193.1 331.1L195.4 331.2L200.6 328.6L201.6 327.1L204.1 326.7L206.3 324.1L210.2 325.7L219.0 321.6L221.7 322.2L221.8 324.4L224.4 326.7ZM314.1 341.3L312.5 340.1L308.4 342.2L308.3 339.8L306.8 338.1L303.9 339.2L301.7 336.6L298.6 336.9L297.2 333.7L299.3 329.2L297.9 328.7L297.1 324.6L298.8 322.5L302.4 321.2L306.4 322.1L306.1 324.5L312.8 327.0L315.7 329.5L317.8 328.7L317.1 337.0L316.1 340.2ZM304.8 303.3L307.7 302.1L308.9 300.3L310.6 301.2L310.8 302.6L316.8 303.7L319.6 305.6L319.6 307.7L321.3 308.1L322.8 310.5L319.8 312.6L319.8 315.6L320.6 315.9L311.8 320.9L312.3 321.5L310.3 325.9L306.1 324.5L307.1 320.7L306.5 318.3L310.0 315.1L310.1 313.2L308.0 312.5L307.8 309.8L306.6 308.8L307.0 305.8ZM324.1 329.4L322.9 327.4L320.0 326.1L318.8 326.4L317.8 328.7L315.7 329.5L310.3 325.9L312.3 321.5L311.8 320.9L320.6 315.9L322.9 317.3L323.7 315.1L325.5 315.3L327.2 317.4L326.0 319.7L327.0 320.5L326.1 322.3L329.1 325.2L328.7 327.8ZM160.7 304.2L162.9 301.4L165.9 300.8L168.1 296.9L169.9 296.4L169.7 294.0L171.7 293.6L171.3 296.4L173.6 299.3L176.4 306.8L169.7 308.9L164.7 308.5L161.5 304.0ZM217.8 270.9L219.0 272.8L217.5 273.3L215.8 272.7L216.8 270.5ZM222.9 249.2L224.3 247.8L226.3 247.7L227.4 250.2L225.9 252.4L226.6 254.0L228.3 254.8L225.7 259.2L226.7 261.5L223.5 262.8L219.1 262.5L214.8 259.1L216.6 256.6L219.6 254.8L219.8 253.6L221.9 252.6ZM236.3 269.4L236.7 268.2L238.2 268.1L240.0 269.8L239.3 273.2L241.1 273.8L242.8 271.7L248.5 273.3L248.8 271.2L252.6 267.8L256.6 266.5L258.6 266.7L258.5 269.1L260.5 270.0L261.0 271.6L255.1 277.1L252.6 277.4L253.6 280.2L249.7 282.5L249.5 289.9L244.2 289.7L242.7 288.7L239.5 290.4L238.8 294.0L233.8 296.1L231.9 294.5L235.7 290.8L236.1 289.0L230.4 288.3L232.7 284.0L234.9 284.4L235.3 283.3L237.2 282.7L238.0 277.8L237.6 272.2ZM257.5 275.0L259.1 275.3L260.4 274.0L262.4 277.6L264.1 277.9L264.7 279.5L259.9 282.9L261.9 284.4L261.7 288.4L265.3 294.4L258.4 300.5L253.0 298.6L251.6 300.8L249.0 302.3L248.6 297.9L247.0 297.5L248.9 291.6L247.6 289.9L249.5 289.9L249.7 282.5L253.6 280.2L252.6 277.4L255.1 277.1ZM311.5 299.3L314.3 301.4L317.3 301.4L319.8 300.2L322.4 296.1L324.6 296.5L327.5 301.5L330.4 302.5L331.8 307.3L335.3 308.3L336.9 306.8L335.8 312.4L338.2 313.6L337.4 317.3L335.0 317.4L332.2 315.5L327.2 317.4L325.5 315.3L323.7 315.1L322.9 317.3L319.8 315.6L319.8 312.6L322.8 310.5L321.3 308.1L319.6 307.7L319.6 305.6L316.8 303.7L310.8 302.6L310.6 301.2L308.9 300.3ZM221.4 270.7L220.4 269.3L223.5 266.8L226.0 267.5L225.1 265.2L228.3 263.8L229.5 264.4L228.2 266.8L229.9 267.7L227.4 273.4L223.1 274.0ZM158.9 358.0L150.5 357.5L148.3 356.3L147.2 353.9L143.2 352.4L141.3 352.8L137.2 351.7L135.6 350.3L136.0 346.1L134.1 343.4L130.9 345.6L128.0 345.2L126.9 343.7L127.5 339.2L126.0 337.2L126.4 335.9L128.9 336.7L129.5 334.0L127.4 334.5L126.1 331.5L128.1 332.2L130.5 329.4L132.9 329.6L134.7 325.4L134.6 323.3L133.2 323.6L131.7 319.9L133.1 318.7L135.6 319.3L140.4 316.2L136.1 313.9L137.0 311.6L142.8 308.6L144.4 306.0L144.0 301.3L145.5 297.5L143.5 294.5L144.1 293.6L140.9 292.9L143.0 288.7L142.1 287.7L139.7 287.4L141.5 285.0L141.9 281.6L143.7 283.9L145.9 282.6L145.6 281.0L143.1 280.2L142.4 277.9L144.0 277.2L145.2 278.8L146.9 279.1L148.5 278.4L148.5 276.4L150.1 274.8L153.5 275.2L149.9 281.3L158.9 282.8L162.7 281.0L165.6 286.0L165.1 288.3L162.7 288.1L161.5 289.4L163.3 292.8L162.0 295.1L163.8 297.1L160.9 299.2L157.1 297.8L155.8 300.3L160.7 304.2L162.9 301.4L165.9 300.8L168.1 296.9L169.9 296.4L169.7 294.0L174.0 292.8L179.3 293.6L180.9 292.3L181.5 294.0L184.2 295.7L185.9 295.4L186.8 293.6L184.8 290.2L184.5 286.2L185.5 285.9L187.9 287.9L189.9 285.7L189.1 282.3L187.3 281.1L185.3 281.4L184.3 280.0L187.3 278.6L185.5 276.5L185.8 274.7L190.2 274.3L192.1 273.1L195.0 273.7L196.7 272.9L197.1 270.5L196.2 266.2L195.0 265.7L193.4 266.9L193.1 268.3L190.2 267.9L187.2 269.1L185.4 268.0L181.4 267.6L178.1 263.6L178.8 256.5L183.9 254.0L186.3 250.7L194.1 246.7L197.7 243.6L199.8 243.4L205.8 240.2L207.5 237.7L209.6 236.9L211.5 237.4L211.9 234.6L217.2 232.9L217.9 234.1L221.1 235.3L226.1 235.2L228.0 236.9L230.9 237.5L230.5 239.6L233.6 242.7L233.7 245.0L230.6 248.3L231.0 250.9L225.7 259.2L226.7 261.5L223.5 262.8L219.1 262.5L217.9 265.5L216.8 265.7L216.9 268.1L219.0 272.8L216.7 273.7L216.9 275.3L213.2 278.4L215.4 283.3L215.2 288.1L217.4 289.9L216.5 290.9L217.8 292.4L220.2 289.5L226.6 294.3L230.1 290.5L230.4 288.3L228.9 287.9L229.7 285.9L228.3 283.8L225.5 284.7L225.9 278.8L223.0 276.6L223.1 274.0L220.4 269.3L223.5 266.8L226.0 267.5L225.1 265.2L228.3 263.8L229.5 264.4L228.2 266.8L229.9 267.7L228.2 272.0L229.3 270.9L232.1 270.5L233.5 273.1L235.9 273.0L237.6 272.2L236.3 269.4L236.7 268.2L238.2 268.1L240.0 269.8L239.3 273.2L241.1 273.8L242.8 271.7L248.5 273.3L248.8 271.2L252.6 267.8L258.6 266.7L258.5 269.1L260.5 270.0L261.0 271.6L257.5 275.0L259.1 275.3L260.4 274.0L262.0 276.1L262.5 274.7L264.5 273.5L267.0 274.1L267.2 272.9L269.8 271.3L270.1 277.2L271.9 276.4L275.6 277.8L278.7 274.4L279.0 271.8L281.4 272.6L284.5 271.3L286.3 271.9L287.2 274.2L292.0 275.9L293.0 279.1L298.8 280.3L300.5 284.6L303.1 284.9L303.1 283.5L308.4 282.5L309.2 283.9L310.6 283.8L311.1 285.3L309.4 289.7L310.6 290.0L310.0 294.6L308.3 295.5L311.5 299.3L308.9 300.3L307.7 302.1L304.8 303.3L298.7 301.8L296.0 302.6L292.6 301.4L290.2 302.8L288.4 302.1L287.2 300.1L285.9 301.0L287.8 304.8L286.1 306.1L286.0 309.5L289.1 308.7L290.4 309.5L292.5 309.0L293.9 311.8L296.4 312.1L298.5 314.1L296.9 318.9L293.0 319.4L293.2 321.6L291.3 323.8L289.5 324.1L289.9 326.6L286.2 331.8L283.1 332.9L280.0 332.1L278.3 333.6L277.2 332.9L275.4 334.1L275.6 337.4L274.0 338.6L273.2 342.7L271.1 341.5L269.2 349.5L267.9 349.5L267.0 351.8L267.5 356.1L265.9 359.3L265.1 359.5L260.9 356.8L257.6 352.8L254.5 352.9L253.3 354.2L250.2 354.9L246.4 353.1L242.1 354.5L241.2 351.5L235.7 350.5L235.5 352.1L229.0 353.4L228.6 355.4L221.9 355.8L220.4 354.8L218.7 355.5L218.5 353.8L216.8 353.4L207.1 358.1L205.4 357.5L198.7 358.3L198.0 354.2L200.6 354.8L199.9 350.8L198.5 349.3L194.3 349.6L192.8 351.1L189.7 350.6L186.3 353.2L185.3 352.9L183.8 354.1L183.8 356.3L180.4 359.3L181.3 360.5L180.4 362.7L177.6 362.5L175.3 365.1L170.8 365.1L170.6 361.1L169.2 358.4ZM294.6 388.6L293.1 387.3L291.6 388.5L291.0 390.8L293.9 394.2L295.6 394.5L294.8 400.9L296.4 400.7L296.9 402.7L298.2 402.7L297.7 407.9L299.4 414.4L298.0 414.7L297.6 417.3L296.2 418.5L292.1 420.3L292.4 422.3L289.0 427.2L287.0 427.4L284.3 429.9L283.3 435.8L281.5 437.8L281.5 439.5L278.7 439.6L276.6 438.6L273.4 440.0L272.6 431.5L270.2 432.1L270.0 430.6L267.9 431.5L266.8 425.3L265.4 423.4L261.5 420.9L259.8 421.4L257.8 418.9L259.3 416.0L257.7 413.8L258.9 409.2L264.1 403.4L264.9 405.4L267.8 405.9L270.0 404.2L270.9 400.6L267.0 398.5L263.5 393.8L261.8 393.7L263.1 391.1L263.4 387.4L262.6 386.4L261.1 386.6L260.3 384.1L265.2 382.1L264.4 380.7L265.0 376.3L262.9 376.5L264.2 374.3L263.5 368.4L261.7 368.5L261.7 362.9L264.9 361.7L267.5 356.1L267.0 351.8L267.9 349.5L269.2 349.5L271.1 341.5L273.2 342.7L274.0 338.6L275.6 337.4L275.4 334.1L277.2 332.9L278.3 333.6L280.0 332.1L283.1 332.9L286.2 331.8L289.9 326.6L289.5 324.1L291.3 323.8L293.2 321.6L293.0 319.4L296.9 318.9L298.5 314.1L296.4 312.1L293.9 311.8L292.5 309.0L290.4 309.5L289.1 308.7L286.0 309.5L286.1 306.1L287.8 304.8L285.9 301.0L287.2 300.1L288.4 302.1L290.2 302.8L292.6 301.4L296.0 302.6L298.7 301.8L304.8 303.3L311.5 299.3L314.3 301.4L317.3 301.4L319.8 300.2L322.4 296.1L324.6 296.5L327.5 301.5L330.4 302.5L331.8 307.3L335.3 308.3L336.9 306.8L335.8 312.4L338.2 313.6L337.4 317.3L337.9 318.6L340.4 321.7L344.5 321.7L345.0 322.9L341.6 328.3L336.8 331.0L337.6 333.4L336.8 335.6L333.7 336.2L328.7 339.4L326.8 343.0L327.9 348.0L325.2 351.0L325.7 352.2L323.5 352.7L322.6 355.4L323.7 359.0L321.1 358.3L319.5 363.7L317.1 364.3L314.7 362.9L307.8 363.3L307.8 364.9L304.0 370.8L301.4 369.7L302.2 375.3L301.3 376.9L303.5 379.8L302.7 388.4L309.4 389.9L309.0 393.6L306.8 394.8L305.5 392.5L303.7 391.7L301.6 393.4L299.9 390.3Z","East":"M528.1 311.9L526.0 309.5L524.6 309.4L522.5 314.4L520.6 313.7L519.2 303.0L519.9 302.2L519.2 299.5L523.9 297.8L525.1 293.4L524.2 291.3L524.9 289.5L527.3 289.0L527.6 293.1L526.6 295.8L531.5 297.3L530.3 302.7L528.1 303.8ZM528.1 306.0L528.1 303.8L530.3 302.7L531.5 297.3L535.5 300.1L536.2 306.4L535.4 309.9L532.9 310.2L532.4 304.4L530.6 306.2ZM531.3 320.0L530.6 317.1L528.2 316.4L528.0 314.2L532.8 312.4L534.7 313.7L535.4 316.2L533.5 321.1ZM524.9 289.5L524.2 291.3L525.1 293.4L523.9 297.8L519.2 299.5L519.3 297.7L517.3 295.9L517.6 292.8L518.9 292.6L519.2 290.8L520.6 290.3L522.1 286.7L522.9 288.8ZM525.7 331.4L525.7 335.0L523.4 336.5L524.0 342.6L520.9 340.6L520.5 343.1L518.6 343.8L516.7 327.1L518.7 327.6L521.9 331.7L523.7 329.8ZM520.2 312.5L520.6 315.7L523.2 316.1L524.2 317.2L524.8 322.4L524.0 324.0L525.4 326.2L525.7 331.4L523.7 329.8L521.9 331.7L518.7 327.6L516.7 327.1L513.8 322.7L514.2 319.9L513.3 315.8L515.6 314.7L516.1 312.3ZM519.2 299.5L520.2 312.5L516.1 312.3L515.6 314.7L514.5 315.2L511.7 306.6L511.4 302.7L512.8 300.4L512.9 294.2L512.2 292.8L514.9 292.7L515.2 294.9L517.2 294.4L517.3 295.9L519.3 297.7ZM524.0 342.6L523.4 336.5L525.7 335.0L526.2 329.7L525.4 326.2L527.3 326.9L529.2 326.1L531.3 337.7L530.5 339.9L528.2 339.5L527.1 342.8L525.8 343.7ZM520.6 313.7L522.5 314.4L524.6 309.4L526.0 309.5L528.1 311.9L528.2 316.4L530.6 317.1L531.3 320.0L529.7 321.1L526.3 318.6L525.5 317.0L520.6 315.7ZM528.1 311.9L528.1 306.0L530.6 306.2L532.4 304.4L532.9 310.2L535.4 309.9L535.6 312.4L534.7 313.7L532.8 312.4L528.0 314.2ZM525.4 326.2L524.0 324.0L524.8 322.4L524.2 317.2L525.5 317.0L526.3 318.6L529.7 321.1L530.2 322.9L529.2 326.1L527.3 326.9ZM392.1 371.5L391.6 370.2L386.6 370.3L384.9 369.0L386.7 361.5L389.8 363.8L392.7 364.4L393.2 363.0L395.9 361.6L397.1 362.0L397.0 364.7L399.6 370.1L399.4 371.6L396.0 372.9ZM372.8 363.7L373.6 367.0L377.2 368.5L376.8 370.4L379.6 371.2L376.6 374.7L374.4 374.4L371.6 377.1L370.1 378.0L368.4 377.5L367.4 378.6L363.1 377.9L360.4 375.6L360.2 374.1L361.8 373.7L363.2 371.5L362.6 365.2L363.5 362.9L366.1 362.7L368.3 364.3L369.3 363.1L372.7 363.0ZM385.0 366.5L384.9 369.0L386.6 370.3L391.6 370.2L392.1 371.5L386.9 376.1L384.8 376.4L382.0 375.0L379.8 376.2L377.1 373.7L379.6 371.2L376.8 370.4L377.2 368.5L373.6 367.0L372.8 363.7L375.2 364.8L379.5 363.7L380.0 366.4L381.8 365.1ZM342.5 368.6L342.2 369.6L340.7 369.4L337.5 371.0L333.1 370.8L331.6 371.9L329.9 376.9L328.2 377.6L328.0 375.9L329.5 373.6L327.6 369.8L328.9 367.3L326.2 367.3L325.1 365.5L328.2 363.9L332.0 363.4L333.2 364.7L334.3 362.9L336.3 363.9L337.0 365.7L339.6 366.1ZM301.4 369.7L304.0 370.8L307.8 364.9L308.9 366.4L308.7 369.5L310.0 375.3L308.7 379.2L312.5 384.5L309.4 389.9L302.7 388.4L303.5 379.8L301.3 376.9L302.2 375.3ZM325.1 365.5L326.2 367.3L328.9 367.3L327.6 369.8L329.5 373.6L328.0 375.9L327.8 378.6L329.5 380.1L329.2 380.8L328.5 381.7L326.0 380.2L324.3 380.5L318.5 385.8L317.3 385.1L311.6 385.9L312.5 384.5L308.7 379.2L310.0 375.3L309.1 373.1L313.1 369.0L315.0 370.4L318.2 370.1L320.0 368.5L321.5 370.2L324.0 369.4L323.6 367.1ZM328.2 377.6L329.9 376.9L331.6 371.9L335.3 370.5L337.5 371.0L340.7 369.4L344.0 370.7L344.2 372.2L350.1 375.5L352.8 375.9L351.0 379.8L347.4 380.7L344.8 377.6L344.4 374.8L342.4 374.4L341.4 376.4L333.9 376.4L332.3 378.1L332.7 379.8L329.5 380.1L327.8 378.6ZM360.4 375.6L363.1 377.9L367.4 378.6L368.4 377.5L370.1 378.0L374.4 374.4L376.6 374.7L377.1 373.7L379.8 376.2L382.0 375.0L384.8 376.4L384.2 381.4L379.1 381.1L379.8 383.1L382.1 384.8L382.8 388.5L381.6 388.9L374.4 379.8L371.0 383.0L364.7 383.3L365.5 381.3L361.1 381.0L354.5 377.7L355.6 376.4L359.3 376.4ZM329.2 380.8L329.5 380.1L332.7 379.8L332.3 378.1L333.9 376.4L341.4 376.4L342.4 374.4L344.4 374.8L344.8 377.6L347.4 380.7L349.0 385.7L346.9 386.5L346.2 385.9L344.3 389.1L344.8 390.7L346.4 390.8L346.4 393.0L340.5 392.9L340.0 395.7L341.7 397.3L341.5 399.4L339.3 398.1L335.3 398.0L333.3 396.3L331.7 396.2L329.1 397.5L328.5 393.9L326.6 393.0L327.3 391.3L325.9 390.6L327.9 387.4L330.5 387.4L332.0 384.8L330.7 384.3L330.4 381.3ZM347.4 380.7L351.0 379.8L352.8 375.9L361.1 381.0L365.5 381.3L364.6 384.5L362.5 384.6L362.2 386.7L365.4 385.4L365.8 388.1L367.4 390.4L366.6 391.4L356.5 392.0L353.5 387.6L354.2 384.7L351.5 383.3L349.0 385.7ZM364.7 383.3L371.0 383.0L374.4 379.8L380.6 387.3L376.8 387.2L376.5 385.5L372.6 386.4L372.1 388.2L369.4 390.5L369.7 393.1L367.0 392.4L363.1 395.0L361.6 397.2L356.5 392.0L366.6 391.4L367.4 390.4L365.8 388.1L365.4 385.4L362.2 386.7L362.5 384.6L364.6 384.5ZM309.4 389.9L311.6 385.9L317.3 385.1L318.5 385.8L324.3 380.5L326.0 380.2L328.5 381.7L329.2 380.8L330.4 381.3L330.7 384.3L332.0 384.8L330.5 387.4L327.9 387.4L325.9 390.6L324.0 391.5L325.0 395.7L323.2 398.4L319.4 402.1L316.9 399.7L312.7 408.1L311.7 405.2L310.3 405.0L305.8 398.5L306.8 394.8L309.0 393.6ZM382.8 388.5L382.1 384.8L379.8 383.1L379.1 381.1L387.3 381.4L394.2 382.9L395.4 382.0L390.2 385.6L386.8 390.1ZM380.6 387.3L381.6 388.9L384.7 388.8L386.8 390.1L383.7 392.0L372.1 395.5L361.4 401.8L361.6 397.2L363.1 395.0L367.0 392.4L369.7 393.1L369.4 390.5L372.1 388.2L372.6 386.4L376.5 385.5L376.8 387.2ZM306.8 394.8L306.2 400.1L310.3 405.0L311.7 405.2L312.4 407.8L311.3 409.1L305.2 407.9L298.6 409.7L297.7 407.9L298.2 402.7L296.9 402.7L296.4 400.7L294.8 400.9L295.6 394.5L293.9 394.2L293.6 392.9L291.5 391.9L291.6 388.5L293.1 387.3L299.9 390.3L301.6 393.4L303.7 391.7L305.5 392.5ZM325.9 390.6L327.3 391.3L326.6 393.0L328.5 393.9L329.1 397.5L331.7 396.2L333.3 396.3L335.3 398.0L336.2 400.4L335.5 403.5L336.9 403.9L336.2 407.6L337.3 409.5L335.5 412.6L331.2 412.7L331.4 414.5L329.5 410.2L325.6 411.7L325.2 413.5L322.8 412.1L321.3 409.7L320.4 411.9L318.3 409.4L315.9 411.1L315.5 409.7L312.7 408.1L316.9 399.7L319.4 402.1L323.2 398.4L325.0 395.7L324.0 391.5ZM312.4 407.8L315.5 409.7L315.9 411.1L318.3 409.4L320.4 411.9L321.3 409.7L322.8 412.1L322.1 413.1L324.1 415.4L319.7 418.2L318.3 417.4L315.4 421.4L316.3 422.9L315.8 424.9L317.1 426.1L314.8 427.1L311.3 425.0L310.6 427.5L306.7 429.0L307.1 427.0L304.5 423.3L303.3 424.2L301.9 421.1L300.9 422.2L297.9 420.8L296.2 418.5L297.6 417.3L298.0 414.7L299.4 414.4L298.6 409.7L305.2 407.9L311.3 409.1ZM296.2 418.5L297.9 420.8L300.9 422.2L301.9 421.1L303.0 422.5L303.3 424.2L301.9 425.7L302.4 426.9L300.9 431.1L301.7 431.4L301.0 434.3L299.5 435.5L297.8 434.1L294.9 433.7L286.1 439.4L281.5 439.5L281.5 437.8L283.3 435.8L284.3 429.9L287.0 427.4L289.0 427.2L292.4 422.3L292.1 420.3ZM423.9 351.5L421.9 351.2L423.5 347.7L424.8 348.1ZM442.8 349.1L444.4 352.0L443.2 352.9ZM423.0 338.1L423.2 335.8L426.3 331.6L428.5 332.5L430.0 331.3L432.3 331.2L433.1 329.8L436.2 334.1L438.7 335.6L438.3 336.6L443.6 341.1L443.9 345.2L442.6 346.3L441.0 345.4L439.3 345.8L437.4 345.1L437.4 343.7L432.5 347.4L430.7 345.9L429.6 346.4L428.9 348.5L426.9 350.3L427.7 351.6L426.8 354.2L424.3 345.1L425.9 341.7L425.1 339.8ZM437.7 347.1L438.4 351.6L437.1 351.8L435.9 349.7L436.4 346.2ZM441.2 346.0L441.3 349.3L438.6 345.7ZM366.1 362.7L363.5 362.9L362.6 365.2L363.2 371.5L361.8 373.7L360.2 374.1L360.4 375.6L359.3 376.4L355.6 376.4L354.5 377.7L352.8 375.9L350.1 375.5L344.2 372.2L344.0 370.7L342.2 369.6L346.8 365.0L349.0 365.5L352.3 363.0L356.5 362.4L355.6 360.9L358.0 360.1L357.6 357.5L361.0 356.3L362.1 355.0L360.8 353.3L361.8 351.8L363.2 351.5L364.8 355.9L364.2 357.4L365.0 358.5L363.8 360.1ZM392.1 371.5L396.0 372.9L397.1 372.0L400.6 372.1L401.3 373.4L395.6 377.5L394.5 379.2L395.4 382.0L394.2 382.9L384.2 381.4L384.8 376.4L386.9 376.1ZM341.5 399.4L341.7 397.3L340.0 395.7L340.5 392.9L346.4 393.0L346.4 390.8L344.8 390.7L344.3 389.1L346.2 385.9L346.9 386.5L349.0 385.7L351.5 383.3L354.2 384.7L353.5 387.6L357.8 394.2L361.6 397.2L361.4 401.8L352.8 410.7L351.4 409.8L349.1 412.0L347.5 411.4L345.5 413.7L344.1 412.3L344.2 408.2L345.6 406.6L344.6 403.7L345.1 399.4ZM335.3 398.0L345.1 399.4L344.6 403.7L345.6 406.6L344.2 408.2L344.1 412.3L345.8 414.2L343.3 417.3L338.5 418.0L337.4 416.7L334.3 416.7L332.2 412.1L335.5 412.6L337.3 409.5L336.2 407.6L336.9 403.9L335.5 403.5L336.2 400.4ZM411.2 322.9L410.9 323.6L413.4 325.5L415.8 326.3L419.0 332.2L413.8 334.1L413.8 336.5L415.4 336.8L415.0 341.2L411.5 342.2L409.8 343.9L409.4 349.4L406.1 348.5L405.2 344.9L403.0 341.2L404.4 337.1L403.5 334.4L401.8 333.3L401.5 324.6L407.7 322.3ZM409.4 349.4L409.8 343.9L411.5 342.2L415.0 341.2L415.4 336.8L413.8 336.5L413.8 334.1L419.0 332.2L420.3 337.7L421.6 339.4L424.1 339.8L425.2 341.7L416.8 350.9L410.4 352.8ZM394.3 339.1L396.7 341.8L400.7 343.2L400.1 346.0L402.2 347.2L403.5 349.3L401.1 350.6L400.2 353.2L392.6 353.6L390.2 357.1L386.6 360.0L384.6 358.2L383.4 360.5L381.4 360.6L381.2 358.9L382.4 357.2L379.1 354.7L380.6 352.5L374.9 348.4L374.6 345.2L372.0 343.7L372.2 342.7L373.6 342.4L374.4 344.4L376.8 344.4L379.1 341.4L379.6 339.7L378.4 338.4L379.5 336.4L379.3 334.3L378.1 333.0L380.3 331.3L381.4 332.9L384.9 333.8L386.5 336.1L388.0 336.8L389.4 336.1ZM361.8 351.8L360.8 353.3L358.8 353.4L358.3 351.5L353.3 351.6L350.6 349.5L349.0 349.6L347.7 348.3L346.7 339.7L342.2 341.7L342.3 343.7L338.1 344.3L336.1 346.2L334.4 345.3L330.7 348.9L327.9 348.0L326.8 343.0L328.7 339.4L333.7 336.2L336.8 335.6L337.6 333.4L336.8 332.0L342.8 336.2L345.7 336.0L347.9 334.3L352.5 334.5L359.2 332.9L359.4 338.2L358.4 341.3L361.6 342.4L363.7 345.0L365.7 345.5L365.0 348.0L363.9 347.8L362.0 349.5ZM386.7 361.5L385.0 366.5L381.8 365.1L380.0 366.4L379.5 363.7L375.2 364.8L372.7 363.0L369.3 363.1L368.3 364.3L366.1 362.7L363.8 360.1L365.0 358.5L364.2 357.4L364.8 355.9L363.2 351.5L361.8 351.8L362.0 349.5L363.9 347.8L365.0 348.0L365.7 345.5L362.8 343.8L363.6 342.2L366.2 340.3L372.2 342.7L372.0 343.7L374.6 345.2L374.9 348.4L380.6 352.5L379.1 354.7L382.4 357.2L381.2 358.9L381.4 360.6L383.4 360.5L384.6 358.2L386.6 360.0ZM397.1 362.0L395.9 361.6L393.2 363.0L392.7 364.4L389.8 363.8L386.7 361.5L386.6 360.0L390.2 357.1L392.6 353.6L400.2 353.2L401.1 350.6L403.5 349.3L402.4 345.8L405.2 344.9L406.1 348.5L409.4 349.4L410.4 352.8L408.7 354.4L403.2 355.3L399.0 358.7ZM327.9 348.0L330.7 348.9L334.4 345.3L336.1 346.2L338.1 344.3L344.2 343.2L344.5 345.1L342.7 349.3L341.6 348.5L336.2 349.1L335.9 353.6L332.3 354.0L330.2 350.9L327.4 349.9L325.7 352.2L325.2 351.0ZM307.8 364.9L307.8 363.3L314.7 362.9L317.1 364.3L319.5 363.7L321.1 358.3L323.7 359.0L322.6 355.4L323.5 352.7L325.7 352.2L327.4 349.9L330.2 350.9L334.2 356.9L333.9 359.9L334.7 361.4L333.2 364.7L332.0 363.4L328.2 363.9L323.6 367.1L324.0 369.4L321.5 370.2L320.0 368.5L318.2 370.1L315.0 370.4L313.1 369.0L309.1 373.1L308.9 366.4ZM401.7 327.5L401.8 333.3L403.5 334.4L404.4 337.1L403.0 341.2L405.2 344.9L402.4 345.8L402.2 347.2L400.1 346.0L400.7 343.2L396.7 341.8L394.3 339.1L396.6 336.6L396.8 334.9L395.1 334.3L395.9 332.9L395.1 330.9L392.9 330.9L392.5 329.1L390.8 327.8L391.8 325.8L397.9 329.2L400.4 329.2ZM360.8 353.3L362.1 355.0L361.0 356.3L357.6 357.5L358.0 360.1L355.6 360.9L356.5 362.4L355.2 363.1L352.3 363.0L352.7 360.6L351.6 359.6L348.6 361.2L348.2 358.4L345.1 358.2L345.3 356.0L347.9 353.9L347.2 351.3L349.9 351.0L350.6 349.5L353.3 351.6L358.3 351.5L358.8 353.4ZM342.5 368.6L339.6 366.1L337.0 365.7L336.3 363.9L334.3 362.9L334.2 356.9L332.3 354.0L335.9 353.6L336.2 349.1L341.6 348.5L342.7 349.3L344.5 345.1L344.2 343.2L342.3 343.7L342.2 341.7L346.7 339.7L347.7 348.3L349.0 349.6L350.6 349.5L349.9 351.0L347.2 351.3L347.9 353.9L345.3 356.0L345.1 358.2L348.2 358.4L348.6 361.2L351.6 359.6L352.7 360.6L352.3 363.0L350.7 364.8L346.8 365.0L344.5 367.9ZM556.8 199.4L555.8 202.6L556.6 205.2L560.9 210.1L555.0 212.3L551.4 210.2L549.8 210.6L549.7 208.5L545.2 205.3L546.0 204.2L544.9 201.7L541.5 199.8L540.1 200.9L536.8 200.9L534.2 199.1L531.8 199.3L531.0 197.3L533.0 195.6L534.3 192.5L547.5 190.9L548.8 192.3L548.3 193.9L550.6 195.2L551.4 197.1L553.4 197.0ZM608.9 215.6L603.9 214.3L604.0 211.4L601.8 209.4L599.4 208.9L596.7 205.7L597.7 203.2L596.9 201.4L597.3 198.4L604.7 196.6L607.1 198.8L609.6 198.5L611.7 200.1L613.7 198.9L619.6 202.2L618.7 205.8L620.0 207.1L619.7 209.2L617.2 209.4L609.7 216.0ZM585.8 213.2L586.0 215.8L588.0 217.5L588.1 220.0L589.4 220.0L591.0 221.5L589.8 222.9L582.9 224.7L579.7 224.4L578.9 223.2L579.5 221.6L576.0 221.1L574.1 219.8L576.2 216.6L575.2 214.4L577.8 213.4L582.3 208.2L590.0 208.1ZM580.3 209.9L577.8 213.4L569.4 214.9L561.5 222.3L560.8 221.3L558.8 222.4L556.5 222.0L554.8 222.7L554.8 220.0L552.8 216.1L553.3 215.5L558.1 217.3L572.3 211.4ZM605.4 214.9L608.9 215.6L610.4 218.3L609.7 219.6L615.3 226.7L613.1 227.9L608.8 225.6L609.3 224.0L607.5 222.0L603.5 221.6L601.4 223.1L597.1 223.2L592.5 224.9L589.5 228.8L586.4 229.6L585.2 232.3L583.9 232.1L583.9 230.3L582.8 229.3L583.6 226.0L582.9 224.7L587.6 223.8L591.0 221.5L589.4 220.0L588.1 220.0L588.0 217.5L590.7 217.2L591.9 216.1L597.2 216.9L598.5 215.2L604.2 215.6ZM575.2 214.4L576.2 216.6L574.1 219.8L576.0 221.1L579.5 221.6L579.0 226.9L577.4 227.8L577.0 226.2L575.5 225.4L572.7 227.1L569.1 226.1L568.8 224.8L566.0 227.9L564.4 227.0L564.6 224.5L561.9 224.3L560.9 226.6L560.2 225.8L563.0 220.3L569.4 214.9L573.6 213.9ZM506.2 218.8L504.7 219.4L498.6 218.9L496.6 215.5L498.0 213.9L498.1 211.7L502.3 211.5L503.4 213.5L507.1 212.6L511.4 210.3L512.5 212.2L515.9 211.4L514.3 214.6L510.6 215.6ZM547.1 218.3L552.9 218.4L544.2 226.6L544.7 228.4L541.2 230.7L535.5 230.9L533.7 228.1L531.7 226.8L533.7 226.0L531.9 224.2L535.0 220.5L532.0 216.2L534.1 214.7L538.0 215.1L540.2 217.8L539.1 219.2L539.8 221.6L541.2 222.3L544.8 221.7L545.2 219.4ZM553.9 218.1L554.8 222.7L556.5 222.0L558.8 222.4L560.8 221.3L561.5 222.3L560.2 225.8L559.0 227.2L553.2 228.3L552.3 229.9L543.4 233.5L541.9 230.1L544.7 228.4L544.2 226.6ZM579.7 224.4L582.9 224.7L583.6 226.0L582.8 229.3L583.9 230.3L583.9 232.1L582.3 234.2L580.0 234.1L577.2 231.4L577.4 227.8L579.0 226.9ZM560.9 226.6L561.9 224.3L564.6 224.5L564.4 227.0L566.0 227.9L565.8 229.6L563.8 231.5L567.8 231.1L563.9 235.7L560.7 236.2L559.1 235.2L558.3 231.4L556.7 231.1ZM556.7 231.1L558.3 231.4L559.1 235.2L560.7 236.2L556.8 238.8L556.8 240.6L554.9 242.3L554.3 240.0L552.2 242.4L551.8 244.8L550.3 241.7L550.4 238.1L547.7 235.3L548.3 233.0ZM565.0 234.6L567.8 231.1L569.1 231.9L573.6 229.2L574.4 237.2L570.8 242.5L572.1 244.8L568.2 245.8L566.9 239.6L564.8 237.7L565.9 236.4ZM528.4 231.7L528.5 235.7L530.2 239.2L526.1 238.6L518.1 241.4L513.2 241.3L513.6 231.6L517.9 230.9L520.2 229.2L524.7 230.0ZM543.4 233.5L546.0 233.0L547.5 231.7L548.3 233.0L547.7 235.3L550.4 238.1L550.3 241.7L551.8 244.8L548.3 248.8L547.3 253.5L547.8 254.7L545.8 256.3L543.6 256.2L543.5 253.5L546.0 244.3L543.9 242.7L541.9 238.9L533.4 239.6L533.6 236.9L535.5 235.6L540.9 236.7L540.8 234.8ZM563.9 235.7L565.0 234.6L565.9 236.4L564.8 237.7L566.9 239.6L567.6 244.1L566.2 244.0L564.5 242.3L563.3 242.6ZM560.7 236.2L563.9 235.7L563.3 242.6L562.1 245.0L560.2 246.2L560.6 247.8L556.2 248.3L554.9 242.3L556.8 240.6L556.8 238.8ZM513.6 231.6L512.9 234.2L513.6 239.5L509.7 237.9L507.5 238.0L506.0 240.4L501.3 241.1L499.7 234.2L503.7 231.8L507.1 233.5L508.3 232.4ZM482.1 239.9L479.4 242.2L474.6 241.2L474.3 243.0L472.1 242.7L470.7 233.1L472.4 232.1L475.2 234.1L478.8 235.2L483.7 235.0ZM499.7 234.2L501.3 241.8L500.1 241.3L498.0 242.6L497.0 239.7L490.1 240.8L489.7 238.9L488.0 238.2L487.1 240.0L484.6 240.8L484.6 239.3L482.1 239.9L483.7 235.0L486.9 234.0L491.9 235.0L493.4 233.6L495.3 234.7ZM514.8 241.0L518.1 241.4L526.1 238.6L529.2 239.6L533.6 236.9L533.4 239.6L527.4 241.2L525.9 245.2L523.2 246.8L522.9 248.2L520.8 249.9L521.9 251.5L517.5 253.4L518.5 250.9L517.2 249.3L514.3 250.2L514.8 246.6L516.6 246.5L518.0 245.1L517.6 244.0L515.3 244.3ZM470.7 233.1L471.6 236.4L471.1 239.8L472.6 244.5L469.7 247.0L469.3 245.2L466.1 245.0L466.5 247.9L464.1 249.9L461.1 249.7L460.6 247.7L461.6 245.1L460.1 243.3L460.6 236.0L464.6 236.1L467.7 235.1L468.5 233.2ZM513.6 239.5L513.2 241.3L509.3 243.4L506.8 246.5L500.6 248.0L502.2 244.3L501.3 241.1L502.3 240.3L506.0 240.4L507.5 238.0L509.7 237.9ZM567.6 244.1L568.2 245.8L569.5 246.1L572.1 244.8L571.9 250.3L573.4 251.1L568.9 255.5L566.2 251.8L562.4 252.3L563.2 251.7L563.2 248.8L560.6 247.8L560.2 246.2L562.1 245.0L563.3 242.6L564.5 242.3L566.2 244.0ZM481.5 240.7L484.6 239.3L484.6 240.8L487.1 240.0L488.0 238.2L489.7 238.9L489.7 242.8L490.6 245.0L489.3 246.7L489.5 248.7L483.6 250.6L478.7 249.8L477.4 247.4L481.8 242.9ZM490.1 240.8L497.0 239.7L495.9 245.1L493.5 246.7L492.3 249.0L489.5 248.7L489.3 246.7L490.6 245.0L489.7 242.8ZM513.2 241.3L514.8 241.0L515.3 244.3L517.6 244.0L518.0 245.1L516.6 246.5L514.8 246.6L514.3 250.2L511.7 251.3L509.6 251.0L509.5 249.0L507.2 248.3L505.5 246.5L506.8 246.5L509.3 243.4ZM497.0 240.7L498.0 242.6L500.1 241.3L501.3 241.8L502.2 244.3L500.6 248.0L496.8 249.3L496.7 250.4L500.1 252.0L500.1 253.5L497.0 252.0L494.2 256.4L491.9 256.4L489.5 259.1L487.9 256.2L484.9 256.8L486.9 252.5L483.6 250.6L486.5 249.1L492.3 249.0L493.5 246.7L495.9 245.1ZM472.1 242.7L474.3 243.0L474.6 241.2L479.4 242.2L481.5 240.7L481.8 242.9L477.4 247.4L477.8 248.1L473.5 248.9L471.9 248.1L473.8 246.0ZM460.3 244.0L461.6 245.1L460.6 247.7L461.1 249.7L464.1 249.9L466.5 247.9L466.1 245.0L469.3 245.2L469.7 247.0L472.6 244.5L473.8 246.0L471.9 248.1L469.2 248.6L467.3 250.4L467.4 251.7L465.4 251.9L460.3 255.5L461.2 254.0L457.3 247.8ZM500.6 248.0L503.4 247.8L505.5 246.5L507.2 248.3L509.5 249.0L509.6 251.0L508.8 251.8L503.1 251.8L502.3 250.1L500.1 252.0L496.7 250.4L496.8 249.3ZM562.4 252.3L566.2 251.8L568.9 255.5L569.8 255.3L570.3 258.7L567.0 262.2L565.6 258.2L563.4 258.3L560.3 255.3ZM477.8 248.1L478.7 249.8L483.6 250.6L486.9 252.5L485.3 255.4L483.7 254.0L478.8 253.8L477.0 254.6L472.8 253.0L466.2 253.7L467.3 250.4L469.2 248.6ZM509.6 251.0L511.7 251.3L509.4 253.5L510.2 254.9L508.7 261.0L502.4 260.9L501.5 257.6L494.5 257.6L491.9 256.4L494.2 256.4L497.0 252.0L500.1 253.5L500.1 252.0L502.3 250.1L503.1 251.8L508.8 251.8ZM528.5 257.9L530.1 259.2L531.5 263.2L534.1 263.2L536.8 265.7L537.1 268.6L535.2 270.1L532.3 276.2L528.4 275.5L527.4 274.2L523.6 274.4L521.2 273.0L517.4 273.3L519.1 273.0L522.7 269.1L519.8 265.7L518.0 265.0L519.9 262.5L521.5 262.0L523.3 258.7L527.5 257.3ZM552.0 284.0L551.2 278.5L551.8 277.3L559.5 274.8L562.1 273.0L563.9 273.0L564.0 274.8L561.7 279.3L561.6 281.2L559.8 282.2L559.6 283.8L556.2 287.1L554.3 284.7ZM484.9 256.8L487.9 256.2L489.5 259.1L491.9 256.4L494.5 257.6L501.5 257.6L502.3 260.2L501.7 261.7L496.2 263.9L495.7 265.2L488.3 265.0L487.7 266.5L485.3 267.9L483.2 267.6L481.4 268.6L479.4 266.9L480.3 264.0L479.7 263.0L482.1 261.7ZM537.1 268.6L536.8 265.7L534.1 263.2L536.2 261.2L538.0 260.2L546.8 260.3L546.9 262.8L543.3 263.3L544.2 265.0L539.9 271.1L537.7 270.0ZM544.2 265.0L543.3 263.3L548.4 262.6L550.6 263.7L552.3 263.1L552.7 264.3L555.1 264.1L558.6 261.8L556.3 267.4L552.8 271.6L548.2 271.2L548.0 267.0ZM518.0 265.0L519.8 265.7L522.7 269.1L521.7 270.9L519.1 273.0L516.1 273.1L515.0 274.9L510.2 272.7L509.6 267.5ZM502.3 260.2L502.4 260.9L508.7 261.0L508.3 268.6L505.7 269.2L506.3 271.3L500.6 271.7L498.4 272.7L497.5 271.6L495.2 272.5L492.8 271.6L492.3 270.3L496.4 269.0L496.2 263.9L501.7 261.7ZM532.3 276.2L535.2 270.1L537.1 268.6L537.7 270.0L539.9 271.1L544.2 265.0L546.6 266.6L546.2 269.4L543.5 272.1L544.4 275.1L543.1 280.3L535.7 279.6L533.0 283.4L531.7 282.5L530.8 280.2ZM479.7 263.0L480.3 264.0L479.4 266.9L483.9 270.0L484.1 271.6L473.0 272.3L468.6 271.0L470.4 265.0L477.2 266.1ZM524.9 289.5L522.9 288.8L520.1 284.9L519.3 278.9L516.9 278.3L515.0 274.9L516.1 273.1L527.4 274.2L528.4 275.5L532.3 276.2L530.8 280.2L529.4 279.9L529.0 285.4L527.3 289.0ZM546.5 277.6L547.1 282.0L545.9 284.4L546.4 285.8L545.5 286.4L543.4 281.7L545.7 275.9ZM516.9 278.3L518.2 278.5L516.7 283.3L515.8 291.6L514.9 292.7L512.2 292.8L510.6 291.9L511.5 290.2L510.5 286.8L512.2 281.6L510.9 277.9L512.3 277.5L514.1 279.0ZM522.1 286.7L518.9 292.6L517.6 292.8L517.2 294.4L515.2 294.9L514.9 292.7L516.7 288.0L516.7 283.3L518.2 278.5L519.3 278.9L520.1 284.9ZM545.5 286.4L546.4 285.8L548.8 286.5L549.1 288.2L547.9 287.3L546.0 290.4L546.2 292.6L545.4 293.7L544.9 288.7ZM543.4 281.7L545.5 286.4L544.9 288.7L545.4 291.7L544.3 289.5L541.9 288.7ZM545.4 293.7L546.2 292.6L546.0 290.4L547.9 287.3L551.4 289.7L551.5 291.5L552.8 292.8L552.4 295.4L553.6 295.8L551.6 301.9L548.7 300.1L546.0 299.5L544.2 300.1L542.8 298.3L543.6 295.3ZM545.5 292.2L542.8 298.3L540.3 298.2L537.8 299.7L535.1 296.4L535.5 294.2L537.3 292.7L537.5 290.1L534.9 289.7L535.3 285.5L536.6 286.1L542.7 284.3L541.9 288.7L544.3 289.5ZM512.2 292.8L512.9 294.2L512.8 300.4L511.4 302.7L511.6 304.6L509.7 305.1L507.7 298.7L510.3 294.3L507.4 292.2L507.4 289.8L508.9 288.7L509.3 286.4L510.5 286.8L511.5 290.2L510.6 291.9ZM507.7 298.7L509.7 305.1L508.5 304.5L507.2 306.6L504.6 305.5L505.4 310.3L504.7 311.5L503.1 311.0L501.8 309.5L501.0 305.5L500.6 292.8L505.3 296.7L507.5 296.2ZM493.7 312.7L491.5 314.9L490.2 310.1L488.2 307.7L488.1 304.3L494.8 304.2L496.7 303.3L493.2 307.9L492.7 311.1ZM501.0 305.5L502.5 312.8L500.8 315.2L497.6 312.0L493.7 312.7L492.7 311.1L493.2 307.9L496.8 303.9L500.0 303.7ZM551.8 244.8L552.2 242.4L554.3 240.0L556.0 245.1L555.9 252.3L547.3 253.5L548.3 248.8ZM556.2 248.3L560.6 247.8L563.2 248.8L563.2 251.7L560.3 255.3L563.4 258.3L557.9 256.1L554.1 257.5L552.8 255.6L554.3 254.1L554.1 252.7L555.9 252.3ZM563.4 258.3L565.6 258.2L566.4 259.2L567.0 262.2L565.4 264.3L562.6 265.4L560.0 264.1L559.9 262.4L558.6 261.8L557.1 263.4L552.7 264.3L552.8 260.3L554.1 257.5L557.9 256.1ZM466.2 253.7L472.8 253.0L471.7 255.8L471.9 263.1L470.4 265.0L468.6 271.0L463.1 268.8L466.8 267.0L467.2 263.2L464.0 261.7L461.5 259.0ZM483.5 258.4L484.4 258.6L482.1 261.7L479.7 263.0L477.2 266.1L470.4 265.0L471.9 263.1L472.1 258.3L474.2 259.2L475.1 257.7L476.8 258.8L481.2 256.9ZM546.9 262.8L546.8 260.3L548.7 256.3L547.8 255.5L548.3 253.0L554.1 252.7L554.3 254.1L552.8 255.6L554.1 257.5L552.3 263.1L550.6 263.7ZM543.5 253.5L543.6 256.2L545.8 256.3L547.8 254.7L547.3 253.5L548.3 253.0L547.8 255.5L548.7 256.3L547.5 259.8L541.1 260.8L538.0 260.2L542.0 256.1L541.5 255.0ZM580.4 209.8L572.3 211.4L570.9 206.6L574.5 206.0L575.7 204.4L574.3 203.6L577.6 200.0L579.3 201.2L581.4 201.3ZM597.1 198.7L597.7 203.2L596.7 205.7L599.4 208.9L601.8 209.4L604.0 211.4L603.9 214.3L605.4 214.9L604.2 215.6L598.5 215.2L597.2 216.9L596.4 212.0L595.0 209.9L591.1 210.6L585.8 213.2ZM531.8 199.3L534.2 199.1L536.8 200.9L536.9 204.5L539.2 205.8L540.2 208.2L539.7 210.9L537.9 213.0L538.0 215.1L534.1 214.7L532.0 216.2L531.3 214.6L532.2 211.7L526.7 211.2L525.7 208.8L521.5 207.7L520.0 205.6L522.9 202.5L524.9 202.5L527.6 200.2ZM547.1 218.3L545.2 219.4L544.8 221.7L541.2 222.3L539.8 221.6L539.1 219.2L540.2 217.8L538.6 215.8L543.4 213.3L543.4 211.9L546.0 211.2L547.0 215.4L549.4 216.1L549.8 218.2ZM515.7 212.1L521.5 207.7L525.7 208.8L526.7 211.2L532.2 211.7L531.3 214.6L535.0 220.5L534.1 221.9L533.5 220.8L531.2 221.3L530.6 222.8L527.2 224.2L527.4 226.7L524.9 226.9L523.1 225.9L522.9 219.7L519.9 218.9L518.5 215.6L517.2 215.6L516.9 213.2ZM506.2 218.8L510.6 215.6L514.3 214.6L515.7 212.1L516.9 213.2L517.2 215.6L518.5 215.6L519.9 218.9L522.9 219.7L523.1 225.9L519.0 226.7L518.6 228.7L519.8 229.8L517.9 230.9L508.3 232.4L508.2 229.8L506.7 228.5L507.4 223.6ZM508.7 261.0L510.7 259.3L515.3 259.2L518.3 263.0L519.9 262.5L518.0 265.0L509.6 267.5L510.2 272.7L506.3 271.3L505.7 269.2L508.3 268.6ZM481.4 268.6L483.2 267.6L485.3 267.9L487.7 266.5L488.3 265.0L491.3 265.5L496.3 264.6L496.4 269.0L490.1 271.1L484.1 271.6L483.9 270.0ZM485.3 255.4L483.5 258.4L481.2 256.9L476.8 258.8L475.1 257.7L474.2 259.2L472.1 258.3L471.7 255.8L472.8 253.0L477.0 254.6L478.8 253.8L483.7 254.0ZM464.0 261.7L467.2 263.2L466.8 267.0L463.1 268.8L460.1 268.8L459.7 266.7L461.0 263.1ZM500.8 295.2L501.0 305.5L500.0 303.7L496.8 303.9L496.4 300.5L494.7 297.0L498.1 296.1L498.7 294.2ZM496.7 303.3L489.6 304.5L489.4 301.2L491.1 298.5L494.7 297.0ZM491.5 314.9L493.7 312.7L497.6 312.0L500.8 315.2L501.8 319.8L500.0 321.5L497.0 322.0L494.4 315.3ZM507.4 289.8L507.4 292.2L510.3 294.3L507.7 298.7L507.5 296.2L505.3 296.7L504.3 295.9L504.3 290.7L505.3 289.7ZM566.0 227.9L568.8 224.8L569.1 226.1L572.7 227.1L575.5 225.4L577.0 226.2L577.3 227.4L569.1 231.9L567.8 231.1L563.8 231.5L565.8 229.6ZM560.2 225.8L560.9 226.6L559.8 228.2L554.7 232.0L552.7 231.6L548.3 233.0L547.5 231.7L552.3 229.9L553.2 228.3L559.0 227.2ZM535.5 230.9L537.1 231.4L541.9 230.1L543.4 233.5L540.8 234.8L540.9 236.7L535.5 235.6L530.2 239.2L528.5 235.7L528.4 231.7ZM521.9 251.5L520.8 249.9L522.9 248.2L525.8 248.1L525.7 249.3L529.4 251.0L532.2 256.3L531.9 259.0L530.1 259.2L526.7 255.6L524.3 255.8L522.1 253.7ZM460.3 255.5L465.4 251.9L467.4 251.7L461.5 259.0L464.0 261.7L461.0 263.1L460.5 265.0L461.1 261.5L459.9 258.9ZM533.4 239.6L541.9 238.9L543.9 242.7L546.0 244.3L543.5 253.5L541.5 255.0L542.0 256.1L534.1 263.2L531.5 263.2L530.1 259.2L531.9 259.0L532.2 256.3L529.4 251.0L525.7 249.3L525.8 248.1L522.9 248.2L523.2 246.8L525.9 245.2L527.4 241.2ZM514.3 250.2L517.2 249.3L518.5 250.9L517.5 253.4L521.9 251.5L522.1 253.7L524.3 255.8L526.7 255.6L528.5 257.9L527.5 257.3L523.3 258.7L521.5 262.0L518.3 263.0L515.3 259.2L510.7 259.3L509.2 260.3L509.2 257.0L510.2 254.9L509.4 253.5ZM536.8 200.9L540.1 200.9L541.5 199.8L544.9 201.7L546.0 204.2L545.2 205.3L547.5 207.6L546.0 211.2L543.4 211.9L543.4 213.3L539.3 215.9L538.0 215.1L537.9 213.0L539.7 210.9L540.2 208.2L539.2 205.8L536.9 204.5ZM588.0 217.5L586.0 215.8L585.8 213.2L591.1 210.6L595.0 209.9L596.4 212.0L597.2 216.9L591.9 216.1L590.7 217.2ZM577.3 227.4L577.2 231.4L580.0 234.1L578.5 236.9L575.9 237.9L574.4 237.2L573.6 229.2L574.9 229.3ZM547.1 282.0L546.5 277.6L548.9 276.4L550.8 277.4L550.0 282.7L547.7 283.2ZM530.8 280.2L531.7 284.9L530.9 286.2L529.0 285.4L529.4 279.9ZM551.8 277.3L551.2 278.5L552.0 284.0L550.8 284.1L549.7 281.4L550.8 277.4L548.9 276.4L546.5 277.6L545.7 275.9L542.7 284.2L542.5 281.4L544.4 275.1L543.5 272.1L546.2 269.4L546.6 266.6L548.0 267.0L548.2 271.2L552.8 271.6L551.1 272.9ZM546.4 285.8L545.9 284.4L547.1 282.0L547.7 283.2L550.0 282.7L550.8 284.1L549.9 286.0L548.8 286.5ZM552.0 284.0L554.3 284.7L556.9 288.4L553.6 295.8L552.4 295.4L552.8 292.8L551.5 291.5L551.4 289.7L549.1 288.2L548.8 286.5L550.8 284.1ZM562.1 273.0L559.5 274.8L551.8 277.3L551.1 272.9L556.3 267.4L558.6 261.8L559.9 262.4L560.0 264.1L562.6 265.4L560.5 270.7ZM531.7 282.5L533.0 283.4L535.7 279.6L543.1 280.3L542.7 284.3L536.6 286.1L533.1 285.5L530.9 286.2ZM531.5 297.3L526.6 295.8L528.2 286.0L535.3 285.5L534.9 289.7L537.5 290.1L537.3 292.7L535.5 294.2L535.1 296.4L532.5 297.9ZM577.6 200.0L574.3 203.6L571.7 201.1L571.2 195.1L565.6 194.2L564.4 194.9L564.1 191.5L561.4 189.0L559.4 188.6L555.5 184.7L552.2 183.4L554.4 181.0L555.9 180.8L557.5 179.1L560.1 179.1L560.6 177.9L562.7 176.8L566.8 180.1L574.8 181.8L574.6 184.3L576.4 185.9L576.3 187.9L577.8 190.2L576.9 191.0L577.8 195.0L576.6 196.3L576.6 198.8ZM555.0 212.3L560.9 210.1L563.5 211.1L564.1 212.3L567.4 211.2L568.5 207.6L567.2 206.8L567.9 206.0L570.9 206.6L572.3 211.4L558.1 217.3L553.3 215.5L552.8 216.1L555.8 213.5ZM577.8 195.0L576.9 191.0L577.8 190.2L576.3 187.9L576.4 185.9L574.6 184.3L575.2 183.2L579.1 183.5L580.9 179.5L592.5 173.8L595.8 179.0L599.0 178.5L598.0 180.5L595.2 181.2L593.8 183.2L595.9 185.7L598.0 183.6L599.9 183.5L601.6 185.1L603.5 189.0L603.0 190.8L600.6 191.5L601.1 192.7L597.5 195.5L596.0 194.9L589.5 196.1L589.5 198.0L587.5 198.8L587.1 200.9L583.8 198.9L581.6 198.7L581.7 196.6ZM597.3 198.4L590.0 208.1L582.3 208.2L580.4 209.8L581.4 201.3L579.3 201.2L577.6 200.0L576.6 198.8L576.6 196.3L577.8 195.0L581.7 196.6L581.6 198.7L583.8 198.9L587.1 200.9L587.5 198.8L589.5 198.0L589.5 196.1L590.9 195.5L592.8 196.0L596.0 194.9L597.5 195.5L598.5 196.1ZM574.3 203.6L575.7 204.4L574.5 206.0L570.9 206.6L567.9 206.0L568.4 204.2L567.2 202.1L564.7 202.9L561.3 201.2L558.6 198.1L560.1 195.9L559.6 193.7L558.2 192.8L558.6 191.4L556.3 186.8L556.6 185.9L564.1 191.5L564.4 194.9L565.6 194.2L571.2 195.1L571.7 201.1ZM556.8 199.4L553.4 197.0L551.4 197.1L550.6 195.2L548.3 193.9L548.8 192.3L547.5 190.9L548.2 187.9L551.6 185.3L552.2 183.4L556.6 185.9L556.3 186.8L558.6 191.4L558.2 192.8L559.6 193.7L560.1 195.9ZM534.1 221.9L531.9 224.2L533.7 226.0L531.7 226.8L533.7 228.1L535.5 230.9L527.6 231.9L524.7 230.0L519.8 229.8L518.6 228.7L519.0 226.7L521.9 225.5L524.9 226.9L527.4 226.7L527.2 224.2L530.6 222.8L531.2 221.3L533.5 220.8ZM558.6 206.6L556.6 205.2L555.8 202.6L556.8 199.4L558.6 198.1L561.3 201.2L564.7 202.9L567.2 202.1L568.4 204.2L567.9 206.0L565.6 207.5L561.0 205.7ZM560.9 210.1L558.6 206.6L561.0 205.7L568.5 207.6L567.4 211.2L564.1 212.3L563.5 211.1ZM547.5 207.6L549.7 208.5L549.8 210.6L551.4 210.2L555.8 213.5L552.8 216.1L553.9 218.1L547.1 218.3L549.8 218.2L549.4 216.1L547.0 215.4L546.0 211.2ZM431.5 221.4L428.8 218.1L427.2 218.3L423.7 215.8L425.4 211.9L425.6 208.5L434.4 204.9L437.3 205.4L439.1 207.3L440.1 210.4L439.2 214.7L437.5 217.1L438.1 220.4ZM423.7 215.8L427.2 218.3L427.2 226.9L424.5 227.4L422.1 224.9L422.2 219.0ZM427.2 218.3L428.8 218.1L431.5 221.4L430.6 223.9L432.5 225.9L430.3 228.3L427.2 226.9ZM431.5 221.4L438.1 220.4L440.7 222.6L436.5 226.8L432.5 225.9L430.6 223.9ZM347.2 239.0L343.6 237.2L342.1 234.4L337.9 232.5L337.6 228.9L335.8 228.0L336.4 226.0L333.0 220.2L336.9 220.0L338.6 218.2L342.0 219.7L342.5 221.0L349.7 222.2L351.2 224.9L350.2 228.9L352.6 229.9L351.0 231.9L352.3 234.2L352.0 235.8ZM422.1 224.9L424.5 227.4L427.2 226.9L430.3 228.3L430.2 231.0L431.4 232.7L430.4 234.1L430.8 236.7L428.3 241.4L426.3 242.3L423.6 240.3L425.2 235.2L423.8 231.1L421.1 227.5ZM429.7 238.1L431.4 232.7L437.4 231.6L438.0 230.1L439.7 230.7L440.8 230.1L444.3 232.5L443.3 238.4L444.7 239.2L445.0 240.8L441.6 240.5L440.9 242.4L437.4 245.3L430.2 239.9ZM347.2 239.0L352.0 235.8L352.3 234.2L351.0 231.9L354.0 229.5L356.9 230.9L358.4 232.8L361.8 232.9L361.5 234.6L363.2 235.6L363.4 238.0L361.5 241.7L363.9 244.7L362.0 246.3L358.7 245.0L355.7 245.3L354.9 246.8L350.4 242.1L349.1 242.8ZM363.4 238.0L363.2 235.6L364.8 235.9L371.4 233.3L373.2 234.3L373.1 237.4L375.0 239.1L375.2 240.4L374.0 242.8L372.7 242.6L372.7 245.9L368.4 244.1L367.5 245.9L365.7 245.1L365.4 241.2L366.3 241.0L364.5 237.7ZM363.9 244.7L361.5 241.7L363.4 238.0L364.5 237.7L366.3 241.0L365.4 241.2L365.7 245.1ZM345.4 238.4L347.2 239.0L349.1 242.8L350.4 242.1L355.6 247.9L352.8 248.0L350.4 246.0L344.3 244.4L342.2 245.7L334.6 242.6L334.8 241.3L337.7 240.2L339.0 238.4L342.5 239.1ZM387.2 259.2L385.1 261.0L380.7 258.6L375.6 260.9L371.5 265.1L369.3 264.6L368.9 263.3L370.7 259.4L371.2 253.8L372.8 250.8L373.0 251.7L376.1 252.1L380.9 255.1L384.0 255.5L383.7 257.9ZM422.7 264.6L420.4 263.3L416.3 266.1L416.7 267.8L418.0 268.8L417.4 270.1L407.9 266.3L404.9 267.0L404.7 265.2L405.8 262.8L405.2 259.7L413.9 260.0L414.6 258.5L417.7 257.6L417.6 255.3L419.0 255.8L420.6 259.0L422.3 259.6L421.7 262.0ZM385.1 261.0L387.7 258.9L389.4 261.5L394.3 260.5L396.7 261.5L396.3 265.4L397.1 267.3L394.3 269.8L392.6 269.4L392.5 266.7L390.4 264.4L389.7 265.9L387.9 264.2L386.4 264.7L384.9 263.2ZM343.3 260.4L343.7 258.6L348.4 260.1L349.3 258.9L354.5 259.6L351.6 268.5L346.4 271.9L343.4 270.7L343.4 268.3L342.4 267.4L344.3 264.8ZM335.1 262.9L338.3 261.1L338.2 259.7L339.6 258.9L340.9 260.3L344.0 261.0L343.5 263.7L344.3 264.8L342.4 267.4L337.6 267.0L337.3 269.5L334.4 269.0L331.9 266.5ZM430.7 262.0L435.0 264.7L438.5 263.6L439.0 267.3L442.2 269.4L440.5 271.5L430.5 271.0L430.8 268.6L427.0 268.7L425.1 267.7L424.8 265.9ZM415.1 269.7L417.4 270.1L418.0 268.8L416.7 267.8L416.3 266.1L420.4 263.3L422.7 264.6L424.4 269.4L430.8 268.6L430.8 273.8L429.6 276.3L427.2 278.1L426.1 276.5L424.2 277.5L425.0 278.8L421.8 282.7L419.0 278.9L420.4 277.4L420.3 276.2L416.9 273.6L416.6 269.9ZM372.7 245.9L372.7 242.6L374.0 242.8L375.1 243.4L375.0 246.1L376.4 246.5L378.5 245.4L379.2 247.1L381.3 248.4L382.1 247.4L385.9 249.2L387.3 253.3L385.6 256.0L387.7 257.5L387.7 258.9L383.7 257.9L384.0 255.5L380.9 255.1L376.1 252.1L373.0 251.7ZM395.8 261.1L394.3 260.5L389.4 261.5L387.7 258.9L387.7 257.5L385.6 256.0L387.1 252.0L388.2 251.1L390.2 253.0L392.6 252.4L394.0 255.8L396.8 255.7L397.4 259.1ZM411.8 268.0L414.8 268.5L415.1 269.7L416.6 269.9L416.9 273.6L420.3 276.2L420.4 277.4L417.9 281.4L410.9 280.2L409.7 277.0L411.1 274.2L410.1 273.2L410.2 271.1ZM326.8 268.5L331.1 266.2L334.4 269.0L333.1 270.0L332.6 276.5L330.6 279.3L332.0 281.1L330.9 283.1L331.1 285.4L330.2 286.1L326.5 285.9L325.8 281.6L323.8 280.6L322.8 271.4ZM342.4 267.4L343.4 268.3L343.4 270.7L346.4 271.9L338.4 280.8L335.3 285.9L330.6 286.9L326.1 286.4L331.1 285.4L330.9 283.1L332.0 281.1L330.6 279.3L333.1 274.7L332.5 273.0L333.1 270.0L334.4 269.0L337.3 269.5L337.6 267.0ZM410.2 271.1L410.1 273.2L411.1 274.2L409.7 277.0L410.9 280.2L409.8 280.3L410.1 283.6L409.0 286.7L408.1 286.9L406.6 284.8L401.2 284.1L402.4 279.0L403.4 278.3L403.2 275.2L406.3 273.6L407.0 270.5ZM389.6 276.0L391.9 284.5L389.5 286.8L388.5 289.9L384.8 287.8L385.3 285.6L381.7 284.6L381.6 282.5L378.4 278.6L376.7 278.5L375.4 275.9L381.4 276.2L386.8 272.3L387.1 274.3ZM403.8 274.3L403.4 278.3L402.4 279.0L401.2 284.1L399.1 284.0L398.2 285.9L394.9 284.3L392.7 285.4L391.9 284.5L389.6 276.0L391.3 276.8L392.7 276.0L393.6 273.0L400.4 273.1ZM375.4 275.9L376.7 278.5L378.4 278.6L380.1 280.6L377.5 281.6L373.4 279.7L371.9 283.3L372.1 284.9L369.2 286.3L365.7 285.8L364.2 284.0L364.7 282.2L363.6 280.6L366.0 276.4L371.2 273.4L374.0 273.8L374.0 275.4ZM366.0 276.4L363.6 280.6L364.7 282.2L364.2 284.0L365.7 285.8L359.6 290.0L358.4 288.9L355.7 289.9L355.4 287.8L353.6 286.3L349.0 289.3L347.1 291.8L343.8 289.4L343.5 287.2L344.9 285.5L346.8 286.7L348.2 286.4L349.3 283.4L351.9 282.4L352.1 281.0L350.1 278.5L351.8 277.3L352.1 274.3L355.8 274.1L359.7 275.8L360.4 274.4L362.0 274.3L365.2 275.1ZM419.0 278.9L424.1 286.8L433.1 291.4L435.7 291.0L437.0 292.6L437.1 295.0L434.6 297.7L432.1 298.6L429.9 302.5L426.3 302.9L422.6 304.7L421.5 302.0L420.6 302.8L418.0 300.5L419.1 294.7L420.9 294.1L421.6 291.3L420.4 290.6L420.8 288.9L418.9 285.5L419.3 281.9L417.9 281.4ZM410.9 280.2L419.3 281.9L418.9 285.5L416.9 285.1L417.1 289.7L415.0 292.5L412.4 292.5L412.0 286.5L409.0 286.7L410.1 283.6L409.8 280.3ZM389.7 265.9L390.4 264.4L392.5 266.7L392.6 269.4L394.3 269.8L392.7 276.0L391.3 276.8L387.1 274.3L386.4 268.8L384.5 269.2ZM380.1 280.6L381.6 282.5L381.7 284.6L385.3 285.6L384.8 287.8L388.5 289.9L389.8 293.3L389.2 294.9L391.0 295.2L389.1 297.9L386.9 297.1L385.1 300.3L383.6 299.8L381.3 301.0L380.4 300.2L378.2 300.8L375.5 296.8L372.8 296.1L372.2 294.4L375.7 293.4L376.1 286.4L374.1 284.5L375.1 281.9ZM401.2 284.1L406.6 284.8L408.1 286.9L412.0 286.5L412.4 292.5L413.6 292.6L413.5 294.7L409.4 298.8L407.2 297.6L403.0 297.6L402.1 296.0L398.7 295.3L401.1 293.3L400.5 291.4L399.3 291.3L397.7 288.2L399.1 284.0ZM418.9 285.5L420.8 288.9L420.4 290.6L421.6 291.3L420.9 294.1L419.1 294.7L418.0 300.5L420.6 302.8L419.4 304.8L420.4 306.8L419.8 308.2L418.0 308.7L415.1 307.1L409.5 307.2L408.8 305.6L403.4 304.1L401.9 302.7L406.3 300.7L405.0 297.8L407.2 297.6L409.4 298.8L413.5 294.7L413.6 292.6L415.0 292.5L417.1 289.7L416.9 285.1ZM388.5 289.9L389.5 286.8L391.9 284.5L392.7 285.4L394.9 284.3L398.2 285.9L397.7 288.2L399.3 291.3L400.5 291.4L401.1 293.3L398.7 295.3L399.5 295.9L399.1 297.5L389.2 294.9L389.8 293.3ZM333.9 286.3L335.3 285.9L336.8 283.8L338.6 286.2L343.5 287.2L343.8 289.4L347.1 291.8L346.1 292.9L348.9 295.2L348.6 300.1L346.5 299.4L339.5 300.6L337.0 302.9L335.3 301.4L333.7 296.8L336.4 296.0L332.7 291.8ZM347.1 291.8L349.0 289.3L353.6 286.3L355.4 287.8L355.7 289.9L358.4 288.9L359.9 290.9L364.6 291.3L363.5 293.5L361.1 293.2L361.3 296.1L363.0 297.6L362.0 299.4L358.3 298.6L358.9 305.0L356.5 304.5L356.4 299.8L348.3 298.2L348.9 295.2L346.1 292.9ZM326.1 286.4L333.9 286.3L332.7 291.8L336.4 296.0L333.7 296.8L335.3 301.4L337.0 302.9L336.9 306.8L335.3 308.3L331.8 307.3L330.4 302.5L327.5 301.5L324.6 296.5L322.4 296.1L324.1 292.3L323.9 287.1ZM425.0 303.9L429.9 302.5L432.1 298.6L436.1 296.2L437.0 297.6L436.8 300.4L434.1 301.6L433.2 305.0L433.8 307.2L437.4 311.0L435.1 323.1L429.9 322.5L431.9 318.8L428.2 314.0L429.5 309.9L426.7 307.1L424.1 306.1ZM389.1 297.9L391.0 295.2L399.1 297.5L400.3 295.1L403.0 297.6L405.0 297.8L406.3 300.7L401.9 302.7L398.3 301.1L395.9 302.2ZM381.3 301.0L383.6 299.8L385.1 300.3L386.9 297.1L391.3 298.8L395.9 302.2L396.4 303.8L395.9 305.3L389.2 306.7L386.8 305.0L384.1 305.0ZM337.0 302.9L339.5 300.6L346.5 299.4L348.6 300.1L348.8 298.7L350.8 298.4L356.4 299.8L357.1 304.9L356.1 306.0L356.6 307.3L354.7 309.1L350.2 306.0L346.5 306.2L343.3 308.0L343.5 310.0L341.8 311.0L342.1 312.5L340.7 313.8L338.2 313.6L335.8 312.4L336.9 309.7ZM377.1 299.6L378.2 300.8L380.4 300.2L384.1 305.0L386.8 305.0L389.2 306.7L386.7 309.1L385.5 311.8L382.7 310.5L382.3 308.5L380.7 308.3L379.8 309.9L376.5 310.2L373.7 306.1L370.8 306.8L370.1 301.8L371.4 299.4L373.4 299.5L375.4 302.0ZM371.4 299.4L370.1 301.8L370.8 306.8L373.7 306.1L376.5 310.2L375.9 311.4L370.5 310.1L369.0 308.4L365.9 308.7L362.1 306.9L362.6 305.6L366.7 305.6L367.7 303.8L366.8 303.1ZM420.6 302.8L421.5 302.0L422.6 304.7L425.0 303.9L424.1 306.1L426.7 307.1L429.5 309.9L428.2 314.0L430.0 316.4L428.9 318.5L424.9 317.7L422.3 321.5L417.9 322.7L415.8 322.0L415.2 321.0L416.3 319.5L409.6 312.1L409.6 310.3L411.4 307.3L415.1 307.1L418.0 308.7L419.8 308.2L420.4 306.8L419.4 304.8ZM391.8 325.8L390.8 327.8L387.8 326.0L388.1 322.9L389.5 321.4L384.0 321.3L379.9 317.9L376.5 317.4L375.1 315.1L376.5 313.2L376.5 310.2L379.8 309.9L380.7 308.3L382.3 308.5L382.7 310.5L385.5 311.8L386.7 309.1L389.2 306.7L394.0 305.3L395.9 305.3L398.0 306.7L397.0 312.4L395.6 313.5L394.1 318.5L394.8 319.2L393.9 324.0ZM357.1 304.9L358.9 305.0L359.4 306.0L362.6 305.6L362.1 306.9L363.4 308.0L369.0 308.4L370.5 310.1L375.9 311.4L376.5 313.2L375.1 315.1L376.5 317.4L374.9 319.5L375.9 323.5L373.2 322.5L370.1 323.6L369.4 322.8L371.6 321.1L370.8 318.6L362.4 316.5L362.6 315.1L359.6 315.7L358.9 318.5L356.6 321.2L355.7 320.0L355.3 312.0L356.6 309.9L355.3 308.5L356.6 307.3L356.1 306.0ZM411.2 322.9L407.7 322.3L401.5 324.6L401.7 327.5L400.4 329.2L397.9 329.2L391.8 325.8L393.9 324.0L394.8 319.2L394.1 318.5L395.6 313.5L397.0 312.4L398.0 306.7L404.5 308.5L409.6 312.1L416.3 319.5L413.7 321.7L413.5 323.5ZM344.8 307.4L346.5 306.2L350.2 306.0L354.7 309.1L355.3 308.5L356.6 309.9L355.3 312.0L355.7 313.8L353.4 314.9L349.1 314.0L345.9 312.0L346.5 308.0ZM338.2 313.6L340.7 313.8L342.1 312.5L341.8 311.0L343.5 310.0L343.3 308.0L344.8 307.4L346.5 308.0L345.9 312.0L349.1 314.0L353.4 314.9L355.7 313.8L355.7 320.0L357.5 323.1L357.3 326.1L354.7 327.6L352.1 325.8L351.3 327.1L348.4 327.3L348.4 326.1L346.1 325.0L345.4 326.6L342.8 326.4L345.0 322.9L344.5 321.7L340.4 321.7L337.9 318.6L337.4 317.3ZM443.6 341.1L438.3 336.6L438.7 335.6L436.2 334.1L434.5 330.8L433.1 329.8L432.3 331.2L430.0 331.3L428.7 329.1L429.9 322.5L435.1 323.1L436.3 315.9L438.4 315.4L441.2 316.5L439.4 322.0L441.8 324.8L440.6 326.8L441.7 328.2L440.9 329.2L442.4 334.3L442.3 337.5L443.8 339.3ZM418.2 330.5L415.8 326.3L413.4 325.5L410.9 323.6L411.2 322.9L413.5 323.5L413.7 321.7L415.2 321.0L417.9 322.7L420.0 322.5L424.4 319.5L424.9 317.7L428.9 318.5L430.0 316.4L431.9 318.8L429.9 321.5L428.7 329.1L421.1 328.0ZM356.6 321.2L358.9 318.5L359.6 315.7L362.6 315.1L362.4 316.5L370.8 318.6L371.6 321.1L369.4 322.8L370.1 323.6L368.7 325.6L365.0 323.9L361.1 325.1L359.3 331.1L358.8 326.9L357.3 326.1L357.5 323.1ZM380.3 331.3L378.1 333.0L376.8 332.1L376.1 327.8L373.1 328.4L372.6 325.6L371.4 324.6L368.7 325.6L370.1 323.6L373.2 322.5L375.9 323.5L374.9 319.5L376.5 317.4L379.9 317.9L384.0 321.3L382.9 322.4L382.1 329.7ZM394.3 339.1L389.4 336.1L388.0 336.8L386.5 336.1L384.9 333.8L381.4 332.9L380.3 331.3L382.1 329.7L382.9 322.4L384.0 321.3L389.5 321.4L388.1 322.9L387.8 326.0L392.5 329.1L392.9 330.9L395.1 330.9L395.9 332.9L395.1 334.3L396.8 334.9L396.6 336.6ZM378.1 333.0L379.5 336.4L378.4 338.4L379.6 339.7L376.8 344.4L374.4 344.4L373.6 342.4L372.2 342.7L366.2 340.3L362.8 343.8L361.6 342.4L358.4 341.3L359.4 338.2L359.3 331.1L361.1 325.1L365.0 323.9L368.7 325.6L371.4 324.6L372.6 325.6L373.1 328.4L376.1 327.8L376.8 332.1ZM426.3 331.6L423.2 335.8L423.0 338.1L420.3 337.7L419.9 334.4L418.2 330.5L421.1 328.0L428.7 329.1L428.1 331.4ZM359.2 332.9L352.5 334.5L347.9 334.3L345.7 336.0L342.8 336.2L336.8 332.0L336.8 331.0L341.6 328.3L342.8 326.4L345.4 326.6L346.1 325.0L348.4 326.1L348.4 327.3L351.3 327.1L352.1 325.8L354.7 327.6L357.3 326.1L358.8 326.9ZM430.0 331.3L428.5 332.5L426.3 331.6L428.1 331.4L428.8 329.6ZM365.7 285.8L369.2 286.3L372.1 284.9L371.9 283.3L373.4 279.7L376.8 280.9L375.1 281.9L374.1 284.5L376.1 286.4L375.7 292.2L366.2 290.2L366.5 289.1L364.8 286.8ZM359.6 290.0L363.9 286.5L366.5 289.1L366.2 290.2L375.7 292.2L375.7 293.4L372.2 294.4L372.8 296.1L375.5 296.8L377.1 299.6L375.4 302.0L373.4 299.5L371.4 299.4L366.8 303.1L367.7 303.8L366.7 305.6L359.4 306.0L358.3 298.6L362.0 299.4L363.0 297.6L361.3 296.1L361.1 293.2L363.5 293.5L364.6 291.3L359.9 290.9ZM371.4 254.7L370.7 259.4L368.9 263.3L366.8 264.3L364.4 263.9L362.6 261.9L361.2 256.5L359.2 253.4L361.5 252.5L368.5 255.0ZM373.3 263.6L375.6 260.9L380.7 258.6L385.1 261.0L384.9 263.2L386.4 264.7L387.9 264.2L389.7 265.9L383.0 269.9L376.3 265.6L376.6 264.2L374.3 263.1ZM404.7 265.2L404.9 267.0L407.9 266.3L411.8 268.0L410.2 271.1L407.0 270.5L406.3 273.6L403.8 274.3L400.4 273.1L393.6 273.0L394.3 269.8L397.1 267.3L396.3 265.4L398.4 264.3L400.5 265.1L402.2 263.7L402.3 265.7ZM386.8 272.3L381.4 276.2L378.4 275.7L378.5 271.5L376.9 269.0L380.4 267.9L383.0 269.9L386.4 268.8ZM378.4 275.7L375.4 275.9L374.0 275.4L374.0 273.8L371.2 273.4L372.3 272.3L372.6 269.6L375.4 270.7L376.8 269.7L378.5 271.5ZM346.4 271.9L351.6 268.5L353.8 270.6L354.7 270.4L354.7 274.2L352.1 274.3L351.1 273.3L346.1 272.7ZM362.0 274.3L360.4 274.4L359.7 275.8L354.7 274.2L354.2 271.0L356.7 268.9L359.4 269.4L361.3 268.6ZM371.2 273.4L366.0 276.4L365.2 275.1L362.0 274.3L361.2 272.3L361.5 266.4L365.4 266.4L366.5 265.1L369.9 266.0L371.4 267.9L373.0 267.3L374.9 269.0L376.9 269.0L375.4 270.7L372.6 269.6L372.3 272.3ZM368.9 263.3L369.3 264.6L371.5 265.1L374.3 263.1L376.6 264.2L376.3 265.6L380.4 267.9L379.3 268.8L374.9 269.0L373.0 267.3L371.4 267.9L369.9 266.0L366.5 265.1L365.4 266.4L361.5 266.4L361.3 268.6L359.4 269.4L356.7 268.9L353.8 270.6L351.6 268.5L354.5 259.6L359.2 259.4L364.4 263.9L366.8 264.3ZM352.8 248.0L355.6 247.9L356.4 252.2L359.2 253.4L361.2 256.5L362.6 261.9L357.7 258.9L350.0 259.4L349.3 258.9L349.8 257.5L346.6 255.5L345.3 252.0L348.6 252.1L349.0 251.1L351.5 250.9ZM337.8 250.5L338.4 247.8L340.3 246.7L339.9 244.2L342.2 245.7L344.3 244.4L350.4 246.0L352.8 248.0L351.5 250.9L349.0 251.1L348.6 252.1L345.3 252.0L346.6 255.5ZM363.8 244.7L367.5 245.9L368.4 244.1L372.7 245.9L373.2 249.9L371.4 254.7L368.5 255.0L361.5 252.5L359.2 253.4L356.4 252.2L354.9 246.8L355.7 245.3L358.7 245.0L362.0 246.3ZM396.3 265.4L396.7 261.5L395.8 261.1L397.4 259.1L396.8 255.7L394.0 255.8L392.6 252.4L395.6 250.7L402.1 251.4L401.8 253.6L399.9 254.0L401.0 256.2L400.6 260.6L402.2 263.7L400.5 265.1L398.4 264.3ZM401.8 253.6L402.6 248.7L401.0 245.5L401.8 243.5L401.0 240.0L402.0 242.5L403.5 243.5L409.9 242.7L414.0 244.1L413.3 244.9L415.0 249.1L413.5 252.4L411.1 253.4ZM392.6 252.4L390.2 253.0L388.2 251.1L390.3 245.9L401.0 240.0L401.8 243.5L401.0 245.5L402.6 248.7L402.1 251.4L395.6 250.7ZM374.0 242.8L375.0 239.1L375.9 239.8L379.5 237.6L390.4 240.5L394.2 243.1L390.3 245.9L387.1 252.0L385.9 249.2L382.1 247.4L381.3 248.4L379.2 247.1L378.5 245.4L376.4 246.5L375.0 246.1L375.1 243.4ZM417.6 255.3L417.7 257.6L414.6 258.5L413.9 260.0L405.2 259.7L405.8 262.8L404.7 265.2L402.3 265.7L400.3 258.6L401.0 256.2L399.9 254.0L411.1 253.4L413.5 252.4L414.6 250.4L418.4 250.6ZM460.5 236.9L460.1 243.3L458.3 242.2L457.2 243.0L453.8 242.7L452.0 241.7L450.7 242.2L445.0 240.8L444.7 239.2L443.3 238.4L444.3 233.6L447.8 234.3L450.4 233.3L452.1 234.6L454.0 234.2L458.6 236.8ZM401.9 302.7L403.4 304.1L408.8 305.6L409.5 307.2L411.4 307.3L409.6 310.3L409.6 312.1L404.5 308.5L395.9 305.3L395.9 302.2L398.3 301.1ZM430.3 228.3L433.9 225.7L439.6 227.5L439.7 230.7L438.0 230.1L437.4 231.6L431.4 232.7L430.2 231.0ZM460.1 243.3L457.3 247.8L457.9 248.9L455.9 247.6L455.3 249.3L456.2 251.3L453.9 252.6L449.8 252.5L447.3 249.9L445.7 249.6L444.5 243.9L442.1 242.6L440.5 244.1L442.1 244.8L440.9 246.3L437.4 245.3L440.9 242.4L441.6 240.5L450.7 242.2L452.0 241.7L453.8 242.7L457.2 243.0L458.3 242.2ZM426.3 242.3L428.3 241.4L431.5 242.3L432.2 244.4L429.8 245.2L428.7 247.7L425.1 249.4L423.2 254.4L424.3 257.5L426.9 257.7L430.9 260.4L430.7 262.0L424.8 265.9L425.1 267.7L427.0 268.7L424.4 269.4L421.7 262.0L422.3 259.6L420.6 259.0L417.5 254.2L418.1 251.9L427.2 245.2ZM423.6 240.3L426.3 242.3L427.2 245.2L418.1 251.9L418.4 250.6L414.6 250.4L415.0 249.1L413.3 244.9L418.1 242.8L422.0 243.8L423.4 242.4ZM336.8 283.8L338.4 280.8L346.1 272.7L351.1 273.3L352.1 274.3L351.8 277.3L350.1 278.5L352.1 281.0L351.9 282.4L349.3 283.4L348.2 286.4L346.8 286.7L344.9 285.5L343.5 287.2L338.6 286.2ZM529.7 321.1L530.2 322.9L529.1 327.9L531.3 337.7L530.5 339.9L528.2 339.5L527.1 342.8L525.8 343.7L520.9 340.6L520.5 343.1L518.6 343.8L516.7 327.1L513.8 322.7L514.2 319.9L513.3 315.8L514.5 315.2L511.7 306.6L511.4 302.7L512.8 300.4L512.9 294.2L512.2 292.8L514.9 292.7L515.2 294.9L517.2 294.4L519.2 290.8L520.6 290.3L522.1 286.7L522.9 288.8L527.3 289.0L527.6 293.1L526.6 295.8L531.5 297.3L535.5 300.1L536.1 309.2L534.7 313.7L535.4 316.2L534.6 319.6L533.5 321.1L531.3 320.0ZM372.2 342.7L373.6 342.4L374.4 344.4L376.8 344.4L379.1 341.4L379.6 339.7L378.4 338.4L379.5 336.4L379.3 334.3L378.1 333.0L380.3 331.3L381.4 332.9L384.9 333.8L386.5 336.1L388.0 336.8L389.4 336.1L394.3 339.1L396.7 341.8L399.3 342.1L400.7 343.2L400.1 346.0L402.2 347.2L402.4 345.8L405.2 344.9L406.1 348.5L409.4 349.4L410.4 352.8L408.7 354.4L403.2 355.3L399.0 358.7L397.1 362.0L397.0 364.7L399.6 370.1L399.4 371.6L396.0 372.9L397.1 372.0L400.6 372.1L401.3 373.4L395.6 377.5L394.5 379.2L395.4 382.0L390.2 385.6L386.8 390.1L383.7 392.0L372.1 395.5L365.2 399.1L357.4 405.3L352.8 410.7L351.4 409.8L349.1 412.0L347.5 411.4L343.3 417.3L338.5 418.0L337.4 416.7L334.3 416.7L332.2 412.1L331.2 412.7L331.4 414.5L329.5 410.2L325.6 411.7L325.2 413.5L322.8 412.1L322.1 413.1L324.1 415.4L319.7 418.2L318.3 417.4L315.4 421.4L316.3 422.9L315.8 424.9L317.1 426.1L314.8 427.1L311.3 425.0L310.6 427.5L306.7 429.0L307.1 427.0L304.5 423.3L301.9 425.7L301.0 434.3L299.5 435.5L297.8 434.1L294.9 433.7L286.1 439.4L281.5 439.5L281.5 437.8L283.3 435.8L284.3 429.9L287.0 427.4L289.0 427.2L292.4 422.3L292.1 420.3L296.2 418.5L297.6 417.3L298.0 414.7L299.4 414.4L297.7 407.9L298.2 402.7L296.9 402.7L296.4 400.7L294.8 400.9L295.6 394.5L293.9 394.2L293.6 392.9L291.5 391.9L291.6 388.5L293.1 387.3L299.9 390.3L301.6 393.4L303.7 391.7L305.5 392.5L306.8 394.8L309.0 393.6L309.4 389.9L302.7 388.4L303.5 379.8L301.3 376.9L302.2 375.3L301.4 369.7L304.0 370.8L307.8 364.9L307.8 363.3L314.7 362.9L317.1 364.3L319.5 363.7L321.1 358.3L323.7 359.0L322.6 355.4L323.5 352.7L325.7 352.2L325.2 351.0L327.9 348.0L326.8 343.0L328.7 339.4L333.7 336.2L336.8 335.6L337.6 333.4L336.8 332.0L342.8 336.2L345.7 336.0L347.9 334.3L352.5 334.5L359.2 332.9L359.4 338.2L358.4 341.3L361.6 342.4L362.8 343.8L366.2 340.3ZM423.9 351.5L421.9 351.2L423.5 347.7L424.8 348.1ZM442.8 349.1L444.4 352.0L443.2 352.9ZM421.2 338.5L421.6 339.4L424.1 339.8L425.2 341.7L416.8 350.9L410.4 352.8L409.4 349.4L406.1 348.5L405.2 344.9L402.4 345.8L402.2 347.2L400.1 346.0L400.7 343.2L396.7 341.8L394.3 339.1L396.6 336.6L396.8 334.9L395.1 334.3L395.9 332.9L395.1 330.9L392.9 330.9L392.5 329.1L387.8 326.0L388.1 322.9L389.5 321.4L384.0 321.3L379.9 317.9L376.5 317.4L375.1 315.1L376.5 313.2L376.5 310.2L379.8 309.9L380.7 308.3L382.3 308.5L382.7 310.5L385.5 311.8L386.7 309.1L389.2 306.7L395.9 305.3L395.9 302.2L398.3 301.1L401.9 302.7L402.8 301.8L405.0 301.7L406.3 300.7L405.0 297.8L407.2 297.6L409.4 298.8L413.5 294.7L413.6 292.6L415.0 292.5L417.1 289.7L416.9 285.1L418.9 285.5L419.3 281.9L417.9 281.4L420.4 277.4L420.3 276.2L416.9 273.6L416.6 269.9L415.1 269.7L417.4 270.1L418.0 268.8L416.7 267.8L416.3 266.1L418.4 264.2L420.4 263.3L422.7 264.6L421.7 262.0L422.3 259.6L420.6 259.0L417.5 254.2L418.1 251.9L427.2 245.2L426.3 242.3L423.6 240.3L425.2 235.2L423.8 231.1L421.1 227.5L422.1 224.9L424.5 227.4L427.2 226.9L430.3 228.3L433.9 225.7L439.6 227.5L439.7 230.7L440.8 230.1L444.3 232.5L445.2 234.4L450.4 233.3L452.1 234.6L454.0 234.2L458.6 236.8L460.5 236.9L460.9 238.8L460.3 244.0L457.3 247.8L457.9 248.9L455.9 247.6L455.3 249.3L456.2 251.3L453.9 252.6L449.8 252.5L447.3 249.9L445.7 249.6L444.5 243.9L442.1 242.6L440.5 244.1L442.1 244.8L440.9 246.3L437.4 245.3L430.2 239.9L429.7 238.1L428.3 241.4L431.5 242.3L432.2 244.4L429.8 245.2L428.7 247.7L425.1 249.4L423.2 254.4L423.4 256.7L426.9 257.7L430.9 260.4L430.7 262.0L435.0 264.7L438.5 263.6L439.0 267.3L442.2 269.4L440.5 271.5L430.5 271.0L430.8 273.8L429.6 276.3L427.2 278.1L426.1 276.5L424.2 277.5L425.0 278.8L421.8 282.7L424.1 286.8L433.1 291.4L435.7 291.0L437.0 292.6L437.1 295.0L436.1 296.2L437.0 297.6L436.8 300.4L434.1 301.6L433.2 305.0L433.8 307.2L437.4 311.0L436.3 315.9L438.4 315.4L441.2 316.5L439.4 322.0L441.8 324.8L440.6 326.8L441.7 328.2L440.9 329.2L442.4 334.3L442.3 337.5L443.8 339.3L443.9 345.2L442.6 346.3L441.0 345.4L439.3 345.8L437.4 345.1L437.4 343.7L432.5 347.4L430.7 345.9L429.6 346.4L428.9 348.5L426.9 350.3L427.7 351.6L426.8 354.2L424.3 345.1L425.9 341.7L425.1 339.8L423.0 338.1ZM437.7 347.1L438.4 351.6L437.1 351.8L435.9 349.7L436.4 346.2ZM441.2 346.0L441.3 349.3L438.6 345.7ZM552.8 216.1L553.9 218.1L548.1 222.5L544.2 226.6L544.7 228.4L541.2 230.7L527.6 231.9L524.7 230.0L520.2 229.2L517.9 230.9L510.1 232.5L508.3 232.4L508.2 229.8L506.7 228.5L507.4 223.6L506.2 218.8L504.7 219.4L498.6 218.9L496.6 215.5L498.0 213.9L498.1 211.7L502.3 211.5L503.4 213.5L507.1 212.6L511.4 210.3L512.5 212.2L517.8 211.2L521.5 207.7L520.0 205.6L520.5 204.3L522.9 202.5L524.9 202.5L527.6 200.2L531.8 199.3L531.0 197.3L533.0 195.6L534.3 192.5L547.5 190.9L548.2 187.9L551.6 185.3L554.4 181.0L555.9 180.8L557.5 179.1L560.1 179.1L560.6 177.9L562.7 176.8L566.8 180.1L574.8 181.8L575.2 183.2L579.1 183.5L580.9 179.5L592.5 173.8L595.8 179.0L599.0 178.5L598.0 180.5L595.2 181.2L593.8 183.2L595.9 185.7L598.0 183.6L599.9 183.5L601.6 185.1L603.5 189.0L603.0 190.8L600.6 191.5L601.1 192.7L597.5 195.5L598.5 196.1L597.3 198.4L604.7 196.6L607.1 198.8L609.6 198.5L611.7 200.1L613.7 198.9L619.6 202.2L618.7 205.8L620.0 207.1L619.7 209.2L617.2 209.4L609.7 216.0L608.9 215.6L610.4 218.3L609.7 219.6L615.3 226.7L613.1 227.9L608.8 225.6L609.3 224.0L607.5 222.0L603.5 221.6L601.4 223.1L597.1 223.2L591.1 225.8L589.5 228.8L586.4 229.6L585.2 232.3L583.9 232.1L582.3 234.2L580.0 234.1L578.5 236.9L575.9 237.9L574.4 237.2L573.6 229.2L579.0 226.9L579.7 224.4L580.8 223.9L582.9 224.7L587.6 223.8L591.0 221.5L589.4 220.0L588.1 220.0L588.0 217.5L586.0 215.8L585.8 213.2L590.0 208.1L582.3 208.2L580.3 209.9L572.3 211.4L558.1 217.3L553.3 215.5ZM577.3 227.4L569.1 231.9L567.8 231.1L563.9 235.7L560.7 236.2L556.8 238.8L556.8 240.6L554.9 242.3L554.3 240.0L548.3 248.8L547.3 253.5L547.8 254.7L545.8 256.3L543.6 256.2L543.5 253.5L541.5 255.0L542.0 256.1L534.1 263.2L536.8 265.7L537.1 268.6L535.2 270.1L530.8 280.2L529.4 279.9L529.0 285.4L527.3 289.0L522.9 288.8L522.1 286.7L520.6 290.3L519.2 290.8L517.2 294.4L515.2 294.9L514.9 292.7L512.2 292.8L510.6 291.9L511.5 290.2L510.5 286.8L512.2 281.6L510.9 277.9L512.3 277.5L514.1 279.0L516.9 278.3L515.0 274.9L516.1 273.1L519.1 273.0L522.7 269.1L519.8 265.7L518.0 265.0L519.9 262.5L518.3 263.0L515.3 259.2L510.7 259.3L509.2 260.3L509.2 257.0L510.2 254.9L509.4 253.5L511.7 251.3L503.1 251.8L502.3 250.1L500.1 252.0L500.1 253.5L497.0 252.0L494.2 256.4L491.9 256.4L489.5 259.1L487.9 256.2L484.9 256.8L485.3 255.4L483.7 254.0L478.8 253.8L477.0 254.6L472.8 253.0L466.2 253.7L461.5 259.0L464.0 261.7L461.0 263.1L460.5 265.0L461.1 261.5L459.9 258.9L461.2 254.0L457.3 247.8L460.3 244.0L460.6 236.0L467.7 235.1L468.5 233.2L472.4 232.1L475.2 234.1L478.8 235.2L485.9 234.9L486.9 234.0L491.9 235.0L493.4 233.6L495.3 234.7L499.7 234.2L503.7 231.8L507.1 233.5L508.3 232.4L517.9 230.9L520.2 229.2L524.7 230.0L527.6 231.9L541.2 230.7L544.7 228.4L544.2 226.6L548.1 222.5L553.9 218.1L552.8 216.1L553.3 215.5L558.1 217.3L572.3 211.4L576.0 210.2L577.5 210.8L582.3 208.2L586.8 207.8L590.0 208.1L585.8 213.2L586.0 215.8L588.0 217.5L588.1 220.0L589.4 220.0L591.0 221.5L587.6 223.8L582.9 224.7L580.8 223.9L579.7 224.4L579.0 226.9ZM569.8 255.3L570.3 258.7L565.4 264.3L562.6 265.4L560.0 264.1L559.9 262.4L558.6 261.8L555.1 264.1L552.7 264.3L552.3 263.1L550.6 263.7L548.4 262.6L543.3 263.3L544.2 265.0L539.9 271.1L537.7 270.0L536.8 265.7L534.1 263.2L542.0 256.1L541.5 255.0L543.5 253.5L543.6 256.2L545.8 256.3L547.8 254.7L547.3 253.5L548.3 248.8L554.3 240.0L554.9 242.3L556.8 240.6L556.8 238.8L560.7 236.2L563.9 235.7L567.8 231.1L569.1 231.9L573.6 229.2L574.4 237.2L570.8 242.5L572.1 244.8L571.9 250.3L573.4 251.1ZM510.2 272.7L506.3 271.3L500.6 271.7L498.4 272.7L497.5 271.6L495.2 272.5L492.8 271.6L492.3 270.3L473.0 272.3L466.0 270.5L463.1 268.8L460.1 268.8L459.7 266.7L461.0 263.1L464.0 261.7L461.5 259.0L466.2 253.7L472.8 253.0L477.0 254.6L478.8 253.8L483.7 254.0L485.3 255.4L484.9 256.8L487.9 256.2L489.5 259.1L491.9 256.4L494.2 256.4L497.0 252.0L500.1 253.5L500.1 252.0L502.3 250.1L503.1 251.8L511.7 251.3L509.4 253.5L510.2 254.9L509.2 257.0L509.2 260.3L510.7 259.3L515.3 259.2L518.3 263.0L519.9 262.5L518.0 265.0L519.8 265.7L522.7 269.1L519.1 273.0L516.1 273.1L515.0 274.9ZM562.1 273.0L563.9 273.0L562.8 278.1L561.7 279.3L561.6 281.2L559.8 282.2L559.6 283.8L556.2 287.1L556.9 288.4L551.6 301.9L548.7 300.1L546.0 299.5L544.2 300.1L542.8 298.3L540.3 298.2L537.8 299.7L535.1 296.4L532.5 297.9L526.6 295.8L529.4 279.9L530.8 280.2L535.2 270.1L537.1 268.6L537.7 270.0L539.9 271.1L544.2 265.0L543.3 263.3L548.4 262.6L550.6 263.7L552.3 263.1L552.7 264.3L557.1 263.4L558.6 261.8L559.9 262.4L560.0 264.1L562.6 265.4L560.5 270.7ZM500.8 315.2L501.8 319.8L500.0 321.5L497.0 322.0L494.4 315.3L491.5 314.9L490.2 310.1L488.2 307.7L488.1 304.3L489.6 304.5L489.4 301.2L491.1 298.5L498.1 296.1L498.7 294.2L500.8 295.2L500.6 292.8L504.3 295.9L504.3 290.7L508.9 288.7L509.3 286.4L510.5 286.8L511.5 290.2L510.6 291.9L512.9 294.2L512.8 300.4L511.4 302.7L511.6 304.6L508.5 304.5L507.2 306.6L504.6 305.5L505.4 310.3L504.7 311.5L501.8 309.5L502.5 312.8ZM427.2 226.9L424.5 227.4L422.1 224.9L422.2 219.0L425.4 211.9L425.6 208.5L434.4 204.9L437.3 205.4L439.1 207.3L440.1 210.4L439.2 214.7L437.5 217.1L438.1 220.4L440.7 222.6L440.4 223.7L438.3 224.3L438.0 225.7L436.5 226.8L432.5 225.9L430.3 228.3ZM345.4 238.4L343.6 237.2L342.1 234.4L337.9 232.5L337.6 228.9L335.8 228.0L336.4 226.0L333.7 222.3L333.0 220.2L336.9 220.0L338.6 218.2L342.0 219.7L342.5 221.0L349.7 222.2L351.2 224.9L350.2 228.9L356.9 230.9L358.4 232.8L361.8 232.9L361.5 234.6L363.2 235.6L364.8 235.9L371.4 233.3L373.2 234.3L373.1 237.4L375.9 239.8L379.5 237.6L390.4 240.5L394.2 243.1L396.9 242.7L401.0 240.0L402.0 242.5L403.5 243.5L409.9 242.7L414.0 244.1L418.1 242.8L422.0 243.8L423.4 242.4L423.6 240.3L426.3 242.3L427.2 245.2L418.1 251.9L417.5 254.2L420.6 259.0L422.3 259.6L421.7 262.0L422.7 264.6L420.4 263.3L416.3 266.1L416.7 267.8L418.0 268.8L417.4 270.1L415.1 269.7L414.8 268.5L411.8 268.0L410.2 271.1L407.0 270.5L406.3 273.6L403.2 275.2L403.4 278.3L402.4 279.0L401.2 284.1L399.1 284.0L398.2 285.9L394.9 284.3L392.7 285.4L391.9 284.5L389.5 286.8L388.5 289.9L384.8 287.8L385.3 285.6L381.7 284.6L381.6 282.5L380.1 280.6L377.5 281.6L373.4 279.7L371.9 283.3L372.1 284.9L369.2 286.3L365.7 285.8L359.6 290.0L358.4 288.9L355.7 289.9L355.4 287.8L353.6 286.3L349.0 289.3L347.1 291.8L343.8 289.4L343.5 287.2L338.6 286.2L336.8 283.8L335.3 285.9L333.9 286.3L326.1 286.4L326.6 283.7L325.8 281.6L323.8 280.6L322.9 275.4L322.8 271.4L326.8 268.5L333.3 265.7L335.1 262.9L338.3 261.1L338.2 259.7L339.6 258.9L340.9 260.3L343.3 260.4L343.7 258.6L348.4 260.1L349.3 258.9L349.8 257.5L337.8 250.5L338.4 247.8L340.3 246.7L339.9 244.2L334.6 242.6L334.8 241.3L337.7 240.2L339.0 238.4L342.5 239.1ZM413.6 292.6L413.5 294.7L409.4 298.8L407.2 297.6L405.0 297.8L406.3 300.7L405.0 301.7L402.8 301.8L401.9 302.7L398.3 301.1L395.9 302.2L395.9 305.3L389.2 306.7L386.7 309.1L385.5 311.8L382.7 310.5L382.3 308.5L380.7 308.3L379.8 309.9L376.5 310.2L376.5 313.2L375.1 315.1L376.5 317.4L379.9 317.9L384.0 321.3L389.5 321.4L388.1 322.9L387.8 326.0L392.5 329.1L392.9 330.9L395.1 330.9L395.9 332.9L395.1 334.3L396.8 334.9L396.6 336.6L394.3 339.1L389.4 336.1L388.0 336.8L386.5 336.1L384.9 333.8L381.4 332.9L380.3 331.3L378.1 333.0L379.5 336.4L378.4 338.4L379.6 339.7L376.8 344.4L374.4 344.4L373.6 342.4L372.2 342.7L366.2 340.3L362.8 343.8L361.6 342.4L358.4 341.3L359.4 338.2L359.2 332.9L352.5 334.5L347.9 334.3L345.7 336.0L342.8 336.2L336.8 332.0L338.4 329.5L341.6 328.3L345.0 322.9L344.5 321.7L340.4 321.7L337.9 318.6L337.4 317.3L338.2 313.6L335.8 312.4L336.9 306.8L335.3 308.3L331.8 307.3L330.4 302.5L327.5 301.5L324.6 296.5L322.4 296.1L324.1 292.3L323.9 287.1L333.9 286.3L335.3 285.9L336.8 283.8L338.6 286.2L343.5 287.2L343.8 289.4L347.1 291.8L349.0 289.3L353.6 286.3L355.4 287.8L355.7 289.9L358.4 288.9L359.6 290.0L365.7 285.8L369.2 286.3L372.1 284.9L371.9 283.3L373.4 279.7L377.5 281.6L380.1 280.6L381.6 282.5L381.7 284.6L385.3 285.6L384.8 287.8L388.5 289.9L389.5 286.8L391.9 284.5L392.7 285.4L394.9 284.3L398.2 285.9L399.1 284.0L401.2 284.1L402.4 279.0L403.4 278.3L403.2 275.2L406.3 273.6L407.0 270.5L410.2 271.1L411.8 268.0L414.8 268.5L415.1 269.7L416.6 269.9L416.9 273.6L420.3 276.2L420.4 277.4L417.9 281.4L419.3 281.9L418.9 285.5L416.9 285.1L417.1 289.7L415.0 292.5Z","South":"M254.6 548.7L251.8 548.4L249.6 550.7L246.4 547.9L246.3 545.6L244.4 544.9L244.1 546.1L240.9 546.2L238.2 548.0L236.7 546.6L240.0 545.1L239.3 541.8L242.1 541.7L242.0 543.1L245.2 542.7L245.6 544.4L247.6 544.0L247.9 542.2L250.8 541.1L252.7 537.1L253.6 538.3L256.5 538.2L257.6 536.5L259.2 543.1L258.3 546.2L256.2 546.6L255.9 548.9ZM254.6 548.7L255.9 548.9L256.2 546.6L258.3 546.2L257.2 552.7L255.8 552.4L256.1 550.4ZM218.3 557.9L218.8 563.2L221.6 563.4L223.3 566.4L223.0 568.2L219.6 568.3L213.3 564.6L209.8 560.8L205.4 568.2L199.1 567.7L198.4 566.7L201.4 563.9L201.1 557.0L204.0 556.9L206.0 552.5L208.1 551.9L209.1 553.2L210.8 552.8ZM223.0 568.2L228.6 556.5L230.1 556.9L232.1 555.2L235.2 554.3L240.3 556.9L240.8 552.3L246.5 555.3L246.4 556.4L244.1 557.8L246.1 558.9L246.5 560.1L244.3 564.4L240.3 563.7L238.8 562.1L235.6 562.1L235.2 568.5L236.6 569.4L234.4 571.6L229.9 570.7L228.9 571.9L224.8 572.6ZM204.2 568.7L208.0 564.8L209.8 560.8L213.3 564.6L219.6 568.3L223.0 568.2L224.1 569.8L224.6 574.6L222.2 576.3L219.8 576.4L214.5 572.7L209.0 574.6L204.5 574.3L202.7 573.3L204.8 569.7ZM202.7 573.3L209.0 574.6L214.5 572.7L219.8 576.4L222.2 576.3L222.9 575.2L224.2 575.6L223.2 577.5L224.4 579.3L226.6 579.4L226.2 581.1L227.4 583.1L226.5 586.3L225.2 584.8L223.4 585.4L223.6 586.6L221.7 587.9L218.8 583.9L218.9 581.4L210.6 583.2L206.4 585.6L203.1 584.5L204.2 579.9L203.0 580.2ZM202.7 573.3L203.0 580.2L204.2 579.9L202.9 585.6L206.7 589.5L207.3 593.3L208.2 593.5L206.8 594.5L201.9 592.3L199.0 592.6L198.5 588.2L191.5 587.4L190.7 585.8L187.6 586.3L188.6 582.6L185.2 581.5L184.9 580.1L186.5 576.9L188.7 576.5L190.7 577.4L194.6 576.6L197.5 577.6L198.9 573.5ZM185.2 581.5L188.6 582.6L187.6 586.3L180.5 590.6L176.6 590.5L178.3 588.9L178.8 586.9L172.2 583.0L178.0 579.0L179.1 580.9ZM203.1 584.5L206.4 585.6L210.6 583.2L218.9 581.4L218.8 583.9L219.9 585.8L216.4 593.7L214.4 592.9L213.0 594.8L211.7 593.3L207.3 593.3L206.7 589.5L202.9 585.6ZM223.6 586.6L223.4 585.4L225.2 584.8L226.5 586.3L227.4 583.1L229.9 585.8L234.3 586.4L232.2 592.6L230.9 593.7L230.2 592.9L225.7 592.3L224.4 593.5L223.9 591.3L225.1 587.7ZM187.6 586.3L190.7 585.8L191.8 591.2L193.9 592.5L191.9 593.6L192.3 596.9L193.7 597.3L194.6 599.4L191.2 600.8L191.7 604.6L190.3 604.6L190.3 610.7L188.0 612.9L186.1 612.0L184.5 608.2L186.4 600.0L184.8 598.0L181.1 596.5L183.2 593.3L181.9 591.3L182.2 589.6ZM219.9 585.8L221.7 587.9L223.6 586.6L225.1 587.7L223.9 591.3L224.4 593.5L225.7 592.3L230.2 592.9L230.8 597.5L228.0 599.0L228.1 601.8L225.6 603.0L222.9 601.8L219.3 603.9L220.3 609.3L219.1 608.5L216.6 611.0L215.5 609.4L216.2 607.9L214.0 607.1L213.8 605.4L214.7 605.3L215.8 601.9L217.6 601.3L218.5 602.6L221.6 601.0L221.8 598.5L218.1 596.0L214.0 596.0L213.0 594.8L214.4 592.9L216.4 593.7ZM230.9 593.7L232.2 592.6L234.3 586.4L238.5 585.5L241.3 587.9L241.5 589.8L239.2 593.6L235.4 596.3L230.8 597.5ZM191.5 587.4L198.5 588.2L199.0 592.6L201.9 592.3L204.9 593.7L204.5 597.3L206.0 598.8L205.5 602.4L203.7 603.5L200.8 602.3L199.3 604.1L197.2 603.6L196.2 605.3L196.5 607.6L195.2 607.7L195.4 611.7L194.5 612.6L192.2 609.7L190.3 610.7L190.3 604.6L191.7 604.6L191.2 600.8L194.6 599.4L193.7 597.3L192.3 596.9L191.9 593.6L193.9 592.5L191.8 591.2ZM228.1 601.8L228.0 599.0L229.6 597.6L235.4 596.3L241.9 590.9L242.4 592.1L240.4 598.2L237.1 596.5L236.4 597.6L238.1 604.4L240.9 606.7L240.7 610.7L238.9 610.6L236.4 612.7L235.5 614.8L233.2 612.5L234.0 609.9L233.5 606.5L234.7 605.5L233.5 603.5ZM204.9 593.7L206.8 594.5L209.4 592.9L211.7 593.3L214.0 596.0L218.1 596.0L221.8 598.5L221.6 601.0L218.5 602.6L217.6 601.3L215.8 601.9L214.7 605.3L213.8 605.4L212.5 605.1L213.4 599.4L210.2 600.1L209.0 603.3L205.5 602.4L206.0 598.8L204.5 597.3ZM241.3 609.5L240.9 606.7L238.1 604.4L236.4 597.6L237.1 596.5L240.4 598.2L241.7 595.7L244.5 594.8L246.6 595.2L246.1 596.7L244.9 597.0L244.5 599.1L244.9 604.8L245.9 605.9L244.1 609.3ZM213.8 605.4L214.0 607.1L216.2 607.9L215.5 609.4L216.6 611.0L215.5 614.9L214.1 616.2L212.6 613.7L210.0 613.7L207.7 616.2L203.5 617.3L203.0 614.6L201.5 613.0L198.2 614.9L194.1 615.1L195.4 611.7L195.2 607.7L196.5 607.6L196.2 605.3L197.2 603.6L199.3 604.1L200.8 602.3L203.7 603.5L205.5 602.4L209.0 603.3L210.2 600.1L213.4 599.4L212.5 605.1ZM216.6 611.0L217.6 610.8L219.1 615.2L219.2 618.6L216.3 619.6L216.1 621.0L212.8 620.7L212.5 623.3L210.9 623.8L209.9 626.1L204.8 626.1L203.4 624.6L202.0 624.6L199.8 625.6L198.5 627.9L203.5 617.3L207.7 616.2L210.0 613.7L212.6 613.7L214.1 616.2L215.5 614.9ZM194.1 615.1L198.2 614.9L201.5 613.0L203.0 614.6L202.8 619.4L198.5 627.9L196.9 628.7L195.1 627.0L194.1 627.7L191.9 626.7L193.6 622.5L192.9 620.7ZM198.5 627.9L199.8 625.6L202.0 624.6L203.4 624.6L204.8 626.1L209.9 626.1L210.9 623.8L214.0 622.9L218.6 626.8L216.6 630.2L214.2 630.5L214.7 632.0L217.3 632.7L215.5 634.2L213.1 632.3L210.8 632.4L207.8 636.1L195.6 631.1L196.9 628.7ZM206.3 635.1L207.8 636.1L210.8 632.4L213.1 632.3L215.5 634.2L215.4 637.3L217.7 638.6L213.2 643.4L213.7 646.4L212.5 648.0L212.6 652.1L210.8 654.9L208.9 655.9L206.4 655.4L205.4 643.5L203.0 640.7L202.7 637.6L206.6 636.2ZM200.9 660.3L200.2 661.8L197.8 661.4L193.3 659.4L190.5 656.9L193.9 652.0L194.1 650.3L197.7 652.0L200.3 657.2L199.8 659.9ZM234.4 571.6L236.6 569.4L235.2 568.5L235.6 562.1L238.8 562.1L240.3 563.7L247.7 564.6L249.2 566.7L252.1 566.4L248.9 571.1L248.5 573.3L247.4 573.2L245.3 575.0L245.2 576.9L240.7 575.2L239.9 577.2L237.2 575.1L235.4 575.1L235.8 573.3ZM234.3 586.4L229.9 585.8L228.2 583.5L229.1 582.1L234.6 581.7L236.0 580.2L240.1 579.9L240.7 575.2L245.2 576.9L245.3 575.0L247.8 575.8L246.8 580.8L247.9 585.4L243.4 590.1L240.4 591.6L241.5 589.8L241.3 587.9L238.5 585.5ZM220.3 609.3L219.3 603.9L221.2 602.2L222.9 601.8L225.6 603.0L228.1 601.8L233.5 603.5L234.7 605.5L233.5 606.5L234.0 609.9L233.2 612.5L235.5 614.8L236.1 616.7L233.3 621.3L231.6 621.0L229.3 617.4L229.9 615.7L228.9 614.6L224.6 613.9ZM217.6 610.8L219.1 608.5L224.6 613.9L228.9 614.6L229.9 615.7L229.3 617.4L230.6 619.2L227.2 621.7L225.0 620.8L224.9 624.5L226.5 624.8L227.0 627.0L223.5 628.2L218.6 626.8L214.0 622.9L212.5 623.3L212.8 620.7L216.1 621.0L216.3 619.6L219.2 618.6L219.1 615.2ZM215.5 634.2L217.3 632.7L214.7 632.0L214.2 630.5L216.6 630.2L218.6 626.8L223.5 628.2L227.0 627.0L226.5 624.8L224.9 624.5L225.0 620.8L227.2 621.7L230.6 619.2L231.6 621.0L233.2 621.3L230.3 625.0L228.6 630.0L232.1 633.9L237.8 633.6L237.2 635.1L235.1 634.4L228.0 634.8L221.8 637.7L217.7 638.6L215.4 637.3ZM240.7 610.7L241.3 609.5L244.1 609.3L245.9 605.9L244.9 604.8L244.5 599.1L244.9 597.0L246.1 596.7L248.7 599.0L248.9 611.4ZM241.7 595.7L242.4 592.1L241.9 590.9L247.9 585.4L248.8 595.4L244.5 594.8ZM224.8 572.6L228.9 571.9L229.9 570.7L234.4 571.6L235.8 573.3L235.4 575.1L237.2 575.1L239.9 577.2L240.1 579.9L236.0 580.2L234.6 581.7L231.1 581.5L228.2 583.5L227.4 583.1L226.2 581.1L226.6 579.4L224.4 579.3L223.2 577.5L224.2 575.6L222.9 575.2L224.6 574.6ZM195.6 631.1L199.2 633.1L206.3 635.1L206.6 636.2L202.7 637.6L203.3 639.3L201.6 639.4L200.4 645.5L198.3 645.0L197.8 646.6L194.8 645.5L191.5 640.5L194.1 637.2ZM203.3 639.3L203.0 640.7L205.4 643.5L206.4 655.4L208.9 655.9L206.2 657.8L205.2 659.8L200.9 660.3L199.8 659.9L200.3 657.2L197.7 652.0L194.1 650.3L192.0 646.6L193.8 643.5L196.0 646.4L197.8 646.6L198.3 645.0L200.4 645.5L201.6 639.4ZM235.2 554.3L232.1 555.2L230.1 556.9L228.6 556.5L227.7 557.8L225.8 556.7L226.7 553.9L226.1 552.9L222.3 553.0L222.7 549.6L228.5 547.2L234.2 549.0L235.7 547.0L235.6 545.8L237.0 548.2L234.7 551.4ZM236.7 546.6L238.2 548.0L240.9 546.2L244.1 546.1L244.4 544.9L246.3 545.6L246.4 547.9L247.5 548.9L242.8 551.7L242.6 552.8L240.8 552.3L240.3 556.9L235.2 554.3L234.7 551.4L237.0 548.2L235.6 545.8ZM223.3 566.4L221.6 563.4L218.8 563.2L218.3 557.9L223.5 552.5L226.1 552.9L226.7 553.9L225.8 556.7L227.7 557.8L224.6 565.3ZM247.5 548.9L249.6 550.7L251.8 548.4L253.6 549.1L254.1 550.4L253.0 553.1L250.0 554.0L250.3 557.8L248.2 558.1L246.5 560.1L246.1 558.9L244.1 557.8L246.4 556.4L246.5 555.3L242.6 552.8L242.8 551.7ZM253.6 549.1L254.6 548.7L256.1 550.4L255.8 552.4L257.2 552.7L257.3 554.5L254.8 562.9L252.1 566.4L249.2 566.7L247.7 564.6L244.3 564.4L248.2 558.1L250.3 557.8L250.0 554.0L253.0 553.1L254.1 550.4ZM331.4 414.5L331.2 412.7L332.2 412.1L334.3 416.7L337.4 416.7L338.5 418.0L343.3 417.3L347.5 411.4L349.1 412.0L351.4 409.8L352.8 410.8L339.3 428.2L329.5 433.5L328.2 431.7L329.7 428.9L327.8 427.2L327.7 423.9L328.8 422.6L327.0 421.7L326.8 419.1L332.6 416.2ZM322.8 412.1L325.2 413.5L325.6 411.7L329.5 410.2L332.6 416.2L326.8 419.1L327.0 421.7L328.8 422.6L327.7 423.9L327.8 427.2L329.7 428.9L328.2 431.7L329.5 433.5L326.4 436.5L324.0 436.2L324.0 434.0L321.2 435.2L319.8 438.8L318.5 438.9L317.8 436.7L315.6 435.8L315.6 434.5L318.6 431.0L321.1 430.4L315.8 424.9L316.3 422.9L315.4 421.4L318.3 417.4L319.7 418.2L324.1 415.4L322.1 413.1ZM303.3 424.2L304.5 423.3L307.1 427.0L306.7 429.0L310.6 427.5L311.3 425.0L314.8 427.1L317.1 426.1L321.1 430.4L318.6 431.0L315.6 434.5L315.6 435.8L317.8 436.7L318.5 438.9L319.8 438.8L321.2 435.2L324.0 434.0L324.0 436.2L326.4 436.5L318.8 445.4L305.7 452.5L305.2 451.3L307.0 447.8L304.1 446.4L302.9 447.5L301.6 446.9L299.1 443.6L294.3 443.6L294.6 441.0L293.4 439.5L294.9 436.7L293.8 435.0L294.9 433.7L297.8 434.1L299.5 435.5L301.0 434.3L301.9 425.7ZM194.0 425.7L195.1 425.2L197.1 426.4L197.1 428.4L198.2 429.4L201.1 429.0L200.3 430.9L201.1 433.2L200.0 434.1L202.3 436.0L198.0 442.3L197.8 444.8L190.2 443.1L181.7 442.6L182.4 440.3L184.0 439.3L182.9 437.6L186.7 437.3L187.4 431.0L189.8 431.9ZM176.3 443.5L177.6 443.1L178.3 440.8L181.7 442.6L194.2 443.8L202.9 446.5L199.3 448.3L196.4 453.1L196.1 454.5L198.2 455.6L198.9 457.2L198.1 460.8L195.6 459.7L191.9 459.5L188.8 460.5L188.0 462.5L186.5 462.3L185.7 460.5L182.8 463.2L177.0 464.7L175.1 463.4L176.6 458.2L175.0 454.4L172.4 454.7L168.8 450.5L172.1 449.8L175.3 451.0L174.3 444.5ZM269.5 487.7L270.9 482.7L268.8 474.5L266.0 470.2L261.9 467.5L254.9 466.1L254.6 462.7L253.4 462.3L252.7 460.0L253.1 458.8L255.9 457.7L258.4 458.7L258.7 460.9L264.0 463.5L264.6 460.0L259.6 458.7L264.2 454.7L269.8 457.2L270.5 456.6L272.2 457.0L271.7 459.1L273.0 462.1L272.5 466.7L273.5 468.2L277.1 466.8L278.1 467.5L280.1 465.2L280.3 468.9L282.4 471.9L284.8 473.1L281.4 473.1L278.7 474.6L276.2 481.6L273.4 484.4L273.3 486.6ZM139.6 464.6L144.3 463.3L144.0 461.3L147.9 459.2L149.9 461.7L151.7 461.3L152.0 459.2L153.3 459.2L153.2 461.1L154.4 461.3L153.7 464.7L151.7 465.2L151.5 467.1L149.4 467.7L148.5 469.4L146.2 468.7L145.8 470.6L147.0 473.7L148.8 473.4L152.1 476.4L151.9 478.2L155.4 478.3L155.6 483.6L153.8 485.4L152.5 485.3L152.2 488.0L150.3 489.0L148.4 488.5L147.3 489.4L145.4 488.9L142.5 489.9L142.1 492.3L140.5 494.3L137.9 495.6L136.4 494.3L133.9 494.7L133.8 493.2L130.8 491.9L129.9 488.5L127.4 488.9L127.5 487.4L129.3 486.0L129.9 486.8L132.2 486.4L134.7 480.1L132.8 479.2L135.1 477.6L135.1 475.3L132.8 475.0L131.6 473.5L131.6 468.3L136.4 466.5L137.0 468.3L138.2 467.7ZM177.0 464.7L182.8 463.2L185.7 460.5L186.5 462.3L188.0 462.5L188.8 460.5L191.9 459.5L195.6 459.7L198.1 460.8L198.5 463.0L197.4 464.5L198.0 466.9L197.2 469.2L193.5 470.2L194.5 471.6L192.1 471.1L188.6 468.5L174.3 476.5L173.6 474.5L174.3 471.2L176.4 471.0L175.8 466.3ZM155.6 483.6L155.4 478.3L151.9 478.2L152.1 476.4L148.8 473.4L147.0 473.7L145.8 470.6L146.2 468.7L148.5 469.4L149.4 467.7L151.5 467.1L151.7 465.2L153.7 464.7L154.2 463.5L156.5 467.4L156.6 470.6L159.5 470.9L166.3 474.2L167.8 474.1L168.7 476.2L172.3 477.2L174.2 482.5L168.5 483.1L168.9 480.7L166.7 481.3L166.7 484.1L164.6 482.5L164.0 484.7L161.7 485.2L157.3 484.8ZM253.4 462.3L254.6 462.7L254.9 466.1L261.9 467.5L266.0 470.2L268.8 474.5L270.9 482.7L269.5 487.7L268.3 484.0L266.3 483.6L260.8 485.6L260.0 481.6L258.3 479.7L254.8 479.9L251.9 477.8L249.3 477.4L249.1 480.3L250.0 481.7L249.1 482.9L247.5 482.3L246.5 484.2L242.6 482.2L242.7 474.7L240.2 474.1L238.6 475.7L235.3 475.8L235.4 469.3L236.2 468.1L239.6 467.7L247.5 464.4L250.8 466.5ZM281.5 439.5L286.1 439.4L290.2 436.3L293.8 435.0L294.9 436.7L293.4 439.5L294.6 441.0L294.3 443.6L299.1 443.6L301.6 446.9L302.9 447.5L304.1 446.4L307.0 447.8L305.2 451.3L305.7 452.5L301.0 456.8L299.5 459.6L299.5 461.0L301.9 461.4L300.9 468.2L288.2 473.9L287.9 471.3L291.0 469.7L291.3 467.3L288.8 462.9L288.9 459.0L286.6 453.4L287.5 451.2L283.7 447.2L279.7 447.0L278.5 445.3L273.2 443.4L271.2 443.6L270.7 441.4L272.5 440.4L274.3 441.4L276.6 438.6L278.7 439.6ZM174.3 476.5L188.6 468.5L192.1 471.1L201.0 473.1L199.0 475.9L199.3 482.6L193.7 481.8L189.9 483.0L184.9 489.5L184.4 491.3L181.8 490.2L181.8 489.0L179.9 487.8L176.7 487.0L176.3 483.5L173.2 483.2L174.2 482.5L172.3 477.2ZM235.3 475.8L238.6 475.7L240.2 474.1L242.7 474.7L242.6 482.2L246.5 484.2L247.5 482.3L249.1 482.9L250.0 481.7L249.1 480.3L249.3 477.4L251.9 477.8L254.8 479.9L258.3 479.7L260.0 481.6L260.8 485.6L257.2 489.3L256.3 493.0L254.5 495.4L252.8 503.6L247.2 504.2L246.3 503.1L236.3 501.3L233.5 501.5L232.6 502.7L232.5 499.8L225.6 499.6L225.9 488.6L226.7 484.0L228.8 481.7L228.7 476.8L231.4 475.5ZM189.0 484.7L189.9 483.0L193.7 481.8L202.4 483.6L207.2 483.1L211.3 484.6L213.3 484.2L215.0 480.8L218.3 479.3L221.2 480.0L222.3 479.0L225.2 480.8L229.1 478.2L228.8 481.7L226.7 484.0L225.9 488.6L225.2 496.4L225.7 502.5L222.9 505.5L220.6 504.8L219.8 502.9L215.7 502.2L214.2 504.8L212.4 504.7L210.3 503.8L209.9 501.2L204.2 500.8L202.5 498.4L200.7 499.5L199.5 498.7L196.1 498.7L194.6 500.8L191.7 500.6L191.6 497.1L189.7 496.3L189.1 493.9L189.1 489.3L190.2 488.7ZM164.0 484.7L164.6 482.5L166.7 484.1L166.7 481.3L168.9 480.7L168.5 483.1L170.3 483.4L172.0 482.3L176.3 483.5L176.7 487.0L179.9 487.8L181.8 489.0L181.8 490.2L184.4 491.3L184.0 492.4L180.2 493.7L177.7 495.9L174.9 496.3L173.4 498.7L171.2 498.4L169.2 500.1L166.2 500.6L163.5 493.5L165.1 489.9L164.5 487.5L166.5 488.3L168.0 486.3ZM152.2 488.0L152.5 485.3L153.8 485.4L155.6 483.6L157.3 484.8L164.0 484.7L168.0 486.3L166.5 488.3L164.5 487.5L165.1 489.9L163.5 493.5L166.2 500.6L165.3 502.0L160.9 505.0L159.4 504.0L157.5 504.3L157.3 503.0L154.9 501.9L154.8 497.0L156.5 495.4L156.3 491.7L157.3 491.5L157.8 489.0L156.4 487.8L154.8 488.8ZM184.4 491.3L184.9 489.5L189.0 484.7L190.2 488.7L189.1 489.3L189.1 493.9L189.7 496.3L191.6 497.1L191.7 500.6L190.8 503.0L188.1 503.5L184.2 501.6L179.2 513.2L176.2 513.3L174.3 510.9L171.1 510.6L170.0 506.8L166.5 505.9L163.0 506.8L163.5 508.8L162.1 509.9L160.1 508.2L160.9 505.0L166.2 500.6L169.2 500.1L171.2 498.4L173.4 498.7L174.9 496.3L177.7 495.9L180.2 493.7L184.0 492.4ZM140.5 494.3L142.1 492.3L142.5 489.9L145.4 488.9L150.3 489.0L152.2 488.0L154.8 488.8L156.4 487.8L157.8 489.0L157.3 491.5L156.3 491.7L156.5 495.4L154.8 497.0L154.9 501.9L151.1 501.3L143.7 502.6L142.5 500.7L144.2 498.1L143.7 496.8ZM130.8 491.9L133.8 493.2L133.9 494.7L136.4 494.3L137.9 495.6L140.5 494.3L143.7 496.8L144.2 498.1L142.5 500.7L143.7 502.6L147.8 502.1L148.2 505.3L146.9 508.9L147.1 511.8L146.6 513.6L144.0 514.5L143.5 516.7L145.0 517.5L143.7 521.7L139.0 522.4L139.0 526.9L137.2 528.3L135.8 526.9L130.5 510.2L127.4 508.7L126.7 506.1L129.2 505.5L131.2 502.9L130.7 501.6L131.6 499.5L130.4 498.3L131.8 497.3L131.7 495.4L130.7 494.9ZM154.9 501.9L157.3 503.0L157.5 504.3L159.4 504.0L160.9 505.0L160.1 508.2L163.0 511.1L163.4 513.7L162.9 515.0L160.7 515.4L158.8 519.9L157.3 520.3L153.4 517.8L153.4 515.5L151.7 513.4L147.1 511.8L146.9 508.9L148.2 505.3L147.8 502.1L151.1 501.3ZM191.7 500.6L194.6 500.8L196.1 498.7L199.5 498.7L200.7 499.5L202.5 498.4L204.2 500.8L209.9 501.2L210.3 503.8L212.4 504.7L210.6 508.9L212.4 510.8L209.6 515.5L215.6 520.0L218.5 520.1L218.2 522.0L219.3 524.2L219.4 526.7L215.3 529.0L214.6 531.7L212.2 531.5L212.0 529.7L209.6 529.5L208.9 527.5L205.6 528.5L205.0 530.7L198.1 533.9L197.6 530.2L192.2 529.6L189.8 532.5L187.9 530.6L189.3 528.1L187.0 525.9L187.3 524.4L186.2 522.8L188.5 522.1L189.0 525.2L192.1 526.3L195.2 525.9L196.8 528.7L197.6 527.8L196.7 526.2L196.7 522.6L199.1 523.0L198.5 520.0L197.1 518.9L196.1 520.3L191.3 518.9L190.9 521.5L188.5 521.2L185.9 518.7L186.3 515.4L184.3 514.5L183.4 512.9L185.6 504.5L183.3 504.0L184.2 501.6L188.1 503.5L190.8 503.0ZM257.6 536.5L256.5 538.2L253.6 538.3L251.0 533.3L249.1 533.3L248.8 531.9L243.6 530.0L243.3 528.4L240.2 529.1L239.9 521.1L235.9 513.1L234.7 512.5L232.5 504.2L233.5 501.5L246.3 503.1L247.2 504.2L252.8 503.6L254.2 509.8L256.1 513.2L254.7 521.4L255.0 525.7L257.3 531.6L256.8 534.4ZM178.0 513.8L180.9 510.4L182.3 506.9L181.6 506.4L183.3 504.0L185.6 504.5L183.4 512.9L184.3 514.5L186.3 515.4L185.9 518.7L188.5 522.1L186.2 522.8L187.3 524.4L185.2 527.5L182.5 528.9L183.4 531.6L180.4 532.7L176.9 536.0L174.7 536.2L173.0 535.6L171.9 533.2L170.4 532.7L169.3 529.7L167.9 528.9L168.5 521.5L170.1 520.8L170.1 516.9L171.6 516.2L173.2 517.4L176.9 516.3ZM162.1 509.9L163.5 508.8L163.0 506.8L166.5 505.9L170.0 506.8L171.1 510.6L174.3 510.9L176.2 513.3L178.0 513.8L176.9 516.3L173.2 517.4L171.6 516.2L170.1 516.9L170.5 519.3L170.1 520.8L168.5 521.5L168.2 527.9L166.2 530.8L164.6 530.9L163.1 524.4L158.9 525.3L154.5 523.9L157.3 520.3L158.8 519.9L160.7 515.4L162.9 515.0L163.4 513.7ZM139.0 526.9L139.0 522.4L143.7 521.7L145.0 517.5L143.5 516.7L144.0 514.5L146.6 513.6L147.1 511.8L151.7 513.4L153.4 515.5L153.4 517.8L157.3 520.3L154.5 523.9L158.9 525.3L163.1 524.4L164.6 530.9L157.5 532.8L156.8 532.0L153.3 536.2L151.2 535.6L150.9 537.4L149.0 538.9L146.8 536.5L146.0 532.6L144.5 531.9L144.1 529.9L139.9 528.8ZM137.2 528.3L139.0 526.9L139.9 528.8L144.1 529.9L144.5 531.9L146.0 532.6L146.8 536.5L149.0 538.9L148.4 541.3L150.4 545.4L149.2 546.6L145.9 545.8L143.5 547.2L142.4 546.5L141.3 547.8ZM167.9 528.9L169.3 529.7L170.4 532.7L171.9 533.2L173.0 535.6L174.9 537.0L170.7 538.7L170.2 540.2L168.1 540.9L167.5 543.5L165.1 543.4L161.5 545.2L162.0 546.3L161.1 548.6L159.7 548.7L157.9 551.4L157.0 547.9L154.7 546.9L154.1 545.3L152.3 546.6L149.6 544.8L148.4 541.3L149.0 538.9L150.9 537.4L151.2 535.6L153.3 536.2L156.8 532.0L157.5 532.8L166.2 530.8ZM198.1 533.9L205.0 530.7L205.6 528.5L208.9 527.5L209.6 529.5L212.0 529.7L212.1 533.1L213.9 535.1L214.0 538.7L212.4 540.2L212.4 542.9L208.7 542.9L208.2 543.8L205.5 543.5L205.9 542.0L197.6 538.8L196.1 535.1ZM157.9 551.4L159.7 548.7L161.1 548.6L162.0 546.3L161.5 545.2L165.1 543.4L167.5 543.5L168.1 540.9L170.2 540.2L170.7 538.7L174.9 537.0L176.1 539.2L174.6 541.2L175.1 544.9L177.2 546.0L177.2 547.3L180.0 547.8L180.6 549.3L179.5 552.9L177.6 553.8L175.4 553.0L173.7 558.3L171.9 557.5L169.6 560.2L165.9 557.9L166.5 556.0L164.7 553.2L163.7 556.9L160.2 556.3L159.2 552.7ZM214.5 554.9L210.8 552.8L209.1 553.2L208.1 551.9L206.7 552.4L206.7 550.0L207.5 547.1L208.6 546.5L208.7 542.9L212.4 542.9L212.4 540.2L214.0 538.7L213.7 536.5L218.1 536.0L217.3 541.2L222.0 543.1L219.5 550.0L217.1 550.8L216.2 552.6L214.9 552.6ZM206.7 552.4L204.6 550.7L204.9 548.7L203.8 545.0L199.9 544.1L198.6 544.3L195.4 550.0L195.0 547.0L192.3 544.2L193.2 542.3L195.6 541.9L196.6 537.7L197.6 538.8L205.9 542.0L205.5 543.5L208.2 543.8L208.6 546.5L207.5 547.1ZM141.3 547.8L142.4 546.5L143.5 547.2L145.9 545.8L149.2 546.6L150.4 545.4L152.3 546.6L154.1 545.3L154.7 546.9L157.0 547.9L156.9 549.9L159.2 552.7L160.2 559.7L156.5 561.2L152.0 560.3L148.4 556.7L143.3 554.9ZM206.0 552.5L204.0 556.9L201.1 557.0L201.4 556.1L199.9 554.6L195.7 552.0L195.4 550.0L198.6 544.3L203.8 545.0L204.9 548.7L204.6 550.7ZM143.3 554.9L148.4 556.7L152.0 560.3L154.9 560.9L153.8 562.8L155.0 563.7L154.8 565.6L150.4 567.9L149.6 570.2ZM160.2 556.3L163.7 556.9L164.7 553.2L166.5 556.0L165.9 557.9L167.6 559.2L165.3 563.2L167.3 564.7L167.8 566.6L170.1 567.9L171.3 571.1L169.6 572.7L167.3 573.8L164.5 573.3L162.9 571.0L158.3 568.7L156.3 565.6L154.8 565.6L155.0 563.7L153.8 562.8L154.9 560.9L156.5 561.2L160.2 559.7ZM204.2 568.7L204.8 569.7L202.7 573.3L198.9 573.5L197.5 577.6L194.6 576.6L190.7 577.4L188.7 576.5L186.5 576.9L184.9 580.1L185.2 581.5L179.1 580.9L178.0 579.0L176.1 579.7L176.0 577.8L179.4 574.0L179.5 572.5L186.6 571.7L188.7 568.5L191.0 567.8L190.7 566.8L191.9 565.2L197.5 565.7L199.1 567.7L201.3 567.5ZM176.1 579.7L175.2 581.3L170.2 584.8L165.3 581.4L165.1 579.7L163.0 577.9L163.2 575.5L165.7 575.9L165.9 573.6L169.6 572.7L169.6 575.5L171.6 575.4L174.0 577.9L176.0 577.8ZM162.9 577.0L165.1 579.7L165.3 581.4L170.2 584.8L169.0 588.1L164.2 590.4L163.5 592.3L160.8 585.4L159.1 584.2L157.8 579.1ZM168.7 601.2L166.3 602.2L163.5 592.3L164.2 590.4L169.0 588.1L170.2 584.8L172.2 583.0L178.8 586.9L175.9 592.4L174.5 593.0L174.0 596.5L168.2 598.5ZM176.6 590.5L180.5 590.6L182.2 589.6L181.9 591.3L183.2 593.3L181.1 596.5L184.8 598.0L186.4 600.0L185.0 604.3L184.9 610.1L180.4 609.2L177.3 606.8L176.0 604.3L177.1 600.7L171.6 600.4L170.7 601.8L168.7 601.2L168.2 598.5L174.0 596.5L174.5 593.0L175.9 592.4ZM184.9 610.1L186.1 612.0L176.0 611.2L173.8 613.8L170.7 613.7L166.3 602.2L168.7 601.2L170.7 601.8L171.6 600.4L177.1 600.7L176.0 604.3L177.3 606.8L180.4 609.2ZM189.5 611.9L192.2 609.7L194.5 612.6L192.9 620.7L193.6 622.5L191.9 626.7L194.1 627.7L195.1 627.0L196.9 628.7L194.4 633.9L192.8 630.5L187.8 630.1L186.0 628.4L186.0 624.4L180.5 621.2L186.9 613.8L190.0 613.8ZM186.3 612.1L188.0 612.9L189.5 611.9L190.0 613.8L186.9 613.8L179.5 622.0L175.1 621.6L173.5 620.4L173.2 622.7L170.8 613.8L173.8 613.8L176.0 611.2ZM173.2 622.7L173.5 620.4L175.1 621.6L175.5 628.2L177.1 627.9L178.2 632.7L180.8 633.0L181.8 636.7L180.6 637.5L177.4 637.9L174.0 629.7ZM180.5 621.2L186.0 624.4L186.0 628.4L187.8 630.1L186.8 631.2L180.8 629.8L178.1 631.6L177.1 627.9L175.5 628.2L175.1 621.6L179.5 622.0ZM194.4 633.9L194.1 637.2L192.7 639.1L186.9 636.7L185.0 638.6L181.9 638.7L180.6 637.5L181.8 636.7L180.8 633.0L178.2 632.7L178.1 631.6L180.8 629.8L186.8 631.2L187.8 630.1L191.7 630.0L192.8 630.5ZM192.7 639.1L191.5 640.5L193.8 643.5L192.6 645.9L189.5 644.2L188.9 645.8L187.0 646.2L184.5 643.9L183.3 645.5L181.7 645.8L178.8 642.9L177.4 637.9L180.6 637.5L181.9 638.7L185.0 638.6L186.9 636.7ZM190.5 656.9L182.0 646.7L181.7 645.8L183.3 645.5L184.5 643.9L187.0 646.2L188.9 645.8L189.5 644.2L192.6 645.9L192.0 646.6L194.1 650.3L193.9 652.0ZM284.8 473.1L282.4 471.9L280.3 468.9L280.1 465.2L278.1 467.5L277.1 466.8L273.5 468.2L272.5 466.7L273.0 462.1L271.7 459.1L272.2 457.0L270.5 456.6L270.5 454.5L271.8 453.2L273.0 453.8L276.8 452.9L277.1 450.7L278.9 450.9L279.6 448.8L273.5 446.3L271.2 443.6L273.2 443.4L278.5 445.3L279.7 447.0L283.7 447.2L287.5 451.2L286.6 453.4L288.9 459.0L288.8 462.9L291.3 467.3L291.0 469.7L287.9 471.3L288.2 473.9ZM227.4 395.1L227.8 396.8L229.8 397.5L229.2 402.0L227.6 402.3L227.7 406.2L223.9 405.3L217.8 408.5L213.7 405.3L214.0 402.7L216.1 401.5L215.4 396.7L216.8 396.2L217.3 394.7L216.1 391.7L218.7 393.9ZM220.7 449.6L219.1 450.8L217.9 449.3L218.2 448.1L220.5 448.1ZM237.1 416.7L235.6 418.2L234.2 418.1L233.1 420.3L230.5 420.1L229.8 422.2L227.6 421.9L223.8 421.1L221.9 418.8L220.9 413.7L224.3 411.8L225.7 412.3L227.4 410.9L231.2 410.8L234.6 415.5ZM241.6 440.0L241.1 443.0L242.9 444.0L241.6 444.8L242.1 446.2L236.2 447.4L233.5 442.8L231.3 441.9L230.6 440.0L227.9 439.3L229.1 438.4L232.7 438.2L234.0 435.5L235.6 435.5L241.4 438.4ZM259.8 421.4L261.5 420.9L265.4 423.4L266.8 425.3L267.9 431.5L261.1 434.0L260.2 433.5L257.6 436.6L255.1 433.4L250.6 433.4L248.5 432.3L250.3 426.1L256.5 426.4L256.8 422.6ZM199.3 482.6L199.0 475.9L201.0 473.1L202.4 472.7L208.1 476.6L208.4 478.3L212.9 482.1L214.8 481.6L213.3 484.2L211.3 484.6L207.2 483.1L205.4 483.8ZM200.1 428.7L199.5 426.8L201.0 422.8L204.0 422.6L205.1 424.0L208.3 424.1L210.0 425.4L210.4 423.5L215.1 425.0L219.0 424.3L220.8 424.7L220.6 429.0L221.6 431.1L217.7 432.5L216.2 431.2L211.1 434.0L209.1 432.7L207.4 433.2L206.8 431.0L204.7 431.0L203.7 429.4ZM229.8 422.2L230.5 420.1L233.1 420.3L235.6 422.4L235.7 425.1L237.2 424.4L241.0 426.5L242.7 425.6L242.7 429.3L240.8 430.1L240.5 431.6L237.1 432.4L235.9 431.1L234.2 431.7L233.0 430.3L233.6 429.1L231.1 428.1L231.3 427.0L229.1 424.5ZM270.5 456.6L269.8 457.2L264.2 454.7L259.6 458.7L264.6 460.0L264.0 463.5L258.7 460.9L258.4 458.7L255.9 457.7L253.1 458.8L252.8 457.0L249.4 454.2L249.8 452.3L247.7 450.8L251.7 449.1L253.0 450.2L254.2 450.0L256.7 446.0L256.2 445.4L259.7 445.2L262.5 449.8L264.5 449.4L266.4 451.4L270.1 449.8L270.5 452.1L271.8 453.2L270.5 454.5ZM247.8 399.3L251.4 403.4L250.4 408.9L245.1 408.0L243.7 406.2L241.7 406.0L242.3 407.9L240.5 408.6L238.2 407.7L236.4 411.8L233.9 410.4L231.9 407.7L227.7 406.2L227.6 402.3L229.2 402.0L229.4 400.0L232.8 400.4L234.8 402.0L236.1 399.0L238.5 399.4L240.7 401.1L246.5 398.6ZM242.1 446.2L241.6 444.8L244.6 443.8L245.4 442.5L249.7 441.5L250.0 439.6L251.8 437.9L251.7 432.9L255.1 433.4L257.6 436.6L256.2 445.4L256.7 446.0L254.2 450.0L253.0 450.2L251.7 449.1L247.7 450.8L246.6 448.1ZM204.3 471.8L205.1 467.9L203.8 461.6L204.5 458.1L207.8 459.9L210.1 457.9L215.4 459.9L217.0 463.2L217.0 465.4L214.5 464.0L212.1 467.4L209.7 467.1L209.2 469.5L207.8 470.9ZM250.4 408.9L248.9 410.2L250.6 411.7L251.1 415.6L249.1 416.3L247.9 419.3L245.1 420.0L239.8 416.0L237.1 416.7L234.6 415.5L232.5 412.0L231.2 410.8L229.6 410.8L228.2 408.9L229.7 406.8L236.4 411.8L238.2 407.7L240.5 408.6L242.3 407.9L241.7 406.0L243.7 406.2L245.1 408.0ZM207.4 433.2L209.1 432.7L211.1 434.0L216.2 431.2L217.7 432.5L221.6 431.1L220.1 442.5L216.3 441.2L216.3 442.6L211.7 439.5L211.2 437.5L208.2 437.9L205.9 435.6ZM220.7 449.6L220.5 448.1L218.2 448.1L216.9 446.2L217.6 443.5L219.1 442.1L223.7 442.2L224.9 442.9L223.9 445.0L225.5 447.4L225.2 449.3ZM235.2 472.8L235.3 475.8L231.4 475.5L228.7 476.8L229.1 478.2L225.2 480.8L222.3 479.0L221.2 480.0L218.3 479.3L215.0 480.8L213.9 479.7L214.4 477.0L213.7 476.2L216.2 472.3L211.7 469.9L212.1 467.4L214.5 464.0L217.0 465.4L217.0 463.2L220.9 462.2L225.5 463.9L224.3 466.8L224.9 468.7L226.5 469.3L227.7 468.4ZM245.1 465.1L239.6 467.7L236.2 468.1L235.2 472.8L227.7 468.4L226.5 469.3L224.9 468.7L224.3 466.8L226.5 462.1L225.0 461.3L227.4 456.5L228.4 457.1L229.7 455.1L230.4 451.8L232.9 452.6L235.1 452.2L236.1 450.0L240.3 450.3L243.3 456.2ZM208.5 416.1L206.2 414.6L205.3 410.8L206.6 408.4L206.7 405.6L208.4 404.7L213.3 407.1L213.7 405.3L217.8 408.5L223.9 405.3L229.7 406.8L228.2 408.9L229.6 410.8L227.4 410.9L225.7 412.3L224.3 411.8L220.9 413.7L218.3 412.4L216.2 413.6L214.7 412.1L210.7 413.9L210.4 415.3ZM204.0 422.6L203.9 420.1L205.1 419.8L206.2 417.0L210.4 415.3L210.7 413.9L214.7 412.1L216.2 413.6L218.3 412.4L220.9 413.7L221.0 416.7L223.8 421.1L220.8 424.7L217.7 425.0L210.4 423.5L210.0 425.4L208.3 424.1L205.1 424.0ZM233.1 420.3L234.2 418.1L235.6 418.2L237.1 416.7L239.8 416.0L245.1 420.0L247.9 419.3L248.3 420.4L247.5 422.3L246.1 422.2L244.9 424.4L241.0 426.5L237.2 424.4L235.7 425.1L235.6 422.4ZM223.8 421.1L229.8 422.2L229.1 424.5L231.3 427.0L230.5 429.4L225.4 430.3L223.2 429.5L221.6 430.3L220.6 429.0L220.8 424.7ZM218.2 448.1L217.9 449.3L219.1 450.8L220.7 449.6L225.2 449.3L225.5 452.2L226.9 452.8L226.0 454.7L227.4 456.5L225.0 461.3L226.5 462.1L225.5 463.9L220.9 462.2L217.0 463.2L215.4 459.9L210.1 457.9L209.5 458.6L207.9 456.7L211.0 452.9L209.7 450.7L211.1 446.8L215.7 448.3L216.9 446.2ZM197.8 444.8L198.0 442.3L202.3 436.0L200.0 434.1L201.1 433.2L200.3 430.9L201.1 429.0L203.7 429.4L204.7 431.0L206.8 431.0L207.4 433.2L205.9 435.6L208.2 437.9L211.2 437.5L211.7 439.5L216.3 442.6L216.3 441.2L219.1 442.1L215.7 448.3L211.1 446.8L210.9 447.9L209.2 447.1L208.1 445.2L205.8 443.7L202.9 446.5ZM235.6 435.5L234.0 435.5L232.7 438.2L229.1 438.4L226.8 442.1L224.9 442.9L220.1 442.5L221.6 430.3L223.2 429.5L225.4 430.3L230.5 429.4L231.1 428.1L233.6 429.1L233.0 430.3L234.2 431.7L235.9 431.1L237.1 432.4ZM253.1 458.8L253.1 464.0L250.8 466.5L247.5 464.4L245.1 465.1L243.3 456.2L238.8 446.5L242.1 446.2L246.6 448.1L247.7 450.8L249.8 452.3L249.4 454.2L252.8 457.0ZM198.3 459.2L198.9 457.2L198.2 455.6L196.1 454.5L196.4 453.1L199.3 448.3L205.8 443.7L210.9 447.9L209.7 450.7L211.0 452.9L207.9 456.7L209.5 458.6L207.8 459.9L203.5 457.2L201.8 457.9L201.0 460.1ZM215.0 480.8L214.8 481.6L212.9 482.1L208.4 478.3L208.1 476.6L202.4 472.7L207.8 470.9L209.2 469.5L209.7 467.1L212.1 467.4L211.7 469.9L216.2 472.3L213.7 476.2L214.4 477.0L213.9 479.7ZM242.9 444.0L241.1 443.0L241.6 440.0L243.9 439.2L245.0 434.7L243.5 433.1L243.0 429.8L245.5 429.0L246.1 430.9L250.6 433.4L251.7 432.9L251.8 437.9L250.0 439.6L249.7 441.5ZM241.6 440.0L241.4 438.4L235.6 435.5L237.1 432.4L240.5 431.6L240.8 430.1L242.7 429.3L243.7 431.0L243.5 433.1L245.0 434.7L243.9 439.2ZM227.9 439.3L230.6 440.0L231.3 441.9L233.5 442.8L236.2 447.4L238.6 447.1L240.3 450.3L236.1 450.0L235.1 452.2L232.9 452.6L230.4 451.8L228.4 457.1L226.0 454.7L226.9 452.8L225.5 452.2L225.5 447.4L223.9 445.0L224.9 442.9L226.8 442.1ZM225.6 499.6L232.5 499.8L232.5 504.2L234.7 512.5L235.9 513.1L239.9 521.1L240.4 531.2L237.2 532.9L235.5 532.1L233.9 529.7L228.3 528.9L223.2 529.6L222.2 527.5L219.4 526.7L218.5 520.1L215.6 520.0L209.6 515.5L212.4 510.8L210.6 508.9L212.4 504.7L214.2 504.8L215.7 502.2L219.8 502.9L220.6 504.8L222.9 505.5L225.7 502.5ZM187.3 524.4L187.0 525.9L189.3 528.1L187.9 530.6L189.8 532.5L192.2 529.6L197.6 530.2L198.3 532.1L198.1 533.9L196.1 535.1L196.6 537.7L195.6 541.9L193.2 542.3L192.3 544.2L192.7 545.2L190.0 546.6L189.9 548.5L191.7 550.4L191.6 554.3L190.0 555.2L186.8 553.9L185.3 549.9L183.5 548.7L180.6 549.3L180.0 547.8L177.2 547.3L177.2 546.0L175.1 544.9L174.6 541.2L176.1 539.2L174.7 536.2L176.9 536.0L180.4 532.7L183.4 531.6L182.5 528.9L185.2 527.5ZM188.5 521.2L190.9 521.5L191.3 518.9L196.1 520.3L197.1 518.9L198.5 520.0L199.1 523.0L196.7 522.6L196.7 526.2L197.6 527.8L196.8 528.7L195.2 525.9L192.1 526.3L189.0 525.2ZM236.7 546.6L235.6 545.8L235.7 547.0L234.2 549.0L228.5 547.2L225.5 548.2L222.7 549.6L222.1 554.4L218.3 557.9L216.0 557.1L214.5 554.9L214.9 552.6L216.2 552.6L217.1 550.8L219.5 550.0L222.0 543.1L217.3 541.2L218.1 536.0L213.7 536.5L212.2 531.5L214.6 531.7L215.3 529.0L219.4 526.7L222.2 527.5L223.2 529.6L228.3 528.9L233.9 529.7L235.5 532.1L237.2 532.9L240.4 531.2L240.2 529.1L243.3 528.4L243.6 530.0L248.8 531.9L249.1 533.3L251.0 533.3L252.7 537.1L250.8 541.1L247.9 542.2L247.6 544.0L245.6 544.4L245.2 542.7L242.0 543.1L242.1 541.7L239.3 541.8L240.0 545.1ZM201.1 557.0L201.4 563.9L198.4 566.7L195.3 565.2L194.1 561.6L191.1 561.2L190.0 555.2L191.6 554.3L191.7 550.4L189.9 548.5L190.0 546.6L191.8 545.1L195.0 547.0L195.7 552.0L199.9 554.6L201.4 556.1ZM174.5 557.3L175.4 553.0L177.6 553.8L179.5 552.9L180.6 549.3L183.5 548.7L185.3 549.9L186.8 553.9L190.0 555.2L191.1 561.2L194.1 561.6L195.3 565.2L191.9 565.2L190.7 566.8L188.3 566.7L187.1 563.1L185.3 562.9L182.7 564.4L176.8 561.2ZM167.6 559.2L169.6 560.2L171.9 557.5L173.7 558.3L174.5 557.3L176.8 561.2L182.7 564.4L185.3 562.9L187.1 563.1L188.3 566.7L190.7 566.8L191.0 567.8L188.7 568.5L186.6 571.7L179.5 572.5L179.4 574.0L176.0 577.8L174.0 577.9L171.6 575.4L169.6 575.5L169.6 572.7L171.3 571.1L170.1 567.9L167.8 566.6L167.3 564.7L165.3 563.2ZM149.6 570.2L150.4 567.9L156.3 565.6L158.3 568.7L165.9 573.6L165.7 575.9L163.2 575.5L162.9 577.0L157.3 579.0L151.4 572.1L150.3 572.1ZM245.3 575.0L247.4 573.2L248.5 573.3L247.8 575.8ZM274.9 440.0L274.3 441.4L272.5 440.4L270.7 441.4L271.2 443.6L273.5 446.3L279.6 448.8L278.9 450.9L277.1 450.7L276.8 452.9L273.0 453.8L270.5 452.1L270.1 449.8L266.4 451.4L264.5 449.4L262.5 449.8L261.4 448.8L259.7 445.2L256.2 445.4L257.6 436.6L260.2 433.5L261.1 434.0L270.0 430.6L270.2 432.1L272.6 431.5L273.4 440.0ZM158.9 451.2L158.2 449.4L159.5 447.3L163.3 449.6L165.0 449.0L165.8 450.8L168.8 450.5L172.4 454.7L175.0 454.4L176.6 458.2L175.1 463.4L177.0 464.7L175.8 466.3L176.4 471.0L174.3 471.2L173.6 474.5L174.3 476.5L172.3 477.2L168.7 476.2L167.8 474.1L166.3 474.2L159.5 470.9L156.6 470.6L156.5 467.4L154.2 463.5L154.4 461.3L153.2 461.1L153.3 459.2L160.1 458.6L160.3 455.5L159.3 454.2L159.9 451.9ZM246.1 596.7L246.6 595.2L248.8 595.4L248.7 599.0ZM250.1 416.5L254.4 419.8L257.8 418.9L259.8 421.4L256.8 422.6L256.5 426.4L250.3 426.1L248.5 432.3L246.1 430.9L245.5 429.0L243.0 429.8L242.7 425.6L244.9 424.4L246.1 422.2L247.5 422.3L249.1 416.3ZM198.1 460.8L198.3 459.2L201.0 460.1L201.8 457.9L203.5 457.2L204.5 458.1L203.8 461.6L205.1 467.9L204.3 471.8L201.0 473.1L194.5 471.6L193.5 470.2L197.2 469.2ZM524.3 566.9L524.0 568.6L522.2 568.1L521.6 566.6L523.1 565.1L524.4 565.3ZM524.4 564.7L521.2 565.4L520.7 564.8L521.4 553.2L522.8 552.6L523.8 540.5L524.4 538.7L526.9 536.3L528.0 537.4L528.5 543.2L527.8 547.6L526.2 549.1L525.2 548.2L524.1 551.7L525.5 552.1L526.6 559.9ZM517.0 597.4L518.6 602.3L517.2 606.0L515.1 605.3L513.9 600.0ZM519.5 583.4L520.6 586.5L518.6 587.2ZM521.9 568.2L521.5 570.3L522.5 571.2L520.9 572.8L522.7 574.9L521.9 579.1L519.9 580.2L521.8 581.2L520.0 583.5L517.6 578.6L517.3 574.3L518.7 575.7L519.9 567.4ZM544.8 680.9L546.5 687.5L545.7 691.1L544.1 691.7L540.8 683.5L541.4 682.0ZM541.9 676.2L542.9 677.8L540.6 680.8L540.0 679.2ZM522.7 635.3L523.4 637.5L521.4 637.8L521.2 635.7ZM252.1 566.4L248.9 571.1L248.5 573.3L247.4 573.2L245.3 575.0L247.8 575.8L246.8 580.8L248.9 591.0L248.8 595.4L246.6 595.2L246.1 596.7L248.7 599.0L248.9 611.4L238.9 610.6L236.4 612.7L235.5 614.8L236.1 616.7L230.3 625.0L228.6 630.0L232.1 633.9L237.8 633.6L237.2 635.1L230.4 634.5L221.8 637.7L218.5 638.0L214.4 641.3L213.2 643.4L213.7 646.4L212.5 648.0L212.6 652.1L210.8 654.9L206.2 657.8L205.2 659.8L200.9 660.3L200.2 661.8L193.3 659.4L190.5 656.9L193.9 652.0L194.1 650.3L192.0 646.6L193.8 643.5L191.5 640.5L194.1 637.2L194.4 633.9L196.9 628.7L195.1 627.0L194.1 627.7L191.9 626.7L193.6 622.5L192.9 620.7L194.5 612.6L192.2 609.7L188.0 612.9L186.1 612.0L184.5 608.2L186.4 600.0L184.8 598.0L181.1 596.5L183.2 593.3L181.9 591.3L182.2 589.6L180.5 590.6L176.6 590.5L178.3 588.9L178.8 586.9L172.2 583.0L178.0 579.0L179.1 580.9L185.2 581.5L184.9 580.1L186.5 576.9L188.7 576.5L190.7 577.4L194.6 576.6L197.5 577.6L198.9 573.5L202.7 573.3L204.8 569.7L204.2 568.7L201.3 567.5L199.1 567.7L198.4 566.7L201.4 563.9L201.1 557.0L204.0 556.9L206.0 552.5L208.1 551.9L209.1 553.2L210.8 552.8L218.3 557.9L222.1 554.4L222.7 549.6L228.5 547.2L234.2 549.0L235.7 547.0L235.6 545.8L236.7 546.6L240.0 545.1L239.3 541.8L242.1 541.7L242.0 543.1L245.2 542.7L245.6 544.4L247.6 544.0L247.9 542.2L250.8 541.1L252.7 537.1L253.6 538.3L256.5 538.2L257.6 536.5L259.2 543.1L254.8 562.9ZM257.6 536.5L256.5 538.2L253.6 538.3L252.7 537.1L250.8 541.1L247.9 542.2L247.6 544.0L245.6 544.4L245.2 542.7L242.0 543.1L242.1 541.7L239.3 541.8L240.0 545.1L236.7 546.6L235.6 545.8L235.7 547.0L234.2 549.0L228.5 547.2L225.5 548.2L222.7 549.6L222.1 554.4L218.3 557.9L216.0 557.1L214.5 554.9L214.9 552.6L216.2 552.6L217.1 550.8L219.5 550.0L222.0 543.1L217.3 541.2L218.1 536.0L213.7 536.5L213.9 535.1L212.1 533.1L212.0 529.7L209.6 529.5L208.9 527.5L205.6 528.5L205.0 530.7L198.1 533.9L197.6 530.2L192.2 529.6L189.8 532.5L187.9 530.6L189.3 528.1L187.0 525.9L187.3 524.4L186.2 522.8L188.5 522.1L189.0 525.2L192.1 526.3L195.2 525.9L196.8 528.7L197.6 527.8L196.7 526.2L196.7 522.6L199.1 523.0L198.5 520.0L197.1 518.9L196.1 520.3L191.3 518.9L190.9 521.5L188.5 521.2L185.9 518.7L186.3 515.4L184.3 514.5L183.4 512.9L185.6 504.5L183.3 504.0L184.2 501.6L188.1 503.5L190.8 503.0L191.7 500.6L191.6 497.1L189.7 496.3L189.1 493.9L189.1 489.3L190.2 488.7L189.0 484.7L189.9 483.0L193.7 481.8L202.4 483.6L207.2 483.1L211.3 484.6L213.3 484.2L215.0 480.8L218.3 479.3L221.2 480.0L222.3 479.0L225.2 480.8L229.1 478.2L228.7 476.8L231.4 475.5L235.3 475.8L235.4 469.3L236.2 468.1L239.6 467.7L247.5 464.4L250.8 466.5L253.1 464.0L253.1 458.8L255.9 457.7L258.4 458.7L258.7 460.9L264.0 463.5L264.6 460.0L259.6 458.7L264.2 454.7L269.8 457.2L271.8 453.2L273.0 453.8L276.8 452.9L277.1 450.7L278.9 450.9L279.6 448.8L273.5 446.3L271.2 443.6L270.7 441.4L272.5 440.4L274.3 441.4L276.6 438.6L278.7 439.6L286.1 439.4L294.9 433.7L297.8 434.1L299.5 435.5L301.0 434.3L301.9 425.7L304.5 423.3L307.1 427.0L306.7 429.0L310.6 427.5L311.3 425.0L314.8 427.1L317.1 426.1L315.8 424.9L316.3 422.9L315.4 421.4L318.3 417.4L319.7 418.2L324.1 415.4L322.1 413.1L322.8 412.1L325.2 413.5L325.6 411.7L329.5 410.2L331.4 414.5L331.2 412.7L332.2 412.1L334.3 416.7L337.4 416.7L338.5 418.0L343.3 417.3L347.5 411.4L349.1 412.0L351.4 409.8L352.8 410.8L339.3 428.2L329.5 433.5L318.8 445.4L305.7 452.5L301.0 456.8L299.5 459.6L299.5 461.0L301.9 461.4L300.9 468.2L288.2 473.9L281.4 473.1L278.7 474.6L276.2 481.6L273.4 484.4L273.3 486.6L271.8 487.5L269.5 487.7L268.3 484.0L266.3 483.6L263.3 484.3L258.6 487.3L254.5 495.4L252.8 503.6L254.2 509.8L256.1 513.2L254.7 521.4L255.0 525.7L257.3 531.6L256.8 534.4ZM194.5 471.6L201.0 473.1L199.0 475.9L199.3 482.6L193.7 481.8L189.9 483.0L189.0 484.7L190.2 488.7L189.1 489.3L189.1 493.9L189.7 496.3L191.6 497.1L191.7 500.6L190.8 503.0L188.1 503.5L184.2 501.6L183.3 504.0L185.6 504.5L183.4 512.9L184.3 514.5L186.3 515.4L185.9 518.7L188.5 521.2L190.9 521.5L191.3 518.9L196.1 520.3L197.1 518.9L198.5 520.0L199.1 523.0L196.7 522.6L196.7 526.2L197.6 527.8L196.8 528.7L195.2 525.9L192.1 526.3L189.0 525.2L188.5 522.1L186.2 522.8L187.3 524.4L187.0 525.9L189.3 528.1L187.9 530.6L189.8 532.5L192.2 529.6L197.6 530.2L198.1 533.9L205.0 530.7L205.6 528.5L208.9 527.5L209.6 529.5L212.0 529.7L212.1 533.1L213.9 535.1L213.7 536.5L218.1 536.0L217.3 541.2L222.0 543.1L219.5 550.0L217.1 550.8L216.2 552.6L214.9 552.6L214.5 554.9L210.8 552.8L209.1 553.2L208.1 551.9L206.0 552.5L204.0 556.9L201.1 557.0L201.4 563.9L198.4 566.7L199.1 567.7L201.3 567.5L204.2 568.7L204.8 569.7L202.7 573.3L198.9 573.5L197.5 577.6L194.6 576.6L190.7 577.4L188.7 576.5L186.5 576.9L184.9 580.1L185.2 581.5L179.1 580.9L178.0 579.0L176.1 579.7L176.0 577.8L174.0 577.9L171.6 575.4L169.6 575.5L169.6 572.7L167.3 573.8L164.5 573.3L162.9 571.0L158.3 568.7L156.3 565.6L154.8 565.6L155.0 563.7L153.8 562.8L154.9 560.9L152.0 560.3L148.4 556.7L143.3 554.9L137.2 528.3L135.8 526.9L130.5 510.2L127.4 508.7L126.7 506.1L129.2 505.5L131.2 502.9L130.7 501.6L131.6 499.5L130.4 498.3L131.8 497.3L129.9 488.5L127.4 488.9L127.5 487.4L129.3 486.0L129.9 486.8L132.2 486.4L134.7 480.1L132.8 479.2L135.1 477.6L135.1 475.3L132.8 475.0L131.6 473.5L131.6 468.3L136.4 466.5L137.0 468.3L138.2 467.7L139.6 464.6L144.3 463.3L144.0 461.3L147.9 459.2L149.9 461.7L151.7 461.3L152.0 459.2L160.1 458.6L160.3 455.5L159.3 454.2L159.9 451.9L158.2 449.4L159.5 447.3L163.3 449.6L165.0 449.0L165.8 450.8L172.1 449.8L175.3 451.0L174.3 444.5L177.6 443.1L178.3 440.8L181.7 442.6L182.4 440.3L184.0 439.3L182.9 437.6L186.7 437.3L187.4 431.0L189.8 431.9L195.1 425.2L197.1 426.4L197.1 428.4L198.2 429.4L201.1 429.0L200.3 430.9L201.1 433.2L200.0 434.1L202.3 436.0L198.0 442.3L197.8 444.8L202.9 446.5L199.3 448.3L196.4 453.1L196.1 454.5L198.2 455.6L198.9 457.2L198.5 463.0L197.4 464.5L198.0 466.9L197.2 469.2L193.5 470.2ZM176.1 579.7L175.2 581.3L172.2 583.0L178.8 586.9L178.3 588.9L176.6 590.5L180.5 590.6L182.2 589.6L181.9 591.3L183.2 593.3L181.1 596.5L184.8 598.0L186.4 600.0L185.0 604.3L184.9 610.1L188.0 612.9L192.2 609.7L194.5 612.6L192.9 620.7L193.6 622.5L191.9 626.7L194.1 627.7L195.1 627.0L196.9 628.7L194.4 633.9L194.1 637.2L191.5 640.5L193.8 643.5L192.0 646.6L194.1 650.3L193.9 652.0L190.5 656.9L178.8 642.9L174.0 629.7L170.8 613.8L171.8 613.5L170.7 613.7L160.8 585.4L159.1 584.2L157.8 579.1L153.6 575.4L151.4 572.1L150.3 572.1L143.3 554.9L148.4 556.7L152.0 560.3L154.9 560.9L153.8 562.8L155.0 563.7L154.8 565.6L156.3 565.6L158.3 568.7L162.9 571.0L164.5 573.3L167.3 573.8L169.6 572.7L169.6 575.5L171.6 575.4L174.0 577.9L176.0 577.8ZM247.8 399.3L251.4 403.4L250.4 408.9L248.9 410.2L250.6 411.7L251.1 415.6L250.1 416.5L254.4 419.8L257.8 418.9L259.8 421.4L261.5 420.9L265.4 423.4L266.8 425.3L267.9 431.5L270.0 430.6L270.2 432.1L272.6 431.5L273.4 440.0L274.9 440.0L274.3 441.4L272.5 440.4L270.7 441.4L271.2 443.6L273.5 446.3L279.6 448.8L278.9 450.9L277.1 450.7L276.8 452.9L273.0 453.8L271.8 453.2L269.8 457.2L264.2 454.7L259.6 458.7L264.6 460.0L264.0 463.5L258.7 460.9L258.4 458.7L255.9 457.7L253.1 458.8L253.1 464.0L250.8 466.5L247.5 464.4L239.6 467.7L236.2 468.1L235.4 469.3L235.3 475.8L231.4 475.5L228.7 476.8L229.1 478.2L225.2 480.8L222.3 479.0L221.2 480.0L218.3 479.3L215.0 480.8L213.3 484.2L211.3 484.6L207.2 483.1L205.4 483.8L199.3 482.6L199.0 475.9L201.0 473.1L194.5 471.6L193.5 470.2L197.2 469.2L198.0 466.9L197.4 464.5L198.5 463.0L198.9 457.2L198.2 455.6L196.1 454.5L196.4 453.1L199.3 448.3L202.9 446.5L197.8 444.8L198.0 442.3L202.3 436.0L200.0 434.1L201.1 433.2L200.3 430.9L201.1 429.0L200.1 428.7L199.5 426.8L201.0 422.8L204.0 422.6L203.9 420.1L205.1 419.8L206.2 417.0L208.5 416.1L206.2 414.6L205.3 410.8L206.6 408.4L206.7 405.6L208.4 404.7L213.3 407.1L214.0 402.7L216.1 401.5L215.4 396.7L216.8 396.2L217.3 394.7L216.1 391.7L218.7 393.9L227.4 395.1L227.8 396.8L229.8 397.5L229.4 400.0L232.8 400.4L234.8 402.0L236.1 399.0L238.5 399.4L240.7 401.1L246.5 398.6ZM245.3 575.0L247.4 573.2L248.5 573.3L247.8 575.8ZM246.1 596.7L246.6 595.2L248.8 595.4L248.7 599.0ZM521.9 568.2L521.5 570.3L522.5 571.2L520.9 572.8L522.7 574.9L521.9 579.1L519.9 580.2L521.8 581.2L520.0 583.5L517.6 578.6L517.3 574.3L518.7 575.7L519.9 567.4ZM524.3 566.9L524.0 568.6L522.2 568.1L521.6 566.6L523.1 565.1L524.4 565.3ZM524.4 564.7L521.2 565.4L520.7 564.8L521.4 553.2L522.8 552.6L523.8 540.5L524.4 538.7L526.9 536.3L528.0 537.4L528.5 543.2L527.8 547.6L526.2 549.1L525.2 548.2L524.1 551.7L525.5 552.1L526.6 559.9ZM517.0 597.4L518.6 602.3L517.2 606.0L515.1 605.3L513.9 600.0ZM519.5 583.4L520.6 586.5L518.6 587.2ZM544.8 680.9L546.5 687.5L545.7 691.1L544.1 691.7L540.8 683.5L541.4 682.0ZM541.9 676.2L542.9 677.8L540.6 680.8L540.0 679.2ZM522.7 635.3L523.4 637.5L521.4 637.8L521.2 635.7Z"}};

const MODE_CFG = {
  mif: {
    chip:'Prospects', chipClass:'', eyebrow:'MARKET OVERVIEW',
    title:'Market Research Dashboard',
    sub:'Steel-consuming manufacturing landscape — active customers & conversion targets.',
    hubTitle:'Search Prospects by', hubEyebrow:'MARKET INTELLIGENCE',
    stripTitle:'Coverage & Data Quality',
  },
  comp: {
    chip:'Competitors', chipClass:'comp', eyebrow:'COMPETITOR INTELLIGENCE',
    title:'Competitor Intelligence Dashboard',
    sub:'Competitors, machine makers, pipe & tube producers, rolls & dies and roll-forming rivals.',
    hubTitle:'Search Competitors by', hubEyebrow:'COMPETITOR INTELLIGENCE',
    stripTitle:'Coverage & Data Quality',
  }
};

// ── SEGMENT FRAMEWORK ─────────────────────────────────────────────
const MKT_ICONS = {
  Mobility:'<path d="M4 13l1.5-4.5A2 2 0 0 1 7.4 7h9.2a2 2 0 0 1 1.9 1.5L20 13v5h-2v2h-2v-2H8v2H6v-2H4v-5Z"/><circle cx="7.5" cy="15.5" r="1.1" fill="currentColor"/><circle cx="16.5" cy="15.5" r="1.1" fill="currentColor"/>',
  'Building & Construction':'<path d="M4 20V9l8-5 8 5v11"/><path d="M9 20v-6h6v6"/>',
  Industry:'<path d="M4 20V11l5 3V11l5 3V7l6 4v9Z"/>',
  'Power Generation':'<path d="M13 3 5 13h5l-1 8 8-11h-5l1-7Z"/>',
  'Agri - Tech':'<path d="M12 21c0-5 0-9 6-12-1 6-3 9-6 12Z"/><path d="M12 21c0-4-1-7-6-9 2 5 3 7 6 9Z"/>',
  Others:'<circle cx="6" cy="12" r="1.8" fill="currentColor"/><circle cx="12" cy="12" r="1.8" fill="currentColor"/><circle cx="18" cy="12" r="1.8" fill="currentColor"/>',
};
const MKT_COLORS = {
  Mobility:'var(--steel)', 'Building & Construction':'var(--accent)', Industry:'var(--purple)',
  'Power Generation':'var(--amber)', 'Agri - Tech':'var(--green)', Others:'#5b6678',
};
const MARKETS = [
  {name:'Mobility', segs:[
    {n:'Automotive — Passenger (LMV)', v:'Auto - Passenger (LMV)'},
    {n:'Automotive — Commercial (LCV/MCV)', v:'Auto - Commercial (LCV_MCV)'},
    {n:'Trucks & Trailers', v:'Trucks & Trailers'},
    {n:'Bus Body', v:'Bus Body'},
    {n:'Construction Equipment', v:'Construction Equipment'},
    {n:'Auto Components', v:'Auto Components'},
    {n:'Shipbuilding', v:'Shipbuilding'},
    {n:'Railways', v:'Railways'},
    {n:'Operator Cabin', v:'Operator Cabin'},
  ]},
  {name:'Building & Construction', segs:[
    {n:'Elevator & Escalator', v:'Elevator & Escalator'},
    {n:'Formwork & Scaffolding', v:'Formwork & Scaffolding'},
    {n:'Mezzanine Floors', v:'Mezzanine Floor'},
    {n:'Window & Facade', v:'Windows & Facades'},
    {n:'Interior Fittings', v:'Interior Fittings'},
    {n:'HVAC', v:'HVAC'},
    {n:'Sun Shading', v:'Sun Shading'},
    {n:'Cranes & Hoists', v:'Cranes & Hoists'},
    {n:'Parking Systems', v:'Parking Systems'},
  ]},
  {name:'Industry', segs:[
    {n:'Furniture', v:'Steel Furniture'},
    {n:'Storage Technology', v:'Storage Technology', sub:['Racks & Storage','Warehouse & Intralogistics']},
    {n:'Home Appliances', v:'Home Appliances'},
    {n:'Food Processing', v:'Food Processing'},
    {n:'Textile', v:'Textile Machinery'},
    {n:'Road Safety', v:'Road Safety'},
    {n:'Cleaning Equipment', v:'Cleaning Equipment'},
    {n:'Metal Hooks & Hardware', v:'Metal Hooks & Hardware'},
    {n:'ISO Freight Containers', v:'ISO Freight Containers'},
    {n:'Electrical Cabinets', v:'Electrical Cabinets'},
  ]},
  {name:'Power Generation', segs:[
    {n:'Solar Mounting Systems', v:'Mounting Structures'},
    {n:'Power Transmission', v:'Power Transmission'},
    {n:'Wind Energy', v:'Wind Energy'},
  ]},
  {name:'Agri - Tech', segs:[
    {n:'Greenhouse', v:'Greenhouse'},
    {n:'Fencing', v:'Fencing'},
    {n:'Agricultural Machinery', v:'Agricultural Machinery'},
  ]},
  {name:'Others', segs:[
    {n:'Traders', v:'Steel Traders'},
    {n:'Tooling', v:'Tooling'},
    {n:'Internal', v:'Internal'},
    {n:'Inter Unit', v:'Inter Unit'},
  ]},
];

// ── UTIL ──────────────────────────────────────────────────────────
function esc(s){return s==null?'':String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function fmt(n){return Number(n).toLocaleString('en-IN');}
function firstChar(v){return (v||'').trim().toUpperCase().charAt(0);}
function count(pred){return DATA.filter(pred).length;}
function segOf(c){let s=(c['Segment']||'').trim(); if(s==='Auto - Commercial (LCV/MCV)') s='Auto - Commercial (LCV_MCV)'; return s;}
function segCount(v){return DATA.filter(c=>segOf(c)===v).length;}
function money(v){v=(v||'').trim();if(!v)return '';return /[a-z₹]/i.test(v)?v:'₹'+v+' Cr';}
function revNum(c){const m=String(c['Revenue (Rs. Cr)']||'').replace(/,/g,'').match(/([\d.]+)/);return m?parseFloat(m[1]):NaN;}

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

// ── REGION NORMALISATION ──────────────────────────────────────────
const ZONE_KEYS = {
  North:['\\bnorth\\b','\\bnorthern\\b','delhi','\\bncr\\b','haryana','punjab','uttar pradesh','\\bup\\b','noida','ghaziabad','faridabad','manesar','gurgaon','gurugram','haridwar','uttarakhand','himachal','jammu','kashmir','ludhiana','chandigarh','meerut','\\bj&k\\b'],
  West:['\\bwest\\b','\\bwestern\\b','gujarat','maharashtra','pune','mumbai','ahmedabad','halol','sanand','\\bgoa\\b','rajkot','vadodara','baroda','surat','nashik','aurangabad','kolhapur','rajasthan','jaipur','jodhpur','bhiwadi','silvassa','daman'],
  South:['\\bsouth\\b','\\bsouthern\\b','tamil nadu','chennai','bangalore','bengaluru','karnataka','telangana','hyderabad','andhra','\\bap\\b','kerala','coimbatore','trichy','sriperumbudur','tumkur','hosur','salem','madurai','vizag','visakhapatnam','\\btn\\b','mysore','hosur'],
  East:['\\beast\\b','\\beastern\\b','west bengal','kolkata','jharkhand','jamshedpur','odisha','orissa','bihar','patna','assam','ranchi','durgapur','siliguri','\\bne\\b','north east','north-east'],
  Central:['\\bcentral\\b','madhya pradesh','\\bmp\\b','chhattisgarh','indore','bhopal','raipur','nagpur','pithampur'],
};
const ZONE_RE = {};
ZONES.forEach(z=>{ZONE_RE[z]=new RegExp('('+ZONE_KEYS[z].join('|')+')','i');});
function regionZones(str){
  const s=(str||'').toLowerCase(); if(!s.trim())return [];
  if(/pan[\s-]*india|national|\bvarious\b|\bmulti\b|all[\s-]*india|international/.test(s)) return ZONES.slice();
  const z=ZONES.filter(zn=>ZONE_RE[zn].test(s));
  return z;
}
function regionCount(zone){return DATA.filter(c=>regionZones(c['Region']).indexOf(zone)>=0).length;}

// ── LOGIN ─────────────────────────────────────────────────────────
async function sha256(str){
  const buf=await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
function enterLogin(){document.getElementById('welcome').style.display='none';document.getElementById('login').style.display='flex';document.getElementById('usr').focus();}
function backWelcome(){document.getElementById('login').style.display='none';document.getElementById('welcome').style.display='flex';setErr('');}
async function tryLogin(){
  const usr=(document.getElementById('usr').value||'').trim().toLowerCase();
  const pwd=document.getElementById('pwd').value;
  if(!usr){setErr('Please enter your username');return;}
  if(!pwd){setErr('Please enter your password');return;}
  const h=await sha256(pwd);
  if(usr===USER_MIF && h===PH_MIF){MODE='mif';launch();}
  else if(usr===USER_COMP && h===PH_COMP){MODE='comp';launch();}
  else{setErr('Invalid username or password.');document.getElementById('pwd').value='';document.getElementById('pwd').focus();}
}
function launch(){
  DATA=DATASETS[MODE];
  document.getElementById('login').style.display='none';
  document.getElementById('app').style.display='block';
  initApp();
}
function setErr(m){const e=document.getElementById('login-err');e.textContent=m;if(m)setTimeout(()=>{if(e.textContent===m)e.textContent='';},3000);}

// ── INIT ──────────────────────────────────────────────────────────
function initApp(){
  const cfg=MODE_CFG[MODE];
  const chip=document.getElementById('mode-chip');
  chip.textContent=cfg.chip; chip.className='hchip '+cfg.chipClass;
  document.getElementById('hts').textContent=(NEWS&&NEWS.updated)?NEWS.updated:TIMESTAMP;
  document.getElementById('synced-ts').textContent=(NEWS&&NEWS.updated)?NEWS.updated:TIMESTAMP;
  document.getElementById('dash-eyebrow').textContent=cfg.eyebrow;
  document.getElementById('dash-title').textContent=cfg.title;
  document.getElementById('dash-sub').textContent=cfg.sub;
  document.getElementById('hub-title').textContent=cfg.hubTitle;
  document.getElementById('hub-eyebrow').textContent=cfg.hubEyebrow;

  // Filter dropdowns
  const segs=[...new Set(DATA.map(segOf).filter(Boolean))].sort();
  const sf=document.getElementById('f-seg');
  segs.forEach(s=>{const o=document.createElement('option');o.value=s;o.textContent=s;sf.appendChild(o);});
  const rf=document.getElementById('f-region');
  ZONES.forEach(r=>{const n=regionCount(r);const o=document.createElement('option');o.value=r;o.textContent=r+' ('+n+')';rf.appendChild(o);});
  const revF=document.getElementById('f-rev');
  REV_BUCKETS.forEach(b=>{const o=document.createElement('option');o.value=b.k;o.textContent=b.k;revF.appendChild(o);});
  const tierF=document.getElementById('f-tier');
  const tiersPresent=[...new Set(DATA.map(c=>normTier(c['Priority Tier']||c['Competitor Tier'])))].filter(t=>t!=='Unclassified');
  TIER_ORDER.filter(t=>tiersPresent.indexOf(t)>=0).forEach(t=>{const o=document.createElement('option');o.value=t;o.textContent=t;tierF.appendChild(o);});

  buildDashboard();
  buildTicker();
  buildNewsDash();
  buildNewsPage();
  buildExhUpcoming();
  buildSegmentsPage();
  buildRegionPage();
  buildExhFilters();
  buildExhRegions();
  renderWorldBase();
  buildExhMapLegend();
  renderExhList();
  renderExhMap();
  goSearch();
}

const REV_BUCKETS=[
  {k:'Under ₹100 Cr',lo:0,hi:100},{k:'₹100 – 500 Cr',lo:100,hi:500},{k:'₹500 – 1000 Cr',lo:500,hi:1000},
  {k:'₹1000 – 5000 Cr',lo:1000,hi:5000},{k:'₹5000 Cr +',lo:5000,hi:1e12},{k:'Not disclosed',lo:-1,hi:-1},
];
function revBucketKey(c){const n=revNum(c);if(isNaN(n))return 'Not disclosed';for(const b of REV_BUCKETS){if(b.lo>=0&&n>=b.lo&&n<b.hi)return b.k;}return 'Not disclosed';}

const TIER_ORDER=['T1 Captive','T1 OEM','T1 Active Prospect','T1 PSU / Defence','T1 Other','T2 Conversion Possible','T2 Other','Pure Prospect','T3 Small','Competitor','Channel / Partner','Not Applicable','Other'];
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
function tierRank(c){const t=normTier(c['Priority Tier']||c['Competitor Tier']);const i=TIER_ORDER.indexOf(t);return i<0?99:i;}

// ── VIEW MGMT ─────────────────────────────────────────────────────
const PAGES=['dash','search','region','segments','compsearch','list','detail','exh','exhdetail','news'];
function showView(v){
  PAGES.forEach(x=>{const el=document.getElementById('view-'+x);if(!el)return;
    if(x==='detail'||x==='exhdetail'){el.classList.toggle('on',x===v);el.style.display=x===v?'flex':'none';}
    else if(x==='dash'||x==='list'){el.style.display=x===v?'block':'none';}
    else{el.style.display=x===v?'block':'none';}
  });
  window.scrollTo(0,0);
}
function setNav(nav){document.querySelectorAll('.sid-item').forEach(i=>i.classList.toggle('active',i.dataset.nav===nav));}
function goDash(){showView('dash');setNav('dashboard');closeSidebar();}
function goSearch(){showView('search');setNav('search');closeSidebar();renderHubArt();}
function goNews(){buildNewsPage();showView('news');setNav('news');closeSidebar();}
function renderHubArt(){
  const el=document.getElementById('region-art');
  if(!el||el.dataset.done)return;
  const W=MAP.vb[0],H=MAP.vb[1];
  const paths=ZONES.map(z=>MAP.paths[z]?'<path d="'+MAP.paths[z]+'" fill="#fff"/>':'').join('');
  el.innerHTML='<svg viewBox="0 0 '+W+' '+H+'" preserveAspectRatio="xMidYMid meet">'+paths+'</svg>';
  el.dataset.done='1';
}
// ── NEWS ENGINE ────────────────────────────────────
function nMoveCls(m){return m==='up'?'up':m==='down'?'down':'flat';}
function nMoveArr(m){return m==='up'?'\u25B2':m==='down'?'\u25BC':'\u25AC';}
function nStatDot(s){return s==='red'?'red':s==='green'?'green':s==='yellow'?'amber':'amber';}
function nStatHex(s){return s==='red'?'#e0455e':s==='green'?'#12a05f':s==='yellow'?'#c58a10':'#96a0b0';}
function nStClass(s){return s==='red'?'st-red':s==='green'?'st-green':s==='yellow'?'st-yellow':'st-neutral';}
function nSbClass(s){return s==='red'?'sb-red':s==='green'?'sb-green':s==='yellow'?'sb-yellow':'sb-neutral';}
function nStatLbl(s){return s==='red'?'Act':s==='green'?'Positive':s==='yellow'?'Watch':'Note';}
function hasNews(){return !!(NEWS&&NEWS.editions&&NEWS.editions.length);}
function buildTicker(){
  var el=document.getElementById('ticker');if(!el||!hasNews())return;
  var ed=NEWS.editions[0];var items=[];
  (ed.indicators||[]).forEach(function(i){items.push('<span class="tk-item"><b>'+esc(i.name)+'</b> <span class="val">'+esc(i.value)+esc(i.unit||'')+'</span> <span class="arr '+nMoveCls(i.move)+'">'+nMoveArr(i.move)+'</span></span>');});
  (ed.stories||[]).forEach(function(s){if(s.marquee){items.push('<span class="tk-item"><span class="tk-dot '+nStatDot(s.status)+'"></span><b>'+esc(s.company&&s.company!=='\u2014'?s.company:s.segment)+'</b> '+esc(s.headline)+'</span>');}});
  if(!items.length)return;
  var row=items.join('');
  document.getElementById('tick-track').innerHTML=row+row;
  el.classList.add('on');
  document.documentElement.style.setProperty('--tick','34px');
}
function buildNewsDash(){
  var card=document.getElementById('news-dash-card');if(!card)return;
  if(!hasNews()){card.style.display='none';return;}
  var ed=NEWS.editions[0];
  var ind=(ed.indicators||[]).filter(function(i){return /HR Coil|Scrap|USD/i.test(i.name);}).slice(0,3);
  var indHTML=ind.map(function(i){return '<div class="ndi"><div class="l">'+esc(i.name)+'</div><div class="v">'+esc(i.value)+esc(i.unit||'')+' <span class="arr" style="color:'+(i.move==='up'?'#0f8a56':i.move==='down'?'#c93350':'#8791a3')+'">'+nMoveArr(i.move)+'</span></div></div>';}).join('');
  var tops=[];NEWS.editions.forEach(function(e){(e.stories||[]).forEach(function(s){if(s.dashboard)tops.push(s);});});tops=tops.slice(0,4);
  var rows=tops.map(function(s){return '<div class="nd-row" onclick="goNews()"><span class="nd-dot" style="background:'+nStatHex(s.status)+'"></span><div><div class="tag">'+esc(s.section)+(s.segment&&s.segment!=='\u2014'?' \u00b7 '+esc(s.segment):'')+'</div><div class="nm">'+esc(s.headline)+'</div><div class="mt">'+esc((s.why||s.what||'').slice(0,140))+'\u2026</div></div></div>';}).join('');
  card.innerHTML='<div class="pc-head"><div class="pc-title">Market News &amp; Customer Intelligence</div><div class="pc-meta">Updated '+esc(NEWS.updated||ed.dateLabel)+'</div></div><div class="newsdash-ind">'+indHTML+'</div><div class="nd-list">'+rows+'</div><div class="nd-more" onclick="goNews()">View all news &rarr;</div>';
  card.style.display='block';
}
function buildTrendCard(tr){
  if(!tr||!tr.points||tr.points.length<2)return '';
  var c=tr.color||'#e07a2e';
  var pts=tr.points;var vals=pts.map(function(p){return p.v;});var min=Math.min.apply(null,vals),max=Math.max.apply(null,vals);
  var W=800,H=160,pad=28;
  var X=function(i){return pad+(W-2*pad)*i/(pts.length-1);};
  var Y=function(v){return H-pad-(H-2*pad)*(v-min)/((max-min)||1);};
  var line=pts.map(function(p,i){return (i?'L':'M')+X(i).toFixed(1)+' '+Y(p.v).toFixed(1);}).join(' ');
  var area=line+' L'+X(pts.length-1).toFixed(1)+' '+(H-pad)+' L'+pad.toFixed(1)+' '+(H-pad)+' Z';
  var dots=pts.map(function(p,i){return '<circle cx="'+X(i).toFixed(1)+'" cy="'+Y(p.v).toFixed(1)+'" r="3.2" fill="'+c+'"/>';}).join('');
  var labs=pts.map(function(p,i){return '<text x="'+X(i).toFixed(1)+'" y="'+(H-8)+'" text-anchor="middle" font-size="9" fill="#96a0b0">'+esc(p.label)+'</text>';}).join('');
  var vlabs=pts.map(function(p,i){return '<text x="'+X(i).toFixed(1)+'" y="'+(Y(p.v)-8).toFixed(1)+'" text-anchor="middle" font-size="9.5" fill="#5b6678" font-weight="600">'+esc(p.disp||p.v)+'</text>';}).join('');
  return '<div class="panel-card news-trend-card"><div class="pc-head"><div class="pc-title">'+esc(tr.title||'Trend')+'</div><div class="pc-meta">'+esc(tr.meta||'')+'</div></div><svg class="nt-svg" viewBox="0 0 '+W+' '+H+'" preserveAspectRatio="none"><path d="'+area+'" fill="'+c+'22"/><path d="'+line+'" fill="none" stroke="'+c+'" stroke-width="2.5"/>'+dots+vlabs+labs+'</svg></div>';
}
function renderEdition(e){
  var note=e.note?'<div class="news-note"><div><div class="h">Input Cost Note for MIF</div><div class="p">'+esc(e.note)+'</div></div></div>':'';
  var cards=(e.stories||[]).map(function(s){
    var co=(s.company&&s.company!=='\u2014')?'<span class="co">'+esc(s.company)+'</span>':'';
    var why=s.why?'<div class="why"><div class="k">Why it matters to MIF</div><div class="t">'+esc(s.why)+'</div></div>':'';
    var src=s.source?'<div class="src">'+esc(s.source)+'</div>':'';
    return '<div class="ncard"><div class="top"><span class="sbar '+nSbClass(s.status)+'"></span><span class="seg">'+esc((s.segment&&s.segment!=='\u2014')?s.segment:s.section)+'</span>'+co+'<span class="stat '+nStClass(s.status)+'">'+nStatLbl(s.status)+'</span></div><div class="bd"><div class="hl">'+esc(s.headline)+'</div><div class="k">What happened</div><div class="t">'+esc(s.what)+'</div>'+why+src+'</div></div>';
  }).join('');
  return '<div class="news-ed"><div class="news-ed-head"><div class="news-ed-date">'+esc(e.dateLabel)+'</div><span class="news-ed-badge">'+esc(e.type||'Daily')+' Briefing</span><div class="news-ed-line"></div></div>'+note+'<div class="news-grid">'+cards+'</div></div>';
}
function buildNewsPage(){
  var ind=document.getElementById('news-indicators'),feed=document.getElementById('news-feed'),trend=document.getElementById('news-trend'),empty=document.getElementById('news-empty');
  if(!feed)return;
  if(!hasNews()){if(empty)empty.style.display='block';return;}
  var ed0=NEWS.editions[0];
  ind.innerHTML=(ed0.indicators||[]).map(function(i){return '<div class="nis"><div class="l">'+esc(i.name)+'</div><div class="v">'+esc(i.value)+' <span class="u">'+esc(i.unit||'')+'</span></div><div class="mv '+nMoveCls(i.move)+'">'+nMoveArr(i.move)+' '+(i.move==='up'?'Up':i.move==='down'?'Down':'Stable')+'</div><div class="dr">'+esc(i.driver||'')+'</div></div>';}).join('');
  trend.innerHTML=(NEWS.trends||(NEWS.trend&&NEWS.trend.points?[NEWS.trend]:[])).map(function(t){return buildTrendCard(t);}).join('');
  feed.innerHTML=NEWS.editions.map(function(e){return renderEdition(e);}).join('');
}
function goExhibitions(){showView('exh');setNav('exh');closeSidebar();}
function goRegion(){buildRegionPage();showView('region');setNav('search');closeSidebar();}
function goSegments(){showView('segments');setNav('search');closeSidebar();}
function goCompSearch(){showView('compsearch');setNav('search');closeSidebar();document.getElementById('cs-input').focus();renderCompSearch();}

// ── DASHBOARD ─────────────────────────────────────────────────────
const KPI_COLORS={accent:'var(--accent)',steel:'#2f6fe0',green:'#12a05f',red:'#e0455e',amber:'#c58a10',purple:'#7a5cf0'};
function buildDashboard(){
  const segCounts={};
  DATA.forEach(c=>{const s=segOf(c); if(s) segCounts[s]=(segCounts[s]||0)+1;});
  const nSeg=Object.keys(segCounts).length;
  const highConf=count(c=>c['Data Confidence']==='High');
  let kpis;
  if(MODE==='comp'){
    kpis=[
      {v:DATA.length, l:'Competitor Entities', s:`across ${nSeg} classes`, c:'red', go:()=>openList('all','All Competitors')},
      {v:segCount('Roll Forming Competitors'), l:'Roll-Forming Rivals', s:'direct competition', c:'accent', go:()=>openSeg('Roll Forming Competitors')},
      {v:segCount('Pipes & Tubes'), l:'Pipe & Tube Makers', s:'tube / hollow section', c:'steel', go:()=>openSeg('Pipes & Tubes')},
      {v:segCount('Machine Makers'), l:'Machine Manufacturers', s:'equipment builders', c:'purple', go:()=>openSeg('Machine Makers')},
      {v:segCount('Rolls & Dies'), l:'Rolls & Dies Makers', s:'tooling suppliers', c:'amber', go:()=>openSeg('Rolls & Dies')},
    ];
  } else {
    const prosp=count(c=>(c._cat||'').indexOf('Prospect')>=0);
    kpis=[
      {v:DATA.length, l:'Companies Tracked', s:`across ${nSeg} segments`, c:'accent', go:()=>openList('all','All Companies')},
      {v:nSeg, l:'Market Segments', s:'tracked verticals', c:'purple', go:()=>goSegments()},
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

  document.getElementById('seg-meta').textContent=`${nSeg} segments · ${fmt(DATA.length)} companies`;
  let covHTML='';
  const usedSegs=new Set();
  const covRow=(name,val,seg)=>`<div class="seg-cov-row" onclick="openSeg('${esc(seg).replace(/'/g,"\\'")}')"><span class="nm" title="${esc(name)}">${esc(name)}</span><span class="ct">${fmt(val)}</span></div>`;
  if(MODE==='comp'){
    const rows=Object.entries(segCounts).sort((a,b)=>b[1]-a[1]);
    covHTML=`<div class="seg-cov-grp"><span class="mk-dot" style="background:var(--red)"></span><span class="mk-nm">Competitor Classes</span><span class="mk-tot">${fmt(DATA.length)}</span></div>`+
      rows.map(([s,v])=>covRow(s,v,s)).join('');
  } else {
    MARKETS.forEach(m=>{
      const rows=m.segs.map(s=>({n:s.n,v:s.v,c:segCount(s.v)})).filter(r=>r.c>0).sort((a,b)=>b.c-a.c);
      if(!rows.length)return;
      rows.forEach(r=>usedSegs.add(r.v));
      const tot=rows.reduce((a,r)=>a+r.c,0), col=MKT_COLORS[m.name]||'#5b6678';
      covHTML+=`<div class="seg-cov-grp"><span class="mk-dot" style="background:${col}"></span><span class="mk-nm">${esc(m.name)}</span><span class="mk-tot">${fmt(tot)}</span></div>`+
        rows.map(r=>covRow(r.n,r.c,r.v)).join('');
    });
    const others=Object.entries(segCounts).filter(([s])=>!usedSegs.has(s)).sort((a,b)=>b[1]-a[1]);
    if(others.length){
      const tot=others.reduce((a,[,v])=>a+v,0);
      covHTML+=`<div class="seg-cov-grp"><span class="mk-dot" style="background:#5b6678"></span><span class="mk-nm">Other Segments</span><span class="mk-tot">${fmt(tot)}</span></div>`+
        others.map(([s,v])=>covRow(s,v,s)).join('');
    }
  }
  document.getElementById('seg-cov').innerHTML=covHTML;

  buildRegionPie();

}
function animateCount(el,target){
  const dur=1100,t0=performance.now();
  function tick(now){const p=Math.min(1,(now-t0)/dur);const e=1-Math.pow(1-p,3);el.textContent=fmt(Math.round(target*e));if(p<1)requestAnimationFrame(tick);}
  requestAnimationFrame(tick);
  setTimeout(()=>{el.textContent=fmt(target);},1200);
}

// ── SEGMENTS PAGE ─────────────────────────────────────────────────
function buildSegmentsPage(){
  const grid=document.getElementById('mkt-grid');
  if(MODE==='comp'){
    const segCounts={};
    DATA.forEach(c=>{const s=segOf(c);if(s)segCounts[s]=(segCounts[s]||0)+1;});
    const rows=Object.entries(segCounts).sort((a,b)=>b[1]-a[1]);
    grid.style.gridTemplateColumns='1fr';
    grid.innerHTML=`<div class="mkt-card" style="max-width:640px">
      <div class="mkt-top"><div class="mkt-ic" style="background:color-mix(in srgb,var(--red) 13%,#fff);color:var(--red)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 20V11l5 3V11l5 3V7l6 4v9Z"/></svg></div>
      <div><div class="mkt-nm">Competitor Classes</div><div class="mkt-ct">${DATA.length} entities · ${rows.length} classes</div></div></div>
      <div class="mkt-body">${rows.map(([v,n])=>`<div class="seg-chip" onclick="openSeg('${esc(v).replace(/'/g,"\\'")}')"><span class="sn">${esc(v)}</span><span class="sc">${n}</span></div>`).join('')}</div></div>`;
    return;
  }
  grid.style.gridTemplateColumns='';
  grid.innerHTML=MARKETS.map(m=>{
    const total=m.segs.reduce((a,s)=>a+segCount(s.v),0);
    const nActive=m.segs.filter(s=>segCount(s.v)>0).length;
    const chips=m.segs.map(s=>{
      const n=segCount(s.v);
      const sub=s.sub?`<div class="seg-sub">${s.sub.map(x=>`<div class="ss">${esc(x)}</div>`).join('')}</div>`:'';
      if(n>0) return `<div class="seg-chip" onclick="openSeg('${esc(s.v).replace(/'/g,"\\'")}')"><span class="sn">${esc(s.n)}</span><span class="sc">${n}</span></div>${sub}`;
      return `<div class="seg-chip void"><span class="sn">${esc(s.n)}</span><span class="sc">&mdash;</span></div>${sub}`;
    }).join('');
    const col=MKT_COLORS[m.name];
    return `<div class="mkt-card">
      <div class="mkt-top"><div class="mkt-ic" style="background:color-mix(in srgb,${col} 13%,#fff);color:${col}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">${MKT_ICONS[m.name]}</svg></div>
        <div><div class="mkt-nm">${esc(m.name)}</div><div class="mkt-ct">${fmt(total)} companies · ${nActive} segments</div></div></div>
      <div class="mkt-body">${chips}</div></div>`;
  }).join('');
}

// ── REGION PAGE ───────────────────────────────────────────────────
const ZONE_COLORS={North:'var(--steel)',West:'var(--accent)',Central:'var(--amber)',East:'var(--purple)',South:'var(--green)'};
function buildRegionPage(){
  const W=MAP.vb[0],H=MAP.vb[1];
  const paths=ZONES.map(z=>MAP.paths[z]?`<path class="rzone-path" data-zone="${z}" d="${MAP.paths[z]}" fill="${ZONE_MAP_FILL[z]}" onclick="openRegion('${z}')"><title>${z} India — ${fmt(regionCount(z))} companies</title></path>`:'').join('');
  const labels=ZONES.map(z=>{const c=MAP.cen[z];if(!c)return '';return `<g class="rzone-lbl"><text class="zn" x="${c[0]}" y="${c[1]}">${z}</text><text class="zc" x="${c[0]}" y="${c[1]+18}">${fmt(regionCount(z))}</text></g>`;}).join('');
  document.getElementById('india-map').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Map of India by region">${paths}${labels}</svg>`;
  const lg=document.getElementById('map-legend');
  if(lg)lg.innerHTML=ZONES.map(z=>`<div class="lg" onclick="openRegion('${z}')"><span class="sw" style="background:${ZONE_MAP_FILL[z]}"></span>${z} <span class="lc">${fmt(regionCount(z))}</span></div>`).join('');
  const side=document.getElementById('region-side');
  side.innerHTML=ZONES.map(z=>{
    const n=regionCount(z);const col=ZONE_COLORS[z];
    return `<div class="rs-item" onclick="openRegion('${z}')">
      <div class="rs-sq" style="background:${col}">${z.charAt(0)}</div>
      <div><div class="rs-nm">${z}</div><div class="rs-mt">${z==='Central'?'Heartland belt':z+' zone'} coverage</div></div>
      <div class="rs-ct">${fmt(n)}</div><div class="rs-arw">&rsaquo;</div></div>`;
  }).join('');
}
function buildRegionPie(){
  const rows=ZONES.map(z=>({z,n:regionCount(z),c:ZONE_COLORS[z]})).filter(d=>d.n>0).sort((a,b)=>b.n-a.n);
  const sum=rows.reduce((a,d)=>a+d.n,0)||1;
  const cx=90,cy=90,R=80,r=50;let a0=-Math.PI/2;
  const slices=rows.map(d=>{
    const frac=d.n/sum, a1=a0+frac*2*Math.PI, large=frac>0.5?1:0;
    const x0=cx+R*Math.cos(a0),y0=cy+R*Math.sin(a0),x1=cx+R*Math.cos(a1),y1=cy+R*Math.sin(a1);
    const xi1=cx+r*Math.cos(a1),yi1=cy+r*Math.sin(a1),xi0=cx+r*Math.cos(a0),yi0=cy+r*Math.sin(a0);
    const dp=`M${x0.toFixed(1)} ${y0.toFixed(1)}A${R} ${R} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)}L${xi1.toFixed(1)} ${yi1.toFixed(1)}A${r} ${r} 0 ${large} 0 ${xi0.toFixed(1)} ${yi0.toFixed(1)}Z`;
    a0=a1;
    return `<path class="pie-slice" d="${dp}" fill="${d.c}" onclick="openRegion('${d.z}')"><title>${d.z}: ${fmt(d.n)}</title></path>`;
  }).join('');
  const svg=`<svg width="180" height="180" viewBox="0 0 180 180">${slices}<text x="90" y="86" text-anchor="middle" style="font-family:var(--disp);font-weight:700;font-size:24px;fill:var(--t1)">${fmt(DATA.length)}</text><text x="90" y="103" text-anchor="middle" style="font-size:9px;letter-spacing:1.5px;font-weight:700;fill:#96a0b0">COMPANIES</text></svg>`;
  document.getElementById('reg-pie').innerHTML=svg;
  document.getElementById('reg-meta').textContent=`${rows.length} zones`;
  document.getElementById('reg-legend').innerHTML=rows.map(d=>`<div class="pie-lg" onclick="openRegion('${d.z}')"><span class="sw" style="background:${d.c}"></span><span class="nm">${d.z}</span><span class="vl">${fmt(d.n)}</span><span class="pc">${Math.round(d.n/sum*100)}%</span></div>`).join('');
}

// ── COMPANY SEARCH PAGE ───────────────────────────────────────────
const CS_HINTS=['CIN','Priority Tier','High Strategic Fit','Roll Forming','IATF Certified'];
function renderCompSearch(){
  document.getElementById('cs-hints').innerHTML=CS_HINTS.map(h=>`<span class="cs-hint" onclick="csFill('${esc(h).replace(/'/g,"\\'")}')">${esc(h)}</span>`).join('');
  onCompSearch();
}
function csFill(t){document.getElementById('cs-input').value=t;onCompSearch();}
let csTimer;
function onCompSearch(){clearTimeout(csTimer);csTimer=setTimeout(()=>{
  const q=(document.getElementById('cs-input').value||'').toLowerCase().trim();
  const res=document.getElementById('cs-res');
  if(!q){res.innerHTML='';return;}
  const hits=DATA.filter(c=>SEARCH_FIELDS.some(f=>c[f]&&String(c[f]).toLowerCase().indexOf(q)>=0));
  if(!hits.length){res.innerHTML=`<div class="no-res"><div class="ic">&#8981;</div><div class="m">No matches</div><div>Try a different company, CIN or tier.</div></div>`;return;}
  const top=hits.slice(0,8);
  res.innerHTML=`<div class="cs-count"><b>${fmt(hits.length)}</b> match${hits.length===1?'':'es'}</div>`+
    top.map(c=>{const i=DATA.indexOf(c);const seg=segOf(c)||c['Primary Segment']||'';
      return `<div class="cs-row" onclick="openDetail(${i},'compsearch')"><div><div class="cn">${esc(c['Company Name']||'—')}</div><div class="cm">${esc([c['Region'],c['CIN']].filter(Boolean).join(' · '))}</div></div>${seg?`<span class="ct">${esc(seg)}</span>`:''}</div>`;
    }).join('')+
    (hits.length>8?`<div class="cs-more"><button onclick="csViewAll()">View all ${fmt(hits.length)} results &rarr;</button></div>`:'');
},180);}
function csViewAll(){const q=document.getElementById('cs-input').value;openList('all','Search results');document.getElementById('search-box').value=q;state.search=q.toLowerCase();state.lastView='compsearch';renderList();updateClear();}

// ── OPENERS (into list) ───────────────────────────────────────────
function resetFilterState(){state.search='';state.seg='';state.region='';state.rev='';state.tier='';state.pv='';state.sort='az';state.page=1;
  ['f-seg','f-region','f-rev','f-tier','f-pv'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('f-sort').value='az';document.getElementById('search-box').value='';}
function openList(ctx,label){state.lastView='search';resetFilterState();state.ctx=ctx;state.label=label;document.getElementById('view-label').textContent=label;showView('list');setNav('search');closeSidebar();renderList();updateClear();}
function openSeg(v){openList('all','Segment · '+v);state.seg=v;document.getElementById('f-seg').value=v;state.lastView='segments';renderList();updateClear();}
function openRegion(z){openList('all','Region · '+z);state.region=z;document.getElementById('f-region').value=z;state.lastView='region';renderList();updateClear();}
function openTier(t){openList('all','Tier · '+t);state.tier=t;document.getElementById('f-tier').value=t;state.lastView='search';renderList();updateClear();}
function listBack(){const v=state.lastView||'search';if(v==='segments')goSegments();else if(v==='region')goRegion();else if(v==='compsearch')goCompSearch();else goSearch();}

// ── QUICK SEARCH (header) ─────────────────────────────────────────
let qTimer;
function onQuickSearch(){clearTimeout(qTimer);qTimer=setTimeout(()=>{
  const q=document.getElementById('search-box').value;
  if(document.getElementById('view-list').style.display!=='block'){openList('all','Search results');}
  state.search=q.toLowerCase();state.lastView='search';state.page=1;renderList();updateClear();
},200);}

// ── FILTER / SEARCH ───────────────────────────────────────────────
function applyFilters(){
  state.seg=document.getElementById('f-seg').value;
  state.region=document.getElementById('f-region').value;
  state.rev=document.getElementById('f-rev').value;
  state.tier=document.getElementById('f-tier').value;
  state.pv=document.getElementById('f-pv').value;
  state.sort=document.getElementById('f-sort').value;
  state.page=1;renderList();updateClear();
}
function clearFilters(){resetFilterState();state.ctx='all';renderList();updateClear();}
function updateClear(){
  const active=state.seg||state.region||state.rev||state.tier||state.pv||state.search;
  document.getElementById('clear-filters').style.display=active?'block':'none';
}
const SEARCH_FIELDS=['Company Name','Segment','Products','BD Notes','Key Products for MIF','HQ Address','Plant Locations','Director Name','Procurement Head','Sales / BD Head','Primary Segment','CIN','GSTIN','Revenue Band','Customer OEMs Served','Parent / Group','Priority Tier','Competitor Tier','Entry Route','Current Steel Suppliers'];
function matches(c){
  const ctx=state.ctx;
  if(ctx==='Customer'&&(c._cat||'').indexOf('Customer')<0)return false;
  if(ctx==='Prospect'&&(c._cat||'').indexOf('Prospect')<0)return false;
  if(state.seg&&segOf(c)!==state.seg)return false;
  if(state.region&&regionZones(c['Region']).indexOf(state.region)<0)return false;
  if(state.rev&&revBucketKey(c)!==state.rev)return false;
  if(state.tier&&normTier(c['Priority Tier']||c['Competitor Tier'])!==state.tier)return false;
  if(state.pv&&firstChar(c['Volume L/M/H'])!==state.pv)return false;
  if(state.search){const q=state.search;return SEARCH_FIELDS.some(f=>c[f]&&String(c[f]).toLowerCase().indexOf(q)>=0);}
  return true;
}
function sortList(arr){
  if(state.sort==='az') arr.sort((a,b)=>(a['Company Name']||'').localeCompare(b['Company Name']||''));
  else if(state.sort==='rev') arr.sort((a,b)=>{const x=revNum(a),y=revNum(b);if(isNaN(x)&&isNaN(y))return 0;if(isNaN(x))return 1;if(isNaN(y))return -1;return y-x;});
  else if(state.sort==='tier') arr.sort((a,b)=>tierRank(a)-tierRank(b)||(a['Company Name']||'').localeCompare(b['Company Name']||''));
  return arr;
}

// ── LIST RENDER ───────────────────────────────────────────────────
function renderList(){
  state.filtered=sortList(DATA.filter(matches));
  const total=state.filtered.length;
  const pages=Math.ceil(total/state.perPage);
  if(state.page>pages)state.page=1;
  const start=(state.page-1)*state.perPage;
  const slice=state.filtered.slice(start,start+state.perPage);
  document.getElementById('result-count').innerHTML= total===0?'No companies found':
    `<b>${fmt(total)}</b> ${total===1?'company':'companies'} &middot; showing ${start+1}\u2013${Math.min(start+state.perPage,total)}`;
  const cl=document.getElementById('company-list');
  if(total===0){cl.innerHTML=`<div class="no-res"><div class="ic">&#8981;</div><div class="m">No companies match</div><div>Try adjusting your filters or search.</div></div>`;document.getElementById('pager').innerHTML='';return;}
  cl.innerHTML=slice.map((c,i)=>{
    const idx=DATA.indexOf(c);const seg=segOf(c)||c['Primary Segment']||'';
    return `<div class="card" onclick="openDetail(${idx},'list')">
      <span class="num">${String(start+i+1).padStart(2,'0')}</span>
      <span class="cname">${esc(c['Company Name']||'\u2014')}</span>
      ${seg?`<span class="seg-tag">${esc(seg)}</span>`:''}
      <span class="chev">&rsaquo;</span></div>`;
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
function goPage(p){state.page=p;renderList();document.getElementById('view-list').scrollIntoView?window.scrollTo(0,0):null;}

// ── DETAIL FULL PAGE ──────────────────────────────────────────────
function fitMeta(c){const f=firstChar(c['Strategic Fit H/M/L']);
  if(f==='H')return{text:'High Fit',bg:'#e6f6ee',fg:'#0f8a56'};
  if(f==='M')return{text:'Medium Fit',bg:'#fdf3e0',fg:'#b5790a'};
  if(f==='L')return{text:'Low Fit',bg:'#eef1f6',fg:'#5b6678'};
  return null;}
function dfield(k,v,opts){opts=opts||{};if(!v||!String(v).trim())return '';
  let val;const s=String(v).trim();
  if(opts.type==='fit'){val=`<div class="dval hl-${firstChar(v)}">${esc(v)}</div>`;}
  else if(opts.type==='conf'){const ch=v==='High'?'H':v==='Medium'?'M':'L';val=`<div class="dval hl-${ch}">${esc(v)}</div>`;}
  else if(s.indexOf('http')===0){const u=esc(s);val=`<div class="dval"><a href="${u}" target="_blank" rel="noopener">${u.replace(/^https?:\/\/(www\.)?/,'').slice(0,42)} \u2197</a></div>`;}
  else val=`<div class="dval${opts.big?' big':''}">${esc(v)}</div>`;
  return `<div class="dfield${opts.full?' full':''}"><div class="dkey">${esc(k)}</div>${val}</div>`;
}
function fieldBlock(rows,cls){
  const ncols=(cls&&cls.indexOf('three')>=0)?3:2;
  const cells=[];let col=0;
  for(const r of rows){
    const html=dfield(r[0],r[1],r[2]);
    if(!html)continue;
    const isFull=r[2]&&r[2].full;
    if(isFull){
      while(col>0&&col<ncols){cells.push('<div class="dfield dfill"></div>');col++;}
      cells.push(html);col=0;
    }else{
      cells.push(html);col++;if(col===ncols)col=0;
    }
  }
  while(col>0&&col<ncols){cells.push('<div class="dfield dfill"></div>');col++;}
  const h=cells.join('');
  return h?`<div class="dfields${cls?' '+cls:''}">${h}</div>`:'';
}

function openDetail(idx,from){
  const c=DATA[idx];state.detailFrom=from||'list';
  document.getElementById('det-eyebrow').textContent=MODE==='comp'?'COMPETITOR DOSSIER':'COMPANY DOSSIER';
  document.getElementById('det-name').textContent=c['Company Name']||'\u2014';
  const fm=fitMeta(c);
  let tags='';
  const seg=segOf(c);if(seg)tags+=`<span class="dtag" style="background:#eaf1fd;color:#2a63cf">${esc(seg)}</span>`;
  if(c['Region'])tags+=`<span class="dtag" style="background:#e6f6ee;color:#0f8a56">${esc(c['Region'].length>18?c['Region'].slice(0,16)+'…':c['Region'])}</span>`;
  const tier=c['Priority Tier']||c['Competitor Tier']||'';
  if(tier)tags+=`<span class="dtag" style="background:color-mix(in srgb,var(--accent) 12%,#fff);color:var(--accent)">${esc(tier.length>26?tier.slice(0,24)+'…':tier)}</span>`;
  if(fm)tags+=`<span class="dtag" style="background:${fm.bg};color:${fm.fg}">${fm.text}</span>`;
  document.getElementById('det-tags').innerHTML=tags;

  // logo
  const dom=logoDomain(c);const logoEl=document.getElementById('det-logo');
  if(dom){logoEl.style.display='flex';const img=document.getElementById('det-logo-img');img.onerror=()=>{logoEl.style.display='none';};img.src='https://icons.duckduckgo.com/ip3/'+dom+'.ico';img.alt=dom;}
  else logoEl.style.display='none';

  // Card 1 — Company Information (Identity + Operations)
  document.getElementById('dc-info').innerHTML=
    '<div class="dsub">Company Identity</div>'+
    fieldBlock([
      ['Legal Entity Name',c['Company Name'],{full:1,big:1}],
      ['CIN',c['CIN']],['GSTIN',c['GSTIN']],
      ['Legal Form',c['Legal Form']],['Year Established',c['Year Est.']],
      ['Parent / Group',c['Parent / Group'],{full:1}],
      ['Website',c['Company Website'],{full:1}],
    ])+
    '<div class="dsub">Operations</div>'+
    fieldBlock([
      ['Employees',c['Employees']],['No. of Plants',c['Plant Count']],
      ['Plant Locations',c['Plant Locations'],{full:1}],
      ['HQ Address',c['HQ Address'],{full:1}],
    ]);

  // Card 2 — Financial Profile
  document.getElementById('dc-fin').innerHTML=
    fieldBlock([
      ['Revenue',money(c['Revenue (Rs. Cr)']),{big:1}],['Revenue Band',c['Revenue Band']],
      ['EBITDA / PAT',money(c['EBITDA or PAT (Rs. Cr)'])],['Growth / 3-yr CAGR',c['Revenue Growth YoY / 3-yr CAGR']],
      ['Revenue Source',c['Revenue Source'],{full:1}],
      ['Capex (last 24m)',c['Capex Announced (Last 24 months)'],{full:1}],
      ['Annual Steel Tonnage (TPA)',c['Approximate Annual Steel Tonnage (TPA)']],['Buying Pattern',c['Buying Pattern']],
    ]);

  // Card 3 — Contact Information
  const contacts=[
    ['Director / Promoter',c['Director Name'],c['Director Contact']],
    ['Procurement Head',c['Procurement Head'],c['Proc. Contact']],
    ['Sales / BD Head',c['Sales / BD Head'],c['Sales Contact']],
  ].filter(x=>x[1]||x[2]);
  let contactHTML='<div class="contact-grid">'+contacts.map(x=>{
    const who=x[1]&&String(x[1]).trim()?esc(x[1]):'<span style="color:#a7b0be">Not listed</span>';
    const cc=x[2]&&String(x[2]).trim()&&x[2]!=='—'?`<div class="cc">${esc(x[2])}</div>`:`<div class="cc muted">No contact on file</div>`;
    return `<div class="ccard"><div class="role">${esc(x[0])}</div><div class="who">${who}</div>${cc}</div>`;
  }).join('')+'</div>';
  contactHTML+=fieldBlock([
    ['HQ Address',c['HQ Address'],{full:1}],
    ['Customer OEMs Served',c['Customer OEMs Served'],{full:1}],
  ],'');
  document.getElementById('dc-contact').innerHTML=contacts.length||c['HQ Address']?contactHTML:'<div style="padding:22px;color:#a7b0be;font-size:12.5px">No contact information on file.</div>';

  // Card 4 — Company Intelligence
  document.getElementById('dc-intel').innerHTML=
    '<div class="dsub">Products & Steel Buying</div>'+
    fieldBlock([
      ['Primary Segment',c['Primary Segment']],['Products',c['Products'],{full:1}],
      ['Coil-to-Component',c['Coil-to-Component']],['In-house Roll Forming',c['In-house Roll Forming']],
      ['HSN Codes',c['HSN Codes']],['Steel Types',c['Steel Types']],
      ['Total Volume Produced',c['Total Volume Produced'],{full:1}],
      ['Roll-Form Source',c['Roll-Form Source'],{full:1}],
      ['Current Steel Suppliers',c['Current Steel Suppliers'],{full:1}],
    ],'three')+
    '<div class="dsub">MIF Business Intelligence</div>'+
    fieldBlock([
      ['Strategic Fit',c['Strategic Fit H/M/L'],{type:'fit'}],['Volume Potential (PV)',c['Volume L/M/H']],
      ['Priority Tier',c['Priority Tier']],['Relationship Status',c['Relationship Status']],
      ['IATF Certified',c['IATF Certified?']],['MIF Proximity (km)',c['MIF Proximity (km est.)']],
      ['Key Products for MIF',c['Key Products for MIF'],{full:1}],
      ['BD Notes',c['BD Notes'],{full:1}],['Entry Route',c['Entry Route'],{full:1}],
      ['Action Owner',c['Action Owner']],['Competitor Tier',c['Competitor Tier']],
      ['Switching Difficulty',c['Switching Difficulty']],
    ],'three')+
    '<div class="dsub">Data Quality</div>'+
    fieldBlock([
      ['Data Confidence',c['Data Confidence'],{type:'conf'}],['Last Verified',c['Last Verified Date']],
      ['Verification Method',c['Verification Method']],
      ['Estimated vs Verified',c['Estimated vs Verified Flag'],{full:1}],
      ['Source Hyperlink(s)',c['Source Hyperlink(s)'],{full:1}],
    ],'three');

  showView('detail');setNav('search');
  document.getElementById('dc-info').scrollTop=0;document.getElementById('dc-intel').scrollTop=0;
}

function closeDetail(){const from=state.detailFrom;if(from==='compsearch')goCompSearch();else if(from==='list'){showView('list');setNav('search');}else goSearch();}

// ── EXHIBITIONS ────────────────────────────────────────────────────
function buildExhUpcoming(){
  const now=new Date();const curM=now.getMonth(),curY=now.getFullYear();
  const nextM=(curM+1)%12,nextY=curM===11?curY+1:curY;
  const rows=EXHIBITIONS.map((e,i)=>({e,i})).filter(o=>o.e.months.some(mo=>(mo.y===curY&&mo.m===curM)||(mo.y===nextY&&mo.m===nextM)));
  rows.sort((a,b)=>{const ma=a.e.months.find(mo=>(mo.y===curY&&mo.m===curM)||(mo.y===nextY&&mo.m===nextM))||{y:9999,m:0};
    const mb=b.e.months.find(mo=>(mo.y===curY&&mo.m===curM)||(mo.y===nextY&&mo.m===nextM))||{y:9999,m:0};
    return (ma.y-mb.y)||(ma.m-mb.m)||a.e.name.localeCompare(b.e.name);});
  document.getElementById('exh-up-meta').textContent=rows.length+' in '+MONTH_NAMES[curM]+' – '+MONTH_NAMES[nextM];
  const list=document.getElementById('exh-up-list');
  if(!rows.length){list.innerHTML='<div class="exh-up-empty">No exhibitions scheduled this month or next.</div>';return;}
  list.innerHTML=rows.map(o=>{
    const e=o.e;const mo=e.months.find(mo=>(mo.y===curY&&mo.m===curM)||(mo.y===nextY&&mo.m===nextM))||{};
    const pc=PRI_CFG[e.priority]||PRI_CFG.WATCH;
    const loc=[e.city,e.country].filter(Boolean).join(', ');
    return '<div class="exh-up-row" onclick="openExhDetail('+o.i+')">'+
      '<div class="datebox"><div class="mo">'+(MONTH_NAMES[mo.m]||'—')+'</div><div class="yr">'+(mo.y||'')+'</div></div>'+
      '<div class="body"><div class="nm">'+esc(e.name)+'</div><div class="mt">'+esc(loc)+'</div></div>'+
      '<span class="pri-badge" style="background:'+pc.bg+';color:'+pc.fg+'">'+pc.label+'</span></div>';
  }).join('');
}

function buildExhFilters(){
  const counts={ALL:EXHIBITIONS.length};
  ['MUST_ATTEND','ATTEND','WATCH','AVOID'].forEach(k=>counts[k]=EXHIBITIONS.filter(e=>e.priority===k).length);
  const order=['ALL','MUST_ATTEND','ATTEND','WATCH','AVOID'];
  document.getElementById('exh-chips').innerHTML=order.map(k=>{
    const cfg=k==='ALL'?{label:'All',fg:'#4c5769'}:PRI_CFG[k];
    const active=exhState.priority===k||(k==='ALL'&&!exhState.priority);
    const style=active?('background:'+cfg.fg+';border-color:'+cfg.fg):'';
    const dotStyle=active?'background:#fff':'background:'+cfg.fg;
    return '<div class="exh-chip'+(active?' active':'')+'" style="'+style+'" onclick="setExhPriority(\''+(k==='ALL'?'':k)+'\')"><span class="dot" style="'+dotStyle+'"></span>'+cfg.label+' <span class="ct">'+counts[k]+'</span></div>';
  }).join('');
  const segs=[...new Set(EXHIBITIONS.map(e=>e.segment).filter(Boolean))].sort();
  const sel=document.getElementById('exh-seg-sel');
  const prevVal=sel.value;
  sel.innerHTML='<option value="">All Segments</option>'+segs.map(s=>'<option value="'+esc(s)+'">'+esc(s)+'</option>').join('');
  sel.value=prevVal;
}
function setExhPriority(p){exhState.priority=p;buildExhFilters();renderExhList();renderExhMap();}
function buildExhRegions(){
  const order=['India','Asia','Europe','North America','South America','Africa','Oceania','Other'];
  const counts={};
  EXHIBITIONS.forEach(e=>{
    if(e.country==='India')counts['India']=(counts['India']||0)+1;
    const c=continentOf(e.country);counts[c]=(counts[c]||0)+1;
  });
  const icons={'India':'IN','Asia':'AS','Europe':'EU','North America':'NA','South America':'SA','Africa':'AF','Oceania':'OC','Other':'\u00b7\u00b7'};
  document.getElementById('exh-regions').innerHTML=order.filter(k=>counts[k]>0).map(k=>{
    const active=exhState.continent===k;
    return '<div class="exh-region-card'+(active?' active':'')+'" onclick="setExhContinent(\''+k+'\')"><div class="rc-ic">'+icons[k]+'</div><div><div class="rc-nm">'+k+'</div><div class="rc-ct">'+counts[k]+' exhibitions</div></div></div>';
  }).join('');
}
function setExhContinent(k){exhState.continent=exhState.continent===k?'':k;buildExhRegions();renderExhList();renderExhMap();}
function applyExhFilters(){
  exhState.search=(document.getElementById('exh-search-input').value||'').toLowerCase();
  exhState.seg=document.getElementById('exh-seg-sel').value;
  renderExhList();renderExhMap();
}
function exhMatches(e){
  if(exhState.priority&&e.priority!==exhState.priority)return false;
  if(exhState.seg&&e.segment!==exhState.seg)return false;
  if(exhState.continent){
    if(exhState.continent==='India'){if(e.country!=='India')return false;}
    else if(continentOf(e.country)!==exhState.continent)return false;
  }
  if(exhState.search){
    const q=exhState.search;
    return ((e.name||'')+' '+(e.segment||'')+' '+(e.country||'')+' '+(e.city||'')+' '+(e.keyCustomers||'')).toLowerCase().indexOf(q)>=0;
  }
  return true;
}
function exhSortKey(e){return e.months&&e.months.length?e.months[0]:{y:9999,m:99};}
function exhDateBox(e){
  if(e.months&&e.months.length){
    const m0=e.months[0];
    return '<div class="datebox"><div class="mo">'+(MONTH_NAMES[m0.m]||'TBD')+'</div><div class="yr">'+(m0.y||'')+'</div></div>';
  }
  return '<div class="datebox"><div class="mo">TBD</div><div class="yr">&mdash;</div></div>';
}
function renderExhList(){
  const rows=EXHIBITIONS.map((e,i)=>({e,i})).filter(o=>exhMatches(o.e));
  rows.sort((a,b)=>{const ma=exhSortKey(a.e),mb=exhSortKey(b.e);return (ma.y-mb.y)||(ma.m-mb.m)||a.e.name.localeCompare(b.e.name);});
  document.getElementById('exh-list-meta').textContent=fmt(rows.length)+' of '+fmt(EXHIBITIONS.length)+' exhibitions';
  const list=document.getElementById('exh-list');
  if(!rows.length){list.innerHTML='<div class="exh-empty">No exhibitions match these filters.</div>';return;}
  list.innerHTML=rows.map(o=>{
    const e=o.e;const pc=PRI_CFG[e.priority]||PRI_CFG.WATCH;
    const loc=[e.city,e.country].filter(Boolean).join(', ');
    return '<div class="exh-card" onclick="openExhDetail('+o.i+')">'+
      exhDateBox(e)+
      '<div class="body"><div class="nm">'+esc(e.name)+'</div><div class="mt">'+esc(loc)+' &middot; '+esc(e.dateRaw)+'</div></div>'+
      '<span class="seg-tag">'+esc(e.segment)+'</span>'+
      '<span class="pri-badge" style="background:'+pc.bg+';color:'+pc.fg+'">'+pc.label+'</span>'+
      '<span class="chev">&rsaquo;</span></div>';
  }).join('');
}
function coordsFor(e){
  const tokens=(e.city||'').split(/[\/,+\n]/).map(s=>s.trim().toLowerCase()).filter(Boolean);
  for(const t of tokens){if(CITY_COORDS[t])return CITY_COORDS[t];}
  return COUNTRY_COORDS[e.country]||null;
}
function buildExhMapLegend(){
  const order=['MUST_ATTEND','ATTEND','WATCH','AVOID'];
  document.getElementById('exh-map-legend').innerHTML=order.map(k=>{const pc=PRI_CFG[k];return '<span class="lg"><span class="sw" style="background:'+pc.fg+'"></span>'+pc.label+'</span>';}).join('');
}
function renderExhMap(){
  const rows=EXHIBITIONS.map((e,i)=>({e,i})).filter(o=>exhMatches(o.e));
  const groups={};
  rows.forEach(o=>{
    const c=coordsFor(o.e);if(!c)return;
    const key=c[0].toFixed(1)+','+c[1].toFixed(1);
    (groups[key]=groups[key]||{lat:c[0],lon:c[1],items:[]}).items.push(o);
  });
  const pins=Object.values(groups);
  document.getElementById('exh-map-meta').textContent=fmt(pins.length)+' locations \u00b7 '+fmt(rows.length)+' exhibitions';
  const order=['MUST_ATTEND','ATTEND','WATCH','AVOID'];
  const contHTML=CONTINENTS.map(c=>{const x=((c.lon+180)/360*100).toFixed(2),y=((90-c.lat)/180*100).toFixed(2);return '<span class="exh-continent" style="left:'+x+'%;top:'+y+'%">'+c.name+'</span>';}).join('');
  const seenC={};
  const countryHTML=pins.map(g=>{const cn=g.items[0].e.country;if(!cn||seenC[cn])return '';seenC[cn]=1;const cc=COUNTRY_COORDS[cn]||[g.lat,g.lon];const x=((cc[1]+180)/360*100).toFixed(2),y=(((90-cc[0])/180*100)-2.8).toFixed(2);return '<span class="exh-clabel" style="left:'+x+'%;top:'+y+'%">'+esc(cn)+'</span>';}).join('');
  const pinsEl=document.getElementById('exh-pins');
  pinsEl.innerHTML=contHTML+countryHTML+pins.map((g,gi)=>{
    const x=((g.lon+180)/360*100).toFixed(2),y=((90-g.lat)/180*100).toFixed(2);
    let dom='WATCH';
    for(const k of order){if(g.items.some(o=>o.e.priority===k)){dom=k;break;}}
    const pc=PRI_CFG[dom];
    const cntBadge=g.items.length>1?'<span class="cnt">'+g.items.length+'</span>':'';
    return '<button class="exh-pin" style="left:'+x+'%;top:'+y+'%;background:'+pc.fg+'" onclick="onPinClick(event,'+gi+')" title="'+esc(g.items.map(o=>o.e.name).join(', '))+'">'+cntBadge+'</button>';
  }).join('')+'<div class="exh-pin-pop" id="exh-pin-pop"></div>';
  window.__exhPinGroups=pins;
}
function onPinClick(ev,gi){
  ev.stopPropagation();
  const g=window.__exhPinGroups[gi];
  if(g.items.length===1){openExhDetail(g.items[0].i);return;}
  const pop=document.getElementById('exh-pin-pop');
  const btn=ev.currentTarget;
  pop.style.left=btn.style.left;pop.style.top=btn.style.top;
  pop.innerHTML='<div class="hd">'+g.items.length+' exhibitions here</div>'+g.items.map(o=>{
    const pc=PRI_CFG[o.e.priority]||PRI_CFG.WATCH;
    return '<div class="row" onclick="openExhDetail('+o.i+')"><span class="nm">'+esc(o.e.name)+'</span><span class="pri-badge" style="background:'+pc.bg+';color:'+pc.fg+'">'+pc.label+'</span></div>';
  }).join('');
  pop.style.display='block';
}
document.addEventListener('click',function(e){
  const pop=document.getElementById('exh-pin-pop');if(!pop)return;
  if(!e.target.closest('.exh-pin')&&!e.target.closest('.exh-pin-pop'))pop.style.display='none';
});
function buildExhPage(){renderWorldBase();buildExhFilters();buildExhRegions();buildExhMapLegend();renderExhList();renderExhMap();}
function openExhDetail(idx){
  const e=EXHIBITIONS[idx];
  document.getElementById('exdet-name').textContent=e.name;
  const pc=PRI_CFG[e.priority]||PRI_CFG.WATCH;
  let tags='<span class="dtag" style="background:#eaf1fd;color:#2a63cf">'+esc(e.segment)+'</span>';
  tags+='<span class="dtag" style="background:'+pc.bg+';color:'+pc.fg+'">'+pc.label+(e.stars?' '+'\u2605'.repeat(e.stars):'')+'</span>';
  if(e.status)tags+='<span class="dtag" style="background:#eff2f6;color:#5b6678">'+esc(e.status)+'</span>';
  document.getElementById('exdet-tags').innerHTML=tags;

  document.getElementById('exdc-overview').innerHTML=fieldBlock([
    ['Exhibition / Conference',e.name,{full:1,big:1}],
    ['Also Known As',e.nameNote,{full:1}],
    ['MIF Segment',e.segment],['Location',e.locationRaw],
    ['Dates (FY 2026-27)',e.dateRaw,{full:1}],
    ['Frequency / Cycle',e.frequency],['Next Edition',e.nextEdition],
    ['Exhibition Type',e.type],
    ['Venue',e.venue,{full:1}],
  ]);
  document.getElementById('exdc-fit').innerHTML=fieldBlock([
    ['MIF Priority',pc.label+(e.stars?' ('+e.stars+' star)':'')],
    ['Priority Rationale',e.priorityNote,{full:1}],
    ['MIF Product Fit',e.productFit,{full:1}],
    ['Key Customers Exhibiting',e.keyCustomers,{full:1}],
    ['Competitor Presence',e.competitors,{full:1}],
    ['Segment & Scale',e.scale,{full:1}],
  ]);
  document.getElementById('exdc-logistics').innerHTML=fieldBlock([
    ['Estimated Cost to Attend',e.cost],['Lead Time for Registration',e.leadTime],
    ['Registration Website',e.website],['Organiser Contact',e.organiser,{full:1}],
    ['Visitor Requirements',e.visitorReq,{full:1}],['Visa for Indians',e.visa,{full:1}],
  ],'three');
  document.getElementById('exdc-source').innerHTML=fieldBlock([
    ['Data Source / Verification',e.source],['Status',e.status],
  ],'three');

  showView('exhdetail');setNav('exh');
}
function closeExhDetail(){showView('exh');setNav('exh');}


// ── SIDEBAR (mobile) ──────────────────────────────────────────────
function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');}
function closeSidebar(){document.getElementById('sidebar').classList.remove('open');}
function logout(){if(confirm('Log out of MIF Market Research?'))location.reload();}
document.getElementById('pwd').addEventListener('keydown',e=>{if(e.key==='Enter')tryLogin();});
document.getElementById('usr').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('pwd').focus();});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){if(document.getElementById('view-detail').classList.contains('on'))closeDetail();else if(document.getElementById('view-exhdetail').classList.contains('on'))closeExhDetail();}});
</script>
</body>
</html>
"""

# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    parse_args()
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: File not found — {EXCEL_FILE}")
        print("Place this script in the same folder as the Excel file, or pass --input.")
        sys.exit(1)

    ts = datetime.now().strftime("%d %b %Y \u00b7 %I:%M %p")
    mif, comp, cols = read_companies()
    exh = read_exhibitions()
    news = parse_news(NEWS_FILE)
    html = generate(mif, comp, cols, ts, exh, news)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT_FILE) // 1024
    print(f"\n\u2713 Portal generated: {OUTPUT_FILE} ({size_kb} KB)")
    print(f"\u2713 Market login  ({PASSWORD_MIF}):  {len(mif)} companies")
    print(f"\u2713 Competitor login ({PASSWORD_COMP}): {len(comp)} companies")
    print(f"\u2713 Exhibitions:  {len(exh)}")
    print(f"\u2713 News editions: {len(news.get('editions', []))}")
    print(f"\u2713 Accent:    {ACCENT}")
    print(f"\u2713 Timestamp: {ts}")
    print(f"\nOpen '{OUTPUT_FILE}' in any browser. Fonts & company marks load from the web when online;")
    print("everything else works fully offline.")

if __name__ == '__main__':
    main()
