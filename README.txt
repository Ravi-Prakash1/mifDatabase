=========================================================
   MIF MARKET INTELLIGENCE PORTAL
   Quick Start Guide
=========================================================
Mother India Forming Pvt. Ltd. | Confidential | June 2026
=========================================================


WHAT IS IN THIS FOLDER?
─────────────────────────────────────────────
  MIF_Intelligence_Portal.html   ← The portal (this is what you open)
  generate_dashboard.py          ← Script to regenerate when Excel updates
  MIF_Master_Database_...xlsx    ← Source Excel (keep this in same folder)
  README.txt                     ← This file


HOW TO OPEN THE PORTAL
─────────────────────────────────────────────
1. Double-click MIF_Intelligence_Portal.html
2. It opens in your default browser (Chrome, Edge, Firefox, Safari)
3. Enter the password when asked

Default Password:  MIF2026

Works fully offline. No internet connection needed.
Works on Windows PC, Mac, iPhone, Android — any modern browser.


HOW TO USE THE PORTAL
─────────────────────────────────────────────
LEFT SIDEBAR — NAVIGATION:
  Dashboard         → Summary stats and charts
  All Companies     → Full list of all 997 companies
  Customers         → Companies where MIF already has a relationship
  Competitors       → T1 Captive / T1 OEM companies
  BD Prospects      → Pure Prospects and T2 Conversion targets
  Phase 1 / Phase 2 → Filter by research phase
  By Segment        → Click any segment name to see only that segment's list

SEARCH & FILTERS (on the list view):
  Search box        → Searches by name, CIN, products, BD notes,
                       procurement contacts, plant locations, and more
  Region filter     → SOUTH / NORTH / WEST / EAST / MULTI
  Tier filter       → T1 Captive, T2 Conversion, Pure Prospect, etc.
  Fit filter        → H (High), M (Medium), L (Low) strategic fit
  Confidence filter → High / Medium / Low data confidence

COMPANY DETAIL VIEW:
  Click any company card to open its full profile on the right panel.
  Shows all 57 fields organized into 8 sections:
    • Company Identity (CIN, GSTIN, Legal Form, Year, Parent Group)
    • Financials (Revenue, EBITDA, Growth, Capex)
    • Operations (Employees, Plants, Locations, HQ)
    • Key Contacts (Director, Procurement Head, Sales Head)
    • Products & Steel Buying (Products, HSN, Steel types, Volume, Tonnage)
    • MIF BD Intelligence (Strategic Fit, BD Notes, Key Products for MIF)
    • Competitor Analysis (Tier, Switching Difficulty, Vendor Reg. Status)
    • Data Quality (Confidence, Last Verified, Source links)

  Press Escape or click the X to close the detail panel.

MOBILE:
  Tap the ☰ (menu icon) in the top left to open the sidebar.
  Company detail opens as full-screen on mobile.


HOW TO UPDATE THE PORTAL AFTER EXCEL CHANGES
─────────────────────────────────────────────
Step 1: Edit and save MIF_Master_Database_...xlsx as usual
Step 2: Keep the Excel file in the same folder as generate_dashboard.py
Step 3: Open Command Prompt / Terminal in that folder
Step 4: Run:  python generate_dashboard.py
Step 5: A new MIF_Intelligence_Portal.html is created (takes ~5 seconds)
Step 6: Open the new HTML file — it shows updated data with today's timestamp

The "Data as of" date shown in the portal header updates automatically.

If you want to change the password:
    python generate_dashboard.py --password NewPassword123


AUTOMATED DAILY UPDATES (NEW)
─────────────────────────────────────────────
The portal can now refresh itself — using FREE tools only, no paid API. A
scheduled GitHub Action (.github/workflows/auto-update.yml) runs every morning
(~08:00 IST) and:

  1. Pulls fresh market news from free news feeds (Google News + industry
     RSS) and the USD/INR rate, then drafts the items that matter to MIF into
     MIF_News_Intelligence.xlsx.
  2. Fills MISSING prospect/exhibition fields from free public data
     (Wikidata/Wikipedia + the GLEIF company register) — only empty cells are
     touched, each filled value is flagged "Estimated", carries a source link,
     and is recorded in the "Fill Audit Log" sheet.
  3. Regenerates index.html and opens a PULL REQUEST for you to review.

Nothing goes live automatically. You review the PR, edit anything that looks
off, and merge — merging deploys the portal to GitHub Pages as before.

HOW THE NEWS IS SUMMARISED (pick one — all free):
  * Groq (recommended): free API key, no credit card, and it does not train on
    your data. Better AI summaries and a "Why it matters to MIF" line.
      1. Get a free key at https://console.groq.com
      2. GitHub: repo Settings -> Secrets and variables -> Actions ->
         New repository secret, name it exactly:  GROQ_API_KEY
  * Gemini: alternative free key from https://aistudio.google.com — set secret
    GEMINI_API_KEY instead. (Note: Google may use free-tier data to improve
    their models.)
  * Nothing at all: if no key is set, the news is auto-drafted straight from the
    headlines with keyword rules, and you fill the "Why it matters" notes in the
    review PR. Fully private, zero setup.

NEW-COMPANY DISCOVERY:
  The daily job also looks for companies NOT yet in your database — from the
  market news, from exhibition exhibitor lists, and from Wikidata by segment.
  A candidate is only added if it (a) is genuinely new (not already in the
  sheet), (b) is confirmed a real company in a free source (Wikidata/GLEIF),
  and (c) clearly belongs to one of MIF's 38 end-use segments (this keeps out
  steel mills and other off-target names). New rows are appended to the Master
  Database with identity fields filled and everything else blank, flagged
  "Needs Verification" — NOTHING is auto-marked Verified. Confirm each new
  company (and fill its financials/contacts) before any outreach.

  Enrichment is prioritised by BD value: existing customers first, then
  Tier 1 / T1 Captive-OEM, Active Prospects, Tier 2 conversion, then the rest.

COVERAGE NOTE:
  Free company data mainly covers larger, well-known firms. Small/obscure
  prospects usually won't be found and their cells simply stay empty for you to
  fill by hand — the tool never guesses. Per-run caps mean a full backfill fills
  in gradually over many days.

RUN THE AUTOMATION LOCALLY (optional):
  pip install -r requirements.txt
  python scripts/fetch_news.py --dry-run          # preview news, no writes
  python scripts/enrich_data.py --dry-run --limit 5
  # For AI summaries locally:  export GROQ_API_KEY=gsk_...  (else rule-based)
  # then drop --dry-run to actually write


SHARING WITH OTHERS
─────────────────────────────────────────────
Share ONLY the MIF_Intelligence_Portal.html file.
Do NOT share the Excel source or the Python script externally.

The HTML file contains all data embedded inside it.
Recipients need: just the .html file + a modern browser + the password.

Recommended: ZIP the file and send via email or WhatsApp.
File size: ~2.6 MB (compresses to ~0.5 MB in ZIP)


REQUIREMENTS (to regenerate only — not to view)
─────────────────────────────────────────────
  Python 3.7 or higher
  pandas library: pip install pandas openpyxl

  To view the portal: any browser, no Python needed.


TECHNICAL NOTES
─────────────────────────────────────────────
  Password protection: SHA-256 hash verified in browser
  Data: 997 companies × 57 fields from Combined Master sheet
  Password stored as hash only — actual password is never in the file
  No data is transmitted anywhere — fully offline
  No cookies, no tracking, no external scripts


─────────────────────────────────────────────
Prepared by: Ravi, BI & R&D Lead, MIF
Tool generated: June 2026
Confidential — MIF Internal Use Only
─────────────────────────────────────────────
