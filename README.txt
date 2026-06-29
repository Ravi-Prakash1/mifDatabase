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
