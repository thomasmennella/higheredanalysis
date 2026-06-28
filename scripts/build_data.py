#!/usr/bin/env python3
"""
Canary build-layer data fetcher.

Pulls the financial / lagging signals that the static page CANNOT fetch
client-side (ProPublica is CORS-blocked; ED files have no live API) and writes
data.json for the page to read. Runs in GitHub Actions on a schedule.

Sources
  - ProPublica Nonprofit Explorer (IRS 990): operating margin, deficit years   [reliable]
  - ED Financial Responsibility Composite Scores (annual .xlsx): composite score [optional]
  - ED Heightened Cash Monitoring (quarterly .xlsx): HCM status                  [optional]

The ProPublica section is the dependable core. The two ED sections are OPTIONAL:
set ED_COMPOSITE_URL / ED_HCM_URL to the current ED file URLs to enable them.
If left as None, those rows simply stay manual in the app — nothing breaks.

Output: data.json  ->  { "generated": <iso>, "data": { "<unitid>": { ... } } }
Each metric is { "v": <value>, "yr": <label> }, matching what the page expects.
"""

import os
import sys
import json
import time
import datetime
import requests

SCORECARD_KEY = os.environ.get("SCORECARD_API_KEY", "DEMO_KEY")
SC = "https://api.data.gov/ed/collegescorecard/v1/schools"
PP = "https://projects.propublica.org/nonprofits/api/v2"

# studentaid.gov throttles non-browser clients, so send a browser-like UA.
ED_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "*/*",
}

# WNE + curated peer roster. MUST match PEER_CHOICES in index.html so that any
# peer a user can pick has pre-computed data. Add names here and in the page.
INSTITUTIONS = [
    "Western New England University",
    "American International College", "Springfield College", "Bay Path University",
    "Assumption University", "Merrimack College", "Nichols College",
    "University of New Haven", "Roger Williams University",
    "Wentworth Institute of Technology", "Quinnipiac University",
    "Sacred Heart University", "Stonehill College", "Suffolk University",
    "Salve Regina University", "Endicott College", "Curry College",
    "Lasell University", "Emmanuel College", "Elms College", "Clark University",
    "Worcester Polytechnic Institute", "Bryant University",
]

# OPTIONAL — set to the current ED file URLs to enable Composite Score + HCM.
# Find them at the Federal Student Aid data center (search "Financial
# Responsibility Composite Scores" and "Heightened Cash Monitoring"). Leave None
# to skip. The parser keys on OPEID8 (pulled from Scorecard) and matches names.
ED_COMPOSITE_URL = "https://studentaid.gov/sites/default/files/ay-22-23-composite-scores.xls"
ED_HCM_URL = "https://studentaid.gov/sites/default/files/Schools-on-hcm-mar-2026.xlsx"


def get(url, **kw):
    kw.setdefault("timeout", 40)
    r = requests.get(url, **kw)
    r.raise_for_status()
    return r


def resolve_scorecard(name):
    """name -> (unitid, opeid8, canonical_name) or (None, None, None)."""
    try:
        r = get(SC, params={
            "api_key": SCORECARD_KEY,
            "school.name": name,
            "fields": "id,ope8_id,school.name,school.city,school.state",
            "per_page": 3,
        })
        res = r.json().get("results", [])
        if res:
            m = res[0]
            return str(m["id"]), str(m.get("ope8_id") or ""), m["school.name"]
    except Exception as e:
        print(f"  scorecard resolve failed for {name}: {e}")
    return None, None, None


def propublica_finance(name):
    """Best-matching 990 filer -> dict(opm, defy, n, year, ein) or None."""
    try:
        r = get(f"{PP}/search.json", params={"q": name})
        orgs = r.json().get("organizations", [])
        if not orgs:
            return None
        # Prefer an org whose name contains a distinctive token of the school
        # and that actually has filings; fall back to the first result.
        token = name.split()[0].lower()
        ranked = sorted(
            orgs,
            key=lambda o: (token in (o.get("name", "").lower()), ),
            reverse=True,
        )
        for o in ranked[:3]:
            ein = o["ein"]
            data = get(f"{PP}/organizations/{ein}.json").json()
            filings = [f for f in data.get("filings_with_data", [])
                       if f.get("totrevenue") is not None]
            if not filings:
                continue
            filings.sort(key=lambda f: f.get("tax_prd_yr", 0), reverse=True)
            latest, last3 = filings[0], filings[:3]
            defy = sum(1 for f in last3
                       if (f["totrevenue"] - (f.get("totfuncexpns") or 0)) < 0)
            rev = latest["totrevenue"]
            opm = round((rev - (latest.get("totfuncexpns") or 0)) / rev * 100, 1) if rev else None
            return {"opm": opm, "defy": defy, "n": len(last3),
                    "year": str(latest.get("tax_prd_yr", "")), "ein": ein,
                    "org": o.get("name", "")}
    except Exception as e:
        print(f"  propublica failed for {name}: {e}")
    return None


# ----- OPTIONAL ED file parsing (enabled only when URLs are set) --------------
def load_ed_table(src):
    """Load an ED .xls/.xlsx (URL or local path) into a list of dict rows. [] on failure."""
    if not src:
        return []
    try:
        import io
        import pandas as pd
        if str(src).startswith("http"):
            r = requests.get(src, headers=ED_HEADERS, timeout=(15, 180))
            r.raise_for_status()
            df = pd.read_excel(io.BytesIO(r.content))   # engine auto: openpyxl (.xlsx) / xlrd (.xls)
        else:
            df = pd.read_excel(src)                      # committed local file fallback
        df.columns = [str(c).strip().lower() for c in df.columns]
        rows = df.to_dict("records")
        print(f"  ED file loaded: {len(rows)} rows, columns: {list(df.columns)}")
        return rows
    except Exception as e:
        print(f"  ED file load failed ({src}): {e}")
        return []


def find_opeid(rows, opeid8, name):
    """Match an ED row by OPEID (preferred) or fuzzy name. Returns the row or None."""
    if not rows:
        return None
    op6 = (opeid8 or "")[:6].lstrip("0")
    for row in rows:
        for k, v in row.items():
            if "opeid" in k or k == "opeid":
                vv = str(v or "").lstrip("0")
                if op6 and (vv == op6 or vv == (opeid8 or "").lstrip("0")):
                    return row
    nl = name.lower().split()[0]
    for row in rows:
        for k, v in row.items():
            if ("school" in k or "name" in k or "institution" in k) and v and nl in str(v).lower():
                return row
    return None


def main():
    out = {"generated": datetime.datetime.now(datetime.timezone.utc).isoformat(), "data": {}}

    comp_rows = load_ed_table(ED_COMPOSITE_URL)
    hcm_rows = load_ed_table(ED_HCM_URL)

    for name in INSTITUTIONS:
        uid, opeid8, canon = resolve_scorecard(name)
        if not uid:
            print(f"skip (no unitid): {name}")
            continue
        rec = {}

        pf = propublica_finance(name)
        if pf:
            if pf["n"] >= 2:
                rec["defy"] = {"v": pf["defy"], "yr": f"990 {pf['year']}"}
            if pf["opm"] is not None:
                rec["opm"] = {"v": pf["opm"], "yr": f"990 {pf['year']}"}
            rec["_src"] = f"990 {pf['year']} ({pf['org']})"
            print(f"{name} [{uid}] -> opm={pf['opm']} defy={pf['defy']} ({pf['year']})")
        else:
            print(f"{name} [{uid}] -> no 990 (public institution?)")

        # Optional ED rows (only if files were provided)
        cr = find_opeid(comp_rows, opeid8, name)
        if cr:
            for k, v in cr.items():
                if "composite" in k and isinstance(v, (int, float)):
                    rec["comp"] = {"v": round(float(v), 2), "yr": "ED"}
        hr = find_opeid(hcm_rows, opeid8, name)
        if hr:
            txt = " ".join(str(v).lower() for v in hr.values())
            rec["hcm"] = {"v": "HCM2" if "hcm2" in txt or "cash monitoring 2" in txt
                          else "HCM1", "yr": "ED"}

        out["data"][uid] = rec
        time.sleep(0.3)

    with open("data.json", "w") as f:
        json.dump(out, f, indent=2)
    print(f"\nWrote data.json: {len(out['data'])} institutions, "
          f"{sum(1 for v in out['data'].values() if v.get('opm'))} with 990 margin.")


if __name__ == "__main__":
    main()
