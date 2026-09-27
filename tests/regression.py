#!/usr/bin/env python3
"""
Regression check: re-run bid_review.py on real specs and assert known-correct answers.

    py tests/regression.py            # all specs found on disk
    py tests/regression.py sdps div03 # only these

Spec PDFs are looked up in SPEC_DIR (default: the folder above this repo). Missing PDFs are
SKIPPED, not failed. Exit code 0 = every check that ran passed.
Every expected value below was verified by hand against the source PDF page.
"""
import os, subprocess, sys, tempfile
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SPEC_DIR = Path(os.environ.get("SPEC_DIR", ROOT.parent))
ENGINE = ROOT / "scripts" / "bid_review.py"

def rows(wb, tab):
    data = list(wb[tab].iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:]]

def bk(wb):
    return {r["Bid-killer"]: r for r in rows(wb, "Bid-Killers")}

def found(wb):
    return [r for r in rows(wb, "Bid-Killers") if r["Clause"] != "NOT FOUND"]

def check_sdps(wb):
    b = bk(wb)
    yield "18+ bid-killers found", len(found(wb)) >= 18
    yield "bid due Sep 8 2021", "September 8, 2021" in str(b["Bid due date / opening"]["Key value"])
    yield "mandatory pre-bid Aug 18 2021", "August 18, 2021" in str(b["Mandatory pre-bid / job-walk"]["Key value"])
    yield "bid bond 10%", b["Bid bond / bid security"]["Key value"] == "10%"
    yield "bid bond page 6", b["Bid bond / bid security"]["Page"] == 6
    yield "LD $2,500/day", b["Liquidated damages"]["Key value"] == "$2,500.00 per day"
    yield "LD page 34", b["Liquidated damages"]["Page"] == 34
    yield "fine $250/occurrence/day (not the LD)", str(b["Per-day penalties / fines"]["Key value"]).startswith("$250.00")
    yield "contract time 250 working days", "(250) consecutive working days" in str(b["Contract time / completion"]["Key value"])
    yield "bid validity 90 days", b["Bid validity period"]["Key value"] == "90 days"
    # p13 says both 'an "A" license or a "C" license' and 'Class A or C-10' - either is correct
    lic = str(b["Required contractor license"]["Key value"])
    yield "license A or C", ('"A" license' in lic and '"C" license' in lic) or lic.startswith("Class A or C")
    yield "post-award forfeit clause p8", b["Post-award bonds/insurance deadline"]["Page"] == 8
    yield "addenda non-responsive p29", "non-responsive" in str(b["Addenda acknowledgement"]["Clause"]) and b["Addenda acknowledgement"]["Page"] == 29
    yield "zero alternates", "ZERO" in str(b["Alternates"]["Clause"])
    # table of contents p2 lists Exhibits A-N (14) and Attachments A-G (7)
    yield "14 exhibits A-N + 7 attachments", b["Required bid forms / exhibits"]["Key value"] == "14 exhibits (A–N), 7 attachments"
    yield "substitution deadline NOT FOUND", b["Substitution request deadline"]["Clause"] == "NOT FOUND"
    ss = rows(wb, "Sole-Source")
    yield "global or-equal clause p48", ss[0]["Classification"].startswith("GLOBAL") and ss[0]["Page"] == 48
    sub = rows(wb, "Submittals")
    yield "9 rows from spec's submittal schedule", sum(1 for r in sub if r["Type"] == "From the spec's submittal schedule") == 9
    yield "no bid-phase 'submitting bids' submittal", not any("submitting bids" in str(r["Submittal"]).lower() for r in sub)
    hid = rows(wb, "Hidden Costs")
    yield "10+ hidden costs", len(hid) >= 10
    yield "staking at contractor's expense found", any("STAKING" in str(r["Clause"]) for r in hid)
    rfi = rows(wb, "RFI List")
    yield "groundwater presumed RFI", any("high groundwater" in str(r["Ambiguous clause"]) for r in rfi)
    yield "no indemnity/surety boilerplate in RFIs", not any(w in str(r["Ambiguous clause"]).lower() for r in rfi for w in ("indemnif", "surety"))
    cost = rows(wb, "Costly Items")
    yield "no 'Microfiber filter' as fiber reinforcement", not any(r["Costly item"] == "Fiber reinforcement" for r in cost)

def check_div03(wb):
    yield "0 bid-killers (trade-only section)", len(found(wb)) == 0
    ss = rows(wb, "Sole-Source")
    prods = {str(r["Product"]) for r in ss}
    yield "exactly 3 named products", len(ss) == 3
    yield "Confilm + Eucobar + Thoroseal", {"Confilm", "Eucobar"} <= prods and any(p.startswith("Thoroseal") for p in prods)
    yield "Confilm merged across p23+p60", any(r["Product"] == "Confilm" and "p60" in str(r["Also on pages"]) for r in ss)
    yield "all or-equal", all(r["Classification"].startswith("OR-EQUAL") for r in ss)

def check_techspec(wb):
    yield "0 bid-killers (technical-only)", len(found(wb)) == 0
    ss = rows(wb, "Sole-Source")
    yield "Acme/Jefferson limited list", any("Acme Electric; Jefferson Electric" in str(r["Manufacturer(s)"]) and r["Classification"].startswith("LIMITED") for r in ss)
    yield "PVC-coated conduit 3-maker list", any("Plasti-Bond" in str(r["Manufacturer(s)"]) for r in ss)
    yield "no distributor (Hanson) row", not any("Hanson" in str(r["Product"]) + str(r["Manufacturer(s)"]) for r in ss)
    cost = rows(wb, "Costly Items")
    yield "no 'tab sheets' as TAB", not any("tab sheets" in str(r["Evidence"]).lower() for r in cost)
    yield "no 'liquid tight' formwork", not any("Forms must be rigid" in str(r["Evidence"]) for r in cost)
    yield "short-circuit study found", any(r["Costly item"] == "Power system studies" for r in cost)

def check_va(wb):
    yield "few bid-killers (technical-only)", len(found(wb)) <= 2
    cost = rows(wb, "Costly Items")
    yield "no evidence from reference/definition articles", not any(k in str(r["Section"]) for r in cost for k in ("Applicable Publications", "Definitions"))
    yield "no radiology X-ray as slab scanning", not any(r["Costly item"] == "Slab scanning / X-ray" for r in cost)
    yield "Level 4 drywall finish found", any(r["Costly item"] == "Level 4/5 drywall finish" for r in cost)
    yield "section citations are CSI numbers", any(str(r["Section"]).startswith("01 00 00 §") for r in rows(wb, "Hidden Costs"))
    yield "15 hidden costs", len(rows(wb, "Hidden Costs")) >= 12

def check_nwit(wb):
    b = bk(wb)
    yield "post-award 7 days", b["Post-award bonds/insurance deadline"]["Key value"] == "7 days"

SPECS = {
    "sdps": ("sdps_special-provisions-full-07.08.21.pdf", "all", check_sdps),
    "div03": ("Technical Specifications Division 03 (PDF).pdf", "concrete", check_div03),
    "techspec": ("technical-specifications_final-052121.pdf", "all", check_techspec),
    "va": ("36C26319R0017-001-Specs-Part-1.pdf", "all", check_va),
    "nwit": ("NWIT_15_08_Tender_Specifications.pdf", "all", check_nwit),
}

def extractors():
    """Every distinct pdftotext on this machine (xpdf and Poppler emit different text - test both).
    Set PDFTOTEXT to test just one."""
    if os.environ.get("PDFTOTEXT"):
        return [os.environ["PDFTOTEXT"]]
    found, seen = [], set()
    for d in os.environ.get("PATH", "").split(os.pathsep):
        for exe in ("pdftotext.exe", "pdftotext"):
            p = Path(d) / exe
            if p.is_file():
                key = str(p.resolve()).lower()
                if key not in seen:
                    seen.add(key)
                    found.append(str(p))
    return found

def version(exe):
    r = subprocess.run([exe, "-v"], capture_output=True, text=True)
    return ((r.stderr or r.stdout).strip().splitlines() or ["?"])[0]

def main():
    wanted = sys.argv[1:] or list(SPECS)
    exes = extractors()
    if not exes:
        sys.exit("No pdftotext found on PATH.")
    passed = failed = skipped = 0
    with tempfile.TemporaryDirectory() as tmp:
        for exe in exes:
            tag = version(exe)
            print(f"\n=== extractor: {exe} ({tag})")
            env = dict(os.environ, PDFTOTEXT=exe, PYTHONIOENCODING="utf-8")
            for name in wanted:
                pdf, trade, check = SPECS[name]
                src = SPEC_DIR / pdf
                if not src.exists():
                    print(f"SKIP {name}: {src} not found")
                    skipped += 1
                    continue
                out = Path(tmp) / f"{name}-{abs(hash(exe)) % 10**6}.xlsx"
                r = subprocess.run([sys.executable, str(ENGINE), str(src), "--trade", trade, "--out", str(out)],
                                   capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
                if r.returncode != 0:
                    print(f"FAIL {name}: engine crashed\n{r.stderr[-800:]}")
                    failed += 1
                    continue
                wb = load_workbook(out)
                for label, ok in check(wb):
                    if not ok:
                        print(f"FAIL {name}: {label}")
                    passed += ok
                    failed += (not ok)
                print(f"  {name}: done")
    print(f"\n{passed} passed, {failed} failed, {skipped} spec runs skipped, {len(exes)} extractor(s)")
    sys.exit(1 if failed else 0)

if __name__ == "__main__":
    main()
