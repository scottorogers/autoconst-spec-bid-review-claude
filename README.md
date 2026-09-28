# Spec Bid-Review

**Read a construction spec book before you bid - automatically. Get back a filled-in spec-review form: every bid-killer, hidden cost, sole-source product, submittal and RFI, each cited to its exact section and page.**

Estimators already keep a spec-review form: the costly requirement, its section, its page. Filling it means reading 200-600 pages by hand. This tool reads the whole thing and fills the form for you in minutes, as an Excel workbook you can check line by line.

Built for GCs, subcontractors, estimators and PMs - anyone who has to read a spec before a bid or before building.

Two modes:

- **`--region uk` (default)** - UK tenders and subcontract enquiries: ITT / Instructions to Tenderers, JCT / NEC / FIDIC, NBS work sections (`Y61`, `SECTION Y63`, NBS clause numbers like `210`), Uniclass codes (`Ss_70_30_45`), £, LADs per week, Employer / Client / Contract Administrator, "or approved equivalent" / "similar approved", day-month dates.
- **`--region us`** - the original engine, unchanged: US CSI MasterFormat specs (5- and 6-digit sections), public-works special provisions (Roman-numeral sections, exhibits) and trade-only sections.

Same honesty rule as the [Drawing Takeoff](https://github.com/hamzaabduljabbar/autoConst-drawing-takeoff-claude) and [Spec Index](https://github.com/hamzaabduljabbar/autoConst-spec-index-claude) tools - every number is sourced, nothing is invented.

---

## UK mode - what it checks

| Tab | UK mode adds |
|---|---|
| **Bid-Killers** | Tender return date/time, mandatory site visit, tender queries deadline, tender validity, form of contract (JCT/NEC/FIDIC + edition), contract amendments / Z clauses, LADs / delay damages, performance bond, parent company guarantee, collateral warranties / third party rights, retention, payment terms, contract period / completion, insurance, design responsibility (CDP), CDM duties, accreditations / cards, variants / qualified tenders, tender addenda, required tender returns (Form of Tender, pricing schedule, CSA, BoQ) |
| **Hidden Costs** | "included in the Contract Sum / rates", "deemed to have allowed", "allow for", "at no additional cost to the Employer / Client / Main Contractor", "make good", contra charges |
| **Sole-Source** | "X or approved equivalent" / "similar approved", NBS `Manufacturer:` / `Product reference:` / `Substitution:` lines |
| **Costly Items** | `electrical` watchlist (LSZH, fire-resistant cable, SWA, ladder/tray/basket, busbar, earthing & bonding, SPDs, segregation, labelling, calcs, EIC/test certs, IST/witness testing, thermography, fibre/copper testing, raised floors, MEWPs, UPS/generator) and `universal_uk` (CDM, cards, accreditations, permits to work, RAMS, BIM/ISO 19650, out-of-hours, live environment, vetting, temporary works, waste, asbestos, attendances, O&Ms, defects period, training, Soft Landings, social value) |
| **RFI List** | Contract Administrator / Project Manager / Employer / Client roles, provisional & PC sums, "to the approval of" |

Also in UK mode:

- **Values come from the matched sentence only.** If a value has to come from the line after a heading, the row drops to MEDIUM and the Notes column says so. A match that runs across two sentences is ignored.
- **NBS section headers are recognised**, so a clause on the Y63 page is cited to Y63, not to the section before it.

**Status:** UK mode is checked against a synthetic tender (`tests/fixtures/uk_tender_synthetic.txt`) and was hand-tuned against one real 39-page UK construction-management trade-contract scope of works (not included - client document). It has not yet been run on a full ITT with contract particulars. Check every row against the page.

Every workbook now labels the page column **PDF page** (1 = first sheet of the PDF), not the printed page number.

---

## What this looks like in practice (US mode)

Run on a real **406-page City of Alameda public-works tender** (`--region us`). The Summary tab:

```
BID-KILLERS                     KEY VALUE                                   PAGE
Bid due date / opening          Wednesday, September 8, 2021; 2:01 p.m       6
Mandatory pre-bid / job-walk    Wednesday, August 18, 2021 at 11:00 am       6
Bid bond / bid security         10%  ("a bid shall not be considered unless")6
Bid validity period             90 days                                      7
Post-award bonds/insurance      5 business days, or forfeit bid guarantee    8
Required contractor license     "A" license or a "C" license                 13
Per-day penalties / fines       $250.00 per occurrence per day               19
Contract time / completion      two hundred fifty (250) consecutive working days  25
Addenda acknowledgement         not acknowledged = non-responsive            29
Liquidated damages              $2,500.00 per day                            34
Alternates                      ZERO add alternates                          40
Required bid forms / exhibits   14 exhibits (A–N), 7 attachments             2
Not found in this document:     Substitution request deadline

HIDDEN COSTS (12) - scope you pay for without a pay item
At Contractor's expense   Construction staking and layout ...                p29
No cost to Owner          ... may require night or weekend work, at no
                          additional cost to the City                        p51
...
```

18 of 19 bid-killer checks found, all correct. Plus 12 hidden costs, the spec's own submittal schedule, and a clause that makes every named product in the spec substitutable (p48).

---

## What's in the workbook

| Tab | What you get |
|---|---|
| **Summary** | Read this first. Every bid-killer with its value and page, top hidden costs, submittals by type, top costly items, top RFIs |
| **Bid-Killers** | 19 checks - bid due date, mandatory pre-bid, bid bond %, performance & payment bonds, post-award deadlines, bid validity, liquidated damages, fines, contract time, insurance, prevailing wage, contractor license, retainage, alternates, addenda, required forms. The key value (date, $, %, days) in its own column |
| **Hidden Costs** | "No separate payment", "incidental", "included in the unit price", "at Contractor's expense", "at no cost to the Owner", retesting at your cost |
| **Sole-Source** | Every named product classified: sole source / limited list / basis of design / or-equal - plus global "or equal" clauses |
| **Submittals** | Register with type and timing. Reads the spec's own submittal schedule table if it has one |
| **Costly Items** | Trade watchlists (concrete, MEP, finishes, universal) - price drivers only, with why each costs money and every page it appears on |
| **RFI List** | Only ambiguity that moves price (unforeseen conditions, "as directed", "to the satisfaction of", match existing...), ranked, with the RFI question drafted |
| **Method & Notes** | How it was produced, confidence rules, the text extractor used |

## How it stays honest

- Every row is a **real sentence from the document**, cited to section + page. Nothing paraphrased or invented.
- Not in the document? It says **NOT FOUND**. It doesn't guess.
- **HIGH** confidence = literal clause (still verify the number). **MEDIUM** = judgment call - read the clause.
- Tables of contents, drawing indexes, reference-standard lists, definitions and running headers are skipped automatically.
- Technical-only spec? The Summary tells you the bid terms are probably in the bid invitation / solicitation and to run that too.
- It flags. You decide.

---

## Setting it up (step by step)

### Step 1 - Install the prerequisites

**Windows** (PowerShell, one at a time, then close and reopen PowerShell):

```powershell
winget install --id Python.Python.3.12 -e
winget install --id oschwartz10612.Poppler -e
py -m pip install openpyxl
```

**Mac:**

```bash
brew install python@3.12 poppler
python3 -m pip install openpyxl
```

Poppler provides `pdftotext`, which pulls the text out of the PDF. Check it with `pdftotext -v`.

### Step 2 - Clone this repo with Claude Code

Open Claude Code and paste:

> *"Clone https://github.com/hamzaabduljabbar/autoConst-spec-bid-review-claude and install it as a skill."*

### Step 3 - Ask for a review

> *"Review this spec for bidding: C:\bids\itt.pdf"*

> *"Run a bid review on spec.pdf and itt.pdf - I'm the electrical sub."*

Claude runs the tool, opens the workbook and walks you through the bid-killers first, with page numbers.

Or run it yourself:

```bash
py scripts/bid_review.py path/to/spec.pdf [more.pdf ...] --trade all --out outputs/spec-review.xlsx
py scripts/bid_review.py path/to/spec.pdf --region us      # US CSI / bid documents
```

`--region`: `uk` (default) or `us`.
`--trade`: `all` (default), `electrical`, `concrete`, `mep`, `finishes`, or a comma list like `electrical,mep`. The site-wide list (`universal_uk` / `universal`) is always added.

**Tip:** pass the whole tender set. LADs, bonds, retention, the form of contract and amendments usually live in the ITT / Contract Particulars / subcontract enquiry, not the NBS spec.

---

## Tested on

| Document | Pages | Result |
|---|---|---|
| City of Alameda public-works tender (special provisions) | 406 | 18/19 bid-killers, all correct; 12 hidden costs; 9-item submittal schedule read |
| VA federal technical specification | 338 | Correctly reported "technical only"; 15 hidden costs; architectural cost drivers (Level 4 finish, VOC limits, moisture testing) |
| Municipal civil / electrical technical spec | 260 | Named products classified (limited lists, or-equal, single-source); short-circuit study, PVC-coated conduit |
| Short tender (NWIT) | 39 | Post-award 7-day insurance/bond deadline |
| Division 03 concrete section | 80 | Correctly no bid-killers; or-equal products merged across sections |

`tests/test_uk.py` checks UK mode against the synthetic fixture. It needs no PDFs and no pdftotext: `py tests/test_uk.py`.

`tests/regression.py` re-runs these (in `--region us`) under every `pdftotext` build on the machine (xpdf and Poppler output differs - the tool normalises it) and checks the hand-verified values. Put the PDFs in the folder above the repo, or set `SPEC_DIR`.

---

## Folder layout

```
autoConst-spec-bid-review-claude/
  README.md
  SKILL.md                 <- Claude Code skill file
  LICENSE
  requirements.txt
  config/watchlists.json   <- trade cost-driver watchlists (edit to suit your scope)
  scripts/bid_review.py    <- the engine
  tests/test_uk.py         <- UK-mode checks on a synthetic tender (no PDFs needed)
  tests/fixtures/          <- the synthetic UK tender text
  tests/regression.py      <- regression checks against real US specs
  examples/                <- a real output workbook (Alameda tender)
  work/                    <- extracted text (auto-generated)
  outputs/                 <- your workbooks (auto-generated)
```

---

## Troubleshooting

**"pdftotext was not found"** - Install Poppler (Step 1) and reopen your terminal, or set `PDFTOTEXT` to the full path of `pdftotext.exe`.

**"no usable text layer ... scanned PDF"** - The PDF is a scan. OCR it first (Adobe Acrobat "Recognize Text", or `ocrmypdf`) and run again.

**"Cannot write ... probably open in Excel"** - Close the workbook, or pass a different `--out`.

**Few or no bid-killers** - You probably ran a technical-only spec. Add the ITT / Instructions to Tenderers / Contract Particulars (UK) or bid invitation / general conditions (US) to the same command.

**US spec giving odd results** - Add `--region us`. UK is the default.

**A costly item that doesn't apply to you** - Watchlists are keyword-based. Edit `config/watchlists.json`, or run with `--trade` for your trade only.

---

## Limits

Text-layer PDFs only. Tables are read as laid out, so check due dates that come from tables. Drawings are not read. Costly Items and the RFI List need an estimator's judgment.

---

Built by [Hamza Jabbar](https://hamzajabbar.online) - [AutoConst](https://hamzajabbar.online).
