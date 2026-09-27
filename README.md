# Spec Bid-Review

**Read a construction spec book before you bid - automatically. Get back a filled-in spec-review form: every bid-killer, hidden cost, sole-source product, submittal and RFI, each cited to its exact section and page.**

Estimators already keep a spec-review form: the costly requirement, its section, its page. Filling it means reading 200-600 pages by hand. This tool reads the whole thing and fills the form for you in minutes, as an Excel workbook you can check line by line.

Built for GCs, subcontractors, estimators and PMs - anyone who has to read a spec before a bid or before building. Works on US CSI MasterFormat specs (5- and 6-digit sections), public-works special provisions (Roman-numeral sections, exhibits) and trade-only sections.

Same honesty rule as the [Drawing Takeoff](https://github.com/hamzaabduljabbar/autoConst-drawing-takeoff-claude) and [Spec Index](https://github.com/hamzaabduljabbar/autoConst-spec-index-claude) tools - every number is sourced, nothing is invented.

---

## What this looks like in practice

Run on a real **406-page City of Alameda public-works tender**. The Summary tab:

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

> *"Review this spec for bidding: C:\bids\city-tender.pdf"*

> *"Run a bid review on spec.pdf and invitation-to-bid.pdf - I'm the electrical sub."*

Claude runs the tool, opens the workbook and walks you through the bid-killers first, with page numbers.

Or run it yourself:

```bash
py scripts/bid_review.py path/to/spec.pdf [more.pdf ...] --trade all --out outputs/spec-review.xlsx
```

`--trade`: `all` (default), `concrete`, `mep`, `finishes`, or a comma list like `mep,finishes`.

**Tip:** pass the whole bid set. Bid bonds, liquidated damages and wage rates usually live in the bid invitation / general conditions, not the technical specification.

---

## Tested on

| Document | Pages | Result |
|---|---|---|
| City of Alameda public-works tender (special provisions) | 406 | 18/19 bid-killers, all correct; 12 hidden costs; 9-item submittal schedule read |
| VA federal technical specification | 338 | Correctly reported "technical only"; 15 hidden costs; architectural cost drivers (Level 4 finish, VOC limits, moisture testing) |
| Municipal civil / electrical technical spec | 260 | Named products classified (limited lists, or-equal, single-source); short-circuit study, PVC-coated conduit |
| Short tender (NWIT) | 39 | Post-award 7-day insurance/bond deadline |
| Division 03 concrete section | 80 | Correctly no bid-killers; or-equal products merged across sections |

`tests/regression.py` re-runs these under every `pdftotext` build on the machine (xpdf and Poppler output differs - the tool normalises it) and checks the hand-verified values. Put the PDFs in the folder above the repo, or set `SPEC_DIR`.

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
  tests/regression.py      <- regression checks against real specs
  examples/                <- a real output workbook (Alameda tender)
  work/                    <- extracted text (auto-generated)
  outputs/                 <- your workbooks (auto-generated)
```

---

## Troubleshooting

**"pdftotext was not found"** - Install Poppler (Step 1) and reopen your terminal, or set `PDFTOTEXT` to the full path of `pdftotext.exe`.

**"no usable text layer ... scanned PDF"** - The PDF is a scan. OCR it first (Adobe Acrobat "Recognize Text", or `ocrmypdf`) and run again.

**"Cannot write ... probably open in Excel"** - Close the workbook, or pass a different `--out`.

**Few or no bid-killers** - You probably ran a technical-only spec. Add the bid invitation / instructions to bidders / general conditions PDF to the same command.

**A costly item that doesn't apply to you** - Watchlists are keyword-based. Edit `config/watchlists.json`, or run with `--trade` for your trade only.

---

## Limits

Text-layer PDFs only. Tables are read as laid out, so check due dates that come from tables. Drawings are not read. Costly Items and the RFI List need an estimator's judgment.

---

Built by [Hamza Jabbar](https://hamzajabbar.online) - [AutoConst](https://hamzajabbar.online).
