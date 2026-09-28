# UK Spec Bid-Review (Claude skill)

**Read a UK tender, spec or scope of works before you price or sign it. Get back a filled-in spec-review form: every bid-killer, hidden cost, sole-source product, timed submittal and RFI, each cited to its section and PDF page.**

Built for UK PMs, estimators, QSs and specialist subcontractors. Reads:

- ITT / Instructions to Tenderers, subcontract enquiries
- JCT / NEC / FIDIC wording, contract particulars and amendments
- NBS work sections (`Y61`, `SECTION Y63`) and clauses (`210 CABLE LADDER`), Uniclass codes (`Ss_70_30_45`)
- Construction-management trade-contract scopes (`Section 6 – Valuation Dates`, `1.5 Testing and Commissioning`)
- £ amounts, day-month dates, LADs per week, Employer / Client / Contract Administrator / Construction Manager

Every row is a real sentence from the document. Not in the document? It says **NOT FOUND** - it never guesses.

---

## What's in the workbook

| Tab | What you get |
|---|---|
| **Summary** | Read first. Bid-killers with value and page, top hidden costs, submittals by type, cost drivers, top RFIs |
| **Bid-Killers** | Tender return date/time, mandatory site visit, queries deadline, tender validity, form of contract / pricing basis (JCT/NEC/FIDIC, lump sum fixed price), amendments / Z clauses, LADs / delay damages, performance bond, PCG, collateral warranties, retention, payment terms, contract period, defects period / warranty, insurance, CDP, CDM, accreditations / cards, variants / qualified tenders, addenda, tender returns. Key value (date, £, %, weeks) in its own column |
| **Hidden Costs** | "Deemed included", "allow for", "at the (Trade) Contractor's expense", scope not shown on the drawings deemed yours, no adjustment to the Contract Sum, chargeable to / contra charges, at no extra cost, make good, retest costs |
| **Sole-Source** | "X or approved equivalent" / "similar approved", NBS `Manufacturer:` / `Product reference:` / `Substitution:` lines, manufacturer lists, basis of design |
| **Submittals** | Register with type and timing - including timed deliverables ("ITP 8 weeks prior to commencement", "programme within 1 month of appointment") |
| **Costly Items** | Watchlists: `electrical` (LSZH, fire-resistant cable, SWA, ladder/tray/basket, busbar, earthing, SPDs, segregation, labelling, IST/witness testing, fibre testing, raised floors, MEWPs, UPS...), `mep`, `concrete`, `finishes`, plus `universal` (CDM, cards, permits to work, RAMS, BIM, out-of-hours, live environment, vetting, attendances, O&Ms, defects period, 24-hour call-out, branded PPE, mandated software, no parking, s106...) |
| **RFI List** | Only ambiguity that moves price or programme, ranked, RFI/TQ question drafted |
| **Method & Notes** | How it was produced, confidence rules, extractor used |

Page numbers are **PDF pages** (1 = first sheet of the PDF), not the printed footer number.

## How it stays honest

- Every row is a sentence from the document with section + PDF page. Nothing paraphrased or invented.
- Values are read from the matched sentence only. If a value had to come from the line after a heading, the row drops to **MEDIUM** and says so.
- **HIGH** = literal clause (still verify the number). **MEDIUM** = judgment call - read the clause.
- Contents pages, reference-standard lists, definitions and running headers are skipped.
- It flags. You decide.

---

## Set up

### 1. Prerequisites

**Windows** (PowerShell, then close and reopen it):

```powershell
winget install --id Python.Python.3.12 -e
winget install --id oschwartz10612.Poppler -e
py -m pip install openpyxl
```

**Mac:** `brew install python@3.12 poppler && python3 -m pip install openpyxl`

Check with `pdftotext -v`.

### 2. Install the skill

- **Claude (app / claude.ai):** open the `uk-spec-bid-review.skill` file and click **Save skill** (if your organisation allows skills).
- **Claude Code:** copy this folder to `~/.claude/skills/uk-spec-bid-review/`.

### 3. Use it

> *"Review this tender before I price it: C:\bids\job123\ITT.pdf and the NBS spec"*

> *"What am I signing up to in this scope of works? I'm the electrical sub."*

> *"Pull the LADs, retention and anything deemed included from these docs and draft the RFIs."*

Or run it yourself:

```bash
python scripts/bid_review.py itt.pdf spec.pdf --trade electrical --out job123-spec-review.xlsx
```

`--trade`: `all` (default), `electrical`, `mep`, `concrete`, `finishes`, or a comma list. `universal` is always added.

**Tip:** pass the whole tender set in one go - LADs, retention, bonds and the form of contract usually sit in the ITT / Contract Particulars, not the spec.

---

## Tested on

| Document | Pages | Result |
|---|---|---|
| Synthetic UK tender (`tests/fixtures/uk_tender_synthetic.txt`) | 6 | 18 checks (`tests/test_uk.py`) - no PDFs needed |
| Real UK construction-management trade-contract scope of works (client document, not included) | 39 | Lump-sum fixed price, monthly interim applications, 12-month defects/warranty; 31 hidden costs; 24 timed submittals; retention / LADs / contract period correctly NOT FOUND (they sit in the trade contract, not the scope) |

Not yet run on a full ITT with contract particulars - check bid-killer rows hardest.

`tests/regression.py` re-runs your own hand-checked real specs (kept outside the repo, in `SPEC_DIR`) - see the file for how to add one.

---

## Folder layout

```
SKILL.md                 <- the skill (instructions Claude follows)
scripts/bid_review.py    <- the engine
config/watchlists.json   <- trade cost-driver watchlists (edit to suit your scope)
tests/test_uk.py         <- checks on the synthetic tender
tests/fixtures/          <- the synthetic tender text
tests/regression.py      <- runner for your own real specs (PDFs stay outside the repo)
LICENSE
```

## Troubleshooting

- **"pdftotext was not found"** - install Poppler (above) and reopen the terminal, or set `PDFTOTEXT` to the full path of `pdftotext.exe`.
- **"no usable text layer ... scanned PDF"** - OCR it first (Acrobat "Recognize Text", or `ocrmypdf`).
- **"Cannot write ... probably open in Excel"** - close the workbook or pass a different `--out`.
- **Few bid-killers** - you ran a spec or scope only. Add the ITT / Contract Particulars / subcontract order to the same command.
- **Irrelevant cost items** - pick your trade with `--trade`, or edit `config/watchlists.json`.

## Limits

Text-layer PDFs only. Tables are read as laid out - check dates and values from tables. Drawings are not read. Costly Items and the RFI List need an estimator's judgment.

---

Based on [AutoConst Spec Bid-Review](https://github.com/hamzaabduljabbar/autoConst-spec-bid-review-claude) by Hamza Jabbar ([hamzajabbar.online](https://hamzajabbar.online)), UK version. Licence: see `LICENSE` - free for your own / internal use; no resale or rebranding as a product.
