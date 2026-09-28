---
name: spec-bid-review
description: UK-first (use --region us for US CSI specs). Read a construction specification / tender PDF (or a whole bid set) and produce the estimator's spec-review form as Excel - every price-impacting requirement pulled out and cited to its section and page, so a bidder does not read 300 pages by hand. Tabs - Summary; BID-KILLERS (bid due date, mandatory pre-bid, bid bond %, performance/payment bonds, post-award deadlines, bid validity, liquidated damages $/day, fines, contract time, insurance, prevailing wage, license, retainage, alternates, addenda, required forms) with the key value pulled out; HIDDEN COSTS (no separate payment, incidental, at Contractor's expense); SOLE-SOURCE / limited list / basis of design / or-equal; SUBMITTAL register; trade COSTLY ITEMS (concrete/MEP/finishes/universal watchlists); ranked RFI list with drafted questions. Every row has section + page + confidence; missing bid-killers are reported NOT FOUND, never invented. Use when a user asks to review a spec before bidding, find bid-killers, find hidden costs, pull submittals, catch sole-source products, draft RFIs, or build a spec-review checklist. Text-layer PDFs only; scanned specs need OCR first.
---

# Spec Bid-Review

Reading a spec book to price a bid takes an estimator 1-3 hours per job (Red Rhino and
MEP Academy both say so) - and the things that lose money are buried: a liquidated-damages
clause, a sole-source manufacturer, a submittal nobody priced, "no separate payment" scope,
an addendum not acknowledged that makes the bid non-responsive.

Estimators already keep a **spec-review form** - a sheet listing each price-impacting
requirement with its section and page. This skill **fills that form automatically**.

## The reliability rule (do not skip)

Pattern extraction is high-recall, not zero-noise. Every row is a real sentence from the
document, cited to Section + Page, so the estimator can verify it.

| Tab | How it's found | Confidence |
|---|---|---|
| Bid-Killers | literal clauses; the most informative clause wins (amounts, dates, "non-responsive" beat passing mentions); the next sentence is read too because values often sit there | **HIGH** - verify the value |
| Hidden Costs | "no separate payment", "incidental", "included in the unit price", "at Contractor's expense", "at no cost to the Owner", retesting | **HIGH** |
| Sole-Source | "manufactured by X", "X or Y" lists, "one of the following", basis of design, global or-equal clauses | **HIGH** with explicit substitution language, **MEDIUM** otherwise |
| Submittals | the spec's own submittal schedule table first, then every sentence requiring a deliverable; claims, notices, RFIs, field sampling and bid-phase items excluded | **HIGH** |
| Costly Items | trade watchlists (`config/watchlists.json`), one row per item with every page; prohibitions/reimbursements flagged | **MEDIUM** - confirm it is in your scope |
| RFI List | only price-relevant ambiguity: unforeseen conditions, risk assumed by bidder, scope set later "as directed", subjective/undefined acceptance, match existing, undefined standards; legal/admin boilerplate excluded | **MEDIUM** - judgment call |

Automatically ignored everywhere: table-of-contents and drawing-index pages, "Applicable
Publications / References / Definitions / Related Requirements" articles, standards lists,
running headers and footers.

## Regions

- `--region uk` (default): ITT / JCT / NEC / FIDIC, NBS work sections (`Y61`) and clauses (`210`), Uniclass,
  £, LADs per week, Employer / Client / Contract Administrator, "or approved equivalent", NBS `Manufacturer:`
  lines, `electrical` + `universal_uk` watchlists. Values come from the matched sentence only; a value
  borrowed from the line after a heading is MEDIUM with a note. Tested on a synthetic fixture only so far.
- `--region us`: the original engine, unchanged.

Page numbers in every tab are **PDF page** numbers, not the printed footer page.

## Workflow

Needs Python 3 + `openpyxl`, and Poppler `pdftotext` on PATH.

1. **Run it on the PDF(s).** Pass the whole bid set - bid-killers usually live in the bid
   invitation / solicitation, not the technical specification:
   ```
   py scripts/bid_review.py spec.pdf [itt.pdf ...] --trade all --out outputs/spec-review.xlsx
   ```
   `--region`: `uk` (default) or `us`. `--trade`: `all` (default), `electrical`, `concrete`, `mep`,
   `finishes`, or a comma list.
   Scanned PDFs (no text layer) are refused with an OCR message.

2. **Open the workbook.** Tabs: Summary (read first), Bid-Killers, Hidden Costs,
   Sole-Source, Submittals, Costly Items, RFI List, Method & Notes. Section column reads
   like `01 00 00 §1.10 Restoration`, `Section IV. Control` or `Exhibit I`.

## How to use this as Claude

- Run `bid_review.py`, then read the Summary back to the user: bid-killers with their
  values and pages first, then hidden costs, then the top RFIs.
- Cite the section and PDF page on every requirement. If a bid-killer is NOT FOUND, say so and
  point out it may be in a document that was not included (UK: ITT / Contract Particulars /
  subcontract enquiry; US: solicitation / SF-1442).
- A row marked MEDIUM with "value is from the sentence after the heading" - read the page before
  quoting the value.
- If the Summary warns there is little tender / Division 00/01 content, ask for the ITT / bid invitation.
- Only open a raw PDF page when a table or figure has no prose around it.

## Tuning

- Trade watchlists: `config/watchlists.json` - each entry has `item`, `match` (regex; a
  space matches space or hyphen) and `why`. Keep them to price drivers, not boilerplate.
- Bid-killer patterns: `BID_KILLERS_UK` / `BID_KILLERS` (US) in `scripts/bid_review.py`; region switch in `configure()`; scoring in
  `hit_score()`.
- Tested on: a City of Alameda public-works tender (406 pp, 18/19 bid-killers, all
  correct), a VA federal technical spec (338 pp), a 260-page municipal electrical/civil
  technical spec, a short tender (NWIT) and a Division 03 concrete section.

## Regression tests

`py tests/test_uk.py` checks UK mode against `tests/fixtures/uk_tender_synthetic.txt` (no PDFs or
pdftotext needed). It must stay at 0 failed.

`py tests/regression.py` re-runs the tested specs under every `pdftotext` on the machine
(xpdf and Poppler emit different characters - text is normalised on load so both give the
same answers) and checks the hand-verified values. Run it after any change to patterns or
watchlists; it must stay at 0 failed. Specs are read from `SPEC_DIR` (default: the folder
above this repo); missing ones are skipped. Set `PDFTOTEXT` to pin one extractor.

## Limits

Text-layer PDFs only. Tables are read as laid out (misaligned columns can shift a due
date - verify). Drawings are not read. It flags; the estimator decides.
