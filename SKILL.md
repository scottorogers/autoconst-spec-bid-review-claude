---
name: uk-spec-bid-review
description: Review a UK tender, spec or scope of works before pricing or signing. Reads the PDF(s) and builds a cited Excel spec-review form - bid-killers (tender return, validity, JCT/NEC form and amendments, LADs, bonds, retention, payment, completion, defects period, CDP, CDM), hidden costs ("deemed included", "allow for", at the Contractor's expense, scope not shown on drawings), "or approved equivalent" products, timed submittals, trade cost drivers and drafted RFIs. Every row cites section and PDF page; anything missing is NOT FOUND, never guessed. Use whenever someone in UK construction wants an ITT, tender pack, NBS spec, subcontract enquiry or trade-contract scope reviewed, checked or summarised - "what am I signing up to", "find the LADs / retention", "hidden costs", "what's deemed included", "draft RFIs / TQs", "tender review", "bid/no-bid" - even if they never say "bid review". Text-layer PDFs; scans need OCR.
---

# UK Spec Bid-Review

Turns a UK tender or spec PDF into the estimator's spec-review form: an Excel workbook where
every price-impacting requirement is pulled out and cited to its section and PDF page, so the
user checks lines instead of reading 40-400 pages cold.

The engine is `scripts/bid_review.py` (pattern extraction, no AI in the loop). Your job is to run
it, verify what matters against the page, and report it in a way a UK site PM can act on.

## Why the rules below matter

The people using this sign contracts off the back of it. A wrong LAD figure, a retention % from
the wrong clause, or a "deemed included" clause that gets missed costs real money, and a single
invented reference destroys trust in the whole output. So: quote only what the workbook or the
page says, cite it, and say NOT FOUND when it's not there.

## 1. Set up (once per machine)

Needs Python 3, `openpyxl`, and Poppler's `pdftotext`.

```bash
python3 -c "import openpyxl" || python3 -m pip install openpyxl
pdftotext -v
```

If `pdftotext` is missing, install Poppler - Windows: `winget install --id oschwartz10612.Poppler -e`;
macOS: `brew install poppler`; Debian/Ubuntu: `apt-get install -y poppler-utils` - or set the
`PDFTOTEXT` environment variable to its full path. If you cannot install it, stop and tell the
user; don't substitute a different extractor, the patterns are tuned to `pdftotext -layout` output.

## 2. Run it

Use the absolute path to this skill's `scripts/bid_review.py`. Write the workbook to the user's
working folder (the skill folder may be read-only):

```bash
python3 <skill-dir>/scripts/bid_review.py itt.pdf spec.pdf scope.pdf --trade electrical --out <work-folder>/<job>-spec-review.xlsx
```

- **Pass the whole tender set in one command.** LADs, retention, bonds and the form of contract
  usually live in the ITT / Contract Particulars / subcontract order, not the NBS spec or scope.
  Rows then cite which document they came from.
- **`--trade`**: `electrical` for an electrical / data / containment package; `mep`, `concrete`,
  `finishes`, or a comma list; `all` (default) for a main-contractor view. The `universal` list
  (CDM, permits, out-of-hours, live environment, attendances, O&Ms, defects...) is always added.
  Picking the trade cuts noise - `all` on an IT package flags pipework and concrete items too.
- A scanned PDF is refused with an OCR message - tell the user to OCR it (Acrobat "Recognize
  Text", or `ocrmypdf`). A `pdftotext -layout` `.txt` file also works as input.

## 3. Verify before you report

The workbook is high-recall, not zero-noise. Before quoting anything the user will price or sign
against - every bid-killer value, the top hidden costs, any date - read that PDF page:

```bash
pdftotext -layout -f <page> -l <page> file.pdf - | less
```

Check that the value belongs to that item (not the next clause), and that "hidden cost" rows are
really the contractor's cost (e.g. "free issue ... free of charge to the Contractor" is the
Client paying). Drop or correct anything that doesn't stand up, and say you did.

Rows marked **MEDIUM** need that read most: a value taken from the line after a heading, a timed
deliverable found without an explicit "submit", every Costly Item and every RFI.

## 4. Report back

Lead with the answer. Use UK site / QS language, tight wording, a citation on every line
(`Section / clause, PDF p.N`). Page numbers are **PDF pages** (1 = first sheet of the file), not
the printed footer number - say so once.

1. **Bid-killers** - table: item | value | section | PDF page. Then one line listing everything
   **NOT FOUND**, and which document it normally sits in (ITT, Contract Particulars, trade
   contract / subcontract order, Employer's Requirements). Never fill a gap with a typical value.
2. **Biggest price risks** - the 5-10 hidden costs that move money most (unshown scope deemed
   yours, no adjustment to the Contract Sum, retest / other parties' costs, out-of-hours
   assumptions, 24-hour call-out...), each with its page.
3. **Timed submittals** - what, when, page (e.g. ITP 8 weeks before start, PDF p.32).
4. **RFIs / TQs to raise** - the few that change price or programme, using the drafted question
   from the RFI List, edited to be specific.
5. **Next action** - e.g. "send the ITT and Contract Particulars to complete the NOT FOUND items",
   or "qualify the tender on X".

Attach or link the workbook. Offer, don't dump - the full tabs are in the file.

## Confidentiality

Tender documents are usually confidential. Keep the PDF, extracted text and workbook in the
user's own folders. Don't upload them anywhere, paste large extracts into other services, or
commit them (or client names / quoted clauses) to a repository unless the user explicitly says so.

## What the workbook contains

| Tab | How it's found | Confidence |
|---|---|---|
| Summary | read first: bid-killers, hidden costs, submittals by type, costly items, top RFIs | - |
| Bid-Killers | literal clauses; the most informative wins (amounts, dates, "non-compliant", "rejected"); value from the matched sentence only | HIGH (MEDIUM if the value came from the line after a heading) |
| Hidden Costs | deemed included / allow for / at the Contractor's expense / not shown on drawings / no adjustment to Contract Sum / chargeable / contra charges / make good / retest | HIGH |
| Sole-Source | "X or approved equivalent", "similar approved", NBS `Manufacturer:` / `Product reference:` / `Substitution:` lines, manufacturer lists, basis of design | HIGH / MEDIUM |
| Submittals | "submit / submission", and timed deliverables ("8 weeks prior to", "within 1 month of appointment", "each week"); payment applications, TQs, CEs excluded | HIGH / MEDIUM |
| Costly Items | `config/watchlists.json` - one row per item with every page it appears on | MEDIUM - confirm it's in your scope |
| RFI List | only price-relevant ambiguity: unforeseen conditions, risk you must assume, scope set later by CA/CM, subjective / approval-based acceptance, match existing, provisional / PC sums, undefined standards | MEDIUM |
| Method & Notes | extractor used, confidence rules, limits | - |

Section labels read like `Y63 SUPPORT COMPONENTS - CABLES §210 Cable Ladder`,
`Section 6 Valuation Dates` or `1. SCOPE OF WORKS §1.5 Testing And Commissioning`.
Contents pages, reference-standard lists, definitions and running headers are skipped.

## Tuning

- Cost drivers: `config/watchlists.json` - each entry is `item`, `match` (case-insensitive
  regex; a space matches space or hyphen; `(?-i:\bUPS\b)` for capitals only) and `why`.
  Keep to things that move price.
- Bid-killer patterns: `BID_KILLERS` in `scripts/bid_review.py`; scoring in `hit_score()`.
- After any change, run `python3 tests/test_uk.py` from the repo (18 checks on a synthetic
  tender, no PDFs needed) if the tests are available.

## Limits

Text-layer PDFs only. Tables are read as laid out - check any date or value that came from a
table. Drawings are not read. Tested on a synthetic UK tender and one real 39-page CM trade-
contract scope; not yet on a full ITT with contract particulars, so check bid-killer rows hardest.
It flags; the estimator decides.

Based on AutoConst Spec Bid-Review by Hamza Jabbar (see LICENSE - internal use; no resale or
rebranding as a product).
