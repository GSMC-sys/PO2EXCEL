# PO2SO App — permanent record

Read `../PO2SO_SPEC.md` first. This file restates the parts that matter day-to-day and
logs concrete facts learned from the real sample files (traps, column layouts, edge
cases). Update this file whenever a new parser fact is discovered.

## The rule (§1)

`PARSE → VALIDATE → if uncertain: STOP, FLAG`. Three statuses only: READY / REVIEW /
BLOCKED. Never guess, never drop a line item, never round to make a total agree.

## Export contract (§3) — do not alter

One sheet `Sheet1`, 16 headers in this exact order:

```
PO# | Customer | Branch | Date | SKU | Item Code | Description | Order Qty. |
Price | Discount 1 | Discount 2 | Discount 3 | Total Amount | Area | P.O Date Exp | Remarks/P.O Note
```

`PO#` and `SKU` are TEXT (`@` format, leading zeros preserved). `Item Code`,
`Discount 1-3`, `Area` are ALWAYS BLANK — the ERP fills `Item Code` from `SKU` on
import. `Order Qty.` is always in PIECES. File name:
`SO_<Customer>_<YYYYMMDD-HHMM>.xlsx`.

## SKU verification (§4)

`sku_master.xlsx` (Customer | Customer SKU | Megabox Item Code | Description |
Date Added) lives next to `index.html`. Customer SKU compared as TEXT. Unknown
SKU → BLOCKED (`Unknown SKU <code> — not in SKU list`). Missing file → every PO
REVIEW (`SKU list not loaded — cannot verify item codes`). Newest `Date Added`
>60 days old → amber banner + REVIEW flag 5b. In-app "Add SKU" only updates the
app's checklist, never the ERP — README must say this loudly.

## Validation rules (§6)

BLOCK: (0) SKU not in master, (1) PO# missing/wrong shape, (2) line missing
SKU/desc/qty/price, (3) line count ≠ declared, (4) printed total ≠ sum of computed
line totals beyond ±₱0.05, (5) qty≤0 or price<0, (6) vendor isn't Good Shepherd /
Able & Good Shepherd Marketplace Corp, (7) date unparseable or >180 days from
**today** (real wall-clock date — parsers accept an injectable `now` for testing
so historical fixtures don't start failing as the calendar moves on), (8) duplicate
PO#+SKU in one batch.

REVIEW: (1) any case→piece conversion happened — always, no exceptions, (2) price÷pack
doesn't divide to 2dp cleanly, (3) unit price >20% different from same SKU elsewhere
in the batch, (4) line-level discount applied, (5) any Metro PO until §5.3 Test A
passes, (5b) SKU master missing/stale, (5c) SKU found but descriptions differ
materially, (6) PO expiry in the past or missing, (7) qty >5,000 pcs on one line.

## Per-customer facts learned from real samples

### Robinsons Handyman (`Handyman PO/*.xls`, actually XLSX — detect via `PK` zip magic)

- Fixed-width "print form" layout: every visible field lives in ONE cell of a run of
  merged narrow columns. Non-empty cells only; find fields by scanning for the label
  text, not fixed row numbers (row offsets shift by ±1 between samples depending on
  wrap height).
- Header block: `I10` = vendor code + name (accept `242437` or `242240`, reject
  anything else — wrong vendor's PO), `BB10` = PO No., `AH10` = Approved Date → `Date`,
  `AH11` = Delivery Date → `P.O Date Exp`, `I14` = Ship To (`CODE-NAME`; Branch =
  text after the first `-`).
- Item table: header row has `B{r}=="SKU CODE"`. Item rows follow with `B`=SKU,
  `Q`=Description, `AF`=Buying UOM (always `PIECE` in samples — no case conversion
  seen yet, but parser still checks), `AM`=Qty Ordered, `AV`=Unit Cost → `Price`,
  `BD`=Total Net Cost. Table ends at the row where `B{r}=="TOTAL GROSS AMOUNT"`
  (offset from header varies — scan for the label, don't hardcode).
- `TOTAL GROSS AMOUNT` / `LESS: TOTAL DISCOUNTS` / `TOTAL NET AMOUNT` are the three
  rows immediately after the item block, in that order — used for the total
  reconciliation (rule 4).
- `PO NOTES :` cell (label `D{r}=="PO NOTES  :"`, value in the next non-empty cell on
  that row) → `Remarks/P.O Note`, e.g. `DC FEE L-5%`.
- The sheet prints the SAME PO twice (a second, shorter summary table further down
  the sheet, no pricing) — ignore the second occurrence, only the first (priced)
  table is authoritative.
- 20 samples, every one exactly one line item. No case UOM seen. Two samples price
  a line at ₱0.01/pc (`PROMO SLIM DRAWER…`, `PROMO DRAWER BLACK…`) — internally
  consistent (qty×price=printed total) so no spec rule fires on them; flagged to
  Dana as an open question (see OPEN ITEMS below), not force-flagged as REVIEW.
- **Assumption, confirmed against §3's own worked example:** this exact sample file
  (`purchase_order1_202606221121037_…`) IS the source of the "real verified sample
  row" printed in §3 of the spec — every field matches exactly.

### The D.I.Y Shop Corp (`The DIY Shop PO.pdf`, SAP Business One, 41 pages / 35 POs)

- One PO is almost always one page; a PO that overflows its page continues onto the
  next page with the SAME PO# repeated in the header (item numbering keeps
  incrementing, table header repeats). Group consecutive pages by PO# in the
  `x≈480-524, y≈655-663` header zone.
- Per-page header (repeats on every page): PO# at `x≈440-570, y≈655-665`; `Order
  Date:` value at `x≈523, y≈732`; `Valid Until:` value at `x≈523, y≈713` →
  `P.O Date Exp`; `For Branch` code (`DIY-000N`) at `x≈275, y≈715` followed by 1-2
  address lines at `x≈275, y≈695-703` → `Branch` (human-readable location, code
  stripped, mirrors the Handyman convention).
- Item rows: each logical item spans 2-3 raw text lines clustered by y (±~12pt
  band, anchored on the line item number). Classify tokens by CONTENT, not fixed
  x: a token starting with `PHP` followed by a number closest to the `Price`
  column is price, the next `PHP` token to its right is the line total, a bare
  1-3 digit integer near the `QTY`/`Qty(Stock UoM)` columns is quantity, `PC` is
  the UOM, everything before the first `PHP`/bare-qty token is `# ItemCode
  Description [trailing pack-size digit]`. The trailing digit on the description
  line (e.g. `MEGABOX STORGE BOX 95L MG698  6`) is a reference pack-size shown by
  SAP, NOT used in any calculation — UOM is always `PC` (pieces) in every sample,
  so no case conversion ever fires for this customer.
- UOM is always `PC` (pieces) in every sample seen — no case conversion ever
  fires for this customer. If a future sample DOES show a non-PC UOM, treat it
  the same way as Puregold's `C0N` handling and flag REVIEW rule 1. **Not yet
  observed in the 35 samples — flag for Dana if this ever appears.**
- **Pricing basis:** the printed `Price` column is VAT-INCLUSIVE (same
  convention as Handyman's "TAX TYPE: VAT Inclusive with EWT"). Confirmed
  against every sample: `sum(qty × price)` per PO lands within a cent of the
  page's `Total:` line (VAT-inclusive grand total) and is ~12% (the VAT rate)
  away from `Total Before VAT:`. Reconcile against `Total:`, not `Total Before
  VAT:` — using the wrong one silently overstated every computed total by the
  VAT amount during development and was caught by the golden-test total check.

### Puregold Price Club (`PUREGOLD Sample PO/*.html`, EDI HTML, 44 files)

- One HTML file = one PO. Well-formed nested `<table>` markup, values inside
  `<font>`/`<center>`/`<left>`/`<right>` tags — extract by plain text stripped of
  tags, not by tag type.
- Header: `Purchase Order No.` (bold `<u>` text) → `PO#`. Vendor line has a
  15-digit vendor number then `(CODE_GOOD SHEPHERD MANUFACTURING CORP.)` — verify
  "GOOD SHEPHERD" appears, vendor number itself not otherwise validated (spec only
  names the Handyman two-code check explicitly). `Entry Date` → `Date`. `Cancel
  Date` → `P.O Date Exp`. `Delivery Location` line has `CODE_ (NAME)` → `Branch`
  = the `(...)`  content, code stripped, mirrors Handyman.
- Item table columns: `No.` `SKU NUMBER` `UPC` `DESCRIPTION` `BUY U/M` `BUY COST`
  `Qty Ord/CS` `Vendor/Spec.Ds` `Net Amount`. `BUY U/M` is either `PC` (pieces,
  no conversion) or `C0N` (case of N pieces — `C04`=4, `C03`=3, etc). Net Amount
  as printed is `BUY COST × Qty Ord`, i.e. **already in the ORIGINAL unit** (cases
  when U/M is `C0N`) — it is NOT pre-converted to pieces.
- Conversion (only when `C0N`): `Order Qty.(pcs) = QtyOrd × N`, `Price(per pc) =
  round(BUY COST / N, 2)`, `Total Amount = Order Qty.(pcs) × Price(per pc)`, which
  must equal the printed `Net Amount` for that line within tolerance.
- **Trap confirmed (§5.2 — validate BEFORE converting):** the `Sub Total` row's
  quantity figure is the sum of `Qty Ord/CS` **as printed** (i.e. in cases in
  the C0N rows, not pieces) — cross-check the printed Sub Total qty against the
  as-printed sum BEFORE any piece conversion, or a real mismatch gets hidden by
  the conversion math.
- `Sub Total` / `Total` Net Amount rows both equal the pre-discount line sum;
  `SKU Discounts` and `Other Discounts` rows are separate (both `.00` in every
  sample seen) — the grand total to reconcile against is the `Total` row.
- 44 samples, 17 use a `C0N` UOM somewhere → REVIEW rule 1 fires on exactly those
  17. The other 27 have only `PC` lines → READY (baseline 27/17/0 matches exactly).

### Metro Retail Stores Group (`METRO SAMPLE POS/*.pdf`, 27 files, text PDF)

- One PDF = one PO, one page (`Page: 1 OF 1` in every sample seen).
- Header: `No. <7-digit>` → `PO#`. `Vendor: 906683 - ABLE & GOOD SHEPHERD
  MARKETPLACE CORPORATION` — validate "GOOD SHEPHERD" appears (rule 6). `Entry
  Date:` → `Date`. `Cancel Date:` → `P.O Date Exp`. `Ship To: <code> - <STORE
  NAME>` → `Branch` (store name after the code, mirrors Handyman/Puregold).
- Item table columns (by x-range, left→right): `SKU` `VPN` `DESCRIPTION/BRAND`
  `UPC` `QTY ORDERED` `UOP PACK SIZE` `STD UOM` `UOP COST (Vat Excl)` `UOP COST
  (Vat Incl)` `DISCOUNT TYPE` `DISCOUNT VALUE` `NET UOP COST (Vat Excl)` `VAT
  RATE` `TOTAL EXT COST (Vat Excl)` `SELL UOM` `UNIT RETAIL`.
- **Confirmed by Dana 2026-09-14 — pricing basis:** the price we export is derived
  from `UOP COST (Vat Incl)` (VAT-inclusive, GROSS of any discount), netted by the
  printed discount:
  `Price = round(UOP_COST_VAT_INCL × (1 − discount_pct/100), 2)`,
  `Total Amount = round(Order Qty. × Price, 2)`.
  Reconcile the PO-level sum against the printed `PO Total Cost Net of Discount
  (VAT Incl)` line (±₱0.05) — **not** the VAT-Excl total, and **not** the printed
  per-line `TOTAL EXT COST (Vat Excl)` cell, which is itself VAT-excl and was
  observed off by a few centavos from qty×unit-cost in at least one real sample
  (`21085878`: 7×330.36=2312.52 printed as 2312.50) — the *document*-level VAT-Incl
  total is the one that reconciles exactly, confirmed against two real samples
  (one with `Percnt 7` discount, one with none).
- Discount column: `TYPE` cell is either `-` (blank/no discount) or `Percnt`; `VALUE`
  cell is the percent number (e.g. `7`). 14 of 27 samples carry a `Percnt`
  discount; 13 have none.
- **Open item, per §5.3/§9c: every Metro PO is REVIEW until the ERP Test A import
  passes — no exceptions, regardless of whether a discount fired.** This makes the
  spec's own stated baseline of "2 READY / 25 REVIEW" for this customer
  unreachable as written — see OPEN ITEMS below.

## OPEN ITEMS — do not invent answers (Dana to close)

Carried from spec §9, plus two new ones found during the build:

1. **`sku_master.xlsx` does not exist yet.** Built against a hand-made sample
   covering the SKUs seen across all 126 sample POs so the verification layer is
   testable; swap in Dana's real ERP export when available.
2. **Metro net-price Test A not yet run.** Every Metro PO exports as REVIEW until
   it passes (§5.3/§9c).
3. **Trial ERP imports not yet done** for Puregold, Metro, D.I.Y (§9d). Only
   Robinsons Handyman has ever been proven against the live ERP.
4. **NEW — spec's stated Metro baseline (2 READY / 25 REVIEW) conflicts with the
   spec's own "every Metro PO is REVIEW until Test A passes" rule (§5.3, §6 rule
   5, §9c).** As written those two statements can't both be true — 0 Metro POs
   can be READY while that rule is active. Built to the explicit rule (0 READY /
   27 REVIEW / 0 BLOCKED for the 27 Metro samples). Flag to Dana: if the "2 READY"
   figure is actually authoritative, tell us which 2 POs and why they should
   bypass the always-REVIEW rule.
5. **NEW — Handyman's stated baseline (18 READY / 2 REVIEW) has no data-driven
   explanation.** All 20 samples pass every BLOCK/REVIEW rule in §6 cleanly except
   two POs priced at ₱0.01/pc (`PROMO SLIM DRAWER MG164RT 3L L`, `PROMO DRAWER
   BLACK MG164 3L` — both internally consistent, qty×price=printed total exactly).
   No enumerated rule in §6 flags a low-but-consistent unit price. Built to the
   literal rule set (20 READY / 0 REVIEW / 0 BLOCKED). Flag to Dana: if those 2
   promo-priced POs should be REVIEW, tell us the rule (e.g. "unit price under
   ₱X is always worth a look") and it'll be added as an explicit numbered rule,
   not inferred.
6. **Golden-test date stability:** rule 7 ("PO date >180 days from today") uses
   real wall-clock time by design (catches misread years in production). All
   parsers take an injectable `now` (defaults to `new Date()`); the test harness
   pins it to a fixed reference date so the golden fixtures don't start failing
   as the calendar moves past the 2026 sample dates. This is an engineering
   decision, not a spec change — flagging it here so it isn't mistaken for a
   loosened rule.

## Workflow

```
npm install xlsx@0.18.5 pdfjs-dist@3.11.174   (one-time, needed by both build.py and finaltest.js)
edit src/* → python3 build.py → node test/finaltest.js
```

Must print `ALL GOLDEN SAMPLES PASS` before a build is considered releasable.
Node harness gotcha: strip a leading `"use strict"` from `parsers.js` before
`eval`-ing it (strict eval hides function declarations from the enclosing scope).

`build.py` inlines `node_modules/xlsx/dist/xlsx.full.min.js`,
`node_modules/pdfjs-dist/build/pdf.min.js`, and the PDF.js worker
(`pdf.worker.min.js`, base64-encoded so it can sit inside a `<script>` tag
without any `</script>`-escaping risk) directly into `index.html` — the
resulting file is ~2.6MB and fully offline. It refuses to build if any inlined
source literally contains `</script`, which would break out of its tag.

## Known limitation: SKU list auto-load

A double-clicked local HTML file cannot reliably read a sibling file
(`fetch('./sku_master.xlsx')` is blocked by Chrome under `file://`, though it
works in some other browsers). `index.html` tries the auto-fetch first as a
nicety and silently falls back — the **Step 0 file picker is the reliable
path** and is what the README tells office staff to use every time.

## PDF coordinate gotcha (Metro + D.I.Y Shop parsers)

Don't trust coordinates read from any tool other than `pdfjs-dist` itself —
different extractors (`pdfminer.six`, the Read tool's own PDF preview) scale
PDF user-space coordinates differently per file. Metro's PDFs report a
~1150×842 space via `pdfjs-dist`; D.I.Y Shop's report a standard 612×792.
Both were verified directly with a throwaway `pdfjsLib.getDocument(...).
getPage(1).getTextContent()` dump before any x/y range was hardcoded into
`parsers.js` — an assumption carried over from `pdfminer` coordinates caused a
real bug during development (unrelated header/footer text got merged into
item rows because line-grouping had no max-distance cutoff). If a future
customer's PDF parser misbehaves, dump real `pdfjs-dist` coordinates first.
