# PO → SO Converter

Converts a retailer Purchase Order (PDF / HTML / XLS) into the Sales Order
Excel format the ERP imports. Runs entirely in your browser — nothing is
uploaded anywhere, there is no login, and no data is saved when you close the
tab.

## How to use it

1. **Double-click `index.html`.** It opens in your browser. No install.
2. **Step 0 — Load the SKU list.** Click the file picker and choose
   `sku_master.xlsx` from the shared drive (it sits next to `index.html`). A
   green banner means it's loaded and current. Amber means it loaded but is
   more than 60 days old — treat unmatched SKUs with extra care. Red means it
   isn't loaded yet — you can still convert, but every PO will be marked
   REVIEW and unknown SKUs will still correctly BLOCK.
3. **Step 1 — Choose the customer** from the dropdown. Don't mix customers in
   one upload.
4. **Step 2 — Drag and drop the PO files** (or the whole folder you
   downloaded from the customer's portal) onto the box, then click
   **Convert**.
5. **Step 3 — Check the results table.** Every PO gets one of three colored
   tags:
   - **READY (green)** — every field extracted and every check passed.
     Nothing to do.
   - **REVIEW (amber)** — it's in the export, but click the row to see why it
     needs a quick human look before you send the SO on (e.g. a case-to-piece
     conversion happened, or a discount applied).
   - **BLOCKED (red)** — it is **not** in the export. Click the row to see
     the exact reason. You'll need to encode that PO by hand or send the file
     to IT.
6. **Step 4 — Download the Excel.** This is the file you import into the ERP
   as Sales Orders. If anything was BLOCKED, the app tells you how many and
   reminds you they're not in the file.
7. Use **Copy exception list** to paste every REVIEW/BLOCKED reason into an
   email or chat — that's your checker's worklist.

## Adding a SKU the app doesn't recognize

If a PO is BLOCKED because of an unknown SKU, the row shows an **Add SKU**
form pre-filled with what's on the PO. Fill in the Megabox item code and
save — the app re-checks everything immediately, no need to re-upload.

**Important — this does NOT create the item in the ERP.** It only updates
this app's own checklist for this session. If the item genuinely doesn't
exist in the ERP yet, the Sales Order import will still fail. The correct
order is:

1. Create the item in the ERP first.
2. Re-export `sku_master.xlsx` from the ERP.
3. *Then* convert the PO.

The in-app "Add SKU" exists so one new SKU doesn't stall an entire batch —
not as a way around setting the item up properly.

When you've added SKUs during a session, two buttons appear:
- **Download updated SKU list** — a fresh `sku_master.xlsx` with your
  additions folded in. Save it over the shared-drive copy.
- **Download new SKU report** — just the additions, for whoever maintains the
  ERP item master.

## Who owns `sku_master.xlsx`

One named person re-exports it from the ERP monthly. Without an owner, the
list goes stale in a few months and the SKU-checking layer stops catching
anything real.

## What "READY / REVIEW / BLOCKED" actually means

The app never guesses. If it isn't sure about a number, it stops and tells
you exactly what to check — a flagged file costs you 30 seconds; a silently
wrong number costs a wrong delivery. See `CLAUDE.md` in this folder for the
full rule list if you want the details behind a particular flag.

## Adding a new customer later

1. Collect 5–10 real PO files from the new customer — at least one multi-line
   PO, one with a discount, and one with a case/pack unit of measure.
2. Hand them to Claude Code along with `PO2SO_SPEC.md`. It writes a new parser
   in `src/parsers.js` plus a golden test set.
3. `python3 build.py` then `node test/finaltest.js` must print
   `ALL GOLDEN SAMPLES PASS`.
4. **Parallel run it first.** For two weeks, that customer's POs get encoded
   both by hand and through the app, and someone compares them. Only after
   two clean weeks does manual encoding stop for that customer.
5. Log whatever new format quirk you hit in `CLAUDE.md` so it isn't
   re-discovered next time.

Never put a new customer straight into live use off one sample file.

## For developers

See `CLAUDE.md` for the export contract, validation rules, and every
per-customer layout fact learned from the real samples. Workflow:

```
npm install xlsx@0.18.5 pdfjs-dist@3.11.174
edit src/*
python3 build.py
node test/finaltest.js
```
