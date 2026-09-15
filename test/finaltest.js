// Golden regression harness. Run with: node test/finaltest.js
// Must print "ALL GOLDEN SAMPLES PASS" to be a releasable build (see CLAUDE.md).
//
// Node gotcha: strict-mode eval hides function declarations from the enclosing
// scope, so we strip a leading "use strict" from parsers.js before eval-ing it.

const fs = require("fs");
const path = require("path");

const APP_DIR = path.join(__dirname, "..");
const SAMPLES_ROOT = path.join(APP_DIR, ".."); // sample folders sit one level above the repo
const GOLDEN_DIR = path.join(__dirname, "golden");

const XLSX = require(path.join(APP_DIR, "node_modules/xlsx"));
global.XLSX = XLSX;
const pdfjsLib = require(path.join(APP_DIR, "node_modules/pdfjs-dist/legacy/build/pdf.js"));
global.pdfjsLib = pdfjsLib;

let src = fs.readFileSync(path.join(APP_DIR, "src/parsers.js"), "utf8");
src = src.replace(/^"use strict";\n/, "");
eval(src);

// Reference "now" per customer: pinned just before every sample's own P.O Date
// Exp so rule 6 ("expiry in the past") never fires from the calendar simply
// moving on. Rule 7 ("PO date >180 days from today") still uses a value close
// to the samples' own dates so historical fixtures don't start BLOCKing either
// — see CLAUDE.md "Golden-test date stability".
const NOW_BY_CUSTOMER = {
  handyman: new Date("2026-01-05T00:00:00Z"),
  puregold: new Date("2026-04-05T00:00:00Z"),
  metro: new Date("2026-06-10T00:00:00Z"),
  diy: new Date("2026-06-10T00:00:00Z")
};

function freshSkuIndex(entries, now) {
  const idx = buildSkuMasterIndex(entries);
  idx.newestDate = new Date(now.getTime() - 86400000); // 1 day old -> never stale
  return idx;
}

function snapshotPo(po) {
  return {
    sourceFile: po.sourceFile,
    poNumber: po.poNumber,
    branch: po.branch,
    date: po.date ? dateToMMDDYYYY(po.date) : null,
    expDate: po.expDate ? dateToMMDDYYYY(po.expDate) : null,
    remarks: po.remarks || "",
    printedTotal: po.printedTotal,
    status: po.status,
    reasons: po.reasons.map(r => `${r.severity} ${r.rule}: ${r.message}`).sort(),
    lines: po.lines.map(l => ({
      sku: l.sku,
      description: l.description,
      qty: l.qty,
      price: l.price,
      totalAmount: l.totalAmount
    }))
  };
}
function compareSnapshots(customerKey, actual, golden) {
  const errors = [];
  if (actual.length !== golden.length) {
    errors.push(`PO count mismatch: parsed ${actual.length}, golden has ${golden.length}`);
  }
  const byFile = new Map(golden.map(g => [g.sourceFile, g]));
  actual.forEach(a => {
    const g = byFile.get(a.sourceFile);
    if (!g) { errors.push(`New/unexpected PO in output: ${a.sourceFile}`); return; }
    byFile.delete(a.sourceFile);
    const aStr = JSON.stringify(a, null, 2);
    const gStr = JSON.stringify(g, null, 2);
    if (aStr !== gStr) {
      errors.push(`MISMATCH ${a.sourceFile}:\n--- golden ---\n${gStr}\n--- actual ---\n${aStr}`);
    }
  });
  byFile.forEach((g, file) => errors.push(`Missing from output (present in golden): ${file}`));
  return errors;
}

async function loadHandyman() {
  const dir = path.join(SAMPLES_ROOT, "Handyman PO");
  const files = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith(".xls")).sort();
  let pos = [];
  files.forEach(f => {
    const bytes = new Uint8Array(fs.readFileSync(path.join(dir, f)));
    pos = pos.concat(parseHandyman(bytes, f));
  });
  return pos;
}

async function loadPuregold() {
  const dir = path.join(SAMPLES_ROOT, "PUREGOLD Sample PO");
  const files = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith(".html")).sort();
  let pos = [];
  files.forEach(f => {
    const html = fs.readFileSync(path.join(dir, f), "utf8");
    pos = pos.concat(parsePuregold(html, f));
  });
  return pos;
}

async function loadMetro() {
  const dir = path.join(SAMPLES_ROOT, "METRO SAMPLE POS");
  const files = fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith(".pdf")).sort();
  let pos = [];
  for (const f of files) {
    const bytes = new Uint8Array(fs.readFileSync(path.join(dir, f)));
    pos = pos.concat(await parseMetro(bytes, f));
  }
  return pos;
}

async function loadDiy() {
  const file = path.join(SAMPLES_ROOT, "The DIY Shop PO.pdf");
  const bytes = new Uint8Array(fs.readFileSync(file));
  return await parseDiy(bytes, "The DIY Shop PO.pdf");
}

function assertExportContract(pos, customerKey) {
  const errors = [];
  const wb = buildWorkbook(pos, customerKey);
  const sheetNames = wb.SheetNames;
  if (sheetNames.length !== 1 || sheetNames[0] !== "Sheet1") {
    errors.push(`Export sheet must be named exactly "Sheet1" (got: ${JSON.stringify(sheetNames)})`);
  }
  const ws = wb.Sheets["Sheet1"];
  EXPORT_HEADERS.forEach((h, c) => {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    const cell = ws[addr];
    if (!cell || cell.v !== h) errors.push(`Header column ${c} expected "${h}", got ${cell ? JSON.stringify(cell.v) : "undefined"}`);
  });
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; R++) {
    const poCell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
    const skuCell = ws[XLSX.utils.encode_cell({ r: R, c: 4 })];
    if (poCell && poCell.v !== "" && poCell.t !== "s") errors.push(`Row ${R + 1}: PO# must be TEXT type, got t=${poCell.t}`);
    if (skuCell && skuCell.v !== "" && skuCell.t !== "s") errors.push(`Row ${R + 1}: SKU must be TEXT type, got t=${skuCell.t}`);
    [5, 9, 10, 11, 13].forEach(c => { // Item Code, Discount 1-3, Area
      const cell = ws[XLSX.utils.encode_cell({ r: R, c })];
      if (cell && cell.v !== "") errors.push(`Row ${R + 1}: column ${c} (${EXPORT_HEADERS[c]}) must be blank, got ${JSON.stringify(cell.v)}`);
    });
  }
  return errors;
}

async function runCustomer(customerKey, loader, updateGolden) {
  const skuBytes = new Uint8Array(fs.readFileSync(path.join(APP_DIR, "sku_master.SAMPLE.xlsx")));
  const skuEntries = parseSkuMasterWorkbook(skuBytes);
  const now = NOW_BY_CUSTOMER[customerKey];

  const pos = await loader();
  validateBatch(customerKey, pos, { now, skuMasterIndex: freshSkuIndex(skuEntries, now) });

  const actual = pos.map(snapshotPo);
  const goldenPath = path.join(GOLDEN_DIR, customerKey + ".json");

  if (updateGolden) {
    fs.mkdirSync(GOLDEN_DIR, { recursive: true });
    fs.writeFileSync(goldenPath, JSON.stringify(actual, null, 2) + "\n");
    console.log(`[${customerKey}] golden written (${actual.length} POs)`);
    return { errors: [], counts: statusCounts(pos) };
  }

  if (!fs.existsSync(goldenPath)) {
    return { errors: [`No golden file at ${goldenPath} — run with --update-golden after reviewing the diff by hand`], counts: statusCounts(pos) };
  }
  const golden = JSON.parse(fs.readFileSync(goldenPath, "utf8"));
  const errors = compareSnapshots(customerKey, actual, golden);
  errors.push(...assertExportContract(pos, customerKey));
  return { errors, counts: statusCounts(pos) };
}

function statusCounts(pos) {
  const c = { READY: 0, REVIEW: 0, BLOCKED: 0 };
  pos.forEach(po => c[po.status]++);
  return c;
}

async function main() {
  const updateGolden = process.argv.includes("--update-golden");
  const customers = [
    ["handyman", loadHandyman],
    ["diy", loadDiy],
    ["puregold", loadPuregold],
    ["metro", loadMetro]
  ];

  let anyError = false;
  for (const [key, loader] of customers) {
    const { errors, counts } = await runCustomer(key, loader, updateGolden);
    console.log(`[${key}] READY:${counts.READY} REVIEW:${counts.REVIEW} BLOCKED:${counts.BLOCKED}`);
    if (errors.length > 0) {
      anyError = true;
      console.log(`[${key}] ${errors.length} FAILURE(S):`);
      errors.forEach(e => console.log(e));
    }
  }

  if (updateGolden) {
    console.log("\nGolden files written. Re-run without --update-golden to verify.");
    return;
  }

  if (anyError) {
    console.log("\nGOLDEN TEST FAILED");
    process.exit(1);
  } else {
    console.log("\nALL GOLDEN SAMPLES PASS");
  }
}

main().catch(e => { console.error(e); process.exit(1); });
