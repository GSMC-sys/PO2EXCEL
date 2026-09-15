"use strict";
/* PO2SO parsers + validators + SKU verification + export builder.
   Runs unmodified in the browser (globals: XLSX, pdfjsLib) and under Node
   (test/finaltest.js requires xlsx + pdfjs-dist and assigns them to the same
   global names before eval-ing this file). No AI/OCR/fuzzy matching anywhere —
   every extraction is a deterministic rule tied to a known layout. */

// ---------------------------------------------------------------------------
// Shared constants
// ---------------------------------------------------------------------------

var EXPORT_HEADERS = [
  "PO#", "Customer", "Branch", "Date", "SKU", "Item Code", "Description",
  "Order Qty.", "Price", "Discount 1", "Discount 2", "Discount 3",
  "Total Amount", "Area", "P.O Date Exp", "Remarks/P.O Note"
];

var GOOD_SHEPHERD_NAME_RE = /GOOD\s+SHEPHERD/i;

var CUSTOMERS = {
  handyman: { key: "handyman", label: "Robinsons Handyman", legalName: "ROBINSONS HANDYMAN INC." },
  diy:      { key: "diy",      label: "The D.I.Y Shop Corp", legalName: "THE D.I.Y SHOP CORP." },
  puregold: { key: "puregold", label: "Puregold Price Club", legalName: "PUREGOLD PRICE CLUB INC." },
  metro:    { key: "metro",    label: "Metro Retail Stores Group", legalName: "METRO RETAIL STORES GROUP, INC." }
};

// ---------------------------------------------------------------------------
// Small shared utilities
// ---------------------------------------------------------------------------

function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

function cleanText(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/\s+/g, " ").trim();
}

function stripTags(html) {
  return html.replace(/<[^>]*>/g, " ");
}

function decodeEntities(s) {
  return s
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

// Accepts "1,792.00", "1792.0000", "758", "-" (=> null)
function parseMoney(raw) {
  if (raw === null || raw === undefined) return null;
  var s = cleanText(String(raw));
  if (s === "" || s === "-" || s === "--" || s === "---") return null;
  s = s.replace(/^PHP\s*/i, "").replace(/,/g, "").trim();
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseIntSafe(raw) {
  var n = parseMoney(raw);
  return n === null ? null : Math.round(n);
}

// Returns {y:number,m:number,d:number} or null. Accepts:
//  MM/DD/YYYY, DD-MON-YYYY, YYYY-MM-DDTHH:MM:SS, DD/MM/YYYY(Puregold entry uses ISO)
function parseDateFlexible(raw) {
  if (!raw) return null;
  var s = cleanText(String(raw));
  if (s === "") return null;

  var m;
  // ISO: 2026-04-07T00:00:00 or 2026-04-07
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return { y: +m[1], m: +m[2], d: +m[3] };

  // DD-MON-YYYY  e.g. 13-JUN-2026
  m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (m) {
    var months = { JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12 };
    var mon = months[m[2].toUpperCase()];
    if (mon) return { y: +m[3], m: mon, d: +m[1] };
  }

  // MM/DD/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return { y: +m[3], m: +m[1], d: +m[2] };

  return null;
}

function dateToMMDDYYYY(d) {
  if (!d) return null;
  var mm = String(d.m).padStart(2, "0");
  var dd = String(d.d).padStart(2, "0");
  return mm + "/" + dd + "/" + d.y;
}

function dateToJs(d) {
  if (!d) return null;
  return new Date(Date.UTC(d.y, d.m - 1, d.d));
}

function daysBetween(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function tokenOverlapRatio(a, b) {
  var ta = cleanText(a).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  var tb = cleanText(b).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  if (ta.length === 0 || tb.length === 0) return 1;
  var setB = new Set(tb);
  var common = ta.filter(function (t) { return setB.has(t); }).length;
  return common / Math.max(ta.length, tb.length);
}

function isZipMagic(bytes) {
  return bytes && bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b; // 'PK'
}

// ---------------------------------------------------------------------------
// PO / line item shape
//
// po = {
//   sourceFile, customerKey, poNumber, branch, date /*{y,m,d}*/, expDate,
//   remarks, vendorOk /*bool*/, declaredLineCount /*number|null*/,
//   printedTotal /*number|null*/, printedTotalLabel /*string*/,
//   lines: [{ sku, description, uom, qtyOriginal, packSize, qty, price,
//              totalAmount, discountPct, caseConverted, printedLineTotal }],
//   parseErrors: [string]   // hard failures found DURING extraction (missing
//                            // header cell etc) -> always BLOCK, rule 2/1 style
// }
// ---------------------------------------------------------------------------

function newPo(customerKey, sourceFile) {
  return {
    sourceFile: sourceFile,
    customerKey: customerKey,
    poNumber: null,
    branch: null,
    date: null,
    expDate: null,
    remarks: "",
    vendorOk: false,
    declaredLineCount: null,
    printedTotal: null,
    printedTotalLabel: null,
    lines: [],
    parseErrors: []
  };
}

// ---------------------------------------------------------------------------
// 1. ROBINSONS HANDYMAN  (xlsx cells, printed-form layout)
// ---------------------------------------------------------------------------

function handymanCellText(ws, addr) {
  var cell = ws[addr];
  if (!cell) return "";
  var v = cell.v;
  if (v === undefined || v === null) return "";
  return cleanText(String(v));
}

function handymanFindCell(ws, range, matchFn) {
  for (var R = range.s.r; R <= range.e.r; R++) {
    for (var C = range.s.c; C <= range.e.c; C++) {
      var addr = XLSX.utils.encode_cell({ r: R, c: C });
      var txt = handymanCellText(ws, addr);
      if (txt && matchFn(txt, R, C)) return { row: R, col: C, addr: addr, text: txt };
    }
  }
  return null;
}

function handymanRowCells(ws, row, range) {
  var out = {};
  for (var C = range.s.c; C <= range.e.c; C++) {
    var addr = XLSX.utils.encode_cell({ r: row, c: C });
    var txt = handymanCellText(ws, addr);
    if (txt) out[XLSX.utils.encode_col(C)] = txt;
  }
  return out;
}

function parseHandyman(bytes, filename) {
  var po = newPo("handyman", filename);

  if (!isZipMagic(bytes)) {
    po.parseErrors.push("File is not a real XLSX (missing ZIP/PK signature) despite .xls name");
    return [po];
  }

  var wb;
  try {
    wb = XLSX.read(bytes, { type: "array" });
  } catch (e) {
    po.parseErrors.push("Could not open workbook: " + e.message);
    return [po];
  }

  var ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws["!ref"]) {
    po.parseErrors.push("Worksheet is empty");
    return [po];
  }
  var range = XLSX.utils.decode_range(ws["!ref"]);

  // Vendor + PO# live on the row that has "VENDOR" in column B.
  var vendorLabelCell = handymanFindCell(ws, range, function (t) { return t === "VENDOR"; });
  if (vendorLabelCell) {
    var row = handymanRowCells(ws, vendorLabelCell.row, range);
    var vendorText = row["I"] || "";
    po.vendorOk = /24243[7]|242240/.test(vendorText) && GOOD_SHEPHERD_NAME_RE.test(vendorText);
    po.poNumber = (row["BB"] || "").trim() || null;
    po.date = parseDateFlexible(row["AH"]);
  } else {
    po.parseErrors.push("Could not find VENDOR row");
  }

  // Delivery date sits one row below (AA/AH pair repeats: DELIVERY DATE).
  if (vendorLabelCell) {
    var delivRow = handymanRowCells(ws, vendorLabelCell.row + 1, range);
    po.expDate = parseDateFlexible(delivRow["AH"]);
  }

  // Branch: "SHIP TO" row, column I = "CODE-NAME"
  var shipToCell = handymanFindCell(ws, range, function (t) { return t === "SHIP TO"; });
  if (shipToCell) {
    var shipRow = handymanRowCells(ws, shipToCell.row, range);
    var raw = shipRow["I"] || "";
    var dash = raw.indexOf("-");
    po.branch = dash >= 0 ? raw.slice(dash + 1).trim() : raw;
  } else {
    po.parseErrors.push("Could not find SHIP TO row");
  }

  // Item table: header row has B == "SKU CODE"
  var hdrCell = handymanFindCell(ws, range, function (t) { return t === "SKU CODE"; });
  if (!hdrCell) {
    po.parseErrors.push("Could not find SKU CODE item table header");
    return [po];
  }

  var r = hdrCell.row + 1;
  var totalsRow = null;
  var maxScan = hdrCell.row + 30;
  while (r <= Math.min(maxScan, range.e.r)) {
    var cells = handymanRowCells(ws, r, range);
    var bVal = cells["B"];
    if (bVal === "TOTAL GROSS AMOUNT") { totalsRow = r; break; }
    if (bVal && bVal !== "LESS: TOTAL DISCOUNTS" && bVal !== "TOTAL NET AMOUNT") {
      var qty = parseMoney(cells["AM"]);
      var price = parseMoney(cells["AV"]);
      var lineTotal = parseMoney(cells["BD"]);
      po.lines.push({
        sku: bVal,
        description: cells["Q"] || "",
        uom: (cells["AF"] || "").toUpperCase(),
        qtyOriginal: qty,
        packSize: 1,
        qty: qty,
        price: price,
        totalAmount: lineTotal !== null ? lineTotal : (qty !== null && price !== null ? round2(qty * price) : null),
        discountPct: 0,
        caseConverted: false,
        printedLineTotal: lineTotal
      });
    }
    r++;
  }

  if (totalsRow !== null) {
    var grossRow = handymanRowCells(ws, totalsRow, range);
    po.printedTotal = parseMoney(grossRow["BB"]);
    po.printedTotalLabel = "TOTAL GROSS AMOUNT";
    // LESS: TOTAL DISCOUNTS / TOTAL NET AMOUNT are usually 2 rows further down,
    // but scan a few rows to be safe rather than assume a fixed offset.
    for (var rr = totalsRow + 1; rr <= totalsRow + 6 && rr <= range.e.r; rr++) {
      var c2 = handymanRowCells(ws, rr, range);
      if (c2["B"] === "TOTAL NET AMOUNT") {
        po.printedTotal = parseMoney(c2["BB"]);
        po.printedTotalLabel = "TOTAL NET AMOUNT";
        break;
      }
    }
  } else {
    po.parseErrors.push("Could not find TOTAL GROSS AMOUNT row");
  }

  // PO NOTES : <remarks>  -- label cell text is "PO NOTES  :" with variable spacing
  var notesCell = handymanFindCell(ws, range, function (t) { return /^PO NOTES\b/i.test(t); });
  if (notesCell) {
    var notesRow = handymanRowCells(ws, notesCell.row, range);
    var parts = [];
    for (var C = notesCell.col + 1; C <= range.e.c; C++) {
      var addr = XLSX.utils.encode_cell({ r: notesCell.row, c: C });
      var t = handymanCellText(ws, addr);
      if (t) parts.push(t);
    }
    po.remarks = parts.join(" ").replace(/,+$/, "").trim();
  }

  po.declaredLineCount = null; // Handyman PO doesn't print a line-count field
  return [po];
}

// ---------------------------------------------------------------------------
// 2. PUREGOLD PRICE CLUB  (EDI HTML)
// ---------------------------------------------------------------------------

function puregoldCellsFromRow(rowHtml) {
  var cells = [];
  var re = /<td[^>]*>([\s\S]*?)<\/td>/gi;
  var m;
  while ((m = re.exec(rowHtml))) {
    var text = decodeEntities(cleanText(stripTags(m[1])));
    cells.push(text);
  }
  return cells;
}

function parsePuregold(html, filename) {
  var po = newPo("puregold", filename);

  var poNoMatch = html.match(/Purchase Order No\.[\s\S]*?<u>([\s\S]*?)<\/u>/i);
  po.poNumber = poNoMatch ? cleanText(stripTags(poNoMatch[1])) : null;

  var vendorMatch = html.match(/Vendor\s*:<\/b><\/font><\/td><td[^>]*><font[^>]*>([\s\S]*?)<\/font><\/td>/i);
  var vendorText = vendorMatch ? decodeEntities(cleanText(stripTags(vendorMatch[1]))) : "";
  po.vendorOk = GOOD_SHEPHERD_NAME_RE.test(vendorText);

  var entryDateMatch = html.match(/Entry Date\s*-\s*<\/b>([\s\S]*?)<\/font>/i);
  po.date = entryDateMatch ? parseDateFlexible(stripTags(entryDateMatch[1])) : null;

  var cancelDateMatch = html.match(/Cancel Date\s*<\/b>\s*-\s*<\/font><\/td><td[^>]*><font[^>]*><b>([\s\S]*?)<\/b>/i);
  if (!cancelDateMatch) cancelDateMatch = html.match(/Cancel Date\s*\t?-\s*<\/b>([\s\S]*?)<\/font>/i);
  po.expDate = cancelDateMatch ? parseDateFlexible(stripTags(cancelDateMatch[1])) : null;

  var deliveryLocMatch = html.match(/Delivery Location\s*:<\/b><\/font><\/td><td[^>]*><font[^>]*>([\s\S]*?)<\/font><\/td>/i);
  if (deliveryLocMatch) {
    var raw = decodeEntities(cleanText(stripTags(deliveryLocMatch[1])));
    var m = raw.match(/\(([^)]+)\)/);
    po.branch = m ? m[1].replace(/^\d+_?/, "").trim() : raw;
  }

  // Item table: rows inside the <table style="line-height:75%"> that follows the
  // green header row containing "SKU NUMBER". We find that table's <tr> blocks by
  // scanning the whole doc for rows whose FIRST cell is a bare row number (1,2,3…)
  // immediately followed by a SKU-looking second cell.
  var declaredCount = 0;
  var rowRe = /<tr>([\s\S]*?)<\/tr>/gi;
  var rowMatch;
  var lastNo = 0;
  while ((rowMatch = rowRe.exec(html))) {
    var cells = puregoldCellsFromRow(rowMatch[1]);
    if (cells.length < 9) continue;
    var no = cells[0];
    if (!/^\d+$/.test(no)) continue;
    var noInt = parseInt(no, 10);
    if (noInt !== lastNo + 1) continue; // not a sequential item row
    lastNo = noInt;
    declaredCount = noInt;

    var sku = cells[1];
    var description = cells[3];
    var uom = cells[4].toUpperCase();
    var buyCost = parseMoney(cells[5]);
    var qtyOrdCS = parseMoney(cells[6]);
    var netAmount = parseMoney(cells[8]);

    var packMatch = uom.match(/^C0*(\d+)$/);
    var packSize = packMatch ? parseInt(packMatch[1], 10) : 1;
    var caseConverted = packSize > 1;

    var qtyPieces = qtyOrdCS !== null ? qtyOrdCS * packSize : null;
    var pricePerPiece = buyCost !== null ? round2(buyCost / packSize) : null;
    var totalAmount = qtyPieces !== null && pricePerPiece !== null ? round2(qtyPieces * pricePerPiece) : null;

    po.lines.push({
      sku: sku,
      description: description,
      uom: uom,
      qtyOriginal: qtyOrdCS,
      packSize: packSize,
      qty: qtyPieces,
      price: pricePerPiece,
      totalAmount: totalAmount,
      discountPct: 0,
      caseConverted: caseConverted,
      printedLineTotal: netAmount,
      // kept for the §5.2 pre-conversion subtotal check
      _rawQtyForSubtotal: qtyOrdCS
    });
  }
  po.declaredLineCount = declaredCount || null;

  // Sub Total row: <td ...>Sub Total...</td>...<td align="center">QTY</td>...<td align="right">AMOUNT</td>
  var subTotalMatch = html.match(/Sub Total[\s\S]*?<\/b><\/font><\/td><td[^>]*><font[^>]*color="#F0E68C">([\d.,]+)<\/font><\/td><td[^>]*><\/td><td[^>]*align="right"[^>]*><font[^>]*color="#F0E68C">([\d.,]+)<\/font>/i);
  po._subTotalQty = subTotalMatch ? parseMoney(subTotalMatch[1]) : null;

  var totalMatch = html.match(/<b>Total\s*-[\s\S]*?<\/b><\/td><td[^>]*><\/td><td[^>]*align="right"[^>]*><b>([\d.,]+)<\/b>/i);
  po.printedTotal = totalMatch ? parseMoney(totalMatch[1]) : null;
  po.printedTotalLabel = "Total";

  return [po];
}

// ---------------------------------------------------------------------------
// Shared PDF text-extraction helpers (pdfjs-dist), used by Metro + D.I.Y Shop
// ---------------------------------------------------------------------------

async function pdfPageTextItems(page) {
  var content = await page.getTextContent();
  return content.items
    .filter(function (it) { return it.str && cleanText(it.str) !== ""; })
    .map(function (it) {
      return { x: it.transform[4], y: it.transform[5], text: cleanText(it.str) };
    });
}

// Groups raw text items into "lines" (same visual row) by y-proximity, each
// line's tokens sorted left-to-right.
function groupIntoLines(items, yTol) {
  var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
  var lines = [];
  sorted.forEach(function (it) {
    var line = lines.find(function (l) { return Math.abs(l.y - it.y) <= yTol; });
    if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
    line.items.push(it);
  });
  lines.forEach(function (l) { l.items.sort(function (a, b) { return a.x - b.x; }); });
  lines.sort(function (a, b) { return b.y - a.y; });
  return lines;
}

function linesToText(lines) {
  return lines.map(function (l) { return l.items.map(function (it) { return it.text; }).join(" "); }).join("\n");
}

// Groups lines into logical item-rows by nearest-anchor assignment: an anchor
// line is one whose leftmost token (x <= anchorXMax) matches anchorRe. Every
// other line attaches to whichever anchor's y is closest, but ONLY if within
// maxDist points (handles wrapped description lines that print just
// above/below their row, per the D.I.Y Shop SAP layout, without dragging in
// unrelated header/footer text).
function groupRowsByAnchor(lines, anchorRe, anchorXMax, maxDist) {
  var anchors = [];
  lines.forEach(function (l) {
    var first = l.items[0];
    if (first && first.x <= anchorXMax && anchorRe.test(first.text)) {
      anchors.push({ y: l.y, lines: [l] });
    }
  });
  if (anchors.length === 0) return [];
  lines.forEach(function (l) {
    var isAnchorLine = anchors.some(function (a) { return a.lines[0] === l; });
    if (isAnchorLine) return;
    var nearest = null;
    var best = Infinity;
    for (var i = 0; i < anchors.length; i++) {
      var d = Math.abs(anchors[i].y - l.y);
      if (d < best) { best = d; nearest = anchors[i]; }
    }
    if (nearest && best <= maxDist) nearest.lines.push(l);
  });
  return anchors;
}

function findLineContaining(lines, re) {
  for (var i = 0; i < lines.length; i++) {
    var text = lines[i].items.map(function (it) { return it.text; }).join(" ");
    if (re.test(text)) return { line: lines[i], text: text };
  }
  return null;
}

// ---------------------------------------------------------------------------
// 3. METRO RETAIL STORES GROUP  (single-page PDF per PO)
// ---------------------------------------------------------------------------

async function parseMetroPage(page, filename) {
  var po = newPo("metro", filename);
  var items = await pdfPageTextItems(page);
  var lines = groupIntoLines(items, 2.0);
  var fullText = linesToText(lines);

  var poMatch = fullText.match(/No\.\s*(\d{5,9})/);
  po.poNumber = poMatch ? poMatch[1] : null;

  var vendorLine = findLineContaining(lines, /^Vendor:/);
  po.vendorOk = vendorLine ? GOOD_SHEPHERD_NAME_RE.test(vendorLine.text) : GOOD_SHEPHERD_NAME_RE.test(fullText);

  var entryMatch = fullText.match(/Entry Date:\s*(\d{1,2}-[A-Z]{3}-\d{4})/);
  po.date = entryMatch ? parseDateFlexible(entryMatch[1]) : null;

  var cancelMatch = fullText.match(/Cancel Date:\s*(\d{1,2}-[A-Z]{3}-\d{4})/);
  po.expDate = cancelMatch ? parseDateFlexible(cancelMatch[1]) : null;

  var shipToLine = findLineContaining(lines, /^Ship To:/);
  if (shipToLine) {
    var m = shipToLine.text.match(/^Ship To:\s*\d+\s*-\s*(.+)$/);
    po.branch = m ? m[1].trim() : null;
  }

  var commentsLine = findLineContaining(lines, /^COMMENTS:/);
  po.remarks = commentsLine ? commentsLine.text.replace(/^COMMENTS:\s*/, "").trim() : "";

  // Totals footer: label and its value share the same visual row.
  var totalInclLine = findLineContaining(lines, /PO Total Cost Net of Discount\s*\(VAT Incl\):/);
  if (totalInclLine) {
    var numMatch = totalInclLine.text.match(/([\d,]+\.\d{2})/);
    po.printedTotal = numMatch ? parseMoney(numMatch[1]) : null;
  }
  po.printedTotalLabel = "PO Total Cost Net of Discount (VAT Incl)";

  // Item rows: restrict to the table body (below the column headers, above the
  // "Total Order Quantity:" footer line), anchor = leading 6-9 digit SKU at x<90.
  var headerLine = findLineContaining(lines, /VALUE$/);
  var footerLine = findLineContaining(lines, /^Total Order Quantity:/);
  var top = headerLine ? headerLine.line.y : 10000;
  var bottom = footerLine ? footerLine.line.y : -10000;
  var bodyLines = lines.filter(function (l) { return l.y < top && l.y > bottom; });

  var anchorRe = /^\d{6,9}$/;
  var rows = groupRowsByAnchor(bodyLines, anchorRe, 90, 6);

  rows.forEach(function (group) {
    var allItems = [];
    group.lines.forEach(function (l) { allItems = allItems.concat(l.items); });
    allItems.sort(function (a, b) { return b.y - a.y || a.x - b.x; });

    var skuToken = allItems.find(function (it) { return it.x <= 90 && anchorRe.test(it.text); });
    var sku = skuToken ? skuToken.text : null;

    var numTokens = allItems.filter(function (it) { return /^[\d,]+\.\d{2}$/.test(it.text); });
    var qtyToken = allItems.find(function (it) { return /^\d+\.\d{2}$/.test(it.text) && it.x > 400 && it.x < 460; });
    var uomToken = allItems.find(function (it) { return /^(EA|PC|CS|BOX|PK)$/i.test(it.text); });
    var discTypeToken = allItems.find(function (it) { return /^Percnt$/i.test(it.text); });
    var discValueToken = discTypeToken ? allItems.find(function (it) { return it.x > discTypeToken.x && it.y === discTypeToken.y && /^\d+(\.\d+)?$/.test(it.text); }) : null;
    var discountPct = discTypeToken && discValueToken ? parseFloat(discValueToken.text) : 0;

    // UOP COST (Vat Excl) sits ~x600-640, (Vat Incl) ~x660-700 on this layout.
    var uopExclToken = allItems.find(function (it) { return numTokens.includes(it) && it.x > 590 && it.x < 650; });
    var uopInclToken = allItems.find(function (it) { return numTokens.includes(it) && it.x > 650 && it.x < 710; });
    var uopExcl = uopExclToken ? parseMoney(uopExclToken.text) : null;
    var uopIncl = uopInclToken ? parseMoney(uopInclToken.text) : null;

    var excludeSet = [skuToken, qtyToken, uomToken, discTypeToken, discValueToken].concat(numTokens);
    var descTokens = allItems.filter(function (it) {
      return excludeSet.indexOf(it) < 0 && it.x > 90 && it.x < 400 &&
        !/^\d{10,14}$/.test(it.text) && it.text !== "-";
    });
    var description = descTokens.map(function (it) { return it.text; }).join(" ").trim();

    var qty = qtyToken ? parseFloat(qtyToken.text) : null;
    var price = uopIncl !== null ? round2(uopIncl * (1 - discountPct / 100)) : null;
    var totalAmount = qty !== null && price !== null ? round2(qty * price) : null;

    po.lines.push({
      sku: sku,
      description: description,
      uom: uomToken ? uomToken.text.toUpperCase() : "",
      qtyOriginal: qty,
      packSize: 1,
      qty: qty,
      price: price,
      totalAmount: totalAmount,
      discountPct: discountPct,
      caseConverted: false,
      printedLineTotal: null,
      _uopExcl: uopExcl,
      _uopIncl: uopIncl
    });
  });

  po.declaredLineCount = null;
  return po;
}

async function parseMetro(bytes, filename) {
  var loadingTask = pdfjsLib.getDocument({ data: bytes });
  var doc = await loadingTask.promise;
  if (doc.numPages < 1) return [];
  var page = await doc.getPage(1);
  var po = await parseMetroPage(page, filename);
  return [po];
}

// ---------------------------------------------------------------------------
// 4. THE D.I.Y SHOP CORP  (multi-PO SAP Business One PDF)
// ---------------------------------------------------------------------------

async function parseDiyPage(page) {
  var items = await pdfPageTextItems(page);
  var lines = groupIntoLines(items, 2.0);
  var fullText = linesToText(lines);

  var poMatch = items.find(function (it) { return it.x >= 440 && it.x <= 570 && it.y >= 640 && it.y <= 670 && /^\d{6,8}$/.test(it.text); });
  var poNumber = poMatch ? poMatch.text : null;

  var orderDateMatch = fullText.match(/Order Date:\s*(\d{2}\/\d{2}\/\d{4})/);
  var validUntilMatch = fullText.match(/Valid Until:\s*(\d{2}\/\d{2}\/\d{4})/);

  // Branch = the address line(s) printed under "For Branch"/the branch code,
  // right column (x 270-330), between the code row and "Terms".
  var branchLines = items.filter(function (it) {
    return it.x >= 270 && it.x <= 330 && it.y >= 682 && it.y <= 712 &&
      !/^(For Branch|Terms|DIY-?|0\d{3})$/i.test(it.text);
  });
  branchLines.sort(function (a, b) { return b.y - a.y; });
  var branch = branchLines.map(function (it) { return it.text; }).join(", ").replace(/\s+,/g, ",").trim();

  // Line prices print VAT-inclusive (matches Handyman's "VAT Inclusive with EWT"
  // convention) so the grand total to reconcile against is "Total:" (VAT-incl),
  // not "Total Before VAT:" -- confirmed: sum(qty*price) landed within a cent of
  // "Total:" and ~12% (the VAT rate) away from "Total Before VAT:" on every
  // sample checked.
  var grandTotalLine = findLineContaining(lines, /^Total:/);
  var totalMatch = grandTotalLine ? grandTotalLine.text.match(/([\d,]+\.\d{2})/) : null;
  var totalQtyMatch = fullText.match(/TOTAL QTY:\s*(\d+)/i);
  var continuedLine = findLineContaining(lines, /^Continue$/);

  // Item rows: restrict to the table body (below the column headers ~588, above
  // the page footer ~96); anchor = the lone item-number token at x<=45.
  var itemAreaLines = lines.filter(function (l) { return l.y < 588 && l.y > 96; });
  var anchorRe = /^\d{1,2}$/;
  var rows = groupRowsByAnchor(itemAreaLines, anchorRe, 45, 9);

  var lineItems = [];
  rows.forEach(function (group) {
    var allItems = [];
    group.lines.forEach(function (l) { allItems = allItems.concat(l.items); });

    var itemNoItem = allItems.find(function (it) { return it.x <= 45 && anchorRe.test(it.text); });
    if (!itemNoItem) return;
    var itemCodeItem = allItems.find(function (it) { return it.x > 45 && it.x <= 95; });

    var phpItems = allItems.filter(function (it) { return /^PHP\s*[\d,]+\.\d{2}$/.test(it.text); });
    var priceItem = phpItems.length > 0 ? phpItems[0] : null;
    var totalItem = phpItems.length > 1 ? phpItems[1] : null;

    var qtyItem = allItems.find(function (it) {
      return it !== itemNoItem && it !== itemCodeItem && /^\d{1,5}$/.test(it.text) && it.x > 340 && it.x < 400;
    });

    var uomItem = allItems.find(function (it) { return it.text === "PC"; });

    var descTokens = allItems.filter(function (it) {
      return it !== itemNoItem && it !== itemCodeItem && it !== priceItem && it !== totalItem &&
        it !== qtyItem && it !== uomItem && it.x > 45 && it.x < 340 &&
        !/^\d{1,3}$/.test(it.text);
    });
    descTokens.sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var description = descTokens.map(function (it) { return it.text; }).join(" ").replace(/\s+/g, " ").trim();

    var qty = qtyItem ? parseInt(qtyItem.text, 10) : null;
    var price = priceItem ? parseMoney(priceItem.text) : null;
    var printedLineTotal = totalItem ? parseMoney(totalItem.text) : null;
    var totalAmount = qty !== null && price !== null ? round2(qty * price) : null;

    lineItems.push({
      _itemNo: parseInt(itemNoItem.text, 10),
      sku: itemCodeItem ? itemCodeItem.text : null,
      description: description,
      uom: uomItem ? "PC" : "",
      qtyOriginal: qty,
      packSize: 1,
      qty: qty,
      price: price,
      totalAmount: totalAmount,
      discountPct: 0,
      caseConverted: false,
      printedLineTotal: printedLineTotal
    });
  });
  lineItems.sort(function (a, b) { return a._itemNo - b._itemNo; });

  return {
    poNumber: poNumber,
    date: orderDateMatch ? parseDateFlexible(orderDateMatch[1]) : null,
    expDate: validUntilMatch ? parseDateFlexible(validUntilMatch[1]) : null,
    branch: branch,
    vendorOk: GOOD_SHEPHERD_NAME_RE.test(fullText),
    printedTotal: totalMatch ? parseMoney(totalMatch[1]) : null,
    totalQtyDeclared: totalQtyMatch ? parseInt(totalQtyMatch[1], 10) : null,
    isFinalPageOfPo: !continuedLine,
    lineItems: lineItems
  };
}

async function parseDiy(bytes, filename) {
  var loadingTask = pdfjsLib.getDocument({ data: bytes });
  var doc = await loadingTask.promise;

  var poGroups = []; // [{ poNumber, pages:[pageData,...] }]
  for (var p = 1; p <= doc.numPages; p++) {
    var page = await doc.getPage(p);
    var pageData = await parseDiyPage(page);
    if (!pageData.poNumber) continue;
    var last = poGroups.length > 0 ? poGroups[poGroups.length - 1] : null;
    if (last && last.poNumber === pageData.poNumber) {
      last.pages.push(pageData);
    } else {
      poGroups.push({ poNumber: pageData.poNumber, pages: [pageData] });
    }
  }

  return poGroups.map(function (group) {
    var po = newPo("diy", filename);
    var first = group.pages[0];
    po.poNumber = group.poNumber;
    po.date = first.date;
    po.expDate = first.expDate;
    po.branch = first.branch;
    po.vendorOk = group.pages.some(function (pg) { return pg.vendorOk; });
    po.printedTotal = null;
    po.printedTotalLabel = "Total (VAT Incl)";
    var declaredQty = null;

    group.pages.forEach(function (pg) {
      pg.lineItems.forEach(function (li) {
        po.lines.push(li);
      });
      if (pg.printedTotal !== null) po.printedTotal = pg.printedTotal;
      if (pg.totalQtyDeclared !== null) declaredQty = pg.totalQtyDeclared;
    });

    po.declaredLineCount = null; // DIY declares total QTY, not line count
    po._declaredTotalQty = declaredQty;
    po.sourceFile = filename + " (PO " + group.poNumber + ")";
    return po;
  });
}

// ---------------------------------------------------------------------------
// SKU MASTER
// ---------------------------------------------------------------------------

function parseSkuMasterWorkbook(bytes) {
  var wb = XLSX.read(bytes, { type: "array" });
  var ws = wb.Sheets[wb.SheetNames[0]];
  var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  var header = rows[0].map(function (h) { return cleanText(String(h)).toLowerCase(); });
  var idx = {
    customer: header.indexOf("customer"),
    sku: header.indexOf("customer sku"),
    itemCode: header.indexOf("megabox item code"),
    description: header.indexOf("description"),
    dateAdded: header.indexOf("date added")
  };
  var entries = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r || r.every(function (c) { return cleanText(String(c)) === ""; })) continue;
    entries.push({
      customer: cleanText(String(r[idx.customer] || "")),
      sku: cleanText(String(r[idx.sku] || "")),
      itemCode: cleanText(String(r[idx.itemCode] || "")),
      description: cleanText(String(r[idx.description] || "")),
      dateAdded: cleanText(String(r[idx.dateAdded] || ""))
    });
  }
  return entries;
}

function skuMasterKey(customerLabel, sku) {
  return customerLabel.toUpperCase() + "||" + sku.trim();
}

function buildSkuMasterIndex(entries) {
  var map = {};
  var newestDate = null;
  entries.forEach(function (e) {
    map[skuMasterKey(e.customer, e.sku)] = e;
    var d = parseDateFlexible(e.dateAdded) || parseMMDDYYYY(e.dateAdded);
    if (d) {
      var js = dateToJs(d);
      if (!newestDate || js > newestDate) newestDate = js;
    }
  });
  return { map: map, entries: entries, newestDate: newestDate };
}

function parseMMDDYYYY(s) {
  var m = cleanText(String(s)).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? { y: +m[3], m: +m[1], d: +m[2] } : null;
}

// ---------------------------------------------------------------------------
// VALIDATION ENGINE (§6) — customer-agnostic, operates on already-parsed POs
// ---------------------------------------------------------------------------

var PO_NUMBER_SHAPE = {
  handyman: /^\d{10}$/,
  diy: /^\d{6,8}$/,
  puregold: /^\d{6,12}-\d{2}$/,
  metro: /^\d{6,9}$/
};

function addReason(po, severity, rule, message) {
  po._reasons.push({ severity: severity, rule: rule, message: message });
}

function validateBatch(customerKey, pos, opts) {
  opts = opts || {};
  var now = opts.now || new Date();
  var skuIndex = opts.skuMasterIndex || null; // {map, entries, newestDate} or null
  var customer = CUSTOMERS[customerKey];

  pos.forEach(function (po) { po._reasons = []; po._blocked = false; });

  // Batch-level: duplicate PO#+SKU
  var seen = {};
  pos.forEach(function (po) {
    if (!po.poNumber) return;
    po.lines.forEach(function (line) {
      if (!line.sku) return;
      var key = po.poNumber + "|" + line.sku;
      seen[key] = seen[key] || [];
      seen[key].push(po);
    });
  });
  Object.keys(seen).forEach(function (key) {
    if (seen[key].length > 1) {
      seen[key].forEach(function (po) {
        addReason(po, "BLOCK", 8, "Duplicate PO#+SKU in this batch: " + key + " appears " + seen[key].length + " times");
      });
    }
  });

  // Batch-level: same SKU, price varies >20% across the batch (REVIEW rule 3)
  var priceBySku = {};
  pos.forEach(function (po) {
    po.lines.forEach(function (line) {
      if (!line.sku || line.price === null || line.price === undefined) return;
      priceBySku[line.sku] = priceBySku[line.sku] || [];
      priceBySku[line.sku].push({ po: po, price: line.price, sku: line.sku });
    });
  });
  Object.keys(priceBySku).forEach(function (sku) {
    var entries = priceBySku[sku];
    if (entries.length < 2) return;
    var prices = entries.map(function (e) { return e.price; });
    var min = Math.min.apply(null, prices);
    var max = Math.max.apply(null, prices);
    if (min > 0 && (max - min) / min > 0.20) {
      var poSet = new Set(entries.map(function (e) { return e.po; }));
      poSet.forEach(function (po) {
        addReason(po, "REVIEW", 3, "SKU " + sku + " price varies >20% across this batch (₱" + min.toFixed(2) + " - ₱" + max.toFixed(2) + ")");
      });
    }
  });

  // SKU master staleness applies to the whole batch
  var skuMasterMissing = !skuIndex;
  var skuMasterStale = skuIndex && skuIndex.newestDate ? daysBetween(skuIndex.newestDate, now) > 60 : false;

  pos.forEach(function (po) {
    // --- structural parse failures -> BLOCK outright
    if (po.parseErrors.length > 0) {
      po.parseErrors.forEach(function (e) { addReason(po, "BLOCK", "parse", e); });
    }

    // BLOCK 1: PO# shape
    var shapeRe = PO_NUMBER_SHAPE[customerKey];
    if (!po.poNumber) {
      addReason(po, "BLOCK", 1, "PO# missing");
    } else if (shapeRe && !shapeRe.test(po.poNumber)) {
      addReason(po, "BLOCK", 1, "PO# '" + po.poNumber + "' does not match the expected format for " + customer.label);
    }

    // BLOCK 6: vendor
    if (!po.vendorOk) {
      addReason(po, "BLOCK", 6, "Vendor on this PO is not identifiable as Good Shepherd Manufacturing Corp / Able & Good Shepherd Marketplace Corp");
    }

    // BLOCK 2/5: line-level required fields, qty/price sanity
    if (po.lines.length === 0) {
      addReason(po, "BLOCK", 2, "No line items could be extracted from this PO");
    }
    po.lines.forEach(function (line, i) {
      var n = i + 1;
      if (!line.sku) addReason(po, "BLOCK", 2, "Line " + n + ": missing SKU");
      if (!line.description) addReason(po, "BLOCK", 2, "Line " + n + " (SKU " + line.sku + "): missing description");
      if (line.qty === null || line.qty === undefined) addReason(po, "BLOCK", 2, "Line " + n + " (SKU " + line.sku + "): missing quantity");
      if (line.price === null || line.price === undefined) addReason(po, "BLOCK", 2, "Line " + n + " (SKU " + line.sku + "): missing unit price");
      if (line.qty !== null && line.qty !== undefined && line.qty <= 0) addReason(po, "BLOCK", 5, "Line " + n + " (SKU " + line.sku + "): quantity " + line.qty + " is <= 0");
      if (line.price !== null && line.price !== undefined && line.price < 0) addReason(po, "BLOCK", 5, "Line " + n + " (SKU " + line.sku + "): negative price " + line.price);
      if (line.qty !== null && line.qty !== undefined && line.qty > 5000) addReason(po, "REVIEW", 7, "Line " + n + " (SKU " + line.sku + "): quantity " + line.qty + " is unusually large (>5,000 pcs) — check for a decimal misread");
    });

    // BLOCK 3: declared line count
    if (po.declaredLineCount !== null && po.declaredLineCount !== po.lines.length) {
      addReason(po, "BLOCK", 3, "PO declares " + po.declaredLineCount + " line item(s) but " + po.lines.length + " were parsed");
    }

    // BLOCK 4: total reconciliation
    if (po.printedTotal === null || po.printedTotal === undefined) {
      addReason(po, "BLOCK", 4, "Could not find a printed document total to reconcile against");
    } else {
      var sum = po.lines.reduce(function (acc, l) { return acc + (l.totalAmount || 0); }, 0);
      sum = round2(sum);
      var diff = Math.abs(sum - po.printedTotal);
      if (diff > 0.05) {
        addReason(po, "BLOCK", 4, "Printed total ₱" + po.printedTotal.toFixed(2) + " (" + (po.printedTotalLabel || "") + ") does not match the sum of computed line totals ₱" + sum.toFixed(2) + " (diff ₱" + diff.toFixed(2) + ")");
      }
    }

    // BLOCK 7: date sanity
    if (!po.date) {
      addReason(po, "BLOCK", 7, "PO date is missing or unparseable");
    } else {
      var jsDate = dateToJs(po.date);
      if (Math.abs(daysBetween(jsDate, now)) > 180) {
        addReason(po, "BLOCK", 7, "PO date " + dateToMMDDYYYY(po.date) + " is more than 180 days from today — check for a misread field");
      }
    }

    // Puregold-specific: subtotal must reconcile BEFORE case->piece conversion (§5.2)
    if (customerKey === "puregold" && po._subTotalQty !== null && po._subTotalQty !== undefined) {
      var rawSum = po.lines.reduce(function (acc, l) { return acc + (l._rawQtyForSubtotal || 0); }, 0);
      rawSum = round2(rawSum);
      if (Math.abs(rawSum - po._subTotalQty) > 0.01) {
        addReason(po, "BLOCK", 4, "Printed Sub Total quantity " + po._subTotalQty + " does not match the sum of as-printed line quantities " + rawSum + " (checked BEFORE case-to-piece conversion)");
      }
    }

    // --- REVIEW rules ---
    po.lines.forEach(function (line, i) {
      var n = i + 1;
      if (line.caseConverted) {
        addReason(po, "REVIEW", 1, "Line " + n + " (SKU " + line.sku + "): case-to-piece conversion applied (" + line.qtyOriginal + " case(s) x " + line.packSize + " = " + line.qty + " pcs)");
        if (line.qty !== null && line.packSize) {
          var rawUnit = line.price !== null ? line.price * line.packSize : null;
        }
        // rule 2: price/pack doesn't divide to 2dp cleanly
        if (line.price !== null && line.packSize > 0) {
          var exact = (Math.round((line.price * line.packSize) * 10000) / 10000);
          var rounded = round2(exact);
          if (Math.abs(exact - rounded) > 0.001) {
            addReason(po, "REVIEW", 2, "Line " + n + " (SKU " + line.sku + "): price ÷ pack size does not divide to exactly 2 decimals");
          }
        }
      }
      if (line.discountPct && line.discountPct > 0) {
        addReason(po, "REVIEW", 4, "Line " + n + " (SKU " + line.sku + "): line-level discount of " + line.discountPct + "% applied");
      }
    });

    if (customerKey === "metro") {
      addReason(po, "REVIEW", 5, "Metro net-price calculation not yet ERP-verified (see §5.3 Test A) — every Metro PO is REVIEW until that trial import passes");
    }

    if (skuMasterMissing) {
      addReason(po, "REVIEW", "5b", "SKU list not loaded — cannot verify item codes");
    } else if (skuMasterStale) {
      addReason(po, "REVIEW", "5b", "SKU list last updated " + (skuIndex.newestDate ? dateToMMDDYYYY({ y: skuIndex.newestDate.getUTCFullYear(), m: skuIndex.newestDate.getUTCMonth() + 1, d: skuIndex.newestDate.getUTCDate() }) : "unknown") + " — more than 60 days ago");
    }

    if (skuIndex) {
      po.lines.forEach(function (line, i) {
        var n = i + 1;
        if (!line.sku) return;
        var entry = skuIndex.map[skuMasterKey(customer.legalName, line.sku)];
        if (!entry) {
          addReason(po, "BLOCK", 0, "Line " + n + ": Unknown SKU " + line.sku + " — not in SKU list");
        } else if (entry.description && line.description) {
          var overlap = tokenOverlapRatio(entry.description, line.description);
          if (overlap < 0.4) {
            addReason(po, "REVIEW", "5c", "Line " + n + " (SKU " + line.sku + "): PO description '" + line.description + "' differs materially from SKU list description '" + entry.description + "'");
          }
        }
      });
    }

    if (!po.expDate) {
      addReason(po, "REVIEW", 6, "PO expiry date (P.O Date Exp) is missing");
    } else {
      var expJs = dateToJs(po.expDate);
      if (expJs.getTime() < now.getTime()) {
        addReason(po, "REVIEW", 6, "PO expiry date " + dateToMMDDYYYY(po.expDate) + " is in the past");
      }
    }

    // Final status
    var hasBlock = po._reasons.some(function (r) { return r.severity === "BLOCK"; });
    var hasReview = po._reasons.some(function (r) { return r.severity === "REVIEW"; });
    po.status = hasBlock ? "BLOCKED" : (hasReview ? "REVIEW" : "READY");
    po.reasons = po._reasons;
  });

  return pos;
}

// ---------------------------------------------------------------------------
// EXPORT BUILDER (§3)
// ---------------------------------------------------------------------------

function buildExportRows(pos, customerKey) {
  var customer = CUSTOMERS[customerKey];
  var rows = [];
  pos.forEach(function (po) {
    if (po.status === "BLOCKED") return;
    po.lines.forEach(function (line) {
      rows.push([
        po.poNumber || "",
        customer.legalName,
        po.branch || "",
        po.date ? dateToMMDDYYYY(po.date) : "",
        line.sku || "",
        "",
        line.description || "",
        line.qty,
        line.price,
        "", "", "",
        line.totalAmount,
        "",
        po.expDate ? dateToMMDDYYYY(po.expDate) : "",
        po.remarks || ""
      ]);
    });
  });
  return rows;
}

function buildWorkbook(pos, customerKey) {
  var rows = buildExportRows(pos, customerKey);
  var wb = XLSX.utils.book_new();
  var ws = {};
  var range = { s: { r: 0, c: 0 }, e: { r: rows.length, c: EXPORT_HEADERS.length - 1 } };

  EXPORT_HEADERS.forEach(function (h, c) {
    ws[XLSX.utils.encode_cell({ r: 0, c: c })] = { t: "s", v: h };
  });

  var textCols = [0, 4]; // PO#, SKU
  var intCols = [7]; // Order Qty.
  var moneyCols = [8, 12]; // Price, Total Amount

  rows.forEach(function (row, r) {
    row.forEach(function (val, c) {
      var addr = XLSX.utils.encode_cell({ r: r + 1, c: c });
      if (val === "" || val === null || val === undefined) {
        ws[addr] = { t: "s", v: "" };
        return;
      }
      if (textCols.indexOf(c) >= 0) {
        ws[addr] = { t: "s", v: String(val), z: "@" };
      } else if (intCols.indexOf(c) >= 0) {
        ws[addr] = { t: "n", v: val, z: "#,##0" };
      } else if (moneyCols.indexOf(c) >= 0) {
        ws[addr] = { t: "n", v: val, z: "#,##0.00" };
      } else {
        ws[addr] = { t: "s", v: String(val) };
      }
    });
  });

  ws["!ref"] = XLSX.utils.encode_range(range);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return wb;
}

function exportFileName(customerKey, when) {
  when = when || new Date();
  var pad = function (n) { return String(n).padStart(2, "0"); };
  var stamp = when.getFullYear() + pad(when.getMonth() + 1) + pad(when.getDate()) + "-" + pad(when.getHours()) + pad(when.getMinutes());
  return "SO_" + CUSTOMERS[customerKey].label.replace(/\s+/g, "") + "_" + stamp + ".xlsx";
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

var POSO = {
  CUSTOMERS: CUSTOMERS,
  EXPORT_HEADERS: EXPORT_HEADERS,
  parseHandyman: parseHandyman,
  parsePuregold: parsePuregold,
  parseMetro: parseMetro,
  parseDiy: parseDiy,
  parseSkuMasterWorkbook: parseSkuMasterWorkbook,
  buildSkuMasterIndex: buildSkuMasterIndex,
  validateBatch: validateBatch,
  buildExportRows: buildExportRows,
  buildWorkbook: buildWorkbook,
  exportFileName: exportFileName,
  _internal: { parseMoney: parseMoney, parseDateFlexible: parseDateFlexible, dateToMMDDYYYY: dateToMMDDYYYY, round2: round2 }
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = POSO;
}
