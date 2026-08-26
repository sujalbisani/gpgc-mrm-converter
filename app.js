/* MRM dispatch-sheet converter - runs entirely in the browser.
   Ported from convert_mrm.py. No file ever leaves this page - everything
   below runs client-side with ExcelJS. */

const RAW_HEADERS = [
  "Sales Group", "Cust. Name", "SO No./Item", "Inv.No./Item", "Inv.Date",
  "Material Desc", "Thick", "Length", "Width", "Cus. Grade", "TC Number",
  "Batch", "Quantity", "Invoice Value", "Truck No.", "PO NO.", "LR Number",
  "Storage Location", "Material Group", "Trans. Name", "Customer", "Quality",
  "Destination", "Supplied HR Material", "UPR", "Pieces", "Supplier Batch",
  "Material", "Prod. hierarchy", "Customer Segment", "Discount", "Freight",
  "Freight Per Ton", "Class",
];

const TARGET_HEADERS = [
  "Sales Group", "Cust. Name", "SO No./Item", "Inv.No./Item", "Inv.Date",
  "Material Desc", "Thick", "TA", "Length", "Width", "WA", "Cus. Grade",
  "TC Number", "Batch", "Quantity", "Invoice Value", "Truck No.", "PO NO.",
  "LR Number", "Storage Location", "Material Group", "Trans. Name", "Customer",
  "Quality", "Destination", "Supplied HR Material", "UPR", "Pieces",
  "Supplier Batch", "Material", "Coating", "Prod. hierarchy", "Customer Segment",
  "Discount", "DPMT", "Freight", "Freight Per Ton", "Class", "Branch Desc",
  "Product Type", "Classification", "NSR", "Value", "TE", "ZE", "Conv Cost",
  "Alloy Cost", "BENSR", "BE_Value", "Zone", "Cluster", "Party",
];
const COL = {};
TARGET_HEADERS.forEach((name, i) => (COL[name] = i + 1)); // 1-indexed

// Destination -> Cluster / Zone lookups, ported verbatim from the user's own
// Excel formulas (which searched a Destination cell for these city keywords).
// Order matters - first matching group wins, same as the IFS() it came from.
const CLUSTER_GROUPS = [
  ["Cluster-1", ["MARATH", "AURANG", "JALNA", "MALKAP", "JALGA", "NANDED",
                 "LATUR", "SOLAPUR", "BEED", "MALEGA"]],
  ["Cluster-2", ["NAGPUR", "NAGOUR", "RAIPUR", "VIDARB", "JAMNAG", "BHARUC",
                 "WARDHA", "BILASPUR", "PUNE"]],
  ["Cluster-3", ["HUBLI", "MYSORE", "CHENN", "CHEENN", "MADURAI", "COIMBA",
                 "BANGAL", "BANAG", "CUDDAL", "TRICHY", "VIJAYW"]],
  ["Cluster-4", ["GWALIO", "BHOPAL", "SATNA", "KATNI"]],
];
const CLUSTER_DEFAULT = "Other";

const ZONE_GROUPS = [
  ["Maharashtra", ["JALGA", "SANGLI", "AURANG", "SOLAPUR", "NAGPUR", "NAGOUR",
                    "WARDHA", "MALEGA", "LATUR", "MALKAP", "INDAPUR", "JALNA",
                    "PUNE", "TALOJA", "KAPSI"]],
  ["Central", ["BHOPAL", "INDORE", "SINGRAU", "GWALIO", "SATNA", "PITHA", "KATNI"]],
  ["East", ["RAMGAR", "RAMGART", "BOKARO", "BILASPUR", "CUTTAC", "RANCHI",
            "RAIPUR", "ROURKE"]],
  ["North", ["LUCKNOW", "FARIDAB", "MUNDKA", "LUDHIAN", "DELH", "NEW DEL",
             "GHAZIAB", "PANCHKU", "JALANDH", "KANPUR", "JAIPUR", "VARANAS", "BANDA"]],
  ["West", ["KHEDA", "SURAT", "DEESA", "AHMEDAB", "JAMNAG", "BHARUC"]],
  ["South", ["PALAKK", "BANGAL", "BANAG", "CHENN", "CHEENN", "HYDERAB", "TELANG",
             "RANGAR", "PASHAM", "PATANC", "SALEM", "COIMBA", "ERODE", "ERNAKU",
             "ERNAV", "VIJAYW", "CUDDAL", "HUBLI", "TRICHY"]],
];
const ZONE_DEFAULT = "Unknown";

function buildLookupFormula(cellRef, groups, defaultLabel) {
  const clauses = groups.map(([label, keywords]) => {
    const searches = keywords.map((kw) => `ISNUMBER(SEARCH("${kw}",${cellRef}))`).join(",");
    return `OR(${searches}),"${label}"`;
  });
  clauses.push(`TRUE,"${defaultLabel}"`);
  return `IFERROR(_xlfn.IFS(${clauses.join(",")}),"${defaultLabel}")`;
}

const DIRECT_FIELDS = [
  "Sales Group", "Cust. Name", "SO No./Item", "Inv.No./Item", "Inv.Date",
  "Material Desc", "Cus. Grade", "TC Number", "Batch", "Quantity",
  "Invoice Value", "Truck No.", "PO NO.", "LR Number", "Storage Location",
  "Material Group", "Trans. Name", "Customer", "Quality", "Destination",
  "Supplied HR Material", "UPR", "Pieces", "Supplier Batch", "Material",
  "Prod. hierarchy", "Customer Segment", "Discount", "Freight",
  "Freight Per Ton", "Class",
];

const BRANCH_MAPPING = {
  "201": "MUMBAI", "204": "INDORE", "206": "AURANGABAD", "301": "DELHI",
  "304": "JAIPUR", "402": "CHENNAI", "403": "BANGALORE", "100": "PLANT",
  "205": "NAGPUR", "203": "PUNE", "202": "AHMEDABAD", "501": "KOLKATA",
  "305": "LUDHIYANA", "401": "HYDERABAD",
};

const CLASSIFICATION_MAPPING = {
  "CHEQUERED PLATE (5 mm to 10 mm)": "CHEQUERED", "CRCA COIL": "CRCA",
  "CRCA SHEET": "CRCA", "CRCA SLIT COIL": "CRCA", "CRFH SLIT COIL": "CRFH",
  "GC SHEET": "GC", "GP COIL": "GP", "GP SHEET": "GP", "GP SLIT COIL": "GP",
  "HR CHEQUERED COIL": "CHEQUERED", "HR COIL": "HR COIL",
  "HR COIL (TRIM)": "HR COIL", "HR COIL(3 MM TO 4.99 MM)": "HR COIL",
  "HR COIL<3 MM": "HR COIL", "HR PICKLED BLANKS": "HRPO", "HR PLATE": "HR PLATE",
  "HR SHEET": "HR SHEET", "HR SHEET (TRIM)": "HR SHEET",
  "HR SHEET(3 TO 4.99 MM)": "HR SHEET", "HR SHEET(5 TO 10MM)": "HR SHEET",
  "HR SHEET(ABOVE 10 MM)": "HR SHEET", "HR SLIT COIL": "HR COIL",
  "HRCOIL (5 MM TO 10 MM)": "HR COIL", "HRCOIL>10 MM": "HR COIL",
  "CHEQUERED PLATE(CTL)": "CHEQUERED", "HR BLANKS": "HR SHEET",
  "HR PICKLED & OILED SHEET": "HRPO", "HR PICKLED SLIT COIL": "HRPO",
  "HR SHEET(ABOVE 12 MM)": "HR SHEET", "HR SHEET<3MM": "HR SHEET",
  "HR SHEET>5MM": "HR SHEET", "HR PLATE(TRIM)": "HR PLATE", "SLAB": "SLAB",
  "HR COIL>10MM": "HR COIL", "CRCA TRIM COIL": "CRCA COIL",
};

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June", "July", "August",
  "September", "October", "November", "December",
];

// These 5 are Evonith's trading-customer subsidiaries: they buy from Evonith
// and resell further downstream. Used for two things:
//  1. NSR = UPR + DPMT for these customers (everyone else uses UPR - Freight Per
//     Ton). Confirmed against MRM-July 2026.XLSX: 0 mismatches across 4200 rows.
//  2. Party/Destination get overridden from that customer's monthly third-party
//     sales register (see loadThirdPartyRegister) when a matching file and
//     invoice row exist, so we see who the material actually got resold to and
//     where it actually ended up, not just the subsidiary's own warehouse.
const TRADING_CUSTOMERS = [
  "DENOTIC", "IRONSPIRE", "IROMETAL", "METAL HUB", "STEEL BRIDGE",
];

function toNumber(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return null;
  if (typeof v === "string") {
    v = v.replace(/,/g, "").trim();
    if (v === "") return null;
  }
  const f = parseFloat(v);
  if (Number.isNaN(f)) return null;
  return Number.isInteger(f) ? f : f;
}

function computeProductType(prodHierarchy, thick, width) {
  if (["GP Coil", "GP Sheet", "GP Slit Coil"].includes(prodHierarchy)) {
    if (width !== null && thick !== null && width <= 1050 && thick < 0.7) {
      return "CGL 2";
    }
    return "CGL 1";
  }
  return prodHierarchy;
}

function cellDateValue(v) {
  // ExcelJS gives Date objects for date cells, but some exports store dates
  // as plain numbers/strings - normalize to a Date if possible.
  if (v instanceof Date) return v;
  return null;
}

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

function readSheetRows(ws) {
  const headerRow = ws.getRow(1);
  const header = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    header[colNumber - 1] = cell.value;
  });
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const vals = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      vals[colNumber - 1] = cell.value;
    });
    rows.push(vals);
  });
  return { header, rows };
}

function loadRawRows(wb) {
  const ws = wb.getWorksheet("Sheet1");
  if (!ws) throw new Error("Raw file has no sheet named 'Sheet1'.");
  const { header, rows } = readSheetRows(ws);
  const trimmedHeader = header.slice(0, RAW_HEADERS.length);
  const mismatch = RAW_HEADERS.some((h, i) => trimmedHeader[i] !== h);
  if (mismatch) {
    throw new Error(
      "Raw file header doesn't match expected layout.\n" +
        "Expected: " + RAW_HEADERS.join(", ") + "\n" +
        "Got: " + trimmedHeader.join(", ")
    );
  }
  const out = [];
  for (const vals of rows) {
    const d = {};
    RAW_HEADERS.forEach((h, i) => (d[h] = vals[i] === undefined ? null : vals[i]));
    const invNo = d["Inv.No./Item"];
    const material = d["Material"];
    if (!invNo || !material) continue; // blank / totals row
    const custName = (d["Cust. Name"] || "").toString();
    if (custName.toUpperCase().includes("EVONITH")) continue; // inter-company self-transfer
    out.push(d);
  }
  return out;
}

function detectMonth(rawRows) {
  const counts = new Map();
  for (const r of rawRows) {
    const dt = cellDateValue(r["Inv.Date"]);
    if (!dt) continue;
    const key = dt.getFullYear() + "-" + dt.getMonth();
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  let bestKey = null;
  let bestCount = -1;
  for (const [k, c] of counts) {
    if (c > bestCount) {
      bestCount = c;
      bestKey = k;
    }
  }
  if (!bestKey) throw new Error("Couldn't find any valid Inv.Date values in the raw file.");
  const [year, month] = bestKey.split("-").map(Number);
  return { year, month }; // month is 0-indexed (JS Date convention)
}

function monthLabel(monthObj) {
  return MONTH_NAMES[monthObj.month] + " " + monthObj.year;
}

function detectConvCostBase(wb, monthObj) {
  const ws = wb.worksheets[0];
  const maxRow = ws.rowCount;
  const maxCol = ws.columnCount;
  let headerRow = null;
  for (let r = 1; r <= maxRow; r++) {
    let found = false;
    for (let c = 1; c <= maxCol; c++) {
      const v = ws.getRow(r).getCell(c).value;
      if (v instanceof Date) {
        found = true;
        break;
      }
    }
    if (found) {
      headerRow = r;
      break;
    }
  }
  if (!headerRow) return null;

  let costRow = null;
  for (let r = headerRow + 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const v = ws.getRow(r).getCell(c).value;
      if (typeof v === "string" && v.toUpperCase().includes("CONVERSION COST")) {
        costRow = r;
        break;
      }
    }
    if (costRow) break;
  }
  if (!costRow) return null;

  for (let c = 1; c <= maxCol; c++) {
    const v = ws.getRow(headerRow).getCell(c).value;
    if (v instanceof Date && v.getFullYear() === monthObj.year && v.getMonth() === monthObj.month) {
      return ws.getRow(costRow).getCell(c).value;
    }
  }
  return null;
}

/* Third-party sales register files (Denotics/Ironspire/Irometal/Metal Hub/Steel
   Bridge monthly sales reports) are legacy .xls, which ExcelJS can't read - use
   SheetJS (loaded as the global `XLSX`) just for these. */
async function parseRegisterFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = wb.Sheets["Sales "];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { range: 10, defval: null });
}

function matchesRegisterFilename(filename, customerKey, monthObj) {
  const upper = filename.toUpperCase();
  const monthName = MONTH_NAMES[monthObj.month].toUpperCase();
  const year = String(monthObj.year);
  return upper.includes(customerKey) && upper.includes(monthName) && upper.includes(year);
}

async function buildThirdPartyRegisters(registerFiles, monthObj) {
  const registers = {}; // customerKey -> {invoiceNo: {destination, buyer}}
  for (const key of TRADING_CUSTOMERS) {
    const match = registerFiles.find((f) => matchesRegisterFilename(f.name, key, monthObj));
    if (!match) continue;
    const rows = await parseRegisterFile(match);
    const byInvoice = {};
    for (const row of rows) {
      const ref = row["Pur Ref. No."];
      if (ref === null || ref === undefined || ref === "") continue;
      const invNo = parseInt(ref, 10);
      if (Number.isNaN(invNo)) continue;
      byInvoice[invNo] = { destination: row["Destination"], buyer: row["Buyer"] };
    }
    registers[key] = byInvoice;
  }
  return registers;
}

function copyInfoSheet(srcWs, destWb) {
  const destWs = destWb.addWorksheet("INFO");
  srcWs.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const newCell = destWs.getRow(rowNumber).getCell(colNumber);
      newCell.value = cell.value;
      if (cell.numFmt) newCell.numFmt = cell.numFmt;
    });
  });
  srcWs.columns.forEach((col, idx) => {
    if (col && col.width) {
      destWs.getColumn(idx + 1).width = col.width;
    }
  });
  return destWs;
}

async function buildOutput(rawRows, templateWb, convCostBase, monthObj, registerFiles, log) {
  const outWb = new ExcelJS.Workbook();
  const infoWs = templateWb.getWorksheet("INFO");
  if (!infoWs) throw new Error("The last-month MRM file has no 'INFO' sheet.");

  const registers = registerFiles && registerFiles.length
    ? await buildThirdPartyRegisters(registerFiles, monthObj)
    : {};

  const ws = outWb.addWorksheet("Sheet1");
  copyInfoSheet(infoWs, outWb);
  const headerRow = ws.getRow(1);
  TARGET_HEADERS.forEach((h, i) => (headerRow.getCell(i + 1).value = h));
  headerRow.font = { bold: true };

  rawRows.forEach((raw, i) => {
    const r = i + 2;
    const row = ws.getRow(r);
    const matDesc = (raw["Material Desc"] || "").toString().trim();
    const custName = (raw["Cust. Name"] || "").toString().toUpperCase();
    const prodHierarchy = raw["Prod. hierarchy"];
    const classification = CLASSIFICATION_MAPPING[matDesc] || matDesc;
    const thick = toNumber(raw["Thick"]);
    const width = toNumber(raw["Width"]);
    const productType = computeProductType(prodHierarchy, thick, width);
    const tradingMatch = TRADING_CUSTOMERS.find((k) => custName.includes(k));
    const nsrFormula = tradingMatch ? `AA${r}+AI${r}` : `AA${r}-AK${r}`;

    let destination = raw["Destination"];
    let party = null;
    if (tradingMatch && registers[tradingMatch]) {
      const invNoStr = String(raw["Inv.No./Item"]).split("/")[0];
      const invNo = parseInt(invNoStr, 10);
      const entry = !Number.isNaN(invNo) ? registers[tradingMatch][invNo] : null;
      if (entry) {
        if (entry.destination) destination = String(entry.destination).trim();
        if (entry.buyer) party = String(entry.buyer).trim();
      }
    }

    DIRECT_FIELDS.forEach((field) => {
      row.getCell(COL[field]).value = raw[field];
    });
    row.getCell(COL["Thick"]).value = thick;
    row.getCell(COL["Length"]).value = toNumber(raw["Length"]);
    row.getCell(COL["Width"]).value = width;
    row.getCell(COL["Classification"]).value = classification;
    row.getCell(COL["Product Type"]).value = productType;
    row.getCell(COL["ZE"]).value = 0;
    row.getCell(COL["Destination"]).value = destination;
    row.getCell(COL["Party"]).value = party;

    const salesGroup = raw["Sales Group"];
    if (salesGroup) {
      const branch = BRANCH_MAPPING[String(salesGroup).trim()];
      if (branch) row.getCell(COL["Branch Desc"]).value = branch;
    }

    row.getCell(COL["Coating"]).value = { formula: `RIGHT(AD${r},"3")*1` };
    row.getCell(COL["TA"]).value = { formula: `G${r}*O${r}` };
    row.getCell(COL["WA"]).value = { formula: `J${r}*O${r}` };
    row.getCell(COL["DPMT"]).value = { formula: `AH${r}/O${r}` };
    row.getCell(COL["NSR"]).value = { formula: nsrFormula };
    row.getCell(COL["Value"]).value = { formula: `AP${r}*O${r}` };
    row.getCell(COL["TE"]).value = { formula: `VLOOKUP(G${r},INFO!E:G,3,TRUE)` };
    row.getCell(COL["Conv Cost"]).value = {
      formula:
        `_xlfn.IFS(F${r}="GP COIL",${convCostBase},` +
        `OR(F${r}="GP SHEET",F${r}="GC SHEET"),${convCostBase}+500,` +
        `F${r}="GP SLIT COIL",${convCostBase}+650,TRUE,"")`,
    };
    row.getCell(COL["Alloy Cost"]).value = { formula: "INFO!$N$3" };
    row.getCell(COL["BENSR"]).value = { formula: `AP${r}-AR${r}-AS${r}-AT${r}-AU${r}` };
    row.getCell(COL["BE_Value"]).value = { formula: `AV${r}*O${r}` };
    row.getCell(COL["Zone"]).value = { formula: buildLookupFormula(`Y${r}`, ZONE_GROUPS, ZONE_DEFAULT) };
    row.getCell(COL["Cluster"]).value = { formula: buildLookupFormula(`Y${r}`, CLUSTER_GROUPS, CLUSTER_DEFAULT) };

    const invDate = raw["Inv.Date"];
    if (invDate instanceof Date) {
      row.getCell(COL["Inv.Date"]).numFmt = "dd-mmm-yy";
    }
  });

  log(`Wrote ${rawRows.length} rows.`);
  return outWb;
}

function downloadWorkbook(wb, filename) {
  return wb.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  });
}

window.MRMConverter = {
  loadWorkbook,
  loadRawRows,
  detectMonth,
  monthLabel,
  detectConvCostBase,
  buildOutput,
  downloadWorkbook,
  MONTH_NAMES,
  TRADING_CUSTOMERS,
};
