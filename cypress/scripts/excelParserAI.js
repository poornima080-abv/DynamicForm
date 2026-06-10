const XLSX    = require("xlsx");
const fs      = require("fs");
const path    = require("path");
const ExcelJS = require("exceljs");

console.log("excelParserAI.js loaded");

// ======================= TYPE MAP =======================
const TYPE_MAP = {
  "Text":                        "string",
  "Number":                      "string",
  "Email":                       "string",
  "Multi Lines of Text":         "textarea",
  "Textarea":                    "textarea",
  "Selection":                   "select",
  "Dropdown":                    "select",
  "Select":                      "select",
  "Select Date / Auto Populate": "date",
  "Date":                        "date",
  "Fixed":                       "disableString",
};

const MANDATORY_ONLY_PATTERNS = [
  /this field to come up as mandatory/i,
  /come up as mandatory if/i,
  /come up as mandatory when/i,
  /will be mandatory when/i,
  /mandatory on publish form/i,
];

// ======================= HELPERS =======================
function normalizePunct(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[/\\.,\-"'=]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isActualRule(text) {
  if (!text || text.trim() === "") return false;
  const t = text.toLowerCase();
  return (
    t.includes("dependent upon")       ||
    t.includes("open this field")      ||
    t.includes("mandatory when")       ||
    t.includes("don't show")           ||
    t.includes("do not show")          ||
    t.includes("disable this")         ||
    t.includes("show this field")      ||
    t.includes("come up as mandatory") ||
    t.includes("will be shown")        ||
    t.includes("will be mandatory")    ||
    t.includes("if selected")          ||
    t.includes("if selection")         ||
    t.includes("selected as")          ||
    t.includes("publish form")         ||
    t.includes("request form")         ||
    t.includes("opened up when")       ||
    t.includes("shown when")
  );
}

function fuzzyMatch(a, b) {
  const al = String(a || "").toLowerCase().trim();
  const bl = String(b || "").toLowerCase().trim();
  if (!al || !bl) return false;
  if (al === bl) return true;

  const anorm = normalizePunct(al);
  const bnorm = normalizePunct(bl);
  if (anorm === bnorm) return true;

  const poNorm = s => s
    .replace(/\bpo\s*no\.?\s*[\/]?\s*capex\s*no\.?\b/gi, "po capex field")
    .replace(/\bpo\s*no\.?\b/gi,    "po number")
    .replace(/\bcapex\s*no\.?\b/gi, "capex number");

  const ap = poNorm(anorm);
  const bp = poNorm(bnorm);
  if (ap === bp) return true;

  if (ap.includes(bp) || bp.includes(ap)) {
    const aW = ap.split(/\s+/).filter(w => w.length > 2);
    const bW = bp.split(/\s+/).filter(w => w.length > 2);
    const sh = aW.length <= bW.length ? aW : bW;
    const lo = aW.length <= bW.length ? bW : aW;
    if (lo.length > sh.length + 1) return false;
    return true;
  }

  // bidirectional word match — fixes PO No / Capex No cross-match
  const aWords = anorm.split(/\s+/).filter(w => w.length > 2);
  const bWords = bnorm.split(/\s+/).filter(w => w.length > 2);
  if (!aWords.length || !bWords.length) return false;
  return aWords.every(w => bnorm.includes(w)) && bWords.every(w => anorm.includes(w));
}

function splitAllowed(raw) {
  return raw.split(/\s*[\/,]\s*/).map(s => s.trim()).filter(Boolean);
}

function findBestParent(selectedValues, parentName) {
  const pNorm = normalizePunct(parentName);
  return Object.keys(selectedValues)
    .filter(k => {
      const kNorm = normalizePunct(k);
      return kNorm === pNorm      ||
             kNorm.includes(pNorm)||
             pNorm.includes(kNorm)||
             fuzzyMatch(k, parentName);
    })
    .sort((a, b) => {
      const score = k => {
        const kn = normalizePunct(k);
        if (kn === pNorm)         return 1000;
        if (kn.startsWith(pNorm)) return 500;
        if (kn.includes(pNorm))   return 100;
        return k.length;
      };
      return score(b) - score(a);
    })[0];
}

// ======================= DEPENDENCY ENGINE =======================
function buildDependencyMap(fields) {
  const selectedValues = {};
  const excludedFields = new Set();

  // pass 1 — collect selected values for all select fields
  for (const field of fields) {
    const fieldName = String(field["Field Name"]       || "").trim();
    const fieldType = String(field["Field Input Type"] || "").trim();
    const raw       = String(field["Value"]            || "").trim();
    const rule      = String(field["Business Rule"]    || "").trim();
    if ((TYPE_MAP[fieldType] || "") !== "select") continue;

    let source = raw;
    if (!source || source.toLowerCase().includes("matrix")) {
      if (rule && !isActualRule(rule)) source = rule;
    }
    const selected = source.includes("|") ? source.split("|")[0].trim() : source.trim();
    if (selected && !selected.toLowerCase().includes("matrix")) {
      selectedValues[fieldName] = selected;
    }
  }

  // pass 2 — evaluate rules for non-select fields
  for (const field of fields) {
    const fieldName = String(field["Field Name"]       || "").trim();
    const fieldType = String(field["Field Input Type"] || "").trim();
    const rule      = String(field["Business Rule"]    || "").trim();
    if ((TYPE_MAP[fieldType] || "") === "select") continue;

    if (
      rule.toLowerCase().includes("don't show this field on request form") ||
      rule.toLowerCase().includes("do not show")
    ) { excludedFields.add(fieldName); continue; }

    if (MANDATORY_ONLY_PATTERNS.some(p => p.test(rule))) continue;
    if (!rule) continue;

    if (!checkDependency(rule, selectedValues, fields, fieldName)) {
      excludedFields.add(fieldName);
    }
  }

  return excludedFields;
}

function checkDependency(rule, selectedValues, allFields, currentFieldName) {
  const r = rule.trim();

  // P1a: mandatory when X is selected in field = Y
  const p1a = r.match(/mandatory when\s+['"]?([^'"=\n]+?)['"]?\s+is selected in field\s*[="]*\s*['"]?([^'"\n]+?)['"]?\s*$/i);
  if (p1a) {
    const key = findBestParent(selectedValues, p1a[2].trim());
    if (!key) return false;
    return fuzzyMatch(selectedValues[key], p1a[1].trim());
  }

  // P1b: Open this field if Selection = X
  const p1b = r.match(/open this field if selection[s]?\s*[=:]\s*['"]?([^,.'"\n\r]+)['"]?/i);
  if (p1b) return Object.values(selectedValues).some(v => fuzzyMatch(v, p1b[1].trim()));

  // P1c: Open this field if FIELDNAME is selected as VALUE
  const p1c = r.match(/open this field if\s+(.+?)\s+is selected as\s+['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p1c) {
    const key = findBestParent(selectedValues, p1c[1].trim());
    if (!key) return false;
    return splitAllowed(p1c[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  // P2: Open this field if Selection 'X,Y,Z'
  const p2 = r.match(/open this field if selection\s*['"]([^'"]+)['"]/i);
  if (p2) {
    const allowed = p2[1].split(",").map(s => s.trim());
    return Object.values(selectedValues).some(v => allowed.some(a => fuzzyMatch(v, a)));
  }

  // P3: if X option is selected in above field
  const p3 = r.match(/if\s+(.+?)\s+option is selected in above field/i);
  if (p3) {
    const required     = p3[1].trim();
    const currentIndex = allFields.findIndex(f => f["Field Name"] === currentFieldName);
    const isAddrSelect = f => {
      if ((TYPE_MAP[f["Field Input Type"]] || "") !== "select") return false;
      const fName = String(f["Field Name"] || "").toLowerCase();
      const fVal  = String(f["Value"]      || "").toLowerCase();
      const fRule = String(f["Business Rule"] || "").toLowerCase();
      return fName.includes("address") ||
             fVal.includes("registered") || fVal.includes("corporate") ||
             fVal.includes("branch")     || fVal.includes("principal") ||
             fVal.includes("head office")||
             fRule.includes("registered")|| fRule.includes("corporate") ||
             fRule.includes("branch")    || fRule.includes("principal");
    };
    const prev =
      [...allFields].slice(0, currentIndex).reverse().find(isAddrSelect) ||
      [...allFields].slice(0, currentIndex).reverse()
        .find(f => (TYPE_MAP[f["Field Input Type"]] || "") === "select");
    if (!prev) return false;
    return fuzzyMatch(selectedValues[prev["Field Name"]] || "", required);
  }

  // P4: if above field is selected as X
  const p4 = r.match(/if above field is selected as\s+['"]?([\w][\w\s]*)['"]?/i);
  if (p4) {
    const currentIndex = allFields.findIndex(f => f["Field Name"] === currentFieldName);
    const prev = [...allFields].slice(0, currentIndex).reverse()
      .find(f => (TYPE_MAP[f["Field Input Type"]] || "") === "select");
    if (!prev) return false;
    return fuzzyMatch(selectedValues[prev["Field Name"]] || "", p4[1].trim());
  }

  // P5: if X is selected as 'Y'
  const p5 = r.match(/if\s+(.+?)\s+is selected as\s+['"]([^'"]+)['"]/i);
  if (p5) {
    const key = findBestParent(selectedValues, p5[1].trim());
    if (!key) return false;
    return splitAllowed(p5[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  // P6: shown/opened when X is selected as Y (supports Y/Z multi-value)
  const p6 = r.match(/(?:shown|opened?).*?when\s+(.+?)\s+is selected as[=\s]*['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p6) {
    const key = findBestParent(selectedValues, p6[1].trim());
    if (!key) return false;
    return splitAllowed(p6[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  // P7: mandatory when X is selected
  const p7 = r.match(/mandatory when\s+['"]?([^'"=\n]+?)['"]?\s+is selected/i);
  if (p7) return Object.values(selectedValues).some(v => fuzzyMatch(v, p7[1].trim()));

  // P8: opened up when X is selected as = Y/Z
  const p8 = r.match(/opened?\s+up\s+when\s+(.+?)\s+is selected as[=\s]*['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p8) {
    const key = findBestParent(selectedValues, p8[1].trim());
    if (!key) return false;
    return splitAllowed(p8[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  // P9: mandatory to upload when response to field X is Y
  const p9 = r.match(/mandatory to upload when response to field\s+["']?([^"'\n]+?)["']?\s+is\s+(\w+)/i);
  if (p9) {
    const key = findBestParent(selectedValues, p9[1].trim());
    if (!key) return false;
    return fuzzyMatch(selectedValues[key], p9[2].trim());
  }

  // P10: opened/shown when X = Y (equals sign variant)
  const p10 = r.match(/(?:opened?|shown|display).*?when\s+(.+?)\s*[=:]\s*([^.\n\r]+?)\s*$/i);
  if (p10) {
    const key = findBestParent(selectedValues, p10[1].trim());
    if (!key) return false;
    return splitAllowed(p10[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  // P11: dependent upon X - Y
  const p11 = r.match(/dependent upon\s+(.+?)\s*[-–]\s*(.+)/i);
  if (p11) {
    const key = findBestParent(selectedValues, p11[1].trim());
    if (!key) return false;
    return splitAllowed(p11[2].trim()).some(a => fuzzyMatch(selectedValues[key], a));
  }

  return true; // no pattern matched → show
}

// ======================= EXCEL PARSER =======================
function parseExcelToContracts(filePath) {
  const wb = XLSX.readFile(filePath);
  const contractsMap = {};

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (!rows.length) continue;

    const headerRow = rows.find(r => r.map(c => String(c).trim()).includes("Field Name"));
    if (!headerRow) { console.log(`Skipping: ${sheetName} — no header`); continue; }

    const headerIndex       = rows.indexOf(headerRow);
    const headers           = headerRow.map(h => String(h).trim());
    const fieldNameIndex    = headers.indexOf("Field Name");
    const typeIndex         = headers.indexOf("Field Input Type");
    const mandatoryIndex    = headers.indexOf("Mandatory");
    const valuesIndex       = headers.indexOf("Values For Selection Fields");
    const businessRuleIndex = headers.indexOf("Business Rules");
    const visibilityIndex   = headers.indexOf("Visibility");

    if (fieldNameIndex === -1 || typeIndex === -1) {
      console.log(`Skipping: ${sheetName} — missing columns`); continue;
    }

    const contractStartIndex = businessRuleIndex + 1;
    const contractNames      = headers.slice(contractStartIndex).filter(Boolean);

    contractNames.forEach(name => {
      if (!contractsMap[name]) contractsMap[name] = { "Contract Type": name, fields: [] };
    });

    for (let i = headerIndex + 1; i < rows.length; i++) {
      const row       = rows[i];
      const fieldName = String(row[fieldNameIndex] || "").trim();
      const fieldType = String(row[typeIndex]      || "").trim();
      if (!fieldName || !fieldType) continue;

      contractNames.forEach((contractName, idx) => {
        const enabled = String(row[contractStartIndex + idx] || "").trim().toLowerCase();
        if (enabled !== "yes") return;
        contractsMap[contractName].fields.push({
          "Field Name":       fieldName,
          "Field Input Type": fieldType,
          "Mandatory":        String(row[mandatoryIndex]    || "").trim(),
          "Visibility":       String(row[visibilityIndex]   || "Yes").trim(),
          "Value":            String(row[valuesIndex]       || "").trim(),
          "Business Rule":    String(row[businessRuleIndex] || "").trim(),
        });
      });
    }
    console.log(`✅ Sheet: ${sheetName} | Contracts: ${contractNames.length}`);
  }

  const contracts = Object.values(contractsMap);
  console.log(`✅ Parsed ${contracts.length} contracts`);
  return contracts;
}

// ======================= GENERATE JSON (step 1 task) =======================
async function generateAI(filePath) {
  const contracts  = parseExcelToContracts(filePath);
  const outputPath = path.join(process.cwd(), "cypress", "fixtures", "categoriesAI.json");
  fs.writeFileSync(outputPath, JSON.stringify(contracts, null, 2), "utf-8");
  console.log(`✅ categoriesAI.json — ${contracts.length} contracts`);
  return contracts;
}

// ======================= WRITE ONE CONTRACT EXCEL =======================
async function writeContractExcel(contract, outputDir) {
  const contractType   = contract["Contract Type"];
  const excludedFields = buildDependencyMap(contract.fields || []);

  // build selectedValues for upload field logic
  const selectedValues = {};
  for (const field of contract.fields || []) {
    const fName   = String(field["Field Name"]       || "").trim();
    const fType   = String(field["Field Input Type"] || "").trim();
    const rawVal  = String(field["Value"]            || "").trim();
    const rawRule = String(field["Business Rule"]    || "").trim();
    if ((TYPE_MAP[fType] || "") !== "select") continue;
    let source = rawVal;
    if (!source || source.toLowerCase().includes("matrix")) {
      if (rawRule && !isActualRule(rawRule)) source = rawRule;
    }
    const sel = source.includes("|") ? source.split("|")[0].trim() : source.trim();
    if (sel && !sel.toLowerCase().includes("matrix")) selectedValues[fName] = sel;
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Fields");
  ws.columns = [
    { header: "name", key: "name", width: 50 },
    { header: "data", key: "data", width: 50 },
    { header: "type", key: "type", width: 20 },
  ];
  ws.getRow(1).font = { name: "Arial" };

  let docx3Added = false;
  ws.addRow({ name: "Requestor Name", data: "Tushar Chugh", type: "disableString" });

  for (const field of contract.fields || []) {
    const fieldName  = String(field["Field Name"]       || "").trim();
    const fieldType  = String(field["Field Input Type"] || "").trim();
    const mandatory  = String(field["Mandatory"]        || "").toLowerCase().trim();
    const rule       = String(field["Business Rule"]    || "").toLowerCase();
    const visibility = String(field.Visibility          || "").toLowerCase();

    if (visibility === "no") continue;
    if (rule.includes("don't show this field on request form")) continue;
    if (rule.includes("do not show")) continue;
    if (fieldName === "Requestor Name") continue;
    if (excludedFields.has(fieldName)) { console.log(`  Excluded: ${fieldName}`); continue; }

    if (fieldName === "Requestor Email") {
      ws.addRow({ name: "Requestor Email", data: "tchugh@srtekbox.com", type: "disableString" });
      continue;
    }

    if (fieldType === "Upload") {
      if (fieldName === "Contract Document") {
        const draftVal = selectedValues["Draft Type"] || "";
        if (draftVal.toLowerCase().includes("counterparty")) {
          ws.addRow({ name: "Contract Document", data: "doc1.docx", type: "docxNew" });
        }
        continue;
      }
      if (fieldName === "9-Point Declaration") {
        ws.addRow({ name: "9-Point Declaration", data: "doc3.docx", type: "docx" });
        continue;
      }
      if (mandatory === "no") {
        if (docx3Added) continue;
        docx3Added = true;
        ws.addRow({ name: "Enter Document Name", data: "adoc1.pdf", type: "docx3" });
        continue;
      }
      ws.addRow({ name: fieldName, data: "adoc1.pdf", type: "file" });
      continue;
    }

    const excelType = TYPE_MAP[fieldType] || "string";
    let value = "";
    const rawValue = String(field.Value            || "").trim();
    const rawRule  = String(field["Business Rule"] || "").trim();

    if (excelType === "disableString") {
      value = (rawValue && !rawValue.toLowerCase().includes("matrix")) ? rawValue : "";
    } else if (excelType === "date") {
      value = "31-Jan-26";
    } else if (excelType === "select") {
      if (fieldName === "Contract Type") {
        value = contractType;
      } else {
        let source = rawValue;
        if (!source || source.toLowerCase().includes("matrix")) {
          if (rawRule && !isActualRule(rawRule)) source = rawRule;
        }
        value = source.includes("|")
          ? source.split("|")[0].trim()
          : (source && !source.toLowerCase().includes("matrix") ? source.trim() : "");
      }
    } else {
      let source = "";
      if (rawValue && !rawValue.includes("|") && !rawValue.toLowerCase().includes("matrix")) {
        source = rawValue;
      } else if (rawRule && !isActualRule(rawRule)) {
        source = rawRule;
      }
      value = source || "Apart from counting words and characters";
    }

    ws.addRow({ name: fieldName, data: value, type: excelType });
  }

  const fileName = contractType.replace(/[^\w\s-]/g, "_").trim() + ".xlsx";
  const filePath  = path.join(outputDir, fileName);
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    await wb.xlsx.writeFile(filePath);
    console.log(`✅ Written: ${fileName}`);
    return fileName;
  } catch (err) {
    if (err.code === "EBUSY") console.error(`❌ File open in Excel: ${fileName}`);
    else                      console.error(`❌ Failed: ${fileName} | ${err.message}`);
    return null;
  }
}

// ======================= CONVERT JSON → EXCEL =======================
async function convertJsonToExcelAI(contracts) {
  const outputDir = path.join(process.cwd(), "cypress/fixtures/File/generated-ai");
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const results = { success: [], failed: [] };
  for (const contract of contracts) {
    const fileName = await writeContractExcel(contract, outputDir);
    if (fileName) results.success.push(fileName);
    else          results.failed.push(contract["Contract Type"]);
  }
  return `Done — ${results.success.length} Excel files written to ${outputDir}`;
}

// ======================= SINGLE-STEP: Sheet → Excel =======================
async function generateExcelFromSheet(filePath) {
  const contracts  = parseExcelToContracts(filePath);
  const outputPath = path.join(process.cwd(), "cypress", "fixtures", "categoriesAI.json");
  fs.writeFileSync(outputPath, JSON.stringify(contracts, null, 2), "utf-8");

  const outputDir = path.join(process.cwd(), "cypress/fixtures/File/generated-ai");
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const results = { success: [], failed: [] };
  for (const contract of contracts) {
    const fileName = await writeContractExcel(contract, outputDir);
    if (fileName) results.success.push(fileName);
    else          results.failed.push(contract["Contract Type"]);
  }

  console.log(`\n✅ Done — ${results.success.length} files written, ${results.failed.length} failed`);
  return {
    message: `Generated ${results.success.length} of ${contracts.length} contracts`,
    success: results.success,
    failed:  results.failed,
  };
}

// ======================= EXPORTS =======================
const exportedKeys = ["generateAI", "convertJsonToExcelAI", "generateExcelFromSheet"];
console.log("ALL EXPORTED KEYS:", exportedKeys);

module.exports = {
  generateAI,
  convertJsonToExcelAI,
  generateExcelFromSheet,
};