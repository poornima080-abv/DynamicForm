const XLSX    = require("xlsx");
const fs      = require("fs");
const path    = require("path");
const ExcelJS = require("exceljs");

const DEBUG_DEPENDENCY = false;

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

// rules that only affect mandatory — field always visible
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
    t.includes("dependent upon")      ||
    t.includes("open this field")     ||
    t.includes("mandatory when")      ||
    t.includes("don't show")          ||
    t.includes("do not show")         ||
    t.includes("disable this")        ||
    t.includes("show this field")     ||
    t.includes("come up as mandatory")||
    t.includes("will be shown")       ||
    t.includes("will be mandatory")   ||
    t.includes("if selected")         ||
    t.includes("if selection")        ||
    t.includes("selected as")         ||
    t.includes("publish form")        ||
    t.includes("request form")        ||
    t.includes("opened up when")      ||
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

  const aWords = anorm.split(/\s+/).filter(w => w.length > 2);
  const bWords = bnorm.split(/\s+/).filter(w => w.length > 2);
  if (!aWords.length || !bWords.length) return false;
  const shorter    = aWords.length <= bWords.length ? aWords : bWords;
  const longerStr  = aWords.length <= bWords.length ? bnorm : anorm;
  const matchCount = shorter.filter(w => longerStr.includes(w)).length;
  return matchCount === shorter.length;
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
        if (kn === pNorm)           return 1000;
        if (kn.startsWith(pNorm))   return 500;
        if (kn.includes(pNorm))     return 100;
        return k.length;
      };
      return score(b) - score(a);
    })[0];
}

// ======================= DEPENDENCY MAP =======================
function buildDependencyMap(fields) {

  const selectedValues = {};
  const excludedFields = new Set();

  // first pass — resolve select values
  for (const field of fields) {
    const fieldName = String(field["Field Name"]       || "").trim();
    const fieldType = String(field["Field Input Type"] || "").trim();
    const raw       = String(field["Value"]            || "").trim();
    const rule      = String(field["Business Rule"]    || "").trim();
    const excelType = TYPE_MAP[fieldType] || "string";

    if (excelType === "select") {
      // use Value first, fallback to Business Rule if not actual rule
      let source = raw;
      if (!source || source.toLowerCase().includes("matrix")) {
        if (rule && !isActualRule(rule)) source = rule;
      }
      const selected = source.includes("|") ? source.split("|")[0].trim() : source.trim();
      if (selected && !selected.toLowerCase().includes("matrix")) {
        selectedValues[fieldName] = selected;
      }
    }
  }

  if (DEBUG_DEPENDENCY) console.log("Selected values:", JSON.stringify(selectedValues, null, 2));

  // second pass — check rules for all non-select fields
  for (const field of fields) {
    const fieldName = String(field["Field Name"]       || "").trim();
    const fieldType = String(field["Field Input Type"] || "").trim();
    const rule      = String(field["Business Rule"]    || "").trim();
    const excelType = TYPE_MAP[fieldType] || "string";

    // never exclude select fields — they are parents
    if (excelType === "select") continue;

    // handle don't show rules
    if (
      rule.toLowerCase().includes("don't show this field on request form") ||
      rule.toLowerCase().includes("do not show")
    ) {
      excludedFields.add(fieldName);
      continue;
    }

    // skip mandatory-only rules — field always visible
    if (MANDATORY_ONLY_PATTERNS.some(p => p.test(rule))) {
      if (DEBUG_DEPENDENCY) console.log(`MANDATORY-ONLY | ${fieldName} | Always visible`);
      continue;
    }

    if (!rule) continue;

    const include = checkDependency(rule, selectedValues, fields, fieldName);
    if (!include) {
      excludedFields.add(fieldName);
    }
  }

  if (DEBUG_DEPENDENCY) console.log("Excluded:", [...excludedFields]);
  return excludedFields;
}

// ======================= CHECK DEPENDENCY =======================
function checkDependency(rule, selectedValues, allFields, currentFieldName) {

  const r   = rule.trim();
  const log = msg => { if (DEBUG_DEPENDENCY) console.log(msg); };

  // ── P1a: "mandatory when X is selected in field = Y" ───────
  const p1a = r.match(/mandatory when\s+['"]?([^'"=\n]+?)['"]?\s+is selected in field\s*[="]*\s*['"]?([^'"\n]+?)['"]?\s*$/i);
  if (p1a) {
    const req = p1a[1].trim();
    const par = p1a[2].trim();
    const key = findBestParent(selectedValues, par);
    if (!key) { log(`P1a | ${currentFieldName} | Parent "${par}" NOT FOUND`); return false; }
    const matched = fuzzyMatch(selectedValues[key], req);
    log(`P1a | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Required: "${req}" | Matched: ${matched}`);
    return matched;
  }

  // ── P1b: "Open this field if Selection = X" ─────────────────
  const p1b = r.match(/open this field if selection[s]?\s*[=:]\s*['"]?([^,.'"\n\r]+)['"]?/i);
  if (p1b) {
    const req     = p1b[1].trim();
    const matched = Object.values(selectedValues).some(v => fuzzyMatch(v, req));
    log(`P1b | ${currentFieldName} | Required: "${req}" | Matched: ${matched}`);
    return matched;
  }

  // ── P1c: "Open this field if X is selected as Y" ────────────
  // e.g. "Open this field if Contract Term is selected as Auto Renewal"
  // NOTE: must come before P2 to avoid swallowing the quotes variant
  const p1c = r.match(/open this field if\s+(.+?)\s+is selected as\s+['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p1c) {
    const parentName = p1c[1].trim();
    const required   = p1c[2].trim();
    const key        = findBestParent(selectedValues, parentName);
    if (!key) { log(`P1c | ${currentFieldName} | Parent "${parentName}" NOT FOUND`); return false; }
    const matched = fuzzyMatch(selectedValues[key], required);
    log(`P1c | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  // ── P2: "Open this field if Selection 'X,Y,Z'" ──────────────
  const p2 = r.match(/open this field if selection\s*['"]([^'"]+)['"]/i);
  if (p2) {
    const allowed = p2[1].split(",").map(s => s.trim());
    const matched = Object.values(selectedValues).some(v => allowed.some(a => fuzzyMatch(v, a)));
    log(`P2 | ${currentFieldName} | Allowed: [${allowed}] | Matched: ${matched}`);
    return matched;
  }

  // ── P3: "if X option is selected in above field" ────────────
  const p3 = r.match(/if\s+(.+?)\s+option is selected in above field/i);
  if (p3) {
    const required     = p3[1].trim();
    const currentIndex = allFields.findIndex(f => f["Field Name"] === currentFieldName);

    // find nearest prev select by field name containing "address"
    let prevSelect = [...allFields].slice(0, currentIndex).reverse()
      .find(f => {
        if ((TYPE_MAP[f["Field Input Type"]] || "") !== "select") return false;
        const fName = String(f["Field Name"] || "").toLowerCase();
        return fName.includes("company address") || fName.includes("counterparty address");
      });

    // fallback — find by value/rule containing office options
    if (!prevSelect) {
      prevSelect = [...allFields].slice(0, currentIndex).reverse()
        .find(f => {
          if ((TYPE_MAP[f["Field Input Type"]] || "") !== "select") return false;
          const fVal  = String(f["Value"]         || "").toLowerCase();
          const fRule = String(f["Business Rule"] || "").toLowerCase();
          const src   = fVal || fRule;
          return (
            src.includes("registered") || src.includes("corporate") ||
            src.includes("branch")     || src.includes("principal") ||
            src.includes("head office")
          );
        });
    }

    if (!prevSelect) { log(`P3 | ${currentFieldName} | No address parent found`); return false; }
    const prevVal = selectedValues[prevSelect["Field Name"]] || "";
    const matched = fuzzyMatch(prevVal, required);
    log(`P3 | ${currentFieldName} | PrevField: "${prevSelect["Field Name"]}" | PrevVal: "${prevVal}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  // ── P4: "if above field is selected as X" ───────────────────
  const p4 = r.match(/if above field is selected as\s+['"]?(\w+)['"]?/i);
  if (p4) {
    const required     = p4[1].trim();
    const currentIndex = allFields.findIndex(f => f["Field Name"] === currentFieldName);
    const prev         = [...allFields].slice(0, currentIndex).reverse()
      .find(f => (TYPE_MAP[f["Field Input Type"]] || "") === "select");
    if (!prev) return false;
    const prevVal = selectedValues[prev["Field Name"]] || "";
    const matched = fuzzyMatch(prevVal, required);
    log(`P4 | ${currentFieldName} | PrevField: "${prev["Field Name"]}" | PrevVal: "${prevVal}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  // ── P5: "if X is selected as 'Y'" ───────────────────────────
  const p5 = r.match(/if\s+(.+?)\s+is selected as\s+['"]([^'"]+)['"]/i);
  if (p5) {
    const parentName = p5[1].trim();
    const required   = p5[2].trim();
    const key        = findBestParent(selectedValues, parentName);
    if (!key) { log(`P5 | ${currentFieldName} | Parent "${parentName}" NOT FOUND`); return false; }
    const matched = fuzzyMatch(selectedValues[key], required);
    log(`P5 | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  // ── P6: "shown/opened when X is selected as Y" ──────────────
  const p6 = r.match(/(?:shown|opened?).*?when\s+(.+?)\s+is selected as[=\s]*['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p6) {
    const parentName = p6[1].trim();
    const required   = p6[2].trim();
    const key        = findBestParent(selectedValues, parentName);
    if (!key) { log(`P6 | ${currentFieldName} | Parent "${parentName}" NOT FOUND`); return false; }
    const matched = fuzzyMatch(selectedValues[key], required);
    log(`P6 | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  // ── P7: "mandatory when X is selected" ──────────────────────
  const p7 = r.match(/mandatory when\s+['"]?([^'"=\n]+?)['"]?\s+is selected/i);
  if (p7) {
    const req     = p7[1].trim();
    const matched = Object.values(selectedValues).some(v => fuzzyMatch(v, req));
    log(`P7 | ${currentFieldName} | Required: "${req}" | Matched: ${matched}`);
    return matched;
  }

  // ── P8: "opened up when X is selected as = Y/Z" ─────────────
  const p8 = r.match(/opened?\s+up\s+when\s+(.+?)\s+is selected as[=\s]*['"]?([^'".\n\r]+?)['"]?\s*$/i);
  if (p8) {
    const parentName    = p8[1].trim();
    const allowedValues = p8[2].trim().split(/[\/,]/).map(s => s.trim());
    const key           = findBestParent(selectedValues, parentName);
    if (!key) { log(`P8 | ${currentFieldName} | Parent "${parentName}" NOT FOUND`); return false; }
    const matched = allowedValues.some(a => fuzzyMatch(selectedValues[key], a));
    log(`P8 | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Allowed: [${allowedValues}] | Matched: ${matched}`);
    return matched;
  }

  // ── P9: "mandatory to upload when response to field X is Y" ─
  const p9 = r.match(/mandatory to upload when response to field\s+["']?([^"'\n]+?)["']?\s+is\s+(\w+)/i);
  if (p9) {
    const parentField = p9[1].trim();
    const required    = p9[2].trim();
    const key         = findBestParent(selectedValues, parentField);
    if (!key) { log(`P9 | ${currentFieldName} | Parent "${parentField}" NOT FOUND`); return false; }
    const matched = fuzzyMatch(selectedValues[key], required);
    log(`P9 | ${currentFieldName} | Parent: "${key}" | Selected: "${selectedValues[key]}" | Required: "${required}" | Matched: ${matched}`);
    return matched;
  }

  log(`NO PATTERN | ${currentFieldName} | Rule: "${rule}" | Defaulting to: true`);
  return true;
}

// ======================= AI ENGINE ENTRY =======================
async function runExcelEngine(filePath) {

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

  const contracts  = Object.values(contractsMap);
  const outputPath = path.join(process.cwd(), "cypress", "fixtures", "categories.json");
  fs.writeFileSync(outputPath, JSON.stringify(contracts, null, 2), "utf-8");
  console.log(`✅ categories.json — ${contracts.length} contracts`);
  return contracts;
}

// ======================= JSON → EXCEL =======================
async function convertToExcel(contracts) {

  const outputDir = path.join(process.cwd(), "cypress/generated");
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  for (const contract of contracts) {

    const contractType   = contract["Contract Type"];
    const excludedFields = buildDependencyMap(contract.fields || []);
    if (DEBUG_DEPENDENCY) console.log(`Excluded: ${[...excludedFields].join(", ")}`);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Fields");

    ws.columns = [
      { header: "name", key: "name", width: 50 },
      { header: "data", key: "data", width: 50 },
      { header: "type", key: "type", width: 20 }
    ];

    ws.getRow(1).font = { name: "Arial" };

    let docx3Added = false;

    // first row always Requestor Name
    ws.addRow({ name: "Requestor Name", data: "Tushar Chugh", type: "disableString" });

    // build selectedValues for upload checks
    const selectedValues = {};
    for (const field of contract.fields || []) {
      const fName   = String(field["Field Name"]       || "").trim();
      const fType   = String(field["Field Input Type"] || "").trim();
      const rawVal  = String(field["Value"]            || "").trim();
      const rawRule = String(field["Business Rule"]    || "").trim();
      if ((TYPE_MAP[fType] || "") === "select") {
        let source = rawVal;
        if (!source || source.toLowerCase().includes("matrix")) {
          if (rawRule && !isActualRule(rawRule)) source = rawRule;
        }
        const sel = source.includes("|") ? source.split("|")[0].trim() : source.trim();
        if (sel && !sel.toLowerCase().includes("matrix")) selectedValues[fName] = sel;
      }
    }

    for (const field of contract.fields || []) {

      const fieldName  = String(field["Field Name"]       || "").trim();
      const fieldType  = String(field["Field Input Type"] || "").trim();
      const mandatory  = String(field["Mandatory"]        || "").toLowerCase().trim();
      const rule       = String(field["Business Rule"]    || "").toLowerCase();
      const visibility = String(field.Visibility         || "").toLowerCase();

      // skip hidden
      if (visibility === "no") continue;
      if (rule.includes("don't show this field on request form")) continue;
      if (rule.includes("do not show")) continue;

      // skip already added
      if (fieldName === "Requestor Name") continue;

      // skip excluded
      if (excludedFields.has(fieldName)) {
        if (DEBUG_DEPENDENCY) console.log(`Skipped: ${fieldName}`);
        continue;
      }

      // Requestor Email
      if (fieldName === "Requestor Email") {
        ws.addRow({ name: "Requestor Email", data: "tchugh@srtekbox.com", type: "disableString" });
        continue;
      }

      // Upload fields
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

      // resolve type
      const excelType = TYPE_MAP[fieldType] || "string";

      // resolve value
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
    const filePath = path.join(outputDir, fileName);

    try {
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      await wb.xlsx.writeFile(filePath);
      console.log(`✅ Written: ${fileName}`);
    } catch (err) {
      if (err.code === "EBUSY") {
        console.error(`❌ File open in Excel: ${fileName}`);
      } else {
        console.error(`❌ Failed: ${fileName} | ${err.message}`);
      }
    }
  }

  return `Done`;
}

module.exports = {
  generateAI:         runExcelEngine,
  convertJsonToExcel: convertToExcel
};