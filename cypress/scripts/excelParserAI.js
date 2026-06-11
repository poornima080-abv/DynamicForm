const XLSX    = require("xlsx");
const fs      = require("fs");
const path    = require("path");
const ExcelJS = require("exceljs");
const http    = require("http");

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

// ======================= OLLAMA AI =======================
async function callOllama(prompt) {
  return new Promise((resolve) => {
    const body = JSON.stringify({ model: "llama3", prompt, stream: false });

    const options = {
      hostname: "localhost",
      port:     11434,
      path:     "/api/generate",
      method:   "POST",
      headers:  {
        "Content-Type":   "application/json",
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = http.request(options, (res) => {
      let data = "";
      res.on("data", chunk => { data += chunk; });
      res.on("end", () => {
        try {
          const parsed = JSON.parse(data);
          resolve(parsed.response || "");
        } catch (e) {
          console.error(`Ollama parse error: ${e.message}`);
          resolve("");
        }
      });
    });

    req.on("error", (e) => {
      console.error(`Ollama error: ${e.message}`);
      resolve("");
    });

    // no timeout — let AI take as long as it needs
    req.write(body);
    req.end();
  });
}

// ======================= AI DEPENDENCY RESOLVER =======================
async function resolveFieldsWithAI(contractType, fields) {

  const selectedValues = {};
  for (const f of fields) {
    if ((TYPE_MAP[f["Field Input Type"]] || "") === "select") {
      const raw = String(f["Value"] || "").trim();
      const sel = raw.includes("|") ? raw.split("|")[0].trim() : raw.trim();
      if (sel && !sel.toLowerCase().includes("matrix")) {
        selectedValues[f["Field Name"]] = sel;
      }
    }
  }

  const fieldList = fields.map(f => ({
    name:      f["Field Name"],
    type:      f["Field Input Type"],
    mandatory: f["Mandatory"],
    rule:      f["Business Rule"],
  }));

  const prompt = `
You are a form dependency engine for enterprise contract management.

CONTRACT TYPE: "${contractType}"

SELECTED VALUES (first option chosen for each dropdown):
${JSON.stringify(selectedValues, null, 2)}

TASK:
For each field below, read its business rule and decide:
- true  = field should be VISIBLE on the form
- false = field should be HIDDEN

DECISION RULES:
1. Rule says "don't show on request form" or "do not show" → false
2. Rule says "come up as mandatory" or "will be mandatory" → true (always visible, just conditionally mandatory)
3. Rule says "open/show if selection = X" → true only if any selected value matches X
4. Rule says "if X option is selected in above field" → check nearest preceding address dropdown's selected value
5. Rule says "if X is selected as Y" → find field X in selected values, check if it equals Y
6. Rule says "shown when X is selected as Y" → find field X in selected values, check if it equals Y
7. Rule says "opened up when X is selected as Y or Z" → true if field X's value is in the list
8. Rule says "mandatory when X is selected in field = Y" → true only if field Y's value equals X
9. Rule says "mandatory to upload when response to field X is Y" → true if field X's value equals Y
10. No rule or unclear rule → true

CRITICAL — return ONLY this JSON format, nothing else:
{
  "Field Name 1": true,
  "Field Name 2": false
}

FIELDS:
${JSON.stringify(fieldList, null, 2)}
`;

  console.log(`\nSending to AI: ${contractType} | Fields: ${fieldList.length} | Prompt: ${prompt.length} chars`);

  const result = await callOllama(prompt);

  try {
    const clean     = result.replace(/```json|```/g, "").trim();
    const jsonMatch = clean.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error("No JSON found in response");
    const parsed = JSON.parse(jsonMatch[0]);
    console.log(`✅ AI resolved ${Object.keys(parsed).length} fields for: ${contractType}`);
    return parsed;
  } catch (e) {
    console.error(`❌ AI failed for ${contractType}: ${e.message}`);
    console.error(`Raw (first 300 chars): ${result.slice(0, 300)}`);
    return null;
  }
}

// ======================= HELPERS =======================
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

// ======================= GENERATE JSON =======================
async function generateAI(filePath) {
  const contracts  = parseExcelToContracts(filePath);
  const outputPath = path.join(process.cwd(), "cypress", "fixtures", "categoriesAI.json");
  fs.writeFileSync(outputPath, JSON.stringify(contracts, null, 2), "utf-8");
  console.log(`✅ categoriesAI.json — ${contracts.length} contracts`);
  return contracts;
}

// ======================= WRITE ONE CONTRACT EXCEL =======================
async function writeContractExcel(contract, outputDir) {
  const contractType = contract["Contract Type"];

  // AI decides all field visibility
  const aiDecisions = await resolveFieldsWithAI(contractType, contract.fields || []);

  // build selectedValues for upload logic
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

    // hard filters — always applied
    if (visibility === "no") continue;
    if (rule.includes("don't show this field on request form")) continue;
    if (rule.includes("do not show")) continue;
    if (fieldName === "Requestor Name") continue;

    // AI decision — if AI says false, hide it
    if (aiDecisions && aiDecisions[fieldName] === false) {
      console.log(`  AI hidden: ${fieldName}`);
      continue;
    }

    // system fields
    if (fieldName === "Requestor Email") {
      ws.addRow({ name: "Requestor Email", data: "tchugh@srtekbox.com", type: "disableString" });
      continue;
    }

    // upload fields
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

    // resolve type and value
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
    if (err.code === "EBUSY") console.error(`❌ File open: ${fileName}`);
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

// ======================= SINGLE STEP: Sheet → Excel =======================
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

  console.log(`\n✅ Done — ${results.success.length} written, ${results.failed.length} failed`);
  return {
    message: `Generated ${results.success.length} of ${contracts.length} contracts`,
    success: results.success,
    failed:  results.failed,
  };
}

// ======================= EXPORTS =======================
module.exports = {
  generateAI,
  convertJsonToExcelAI,
  generateExcelFromSheet,
};