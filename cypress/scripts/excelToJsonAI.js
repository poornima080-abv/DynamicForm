const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { runExcelEngine } = require("./cypress/scripts/excelParser.js");

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

function shouldShowField(field) {
  const visibility = String(field.Visibility || "").trim().toLowerCase();
  const rule       = String(field["Business Rule"] || "").toLowerCase();
  if (visibility === "no") return false;
  if (rule.includes("don't show this field on request form")) return false;
  if (rule.includes("do not show")) return false;
  return visibility === "yes";
}

async function extractDependency(rule, fieldName) {
  if (!rule || rule.trim() === "") return null;

  const r = rule.trim();

  // -------------------------
  // 1. Try existing regex first
  // -------------------------
  const openIfMatch = r.match(/open this field if selection[s]?\s*[=:]\s*['"]?([^,.'"\n]+)['"]?/i);
  if (openIfMatch) {
    return {
      type: "open_if_selection",
      condition: openIfMatch[1].trim().toLowerCase()
    };
  }

  const fieldValueMatch = r.match(/if\s+(.+?)\s+is selected as\s+['"]([^'"]+)['"]/i);
  if (fieldValueMatch) {
    return {
      type: "field_value",
      parentField: fieldValueMatch[1].trim().toLowerCase(),
      condition: fieldValueMatch[2].trim().toLowerCase()
    };
  }

  // -------------------------
  // 2. AI fallback (ONLY if regex fails)
  // -------------------------
  try {
    const response = await ollama.chat({
      model: "llama3",
      messages: [
        {
          role: "user",
          content: `
Convert this business rule into structured JSON.

Field Name: ${fieldName}
Rule: ${r}

Return ONLY JSON in this format:

{
  "type": "field_value",
  "parentField": "",
  "condition": ""
}

or

{
  "type": "open_if_selection",
  "condition": ""
}
          `
        }
      ]
    });

    let text = response.message.content
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();

    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

function normalizeText(str) {
  return String(str || "")
    .toLowerCase()
    .replace(/\./g, "")
    .replace(/\bno\b/g, "number")
    .replace(/\bpo no\b/g, "po number")
    .replace(/\s+/g, " ")
    .trim();
}

function fuzzyMatch(a, b) {
  const al = normalizeText(a);
  const bl = normalizeText(b);

  if (al === bl) return true;

  if (al.includes(bl) || bl.includes(al))
    return true;

  return false;
}

async function shouldIncludeField(field, selectedValues, previousFieldName) {
  const rule = String(field["Business Rule"] || "");
  const dep = await extractDependency(rule, field["Field Name"]);
  if (!dep) return true;

  switch (dep.type) {

    case "open_if_selection": {
      return Object.values(selectedValues).some(val =>
        fuzzyMatch(val, dep.condition)
      );
    }

    case "open_if_list": {
      return Object.values(selectedValues).some(val =>
        dep.condition.some(opt => fuzzyMatch(val, opt))
      );
    }

    case "above_field_option": {
      const aboveValue = selectedValues[previousFieldName] || "";
      return fuzzyMatch(aboveValue, dep.condition);
    }

    case "above_field_value": {
      const aboveValue = selectedValues[previousFieldName] || "";
      return fuzzyMatch(aboveValue, dep.condition);
    }

    case "field_value": {
      const parentKey = Object.keys(selectedValues).find(k =>
        fuzzyMatch(k, dep.parentField)
      );
      if (!parentKey) return false;
      return fuzzyMatch(selectedValues[parentKey], dep.condition);
    }

    default:
      return true;
  }
}

function resolveUploadType(field) {
  const fieldName = String(field["Field Name"] || "").trim();
  const mandatory = String(field["Mandatory"] || "").toLowerCase().trim();
  if (fieldName === "Contract Document")   return "docxNew";
  if (fieldName === "9-Point Declaration") return "docx";
  if (mandatory === "no")                  return "docx3";
  return "file";
}

function resolveUploadValue(field) {
  const fieldName = String(field["Field Name"] || "").trim();
  if (fieldName === "Contract Document")   return "doc1.docx";
  if (fieldName === "9-Point Declaration") return "doc3.docx";
  return "adoc1.pdf";
}

function resolveUploadName(field) {
  const fieldName = String(field["Field Name"] || "").trim();
  const mandatory = String(field["Mandatory"] || "").toLowerCase().trim();
  if (fieldName === "Contract Document")   return "Contract Document";
  if (fieldName === "9-Point Declaration") return "9-Point Declaration";
  if (mandatory !== "no")                  return fieldName;
  return "Enter Document Name";
}

function resolveValue(field, excelType) {
  const raw       = String(field.Value || "").trim();
  const fieldName = String(field["Field Name"] || "").trim();

  if (excelType === "date")          return "31-Jan-26";
  if (excelType === "disableString") return "";

  if (excelType === "select") {
    if (fieldName === "Contract Type") return field["_contractType"] || "";
    if (raw && raw.includes("|"))      return raw.split("|")[0].trim();
    if (raw && !raw.toLowerCase().includes("matrix") && raw.trim()) return raw;
    return "";
  }

  if (excelType === "string" || excelType === "textarea") {
    if (
      raw &&
      !raw.includes("|") &&
      !raw.toLowerCase().includes("matrix") &&
      raw.trim()
    ) return raw;
    return "Apart from counting words and characters";
  }

  return "";
}

function safeFilename(name) {
  return name.replace(/[^a-zA-Z0-9 _-]/g, "_").trim() + ".xlsx";
}

function readExcelFile(filePath) {
  const workbook  = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet);
}

module.exports = {
  e2e: {
    setupNodeEvents(on, config) {

      on('task', {

        readExcelFile(filePath) {
          return readExcelFile(filePath);
        },

        async convertJsonToExcel(_) {
          const jsonPath  = path.join(__dirname, 'cypress', 'fixtures', 'categories.json');
          const outputDir = path.join(__dirname, 'cypress', 'fixtures', 'File', 'generated');

          if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

          const meta = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));

          for (const contract of meta) {
            const contractType  = contract['Contract Type'];
            const visibleFields = (contract.fields || []).filter(shouldShowField);

            console.log(`Contract: ${contractType} | Fields: ${visibleFields.length}`);

            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet('Fields');

            ws.columns = [
              { header: 'name', key: 'name', width: 40 },
              { header: 'data', key: 'data', width: 40 },
              { header: 'type', key: 'type', width: 20 },
            ];

            ws.getRow(1).font = { name: 'Arial' };

            const selectedValues  = {};
            let docx3Added        = false;
            let previousFieldName = "";

            for (const field of visibleFields) {
              const fieldName      = field["Field Name"];
              const fieldInputType = field["Field Input Type"];
              field["_contractType"] = contractType;

              if (!(await shouldIncludeField(field, selectedValues, previousFieldName))) {
                console.log(`Skipping: ${fieldName}`);
                continue;
              }

              let excelType, value, rowName;

              if (fieldInputType === "Upload") {
                excelType = resolveUploadType(field);
                value     = resolveUploadValue(field);
                rowName   = resolveUploadName(field);

                if (excelType === "docx3") {
                  if (docx3Added) continue;
                  docx3Added = true;
                }

              } else {
                excelType = TYPE_MAP[fieldInputType] || "string";
                value     = resolveValue(field, excelType);
                rowName   = fieldName;

                if (excelType === "select" && value) {
                  selectedValues[fieldName] = value;
                  previousFieldName         = fieldName;
                }
              }

              ws.addRow({ name: rowName, data: value, type: excelType });
            }

            const filename = safeFilename(contractType);
            const filepath = path.join(outputDir, filename);

            try {
              await wb.xlsx.writeFile(filepath);
              console.log(`✅ Written: ${filename}`);
            } catch (err) {
              console.error(`❌ Failed: ${filename} | ${err.message}`);
            }
          }

          const files = fs.readdirSync(outputDir);
          console.log(`\nTotal files written: ${files.length}`);
          return `Done — ${files.length} Excel files written to ${outputDir}`;
        }

      });

      return config;
    }
  }
};
