const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

// Excel file path
const filePath = path.join(__dirname, "../../Implementation_SRF.xlsx");

// Read workbook
const workbook = XLSX.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Read sheet as matrix
const rows = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: ""
});

// Find header row
const headerIndex = rows.findIndex(
  (row) =>
    row.includes("Field Name") &&
    row.includes("Business Rules")
);

if (headerIndex === -1) {
  throw new Error("Header row not found");
}

const headers = rows[headerIndex];

// Column indexes
const serialIndex = headers.indexOf("#");
const fieldNameIndex = headers.indexOf("Field Name");
const fieldTypeIndex = headers.indexOf("Field Input Type");
const mandatoryIndex = headers.indexOf("Mandatory");
const valueIndex = headers.indexOf("Values For Selection Fields");
const businessRulesIndex = headers.indexOf("Business Rules");

if (
  fieldNameIndex === -1 ||
  fieldTypeIndex === -1 ||
  mandatoryIndex === -1 ||
  businessRulesIndex === -1
) {
  throw new Error("Required columns not found in sheet");
}

// Contract types start after "Business Rules"
const contractTypes = headers
  .slice(businessRulesIndex + 1)
  .filter(Boolean);

console.log(`Found ${contractTypes.length} contract types`);

// Create contract buckets
const contractsMap = new Map();

contractTypes.forEach((contractType, idx) => {
  contractsMap.set(contractType, {
    "#": idx + 1,
    "Contract Type": contractType,
    fields: []
  });
});

// Process rows
for (let i = headerIndex + 1; i < rows.length; i++) {
  const row = rows[i];

  if (!row || row.length === 0) continue;

  const fieldName = row[fieldNameIndex];

  // Skip empty rows
  if (!fieldName) continue;

  // Skip section headers
  if (
    typeof fieldName === "string" &&
    !row[fieldTypeIndex] &&
    !row[mandatoryIndex]
  ) {
    continue;
  }

  contractTypes.forEach((contractType, idx) => {
    const visibilityColumn = businessRulesIndex + 1 + idx;

    const visibility =
      row[visibilityColumn] && String(row[visibilityColumn]).trim()
        ? String(row[visibilityColumn]).trim()
        : "No";
    const contract = contractsMap.get(contractType);

    contract.fields.push({
      "Field Name": fieldName,
      "Visibility": visibility,
      "Field Input Type": row[fieldTypeIndex] || "",
      "Mandatory": row[mandatoryIndex] || "",
      "Value": row[valueIndex] || "",
      "Business Rule": row[businessRulesIndex] || ""
    });
  });
}
console.log(headers);
// Convert map to array
const output = Array.from(contractsMap.values());

// Output file
const outputPath = path.join(
  __dirname,
  "../fixtures/categories.json"
);

// Write JSON
fs.writeFileSync(
  outputPath,
  JSON.stringify(output, null, 2),
  "utf8"
);

console.log(`Generated ${output.length} contract types`);
console.log("Done");