const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Load parsed metadata
const data = require("../fixtures/categories.json");

// Ensure output folder exists
const jsonOutDir = path.join(__dirname, "../generated");
const excelOutDir = path.join(__dirname, "../output");

if (!fs.existsSync(jsonOutDir)) {
  fs.mkdirSync(jsonOutDir, { recursive: true });
}

if (!fs.existsSync(excelOutDir)) {
  fs.mkdirSync(excelOutDir, { recursive: true });
}

// -------------------------
// SAFE STRING NORMALIZER
// -------------------------
const normalize = (val) =>
  String(val ?? "").trim().toLowerCase();

// -------------------------
// SAFE FILE NAME CLEANER
// -------------------------
const safeFileName = (name) =>
  String(name).replace(/[^a-z0-9-_]/gi, "_");

// Loop all contract types
data.contractTypes.forEach((contractType) => {

  const fieldsForType = [];

  data.fields.forEach((field) => {

    const visibility = field.visibility?.[contractType];

    if (normalize(visibility) === "yes") {
      fieldsForType.push({
        name: field.name || "",
        type: field.type || "",
        mandatory: field.mandatory || "",
        rule: field.businessRule || "",
        value: ""
      });
    }
  });

  // =========================
  // 1. Write JSON file
  // =========================
  const jsonFilePath = path.join(
    jsonOutDir,
    `generated_${safeFileName(contractType)}.json`
  );

  fs.writeFileSync(
    jsonFilePath,
    JSON.stringify(
      {
        contractType,
        fields: fieldsForType
      },
      null,
      2
    )
  );

  // =========================
  // 2. Write Excel file
  // =========================
  const excelRows = fieldsForType.map((f) => ({
    Name: f.name,
    Value: "",
    Type: f.type
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(excelRows);

  XLSX.utils.book_append_sheet(wb, ws, "Fields");

  const excelFilePath = path.join(
    excelOutDir,
    `${safeFileName(contractType)}.xlsx`
  );

  XLSX.writeFile(wb, excelFilePath);
});

console.log("✅ Category sheets + Excel files generated successfully");