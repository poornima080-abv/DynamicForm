const XLSX = require("xlsx");
const fs = require("fs");

const filePath = "Implementation_SRF.xlsx";
const wb = XLSX.readFile(filePath);

// take FIRST sheet only
const sheet = wb.Sheets[wb.SheetNames[0]];

// convert to raw matrix (VERY IMPORTANT)
const data = XLSX.utils.sheet_to_json(sheet, {
  header: 1,
  defval: ""
});

if (!data || data.length < 2) {
  throw new Error("Sheet 1 is empty or invalid");
}

// find real header row dynamically
let headerIndex = data.findIndex(r => r.includes("Field Name"));
if (headerIndex === -1) headerIndex = 0;

const headers = data[headerIndex];

// build structured fields safely
const fields = [];

for (let i = headerIndex + 1; i < data.length; i++) {
  const row = data[i];

  // skip empty / section headers like "Contract Details"
  if (!row || !row[1]) continue;
  if (typeof row[1] === "string" && row[1].includes("Agreement")) continue;

  const obj = {};
  headers.forEach((h, idx) => {
    obj[h] = row[idx];
  });

  fields.push({
    name: obj["Field Name"],
    type: obj["Field Input Type"],
    mandatory: String(obj["Mandatory"] || "").trim(),
    rule: obj["Business Rules"] || "",
    visibility: obj
  });
}

const output = {
  contractTypes: [],
  fields
};

fs.writeFileSync(
  "cypress/fixtures/categories.json",
  JSON.stringify(output, null, 2)
);

console.log("categories.json generated safely");