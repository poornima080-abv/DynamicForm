const path = require('path');
const fs = require("fs");
const XLSX = require("xlsx");

const excelPath = path.resolve(
  process.cwd(),
  'Implementation_SRF.xlsx'
);

console.log("Reading Excel from:", excelPath);

const workbook = XLSX.readFile(excelPath);
const sheet =
  workbook.Sheets[workbook.SheetNames[0]];

const rows =
  XLSX.utils.sheet_to_json(sheet);

const fixedColumns = [
  "Field Name",
  "Field Type",
  "Mandatory",
  "Business Rules"
];

const contractTypes =
  Object.keys(rows[0]).filter(
    col => !fixedColumns.includes(col)
  );

const result = {
  contractTypes,
  fields: []
};

rows.forEach(row => {

  const field = {

    name: row["Field Name"] || "",

    type: row["Field Type"] || "",

    mandatory: row["Mandatory"] || "",

    businessRule:
      row["Business Rules"] || "",

    visibility: {}

  };

  contractTypes.forEach(type => {

    field.visibility[type] =
      row[type] || "";

  });

  result.fields.push(field);

});

fs.writeFileSync(
  "cypress/fixtures/categories.json",
  JSON.stringify(result, null, 2)
);

console.log("categories.json generated");