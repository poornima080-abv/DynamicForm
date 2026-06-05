const fs = require("fs");
const XLSX = require("xlsx");

const { Ollama } = require("ollama");
const ollama = new Ollama({
  host: "http://127.0.0.1:11434"
});

const data = JSON.parse(
  fs.readFileSync("cypress/fixtures/contract-output.json", "utf8")
);

Object.keys(data).forEach((contractType) => {

  const rows = data[contractType].map(f => ({
    "Field Name": f.fieldName,
    "Value": f.value,
    "Field Input Type": f.type
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);

  XLSX.utils.book_append_sheet(wb, ws, "Fields");

  XLSX.writeFile(wb, `output/${contractType}.xlsx`);
});

console.log("Excel files generated successfully");