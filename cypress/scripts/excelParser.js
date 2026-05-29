const XLSX = require('xlsx');
const fs = require('fs-extra');

const workbook = XLSX.readFile('./cypress/fixtures/Implementation_SRF.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet);

const categories = {};

rows.forEach(row => {

  const category = row['Category']; // MUST exist in Excel

  if (!category) return;

  if (!categories[category]) {
    categories[category] = [];
  }

  categories[category].push({
    name: row['Field Name'],
    type: row['Field Type']
  });

});

fs.writeJsonSync('./cypress/fixtures/categories.json', categories, { spaces: 2 });

console.log("DONE");