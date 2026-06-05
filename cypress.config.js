const XLSX = require('xlsx');

const {
  generateAI: generateAIEngine,
  convertJsonToExcel
} = require("./cypress/scripts/excelParser.js");
function readExcelFile(filePath) {

  const workbook = XLSX.readFile(filePath);

  const sheetName = workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet);
}

module.exports = {
  e2e: {

    setupNodeEvents(on, config) {

      on("task", {

       async generateAI({ filePath }) {
  console.log("FILE PATH:", filePath);
  return await generateAIEngine(filePath);
},

        async convertJsonToExcel({ contracts }) {
          return await convertJsonToExcel(contracts);
        }

      });

      return config;
    },

    defaultCommandTimeout: 120000,
    taskTimeout: 600000
  }
};