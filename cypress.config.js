const XLSX = require('xlsx');
const {
 convertJsonToExcelAI,  generateExcelFromSheet, generateAI
} = require("./cypress/scripts/excelParserAI.js");

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
  

        readExcelFile(filePath) {
          return readExcelFile(filePath);
        },

        // step 1 — Excel → JSON (no AI)
        async generateAI({ filePath }) {
          return await generateAI(filePath);
        },

        // step 2 — JSON → Excel (AI decides visibility)
        async convertJsonToExcelAI({ contracts }) {
          return await convertJsonToExcelAI(contracts);
        },
        // step 3 — UI Validation (placeholder)
         async generateExcelFromSheet({ contracts }) {
          return await generateExcelFromSheet(contracts);
        }
      });

      return config;
    },

    defaultCommandTimeout: 120000,
    taskTimeout: 3600000
  }
};