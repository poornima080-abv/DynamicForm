const XLSX = require('xlsx');

function readExcelFile(filePath) {

  const workbook = XLSX.readFile(filePath);

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
        }

      });

      return config;
    }

  }

};
