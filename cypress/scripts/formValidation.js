const fs = require("fs");

function getFields(contractType) {

  const file =
    `generated_${contractType}.json`;

  return JSON.parse(
    fs.readFileSync(file)
  );

}

module.exports = {
  getFields
};