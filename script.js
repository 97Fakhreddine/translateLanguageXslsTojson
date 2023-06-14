const XLSX = require('xlsx');
const json = require('./fr.json')
const fs = require('fs');
const path = require('path');
function createExcelFromJSON(jsonData, filePath) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();
  const sheetName = 'Sheet1';

  // Convert JSON object to an array of key-value pairs
  // const data = Object.entries(jsonData);
  const data = convertJSON(json);
  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(data);

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

  // Write workbook to file
  XLSX.writeFile(workbook, filePath);
}

const convertJSON = (json) => {
  const data = [];

  if (typeof json === 'object' && json !== null) {
    for (let key in json) {
      if (typeof json[key] === 'string') {
        data.push([key, json[key]]);
      } else {
        const nestedData = convertJSON(json[key]);
        nestedData.forEach((nestedRow) => {
          data.push([key, ...nestedRow]);
        });
      }
    }
  }

  return data;

}

function createJSONFromSheets(workbook) {
  const jsonData = {};

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const worksheet = XLSX.utils.sheet_to_json(sheet);

    let sheetData = jsonData;
    const nestedKeys = sheetName.split('.');
    nestedKeys.forEach((key, index) => {
      if (!sheetData[key]) {
        sheetData[key] = {};
      }
      if (index === nestedKeys.length - 1) {
        worksheet.forEach((row) => {
          const [rowKey, value] = Object.values(row);
          sheetData[key][rowKey] = value;
        });
      } else {
        sheetData = sheetData[key];
      }
    });
  });

  return jsonData;
}


// Example usage
function getSheetFilePath(sheetName) {
  const scriptDir = __dirname;
  const sheetFileName = `${sheetName}.xlsx`;
  const sheetFilePath = path.join(scriptDir, sheetFileName);
  return sheetFilePath;
}




function runCreateSheet() {
  const filePath = 'output.xlsx';
  createExcelFromJSON(json, filePath);
  console.log('Done!!')
}

function runCreateJson() {
  const jsonFilePath = 'ar.json';
  // Write JSON object to a file
  fs.writeFileSync(jsonFilePath, JSON.stringify(json, null, 2));
  const workbook = XLSX.readFile(getSheetFilePath('translationLanguageInFr'));
  createJSONFromSheets(workbook);
  console.log('Done!!')
}
// do not touch the above code we might need it to translate to other language in the future
//
module.exports = {
  runCreateSheet,
  runCreateJson
}
