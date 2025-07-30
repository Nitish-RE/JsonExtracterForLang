const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Load the workbook
const workbook = XLSX.readFile('MiY-Germany-UI-Strings.xlsx');

// Prepare objects for en.json and de.json
const en = {};
const de = {};

// Loop through all sheets
for (const sheetName of workbook.SheetNames) {
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Start from row 2 (index 1) if row 1 is header
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const label = row[1]; // Column B
    const enText = row[2]; // Column C
    const deText = row[3]; // Column D

    if (label) {
      en[label] = enText || '';
      de[label] = deText || '';
    }
  }
}

// Ensure the result directory exists
const resultDir = path.join(__dirname, 'result');
if (!fs.existsSync(resultDir)) {
  fs.mkdirSync(resultDir);
}

// Write to result/en.json and result/de.json
fs.writeFileSync(path.join(resultDir, 'en.json'), JSON.stringify(en, null, 2), 'utf8');
fs.writeFileSync(path.join(resultDir, 'de.json'), JSON.stringify(de, null, 2), 'utf8');

console.log('en.json and de.json have been created in the result folder from all sheets.'); 