const xlsx = require("xlsx");
const fs = require("fs");

// Load the master workbook
const workbook = xlsx.readFile("./PWMS PHED Location Master .xlsx");

// Get the "Nadia" sheet (change if needed)
const sheetName = "Nadia";
const sheet = workbook.Sheets[sheetName];

// Convert sheet to JSON
const data = xlsx.utils.sheet_to_json(sheet);

// Sort data alphabetically by "PWSS" (Scheme) field
const sortedData = data.sort((a, b) => {
    const schemeA = (a["PWSS"] || "").toUpperCase();
    const schemeB = (b["PWSS"] || "").toUpperCase();
    if (schemeA < schemeB) return -1;
    if (schemeA > schemeB) return 1;
    return 0;
});

// Write sorted data to a new Excel file
const ws = xlsx.utils.json_to_sheet(sortedData);
const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, ws, "SortedSchemes");

// Ensure output directory exists
if (!fs.existsSync('./output')) fs.mkdirSync('./output');

xlsx.writeFile(wb, './output/Sorted_Schemes.xlsx');

console.log("✅ Sorted Excel file created at ./output/Sorted_Schemes.xlsx");