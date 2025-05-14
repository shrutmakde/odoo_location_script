const xlsx = require("xlsx");
const fs = require("fs");

// Load the master workbook
const workbook = xlsx.readFile("./PWMS PHED Location Master .xlsx");

// Get the "Nadia" sheet
const sheetName = "Nadia";
const sheet = workbook.Sheets[sheetName];

// Convert sheet to JSON
const data = xlsx.utils.sheet_to_json(sheet);

// Constants
const STATE = "West Bengal";

// --- 1. Block Master Excel ---
const blockMap = new Map();
data.forEach(row => {
    const block = row["Block"];
    const district = row["District"];
    if (block && !blockMap.has(block)) {
        blockMap.set(block, { "Block Name": block, State: STATE, District: district });
    }
});
const blockMaster = Array.from(blockMap.values());

// --- 2. Scheme Master Excel ---
const schemeMap = new Map();
data.forEach(row => {
    const scheme = row["PWSS"];
    const block = row["Block"];
    const district = row["District"];
    if (scheme && !schemeMap.has(scheme)) {
        schemeMap.set(scheme, {
            "Scheme Name": scheme,
            Block: block,
            District: district,
            State: STATE
        });
    }
});
const schemeMaster = Array.from(schemeMap.values());

// --- 3. Zone Master Excel ---
const zoneMap = new Map();
data.forEach(row => {
    const scheme = row["PWSS"];
    const zone = row["Zone"] || "N/A";
    const block = row["Block"];
    const district = row["District"];

    const key = `${scheme}-${zone}`;
    if (!zoneMap.has(key)) {
        zoneMap.set(key, {
            "Zone Name": zone,
            Scheme: scheme,
            Block: block,
            District: district
        });
    }
});
const zoneMaster = Array.from(zoneMap.values());

// // --- 4. Pump Master Excel ---
// const pumpMaster = data.map(row => ({
//     "Pump House Name": row["Pump House No"],
//     "Pump House Type": row["Pump Type"],
//     Scheme: row["PWSS"],
//     Block: row["Block"],
//     District: row["District"],
//     Zone: row["Zone"] || "N/A",
//     Latitude: row["Latitude"],
//     Longitude: row["Longitude"]
// }));


// --- 4. Pump Master Excel ---
const pumpMaster = data.map(row => {
    let type = ""; // Declare type before using it
    const pumpType = row["Pump Type"].trim(); // Trim spaces from input

    if (pumpType === "Basic") {
        type = "type_a";
    } else if (pumpType === "Intermediate") {
        type = "type_b";
    }

    return { // Ensure return statement
        "Pump House Name": row["Pump House No"],
        "Pump House Type": type, // Properly assigned type
        State: STATE,
        District: row["District"],
        Block: row["Block"],
        Scheme: row["PWSS"],
        Zone: row["Zone"] || "N/A",
        Latitude: row["Latitude"],
        Longitude: row["Longitude"]
    };
});



// --- Write all outputs ---
function writeExcel(filename, data) {
    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
    xlsx.writeFile(wb, `./output/${filename}`);
}

// Ensure output directory exists
if (!fs.existsSync('./output')) fs.mkdirSync('./output');

writeExcel("Block_Master.xlsx", blockMaster);
writeExcel("Scheme_Master.xlsx", schemeMaster);
writeExcel("Zone_Master.xlsx", zoneMaster);
writeExcel("Pump_Master.xlsx", pumpMaster);

console.log("✅ All Excel files created in ./output");
