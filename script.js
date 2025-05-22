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
    const block = row["Block"].trim();
    const district = row["District"].trim();
    if (block && !blockMap.has(block)) {
        blockMap.set(block, { "Block Name": block, State: STATE, District: district });
    }
});
const blockMaster = Array.from(blockMap.values());

// --- 2. Scheme Master Excel ---
const schemeMap = new Map();
data.forEach(row => {
    const scheme = row["PWSS"].trim();
    const block = row["Block"].trim();
    const district = row["District"].trim();
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
    const scheme = row["PWSS"].trim();
    const zone = row["Zone"] ? row["Zone"].trim() : "N/A";
    const block = row["Block"].trim();
    const district = row["District"].trim();

    const key = `${scheme}-${zone}`;
    if (!zoneMap.has(key)) {
        zoneMap.set(key, {
            "Zone Name": zone,
            State: STATE,
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
    // let type = ""; // Declare type before using it
    const pumpType = row["Pump Type"].trim(); // Trim spaces from input
    const zone = (row["Zone"] ? row["Zone"].trim() : "N/A") + " */* " + row["PWSS"].trim();

    // if (pumpType === "Basic") {
    //     type = "type_a";
    // } else if (pumpType === "Intermediate") {
    //     type = "type_b";
    // }

    return { // Ensure return statement
        "Pump House Name": row["Pump House No"],
        "Pump House Type": pumpType, // Properly assigned type
        State: STATE,
        District: row["District"].trim(),
        Block: row["Block"].trim(),
        Scheme: row["PWSS"].trim(),
        // Zone: (row["Zone"] ? row["Zone"].trim() : "N/A") + " / " + row["PWSS"].trim(),
        Zone: zone,
        Latitude: row["Latitude"] ? String(row["Latitude"]).trim() : "0.00",
        Longitude: row["Longitude"] ? String(row["Longitude"]).trim() : "0.00"
    };
});


// --- Write all outputs ---
function writeExcel(filename, data) {
    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
    xlsx.writeFile(wb, `./output/${sheetName}/${filename}`);
}

// Ensure output directory exists
if (!fs.existsSync(`./output/${sheetName}`)) fs.mkdirSync(`./output/${sheetName}`);

writeExcel("Block_Master.xlsx", blockMaster);
writeExcel("Scheme_Master.xlsx", schemeMaster);
writeExcel("Zone_Master.xlsx", zoneMaster);
writeExcel("Pump_Master.xlsx", pumpMaster);

console.log(`✅ All Excel files created in ./output/${sheetName}`);
