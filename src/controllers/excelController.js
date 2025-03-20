const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const processExcel = async () => {
    const uploadDir = "uploads/";
    const outputDir = "output/";

    // Get the latest uploaded file
    const files = fs.readdirSync(uploadDir);
    if (files.length === 0) throw new Error("No file found in uploads");

    const latestFile = files[files.length - 1]; // Get the last uploaded file
    const filePath = path.join(uploadDir, latestFile);

    // Read and process the file
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert to JSON
    let jsonData = XLSX.utils.sheet_to_json(worksheet);

    // Example: Adding a Total column dynamically
    jsonData.forEach((row) => {
        row["Total"] = (row.Math || 0) + (row.Science || 0) + (row.English || 0);
    });

    // Convert back to Excel
    const newSheet = XLSX.utils.json_to_sheet(jsonData);
    workbook.Sheets[sheetName] = newSheet;

    // Save output file
    const outputFilePath = path.join(outputDir, "processed-output.xlsx");
    XLSX.writeFile(workbook, outputFilePath);

    return outputFilePath; // Return path for download
};

module.exports = { processExcel };
