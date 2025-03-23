const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

// Multer setup for file upload
const storage = multer.diskStorage({
    destination: "uploads/",
    filename: (req, file, cb) => {
        cb(null, "uploaded.xlsx");
    },
});
const upload = multer({ storage });

// Process Excel file
app.post("/upload", upload.single("file"), async (req, res) => {
    try {
        const filePath = path.join(__dirname, "uploads/uploaded.xlsx");
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);

        // Process "PO PSO SPPU ATT" sheet
        const sheetName2 = "PO PSO SPPU ATT";
        const worksheet2 = workbook.getWorksheet(sheetName2);
        
        if (worksheet2) {
            console.log(`Sheet "${sheetName2}" found. Checking data...`);
        
            const attainmentRow = 8; // Row containing PO % Attainment values
            const byDirectRow = 14; // Row where by Direct-Internal Assessment values should be mapped
        
            console.log("Row 8 Values:");
            worksheet2.getRow(attainmentRow).eachCell((cell, colNumber) => {
                console.log(`Column ${colNumber}: ${cell.value}`);
            });
        
            worksheet2.getRow(byDirectRow).eachCell((cell, colNumber) => {
                if (colNumber > 1) { // Skip the first column (header)
                    const attainmentValue = worksheet2.getCell(attainmentRow, colNumber).value;
                    if (attainmentValue !== null && attainmentValue !== undefined) {
                        console.log(`Mapping Column ${colNumber}: ${attainmentValue}`);
                        worksheet2.getCell(byDirectRow, colNumber).value = attainmentValue;
                    } else {
                        console.log(`Column ${colNumber}: No value found in row 8`);
                    }
                }
            });
        }
        

        // Ensure processed files directory exists
        const outputDir = path.join(__dirname, "processed_files");
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        const outputPath = path.join(outputDir, "Processed_File.xlsx");
        await workbook.xlsx.writeFile(outputPath);

        res.json({ message: "File processed successfully", downloadPath: "/processed_files/Processed_File.xlsx" });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Serve processed files
app.use("/processed_files", express.static(path.join(__dirname, "processed_files")));

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});

