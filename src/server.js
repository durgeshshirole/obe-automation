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

        // Process "CO internal ATTAINMENT"
        const sheetName1 = "CO internal ATTAINMENT";
        const worksheet1 = workbook.getWorksheet(sheetName1);
        
        if (worksheet1) {
            let studentStartRow = 0;
            let studentEndRow = 0;
            worksheet1.eachRow((row, rowNumber) => {
                const firstCellValue = row.getCell(1).value;
                if (typeof firstCellValue === "number") {
                    if (studentStartRow === 0) studentStartRow = rowNumber;
                    studentEndRow = rowNumber;
                }
            });

            const formulaMappings = {
                5: `IF(D{row}>=8,"Y","N")`,
                7: `IF(F{row}>=7,"Y","N")`,
                9: `IF(H{row}>=18,"Y","N")`,
                11: `IF(J{row}>=16,"Y","N")`,
                13: `IF(L{row}>=2,"Y","N")`,
                17: `IF(P{row}>=7,"Y","N")`,
                19: `IF(R{row}>=18,"Y","N")`,
                21: `IF(T{row}>=15,"Y","N")`,
                23: `IF(V{row}>=2,"Y","N")`,
                25: `((P{row}+R{row}+T{row})/3)*0.85 + V{row}*0.15`,
                27: `IF(Z{row}>=7,"Y","N")`,
                29: `IF(AB{row}>=7,"Y","N")`,
                31: `IF(AD{row}>=18,"Y","N")`,
                33: `IF(AF{row}>=2,"Y","N")`,
                35: `IF(AH{row}>2,"Y","N")`,
                39: `IF(AL{row}>=7,"Y","N")`,
                41: `IF(AN{row}>=7,"Y","N")`,
                43: `IF(AP{row}>=16,"Y","N")`
            };
            
            for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
                worksheet1.getCell(`N${rowIndex}`).value = {
                    formula: `=((D${rowIndex}+F${rowIndex}+H${rowIndex})/3)*0.75 + J${rowIndex}*0.15 + L${rowIndex}*0.1`
                };
                Object.keys(formulaMappings).forEach((col) => {
                    const cell = worksheet1.getCell(rowIndex, Number(col));
                    if (!cell.value || cell.value === "") {
                        cell.value = { formula: formulaMappings[col].replace(/{row}/g, rowIndex) };
                    }
                });
            }
        }

        // Process "PO PSO SPPU ATT"
        const sheetName2 = "PO PSO SPPU ATT";
        const worksheet2 = workbook.getWorksheet(sheetName2);
        
        if (worksheet2) {
            worksheet2.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    if (cell.formula) {
                        row.getCell(colNumber).value = ""; // Remove formula values
                    }
                });
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
