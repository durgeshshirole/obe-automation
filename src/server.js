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

        const sheetName = " PO PSO SPPU ATT "; // Working on the 5th sheet
        const worksheet5 = workbook.getWorksheet(sheetName);
        const sheet4 = "CO internal ATTAINMENT";
        const worksheet4 = workbook.getWorksheet(sheet4);

        if (!worksheet5) {
            return res.status(400).json({ error: `Sheet '${sheetName5}' not found` });
        }

        // Identify student data start & end rows based on roll number (assumed to be in column A)
        let studentStartRow = 0;
        let studentEndRow = 0;

        worksheet5.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(2).value; // Assuming roll number is in Column A

            if (typeof firstCellValue === "number") {
                if (studentStartRow === 0) {
                    studentStartRow = rowNumber; // Set the start row at the first number found
                }
                studentEndRow = rowNumber; // Keep updating this to get the last valid roll number row
            }
        });

        if (studentStartRow === 0 || studentEndRow === 0) {
            return res.status(400).json({ error: "No student data found" });
        }

        // **✅ Apply Formula for Theory Marks (Column F) without removing any existing logic**
        for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
            worksheet5.getCell(`F${rowIndex}`).value = {
                formula: `D${rowIndex} + E${rowIndex}`
            };
        }
        
        
        const rowTarget = 8; // PO % ATT row
        const rowSource = 7; // A by M percentage source row
    
        // Column pairs for applying the formula
        const columnMappings = [
            { numerator: "N", denominator: "M", result: "M" },
            { numerator: "P", denominator: "O", result: "O" },
            { numerator: "R", denominator: "Q", result: "Q" },
            { numerator: "T", denominator: "S", result: "S" },
            { numerator: "V", denominator: "U", result: "U" },
            { numerator: "X", denominator: "W", result: "W" },
            { numerator: "Z", denominator: "Y", result: "Y" },
            { numerator: "AB", denominator: "AA", result: "AA" },
            { numerator: "AD", denominator: "AC", result: "AC" },
            { numerator: "AF", denominator: "AE", result: "AE" },
            { numerator: "AH", denominator: "AG", result: "AG" }
        ];
    
        columnMappings.forEach(({ numerator, denominator, result }) => {
            worksheet5.getCell(`${result}${rowTarget}`).value = {
                formula: `${numerator}${rowSource}/${denominator}${rowSource}*100`
            };
        });


        worksheet4.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(2).value; // Assuming roll number is in Column A

            if (typeof firstCellValue === "number") {
                if (studentStartRow === 0) {
                    studentStartRow = rowNumber; // Set the start row at the first number found
                }
                studentEndRow = rowNumber; // Keep updating this to get the last valid roll number row
            }
        });

        // **Existing formula mappings (preserved)**
        const formulaMappings = {
            5: `IF(D{row}>=8,"Y","N")`,  // Column E
            7: `IF(F{row}>=7,"Y","N")`,  // Column G
            9: `IF(H{row}>=18,"Y","N")`, // Column I
            11: `IF(J{row}>=16,"Y","N")`,  // Column K
            13: `IF(L{row}>=2,"Y","N")`,   // Column M
            17: `IF(P{row}>=7,"Y","N")`,   // Column Q
            19: `IF(R{row}>=18,"Y","N")`,  // Column S
            21: `IF(T{row}>=15,"Y","N")`,  // Column U
            23: `IF(V{row}>=2,"Y","N")`,   // Column W
            25: `((P{row}+R{row}+T{row})/3)*0.85 + V{row}*0.15`,// Column Y
            
            27: `IF(Z{row}>=7,"Y","N")`,  // Column AA
            29: `IF(AB{row}>=7,"Y","N")`, // Column AC
            31: `IF(AD{row}>=18,"Y","N")`, // Column AE
            33: `IF(AF{row}>=2,"Y","N")`,  // Column AG
            35: `IF(AH{row}>2,"Y","N")`,   // Column AI
            39: `IF(AL{row}>=7,"Y","N")`,  // Column AM
            41: `IF(AN{row}>=7,"Y","N")`,  // Column AO
            43: `IF(AP{row}>=16,"Y","N")`  // Column AQ
        };

        // **Preserved: Apply formulas dynamically**
        for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
            Object.keys(formulaMappings).forEach((col) => {
                const cell = worksheet4.getCell(rowIndex, Number(col));
                if (!cell.value || cell.value === "") {
                    cell.value = { formula: formulaMappings[col].replace(/{row}/g, rowIndex) };
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
