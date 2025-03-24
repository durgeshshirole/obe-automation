// const express = require("express");
// const excelRoutes = require("./routes/excelRoutes");

// const app = express();
// const PORT = 5000;

// app.use(express.json());
// app.use("/api/excel", excelRoutes);

// app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
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

        const sheetName = "CO internal ATTAINMENT";
        const worksheet = workbook.getWorksheet(sheetName);

        if (!worksheet) {
            return res.status(400).json({ error: `Sheet '${sheetName}' not found` });
        }

        // Identify student data start & end rows based on roll number (assumed to be in column 1)
        let studentStartRow = 0;
        let studentEndRow = 0;

        worksheet.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(1).value; // Assuming roll number is in Column A

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

        // Define formulas for relevant columns
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

        // **Manually set column N formula without mapping issue**
        for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
            worksheet.getCell(`N${rowIndex}`).value = {
                formula: `=((D${rowIndex}+F${rowIndex}+H${rowIndex})/3)*0.75 + J${rowIndex}*0.15 + L${rowIndex}*0.1`
            };
        }

        // Apply formulas from mappings
        for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
            Object.keys(formulaMappings).forEach((col) => {
                if (col !== "12") { // Skip column N since it is manually handled
                    const cell = worksheet.getCell(rowIndex, Number(col));
                    if (!cell.value || cell.value === "") {
                        cell.value = { formula: formulaMappings[col].replace(/{row}/g, rowIndex) };
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
