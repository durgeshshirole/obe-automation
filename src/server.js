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

        // Identify student data start & end rows
        let studentStartRow = 0;
        let studentEndRow = 0;
        worksheet.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(1).value;

            if (typeof firstCellValue === "number" && studentStartRow === 0) {
                studentStartRow = rowNumber; // First student row
            }
            if (studentStartRow > 0 && (firstCellValue === null || firstCellValue === "") && studentEndRow === 0) {
                studentEndRow = rowNumber - 1; // Last student row
            }
        });

        if (studentStartRow === 0) {
            return res.status(400).json({ error: "No student data found" });
        }
        if (studentEndRow === 0) {
            studentEndRow = worksheet.rowCount; // If no empty row is found, process till the last row
        }

        // Define formulas for relevant columns (Fixed Formula Structure)
        const formulaMappings = {
            5: `IF(D{row}>=8,"Y",'CO internal ATTAINMENT'!K1)`, // Column E
            7: `IF(F{row}>=7,"Y",'CO internal ATTAINMENT'!K1)`, // Column G
            9: `IF(H{row}>=18,"Y",'CO internal ATTAINMENT'!K1)`, // Column I
            13: `IF(L{row}>=2,"Y",'CO internal ATTAINMENT'!K1)`  // Column M
        };

        // Apply formulas only to student rows
        for (let rowIndex = studentStartRow; rowIndex <= studentEndRow; rowIndex++) {
            Object.keys(formulaMappings).forEach((col) => {
                const cell = worksheet.getCell(rowIndex, Number(col));
                if (!cell.value) { // Only update if the cell is empty
                    cell.value = { formula: formulaMappings[col].replace("{row}", rowIndex) };
                }
            });
        }

        // Preserve formatting and styles
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.style = { ...cell.style }; // Preserve original style
            });
        });

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
