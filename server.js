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

        // Identify where student rows start (first row with a numeric roll number)
        let studentStartRow = 0;
        worksheet.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(1).value;
            if (typeof firstCellValue === "number" && studentStartRow === 0) {
                studentStartRow = rowNumber;
            }
        });

        if (studentStartRow === 0) {
            return res.status(400).json({ error: "No student data found" });
        }

        // Define formulas for relevant columns
        const formulaMappings = {
            5: "IF(VALUE(D{row})>=8,\"Y\",\"N\")", // Column E
            7: "IF(VALUE(F{row})>=7,\"Y\",\"N\")", // Column G
            9: "IF(VALUE(H{row})>=18,\"Y\",\"N\")", // Column I
            13: "IF(VALUE(L{row})>=2,\"Y\",\"N\")" // Column M
        };

        // Apply formulas to the correct rows
        for (let rowIndex = studentStartRow; rowIndex <= worksheet.rowCount; rowIndex++) {
            Object.keys(formulaMappings).forEach((col) => {
                worksheet.getCell(rowIndex, Number(col)).value = { formula: formulaMappings[col].replace("{row}", rowIndex) };
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
