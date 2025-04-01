const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const port = 3000;

// Ensure necessary directories exist
const uploadDir = path.join(__dirname, "uploads");
const outputDir = path.join(__dirname, "processed_files");
const templateDir = path.join(__dirname, "template");

if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);
if (!fs.existsSync(templateDir)) fs.mkdirSync(templateDir);

// Multer storage: Save uploaded file with its original name
const storage = multer.diskStorage({
    destination: uploadDir,
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    },
});
const upload = multer({ storage });

// Define template file path
const templatePath = path.join(templateDir, "Template.xlsx");
const outputFilePath = path.join(outputDir, "Processed_Template.xlsx");

// Function to check if a file exists
const fileExists = (filePath) => fs.existsSync(filePath);

// Column mappings for Sheet 4
const columnMappingsSheet4 = {
    "B": "B",
    "C": "C", 
    "D": "D", 
    "E": "F", 
    "O": "H", 
    "AG": "J",
    "U": "L", 
    "AA": "N", 
    "F": "R", //UNIT 2
    "P": "T", 
    "AH": "V",
     "V": "X",
    "AB": "Z", 
    "J": "AD", //UNIT 3
    "H": "AF", 
    "Q": "AH", 
    "AI": "AJ", 
    "W": "AL",
    "AC": "AN", 
    "K": "AT", //UNIT 4
    "I": "AV", 
    "R": "AX", 
    "X": "BB", 
    "AD": "BD", 
    "L": "BH" ,//UNIT 5
    "S":"BL",
    "AK":"BN",
    "Y":"BP",
    "AE":"BR",
    "M":"BV", //UNIT 6
    "T":"BX",
    "AL":"BZ",
    "Z":"CB",
    "AF":"CD"
};

// Column mappings for Sheet 5
const columnMappingsSheet5 = {
    "B": "B",
    "C": "C",
    "G":"D",
    "N":"E"
}; 

// Mapping for additional sheet (PO-PSO int and Ext att)
const additionalMappings = {
   "H2":"H2",
   "A8":["A8", "A19"],
   "B8":["B8", "B19"],
   "C8":["C8", "C19"],
   "D8":["D8", "D19"],
   "F8":"F8",
   "F9":"F9",
   "F10":"F10",
   "F11":"F11",
   "F12":"F12",
   "F13":"F13",
   "G8":"G8",
   "G9":"G9",
   "G10":"G10",
   "G11":"G11",
   "G12":"G12",
   "G13":"G13",
   "H8":"H8",
   "H9":"H9",
   "H10":"H10",
   "H11":"H11",
   "H12":"H12",
   "H13":"H13",
   "G8": "G8",
    "G9": "G9",
    "G10": "G10",
    "G11": "G11",
    "G12": "G12",
    "G13": "G13",
    "I8": "I8",
    "I9": "I9",
    "I10": "I10",
    "I11": "I11",
    "I12": "I12",
    "I13": "I13",
    "J8": "J8",
    "J9": "J9",
    "J10": "J10",
    "J11": "J11",
    "J12": "J12",
    "J13": "J13",
    "K8": "K8",
    "K9": "K9",
    "K10": "K10",
    "K11": "K11",
    "K12": "K12",
    "K13": "K13",
    "L8": "L8",
    "L9": "L9",
    "L10": "L10",
    "L11": "L11",
    "L12": "L12",
    "L13": "L13",
    "M8": "M8",
    "M9": "M9",
    "M10": "M10",
    "M11": "M11",
    "M12": "M12",
    "M13": "M13",
    "N8": "N8",
    "N9": "N9",
    "N10": "N10",
    "N11": "N11",
    "N12": "N12",
    "N13": "N13",
    "O8": "O8",
    "O9": "O9",
    "O10": "O10",
    "O11": "O11",
    "O12": "O12",
    "O13": "O13",
    "P8": "P8",
    "P9": "P9",
    "P10": "P10",
    "P11": "P11",
    "P12": "P12",
    "P13": "P13",
    "Q8": "Q8",
    "Q9": "Q9",
    "Q10": "Q10",
    "Q11": "Q11",
    "Q12": "Q12",
    "Q13": "Q13",
    "R8": "R8",
    "R9": "R9",
    "R10": "R10",
    "R11": "R11",
    "R12": "R12",
    "R13": "R13",
    "S8": "S8",
    "S9": "S9",
    "S10": "S10",
    "S11": "S11",
    "S12": "S12",
    "S13": "S13",
    "T8": "T8",
    "T9": "T9",
    "T10": "T10",
    "T11": "T11",
    "T12": "T12",
    "T13": "T13"
    
};

// Function to process a sheet
const processSheet = (inputSheet, templateSheet, columnMappings) => {
    let inputStartRow = null;
    inputSheet.eachRow((row, rowNumber) => {
        if (row.getCell(1).value === 1 && inputStartRow === null) {
            inputStartRow = rowNumber;
        }
    });
    if (inputStartRow === null) return false;

    let templateStartRow = null;
    templateSheet.eachRow((row, rowNumber) => {
        if (row.getCell(1).value === 1 && templateStartRow === null) {
            templateStartRow = rowNumber;
        }
    });
    if (templateStartRow === null) return false;

    const inputRowCount = inputSheet.rowCount - inputStartRow + 1;
    const templateRowCount = templateSheet.rowCount - templateStartRow + 1;

    if (inputRowCount > templateRowCount) {
        for (let i = templateRowCount; i < inputRowCount; i++) {
            templateSheet.insertRow(templateStartRow + i);
        }
    }

    for (let i = inputStartRow; i <= inputSheet.rowCount; i++) {
        const inputRow = inputSheet.getRow(i);
        const templateRowNumber = templateStartRow + (i - inputStartRow);

        Object.keys(columnMappings).forEach(inputCol => {
            const templateCol = columnMappings[inputCol];
            const inputValue = inputRow.getCell(inputCol).value;

            if (inputValue !== null) {
                templateSheet.getCell(`${templateCol}${templateRowNumber}`).value = inputValue;
            }
        });
    }
    return true;
};

// API to process uploaded file and fill the template for Sheet 4 and Sheet 5
app.post("/process-excel", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: "No file uploaded" });
        }

        const inputFilePath = path.join(uploadDir, req.file.filename);
        if (!fileExists(templatePath) || !fileExists(inputFilePath)) {
            return res.status(400).json({ error: "Template or input file not found" });
        }

        const inputWorkbook = new ExcelJS.Workbook();
        await inputWorkbook.xlsx.readFile(inputFilePath);
        const inputSheet = inputWorkbook.worksheets[0];

        const templateWorkbook = new ExcelJS.Workbook();
        await templateWorkbook.xlsx.readFile(templatePath);
        
        // Process Sheet 4
        const templateSheet4 = templateWorkbook.getWorksheet("CO internal ATTAINMENT");
        if (!templateSheet4) {
            return res.status(400).json({ error: "Sheet 4 not found in template" });
        }
        const sheet4Processed = processSheet(inputSheet, templateSheet4, columnMappingsSheet4);
        
        // Process Sheet 5
        const templateSheet5 = templateWorkbook.getWorksheet(" PO PSO SPPU ATT ");
        if (!templateSheet5) {
            return res.status(400).json({ error: "Sheet 5 not found in template" });
        }
        const sheet5Processed = processSheet(inputSheet, templateSheet5, columnMappingsSheet5);
        
        if (!sheet4Processed || !sheet5Processed) {
            return res.status(400).json({ error: "Could not find starting row in input or template file" });
        }

        // Save intermediate processed file
        await templateWorkbook.xlsx.writeFile(outputFilePath);

        // ==== PO-PSO int and Ext att Sheet Processing ====
        await templateWorkbook.xlsx.readFile(outputFilePath); // Reload updated template
        const templateSheet = templateWorkbook.getWorksheet(" PO-PSO int and Ext att");
        const inputSheet2 = inputWorkbook.getWorksheet("Sheet2");

        if (!inputSheet2) {
            return res.status(400).json({ error: "Sheet2 not found in input.xlsx" });
        }
        if (!templateSheet) {
            return res.status(400).json({ error: "PO-PSO int and Ext att not found in Processed_Template.xlsx" });
        }

        // Mapping logic
        Object.keys(additionalMappings).forEach(inputCell => {
            const inputValue = inputSheet2.getCell(inputCell).value;
            let templateCells = additionalMappings[inputCell];
        
            // console.log(`Mapping ${inputCell} (${inputValue}) to ${templateCells}`);
        
            if (!Array.isArray(templateCells)) {
                templateCells = [templateCells]; // Convert single values to arrays
            }
        
            for (let i = 0; i < templateCells.length; i++) {
                let templateCell = templateCells[i];
                // console.log(`Setting ${templateCell} to ${inputValue}`);
                templateSheet.getCell(templateCell).value = inputValue; // Assign value to each mapped cell
            }
        });
        
        
        

        // Save final processed file
        await templateWorkbook.xlsx.writeFile(outputFilePath);

        // return res.json({ message: "Sheets processed successfully", file: outputFilePath });
        console.log("step1:Sheets processed successfully")

    } catch (error) {
        res.status(500).json({ error: "Internal server error", details: error.message });
    }





// Process Excel file

    try {
        console.log("inside step2")
        const filePath = path.join(__dirname, "processed_files/Processed_Template.xlsx");
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
            { numerator: "AH", denominator: "AG", result: "AG" },
            { numerator: "AJ", denominator: "AI", result: "AI" },
            { numerator: "AL", denominator: "AK", result: "AK" },
            { numerator: "AN", denominator: "AM", result: "AM" },
            { numerator: "AP", denominator: "AO", result: "AO" },
        ];

        columnMappings.forEach(({ numerator, denominator, result }) => {
            const numCell = worksheet5.getCell(`${numerator}${rowSource}`).value;
            const denCell = worksheet5.getCell(`${denominator}${rowSource}`).value;

            // Convert values to numbers or null if they can't be parsed
            const numValue = isNaN(parseFloat(numCell)) ? null : parseFloat(numCell);
            const denValue = isNaN(parseFloat(denCell)) ? null : parseFloat(denCell);

            if (numValue === null || denValue === null || denValue === 0) {
                worksheet5.getCell(`${result}${rowTarget}`).value = "-";
            } else {
                worksheet5.getCell(`${result}${rowTarget}`).value = {
                    formula: `IF(OR(ISBLANK(${numerator}${rowSource}), ISBLANK(${denominator}${rowSource}), ${denominator}${rowSource}=0), "-", ROUND(${numerator}${rowSource} / ${denominator}${rowSource} * 100, 2))`
                };
            }

            // Auto-adjust column width
            worksheet5.getColumn(result).width = 15;
        });





        // processing table po pso attainment for ext sppu attainment

        const columnMapping = {
            11: 13,  // K14 → M8
            12: 15,  // L14 → O8
            13: 17,  // M14 → Q8
            14: 19,  // N14 → S8
            15: 21,  // O14 → U8
            16: 23,  // P14 → W8
            17: 25,  // Q14 → Y8
            18: 27,  // R14 → AA8
            19: 29,  // S14 → AC8
            20: 31,  // T14 → AE8
            21: 33,  // U14 → AG8
            22: 35,  // V14 → AI8
            23: 37,  // W14 → AK8
            24: 39,  // X14 → AM8
            25: 41   // Y14 → AO8
        };


        Object.keys(columnMapping).forEach((targetCol) => {
            const sourceCol = columnMapping[targetCol];

            // Convert column numbers to Excel letter notation
            const sourceCellRef = worksheet5.getCell(8, sourceCol).address;
            const targetCellRef = worksheet5.getCell(14, Number(targetCol)).address;

            // Debugging Log: Check if source cell has a value
            // console.log(`Mapping ${targetCellRef} → ${sourceCellRef}`);

            // Set formula dynamically
            worksheet5.getCell(14, Number(targetCol)).value = { formula: `=${sourceCellRef}` };
        });


        // Processing sheet 4 (CO internal ATTAINMENT)
        worksheet4.eachRow((row, rowNumber) => {
            const firstCellValue = row.getCell(2).value; // Assuming roll number is in Column B

            if (typeof firstCellValue === "number") {
                if (studentStartRow === 0) {
                    studentStartRow = rowNumber;
                }
                studentEndRow = rowNumber;
            }
        });



        // **Preserved Formula Mappings**
        const formulaMappings = {
            5: `IF(D{row}>=8,"Y","N")`,
            7: `IF(F{row}>=7,"Y","N")`,
            9: `IF(H{row}>=18,"Y","N")`,
            11: `IF(J{row}>=16,"Y","N")`,
            13: `IF(L{row}>=2,"Y","N")`,
            14: `((D{row}+F{row}+H{row})/3)*0.75 + J{row}*0.15 + L{row}*0.1`,
            15: `IF(N{row}>=12,"Y","N")`,
            17: `IF(P{row}>=7,"Y","N")`,
            19: `IF(R{row}>=18,"Y","N")`,
            21: `IF(T{row}>=15,"Y","N")`,
            23: `IF(V{row}>=2,"Y","N")`,
            24: `((P{row}+R{row}+T{row})/3)*0.85 + V{row}*0.15`,
            25: `IF(X{row}>=13.32,"Y","N")`,
            27: `IF(Z{row}>=7,"Y","N")`,
            29: `IF(AB{row}>=7,"Y","N")`,
            31: `IF(AD{row}>=18,"Y","N")`,
            33: `IF(AF{row}>=2,"Y","N")`,
            35: `IF(AH{row}>2,"Y","N")`,
            36: `((Z{row}+AD{row}+AH{row})/3)*0.85 + AF{row}*0.15`,
            37: `IF(AJ{row}>=8,"Y","N")`,
            39: `IF(AL{row}>=7,"Y","N")`,
            41: `IF(AN{row}>=7,"Y","N")`,
            43: `IF(AP{row}>=16,"Y","N")`,
            45: `IF(AR{row}>=20,"Y","N")`,
            47: `IF(AT{row}>=2,"Y","N")`,
            48: `((AL{row}+AN{row}+AP{row})/3)*0.75 + AR{row}*0.15 + AT{row}*0.1`,
            49: `IF(AV{row}>=8,"Y","N")`,
            51: `IF(AX{row}>=5,"Y","N")`,
            53: `IF(AZ{row}>=7,"Y","N")`,
            55: `IF(BB{row}>=16,"Y","N")`,
            57: `IF(BD{row}>=2,"Y","N")`,
            58: `((AX{row}+AZ{row}+BB{row})/3)*0.75 + BD{row}*0.15`,
            59: `IF(BF{row}>=8,"Y","N")`,
            61: `IF(BH{row}>=5,"Y","N")`,
            63: `IF(BJ{row}>=16,"Y","N")`,
            65: `IF(BL{row}>=20,"Y","N")`,
            67: `IF(BN{row}>=2,"Y","N")`,
            68: `((BH{row}+BJ{row}+BL{row})/3)*0.85 + BN{row}*0.15`,
            69: `IF(BP{row}>=8,"Y","N")`

        };

        // **Apply formulas dynamically**
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

        if (!res.headersSent) {
            res.json({ 
                message: "File processed successfully", 
                downloadPath: "/processed_files/Processed_File.xlsx" 
            });
            return; // Prevent further execution
        }
    } catch (error) {
        if (!res.headersSent) {
            res.status(500).json({ error: error.message });
        }
    }
});

// Serve processed files
console.log("step2 completed")
app.use("/processed_files", express.static(path.join(__dirname, "processed_files")));

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
