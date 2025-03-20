const xlsx = require("xlsx");

exports.mergeExcel = (templatePath, inputPath) => {
    const templateWorkbook = xlsx.readFile(templatePath);
    const inputWorkbook = xlsx.readFile(inputPath);

    const templateSheet = templateWorkbook.Sheets[templateWorkbook.SheetNames[0]];
    const inputSheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];

    const templateData = xlsx.utils.sheet_to_json(templateSheet);
    const inputData = xlsx.utils.sheet_to_json(inputSheet);

    // Merge data
    inputData.forEach((student) => {
        let entry = templateData.find((e) => e.RollNo === student.RollNo);
        if (!entry) templateData.push(student);
        else Object.assign(entry, student);
    });

    // Calculate subject-wise average
    const totalStudents = templateData.length;
    const subjectSums = {};
    const subjectCounts = {};

    templateData.forEach((entry) => {
        Object.keys(entry).forEach((key) => {
            if (key !== "RollNo" && key !== "Name") {
                if (!subjectSums[key]) {
                    subjectSums[key] = 0;
                    subjectCounts[key] = 0;
                }
                subjectSums[key] += Number(entry[key]) || 0;
                subjectCounts[key] += entry[key] ? 1 : 0;
            }
        });
    });

    // Add SPOSMarks row
    const avgRow = { RollNo: "SPOSMarks", Name: "Average" };
    Object.keys(subjectSums).forEach((subject) => {
        avgRow[subject] = subjectCounts[subject] > 0 ? (subjectSums[subject] / subjectCounts[subject]).toFixed(2) : 0;
    });

    templateData.push(avgRow);

    // Convert back to worksheet
    const updatedSheet = xlsx.utils.json_to_sheet(templateData);
    templateWorkbook.Sheets[templateWorkbook.SheetNames[0]] = updatedSheet;

    return templateWorkbook;
};
