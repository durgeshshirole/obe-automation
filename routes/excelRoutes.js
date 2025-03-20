const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const { processExcel } = require("../controllers/excelController");

const router = express.Router();

// Ensure the uploads and output folders exist
if (!fs.existsSync("uploads/")) fs.mkdirSync("uploads/");
if (!fs.existsSync("output/")) fs.mkdirSync("output/");

// Multer storage configuration (store files with original names)
const storage = multer.diskStorage({
    destination: "uploads/",
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage });

// 1️⃣ Route to upload and store file (No processing)
router.post("/upload", upload.single("file"), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
    }
    res.json({ message: "File uploaded successfully", filename: req.file.filename });
});

// 2️⃣ Route to process and return output file
router.get("/process", async (req, res) => {
    try {
        const outputFilePath = await processExcel(); // Process the stored file
        res.download(outputFilePath, "processed-output.xlsx");
    } catch (error) {
        res.status(500).json({ error: "Error processing file" });
    }
});

module.exports = router;
