import express from "express";
import cors from "cors";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const app = express();
app.use(cors());
app.use(express.json());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ===============================
// POST /fill-dll
// ===============================
app.post("/fill-dll", (req, res) => {
  try {
    const data = req.body;

    // -------------------------------
    // Load template
    // -------------------------------
    const templatePath = path.join(__dirname, "DLL_FORMAT.xlsx");
    if (!fs.existsSync(templatePath)) {
      throw new Error("DLL_FORMAT.xlsx not found");
    }

    const workbook = XLSX.readFile(templatePath);
    const sheetName = "DLL";

    if (!workbook.Sheets[sheetName]) {
      throw new Error("Sheet 'DLL' not found in template");
    }

    const sheet = workbook.Sheets[sheetName];

    // -------------------------------
    // Helper function
    // -------------------------------
    const setCell = (cell, value) => {
      sheet[cell] = { t: "s", v: value ?? "" };
    };

    // ===============================
    // HEADER SECTION
    // ===============================
    setCell("C5", data.teacherName);
    setCell("C6", data.gradeLevel);
    setCell("C7", data.subject);
    setCell("F5", data.quarter);
    setCell("F6", data.weekDate);
    setCell("C8", data.topic);

    // ===============================
    // I. OBJECTIVES
    // ===============================
    (data.objectives || []).forEach((text, i) => {
      setCell(`C${11 + i}`, text);
    });

    // ===============================
    // II. CONTENT
    // ===============================
    setCell("C15", data.contentStandard);
    setCell("C16", data.performanceStandard);
    setCell("C17", data.learningCompetency);

    // ===============================
    // III. LEARNING RESOURCES
    // ===============================
    setCell("C19", data.references?.textbook);
    setCell("C20", data.references?.additional);

    // ===============================
    // IV. PROCEDURES (A–G)
    // ===============================
    const procedures = data.procedures || {};

    const procedureMap = {
      A: "motivation",
      B: "presentation",
      C: "discussion",
      D: "practice",
      E: "generalization",
      F: "evaluation",
      G: "assignment"
    };

    let row = 23;
    Object.keys(procedureMap).forEach(letter => {
      setCell(`A${row}`, letter);
      setCell(`C${row}`, procedures[procedureMap[letter]]);
      row++;
    });

    // ===============================
    // V. REMARKS
    // ===============================
    setCell("C31", data.remarks);

    // ===============================
    // VI. REFLECTION
    // ===============================
    setCell("C33", data.reflection?.learned);
    setCell("C34", data.reflection?.difficulty);
    setCell("C35", data.reflection?.improvement);

    // -------------------------------
    // Save output
    // -------------------------------
    const outputFile = path.join(__dirname, "DLL_FILLED.xlsx");
    XLSX.writeFile(workbook, outputFile);

    res.download(outputFile);

  } catch (err) {
    console.error("DLL ERROR:", err);
    res.status(500).json({
      error: "DLL export failed",
      details: err.message
    });
  }
});

// ===============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DLL Excel Fill Service running on port ${PORT}`);
});