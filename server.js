import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

// helper: write to merged cell safely
function writeCell(sheet, cell, value) {
  const c = sheet.getCell(cell);
  c.value = value || "";
  c.alignment = { wrapText: true, vertical: "top" };
}

// =======================
// FILL DLL ENDPOINT
// =======================
app.post("/fill-dll", async (req, res) => {
  try {
    const lesson = req.body;

    const templatePath = path.join(process.cwd(), "DLL_FORMAT.xlsx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "DLL template not found" });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const sheet = workbook.worksheets[0];

    // =======================
    // HEADER (STATIC)
    // =======================
    writeCell(sheet, "C5", lesson.teacherName);
    writeCell(sheet, "C6", lesson.gradeLevel);
    writeCell(sheet, "C7", lesson.subject);
    writeCell(sheet, "C8", lesson.quarter);
    writeCell(sheet, "C9", lesson.weekDate);

    // =======================
    // I. OBJECTIVES
    // =======================
    writeCell(sheet, "C12", lesson.objectives?.join("\n"));

    // =======================
    // II. CONTENT
    // =======================
    writeCell(sheet, "C16", lesson.topic);

    // =======================
    // III. LEARNING RESOURCES
    // =======================
    writeCell(sheet, "C18", lesson.resources?.join("\n"));

    // =======================
    // IV. PROCEDURES (ALL DAYS)
    // =======================
    const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

    const procedures = [
      lesson.motivation,
      lesson.presentation,
      lesson.discussion,
      lesson.practice,
      lesson.generalization,
      lesson.assessment,
      lesson.assignment
    ];

    /**
     * DLL FORMAT ROW LOGIC (BASED SA TEMPLATE MO)
     * Monday starts at row 23
     * Each day block = 8 rows
     * A–G = 7 rows
     */
    days.forEach((day, dayIndex) => {
      const baseRow = 23 + dayIndex * 8;

      // Day label
      writeCell(sheet, `B${baseRow - 1}`, day);

      procedures.forEach((text, i) => {
        writeCell(sheet, `C${baseRow + i}`, text);
      });
    });

    // =======================
    // EXPORT
    // =======================
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="DLL_FILLED.xlsx"'
    );

    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "DLL export failed" });
  }
});

app.listen(PORT, () => {
  console.log(`DLL Excel Fill Service running on port ${PORT}`);
});