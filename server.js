import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

/* ================= MERGED CELL SAFE WRITE ================= */
function writeMerged(sheet, row, startCol, endCol, value) {
  // write to the top-left cell of merged range
  const cell = sheet.getCell(row, startCol);
  cell.value = value || "";
  cell.alignment = { wrapText: true, vertical: "top" };

  // ensure merged (safe even if already merged)
  try {
    sheet.mergeCells(row, startCol, row, endCol);
  } catch {}
}

/* ================= HEALTH ================= */
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

/* ================= DLL EXPORT ================= */
app.post("/fill-dll", async (req, res) => {
  try {
    const {
      teacherName,
      gradeLevel,
      subject,
      quarter,
      weekDate,
      generatedLesson
    } = req.body;

    if (!generatedLesson) {
      return res.status(400).json({ error: "Missing generatedLesson" });
    }

    const templatePath = path.join(process.cwd(), "DLL_FORMAT.xlsx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "DLL_FORMAT.xlsx not found" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.getWorksheet(1);

    /* ===== HEADER (ACTUAL INPUT AREA: D–H) ===== */
    writeMerged(sheet, 5, 4, 8, teacherName);   // D5:H5
    writeMerged(sheet, 6, 4, 8, gradeLevel);    // D6:H6
    writeMerged(sheet, 7, 4, 8, subject);       // D7:H7
    writeMerged(sheet, 8, 4, 8, quarter);       // D8:H8
    writeMerged(sheet, 9, 4, 8, weekDate);      // D9:H9

    /* ===== I. OBJECTIVES (D12:H14) ===== */
    const objectives = Array.isArray(generatedLesson.I_Objectives)
      ? generatedLesson.I_Objectives.join("\n")
      : generatedLesson.I_Objectives || "";

    writeMerged(sheet, 12, 4, 8, objectives);

    /* ===== II. CONTENT (D16:H16) ===== */
    writeMerged(sheet, 16, 4, 8, generatedLesson.II_Content || "");

    /* ===== III. LEARNING RESOURCES (D18:H20) ===== */
    const resources = Array.isArray(generatedLesson.III_LearningResources)
      ? generatedLesson.III_LearningResources.join("\n")
      : generatedLesson.III_LearningResources || "";

    writeMerged(sheet, 18, 4, 8, resources);

    /* ===== IV. PROCEDURES – WEEKLY (D–H) ===== */
    const procedures = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const dayRowMap = {
      Monday: 23,
      Tuesday: 31,
      Wednesday: 39,
      Thursday: 47,
      Friday: 55
    };

    Object.values(dayRowMap).forEach(startRow => {
      for (let i = 0; i < 7; i++) {
        writeMerged(sheet, startRow + i, 4, 8, procedures[i] || "");
      }
    });

    /* ===== EXPORT ===== */
    const buffer = await wb.xlsx.writeBuffer();
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
    console.error("DLL ERROR:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ DLL Excel Fill Service running on port ${PORT}`);
});