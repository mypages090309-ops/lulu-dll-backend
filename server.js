import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

/* ================= HELPER ================= */
function write(sheet, cell, value) {
  const c = sheet.getCell(cell);
  c.value = value || "";
  c.alignment = { wrapText: true, vertical: "top" };
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

    // 🔒 SINGLE SHEET ONLY (AS PER YOUR TEMPLATE)
    const sheet = wb.getWorksheet(1);

    /* ========== HEADER (MERGED C–G) ========== */
    write(sheet, "C5", teacherName);
    write(sheet, "C6", gradeLevel);
    write(sheet, "C7", subject);
    write(sheet, "C8", quarter);
    write(sheet, "C9", weekDate);

    /* ========== I. OBJECTIVES (C12:G14) ========== */
    const objectivesText = Array.isArray(generatedLesson.I_Objectives)
      ? generatedLesson.I_Objectives.join("\n")
      : generatedLesson.I_Objectives || "";

    write(sheet, "C12", objectivesText);

    /* ========== II. CONTENT (C16:G16) ========== */
    write(sheet, "C16", generatedLesson.II_Content || "");

    /* ========== III. LEARNING RESOURCES (C18:G20) ========== */
    const resourcesText = Array.isArray(generatedLesson.III_LearningResources)
      ? generatedLesson.III_LearningResources.join("\n")
      : generatedLesson.III_LearningResources || "";

    write(sheet, "C18", resourcesText);

    /* ========== IV. PROCEDURES – WEEKLY (CRITICAL FIX) ========== */
    const procedures = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

    /**
     * DLL FORMAT LOGIC (BASED ON YOUR TEMPLATE):
     * - Monday block starts at row 23
     * - Each day block = 7 procedure rows + 1 spacer = 8 rows
     * - A–G rows are merged C–G
     */
    days.forEach((day, dayIndex) => {
      const baseRow = 23 + dayIndex * 8;

      // Day label (Column B)
      write(sheet, `B${baseRow - 1}`, day);

      // Procedures A–G
      for (let i = 0; i < 7; i++) {
        write(sheet, `C${baseRow + i}`, procedures[i] || "");
      }
    });

    /* ========== EXPORT ========== */
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

/* ================= START ================= */
app.listen(PORT, () => {
  console.log(`✅ DLL Excel Fill Service running on port ${PORT}`);
});