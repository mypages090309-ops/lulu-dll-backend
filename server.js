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

    /* ===== LOAD TEMPLATE ===== */
    const templatePath = path.join(process.cwd(), "DLL_FORMAT.xlsx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "DLL_FORMAT.xlsx not found" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    // Single sheet (as per your template)
    const sheet = wb.getWorksheet(1);

    /* ================= HEADER (C–G MERGED) ================= */
    write(sheet, "C5", teacherName);
    write(sheet, "C6", gradeLevel);
    write(sheet, "C7", subject);
    write(sheet, "C8", quarter);
    write(sheet, "C9", weekDate);

    /* ================= I. OBJECTIVES (C12:G14) ================= */
    const objectivesText = Array.isArray(generatedLesson.I_Objectives)
      ? generatedLesson.I_Objectives.join("\n")
      : generatedLesson.I_Objectives || "";

    write(sheet, "C12", objectivesText);

    /* ================= II. CONTENT (C16:G16) ================= */
    write(sheet, "C16", generatedLesson.II_Content || "");

    /* ================= III. LEARNING RESOURCES (C18:G20) ================= */
    const resourcesText = Array.isArray(generatedLesson.III_LearningResources)
      ? generatedLesson.III_LearningResources.join("\n")
      : generatedLesson.III_LearningResources || "";

    write(sheet, "C18", resourcesText);

    /* ================= IV. PROCEDURES (WEEKLY – FINAL FIX) ================= */
    const procedures = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    // EXACT row starts based on YOUR DLL_FORMAT.xlsx
    const dayRowMap = {
      Monday: 23,
      Tuesday: 31,
      Wednesday: 39,
      Thursday: 47,
      Friday: 55
    };

    Object.entries(dayRowMap).forEach(([day, startRow]) => {
      // Day label (Column B)
      write(sheet, `B${startRow - 1}`, day);

      // A–G Procedures (C–G merged per row)
      for (let i = 0; i < 7; i++) {
        write(sheet, `C${startRow + i}`, procedures[i] || "");
      }
    });

    /* ================= EXPORT ================= */
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