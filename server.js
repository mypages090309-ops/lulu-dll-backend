import express from "express";
import ExcelJS from "exceljs";
import cors from "cors";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

/* ================= HEALTH CHECK ================= */
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

/* ================= DLL FILL ================= */
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

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("./DLL_FORMAT.xlsx");

    const sheet1 = workbook.getWorksheet(1);
    const sheet2 = workbook.getWorksheet(2);

    /* ================= HELPERS ================= */
    const safe = (v) => (v ? v : "");

    const join = (v) =>
      Array.isArray(v) ? v.join("\n") : safe(v);

    const setCell = (sheet, cell, value) => {
      const c = sheet.getCell(cell);
      c.value = value;
      c.alignment = { wrapText: true, vertical: "top" };
    };

    /* ================= HEADER ================= */
    setCell(sheet1, "F3", teacherName);
    setCell(sheet1, "I2", gradeLevel);
    setCell(sheet1, "I3", subject);
    setCell(sheet1, "I4", quarter);
    setCell(sheet1, "F4", weekDate);

    /* ================= OBJECTIVES ================= */
    setCell(sheet1, "C7", join(generatedLesson.I_Objectives));
    setCell(sheet1, "C8", "");
    setCell(sheet1, "C9", join(generatedLesson.I_Objectives));

    /* ================= CONTENT ================= */
    setCell(sheet1, "C11", safe(generatedLesson.II_Content));

    /* ================= LEARNING RESOURCES ================= */
    setCell(sheet1, "C14", join(generatedLesson.III_LearningResources));
    setCell(sheet1, "C15", "");
    setCell(sheet1, "C16", "");
    setCell(sheet1, "C17", "");

    /* ================= PROCEDURES ================= */
    const proc = generatedLesson.IV_Procedures || {};

    setCell(sheet1, "C20", join(proc.A_Review));
    setCell(sheet1, "C21", join(proc.B_Purpose));
    setCell(sheet1, "C22", join(proc.C_Presentation));
    setCell(sheet1, "C23", join(proc.D_Practice));
    setCell(sheet1, "C24", join(proc.E_Generalization));
    setCell(sheet1, "C25", join(proc.F_Application));
    setCell(sheet1, "C26", join(proc.G_Evaluation));

    /* ================= REFLECTION (SAFE EMPTY) ================= */
    setCell(sheet2, "C5", "");
    setCell(sheet2, "C6", "");
    setCell(sheet2, "C7", "");
    setCell(sheet2, "C8", "");
    setCell(sheet2, "C9", "");
    setCell(sheet2, "C10", "");
    setCell(sheet2, "C11", "");
    setCell(sheet2, "C12", "");

    /* ================= SAVE & DOWNLOAD ================= */
    const fileName = `DLL_${gradeLevel}_${subject}_${Date.now()}.xlsx`
      .replace(/\s+/g, "_");

    const outputPath = path.join(process.cwd(), fileName);
    await workbook.xlsx.writeFile(outputPath);

    res.download(outputPath, fileName, () => {
      fs.unlinkSync(outputPath);
    });

  } catch (err) {
    console.error("DLL ERROR:", err);
    res.status(500).json({ error: err.message });
  }
});

/* ================= START SERVER ================= */
app.listen(3000, () => {
  console.log("✅ DLL Excel Fill Service running on port 3000");
});