import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

/* ================= HEALTH ================= */
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

/* ================= DLL MACHINE EXPORT ================= */
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

    const templatePath = path.join(process.cwd(), "DLL_MACHINE.xlsx");
    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: "DLL_MACHINE.xlsx not found" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);
    const sheet = wb.getWorksheet("DLL_DATA");

    const write = (cell, value) => {
      sheet.getCell(cell).value = value || "";
    };

    /* ===== HEADER ===== */
    write("B1", teacherName);
    write("B2", gradeLevel);
    write("B3", subject);
    write("B4", quarter);
    write("B5", weekDate);

    /* ===== OBJECTIVES ===== */
    const objectives = generatedLesson.I_Objectives || [];
    write("B7", objectives[0]);
    write("B8", objectives[1]);
    write("B9", objectives[2]);

    /* ===== CONTENT & RESOURCES ===== */
    write("B11", generatedLesson.II_Content);
    write("B12", (generatedLesson.III_LearningResources || []).join("\n"));

    /* ===== PROCEDURES (PER DAY VARIATION) ===== */
    const procedures = generatedLesson.IV_Procedures || [];

    const dayMap = {
      Monday: 15,
      Tuesday: 23,
      Wednesday: 31,
      Thursday: 39,
      Friday: 47
    };

    Object.values(dayMap).forEach(startRow => {
      for (let i = 0; i < 7; i++) {
        write(`B${startRow + i}`, procedures[i] || "");
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
      'attachment; filename="DLL_MACHINE_FILLED.xlsx"'
    );

    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ DLL Machine Service running on port ${PORT}`);
});