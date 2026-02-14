const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

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

    const templatePath = path.join(__dirname, "DLL_MACHINE.xlsx");
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
    write("B7", objectives[0] || "");
    write("B8", objectives[1] || "");
    write("B9", objectives[2] || "");

    /* ===== CONTENT & RESOURCES ===== */
    write("B11", generatedLesson.II_Content || "");
    write(
      "B12",
      Array.isArray(generatedLesson.III_LearningResources)
        ? generatedLesson.III_LearningResources.join("\n")
        : generatedLesson.III_LearningResources || ""
    );

    /* ===== PROCEDURES ===== */
    const steps = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const [
      motivation,
      presentation,
      discussion,
      practice,
      generalization,
      assessment,
      assignment
    ] = steps;

    const dayPlan = {
      Monday: [motivation, generalization],
      Tuesday: [presentation, discussion],
      Wednesday: [practice],
      Thursday: [assignment],
      Friday: [assessment, generalization]
    };

    const dayRowMap = {
      Monday: 15,
      Tuesday: 23,
      Wednesday: 31,
      Thursday: 39,
      Friday: 47
    };

    Object.entries(dayRowMap).forEach(([day, startRow]) => {
      const items = dayPlan[day] || [];
      for (let i = 0; i < 7; i++) {
        write(`B${startRow + i}`, items[i] || "");
      }
    });

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
    console.error("DLL ERROR:", err);
    res.status(500).json({ error: err.message });
  }
});

/* ================= START ================= */
app.listen(PORT, () => {
  console.log(`✅ DLL Machine Service running on port ${PORT}`);
});