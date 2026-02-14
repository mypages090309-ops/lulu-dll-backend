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
/**
 * POST /fill-dll
 * Expects:
 * {
 *   teacherName,
 *   gradeLevel,
 *   subject,
 *   quarter,
 *   weekDate,
 *   generatedLesson
 * }
 */
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
      return res.status(400).json({
        error: "Missing generatedLesson data"
      });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("./DLL_FORMAT.xlsx");

    const sheet1 = workbook.getWorksheet(1);
    const sheet2 = workbook.getWorksheet(2);

    /* ================= HELPERS ================= */
    const safe = (v) => v ?? "";

    const joinIfArray = (v) =>
      Array.isArray(v) ? v.join("\n") : safe(v);

    const setCell = (sheet, cell, value) => {
      const c = sheet.getCell(cell);
      c.value = value;
      c.alignment = {
        wrapText: true,
        vertical: "top"
      };
    };

    /* ================= TABLE 1 ================= */

    // HEADER
    setCell(sheet1, "F3", teacherName);
    setCell(sheet1, "I2", gradeLevel);
    setCell(sheet1, "I3", subject);
    setCell(sheet1, "I4", quarter);
    setCell(sheet1, "F4", weekDate);

    // OBJECTIVES
    setCell(sheet1, "C7", safe(generatedLesson.contentStandards));
    setCell(sheet1, "C8", safe(generatedLesson.performanceStandards));
    setCell(sheet1, "C9", joinIfArray(generatedLesson.objectives));

    // CONTENT
    setCell(sheet1, "C11", safe(generatedLesson.topic));

    // LEARNING RESOURCES
    setCell(
      sheet1,
      "C14",
      safe(generatedLesson.references?.teacherGuide)
    );
    setCell(
      sheet1,
      "C15",
      safe(generatedLesson.references?.learnerMaterials)
    );
    setCell(
      sheet1,
      "C16",
      safe(generatedLesson.references?.textbook)
    );
    setCell(
      sheet1,
      "C17",
      safe(generatedLesson.references?.additional)
    );

    // PROCEDURES
    setCell(
      sheet1,
      "C20",
      safe(generatedLesson.procedures?.review)
    );
    setCell(
      sheet1,
      "C21",
      safe(generatedLesson.procedures?.purpose)
    );
    setCell(
      sheet1,
      "C22",
      safe(generatedLesson.procedures?.presentation)
    );
    setCell(
      sheet1,
      "C23",
      safe(generatedLesson.procedures?.practice)
    );
    setCell(
      sheet1,
      "C24",
      safe(generatedLesson.procedures?.generalization)
    );
    setCell(
      sheet1,
      "C25",
      safe(generatedLesson.procedures?.application)
    );
    setCell(
      sheet1,
      "C26",
      safe(generatedLesson.procedures?.evaluation)
    );

    /* ================= TABLE 2 ================= */

    setCell(
      sheet2,
      "C5",
      safe(generatedLesson.reflection?.mastery80)
    );
    setCell(
      sheet2,
      "C6",
      safe(generatedLesson.reflection?.needRemedial)
    );
    setCell(
      sheet2,
      "C7",
      safe(generatedLesson.reflection?.remedialEffective)
    );
    setCell(
      sheet2,
      "C8",
      safe(generatedLesson.reflection?.caughtUp)
    );
    setCell(
      sheet2,
      "C9",
      safe(generatedLesson.reflection?.stillRemedial)
    );
    setCell(
      sheet2,
      "C10",
      safe(generatedLesson.reflection?.strategiesWorked)
    );
    setCell(
      sheet2,
      "C11",
      safe(generatedLesson.reflection?.difficulties)
    );
    setCell(
      sheet2,
      "C12",
      safe(generatedLesson.reflection?.innovations)
    );

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
    res.status(500).json({
      error: err.message
    });
  }
});

/* ================= START SERVER ================= */
app.listen(3000, () => {
  console.log("✅ DLL Excel Fill Service running on port 3000");
});