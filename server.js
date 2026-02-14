const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

app.post("/fill-dll", async (req, res) => {
  try {
    const { teacherName, gradeLevel, subject, quarter, weekDate, generatedLesson } = req.body;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, "DLL_FORMAT.xlsx"));

    const sheet = wb.worksheets[0];

    const writeByName = (name, value) => {
      const ranges = wb.definedNames.getRanges(name);
      if (!ranges) return;
      ranges.forEach(r => {
        sheet.getCell(r).value = value || "";
      });
    };

    /* ===== HEADER ===== */
    writeByName("teacher_name", teacherName);
    writeByName("grade_level", gradeLevel);
    writeByName("learning_area", subject);
    writeByName("quarter", quarter);
    writeByName("week_date", weekDate);

    /* ===== OBJECTIVES ===== */
    writeByName("obj_content", generatedLesson.I_Objectives?.[0]);
    writeByName("obj_performance", generatedLesson.I_Objectives?.[1]);
    writeByName("obj_learning_competencies", generatedLesson.I_Objectives?.[2]);

    /* ===== PROCEDURES (A–J COMPLETE) ===== */
    const p = generatedLesson.IV_Procedures || [];

    writeByName("proc_A_review", p[0]);
    writeByName("proc_B_motivation", p[1]);
    writeByName("proc_C_presentation", p[2]);
    writeByName("proc_D_discussion", p[3]);
    writeByName("proc_E_practice", p[4]);
    writeByName("proc_F_mastery", p[5]);
    writeByName("proc_G_application", p[6]);
    writeByName("proc_H_generalization", p[7]);
    writeByName("proc_I_evaluation", p[8]);
    writeByName("proc_J_remediation", p[9]);

    const buffer = await wb.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="DLL_FINAL.xlsx"'
    );

    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log("✅ DLL FINAL SERVICE RUNNING");
});