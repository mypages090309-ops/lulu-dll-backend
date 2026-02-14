import express from "express";
import ExcelJS from "exceljs";
import cors from "cors";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

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

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("./DLL_FORMAT.xlsx");

    const s1 = workbook.getWorksheet(1);
    const s2 = workbook.getWorksheet(2);

    const normalize = (v) => {
      if (!v) return "";
      if (Array.isArray(v)) return v.join("\n");
      return String(v).trim();
    };

    const set = (sheet, cell, val) => {
      const c = sheet.getCell(cell);
      c.value = val;
      c.alignment = { wrapText: true, vertical: "top" };
    };

    /* ===== HEADER ===== */
    set(s1, "F3", teacherName);
    set(s1, "I2", gradeLevel);
    set(s1, "I3", subject);
    set(s1, "I4", quarter);
    set(s1, "F4", weekDate);

    /* ===== OBJECTIVES ===== */
    set(s1, "C7", normalize(generatedLesson.I_Objectives));
    set(s1, "C8", "");
    set(s1, "C9", normalize(generatedLesson.I_Objectives));

    /* ===== CONTENT ===== */
    set(s1, "C11", normalize(generatedLesson.II_Content));

    /* ===== LEARNING RESOURCES ===== */
    set(s1, "C14", normalize(generatedLesson.III_LearningResources));
    set(s1, "C15", "");
    set(s1, "C16", "");
    set(s1, "C17", "");

    /* ===== PROCEDURES (ARRAY → A–G ROWS) ===== */
    const steps = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const procRows = ["C20", "C21", "C22", "C23", "C24", "C25", "C26"];

    procRows.forEach((cell, i) => {
      set(s1, cell, steps[i] || "");
    });

    /* ===== REFLECTION (LEFT BLANK) ===== */
    ["C5","C6","C7","C8","C9","C10","C11","C12"].forEach(c =>
      set(s2, c, "")
    );

    const fileName = `DLL_${Date.now()}.xlsx`;
    const out = path.join(process.cwd(), fileName);

    await workbook.xlsx.writeFile(out);
    res.download(out, fileName, () => fs.unlinkSync(out));

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => {
  console.log("✅ DLL Excel Fill Service running on port 3000");
});