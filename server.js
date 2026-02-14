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

    if (!generatedLesson) {
      return res.status(400).json({ error: "Missing generatedLesson" });
    }

    // ✅ ABSOLUTE PATH (RENDER SAFE)
    const templatePath = path.join(process.cwd(), "DLL_FORMAT.xlsx");
    if (!fs.existsSync(templatePath)) {
      throw new Error("DLL_FORMAT.xlsx not found in backend root");
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    // ✅ USE FIRST SHEET ONLY (SAFE)
    const s1 = wb.getWorksheet(1);

    const normalize = (v) => {
      if (!v) return "";
      if (Array.isArray(v)) return v.join("\n");
      return String(v).trim();
    };

    const set = (cell, value) => {
      const c = s1.getCell(cell);
      c.value = value;
      c.alignment = { wrapText: true, vertical: "top" };
    };

    /* ===== HEADER (BASED ON YOUR TEMPLATE) ===== */
    set("E4", teacherName);
    set("J4", gradeLevel);
    set("E5", subject);
    set("J5", quarter);
    set("E6", weekDate);

    /* ===== I. OBJECTIVES ===== */
    set("D9", normalize(generatedLesson.I_Objectives));
    set("D10", normalize(generatedLesson.I_Objectives));
    set("D11", normalize(generatedLesson.I_Objectives));

    /* ===== II. CONTENT ===== */
    set("D13", normalize(generatedLesson.II_Content));

    /* ===== III. LEARNING RESOURCES ===== */
    set("D15", normalize(generatedLesson.III_LearningResources));
    set("D16", "");
    set("D17", "");
    set("D18", "");

    /* ===== IV. PROCEDURES (ARRAY → ROWS) ===== */
    const steps = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const procCells = ["D21","D22","D23","D24","D25","D26","D27"];
    procCells.forEach((cell, i) => {
      set(cell, steps[i] || "");
    });

    /* ===== SAVE & DOWNLOAD ===== */
    const fileName = `DLL_${Date.now()}.xlsx`;
    const outPath = path.join(process.cwd(), fileName);

    await wb.xlsx.writeFile(outPath);
    res.download(outPath, fileName, () => fs.unlinkSync(outPath));

  } catch (err) {
    console.error("DLL ERROR:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => {
  console.log("✅ DLL Excel Fill Service running on port 3000");
});