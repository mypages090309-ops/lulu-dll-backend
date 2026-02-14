import express from "express";
import ExcelJS from "exceljs";
import cors from "cors";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

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

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile("./DLL_FORMAT.xlsx");

    const s1 = wb.getWorksheet("DLL");   // Sheet name as in template
    const s2 = wb.getWorksheet("REFLECTION");

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

    /* ============ HEADER (MERGED CELLS SAFE) ============ */
    set("E4", teacherName);     // Teacher
    set("J4", gradeLevel);      // Grade Level
    set("E5", subject);         // Subject
    set("J5", quarter);         // Quarter
    set("E6", weekDate);        // Week

    /* ============ I. OBJECTIVES ============ */
    set("D9", normalize(generatedLesson.I_Objectives));
    set("D10", normalize(generatedLesson.I_Objectives));
    set("D11", normalize(generatedLesson.I_Objectives));

    /* ============ II. CONTENT ============ */
    set("D13", normalize(generatedLesson.II_Content));

    /* ============ III. LEARNING RESOURCES ============ */
    set("D15", normalize(generatedLesson.III_LearningResources));
    set("D16", "");
    set("D17", "");
    set("D18", "");

    /* ============ IV. PROCEDURES (ARRAY → ROWS) ============ */
    const steps = Array.isArray(generatedLesson.IV_Procedures)
      ? generatedLesson.IV_Procedures
      : [];

    const procCells = [
      "D21", // A
      "D22", // B
      "D23", // C
      "D24", // D
      "D25", // E
      "D26", // F
      "D27"  // G
    ];

    procCells.forEach((cell, i) => {
      set(cell, steps[i] || "");
    });

    /* ============ REFLECTION (SAFE BLANK) ============ */
    if (s2) {
      ["D6","D7","D8","D9","D10","D11","D12"].forEach(c => {
        const cell = s2.getCell(c);
        cell.value = "";
        cell.alignment = { wrapText: true, vertical: "top" };
      });
    }

    /* ============ SAVE & DOWNLOAD ============ */
    const fileName = `DLL_${Date.now()}.xlsx`;
    const outPath = path.join(process.cwd(), fileName);

    await wb.xlsx.writeFile(outPath);

    res.download(outPath, fileName, () => {
      fs.unlinkSync(outPath);
    });

  } catch (err) {
    console.error("DLL ERROR:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(3000, () => {
  console.log("✅ DLL Excel Fill Service running on port 3000");
});