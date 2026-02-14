import express from "express";
import cors from "cors";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const app = express();
app.use(cors());
app.use(express.json());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ===============================
// HELPER: safely set named range
// ===============================
function setNamedRange(workbook, rangeName, value) {
  const definedName = workbook.definedNames.getName(rangeName);
  if (!definedName) return;

  const ranges = definedName.ranges;
  ranges.forEach((r) => {
    const ws = workbook.getWorksheet(r.split("!")[0].replace(/'/g, ""));
    const cellRef = r.split("!")[1];
    ws.getCell(cellRef).value = value ?? "";
  });
}

// ===============================
// MAIN ENDPOINT
// ===============================
app.post("/fill-dll", async (req, res) => {
  try {
    const data = req.body;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(
      path.join(__dirname, "DLL_FORMAT.xlsx")
    );

    // ===============================
    // HEADER FIELDS
    // ===============================
    setNamedRange(workbook, "school_name", data.school);
    setNamedRange(workbook, "teacher_name", data.teacherName);
    setNamedRange(workbook, "grade_level", data.gradeLevel);
    setNamedRange(workbook, "learning_area", data.subject);
    setNamedRange(workbook, "quarter", data.quarter);
    setNamedRange(workbook, "week_date", data.weekDate);

    // ===============================
    // DAYS (MONDAY – FRIDAY)
    // ===============================
    const days = ["monday", "tuesday", "wednesday", "thursday", "friday"];

    days.forEach((day) => {
      const d = data.dll?.[day] || {};

      setNamedRange(workbook, `${day}_objectives`, d.objectives);
      setNamedRange(workbook, `${day}_content`, d.content);

      // PROCEDURES A–J (10 steps)
      setNamedRange(workbook, `${day}_proc_a`, d.procedures?.A);
      setNamedRange(workbook, `${day}_proc_b`, d.procedures?.B);
      setNamedRange(workbook, `${day}_proc_c`, d.procedures?.C);
      setNamedRange(workbook, `${day}_proc_d`, d.procedures?.D);
      setNamedRange(workbook, `${day}_proc_e`, d.procedures?.E);
      setNamedRange(workbook, `${day}_proc_f`, d.procedures?.F);
      setNamedRange(workbook, `${day}_proc_g`, d.procedures?.G);
      setNamedRange(workbook, `${day}_proc_h`, d.procedures?.H);
      setNamedRange(workbook, `${day}_proc_i`, d.procedures?.I);
      setNamedRange(workbook, `${day}_proc_j`, d.procedures?.J);

      setNamedRange(workbook, `${day}_assessment`, d.assessment);
      setNamedRange(workbook, `${day}_assignment`, d.assignment);
      setNamedRange(workbook, `${day}_remarks`, d.remarks);
    });

    // ===============================
    // SEND FILE
    // ===============================
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=DLL_DEPED.xlsx"
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error("DLL EXPORT ERROR:", err);
    res.status(500).json({ error: "DLL export failed" });
  }
});

// ===============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`DLL Excel Fill Service running on port ${PORT}`);
});