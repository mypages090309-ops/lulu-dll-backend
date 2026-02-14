const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

/**
 * Utility: set value to named range
 */
function setNamedRange(workbook, rangeName, value) {
  const name = workbook.Workbook.Names.find(n => n.Name === rangeName);
  if (!name) return;

  const sheetName = name.Ref.split("!")[0].replace(/'/g, "");
  const cellRef = name.Ref.split("!")[1];
  const ws = workbook.Sheets[sheetName];

  XLSX.utils.sheet_add_aoa(ws, [[value]], { origin: cellRef });
}

/**
 * MAIN DLL EXPORT
 */
app.post("/fill-dll", async (req, res) => {
  try {
    const data = req.body;

    const templatePath = path.join(__dirname, "DLL_FORMAT.xlsx");
    const wb = XLSX.readFile(templatePath);

    /* ================= HEADER ================= */
    setNamedRange(wb, "school_name", data.school || "");
    setNamedRange(wb, "teacher_name", data.teacher || "");
    setNamedRange(wb, "grade_level", data.gradeLevel || "");
    setNamedRange(wb, "learning_area", data.subject || "");
    setNamedRange(wb, "quarter", data.quarter || "");
    setNamedRange(wb, "week_date", data.weekDate || "");

    /* ================= OBJECTIVES ================= */
    setNamedRange(wb, "content_standards", data.objectives?.contentStandards || "");
    setNamedRange(wb, "performance_standards", data.objectives?.performanceStandards || "");

    const days = ["mon", "tue", "wed", "thu", "fri"];

    days.forEach(day => {
      setNamedRange(
        wb,
        `learning_competencies_${day}`,
        (data.objectives?.learningCompetencies?.[day] || []).join("\n")
      );
    });

    /* ================= PROCEDURES (A–J) ================= */
    days.forEach(day => {
      const procs = data.procedures?.[day] || {};

      setNamedRange(wb, `proc_a_${day}`, procs.A || "");
      setNamedRange(wb, `proc_b_${day}`, procs.B || "");
      setNamedRange(wb, `proc_c_${day}`, procs.C || "");
      setNamedRange(wb, `proc_d_${day}`, procs.D || "");
      setNamedRange(wb, `proc_e_${day}`, procs.E || "");
      setNamedRange(wb, `proc_f_${day}`, procs.F || "");
      setNamedRange(wb, `proc_g_${day}`, procs.G || "");
      setNamedRange(wb, `proc_h_${day}`, procs.H || "");
      setNamedRange(wb, `proc_i_${day}`, procs.I || "");
      setNamedRange(wb, `proc_j_${day}`, procs.J || "");
    });

    /* ================= SAVE FILE ================= */
    const outPath = path.join(__dirname, "DLL_FILLED.xlsx");
    XLSX.writeFile(wb, outPath);

    res.download(outPath, "DLL.xlsx");
  } catch (err) {
    console.error("DLL EXPORT ERROR:", err);
    res.status(500).json({ error: "DLL export failed" });
  }
});

app.listen(PORT, () => {
  console.log(`✅ DLL Excel Fill Service running on port ${PORT}`);
});