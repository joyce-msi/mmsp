import ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";

interface Task {
  key: string;
  summary: string;
  start: string;
  end: string;
  status: "Done" | "In Progress" | "To Do";
  sprint: number;
}

interface Project {
  key: string;
  name: string;
  color: string;
}

const statusColors: Record<string, string> = {
  "Done": "FF22c55e",
  "In Progress": "FFf59e0b",
  "To Do": "FF6b7280",
};

const statusFills: Record<string, string> = {
  "Done": "FFdcfce7",
  "In Progress": "FFfef9c3",
  "To Do": "FFf3f4f6",
};

const sprintColors: Record<number, string> = {
  1: "FFe0e7ff",
  2: "FFe0f2fe",
  3: "FFd1fae5",
};

async function main() {
  const dir = import.meta.dirname;
  const projects: Project[] = JSON.parse(fs.readFileSync(path.join(dir, "projects.json"), "utf-8"));

  const wb = new ExcelJS.Workbook();
  wb.creator = "MMSP Dev Chart";
  wb.created = new Date();

  for (const project of projects) {
    const dataPath = path.join(dir, "data", `${project.key}.json`);
    if (!fs.existsSync(dataPath)) continue;

    const tasks: Task[] = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
    const ws = wb.addWorksheet(project.name);

    // Column widths
    ws.columns = [
      { header: "Key", key: "key", width: 16 },
      { header: "Summary", key: "summary", width: 45 },
      { header: "Sprint", key: "sprint", width: 10 },
      { header: "Start Date", key: "start", width: 14 },
      { header: "Due Date", key: "end", width: 14 },
      { header: "Status", key: "status", width: 14 },
      { header: "Duration (days)", key: "duration", width: 16 },
    ];

    // Header style
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1e293b" } };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };
    headerRow.height = 24;

    // Group by sprint
    const sprints = [...new Set(tasks.map(t => t.sprint))].sort((a, b) => a - b);
    let rowIndex = 2;

    for (const sprint of sprints) {
      // Sprint header row
      const sprintTasks = tasks.filter(t => t.sprint === sprint);
      const sprintRow = ws.getRow(rowIndex);
      sprintRow.getCell(1).value = `Sprint ${sprint}`;
      sprintRow.getCell(1).font = { bold: true, size: 11 };
      sprintRow.getCell(2).value = `${sprintTasks.length} tasks`;
      sprintRow.getCell(2).font = { italic: true, color: { argb: "FF64748b" } };
      sprintRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: sprintColors[sprint] || "FFf1f5f9" } };
      sprintRow.height = 22;
      rowIndex++;

      for (const task of sprintTasks) {
        const startDate = new Date(task.start);
        const endDate = new Date(task.end);
        const duration = Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));

        const row = ws.getRow(rowIndex);
        row.getCell(1).value = task.key;
        row.getCell(2).value = task.summary;
        row.getCell(3).value = task.sprint;
        row.getCell(3).alignment = { horizontal: "center" };
        row.getCell(4).value = task.start;
        row.getCell(5).value = task.end;
        row.getCell(6).value = task.status;
        row.getCell(7).value = duration;
        row.getCell(7).alignment = { horizontal: "center" };

        // Status cell coloring
        const statusCell = row.getCell(6);
        statusCell.font = { bold: true, color: { argb: statusColors[task.status] || "FF000000" } };
        statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: statusFills[task.status] || "FFFFFFFF" } };
        statusCell.alignment = { horizontal: "center" };

        rowIndex++;
      }
    }

    // Add borders to all data cells
    for (let r = 1; r < rowIndex; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 7; c++) {
        row.getCell(c).border = {
          top: { style: "thin", color: { argb: "FFe2e8f0" } },
          bottom: { style: "thin", color: { argb: "FFe2e8f0" } },
          left: { style: "thin", color: { argb: "FFe2e8f0" } },
          right: { style: "thin", color: { argb: "FFe2e8f0" } },
        };
      }
    }

    // Auto-filter
    ws.autoFilter = { from: "A1", to: "G1" };
  }

  // Summary sheet
  const summary = wb.addWorksheet("Summary", { properties: { tabColor: { argb: "FF1e293b" } } });
  summary.columns = [
    { header: "Project", key: "project", width: 18 },
    { header: "Total Tasks", key: "total", width: 14 },
    { header: "Done", key: "done", width: 10 },
    { header: "In Progress", key: "inProgress", width: 14 },
    { header: "To Do", key: "toDo", width: 10 },
    { header: "Progress %", key: "progress", width: 14 },
  ];

  const sHeaderRow = summary.getRow(1);
  sHeaderRow.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
  sHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1e293b" } };
  sHeaderRow.alignment = { vertical: "middle", horizontal: "center" };
  sHeaderRow.height = 24;

  let sRow = 2;
  for (const project of projects) {
    const dataPath = path.join(dir, "data", `${project.key}.json`);
    if (!fs.existsSync(dataPath)) continue;
    const tasks: Task[] = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
    const done = tasks.filter(t => t.status === "Done").length;
    const inProg = tasks.filter(t => t.status === "In Progress").length;
    const toDo = tasks.filter(t => t.status === "To Do").length;
    const pct = tasks.length > 0 ? Math.round((done / tasks.length) * 100) : 0;

    const row = summary.getRow(sRow);
    row.getCell(1).value = project.name;
    row.getCell(1).font = { bold: true };
    row.getCell(2).value = tasks.length;
    row.getCell(3).value = done;
    row.getCell(3).font = { color: { argb: "FF22c55e" } };
    row.getCell(4).value = inProg;
    row.getCell(4).font = { color: { argb: "FFf59e0b" } };
    row.getCell(5).value = toDo;
    row.getCell(6).value = `${pct}%`;
    row.getCell(6).alignment = { horizontal: "center" };

    for (let c = 1; c <= 6; c++) {
      row.getCell(c).alignment = { ...row.getCell(c).alignment, horizontal: "center" };
      row.getCell(c).border = {
        top: { style: "thin", color: { argb: "FFe2e8f0" } },
        bottom: { style: "thin", color: { argb: "FFe2e8f0" } },
        left: { style: "thin", color: { argb: "FFe2e8f0" } },
        right: { style: "thin", color: { argb: "FFe2e8f0" } },
      };
    }
    sRow++;
  }

  // Move Summary to first position
  wb.removeWorksheet(summary.id);
  const newSummary = wb.addWorksheet("Summary", { properties: { tabColor: { argb: "FF1e293b" } } });
  // Re-create summary in a new approach - just reorder by writing fresh
  // ExcelJS doesn't support reordering, so we'll keep it as last tab

  const outPath = path.join(dir, "MMSP-DevChart.xlsx");
  await wb.xlsx.writeFile(outPath);
  console.log(`Exported to ${outPath}`);
  console.log(`  ${projects.length} project sheets + Summary`);
}

main();
