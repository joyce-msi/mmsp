import { useState, useRef, useCallback, useEffect } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import projectsConfig from "./projects.json";

type Status = "Done" | "In Progress" | "To Do";
interface Task { key: string; summary: string; start: string; end: string; status: Status; sprint: number; }
interface ProjectDef { key: string; name: string; color: string; }

const CHART_START = new Date("2026-03-11");
const CHART_END   = new Date("2026-05-29");
const TOTAL_DAYS  = Math.round((CHART_END.getTime() - CHART_START.getTime()) / 86400000) + 1;

const sprintConfig = {
  1: { label: "Sprint 1", color: "#6366f1", bg: "#1e1b4b", range: "Mar 11 – Mar 31" },
  2: { label: "Sprint 2", color: "#0ea5e9", bg: "#0c2a3f", range: "Apr 1 – Apr 30" },
  3: { label: "Sprint 3", color: "#10b981", bg: "#052e1e", range: "May 1 – May 29" },
  4: { label: "Sprint 4", color: "#f59e0b", bg: "#2a1f0c", range: "Jun+" },
} as Record<number, { label: string; color: string; bg: string; range: string }>;

const statusColors: Record<Status, { bar: string; text: string }> = {
  "Done":        { bar: "#22c55e", text: "#166534" },
  "In Progress": { bar: "#f59e0b", text: "#92400e" },
  "To Do":       { bar: "#64748b", text: "#1e293b"  },
};

const STATUS_CYCLE: Status[] = ["To Do", "In Progress", "Done"];

const toDay = (d: string) => Math.round((new Date(d).getTime() - CHART_START.getTime()) / 86400000);
const fromDay = (day: number) => {
  const d = new Date(CHART_START);
  d.setUTCDate(d.getUTCDate() + day);
  return d.toISOString().slice(0, 10);
};
const fmt = (d: string) => new Date(d).toLocaleDateString("en-GB", { day:"2-digit", month:"short" });

// Build week headers
const weeks: { label: string; start: number; span: number }[] = [];
let _d = new Date(CHART_START);
while (_d <= CHART_END) {
  const dayIdx = Math.round((_d.getTime() - CHART_START.getTime()) / 86400000);
  const mon = new Date(_d);
  const span = Math.min(7, TOTAL_DAYS - dayIdx);
  weeks.push({ label: `${mon.getDate()} ${mon.toLocaleDateString("en-GB",{month:"short"})}`, start: dayIdx, span });
  _d.setDate(_d.getDate() + 7);
}

const projects: ProjectDef[] = projectsConfig;

type DragMode = "move" | "resize-end";

// Cache for loaded project data
const dataCache: Record<string, Task[]> = {};

async function loadProjectData(projectKey: string): Promise<Task[]> {
  if (dataCache[projectKey]) return dataCache[projectKey];
  try {
    const res = await fetch(`${import.meta.env.BASE_URL}data/${projectKey}.json`);
    if (!res.ok) return [];
    const data = await res.json();
    dataCache[projectKey] = data;
    return data;
  } catch {
    return [];
  }
}

function ProjectChart({ project }: { project: ProjectDef }) {
  const [tasks, setTasks] = useState<Task[]>([]);
  const [originalTasks, setOriginalTasks] = useState<Task[]>([]);
  const [loading, setLoading] = useState(true);
  const [modified, setModified] = useState<Set<string>>(new Set());
  const [activeSprint, setActiveSprint] = useState<number | "All">("All");
  const [activeStatus, setActiveStatus] = useState<Status | "All">("All");
  const [hovered, setHovered] = useState<string | null>(null);
  const [editingKey, setEditingKey] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [saving, setSaving] = useState(false);
  const [saveMsg, setSaveMsg] = useState<{ text: string; ok: boolean } | null>(null);
  const [syncing, setSyncing] = useState(false);

  const dragRef = useRef<{
    key: string;
    mode: DragMode;
    startX: number;
    origStart: number;
    origEnd: number;
  } | null>(null);

  const LABEL_W = 230;
  const COL_W = 14;
  const BAR_H = 22;
  const ROW_H = 36;

  // Load data
  useEffect(() => {
    setLoading(true);
    setModified(new Set());
    setSaveMsg(null);
    loadProjectData(project.key).then(data => {
      setTasks(data);
      setOriginalTasks(data);
      setLoading(false);
    });
  }, [project.key]);

  const updateTask = useCallback((key: string, updates: Partial<Task>) => {
    setTasks(prev => prev.map(t => t.key === key ? { ...t, ...updates } : t));
    setModified(prev => new Set(prev).add(key));
  }, []);

  // Drag handlers
  const handleBarMouseDown = useCallback((e: React.MouseEvent, key: string, mode: DragMode) => {
    e.preventDefault();
    e.stopPropagation();
    const task = tasks.find(t => t.key === key);
    if (!task) return;
    dragRef.current = { key, mode, startX: e.clientX, origStart: toDay(task.start), origEnd: toDay(task.end) };
  }, [tasks]);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      const drag = dragRef.current;
      if (!drag) return;
      const dx = e.clientX - drag.startX;
      const dayDelta = Math.round(dx / COL_W);
      if (drag.mode === "move") {
        const newStart = Math.max(0, Math.min(TOTAL_DAYS - 1, drag.origStart + dayDelta));
        const dur = drag.origEnd - drag.origStart;
        const newEnd = Math.min(TOTAL_DAYS - 1, newStart + dur);
        updateTask(drag.key, { start: fromDay(newStart), end: fromDay(newEnd) });
      } else {
        const newEnd = Math.max(drag.origStart + 1, Math.min(TOTAL_DAYS - 1, drag.origEnd + dayDelta));
        updateTask(drag.key, { end: fromDay(newEnd) });
      }
    };
    const handleMouseUp = () => { dragRef.current = null; };
    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
    return () => { window.removeEventListener("mousemove", handleMouseMove); window.removeEventListener("mouseup", handleMouseUp); };
  }, [updateTask]);

  const cycleStatus = useCallback((key: string) => {
    const task = tasks.find(t => t.key === key);
    if (!task) return;
    const idx = STATUS_CYCLE.indexOf(task.status);
    updateTask(key, { status: STATUS_CYCLE[(idx + 1) % STATUS_CYCLE.length] });
  }, [tasks, updateTask]);

  const startEdit = useCallback((key: string, current: string) => {
    setEditingKey(key);
    setEditValue(current);
  }, []);

  const commitEdit = useCallback(() => {
    if (editingKey && editValue.trim()) {
      updateTask(editingKey, { summary: editValue.trim() });
    }
    setEditingKey(null);
  }, [editingKey, editValue, updateTask]);

  const saveToJira = useCallback(async () => {
    if (modified.size === 0) return;
    setSaving(true);
    setSaveMsg(null);
    const origMap = new Map(originalTasks.map(t => [t.key, t]));
    const changes = [...modified].map(key => {
      const task = tasks.find(t => t.key === key)!;
      const orig = origMap.get(key);
      const diff: Record<string, string> = { key };
      if (!orig || task.summary !== orig.summary) diff.summary = task.summary;
      if (!orig || task.start !== orig.start) diff.start = task.start;
      if (!orig || task.end !== orig.end) diff.end = task.end;
      if (!orig || task.status !== orig.status) diff.status = task.status;
      return diff;
    });
    try {
      const res = await fetch("/api/update-tasks", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(changes),
      });
      const data = await res.json();
      const failed = data.results.filter((r: { success: boolean }) => !r.success);
      if (failed.length === 0) {
        setSaveMsg({ text: `Saved ${changes.length} task(s) to Jira`, ok: true });
        setOriginalTasks(tasks);
        setModified(new Set());
        // Update cache
        dataCache[project.key] = tasks;
      } else {
        setSaveMsg({ text: `${failed.length} of ${changes.length} failed`, ok: false });
      }
    } catch (e: unknown) {
      setSaveMsg({ text: `Error: ${e instanceof Error ? e.message : String(e)}`, ok: false });
    } finally {
      setSaving(false);
    }
  }, [modified, tasks, originalTasks, project.key]);

  const discard = useCallback(() => {
    setTasks(originalTasks);
    setModified(new Set());
    setSaveMsg(null);
  }, [originalTasks]);

  const syncFromJira = useCallback(async () => {
    setSyncing(true);
    setSaveMsg(null);
    try {
      const res = await fetch(`/api/sync/${project.key}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data: Task[] = await res.json();
      setTasks(data);
      setOriginalTasks(data);
      setModified(new Set());
      delete dataCache[project.key];
      setSaveMsg({ text: `Synced ${data.length} tasks from Jira`, ok: true });
    } catch (e: unknown) {
      setSaveMsg({ text: `Sync failed: ${e instanceof Error ? e.message : String(e)}`, ok: false });
    } finally {
      setSyncing(false);
    }
  }, [project.key]);

  if (loading) {
    return <div style={{ padding: 40, textAlign: "center", color: "#64748b" }}>Loading {project.name} data...</div>;
  }

  if (tasks.length === 0 && !syncing) {
    return (
      <div style={{ padding: 40, textAlign: "center", color: "#64748b" }}>
        <p>No tasks found for {project.name}.</p>
        <button onClick={syncFromJira}
          style={{ background: "#3b82f6", border: "none", color: "#fff", borderRadius: 8, padding: "8px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer", marginTop: 8 }}>
          Sync from Jira
        </button>
      </div>
    );
  }

  const filtered = tasks.filter(t =>
    (activeSprint === "All" || t.sprint === activeSprint) &&
    (activeStatus === "All" || t.status === activeStatus)
  );

  const sprintNumbers = [...new Set(tasks.map(t => t.sprint))].sort((a, b) => a - b);
  const grouped = sprintNumbers.map(s => ({ sprint: s, tasks: filtered.filter(t => t.sprint === s) })).filter(g => g.tasks.length > 0);
  const todayDay = 0;

  return (
    <>
      {/* Filters + Save */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap", alignItems: "center" }}>
        {([["All", "#94a3b8"] as const, ...sprintNumbers.map(n => [n, sprintConfig[n]?.color || "#94a3b8"] as const)]).map(([s, c]) => (
          <button key={String(s)} onClick={() => setActiveSprint(s as number | "All")}
            style={{ background: activeSprint === s ? c : "#1e293b", border: `1px solid ${activeSprint === s ? c : "#334155"}`, color: activeSprint === s ? "#fff" : "#94a3b8", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", transition: "all 0.15s" }}>
            {s === "All" ? "All Sprints" : `Sprint ${s}`}
          </button>
        ))}
        <div style={{ width: 1, background: "#334155", margin: "0 4px" }} />
        {(["All", "Done", "In Progress", "To Do"] as const).map(s => (
          <button key={s} onClick={() => setActiveStatus(s as Status | "All")}
            style={{ background: activeStatus === s ? (s === "All" ? "#475569" : statusColors[s as Status]?.bar ?? "#475569") : s === "All" ? "#1e293b" : `${statusColors[s as Status]?.bar}22`, border: `1px solid ${activeStatus === s ? (s === "All" ? "#475569" : statusColors[s as Status]?.bar ?? "#475569") : s === "All" ? "#334155" : `${statusColors[s as Status]?.bar}66`}`, color: activeStatus === s ? "#fff" : s === "All" ? "#94a3b8" : statusColors[s as Status]?.bar ?? "#94a3b8", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", transition: "all 0.15s" }}>
            {s === "All" ? "All Status" : s}
          </button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 12, color: "#475569" }}>{filtered.length} tasks shown</span>
          <button onClick={syncFromJira} disabled={syncing || modified.size > 0}
            title={modified.size > 0 ? "Save or discard changes before syncing" : "Fetch latest data from Jira"}
            style={{ background: "#0f172a", border: "1px solid #334155", color: syncing ? "#64748b" : "#94a3b8", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 600, cursor: syncing || modified.size > 0 ? "not-allowed" : "pointer", opacity: modified.size > 0 ? 0.4 : 1, transition: "all 0.15s", display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ display: "inline-block", animation: syncing ? "spin 1s linear infinite" : "none" }}>&#8635;</span>
            {syncing ? "Syncing..." : "Sync from Jira"}
          </button>
          {modified.size > 0 && (
            <>
              <button onClick={saveToJira} disabled={saving}
                style={{ background: "#3b82f6", border: "none", color: "#fff", borderRadius: 8, padding: "6px 16px", fontSize: 12, fontWeight: 700, cursor: saving ? "wait" : "pointer", opacity: saving ? 0.6 : 1 }}>
                {saving ? "Saving..." : `Save ${modified.size} change(s) to Jira`}
              </button>
              <button onClick={discard}
                style={{ background: "transparent", border: "1px solid #475569", color: "#94a3b8", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                Discard
              </button>
            </>
          )}
          {saveMsg && <span style={{ fontSize: 12, color: saveMsg.ok ? "#22c55e" : "#ef4444", fontWeight: 600 }}>{saveMsg.text}</span>}
        </div>
      </div>

      {modified.size === 0 && (
        <div style={{ fontSize: 11, color: "#475569", marginBottom: 12 }}>
          Drag bars to move · Drag right edge to resize · Click status dot to cycle · Double-click summary to edit
        </div>
      )}

      {/* Chart */}
      <div style={{ background: "#111827", borderRadius: 12, border: "1px solid #1f2937", overflow: "auto" }}>
        <div style={{ minWidth: LABEL_W + COL_W * TOTAL_DAYS }}>
          {/* Header */}
          <div style={{ display: "flex", borderBottom: "1px solid #1f2937", background: "#0d1117", position: "sticky", top: 0, zIndex: 10 }}>
            <div style={{ width: LABEL_W, flexShrink: 0, padding: "8px 16px", fontSize: 11, fontWeight: 700, color: "#475569", borderRight: "1px solid #1f2937" }}>TASK</div>
            <div style={{ position: "relative", flex: 1 }}>
              <div style={{ display: "flex" }}>
                {weeks.map((w, i) => (
                  <div key={i} style={{ width: w.span * COL_W, boxSizing: "border-box", borderLeft: "1px solid #1f2937", padding: "4px 6px", fontSize: 10, fontWeight: 700, color: "#64748b", whiteSpace: "nowrap", overflow: "hidden" }}>
                    {w.label}
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", borderTop: "1px solid #1f2937" }}>
                {[
                  { label: "March 2026", days: 21 },
                  { label: "April 2026", days: 30 },
                  { label: "May 2026", days: 29 },
                ].map((m, i) => (
                  <div key={i} style={{ width: m.days * COL_W, boxSizing: "border-box", borderLeft: i > 0 ? "1px solid #334155" : "none", padding: "3px 6px", fontSize: 10, fontWeight: 700, color: "#334155", letterSpacing: 1 }}>
                    {m.label.toUpperCase()}
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Sprint groups */}
          {grouped.map(({ sprint, tasks: sprintTasks }) => {
            const sc = sprintConfig[sprint] || { label: `Sprint ${sprint}`, color: "#94a3b8", bg: "#1a1a2e", range: "" };
            return (
              <div key={sprint}>
                <div style={{ display: "flex", background: sc.bg, borderBottom: "1px solid #1f2937" }}>
                  <div style={{ width: LABEL_W, flexShrink: 0, padding: "7px 16px", borderRight: "1px solid #1f2937", display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: sc.color }} />
                    <span style={{ fontSize: 12, fontWeight: 700, color: sc.color }}>{sc.label}</span>
                    <span style={{ fontSize: 11, color: "#475569" }}>{sc.range}</span>
                  </div>
                  <div style={{ flex: 1 }} />
                </div>

                {sprintTasks.map((task, idx) => {
                  const sc2 = statusColors[task.status];
                  const startD = toDay(task.start);
                  const endD = toDay(task.end);
                  const dur = endD - startD;
                  const isHov = hovered === task.key;
                  const isMod = modified.has(task.key);
                  const barWidth = Math.max(dur * COL_W - 2, 4);

                  return (
                    <div key={task.key}
                      onMouseEnter={() => setHovered(task.key)}
                      onMouseLeave={() => setHovered(null)}
                      style={{ display: "flex", borderBottom: "1px solid #1a2234", background: isMod ? "#1a1a2e" : isHov ? "#1a2535" : idx % 2 === 0 ? "#111827" : "#0f1623", height: ROW_H, transition: "background 0.1s" }}>

                      <div style={{ width: LABEL_W, flexShrink: 0, borderRight: "1px solid #1f2937", padding: "0 12px", display: "flex", flexDirection: "column", justifyContent: "center", gap: 1 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ fontSize: 10, color: sc2.bar, fontWeight: 700, whiteSpace: "nowrap" }}>{task.key}</span>
                          <span onClick={() => cycleStatus(task.key)} title={`Status: ${task.status} (click to change)`}
                            style={{ width: 8, height: 8, borderRadius: "50%", background: sc2.bar, flexShrink: 0, cursor: "pointer", border: "1px solid rgba(255,255,255,0.3)", transition: "transform 0.1s" }}
                            onMouseEnter={e => (e.currentTarget.style.transform = "scale(1.4)")}
                            onMouseLeave={e => (e.currentTarget.style.transform = "scale(1)")} />
                          {isMod && <span style={{ fontSize: 8, color: "#3b82f6", fontWeight: 700 }}>●</span>}
                        </div>
                        {editingKey === task.key ? (
                          <input autoFocus value={editValue} onChange={e => setEditValue(e.target.value)}
                            onBlur={commitEdit} onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditingKey(null); }}
                            style={{ fontSize: 11, color: "#e2e8f0", background: "#1e293b", border: "1px solid #3b82f6", borderRadius: 3, padding: "1px 4px", outline: "none", width: "100%" }} />
                        ) : (
                          <span onDoubleClick={() => startEdit(task.key, task.summary)}
                            style={{ fontSize: 11, color: "#94a3b8", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", cursor: "text" }}
                            title={`${task.summary} (double-click to edit)`}>{task.summary}</span>
                        )}
                      </div>

                      <div style={{ position: "relative", flex: 1, overflow: "hidden" }}>
                        {todayDay >= 0 && todayDay < TOTAL_DAYS && (
                          <div style={{ position: "absolute", left: (todayDay + 0.5) * COL_W, top: 0, bottom: 0, width: 2, background: "#ef4444", opacity: 0.6, zIndex: 2, pointerEvents: "none" }} />
                        )}
                        {[toDay("2026-04-01"), toDay("2026-05-01")].map((dd, i) => (
                          <div key={i} style={{ position: "absolute", left: dd * COL_W, top: 0, bottom: 0, width: 1, background: "#334155", opacity: 0.5, zIndex: 1, pointerEvents: "none" }} />
                        ))}
                        <div onMouseDown={e => handleBarMouseDown(e, task.key, "move")}
                          style={{
                            position: "absolute", left: startD * COL_W + 1, top: (ROW_H - BAR_H) / 2,
                            width: barWidth, height: BAR_H,
                            background: `linear-gradient(90deg, ${sc2.bar}dd, ${sc2.bar}99)`,
                            borderRadius: 5, display: "flex", alignItems: "center", paddingLeft: 6,
                            boxShadow: isHov ? `0 0 0 1.5px ${sc2.bar}, 0 2px 8px ${sc2.bar}44` : isMod ? `0 0 0 1px #3b82f6` : "none",
                            transition: "box-shadow 0.15s", zIndex: 3, overflow: "hidden", cursor: "grab", userSelect: "none",
                          }}>
                          <span style={{ fontSize: 9, color: "#fff", fontWeight: 600, whiteSpace: "nowrap", opacity: 0.9, pointerEvents: "none" }}>
                            {fmt(task.start)} – {fmt(task.end)}
                          </span>
                          <div onMouseDown={e => handleBarMouseDown(e, task.key, "resize-end")}
                            style={{ position: "absolute", right: 0, top: 0, width: 6, height: "100%", cursor: "ew-resize", borderRadius: "0 5px 5px 0", background: isHov ? "rgba(255,255,255,0.2)" : "transparent" }} />
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>
    </>
  );
}

async function exportAllToExcel() {
  const chartStart = new Date("2026-03-11");
  const chartEnd = new Date("2026-05-29");
  const totalDays = Math.round((chartEnd.getTime() - chartStart.getTime()) / 86400000) + 1;

  // Generate all dates
  const allDates: Date[] = [];
  for (let i = 0; i < totalDays; i++) {
    const d = new Date(chartStart);
    d.setDate(d.getDate() + i);
    allDates.push(d);
  }

  const INFO_COLS = 5; // Key, Summary, Status, Start, End
  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const barColors: Record<string, string> = { "Done": "FF22c55e", "In Progress": "FFf59e0b", "To Do": "FF94a3b8" };
  const barBg: Record<string, string> = { "Done": "FFdcfce7", "In Progress": "FFfef9c3", "To Do": "FFe2e8f0" };
  const statusTextColors: Record<string, string> = { "Done": "FF166534", "In Progress": "FF92400e", "To Do": "FF475569" };
  const sprintBg: Record<number, string> = { 1: "FFe0e7ff", 2: "FFe0f2fe", 3: "FFd1fae5" };
  const sprintText: Record<number, string> = { 1: "FF4338ca", 2: "FF0284c7", 3: "FF059669" };

  const headerFill: ExcelJS.FillPattern = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1e293b" } };
  const headerFont: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
  const thinBorder: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFd1d5db" } };
  const cellBorder: Partial<ExcelJS.Borders> = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

  const wb = new ExcelJS.Workbook();
  wb.creator = "MMSP Dev Chart";

  for (const project of projects) {
    let tasks: Task[];
    if (dataCache[project.key]) {
      tasks = dataCache[project.key];
    } else {
      try {
        const res = await fetch(`${import.meta.env.BASE_URL}data/${project.key}.json`);
        if (!res.ok) continue;
        tasks = await res.json();
      } catch { continue; }
    }

    const ws = wb.addWorksheet(project.name);

    // Column widths: info cols + one per day
    ws.columns = [
      { width: 14 }, // Key
      { width: 38 }, // Summary
      { width: 12 }, // Status
      { width: 12 }, // Start
      { width: 12 }, // End
      ...allDates.map(() => ({ width: 3.5 })),
    ];

    // --- Row 1: Month headers ---
    const monthRow = ws.getRow(1);
    monthRow.height = 20;
    for (let c = 1; c <= INFO_COLS; c++) {
      const cell = monthRow.getCell(c);
      cell.fill = headerFill;
      cell.border = cellBorder;
    }
    // Merge month spans
    let monthStart = 0;
    let curMonth = allDates[0].getMonth();
    for (let i = 0; i <= allDates.length; i++) {
      const m = i < allDates.length ? allDates[i].getMonth() : -1;
      if (m !== curMonth) {
        const startCol = INFO_COLS + 1 + monthStart;
        const endCol = INFO_COLS + i;
        if (endCol > startCol) {
          ws.mergeCells(1, startCol, 1, endCol);
        }
        const mCell = monthRow.getCell(startCol);
        mCell.value = `${monthNames[curMonth]} 2026`;
        mCell.fill = headerFill;
        mCell.font = { ...headerFont, size: 11 };
        mCell.alignment = { horizontal: "center", vertical: "middle" };
        mCell.border = cellBorder;
        // Fill remaining merged cells border
        for (let c = startCol + 1; c <= endCol; c++) {
          monthRow.getCell(c).fill = headerFill;
          monthRow.getCell(c).border = cellBorder;
        }
        monthStart = i;
        curMonth = m;
      }
    }

    // --- Row 2: Day number headers ---
    const dayRow = ws.getRow(2);
    dayRow.height = 18;
    const infoHeaders = ["Key", "Summary", "Status", "Start", "End"];
    for (let c = 0; c < INFO_COLS; c++) {
      const cell = dayRow.getCell(c + 1);
      cell.value = infoHeaders[c];
      cell.fill = headerFill;
      cell.font = headerFont;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = cellBorder;
    }
    for (let i = 0; i < allDates.length; i++) {
      const cell = dayRow.getCell(INFO_COLS + 1 + i);
      cell.value = allDates[i].getDate();
      cell.font = { size: 8, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = headerFill;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = cellBorder;
    }

    // --- Row 3: Day name headers ---
    const dayNameRow = ws.getRow(3);
    dayNameRow.height = 16;
    for (let c = 1; c <= INFO_COLS; c++) {
      const cell = dayNameRow.getCell(c);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
      cell.border = cellBorder;
    }
    for (let i = 0; i < allDates.length; i++) {
      const cell = dayNameRow.getCell(INFO_COLS + 1 + i);
      const dow = allDates[i].getDay();
      cell.value = dayNames[dow].charAt(0);
      const isWeekend = dow === 0 || dow === 6;
      cell.font = { size: 7, color: { argb: isWeekend ? "FFef4444" : "FFcbd5e1" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isWeekend ? "FF1a1a2e" : "FF334155" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = cellBorder;
    }

    // --- Task rows ---
    const sprintNums = [...new Set(tasks.map(t => t.sprint))].sort((a, b) => a - b);
    let rowIdx = 4;

    for (const sprint of sprintNums) {
      const sprintTasks = tasks.filter(t => t.sprint === sprint);
      const bg = sprintBg[sprint] || "FFf1f5f9";
      const txtColor = sprintText[sprint] || "FF475569";

      // Sprint header row
      const spRow = ws.getRow(rowIdx);
      spRow.height = 22;
      ws.mergeCells(rowIdx, 1, rowIdx, INFO_COLS);
      const spCell = spRow.getCell(1);
      spCell.value = `Sprint ${sprint}  (${sprintTasks.length} tasks)`;
      spCell.font = { bold: true, size: 11, color: { argb: txtColor } };
      spCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      spCell.alignment = { vertical: "middle" };
      spCell.border = cellBorder;
      for (let c = 2; c <= INFO_COLS; c++) {
        spRow.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
        spRow.getCell(c).border = cellBorder;
      }
      // Light fill across date columns for sprint row
      for (let i = 0; i < allDates.length; i++) {
        const c = spRow.getCell(INFO_COLS + 1 + i);
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
        c.border = { left: thinBorder, right: thinBorder, top: thinBorder, bottom: thinBorder };
      }
      rowIdx++;

      // Task rows
      for (const task of sprintTasks) {
        const row = ws.getRow(rowIdx);
        row.height = 20;

        // Key
        const keyCell = row.getCell(1);
        keyCell.value = task.key;
        keyCell.font = { size: 9, bold: true, color: { argb: statusTextColors[task.status] || "FF475569" } };
        keyCell.alignment = { vertical: "middle" };
        keyCell.border = cellBorder;

        // Summary
        const sumCell = row.getCell(2);
        sumCell.value = task.summary;
        sumCell.font = { size: 9, color: { argb: "FF1e293b" } };
        sumCell.alignment = { vertical: "middle", wrapText: true };
        sumCell.border = cellBorder;

        // Status
        const stCell = row.getCell(3);
        stCell.value = task.status;
        stCell.font = { size: 9, bold: true, color: { argb: statusTextColors[task.status] || "FF475569" } };
        stCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: barBg[task.status] || "FFFFFFFF" } };
        stCell.alignment = { horizontal: "center", vertical: "middle" };
        stCell.border = cellBorder;

        // Start
        const startCell = row.getCell(4);
        startCell.value = task.start;
        startCell.font = { size: 9, color: { argb: "FF475569" } };
        startCell.alignment = { horizontal: "center", vertical: "middle" };
        startCell.border = cellBorder;

        // End
        const endCell = row.getCell(5);
        endCell.value = task.end;
        endCell.font = { size: 9, color: { argb: "FF475569" } };
        endCell.alignment = { horizontal: "center", vertical: "middle" };
        endCell.border = cellBorder;

        // Gantt bar across date columns
        const taskStart = new Date(task.start);
        const taskEnd = new Date(task.end);
        for (let i = 0; i < allDates.length; i++) {
          const cell = row.getCell(INFO_COLS + 1 + i);
          const d = allDates[i];
          const isWeekend = d.getDay() === 0 || d.getDay() === 6;
          const inRange = d >= taskStart && d < taskEnd;

          if (inRange) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: barColors[task.status] || "FF94a3b8" } };
          } else if (isWeekend) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFf8fafc" } };
          }
          cell.border = { left: { style: "hair", color: { argb: "FFe2e8f0" } }, right: { style: "hair", color: { argb: "FFe2e8f0" } }, top: thinBorder, bottom: thinBorder };
        }

        rowIdx++;
      }
    }

    // Freeze panes: freeze info columns + header rows
    ws.views = [{ state: "frozen", xSplit: INFO_COLS, ySplit: 3 }];
  }

  // --- Summary sheet ---
  const summary = wb.addWorksheet("Summary");
  summary.columns = [
    { width: 18 }, { width: 14 }, { width: 10 }, { width: 14 }, { width: 10 }, { width: 14 },
  ];
  const sHeaders = ["Project", "Total Tasks", "Done", "In Progress", "To Do", "Progress %"];
  const sHeaderRow = summary.getRow(1);
  sHeaderRow.height = 24;
  for (let c = 0; c < sHeaders.length; c++) {
    const cell = sHeaderRow.getCell(c + 1);
    cell.value = sHeaders[c];
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = cellBorder;
  }
  let sRow = 2;
  for (const project of projects) {
    const dataPath = `${import.meta.env.BASE_URL}data/${project.key}.json`;
    let tasks: Task[];
    if (dataCache[project.key]) {
      tasks = dataCache[project.key];
    } else {
      try {
        const res = await fetch(dataPath);
        if (!res.ok) continue;
        tasks = await res.json();
      } catch { continue; }
    }
    const done = tasks.filter(t => t.status === "Done").length;
    const inProg = tasks.filter(t => t.status === "In Progress").length;
    const toDo = tasks.filter(t => t.status === "To Do").length;
    const pct = tasks.length > 0 ? Math.round((done / tasks.length) * 100) : 0;

    const row = summary.getRow(sRow);
    row.getCell(1).value = project.name;
    row.getCell(1).font = { bold: true };
    row.getCell(2).value = tasks.length;
    row.getCell(3).value = done;
    row.getCell(3).font = { color: { argb: "FF22c55e" }, bold: true };
    row.getCell(4).value = inProg;
    row.getCell(4).font = { color: { argb: "FFf59e0b" }, bold: true };
    row.getCell(5).value = toDo;
    row.getCell(6).value = `${pct}%`;
    for (let c = 1; c <= 6; c++) {
      row.getCell(c).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(c).border = cellBorder;
    }
    sRow++;
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, `MMSP-DevChart-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

export default function GanttChart() {
  const [activeProject, setActiveProject] = useState<string>(projects[0].key);
  const [exporting, setExporting] = useState(false);
  const current = projects.find(p => p.key === activeProject)!;

  const handleExport = useCallback(async () => {
    setExporting(true);
    try {
      await exportAllToExcel();
    } catch (e) {
      console.error("Export failed:", e);
    } finally {
      setExporting(false);
    }
  }, []);

  return (
    <div style={{ fontFamily: "'Inter',sans-serif", background: "#0a0f1e", minHeight: "100vh", padding: 24, color: "#e2e8f0" }}>
      {/* Title */}
      <div style={{ marginBottom: 20, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <h1 style={{ margin: "0 0 4px", fontSize: 20, fontWeight: 700, color: "#f8fafc" }}>MMSP — Development Gantt Chart</h1>
          <p style={{ margin: 0, fontSize: 13, color: "#64748b" }}>Sprint 1 (Mar) · Sprint 2 (Apr) · Sprint 3 (May)</p>
        </div>
        <button onClick={handleExport} disabled={exporting}
          style={{ background: "#1e293b", border: "1px solid #334155", color: "#94a3b8", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: exporting ? "wait" : "pointer", display: "flex", alignItems: "center", gap: 6, transition: "all 0.15s", opacity: exporting ? 0.6 : 1 }}>
          <span style={{ fontSize: 14 }}>&#128196;</span>
          {exporting ? "Exporting..." : "Export to Excel"}
        </button>
      </div>

      {/* Project tabs */}
      <div style={{ display: "flex", gap: 4, marginBottom: 20, borderBottom: "1px solid #1f2937", paddingBottom: 0 }}>
        {projects.map(p => (
          <button key={p.key} onClick={() => setActiveProject(p.key)}
            style={{
              background: activeProject === p.key ? "#111827" : "transparent",
              border: activeProject === p.key ? "1px solid #1f2937" : "1px solid transparent",
              borderBottom: activeProject === p.key ? "1px solid #111827" : "1px solid transparent",
              marginBottom: -1,
              color: activeProject === p.key ? p.color : "#64748b",
              borderRadius: "8px 8px 0 0",
              padding: "8px 18px",
              fontSize: 13,
              fontWeight: activeProject === p.key ? 700 : 500,
              cursor: "pointer",
              transition: "all 0.15s",
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: activeProject === p.key ? p.color : "#475569" }} />
            {p.name}
          </button>
        ))}
      </div>

      {/* Active project chart */}
      <ProjectChart key={current.key} project={current} />

      {/* Legend */}
      <div style={{ display: "flex", gap: 20, marginTop: 14, flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div style={{ width: 3, height: 14, background: "#ef4444", borderRadius: 2 }} />
          <span style={{ fontSize: 11, color: "#64748b" }}>Today</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div style={{ width: 3, height: 14, background: "#334155", borderRadius: 2 }} />
          <span style={{ fontSize: 11, color: "#64748b" }}>Sprint boundary</span>
        </div>
        {(["Done", "In Progress", "To Do"] as Status[]).map(s => (
          <div key={s} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div style={{ width: 12, height: 12, borderRadius: 3, background: statusColors[s].bar }} />
            <span style={{ fontSize: 11, color: "#64748b" }}>{s}</span>
          </div>
        ))}
        {Object.keys(sprintConfig).map(Number).map(s => {
          const sc = sprintConfig[s];
          return (
            <div key={s} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 12, height: 12, borderRadius: "50%", background: sc.color }} />
              <span style={{ fontSize: 11, color: "#64748b" }}>{sc.label}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
