import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import GanttChart from "../gantt-chart";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <GanttChart />
  </StrictMode>
);
