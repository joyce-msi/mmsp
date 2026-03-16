import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import dotenv from "dotenv";
import * as fs from "fs";
import * as path from "path";

dotenv.config();

const JIRA_BASE = "https://msi-global.atlassian.net";
const AUTH = Buffer.from(
  `${process.env.JIRA_EMAIL}:${process.env.JIRA_API_TOKEN}`
).toString("base64");

export default defineConfig({
  base: "/mmsp/",
  plugins: [
    react(),
    {
      name: "jira-api",
      configureServer(server) {
        // Serve data/*.json files
        server.middlewares.use("/data", (req, res, next) => {
          const filePath = path.join(__dirname, "data", req.url || "");
          if (fs.existsSync(filePath) && filePath.endsWith(".json")) {
            res.setHeader("Content-Type", "application/json");
            res.end(fs.readFileSync(filePath, "utf-8"));
          } else {
            next();
          }
        });

        // Live fetch from Jira for a project: /api/sync/PROJECTKEY
        server.middlewares.use("/api/sync/", async (req, res) => {
          const projectKey = (req.url || "").replace(/^\//, "").replace(/\/$/, "");
          if (!projectKey) {
            res.statusCode = 400;
            res.end(JSON.stringify({ error: "Missing project key" }));
            return;
          }

          const sprintDateRanges: Record<number, { start: string; end: string }> = {
            1: { start: "2026-03-11", end: "2026-03-31" },
            2: { start: "2026-04-01", end: "2026-04-30" },
            3: { start: "2026-05-01", end: "2026-05-29" },
          };

          function mapStatus(s: string) {
            const l = s.toLowerCase();
            if (l === "done" || l === "closed" || l === "resolved" || l === "ready for release" || l === "in review") return "Done";
            if (l === "in progress" || l === "in development") return "In Progress";
            return "To Do";
          }

          function buildSprintMap(issues: Array<{ fields: Record<string, unknown> }>): Map<number, number> {
            const sprintInfo = new Map<number, string>();
            for (const issue of issues) {
              const sprints = issue.fields.customfield_10006 as Array<{ id: number; name: string }> | null;
              if (sprints) {
                for (const sp of sprints) {
                  if (!sprintInfo.has(sp.id)) sprintInfo.set(sp.id, sp.name);
                }
              }
            }
            const sortedIds = [...sprintInfo.keys()].sort((a, b) => a - b);
            const idToNumber = new Map<number, number>();
            for (let i = 0; i < sortedIds.length; i++) {
              const id = sortedIds[i];
              const name = sprintInfo.get(id)!;
              const match = name.match(/(\d+)/);
              idToNumber.set(id, match ? parseInt(match[1], 10) : i + 1);
            }
            return idToNumber;
          }

          function extractSprint(fields: Record<string, unknown>, sprintMap: Map<number, number>): number {
            // 1. Check labels (e.g. "sprint-1-sjt-march")
            const labels = (fields.labels as string[]) || [];
            for (const label of labels) {
              const match = label.match(/^sprint-(\d+)/i);
              if (match) return parseInt(match[1], 10);
            }
            // 2. Check sprint field — only use if sprint name contains a number
            const sprints = fields.customfield_10006 as Array<{ id: number; name: string }> | null;
            if (sprints && sprints.length > 0) {
              const nameMatch = sprints[sprints.length - 1].name.match(/(\d+)/);
              if (nameMatch) return parseInt(nameMatch[1], 10);
            }
            return -1;
          }

          try {
            const issues: Array<{ key: string; fields: Record<string, unknown> }> = [];
            let nextPageToken: string | null = null;

            while (true) {
              const params = new URLSearchParams({
                jql: `project = "${projectKey}" ORDER BY rank ASC`,
                maxResults: "100",
                fields: "summary,status,assignee,customfield_11702,duedate,customfield_10006,labels",
              });
              if (nextPageToken) params.set("nextPageToken", nextPageToken);

              const r = await fetch(`${JIRA_BASE}/rest/api/3/search/jql?${params}`, {
                headers: { Authorization: `Basic ${AUTH}`, Accept: "application/json" },
              });
              if (!r.ok) {
                const text = await r.text();
                res.statusCode = r.status;
                res.end(JSON.stringify({ error: text }));
                return;
              }
              const data = await r.json();
              issues.push(...data.issues);
              if (data.isLast) break;
              nextPageToken = data.nextPageToken;
              if (!nextPageToken) break;
            }

            // Filter to sprint-assigned issues only
            const sprintIssues = issues.filter(i => {
              const sprints = i.fields.customfield_10006 as unknown[] | null;
              return sprints && sprints.length > 0;
            });

            const sprintMap = buildSprintMap(issues);

            const tasks = sprintIssues.map(i => {
              const sprint = extractSprint(i.fields, sprintMap);
              const fallback = sprintDateRanges[sprint] || sprintDateRanges[1];
              const assignee = (i.fields.assignee as { displayName: string } | null)?.displayName;
              return {
                key: i.key,
                summary: (i.fields.summary as string) || "",
                start: (i.fields.customfield_11702 as string) || fallback.start,
                end: (i.fields.duedate as string) || fallback.end,
                status: mapStatus(((i.fields.status as { name: string }) || { name: "To Do" }).name),
                sprint,
                ...(assignee ? { assignee } : {}),
              };
            }).filter(t => t.sprint >= 1);

            tasks.sort((a, b) => a.sprint - b.sprint || a.start.localeCompare(b.start));

            // Also save to local data files (both data/ and public/data/)
            const json = JSON.stringify(tasks, null, 2);
            for (const dir of ["data", path.join("public", "data")]) {
              const dataDir = path.join(__dirname, dir);
              if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
              fs.writeFileSync(path.join(dataDir, `${projectKey}.json`), json);
            }

            res.setHeader("Content-Type", "application/json");
            res.end(JSON.stringify(tasks));
          } catch (e: unknown) {
            res.statusCode = 500;
            res.end(JSON.stringify({ error: e instanceof Error ? e.message : String(e) }));
          }
        });

        server.middlewares.use("/api/update-tasks", async (req, res) => {
          if (req.method !== "POST") {
            res.statusCode = 405;
            res.end("Method not allowed");
            return;
          }

          let body = "";
          req.on("data", (chunk: Buffer) => (body += chunk));
          req.on("end", async () => {
            try {
              const changes: Array<{
                key: string;
                summary?: string;
                start?: string;
                end?: string;
                status?: string;
              }> = JSON.parse(body);

              const results: Array<{ key: string; success: boolean; error?: string }> = [];

              for (const task of changes) {
                try {
                  // Update fields (summary, dates)
                  const fields: Record<string, unknown> = {};
                  if (task.summary !== undefined) fields.summary = task.summary;
                  if (task.start !== undefined) fields.customfield_11702 = task.start;
                  if (task.end !== undefined) fields.duedate = task.end;

                  if (Object.keys(fields).length > 0) {
                    const r = await fetch(`${JIRA_BASE}/rest/api/3/issue/${task.key}`, {
                      method: "PUT",
                      headers: {
                        Authorization: `Basic ${AUTH}`,
                        "Content-Type": "application/json",
                      },
                      body: JSON.stringify({ fields }),
                    });
                    if (!r.ok) {
                      const text = await r.text();
                      results.push({ key: task.key, success: false, error: text });
                      continue;
                    }
                  }

                  // Update status via transitions
                  if (task.status !== undefined) {
                    const tr = await fetch(
                      `${JIRA_BASE}/rest/api/3/issue/${task.key}/transitions`,
                      {
                        headers: {
                          Authorization: `Basic ${AUTH}`,
                          Accept: "application/json",
                        },
                      }
                    );
                    const data = await tr.json();
                    const target = data.transitions.find((t: { name: string }) => {
                      const name = t.name.toLowerCase();
                      if (task.status === "Done")
                        return name === "done" || name.includes("done");
                      if (task.status === "In Progress")
                        return name === "in progress" || name.includes("progress");
                      return name === "to do" || name.includes("to do");
                    });
                    if (target) {
                      await fetch(
                        `${JIRA_BASE}/rest/api/3/issue/${task.key}/transitions`,
                        {
                          method: "POST",
                          headers: {
                            Authorization: `Basic ${AUTH}`,
                            "Content-Type": "application/json",
                          },
                          body: JSON.stringify({ transition: { id: target.id } }),
                        }
                      );
                    }
                  }

                  results.push({ key: task.key, success: true });
                } catch (e: unknown) {
                  results.push({
                    key: task.key,
                    success: false,
                    error: e instanceof Error ? e.message : String(e),
                  });
                }
              }

              res.setHeader("Content-Type", "application/json");
              res.end(JSON.stringify({ results }));
            } catch (e: unknown) {
              res.statusCode = 500;
              res.end(
                JSON.stringify({
                  error: e instanceof Error ? e.message : String(e),
                })
              );
            }
          });
        });
      },
    },
  ],
});
