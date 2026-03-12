import "dotenv/config";
import * as fs from "fs";
import * as path from "path";

const JIRA_BASE = "https://msi-global.atlassian.net";
const EMAIL = process.env.JIRA_EMAIL!;
const TOKEN = process.env.JIRA_API_TOKEN!;

if (!EMAIL || !TOKEN) {
  console.error("Missing JIRA_EMAIL or JIRA_API_TOKEN in .env file");
  process.exit(1);
}

const AUTH = Buffer.from(`${EMAIL}:${TOKEN}`).toString("base64");

interface ProjectConfig {
  key: string;
  name: string;
  color: string;
}

interface JiraIssue {
  key: string;
  fields: {
    summary: string;
    status: { name: string };
    customfield_11702: string | null; // Start Date
    duedate: string | null;
    customfield_10006?: Array<{ name: string; id: number; state: string }> | null; // Sprint
    labels?: string[];
  };
}

interface TaskData {
  key: string;
  summary: string;
  start: string;
  end: string;
  status: "Done" | "In Progress" | "To Do";
  sprint: number;
}

// Sprint date ranges used as fallback for tasks without dates
const sprintDateRanges: Record<number, { start: string; end: string }> = {
  1: { start: "2026-03-11", end: "2026-03-31" },
  2: { start: "2026-04-01", end: "2026-04-30" },
  3: { start: "2026-05-01", end: "2026-05-29" },
};

async function fetchAllIssues(projectKey: string): Promise<JiraIssue[]> {
  const issues: JiraIssue[] = [];
  let nextPageToken: string | null = null;
  const maxResults = 100;

  while (true) {
    const jql = `project = "${projectKey}" ORDER BY rank ASC`;
    const params = new URLSearchParams({
      jql,
      maxResults: String(maxResults),
      fields: "summary,status,customfield_11702,duedate,customfield_10006,labels",
    });
    if (nextPageToken) {
      params.set("nextPageToken", nextPageToken);
    }
    const url = `${JIRA_BASE}/rest/api/3/search/jql?${params}`;

    const res = await fetch(url, {
      headers: {
        Authorization: `Basic ${AUTH}`,
        Accept: "application/json",
      },
    });

    if (!res.ok) {
      const text = await res.text();
      console.error(`  Jira API error (${res.status}): ${text}`);
      return issues;
    }

    const data = await res.json();
    issues.push(...data.issues);

    if (data.isLast) break;
    nextPageToken = data.nextPageToken;
    if (!nextPageToken) break;
  }

  return issues;
}

function mapStatus(jiraStatus: string): "Done" | "In Progress" | "To Do" {
  const lower = jiraStatus.toLowerCase();
  if (lower === "done" || lower === "closed" || lower === "resolved" || lower === "ready for release" || lower === "in review") return "Done";
  if (lower === "in progress" || lower === "in development") return "In Progress";
  return "To Do";
}

function buildSprintMap(issues: JiraIssue[]): Map<number, number> {
  // Collect all unique sprint IDs and names
  const sprintInfo = new Map<number, string>();
  for (const issue of issues) {
    const sprints = issue.fields.customfield_10006;
    if (sprints) {
      for (const sp of sprints) {
        if (!sprintInfo.has(sp.id)) sprintInfo.set(sp.id, sp.name);
      }
    }
  }

  // Sort by ID (creation order) and assign Sprint 1, 2, 3...
  const sortedIds = [...sprintInfo.keys()].sort((a, b) => a - b);
  const idToNumber = new Map<number, number>();

  for (let i = 0; i < sortedIds.length; i++) {
    const id = sortedIds[i];
    const name = sprintInfo.get(id)!;

    // First try to extract a number from the name (e.g. "Sprint 2")
    const match = name.match(/(\d+)/);
    if (match) {
      idToNumber.set(id, parseInt(match[1], 10));
    } else {
      // Fall back to order-based numbering
      idToNumber.set(id, i + 1);
    }
  }

  return idToNumber;
}

function extractSprintNumber(issue: JiraIssue, sprintMap: Map<number, number>): number {
  // 1. Check labels (e.g. "sprint-1-sjt-march", "sprint-2-svc-april")
  const labels = issue.fields.labels || [];
  for (const label of labels) {
    const match = label.match(/^sprint-(\d+)/i);
    if (match) return parseInt(match[1], 10);
  }

  // 2. Check sprint field — only use if the sprint name contains a number
  const sprints = issue.fields.customfield_10006;
  if (sprints && sprints.length > 0) {
    const lastSprint = sprints[sprints.length - 1];
    const nameMatch = lastSprint.name.match(/(\d+)/);
    if (nameMatch) return parseInt(nameMatch[1], 10);
  }

  // No sprint label and no numbered sprint name — exclude
  return -1;
}

async function syncProject(project: ProjectConfig): Promise<void> {
  console.log(`\nSyncing ${project.name} (${project.key})...`);
  const issues = await fetchAllIssues(project.key);
  console.log(`  Found ${issues.length} issues`);

  // Build sprint ID → sprint number mapping
  const sprintMap = buildSprintMap(issues);
  const sprintNames = new Map<number, string>();
  for (const issue of issues) {
    const sprints = issue.fields.customfield_10006;
    if (sprints) {
      for (const sp of sprints) {
        const num = sprintMap.get(sp.id);
        if (num && !sprintNames.has(num)) sprintNames.set(num, sp.name);
      }
    }
  }
  for (const [num, name] of [...sprintNames.entries()].sort((a, b) => a[0] - b[0])) {
    console.log(`  Sprint ${num}: "${name}"`);
  }

  // Only include issues assigned to a sprint
  const sprintIssues = issues.filter(i => {
    const sprints = i.fields.customfield_10006;
    return sprints && sprints.length > 0;
  });

  const backlogCount = issues.length - sprintIssues.length;
  if (backlogCount > 0) {
    console.log(`  Skipped ${backlogCount} backlog issues (no sprint assigned)`);
  }

  const allTasks = sprintIssues.map(i => {
    const sprint = extractSprintNumber(i, sprintMap);
    const fallback = sprintDateRanges[sprint] || sprintDateRanges[1];
    return {
      key: i.key,
      summary: i.fields.summary,
      start: i.fields.customfield_11702 || fallback.start,
      end: i.fields.duedate || fallback.end,
      status: mapStatus(i.fields.status.name),
      sprint,
    };
  });

  // Exclude tasks without a valid sprint label mapping
  const excluded = allTasks.filter(t => t.sprint < 1);
  if (excluded.length > 0) {
    console.log(`  Skipped ${excluded.length} issues without a sprint label`);
  }
  const tasks: TaskData[] = allTasks.filter(t => t.sprint >= 1);

  const withoutDates = sprintIssues.filter(i => !i.fields.customfield_11702 || !i.fields.duedate).length;
  if (withoutDates > 0) {
    console.log(`  ${withoutDates} tasks without start/due dates — using sprint date range as fallback`);
  }

  // Sort by sprint then start date
  tasks.sort((a, b) => a.sprint - b.sprint || a.start.localeCompare(b.start));

  // Write to data file
  const dataDir = path.join(import.meta.dirname, "data");
  if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir);

  const outPath = path.join(dataDir, `${project.key}.json`);
  fs.writeFileSync(outPath, JSON.stringify(tasks, null, 2));
  console.log(`  Wrote ${tasks.length} tasks to data/${project.key}.json`);
}

async function main() {
  const projectsPath = path.join(import.meta.dirname, "projects.json");
  const projects: ProjectConfig[] = JSON.parse(fs.readFileSync(projectsPath, "utf-8"));

  console.log(`Syncing ${projects.length} projects from Jira...`);

  for (const project of projects) {
    await syncProject(project);
  }

  console.log("\nDone!");
}

main();
