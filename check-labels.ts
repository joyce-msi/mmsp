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
    customfield_10006?: Array<{ name: string; id: number; state: string }> | null;
    labels?: string[];
  };
}

async function fetchAllIssues(projectKey: string): Promise<JiraIssue[]> {
  const issues: JiraIssue[] = [];
  let nextPageToken: string | null = null;
  const maxResults = 100;

  while (true) {
    const jql = `project = "${projectKey}" ORDER BY rank ASC`;
    const params = new URLSearchParams({
      jql,
      maxResults: String(maxResults),
      fields: "summary,status,customfield_10006,labels",
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

function getSprintName(issue: JiraIssue): string | null {
  const sprints = issue.fields.customfield_10006;
  if (sprints && sprints.length > 0) {
    return sprints[sprints.length - 1].name;
  }
  return null;
}

function hasSprintLabel(issue: JiraIssue): boolean {
  const labels = issue.fields.labels || [];
  return labels.some(l => /^sprint-\d+/i.test(l));
}

async function checkProject(project: ProjectConfig): Promise<number> {
  console.log(`\n${"=".repeat(60)}`);
  console.log(`${project.name} (${project.key})`);
  console.log("=".repeat(60));

  const issues = await fetchAllIssues(project.key);
  console.log(`Total issues: ${issues.length}`);

  // Group issues by sprint
  const sprintGroups = new Map<string, JiraIssue[]>();
  const backlog: JiraIssue[] = [];

  for (const issue of issues) {
    const sprintName = getSprintName(issue);
    if (sprintName) {
      if (!sprintGroups.has(sprintName)) sprintGroups.set(sprintName, []);
      sprintGroups.get(sprintName)!.push(issue);
    } else {
      backlog.push(issue);
    }
  }

  let totalMissing = 0;

  // Check each sprint
  for (const [sprintName, sprintIssues] of sprintGroups) {
    const missing = sprintIssues.filter(i => !hasSprintLabel(i));
    const icon = missing.length === 0 ? "OK" : "MISSING";
    console.log(`\n  Sprint: "${sprintName}" (${sprintIssues.length} issues) [${icon}]`);

    if (missing.length > 0) {
      totalMissing += missing.length;
      for (const issue of missing) {
        const labels = (issue.fields.labels || []).join(", ") || "(none)";
        console.log(`    ${issue.key} | ${issue.fields.status.name} | ${issue.fields.summary}`);
        console.log(`      Labels: ${labels}`);
      }
    }
  }

  // Check backlog
  const backlogWithLabels = backlog.filter(i => hasSprintLabel(i));
  if (backlogWithLabels.length > 0) {
    console.log(`\n  Backlog issues WITH sprint labels (possibly misplaced):`);
    for (const issue of backlogWithLabels) {
      const labels = (issue.fields.labels || []).join(", ");
      console.log(`    ${issue.key} | ${issue.fields.status.name} | ${issue.fields.summary}`);
      console.log(`      Labels: ${labels}`);
    }
  }

  const backlogNoLabel = backlog.filter(i => !hasSprintLabel(i));
  console.log(`\n  Backlog: ${backlog.length} issues (${backlogNoLabel.length} without labels, ${backlogWithLabels.length} with sprint labels)`);

  if (totalMissing === 0) {
    console.log(`\n  All sprint issues have labels.`);
  } else {
    console.log(`\n  ${totalMissing} sprint issues MISSING labels!`);
  }

  return totalMissing;
}

async function main() {
  const projectsPath = path.join(import.meta.dirname, "projects.json");
  const projects: ProjectConfig[] = JSON.parse(fs.readFileSync(projectsPath, "utf-8"));

  console.log(`Checking sprint labels for ${projects.length} projects...\n`);

  let grandTotal = 0;
  for (const project of projects) {
    grandTotal += await checkProject(project);
  }

  console.log(`\n${"=".repeat(60)}`);
  if (grandTotal === 0) {
    console.log("All sprint issues across all projects have labels.");
  } else {
    console.log(`${grandTotal} total sprint issues missing labels across all projects.`);
  }
}

main();
