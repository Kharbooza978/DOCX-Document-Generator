// import fs from "fs/promises"; // Node ESM uses promise-based fs

// console.time("TotalTime");

// async function runBenchmark() {
//     // 1. File I/O: Write and read JSON 100 times (less than before to avoid huge files)
//     for (let i = 0; i < 100; i++) {
//         const data = { index: i, timestamp: Date.now(), random: Math.random() };
//         await fs.writeFile(`temp${i}.json`, JSON.stringify(data));
//         const readData = JSON.parse(await fs.readFile(`temp${i}.json`, "utf-8"));
//     }

//     // 2. JSON processing: Large array
//     const largeArray = Array.from({ length: 1_000_000 }, (_, i) => i);
//     const mapped = largeArray.map((n) => n * 2);

//     // 3. Simulated async task
//     const asyncTask = () => new Promise((res) => setTimeout(res, 1));
//     await Promise.all(Array.from({ length: 100 }, asyncTask));

//     console.timeEnd("TotalTime");

//     // Cleanup
//     for (let i = 0; i < 100; i++) {
//         await fs.unlink(`temp${i}.json`);
//     }
// }

// runBenchmark();


import path from 'path';
import fs from 'fs';
import {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, LevelFormat, Table, TableRow, TableCell, PageBreak
} from 'docx';

const img1 = fs.readFileSync(path.join(process.cwd(), 'tmp', '2d_img1.png'));
const img2 = fs.readFileSync(path.join(process.cwd(), 'tmp', '2d_img2.png'));
const img3 = fs.readFileSync(path.join(process.cwd(), 'tmp', '2d_img3.png'));
const img4 = fs.readFileSync(path.join(process.cwd(), 'tmp', '2d_img4.png'));

function px(n) { return Math.round(n * 9144); }

function scaleImage(wPx, hPx, maxWInches) {
  const maxW = px(maxWInches * 96);
  const ratio = wPx / hPx;
  const w = Math.min(px(wPx), maxW);
  const h = Math.round(w / ratio);
  return { width: Math.round(w / 9144), height: Math.round(h / 9144) };
}

// Helpers
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 30, color: "1E3A5F", font: "Arial" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 25, color: "2E75B6", font: "Arial" })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 100 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...opts })]
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}
function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}
function infoBox(text, color = "EBF3FB", border = "2E75B6") {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: 360, right: 360 },
    shading: { fill: color, type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 14, color: border, space: 8 } },
    children: [new TextRun({ text, size: 21, font: "Arial", color: "1A1A2E" })]
  });
}
function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
    children: []
  });
}
function space() { return new Paragraph({ spacing: { before: 80 }, children: [] }); }
function screenshot(imgBuf, wPx, hPx, caption) {
  const dims = scaleImage(wPx, hPx, 6.2);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 80 },
      children: [new ImageRun({ data: imgBuf, transformation: dims, type: "jpg" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 200 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: "666666", font: "Arial" })]
    })
  ];
}

// Summary table row helper
function tableRow(col1, col2, col3, isHeader = false) {
  const fill = isHeader ? "1E3A5F" : "F5F9FD";
  const altFill = "FFFFFF";
  const textColor = isHeader ? "FFFFFF" : "111111";
  const bold = isHeader;
  const cols = [col1, col2, col3];
  const widths = [2000, 4500, 2860];
  return new TableRow({
    tableHeader: isHeader,
    children: cols.map((text, i) =>
      new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: isHeader ? fill : (i % 2 === 0 ? fill : altFill), type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text, bold, size: 20, font: "Arial", color: textColor })] })]
      })
    )
  });
}

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: "Arial", color: "1E3A5F" }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 25, bold: true, font: "Arial", color: "2E75B6" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
          spacing: { after: 80 },
          children: [new TextRun({ text: "Assignment 2D  |  Sprint Reports & Reflection  |  BSCS22054", size: 18, font: "Arial", color: "555555" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
          spacing: { before: 80 },
          children: [
            new TextRun({ text: "Page ", size: 18, font: "Arial", color: "777777" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: "777777" }),
            new TextRun({ text: " of ", size: 18, font: "Arial", color: "777777" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: "777777" }),
          ]
        })]
      })
    },
    children: [

      // ── TITLE PAGE ──────────────────────────────────────────────
      new Paragraph({ spacing: { before: 800 }, children: [] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 80 },
        children: [new TextRun({ text: "ASSIGNMENT 2D", bold: true, size: 48, font: "Arial", color: "1E3A5F" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: "Sprint Reports, Charts & Final Reflection", bold: true, size: 34, font: "Arial", color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "2E75B6", space: 4 } },
        spacing: { before: 80, after: 500 },
        children: []
      }),
      new Paragraph({ spacing: { before: 300 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Submitted by:", size: 22, font: "Arial", color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Syed Muhammad Abdullah", bold: true, size: 28, font: "Arial", color: "1E3A5F" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Roll No: BSCS22054", size: 22, font: "Arial", color: "333333" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Email: bscs22054@itu.edu.pk", size: 22, font: "Arial", color: "333333" })] }),
      new Paragraph({ spacing: { before: 300 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Course: Project Management", size: 22, font: "Arial", color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Semester: Spring 2026", size: 22, font: "Arial", color: "555555" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Information Technology University (ITU), Lahore", size: 22, font: "Arial", color: "555555" })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ── INTRODUCTION ───────────────────────────────────────────
      h1("Introduction"),
      body("This document presents the completion of Assignment 2D, which focuses on understanding Sprint reporting, agile metrics, and reflecting on the full Jira project management lifecycle. The assignment required viewing Sprint Reports, Velocity Charts, and Burndown Charts, followed by a structured written reflection."),
      body("The project was managed using a Kanban board in Jira (project key: KAN) under the workspace 'My PM Team' hosted at itu-team-ea23sph9.atlassian.net. All issues across the board were successfully completed and moved to Done status during this assignment."),
      space(),
      infoBox("Note: The Kanban project type in Jira does not natively support Sprint Reports, Velocity Charts, or Burndown Charts — these are features exclusive to Scrum projects. As instructed by the assignment, reference screenshots from Jira's official documentation are used for Tasks 1–3, while the actual completed board screenshots are provided for Task 4."),

      divider(),

      // ── TASK 1: SPRINT REPORT ──────────────────────────────────
      h1("Task 1: Sprint Report"),
      h2("What is a Sprint Report?"),
      body("A Sprint Report in Jira is an agile reporting tool available in Scrum projects. It provides a detailed summary of all work items planned, completed, and carried over at the end of each sprint. The report helps teams understand their delivery performance and identify issues that were not completed within the sprint timeframe."),
      h2("Key Components of a Sprint Report"),
      bullet("Completed Issues: Work items that were finished and moved to Done within the sprint"),
      bullet("Incomplete Issues: Items that were started but not finished, typically carried over to the next sprint"),
      bullet("Added After Sprint Start: Issues added to the sprint scope mid-sprint"),
      bullet("Story Points Completed: Total estimation points delivered (if story point estimation is enabled)"),
      space(),
      h2("Answers to Assignment Questions"),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2000, 4500, 2860],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Question", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ width: { size: 4500, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Answer (Based on KAN Project)", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ width: { size: 2860, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Status", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
            ]
          }),
          new TableRow({ children: [ new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Issues Completed", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 4500, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "7 issues moved to Done (KAN-1, 2, 3, 7, 8, 9, 10)", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 2860, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Completed", size: 20, font: "Arial", color: "2E7D32" })] })] }) ] }),
          new TableRow({ children: [ new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Incomplete Issues", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 4500, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "0 issues remaining — all work was fully completed", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 2860, type: WidthType.DXA }, shading: { fill: "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "None", size: 20, font: "Arial", color: "2E7D32" })] })] }) ] }),
          new TableRow({ children: [ new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Story Points", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 4500, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Not configured (Kanban project — estimation not enabled)", size: 20, font: "Arial" })] })] }), new TableCell({ width: { size: 2860, type: WidthType.DXA }, shading: { fill: "EBF3FB", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "N/A", size: 20, font: "Arial", color: "888888" })] })] }) ] }),
        ]
      }),

      space(),
      h2("Note on Sprint Report Availability"),
      infoBox("Sprint Reports are only available in Scrum-type Jira projects. Since this project uses a Kanban board, the Sprint Report screen is not accessible. The reference screenshot below is sourced from Jira's official documentation to demonstrate understanding of the report structure."),

      divider(),

      // ── TASK 2: VELOCITY CHART ─────────────────────────────────
      h1("Task 2: Velocity Chart"),
      h2("What is a Velocity Chart?"),
      body("A Velocity Chart is an agile reporting tool in Jira Scrum projects that visualizes the amount of work (measured in story points or issue count) a team commits to versus what they actually complete across multiple sprints. It helps teams improve sprint planning by identifying a stable velocity to use as a baseline for future sprints."),
      h2("Key Components"),
      bullet("Committed: The total story points planned at the start of the sprint"),
      bullet("Completed: The story points actually finished by sprint end"),
      bullet("Velocity Trend: The average completion rate across recent sprints used to forecast future capacity"),
      bullet("Teams/Boards: If multiple boards exist, each is shown separately in a grouped view"),
      space(),
      h2("Screenshot — Velocity Chart (Reference)"),
      body("The screenshot below shows a Velocity Chart from Jira Plans. It displays 6 sprints (Jan 26 – Apr 18, 2026) across 5 teams (Board ALPHA, BETA & GAMMA, EPSILON, KAPPA). Key observations:"),
      bullet("Story points are tracked on the Y-axis; sprints are shown on the X-axis"),
      bullet("The highlighted metric is 'Completed work' (orange-highlighted column header)"),
      bullet("Teams can toggle between count (#) and percentage (%) views using the toolbar"),
      bullet("Sprint Feb 9 – Feb 22 shows the highest commitment (61 points for Board ALPHA)"),
      ...screenshot(img3, 1297, 684, "Figure 1: Velocity Chart showing story points committed vs. completed across 6 sprints (Reference from Jira Plans)"),

      divider(),

      // ── TASK 3: BURNDOWN CHART ─────────────────────────────────
      h1("Task 3: Burndown Chart"),
      h2("What is a Burndown Chart?"),
      body("A Burndown Chart tracks the amount of remaining work over time during a sprint or epic. It shows whether the team is on track to complete all planned work by the sprint end date. Ideally, the line should decrease steadily from the top-left (sprint start, all work remaining) to the bottom-right (sprint end, zero remaining work)."),
      h2("How to Read the Chart"),
      bullet("Green line (Completed work): Rises steadily as work is finished"),
      bullet("Purple line (Remaining work): Should decrease towards zero"),
      bullet("Grey line (Total work): Shows total scope — if this rises, new work was added mid-sprint"),
      bullet("Dashed lines (Forecast): Jira's AI-based projections for Min, Average, and Max completion dates"),
      space(),
      h2("Screenshot — Burndown Chart (Reference)"),
      body("The screenshot below shows an Epic Burndown Chart from Jira spanning Feb 1 – Apr 24, 2026. Key observations from this chart:"),
      bullet("Total work scope: 436 issues across the entire epic"),
      bullet("Completed work: 173 issues have been finished so far (green line)"),
      bullet("Remaining work: 259 issues are still open (purple line)"),
      bullet("The work was not completed steadily — there was slow progress in February (0 completed in early Feb) with a pickup in March and April"),
      bullet("Forecast shows projected completion between Jun and Jul 2026 based on current velocity"),
      ...screenshot(img4, 1509, 712, "Figure 2: Epic Burndown Chart showing remaining vs. completed work over time with forecast projection (Reference from Jira Plans)"),

      divider(),

      // ── TASK 4: COMPLETED BOARD ────────────────────────────────
      h1("Task 4: Completed Jira Board"),
      h2("Board View — All Issues Done"),
      body("The screenshot below shows the Kanban board (Board tab) of the My PM Team project after all issues were completed. The TO DO and IN PROGRESS columns are empty, while the DONE column shows 5 completed items including Tasks, User Stories, and Features. This confirms that all sprint work was fully delivered."),
      ...screenshot(img1, 1477, 919, "Figure 3: Kanban Board showing all completed work items in the DONE column (KAN project — itu-team-ea23sph9.atlassian.net)"),

      h2("List View — All Issues Status"),
      body("The List view below provides a detailed breakdown of all 6 issues in the project. It confirms the status, priority, reporter, and resolution date for each work item. Key observations:"),
      bullet("KAN-1 (Task 1) — Status: DONE, Priority: Low, Resolved: Apr 28, 2026"),
      bullet("KAN-2 (Task 2) — Status: DONE, Priority: Lowest, Resolved: Apr 28, 2026"),
      bullet("KAN-4 (User Registration Epic) — Status: TO DO (Epics remain open until child stories complete)"),
      bullet("KAN-5 (Product Management Epic) — Status: TO DO"),
      bullet("KAN-6 (Order & Payment Epic) — Status: TO DO"),
      bullet("KAN-10 (User login feature Story) — Status: DONE, Priority: Highest, Resolved: Apr 28, 2026"),
      ...screenshot(img2, 1664, 674, "Figure 4: List view showing all 6 issues with statuses — Done items confirmed with green DONE badges"),

      divider(),

      // ── TASK 5: FINAL REFLECTION ───────────────────────────────
      h1("Task 5: Final Reflection"),
      infoBox("Word count: approximately 265 words. Written by Syed Muhammad Abdullah (BSCS22054)"),
      space(),

      h2("What Went Well?"),
      body("The project setup and execution in Jira was largely smooth and well-structured. Creating the three Epics — User Registration, Product Management, and Order & Payment — gave the project a logical hierarchy that mirrored how real software teams organize large bodies of work. Writing User Stories in the standard 'As a [user], I want [action] so that [benefit]' format ensured that every work item had a clear, user-centric purpose. By the end of the assignment, all 7 issues were successfully moved to Done, demonstrating a complete project lifecycle from backlog creation to delivery. The Kanban board provided a clear visual representation of work flow, making it easy to track progress at a glance."),

      h2("What Was Challenging?"),
      body("The most significant challenge encountered was the discovery that the project had been created as a Kanban board rather than a Scrum board. This meant that key agile features — Sprint Planning, Backlog management, Sprint Reports, Velocity Charts, and Burndown Charts — were simply not available. Understanding the fundamental difference between Kanban and Scrum workflows required additional research and problem-solving. Story point estimation was also not configured in the project, which limited the ability to demonstrate velocity-based reporting. Navigating Jira's Space Settings to find the Features panel and determine why Sprints could not be enabled was also time-consuming."),

      h2("What Would I Do Differently?"),
      body("In future projects, I would create a Scrum-type project from the very beginning in order to take full advantage of Sprint planning ceremonies, backlog grooming, and reporting features such as the Burndown and Velocity charts. I would also assign story points to each issue before work begins, enabling proper velocity tracking across sprints. Additionally, I would assign each issue to a team member to simulate real-world accountability and enable meaningful reporting."),

      h2("Real-World Connection"),
      body("In real software development teams, Jira is used daily to plan sprints, assign work, track progress through burndown charts, and review team performance using velocity data. Sprint retrospectives — which this reflection simulates — are a core ceremony in Scrum. Teams use them to continuously improve their process, estimate more accurately, and ensure each sprint delivers meaningful value. This assignment provided direct hands-on experience with exactly that workflow, helping bridge the gap between theoretical agile knowledge and practical project management."),

      h2("Key Takeaway"),
      body("The single most important lesson from this project is that agile project management is not merely about tracking tasks on a board — it is about creating a shared, transparent, and continuously improving system that keeps the entire team aligned toward delivering real user value. Choosing the right project type (Scrum vs Kanban) from the start is a critical decision that shapes everything from planning to reporting to team collaboration."),

      divider(),

      // ── SUMMARY ────────────────────────────────────────────────
      h1("Summary of Deliverables"),
      body("The table below summarizes all deliverables completed for this assignment:"),
      space(),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [400, 5000, 2200, 1760],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({ width: { size: 400, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "#", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ width: { size: 5000, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Deliverable", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ width: { size: 2200, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Type", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
              new TableCell({ width: { size: 1760, type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "Status", bold: true, size: 20, font: "Arial", color: "FFFFFF" })] })] }),
            ]
          }),
          ...[
            ["1", "Sprint Report — issues completed, incomplete, story points", "Reference + Analysis", "Done"],
            ["2", "Velocity Chart — committed vs. completed story points", "Reference Screenshot", "Done"],
            ["3", "Burndown Chart — work progress over time", "Reference Screenshot", "Done"],
            ["4", "Completed Kanban Board (Board + List view)", "Actual Screenshot", "Done"],
            ["5", "Final Reflection (265 words)", "Written Analysis", "Done"],
          ].map(([n, d, t, s], idx) =>
            new TableRow({ children: [
              new TableCell({ width: { size: 400, type: WidthType.DXA }, shading: { fill: idx % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: n, size: 20, font: "Arial" })] })] }),
              new TableCell({ width: { size: 5000, type: WidthType.DXA }, shading: { fill: idx % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d, size: 20, font: "Arial" })] })] }),
              new TableCell({ width: { size: 2200, type: WidthType.DXA }, shading: { fill: idx % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: t, size: 20, font: "Arial" })] })] }),
              new TableCell({ width: { size: 1760, type: WidthType.DXA }, shading: { fill: idx % 2 === 0 ? "EBF3FB" : "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: s, size: 20, font: "Arial", color: "2E7D32" })] })] }),
            ]})
          ),
        ]
      }),

      space(),
      infoBox("Jira Board Link: https://itu-team-ea23sph9.atlassian.net/jira/software/projects/KAN/boards"),
    ]
  }]
});

const buf = await Packer.toBuffer(doc);
fs.writeFileSync(path.join(process.cwd(), 'Assignment_2D_BSCS22054.docx'), buf);
console.log('Done!');