/**
 * ============================================================
 *  Gantt Chart Generator — PptxGenJS
 *  Generates a modern, minimalist Gantt Chart as a .pptx file.
 * ============================================================
 *
 * USAGE / HELP:
 *   node gantt_chart.js '<JSON_STRING>'
 *
 * EXAMPLE COMMAND:
 *   node gantt_chart.js '{"projectName":"Project Alpha","tasks":[{"name":"Discovery & Research","start":"2024-01-01","end":"2024-01-10","team":["Alice","Bob"]},{"name":"Design Phase","start":"2024-01-08","end":"2024-01-22","team":["Carol","Dave"]},{"name":"Prototype Build","start":"2024-01-20","end":"2024-02-05","team":["Eve","Frank","Grace"]},{"name":"Milestone 1","date":"2024-01-25","type":"milestone"},{"name":"User Testing","start":"2024-02-03","end":"2024-02-14","team":["Alice","Carol"]},{"name":"Iteration & Polish","start":"2024-02-12","end":"2024-02-26","team":["Dave","Eve"]},{"name":"Final Review","date":"2024-02-28","type":"milestone"}],"gates":["2024-01-15","2024-02-10"]}'
 *
 * OUTPUT:
 *   gantt_chart.pptx (saved in current working directory)
 *
 * SCHEMA:
 * {
 *   "projectName": "string",
 *   "tasks": [
 *     // Regular task:
 *     { "name": "string", "start": "YYYY-MM-DD", "end": "YYYY-MM-DD", "team": ["string"] },
 *     // Milestone:
 *     { "name": "string", "date": "YYYY-MM-DD", "type": "milestone" }
 *   ],
 *   "gates": ["YYYY-MM-DD"]  // optional
 * }
 */

"use strict";

const pptxgen = require("pptxgenjs");

// ─── COLOUR & FONT CONSTANTS ────────────────────────────────────────────────
const C = {
  // Blues
  PRIMARY:        "2563EB",   // Main blue — bars, accents
  PRIMARY_LIGHT:  "DBEAFE",   // Very light blue — row zebra stripe
  PRIMARY_DARK:   "1D4ED8",   // Dark blue — bar stroke
  PRIMARY_MID:    "93C5FD",   // Medium blue — background bars (track)
  MILESTONE:      "F59E0B",   // Amber — milestone diamond
  GATE:           "64748B",   // Slate — gate line
  GATE_LIGHT:     "CBD5E1",   // Light slate — gate label bg

  // Neutrals
  BG:             "FFFFFF",
  HEADER_BG:      "F8FAFC",
  HEADER_TEXT:    "1E3A5F",
  LABEL:          "1E293B",
  SUBLABEL:       "64748B",
  GRID_LINE:      "E2E8F0",
  ROW_ALT:        "F8FAFC",
  WHITE:          "FFFFFF",
};

const FONT = "Segoe UI";

// ─── SLIDE DIMENSIONS (LAYOUT_WIDE = 13.3" × 7.5") ─────────────────────────
const SLIDE_W   = 13.3;
const SLIDE_H   = 7.5;

// ─── LAYOUT ZONES ───────────────────────────────────────────────────────────
const MARGIN_L  = 0.3;   // Left edge of everything
const MARGIN_T  = 0.3;   // Top edge
const LABEL_W   = 2.4;   // Width of the task-name column
const CHART_X   = MARGIN_L + LABEL_W + 0.15; // Where the timeline begins (inches)
const CHART_W   = SLIDE_W - CHART_X - 0.3;   // Width of the timeline area
const HEADER_H  = 0.95;  // Height of the header banner
const FOOT_H    = 0.32;  // Footer strip
const BODY_Y    = MARGIN_T + HEADER_H + 0.12; // Y where rows start
const BODY_H    = SLIDE_H - BODY_Y - FOOT_H - 0.1;
const ROW_GAP   = 0.06;  // Vertical gap between rows

// ─── HELPERS ────────────────────────────────────────────────────────────────
function parseDate(str) { return new Date(str + "T00:00:00Z"); }

function daysBetween(a, b) {
  return (b - a) / (1000 * 60 * 60 * 24);
}

/** Map a date to an X position (inches) within the chart area */
function dateToX(date, startDate, totalDays) {
  const elapsed = daysBetween(startDate, date);
  return CHART_X + (elapsed / totalDays) * CHART_W;
}

/** Format a date as "Jan 1" */
function fmtShort(date) {
  return date.toLocaleDateString("en-US", { month: "short", day: "numeric", timeZone: "UTC" });
}

/** Generate N evenly-spaced tick dates between start and end (inclusive) */
function generateTicks(startDate, endDate, approxCount) {
  const totalDays = daysBetween(startDate, endDate);
  const tickInterval = Math.ceil(totalDays / (approxCount - 1));
  const ticks = [];
  let cur = new Date(startDate);
  while (cur <= endDate) {
    ticks.push(new Date(cur));
    cur = new Date(cur.getTime() + tickInterval * 24 * 60 * 60 * 1000);
  }
  // Always include the end date
  if (ticks[ticks.length - 1].getTime() !== endDate.getTime()) {
    ticks.push(new Date(endDate));
  }
  return ticks;
}

// ─── MAIN ───────────────────────────────────────────────────────────────────
async function generateGantt(jsonInput) {
  // ── Parse input ──
  let data;
  try {
    data = JSON.parse(jsonInput);
  } catch (e) {
    console.error("❌  Invalid JSON input:", e.message);
    process.exit(1);
  }

  const projectName = data.projectName || "Gantt Chart";
  const tasks       = data.tasks  || [];
  const gates       = (data.gates || []).map(parseDate);

  if (!tasks.length) {
    console.error("❌  No tasks provided.");
    process.exit(1);
  }

  // ── Compute timeline bounds ──
  let allDates = [];
  tasks.forEach(t => {
    if (t.type === "milestone") allDates.push(parseDate(t.date));
    else { allDates.push(parseDate(t.start)); allDates.push(parseDate(t.end)); }
  });
  gates.forEach(d => allDates.push(d));

  let minDate = new Date(Math.min(...allDates));
  let maxDate = new Date(Math.max(...allDates));

  // Pad by 5% on each side for breathing room
  const pad = Math.max(2, Math.round(daysBetween(minDate, maxDate) * 0.04));
  minDate = new Date(minDate.getTime() - pad * 24 * 60 * 60 * 1000);
  maxDate = new Date(maxDate.getTime() + pad * 24 * 60 * 60 * 1000);
  const totalDays = daysBetween(minDate, maxDate);

  // Separate regular tasks from milestones (milestones rendered on top of rows)
  const regularTasks  = tasks.filter(t => t.type !== "milestone");
  const milestones    = tasks.filter(t => t.type === "milestone");

  const rowCount  = regularTasks.length;
  const rowH      = Math.min(0.52, (BODY_H - ROW_GAP * (rowCount - 1)) / rowCount);
  const barH      = rowH * 0.48;   // Bar is 48% of row height
  const trackH    = rowH * 0.20;   // Background "track" behind bar

  // ─── Create presentation ──────────────────────────────────────────────────
  const pres   = new pptxgen();
  pres.layout  = "LAYOUT_WIDE"; // 13.3 × 7.5
  pres.title   = projectName;
  pres.author  = "Gantt Generator";

  const slide  = pres.addSlide();
  slide.background = { color: C.BG };

  // ═══════════════════════════════════════════════════════════════════════════
  // HEADER BANNER
  // ═══════════════════════════════════════════════════════════════════════════
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: HEADER_H,
    fill: { color: C.PRIMARY },
    line: { color: C.PRIMARY, width: 0 },
  });

  // Project title
  slide.addText(projectName, {
    x: MARGIN_L, y: 0, w: 6, h: HEADER_H,
    fontFace: FONT, fontSize: 22, bold: true, color: C.WHITE,
    valign: "middle", align: "left", margin: 0,
  });

  // "GANTT CHART" badge on the right
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SLIDE_W - 1.8, y: (HEADER_H - 0.32) / 2, w: 1.5, h: 0.32,
    fill: { color: C.WHITE },
    line: { color: C.WHITE, width: 0 },
    rectRadius: 0.05,
  });
  slide.addText("GANTT CHART", {
    x: SLIDE_W - 1.8, y: (HEADER_H - 0.32) / 2, w: 1.5, h: 0.32,
    fontFace: FONT, fontSize: 7.5, bold: true, color: C.PRIMARY,
    valign: "middle", align: "center", charSpacing: 2, margin: 0,
  });

  // Date range subtitle
  slide.addText(`${fmtShort(minDate)} – ${fmtShort(maxDate)}`, {
    x: 6.5, y: 0, w: SLIDE_W - 6.5 - 2, h: HEADER_H,
    fontFace: FONT, fontSize: 9.5, color: "BFDBFE",
    valign: "middle", align: "right", margin: 0,
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN HEADER STRIP (below banner)
  // ═══════════════════════════════════════════════════════════════════════════
  const colHdrY = MARGIN_T + HEADER_H;
  const colHdrH = 0.30;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: colHdrY, w: SLIDE_W, h: colHdrH,
    fill: { color: C.HEADER_BG },
    line: { color: C.GRID_LINE, width: 0.5 },
  });

  slide.addText("TASK", {
    x: MARGIN_L, y: colHdrY, w: LABEL_W, h: colHdrH,
    fontFace: FONT, fontSize: 7.5, bold: true, color: C.SUBLABEL,
    valign: "middle", align: "left", charSpacing: 1.5, margin: 0,
  });

  slide.addText("TIMELINE", {
    x: CHART_X, y: colHdrY, w: CHART_W, h: colHdrH,
    fontFace: FONT, fontSize: 7.5, bold: true, color: C.SUBLABEL,
    valign: "middle", align: "left", charSpacing: 1.5, margin: 0,
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // TICK MARKS & VERTICAL GRID LINES
  // ═══════════════════════════════════════════════════════════════════════════
  const tickY     = colHdrY + colHdrH;
  const tickH     = 0.28;
  const gridBotY  = tickY + tickH + rowCount * (rowH + ROW_GAP);
  const ticks     = generateTicks(minDate, maxDate, 9);

  ticks.forEach((tick, i) => {
    const tx = dateToX(tick, minDate, totalDays);

    // Vertical grid line (through all rows)
    slide.addShape(pres.shapes.LINE, {
      x: tx, y: tickY + tickH * 0.7,
      w: 0, h: gridBotY - (tickY + tickH * 0.7),
      line: { color: C.GRID_LINE, width: 0.6, dashType: "sysDot" },
    });

    // Tick label — skip if too close to edges
    const labelW = 0.9;
    if (tx - labelW / 2 >= CHART_X - 0.05 && tx + labelW / 2 <= CHART_X + CHART_W + 0.05) {
      slide.addText(fmtShort(tick), {
        x: tx - labelW / 2, y: tickY, w: labelW, h: tickH,
        fontFace: FONT, fontSize: 7.5, color: C.SUBLABEL,
        valign: "middle", align: "center", margin: 0,
      });
    }
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // ROWS — Background + Task label + Bar
  // ═══════════════════════════════════════════════════════════════════════════
  const rowsStartY = tickY + tickH;

  regularTasks.forEach((task, i) => {
    const ry      = rowsStartY + i * (rowH + ROW_GAP);
    const barY    = ry + (rowH - barH) / 2;
    const trackY  = ry + (rowH - trackH) / 2;

    // ── Alternating row background ──
    if (i % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0, y: ry, w: SLIDE_W, h: rowH,
        fill: { color: C.ROW_ALT },
        line: { color: C.GRID_LINE, width: 0.3 },
      });
    }

    // ── Task name ──
    slide.addText(task.name, {
      x: MARGIN_L, y: ry, w: LABEL_W - 0.1, h: rowH,
      fontFace: FONT, fontSize: 9, bold: true, color: C.LABEL,
      valign: "middle", align: "left", margin: 0,
    });

    // ── Start / End dates below the name ──
    const startD = parseDate(task.start);
    const endD   = parseDate(task.end);
    const dur    = Math.round(daysBetween(startD, endD));
    slide.addText(`${fmtShort(startD)} → ${fmtShort(endD)}  (${dur}d)`, {
      x: MARGIN_L, y: ry + rowH * 0.54, w: LABEL_W - 0.1, h: rowH * 0.4,
      fontFace: FONT, fontSize: 7, color: C.SUBLABEL,
      valign: "top", align: "left", margin: 0,
    });

    // ── Background track (full-width "rail") ──
    const xStart = dateToX(startD, minDate, totalDays);
    const xEnd   = dateToX(endD,   minDate, totalDays);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: CHART_X, y: trackY, w: CHART_W, h: trackH,
      fill: { color: C.PRIMARY_LIGHT },
      line: { color: C.GRID_LINE, width: 0.5 },
      rectRadius: 0.04,
    });

    // ── Task bar ──
    const barW = Math.max(0.04, xEnd - xStart);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xStart, y: barY, w: barW, h: barH,
      fill: { color: C.PRIMARY },
      line: { color: C.PRIMARY_DARK, width: 0 },
      rectRadius: 0.06,
    });

    // ── Team members label inside / beside bar ──
    const team = (task.team || []).join(", ");
    if (team) {
      // Try to place label inside the bar; if bar is narrow, place to the right
      const labelFontSz = 7.5;
      const insideX = xStart + 0.07;
      const insideW = barW - 0.12;

      if (insideW > 0.3) {
        slide.addText(team, {
          x: insideX, y: barY, w: insideW, h: barH,
          fontFace: FONT, fontSize: labelFontSz, color: C.WHITE,
          valign: "middle", align: "left", margin: 0,
        });
      } else {
        // Outside the bar (to the right)
        slide.addText(team, {
          x: xEnd + 0.06, y: barY, w: 1.5, h: barH,
          fontFace: FONT, fontSize: labelFontSz, color: C.SUBLABEL,
          valign: "middle", align: "left", margin: 0,
        });
      }
    }
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // GATES — Dashed vertical lines spanning all rows
  // ═══════════════════════════════════════════════════════════════════════════
  const gateLineTop = rowsStartY;
  const gateLineBot = rowsStartY + rowCount * (rowH + ROW_GAP) - ROW_GAP;

  gates.forEach(gateDate => {
    const gx = dateToX(gateDate, minDate, totalDays);
    if (gx < CHART_X || gx > CHART_X + CHART_W) return;

    // Dashed gate line
    slide.addShape(pres.shapes.LINE, {
      x: gx, y: gateLineTop, w: 0, h: gateLineBot - gateLineTop,
      line: { color: C.GATE, width: 1.5, dashType: "dash" },
    });

    // Gate label pill at the top
    const pillW = 0.75;
    const pillH = 0.22;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: gx - pillW / 2, y: gateLineTop - pillH - 0.03,
      w: pillW, h: pillH,
      fill: { color: C.GATE },
      line: { color: C.GATE, width: 0 },
      rectRadius: 0.04,
    });
    slide.addText(`GATE  ${fmtShort(gateDate)}`, {
      x: gx - pillW / 2, y: gateLineTop - pillH - 0.03,
      w: pillW, h: pillH,
      fontFace: FONT, fontSize: 6, bold: true, color: C.WHITE,
      valign: "middle", align: "center", charSpacing: 0.5, margin: 0,
    });
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // MILESTONES — Diamond shapes on the timeline
  // ═══════════════════════════════════════════════════════════════════════════
  // We need to render milestones at the correct Y.
  // Strategy: find the row index of the task *before* the milestone to align it.
  // If none found, place at mid-chart.
  const DIAMOND_SIZE = 0.22;

  milestones.forEach(ms => {
    const msDate = parseDate(ms.date);
    const mx     = dateToX(msDate, minDate, totalDays);
    if (mx < CHART_X || mx > CHART_X + CHART_W) return;

    // Find the last regular task that ends on or before the milestone date
    let targetRowIdx = regularTasks.length - 1;
    for (let i = 0; i < regularTasks.length; i++) {
      const tEnd = parseDate(regularTasks[i].end);
      if (tEnd <= msDate) targetRowIdx = i;
    }

    const ry = rowsStartY + targetRowIdx * (rowH + ROW_GAP);
    const my = ry + rowH / 2; // Vertical center of that row

    // Diamond = rotated square
    slide.addShape(pres.shapes.RECTANGLE, {
      x: mx - DIAMOND_SIZE / 2, y: my - DIAMOND_SIZE / 2,
      w: DIAMOND_SIZE, h: DIAMOND_SIZE,
      fill: { color: C.MILESTONE },
      line: { color: "D97706", width: 1 },
      rotate: 45,
    });

    // Milestone label (above diamond)
    slide.addText(ms.name, {
      x: mx - 0.9, y: my - DIAMOND_SIZE / 2 - 0.22,
      w: 1.8, h: 0.22,
      fontFace: FONT, fontSize: 7.5, bold: true, color: "92400E",
      valign: "middle", align: "center", margin: 0,
    });

    // Date label (below diamond)
    slide.addText(fmtShort(msDate), {
      x: mx - 0.6, y: my + DIAMOND_SIZE / 2 + 0.02,
      w: 1.2, h: 0.18,
      fontFace: FONT, fontSize: 7, color: C.SUBLABEL,
      valign: "top", align: "center", margin: 0,
    });
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // LEGEND
  // ═══════════════════════════════════════════════════════════════════════════
  const legendY   = SLIDE_H - FOOT_H + 0.05;
  const legendItems = [
    { label: "Task", shape: "bar",      color: C.PRIMARY },
    { label: "Milestone", shape: "diamond", color: C.MILESTONE },
    { label: "Gate", shape: "dash",     color: C.GATE },
  ];

  let lx = CHART_X;
  legendItems.forEach(item => {
    const iconSize = 0.13;
    const iconY    = legendY + (FOOT_H - iconSize) / 2;

    if (item.shape === "bar") {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: lx, y: iconY, w: 0.28, h: iconSize,
        fill: { color: item.color },
        line: { color: item.color, width: 0 },
        rectRadius: 0.03,
      });
      lx += 0.32;
    } else if (item.shape === "diamond") {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: lx + 0.05, y: iconY, w: iconSize, h: iconSize,
        fill: { color: item.color },
        line: { color: "D97706", width: 0.5 },
        rotate: 45,
      });
      lx += 0.32;
    } else if (item.shape === "dash") {
      slide.addShape(pres.shapes.LINE, {
        x: lx, y: iconY + iconSize / 2, w: 0.28, h: 0,
        line: { color: item.color, width: 1.5, dashType: "dash" },
      });
      lx += 0.32;
    }

    slide.addText(item.label, {
      x: lx, y: legendY, w: 0.7, h: FOOT_H,
      fontFace: FONT, fontSize: 8, color: C.SUBLABEL,
      valign: "middle", align: "left", margin: 0,
    });
    lx += 0.78;
  });

  // Footer line
  slide.addShape(pres.shapes.LINE, {
    x: 0, y: SLIDE_H - FOOT_H - 0.01, w: SLIDE_W, h: 0,
    line: { color: C.GRID_LINE, width: 0.5 },
  });

  // Generated-by note (right side of footer)
  slide.addText(`Generated ${new Date().toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"})}`, {
    x: SLIDE_W - 2.2, y: legendY, w: 2, h: FOOT_H,
    fontFace: FONT, fontSize: 7.5, color: C.SUBLABEL,
    valign: "middle", align: "right", margin: 0,
  });

  // ─── Write file ───────────────────────────────────────────────────────────
  const outFile = "gantt_chart.pptx";
  await pres.writeFile({ fileName: outFile });
  console.log(`✅  Gantt chart saved → ${outFile}`);
}

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────
const arg = process.argv[2];
if (!arg) {
  console.error("Usage: node gantt_chart.js '<JSON_STRING>'");
  console.error("Run with --help for an example command.");
  process.exit(1);
}
if (arg === "--help" || arg === "-h") {
  console.log(`
Example:
  node gantt_chart.js '{"projectName":"Project Alpha","tasks":[{"name":"Design Phase","start":"2024-01-01","end":"2024-01-15","team":["Alice","Bob"]},{"name":"Milestone 1","date":"2024-01-16","type":"milestone"}],"gates":["2024-01-20"]}'
`);
  process.exit(0);
}

generateGantt(arg).catch(err => {
  console.error("❌  Fatal error:", err.message);
  process.exit(1);
});
