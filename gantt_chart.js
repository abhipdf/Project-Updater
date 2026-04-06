/**
 * ============================================================
 *  Gantt Chart Generator — PptxGenJS  (Split Table + Timeline)
 *  Generates a professional Gantt Chart as an editable .pptx
 * ============================================================
 *
 * USAGE:
 *   node gantt_chart.js '<JSON_STRING>'
 *
 * EXAMPLE:
 *   node gantt_chart.js '{"projectName":"Project Alpha","tasks":[{"name":"Discovery & Research","start":"2024-01-01","end":"2024-01-10","team":["Alice","Bob"]},{"name":"Design Phase","start":"2024-01-08","end":"2024-01-22","team":["Carol","Dave"]},{"name":"Prototype Build","start":"2024-01-20","end":"2024-02-05","team":["Eve","Frank"]},{"name":"User Testing","start":"2024-02-03","end":"2024-02-14","team":["Alice","Carol"]},{"name":"Iteration & Polish","start":"2024-02-12","end":"2024-02-26","team":["Dave","Eve"]}],"milestones":[{"name":"M1: Design Sign-off","date":"2024-01-25"},{"name":"M2: Beta Release","date":"2024-02-28"}],"gates":["2024-01-15","2024-02-10"]}'
 *
 * JSON SCHEMA:
 * {
 *   "projectName": "string",
 *   "tasks": [
 *     { "name": "string", "start": "YYYY-MM-DD", "end": "YYYY-MM-DD", "team": ["string"] }
 *   ],
 *   "milestones": [
 *     { "name": "string", "date": "YYYY-MM-DD" }
 *   ],
 *   "gates": ["YYYY-MM-DD"]
 * }
 *
 * OUTPUT: gantt_chart.pptx (in current working directory)
 */

"use strict";

const pptxgen = require("pptxgenjs");

// ─── DESIGN TOKENS ─────────────────────────────────────────────────────────
const C = {
  // Primary palette
  BLUE:           "2563EB",   // Gantt bars, milestone diamonds, header accents
  BLUE_DARK:      "1D4ED8",   // Bar border / depth
  BLUE_LIGHT:     "DBEAFE",   // Alternating row tint
  BLUE_MID:       "93C5FD",   // Sub-header background

  // Milestone / gate
  MILESTONE_FILL: "DC2626",   // Diamond fill — blue to match bars
  MILESTONE_LINE: "1D4ED8",   // Diamond border
  GATE_LINE:      "94A3B8",   // Dashed gate line — slate

  // Table
  TBL_HEADER_BG:  "1E3A5F",  // Deep navy header background
  TBL_HEADER_FG:  "FFFFFF",  // Header text
  TBL_ROW_ODD:    "FFFFFF",  // Normal row
  TBL_ROW_EVEN:   "F0F5FF",  // Alternating tint
  TBL_BORDER:     "CBD5E1",  // Cell border colour
  TBL_TEXT:       "1E293B",  // Body text
  TBL_SUB:        "64748B",  // Secondary / duration text

  // Timeline header
  TL_MONTH_BG:    "1E3A5F",  // Month band background (matches table header)
  TL_MONTH_FG:    "FFFFFF",  // Month band text
  TL_DAY_BG:      "EFF6FF",  // Day sub-header background
  TL_DAY_FG:      "64748B",  // Day number text
  TL_GRID:        "E2E8F0",  // Vertical grid lines

  // Chrome
  SLIDE_BG:       "FFFFFF",
  TITLE_BG:       "1E3A5F",
  TITLE_FG:       "FFFFFF",
  FOOTER_BG:      "F8FAFC",
  FOOTER_TEXT:    "94A3B8",
};

const FONT = "Calibri";

// ─── SLIDE GEOMETRY ─────────────────────────────────────────────────────────
const SW = 13.3;   // Slide width  (LAYOUT_WIDE)
const SH = 7.5;    // Slide height

// Zones (inches)
const TITLE_H  = 0.55;   // Top title bar height
const FOOTER_H = 0.28;   // Bottom footer height
const MARGIN   = 0.18;   // Outer horizontal margin

// Table section (left side)
const TBL_X          = MARGIN;
const TBL_COL_ID_W   = 0.32;   // "#" column width
const TBL_COL_NAME_W = 2.10;   // Task name column width
const TBL_COL_DUR_W  = 0.60;   // Duration column width
const TBL_W          = TBL_COL_ID_W + TBL_COL_NAME_W + TBL_COL_DUR_W;

// Timeline section (right side)
const TL_X = TBL_X + TBL_W + 0.08;   // Start of timeline area
const TL_W = SW - TL_X - MARGIN;      // Width of timeline area

// Vertical row geometry
const HDR_ROW_H  = 0.30;   // Month header row height
const DAY_ROW_H  = 0.22;   // Day sub-header row height
const DATA_ROW_H = 0.42;   // Data row height

// Y positions
const CHART_Y = TITLE_H + 0.1;                    // Where the header rows start
const DATA_Y  = CHART_Y + HDR_ROW_H + DAY_ROW_H; // Where data rows start
const BAR_PAD = 0.07;                              // Vertical padding inside a row for the bar
const BAR_H   = DATA_ROW_H - BAR_PAD * 2;         // Gantt bar height

// ─── DATE HELPERS ───────────────────────────────────────────────────────────
function parseDate(str) {
  return new Date(str + "T00:00:00Z");
}

function daysBetween(a, b) {
  return (b - a) / 864e5;
}

function fmtMonth(date) {
  return date.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" });
}

function fmtShort(date) {
  return date.toLocaleDateString("en-US", { month: "short", day: "numeric", timeZone: "UTC" });
}

function durLabel(start, end) {
  const d = Math.round(daysBetween(parseDate(start), parseDate(end)));
  return `${d}d`;
}

// ─── COORDINATE MAPPING ─────────────────────────────────────────────────────
/**
 * Maps a Date object to an X-coordinate (inches) within the timeline area.
 * @param {Date}   date       - The date to map
 * @param {Date}   tlStart    - Timeline start date
 * @param {number} totalDays  - Total span of the timeline in days
 * @returns {number} X coordinate in inches
 */
function dateToX(date, tlStart, totalDays) {
  const elapsed = daysBetween(tlStart, date);
  return TL_X + (elapsed / totalDays) * TL_W;
}

// ─── MAIN ────────────────────────────────────────────────────────────────────
async function generateGantt(jsonInput) {

  // ── 1. Parse & validate input ────────────────────────────────────────────
  let data;
  try {
    data = JSON.parse(jsonInput);
  } catch (e) {
    console.error("❌  Invalid JSON:", e.message);
    process.exit(1);
  }

  const projectName = data.projectName || "Gantt Chart";
  const rawTasks    = data.tasks       || [];
  const gates       = (data.gates || []).map(parseDate);

  // Collect milestones from both the dedicated array AND task entries with type:"milestone"
  const milestones = [...(data.milestones || [])];
  rawTasks.forEach(t => {
    if (t.type === "milestone" && t.date) {
      milestones.push({ name: t.name, date: t.date });
    }
  });

  // Regular tasks only (exclude milestone-typed entries)
  const regularTasks = rawTasks.filter(t => t.type !== "milestone");

  if (!regularTasks.length) {
    console.error("❌  No regular tasks provided.");
    process.exit(1);
  }

  // ── 2. Compute timeline bounds ───────────────────────────────────────────
  const allDates = [];
  regularTasks.forEach(t => {
    if (t.start) allDates.push(parseDate(t.start));
    if (t.end)   allDates.push(parseDate(t.end));
  });
  milestones.forEach(m => allDates.push(parseDate(m.date)));
  gates.forEach(d => allDates.push(d));

  const rawMin = new Date(Math.min(...allDates));
  const rawMax = new Date(Math.max(...allDates));

  // Snap start to 1st of the first month, end to last day of the last month
  const tlStart   = new Date(Date.UTC(rawMin.getUTCFullYear(), rawMin.getUTCMonth(), 1));
  const tlEnd     = new Date(Date.UTC(rawMax.getUTCFullYear(), rawMax.getUTCMonth() + 1, 0));
  const totalDays = daysBetween(tlStart, tlEnd);

  // ── 3. Build month segments ──────────────────────────────────────────────
  const monthSegs = [];
  let cur = new Date(tlStart);
  while (cur < tlEnd) {
    const segStart  = new Date(cur);
    const nextMonth = new Date(Date.UTC(cur.getUTCFullYear(), cur.getUTCMonth() + 1, 1));
    const segEnd    = nextMonth > tlEnd ? new Date(tlEnd) : new Date(nextMonth);
    monthSegs.push({
      label:  fmtMonth(segStart),
      xStart: dateToX(segStart, tlStart, totalDays),
      xEnd:   dateToX(segEnd,   tlStart, totalDays),
    });
    cur = nextMonth;
  }

  // ── 4. Build day tick marks (every 3 days) ───────────────────────────────
  const DAY_STEP = 3;
  const dayTicks = [];
  let dayPtr = new Date(tlStart);
  while (dayPtr <= tlEnd) {
    const dom = dayPtr.getUTCDate();
    if (dom === 1 || dom % DAY_STEP === 0) {
      dayTicks.push({
        label: String(dom),
        x:     dateToX(dayPtr, tlStart, totalDays),
      });
    }
    dayPtr = new Date(dayPtr.getTime() + 864e5);
  }

  // ── 5. Create presentation ───────────────────────────────────────────────
  const pres  = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title  = projectName;

  const slide = pres.addSlide();
  slide.background = { color: C.SLIDE_BG };

  // ═════════════════════════════════════════════════════════════════════════
  // TITLE BAR
  // ═════════════════════════════════════════════════════════════════════════
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SW, h: TITLE_H,
    fill: { color: C.TITLE_BG },
    line: { color: C.TITLE_BG, width: 0 },
  });

  // Left accent stripe
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.06, h: TITLE_H,
    fill: { color: C.BLUE },
    line: { color: C.BLUE, width: 0 },
  });

  slide.addText(projectName, {
    x: 0.18, y: 0, w: 7.5, h: TITLE_H,
    fontFace: FONT, fontSize: 16, bold: true, color: C.TITLE_FG,
    valign: "middle", align: "left", margin: 0,
  });

  // "GANTT CHART" pill badge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: SW - 1.95, y: (TITLE_H - 0.26) / 2, w: 1.75, h: 0.26,
    fill: { color: C.BLUE },
    line: { color: C.BLUE, width: 0 },
  });
  slide.addText("GANTT CHART", {
    x: SW - 1.95, y: (TITLE_H - 0.26) / 2, w: 1.75, h: 0.26,
    fontFace: FONT, fontSize: 7, bold: true, color: C.TITLE_FG,
    valign: "middle", align: "center", charSpacing: 1.5, margin: 0,
  });

  // Date range
  slide.addText(`${fmtShort(tlStart)}  –  ${fmtShort(tlEnd)}`, {
    x: SW - 4.0, y: 0, w: 1.9, h: TITLE_H,
    fontFace: FONT, fontSize: 8.5, color: C.BLUE_MID,
    valign: "middle", align: "right", margin: 0,
  });

  // ═════════════════════════════════════════════════════════════════════════
  // TABLE HEADER (spans both month-row and day-row heights)
  // ═════════════════════════════════════════════════════════════════════════
  const tblHdrH = HDR_ROW_H + DAY_ROW_H;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: TBL_X, y: CHART_Y, w: TBL_W, h: tblHdrH,
    fill: { color: C.TBL_HEADER_BG },
    line: { color: C.TBL_HEADER_BG, width: 0 },
  });

  // "#" header
  slide.addText("#", {
    x: TBL_X, y: CHART_Y, w: TBL_COL_ID_W, h: tblHdrH,
    fontFace: FONT, fontSize: 8.5, bold: true, color: C.TBL_HEADER_FG,
    valign: "middle", align: "center", margin: 0,
  });

  // "Task Name" header
  slide.addText("Task Name", {
    x: TBL_X + TBL_COL_ID_W, y: CHART_Y,
    w: TBL_COL_NAME_W, h: tblHdrH,
    fontFace: FONT, fontSize: 8.5, bold: true, color: C.TBL_HEADER_FG,
    valign: "middle", align: "left", margin: [0, 0, 0, 6],
  });

  // "Dur." header
  slide.addText("Dur.", {
    x: TBL_X + TBL_COL_ID_W + TBL_COL_NAME_W, y: CHART_Y,
    w: TBL_COL_DUR_W, h: tblHdrH,
    fontFace: FONT, fontSize: 8.5, bold: true, color: C.TBL_HEADER_FG,
    valign: "middle", align: "center", margin: 0,
  });

  // Table header column dividers
  [TBL_COL_ID_W, TBL_COL_ID_W + TBL_COL_NAME_W].forEach(offset => {
    slide.addShape(pres.shapes.LINE, {
      x: TBL_X + offset, y: CHART_Y, w: 0, h: tblHdrH,
      line: { color: "2D5090", width: 0.5 },
    });
  });

  // ═════════════════════════════════════════════════════════════════════════
  // TIMELINE MONTH HEADER ROW
  // ═════════════════════════════════════════════════════════════════════════
  monthSegs.forEach(seg => {
    const segW = seg.xEnd - seg.xStart;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: seg.xStart, y: CHART_Y, w: segW, h: HDR_ROW_H,
      fill: { color: C.TL_MONTH_BG },
      line: { color: "2D5090", width: 0.5 },
    });
    slide.addText(seg.label, {
      x: seg.xStart + 0.08, y: CHART_Y, w: segW - 0.1, h: HDR_ROW_H,
      fontFace: FONT, fontSize: 8.5, bold: true, color: C.TL_MONTH_FG,
      valign: "middle", align: "left", margin: 0,
    });
  });

  // ─── DAY SUB-HEADER ROW ───────────────────────────────────────────────────
  const dayHdrY = CHART_Y + HDR_ROW_H;

  slide.addShape(pres.shapes.RECTANGLE, {
    x: TL_X, y: dayHdrY, w: TL_W, h: DAY_ROW_H,
    fill: { color: C.TL_DAY_BG },
    line: { color: C.TBL_BORDER, width: 0.4 },
  });

  dayTicks.forEach(tick => {
    if (tick.x < TL_X - 0.01 || tick.x > TL_X + TL_W + 0.01) return;
    slide.addText(tick.label, {
      x: tick.x - 0.22, y: dayHdrY, w: 0.44, h: DAY_ROW_H,
      fontFace: FONT, fontSize: 6.5, color: C.TL_DAY_FG,
      valign: "middle", align: "center", margin: 0,
    });
  });

  // ═════════════════════════════════════════════════════════════════════════
  // DATA ROWS  — table cells + gantt bars, perfectly synchronised by row Y
  // ═════════════════════════════════════════════════════════════════════════
  const chartBot = DATA_Y + regularTasks.length * DATA_ROW_H;

  regularTasks.forEach((task, idx) => {
    const rowY  = DATA_Y + idx * DATA_ROW_H;
    const rowBg = idx % 2 === 1 ? C.TBL_ROW_EVEN : C.TBL_ROW_ODD;

    // ── Table row background ──
    slide.addShape(pres.shapes.RECTANGLE, {
      x: TBL_X, y: rowY, w: TBL_W, h: DATA_ROW_H,
      fill: { color: rowBg },
      line: { color: C.TBL_BORDER, width: 0.4 },
    });

    // ── Timeline row background (same alternating shade, keeps grid visible) ──
    slide.addShape(pres.shapes.RECTANGLE, {
      x: TL_X, y: rowY, w: TL_W, h: DATA_ROW_H,
      fill: { color: rowBg },
      line: { color: C.TBL_BORDER, width: 0.3 },
    });

    // ── Column dividers in table ──
    [TBL_COL_ID_W, TBL_COL_ID_W + TBL_COL_NAME_W].forEach(offset => {
      slide.addShape(pres.shapes.LINE, {
        x: TBL_X + offset, y: rowY, w: 0, h: DATA_ROW_H,
        line: { color: C.TBL_BORDER, width: 0.4 },
      });
    });

    // ── Cell: # ──
    slide.addText(String(idx + 1), {
      x: TBL_X, y: rowY, w: TBL_COL_ID_W, h: DATA_ROW_H,
      fontFace: FONT, fontSize: 8, color: C.TBL_SUB,
      valign: "middle", align: "center", margin: 0,
    });

    // ── Cell: Task name ──
    slide.addText(task.name || "", {
      x: TBL_X + TBL_COL_ID_W + 0.07, y: rowY,
      w: TBL_COL_NAME_W - 0.09, h: DATA_ROW_H,
      fontFace: FONT, fontSize: 8.5, bold: true, color: C.TBL_TEXT,
      valign: "middle", align: "left", margin: 0,
    });

    // ── Cell: Duration ──
    const dur = task.start && task.end ? durLabel(task.start, task.end) : "";
    slide.addText(dur, {
      x: TBL_X + TBL_COL_ID_W + TBL_COL_NAME_W, y: rowY,
      w: TBL_COL_DUR_W, h: DATA_ROW_H,
      fontFace: FONT, fontSize: 8, color: C.TBL_SUB,
      valign: "middle", align: "center", margin: 0,
    });

    // ── GANTT BAR (vertically centred in row) ──────────────────────────────
    if (task.start && task.end) {
      const barXStart = dateToX(parseDate(task.start), tlStart, totalDays);
      const barXEnd   = dateToX(parseDate(task.end),   tlStart, totalDays);
      const barW      = Math.max(0.05, barXEnd - barXStart);
      const barY      = rowY + BAR_PAD;   // Perfectly synchronised with row Y

      // Bar shape — ROUNDED_RECTANGLE for polished look
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: barXStart, y: barY, w: barW, h: BAR_H,
        fill: { color: C.BLUE },
        line: { color: C.BLUE_DARK, width: 0 },
        rectRadius: 0.04,
      });

      // Team members label — immediately to the right of the bar
      const team = (task.team || []).join(", ");
      if (team) {
        const rightEdge = TL_X + TL_W - 0.06;   // Safe right edge
        const labelX    = Math.min(barXEnd + 0.07, rightEdge - 0.1);
        const spaceW    = rightEdge - labelX;

        if (barXEnd < rightEdge - 0.3 && spaceW > 0.3) {
          // Enough room to the right of the bar — place outside
          slide.addText(team, {
            x: labelX, y: barY, w: spaceW, h: BAR_H,
            fontFace: FONT, fontSize: 7.5, color: C.TBL_SUB,
            valign: "middle", align: "left", margin: 0,
          });
        } else if (barW > 0.6) {
          // Bar is wide enough — place label inside (white text)
          slide.addText(team, {
            x: barXStart + 0.07, y: barY, w: barW - 0.1, h: BAR_H,
            fontFace: FONT, fontSize: 7.5, color: "FFFFFF",
            valign: "middle", align: "left", margin: 0,
          });
        }
      }
    }
  });

  // ─── VERTICAL DAY GRID LINES (drawn after rows so they layer on top) ─────
  dayTicks.forEach(tick => {
    if (tick.x < TL_X - 0.01 || tick.x > TL_X + TL_W + 0.01) return;
    slide.addShape(pres.shapes.LINE, {
      x: tick.x, y: DATA_Y, w: 0, h: chartBot - DATA_Y,
      line: { color: C.TL_GRID, width: 0.4 },
    });
  });

  // ─── MONTH BOUNDARY BOLD GRID LINES ──────────────────────────────────────
  monthSegs.forEach((seg, i) => {
    if (i === 0) return;
    slide.addShape(pres.shapes.LINE, {
      x: seg.xStart, y: CHART_Y, w: 0, h: chartBot - CHART_Y,
      line: { color: "94A3B8", width: 0.8 },
    });
  });

  // ═════════════════════════════════════════════════════════════════════════
  // GATES — dashed vertical lines across the full data area
  // ═════════════════════════════════════════════════════════════════════════
  gates.forEach(gateDate => {
    const gx = dateToX(gateDate, tlStart, totalDays);
    if (gx < TL_X - 0.01 || gx > TL_X + TL_W + 0.01) return;

    // Dashed line through data rows
    slide.addShape(pres.shapes.LINE, {
      x: gx, y: DATA_Y, w: 0, h: chartBot - DATA_Y,
      line: { color: C.GATE_LINE, width: 1.5, dashType: "dash" },
    });

    // Label pill in the day sub-header zone
    const pillW = 0.72, pillH = 0.17;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: gx - pillW / 2, y: dayHdrY + (DAY_ROW_H - pillH) / 2,
      w: pillW, h: pillH,
      fill: { color: C.GATE_LINE },
      line: { color: C.GATE_LINE, width: 0 },
    });
    slide.addText(`GATE ${fmtShort(gateDate)}`, {
      x: gx - pillW / 2, y: dayHdrY + (DAY_ROW_H - pillH) / 2,
      w: pillW, h: pillH,
      fontFace: FONT, fontSize: 5.5, bold: true, color: "FFFFFF",
      valign: "middle", align: "center", margin: 0,
    });
  });

  // ═════════════════════════════════════════════════════════════════════════
  // MILESTONES — blue diamond shapes pinned to the timeline rows
  // ═════════════════════════════════════════════════════════════════════════
  const DSIZE = 0.20;

  milestones.forEach(ms => {
    const msDate = parseDate(ms.date);
    const mx     = dateToX(msDate, tlStart, totalDays);
    if (mx < TL_X - 0.01 || mx > TL_X + TL_W + 0.01) return;

    // Pin to the last task whose end date is on or before the milestone date
    let pinRow = regularTasks.length - 1;
    for (let i = 0; i < regularTasks.length; i++) {
      if (regularTasks[i].end && parseDate(regularTasks[i].end) <= msDate) {
        pinRow = i;
      }
    }
    const rowY = DATA_Y + pinRow * DATA_ROW_H;
    const my   = rowY + DATA_ROW_H / 2;   // Vertical centre of pinned row

    // Diamond = 45° rotated square
    slide.addShape(pres.shapes.RECTANGLE, {
      x: mx - DSIZE / 2, y: my - DSIZE / 2,
      w: DSIZE, h: DSIZE,
      fill: { color: C.MILESTONE_FILL },
      line: { color: C.MILESTONE_LINE, width: 1 },
      rotate: 45,
    });

    // Label above diamond — shift left if too close to the right edge
    const labelW    = 1.76;
    const rawLabelX = mx - labelW / 2;
    const safeLabelX = Math.min(rawLabelX, TL_X + TL_W - labelW - 0.04);
    slide.addText(ms.name || "", {
      x: safeLabelX, y: my - DSIZE / 2 - 0.22,
      w: labelW, h: 0.21,
      fontFace: FONT, fontSize: 6.5, bold: true, color: C.BLUE_DARK,
      valign: "middle", align: "center", margin: 0,
    });

    // Date label below diamond
    slide.addText(fmtShort(msDate), {
      x: mx - 0.5, y: my + DSIZE / 2 + 0.02,
      w: 1.0, h: 0.17,
      fontFace: FONT, fontSize: 6.5, color: C.TBL_SUB,
      valign: "top", align: "center", margin: 0,
    });
  });

  // ═════════════════════════════════════════════════════════════════════════
  // CHART BORDERS — crisp outer frames for both sections
  // ═════════════════════════════════════════════════════════════════════════
  const fullH = chartBot - CHART_Y;

  // Table section
  slide.addShape(pres.shapes.RECTANGLE, {
    x: TBL_X, y: CHART_Y, w: TBL_W, h: fullH,
    fill: { type: "none" },
    line: { color: "94A3B8", width: 0.8 },
  });

  // Timeline section
  slide.addShape(pres.shapes.RECTANGLE, {
    x: TL_X, y: CHART_Y, w: TL_W, h: fullH,
    fill: { type: "none" },
    line: { color: "94A3B8", width: 0.8 },
  });

  // ═════════════════════════════════════════════════════════════════════════
  // FOOTER
  // ═════════════════════════════════════════════════════════════════════════
  const footerY = SH - FOOTER_H;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: footerY, w: SW, h: FOOTER_H,
    fill: { color: C.FOOTER_BG },
    line: { color: C.TBL_BORDER, width: 0.5 },
  });

  // Legend
  const legendItems = [
    { label: "Task",      shape: "bar",     color: C.BLUE },
    { label: "Milestone", shape: "diamond", color: C.MILESTONE_FILL },
    { label: "Gate",      shape: "dash",    color: C.GATE_LINE },
  ];
  let lx = TL_X;
  legendItems.forEach(item => {
    const iconY  = footerY + (FOOTER_H - 0.11) / 2;
    const iconSz = 0.11;

    if (item.shape === "bar") {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: lx, y: iconY, w: 0.26, h: iconSz,
        fill: { color: item.color },
        line: { color: item.color, width: 0 },
        rectRadius: 0.03,
      });
    } else if (item.shape === "diamond") {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: lx + 0.04, y: iconY, w: iconSz, h: iconSz,
        fill: { color: item.color },
        line: { color: C.MILESTONE_LINE, width: 0.5 },
        rotate: 45,
      });
    } else if (item.shape === "dash") {
      slide.addShape(pres.shapes.LINE, {
        x: lx, y: iconY + iconSz / 2, w: 0.26, h: 0,
        line: { color: item.color, width: 1.5, dashType: "dash" },
      });
    }

    slide.addText(item.label, {
      x: lx + 0.30, y: footerY, w: 0.80, h: FOOTER_H,
      fontFace: FONT, fontSize: 7.5, color: C.FOOTER_TEXT,
      valign: "middle", align: "left", margin: 0,
    });
    lx += 1.15;
  });

  // Generated timestamp
  const ts = new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
  slide.addText(`Generated ${ts}`, {
    x: SW - 2.0, y: footerY, w: 1.82, h: FOOTER_H,
    fontFace: FONT, fontSize: 7.5, color: C.FOOTER_TEXT,
    valign: "middle", align: "right", margin: 0,
  });

  // ─── Write file ──────────────────────────────────────────────────────────
  const outFile = "gantt_chart.pptx";
  await pres.writeFile({ fileName: outFile });
  console.log(`✅  Gantt chart saved → ${outFile}`);
}

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────
const arg = process.argv[2];
if (!arg || arg === "--help" || arg === "-h") {
  console.log(`
Usage:  node gantt_chart.js '<JSON_STRING>'

Example:
  node gantt_chart.js '{"projectName":"Project Alpha","tasks":[{"name":"Design Phase","start":"2024-01-01","end":"2024-01-15","team":["Alice","Bob"]},{"name":"Dev Sprint","start":"2024-01-13","end":"2024-02-09","team":["Carol"]}],"milestones":[{"name":"Launch","date":"2024-02-10"}],"gates":["2024-01-20"]}'
`);
  process.exit(arg ? 0 : 1);
}

generateGantt(arg).catch(err => {
  console.error("❌  Fatal error:", err.message);
  process.exit(1);
});