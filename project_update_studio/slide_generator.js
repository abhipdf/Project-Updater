#!/usr/bin/env node

/**
 * Slide Generator for Project Update Studio — Premium Edition
 * Generates a single PowerPoint slide with project weekly update data
 *
 * Usage: node slide_generator.js <path_to_json_file>
 */

const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// ─── COLOR PALETTE ────────────────────────────────────────────────────────────
const C = {
  navy:         '002855',   // deep navy — left banner + footer
  ocean:        '0057A8',   // Philips ocean blue — headings, accents
  sky:          '1A7CC1',   // mid blue — sub-accents
  ice:          'D6E8F7',   // very light blue — card backgrounds
  frost:        'EDF4FB',   // near-white blue — alternating rows
  white:        'FFFFFF',
  charcoal:     '1C2B3A',   // body text
  slate:        '4A6075',   // secondary text
  silver:       'B0BEC5',   // dividers, placeholders
  green:        '00875A',   // RAG green
  amber:        'FF8B00',   // RAG amber
  red:          'D9001B',   // RAG red
  tag_green_bg: 'E3F5EE',
  tag_amber_bg: 'FFF3CD',
  tag_red_bg:   'FDECEA',
};

const FONT = 'Calibri';

// Slide dimensions (16:9, inches)
const W = 13.33;
const H = 7.5;

// ─── HELPERS ──────────────────────────────────────────────────────────────────

function ragColor(status) {
  const s = (status || '').toLowerCase();
  if (s === 'red')   return { dot: C.red,   bg: C.tag_red_bg,   label: 'Blocked' };
  if (s === 'amber') return { dot: C.amber, bg: C.tag_amber_bg, label: 'At Risk' };
  return               { dot: C.green, bg: C.tag_green_bg, label: 'On Track' };
}

function label(key, lang) {
  const L = {
    en: {
      this_week:    "Completed This Week",
      next_steps:   "Next Steps",
      team:         "Team",
      decisions:    "Management Decisions",
      kpis:         "KPIs & Metrics",
      risks:        "Risks & Blockers",
      no_data:      "No updates recorded",
      no_decisions: "No decisions required",
      no_risks:     "No blockers identified",
      confidential: "Confidential · Internal Use Only",
      weekly_update:"Weekly Status Update",
      owner:        "Owner",
      urgency_urgent:   "Urgent",
      urgency_week:     "This Week",
      urgency_possible: "When Possible",
    },
    de: {
      this_week:    "Diese Woche abgeschlossen",
      next_steps:   "Nächste Schritte",
      team:         "Team",
      decisions:    "Managemententscheidungen",
      kpis:         "KPIs & Kennzahlen",
      risks:        "Risiken & Blocker",
      no_data:      "Keine Einträge",
      no_decisions: "Keine Entscheidungen erforderlich",
      no_risks:     "Keine Blocker identifiziert",
      confidential: "Vertraulich · Nur für interne Verwendung",
      weekly_update:"Wöchentliches Status-Update",
      owner:        "Verantwortlich",
      urgency_urgent:   "Dringend",
      urgency_week:     "Diese Woche",
      urgency_possible: "Bei Gelegenheit",
    },
  };
  const d = L[lang === 'de' ? 'de' : 'en'];
  return d[key] || key;
}

function safeJson(input, def = []) {
  if (!input) return def;
  if (Array.isArray(input) || (typeof input === 'object')) return input;
  try { return JSON.parse(input); } catch { return def; }
}

function trunc(str, n) {
  if (!str) return '';
  return str.length > n ? str.slice(0, n) + '…' : str;
}

function initials(name) {
  return (name || '?').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
}

function urgencyKey(u) {
  if (!u) return 'urgency_possible';
  const s = u.toLowerCase();
  if (s === 'urgent') return 'urgency_urgent';
  if (s === 'this_week') return 'urgency_week';
  return 'urgency_possible';
}

function urgencyColor(u) {
  const s = (u || '').toLowerCase();
  if (s === 'urgent')    return { fg: C.red,   bg: C.tag_red_bg };
  if (s === 'this_week') return { fg: C.amber, bg: C.tag_amber_bg };
  return                        { fg: C.green, bg: C.tag_green_bg };
}

// ─── SECTION CARD ─────────────────────────────────────────────────────────────
// Draws a card with a coloured top-border accent, title, and returns the
// y-position where content should start.
function addCard(slide, prs, { x, y, w, h, title, accentColor }) {
  // Card background
  slide.addShape(prs.ShapeType.rect, {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: 'DEE6EF', width: 0.75, type: 'solid' },
  });
  // Top accent bar
  slide.addShape(prs.ShapeType.rect, {
    x, y, w, h: 0.05,
    fill: { color: accentColor || C.ocean },
    line: { type: 'none' },
  });
  // Title
  slide.addText(title, {
    x: x + 0.15, y: y + 0.08,
    w: w - 0.3, h: 0.28,
    fontFace: FONT, fontSize: 9, bold: true,
    color: accentColor || C.ocean,
    align: 'left', valign: 'middle',
    charSpacing: 0.5,
  });
  // Title underline — thin rule
  slide.addShape(prs.ShapeType.rect, {
    x: x + 0.15, y: y + 0.38,
    w: w - 0.3, h: 0.008,
    fill: { color: 'DEE6EF' },
    line: { type: 'none' },
  });
  return y + 0.42; // content start Y
}

// ─── BULLET ROWS ──────────────────────────────────────────────────────────────
function addBulletRows(slide, prs, { x, y, w, items, maxItems = 5, fontSize = 8.5, rowH = 0.26 }) {
  const shown = items.slice(0, maxItems);
  shown.forEach((txt, i) => {
    const bg = i % 2 === 1 ? C.frost : C.white;
    slide.addShape(prs.ShapeType.rect, {
      x, y: y + i * rowH, w, h: rowH,
      fill: { color: bg }, line: { type: 'none' },
    });
    // Dot
    slide.addShape(prs.ShapeType.ellipse, {
      x: x + 0.12, y: y + i * rowH + rowH / 2 - 0.045,
      w: 0.09, h: 0.09,
      fill: { color: C.ocean }, line: { type: 'none' },
    });
    slide.addText(trunc(txt, 95), {
      x: x + 0.27, y: y + i * rowH,
      w: w - 0.35, h: rowH,
      fontFace: FONT, fontSize, color: C.charcoal,
      valign: 'middle', wrap: true,
    });
  });
  if (items.length > maxItems) {
    slide.addText(`+${items.length - maxItems} more in full report`, {
      x, y: y + shown.length * rowH + 0.02,
      w, h: 0.2,
      fontFace: FONT, fontSize: 7.5, italic: true, color: C.silver,
      align: 'right',
    });
  }
}

// ─── MAIN GENERATOR ───────────────────────────────────────────────────────────
function generateSlide(data) {
  const prs = new PptxGenJS();
  prs.defineLayout({ name: 'WIDE', width: W, height: H });
  prs.layout = 'WIDE';

  const slide = prs.addSlide();
  slide.background = { color: 'F0F4F9' }; // light blue-grey canvas

  const lang       = data.language || 'en';
  const projectName= data.project_name || 'Project';
  const weekLabel  = data.week_label || '';
  const rag        = ragColor(data.rag_status);
  const aiSummary  = data.ai_summary || '';
  const teamMembers= data.team_members || [];

  const tasks    = safeJson(data.tasks_completed, []);
  const nextTasks= safeJson(data.next_tasks, []);
  const decisions= safeJson(data.management_decisions, []);
  const kpis     = safeJson(data.kpi_updates, []);
  const risks    = safeJson(data.risks_blockers, []);

  // ── LEFT BANNER (vertical navy strip) ───────────────────────────────────────
  const bannerW = 2.6;
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: bannerW, h: H,
    fill: { color: C.navy }, line: { type: 'none' },
  });

  // Decorative accent line on banner
  slide.addShape(prs.ShapeType.rect, {
    x: bannerW - 0.06, y: 0, w: 0.06, h: H,
    fill: { color: C.ocean }, line: { type: 'none' },
  });

  // "WEEKLY STATUS UPDATE" super-label (rotated feel — stacked chars workaround: vertical text box)
  slide.addText(label('weekly_update', lang).toUpperCase(), {
    x: 0.12, y: 0.35, w: bannerW - 0.3, h: 0.35,
    fontFace: FONT, fontSize: 7.5, bold: true,
    color: C.sky, charSpacing: 2,
    align: 'left', valign: 'middle',
  });

  // Project Name — large, wrapping
  slide.addText(projectName, {
    x: 0.18, y: 0.78, w: bannerW - 0.36, h: 2.2,
    fontFace: FONT, fontSize: 22, bold: true,
    color: C.white,
    align: 'left', valign: 'top',
    wrap: true, lineSpacingMultiple: 1.15,
  });

  // Week label pill
  slide.addShape(prs.ShapeType.roundRect, {
    x: 0.18, y: 3.1, w: bannerW - 0.5, h: 0.34,
    fill: { color: C.ocean }, line: { type: 'none' },
    rectRadius: 0.05,
  });
  slide.addText(weekLabel, {
    x: 0.18, y: 3.1, w: bannerW - 0.5, h: 0.34,
    fontFace: FONT, fontSize: 8.5, color: C.white,
    bold: true, align: 'center', valign: 'middle',
  });

  // RAG pill
  slide.addShape(prs.ShapeType.roundRect, {
    x: 0.18, y: 3.55, w: bannerW - 0.5, h: 0.38,
    fill: { color: rag.dot }, line: { type: 'none' },
    rectRadius: 0.05,
  });
  slide.addText(`● ${rag.label.toUpperCase()}`, {
    x: 0.18, y: 3.55, w: bannerW - 0.5, h: 0.38,
    fontFace: FONT, fontSize: 9, color: C.white,
    bold: true, align: 'center', valign: 'middle',
  });

  // ── TEAM block in banner ─────────────────────────────────────────────────────
  slide.addText(label('team', lang).toUpperCase(), {
    x: 0.18, y: 4.15, w: bannerW - 0.36, h: 0.22,
    fontFace: FONT, fontSize: 7, bold: true, color: C.sky,
    charSpacing: 1.5, align: 'left',
  });

  let teamY = 4.42;
  teamMembers.slice(0, 5).forEach((m) => {
    const name = m.name || m;
    const role = m.role || '';
    // Avatar circle
    slide.addShape(prs.ShapeType.ellipse, {
      x: 0.18, y: teamY, w: 0.34, h: 0.34,
      fill: { color: C.ocean }, line: { type: 'none' },
    });
    slide.addText(initials(name), {
      x: 0.18, y: teamY, w: 0.34, h: 0.34,
      fontFace: FONT, fontSize: 8, bold: true, color: C.white,
      align: 'center', valign: 'middle',
    });
    // Name + role
    slide.addText(trunc(name, 22), {
      x: 0.6, y: teamY, w: bannerW - 0.72, h: 0.18,
      fontFace: FONT, fontSize: 8.5, bold: true, color: C.white,
      valign: 'bottom',
    });
    if (role) {
      slide.addText(trunc(role, 26), {
        x: 0.6, y: teamY + 0.18, w: bannerW - 0.72, h: 0.15,
        fontFace: FONT, fontSize: 7.5, color: C.silver,
        valign: 'top',
      });
    }
    teamY += 0.42;
  });

  // Date at bottom of banner
  const now = new Date();
  const dateStr = now.toLocaleDateString(lang === 'de' ? 'de-DE' : 'en-GB', {
    day: 'numeric', month: 'long', year: 'numeric',
  });
  slide.addText(dateStr, {
    x: 0.1, y: H - 0.45, w: bannerW - 0.2, h: 0.3,
    fontFace: FONT, fontSize: 7.5, color: C.silver,
    align: 'center', valign: 'middle',
  });

  // ── RIGHT CONTENT AREA ───────────────────────────────────────────────────────
  const cx = bannerW + 0.18;   // content x
  const cw = W - cx - 0.18;   // content width
  const topY = 0.18;
  const gap = 0.16;

  // ── ROW 1: Completed Tasks + Next Steps (side by side) ─────────────────────
  const r1h = 2.42;
  const col1w = cw * 0.55;
  const col2w = cw * 0.42;
  const col2x = cx + col1w + gap;

  // Card 1: Completed this week
  const c1contentY = addCard(slide, prs, {
    x: cx, y: topY, w: col1w, h: r1h,
    title: label('this_week', lang).toUpperCase(),
    accentColor: C.ocean,
  });

  const taskStrings = tasks.map(t =>
    typeof t === 'string' ? t
    : t.result ? `${t.task}  →  ${t.result}`
    : t.task || ''
  ).filter(Boolean);

  if (taskStrings.length) {
    addBulletRows(slide, prs, {
      x: cx, y: c1contentY, w: col1w,
      items: taskStrings, maxItems: 5, rowH: 0.3,
    });
  } else {
    slide.addText(label('no_data', lang), {
      x: cx + 0.15, y: c1contentY + 0.1, w: col1w - 0.3, h: 0.3,
      fontFace: FONT, fontSize: 8.5, italic: true, color: C.silver,
    });
  }

  // Card 2: Next Steps
  const c2contentY = addCard(slide, prs, {
    x: col2x, y: topY, w: col2w, h: r1h,
    title: label('next_steps', lang).toUpperCase(),
    accentColor: C.sky,
  });

  const nextStrings = nextTasks.map(t => {
    if (typeof t === 'string') return t;
    const owner = t.owner ? `[${t.owner}]` : '';
    const due   = t.due_date ? ` · ${t.due_date}` : '';
    return `${t.task || ''}  ${owner}${due}`.trim();
  }).filter(Boolean);

  if (nextStrings.length) {
    addBulletRows(slide, prs, {
      x: col2x, y: c2contentY, w: col2w,
      items: nextStrings, maxItems: 5, rowH: 0.3,
    });
  } else {
    slide.addText(label('no_data', lang), {
      x: col2x + 0.15, y: c2contentY + 0.1, w: col2w - 0.3, h: 0.3,
      fontFace: FONT, fontSize: 8.5, italic: true, color: C.silver,
    });
  }

  // ── ROW 2: KPIs + Risks (side by side) ─────────────────────────────────────
  const r2y = topY + r1h + gap;
  const r2h = 1.9;
  const kpiW = cw * 0.38;
  const riskW = cw * 0.59;
  const riskX = cx + kpiW + gap;

  // Card 3: KPIs
  const c3contentY = addCard(slide, prs, {
    x: cx, y: r2y, w: kpiW, h: r2h,
    title: label('kpis', lang).toUpperCase(),
    accentColor: '1A7CC1',
  });

  if (kpis.length) {
    kpis.slice(0, 4).forEach((kpi, i) => {
      const metric = typeof kpi === 'string' ? kpi : `${kpi.metric || ''}`;
      const value  = typeof kpi === 'string' ? '' : `${kpi.value || ''}`;
      const trend  = typeof kpi === 'object' && kpi.trend ? kpi.trend : '';
      const rowY   = c3contentY + i * 0.35;
      const bg     = i % 2 === 1 ? C.frost : C.white;

      slide.addShape(prs.ShapeType.rect, {
        x: cx, y: rowY, w: kpiW, h: 0.35,
        fill: { color: bg }, line: { type: 'none' },
      });
      slide.addText(trunc(metric, 30), {
        x: cx + 0.12, y: rowY, w: kpiW * 0.6, h: 0.35,
        fontFace: FONT, fontSize: 8.5, color: C.charcoal, valign: 'middle',
      });
      if (value) {
        slide.addText(value + (trend ? ` ${trend}` : ''), {
          x: cx + kpiW * 0.6, y: rowY, w: kpiW * 0.36, h: 0.35,
          fontFace: FONT, fontSize: 9, bold: true, color: C.ocean,
          align: 'right', valign: 'middle',
        });
      }
    });
  } else {
    slide.addText(label('no_data', lang), {
      x: cx + 0.15, y: c3contentY + 0.1, w: kpiW - 0.3, h: 0.3,
      fontFace: FONT, fontSize: 8.5, italic: true, color: C.silver,
    });
  }

  // Card 4: Risks & Blockers
  const c4contentY = addCard(slide, prs, {
    x: riskX, y: r2y, w: riskW, h: r2h,
    title: label('risks', lang).toUpperCase(),
    accentColor: C.amber,
  });

  if (risks.length) {
    risks.slice(0, 4).forEach((risk, i) => {
      const issue = typeof risk === 'string' ? risk : (risk.issue || risk.risk || '');
      const mit   = typeof risk === 'object' ? (risk.mitigation || '') : '';
      const rowH  = mit ? 0.44 : 0.3;
      const rowY  = c4contentY + i * rowH;
      const bg    = i % 2 === 1 ? C.frost : C.white;

      slide.addShape(prs.ShapeType.rect, {
        x: riskX, y: rowY, w: riskW, h: rowH,
        fill: { color: bg }, line: { type: 'none' },
      });
      slide.addShape(prs.ShapeType.ellipse, {
        x: riskX + 0.12, y: rowY + (rowH / 2) - 0.045,
        w: 0.09, h: 0.09,
        fill: { color: C.amber }, line: { type: 'none' },
      });
      slide.addText(trunc(issue, 80), {
        x: riskX + 0.27, y: rowY + (mit ? 0.02 : 0),
        w: riskW - 0.35, h: 0.26,
        fontFace: FONT, fontSize: 8.5, color: C.charcoal, valign: 'middle',
      });
      if (mit) {
        slide.addText(`↳ ${trunc(mit, 90)}`, {
          x: riskX + 0.27, y: rowY + 0.26,
          w: riskW - 0.35, h: 0.18,
          fontFace: FONT, fontSize: 7.5, italic: true, color: C.slate,
        });
      }
    });
  } else {
    slide.addText(label('no_risks', lang), {
      x: riskX + 0.15, y: c4contentY + 0.1, w: riskW - 0.3, h: 0.3,
      fontFace: FONT, fontSize: 8.5, italic: true, color: C.silver,
    });
  }

  // ── ROW 3: Management Decisions (full width) ────────────────────────────────
  const r3y = r2y + r2h + gap;
  const r3h = H - r3y - 0.5; // remaining space minus footer
  const decContentY = addCard(slide, prs, {
    x: cx, y: r3y, w: cw, h: r3h,
    title: label('decisions', lang).toUpperCase(),
    accentColor: C.red,
  });

  if (decisions.length) {
    decisions.slice(0, 3).forEach((dec, i) => {
      const text    = typeof dec === 'string' ? dec : (dec.decision || '');
      const ctx     = typeof dec === 'object' ? (dec.context || '') : '';
      const urg     = typeof dec === 'object' ? (dec.urgency || '') : '';
      const ucol    = urgencyColor(urg);
      const rowH    = ctx ? 0.46 : 0.32;
      const rowY    = decContentY + i * rowH;
      const bg      = i % 2 === 1 ? C.frost : C.white;
      const urgLabel= label(urgencyKey(urg), lang);

      slide.addShape(prs.ShapeType.rect, {
        x: cx, y: rowY, w: cw, h: rowH,
        fill: { color: bg }, line: { type: 'none' },
      });

      // Number badge
      slide.addShape(prs.ShapeType.ellipse, {
        x: cx + 0.1, y: rowY + rowH / 2 - 0.13,
        w: 0.26, h: 0.26,
        fill: { color: C.navy }, line: { type: 'none' },
      });
      slide.addText(`${i + 1}`, {
        x: cx + 0.1, y: rowY + rowH / 2 - 0.13,
        w: 0.26, h: 0.26,
        fontFace: FONT, fontSize: 8, bold: true, color: C.white,
        align: 'center', valign: 'middle',
      });

      slide.addText(trunc(text, 110), {
        x: cx + 0.45, y: rowY + 0.03,
        w: cw - 1.8, h: 0.26,
        fontFace: FONT, fontSize: 8.5, color: C.charcoal, valign: 'middle',
      });
      if (ctx) {
        slide.addText(`↳ ${trunc(ctx, 120)}`, {
          x: cx + 0.45, y: rowY + 0.27,
          w: cw - 1.8, h: 0.18,
          fontFace: FONT, fontSize: 7.5, italic: true, color: C.slate,
        });
      }

      // Urgency tag (right side)
      slide.addShape(prs.ShapeType.roundRect, {
        x: cx + cw - 1.25, y: rowY + rowH / 2 - 0.12,
        w: 1.1, h: 0.24,
        fill: { color: ucol.bg }, line: { color: ucol.fg, width: 0.5 },
        rectRadius: 0.04,
      });
      slide.addText(urgLabel, {
        x: cx + cw - 1.25, y: rowY + rowH / 2 - 0.12,
        w: 1.1, h: 0.24,
        fontFace: FONT, fontSize: 7.5, bold: true, color: ucol.fg,
        align: 'center', valign: 'middle',
      });
    });
  } else {
    slide.addText(label('no_decisions', lang), {
      x: cx + 0.15, y: decContentY + 0.05, w: cw - 0.3, h: 0.28,
      fontFace: FONT, fontSize: 8.5, italic: true, color: C.silver,
    });
  }

  // ── FOOTER ───────────────────────────────────────────────────────────────────
  const footerY = H - 0.42;
  slide.addShape(prs.ShapeType.rect, {
    x: 0, y: footerY, w: W, h: 0.42,
    fill: { color: C.navy }, line: { type: 'none' },
  });
  // Left: confidential
  slide.addText(label('confidential', lang), {
    x: bannerW + 0.2, y: footerY,
    w: 3.5, h: 0.42,
    fontFace: FONT, fontSize: 7.5, color: C.silver,
    align: 'left', valign: 'middle',
  });
  // Center: AI summary
  if (aiSummary) {
    slide.addText(trunc(aiSummary, 130), {
      x: bannerW + 3.8, y: footerY,
      w: W - bannerW - 7.2, h: 0.42,
      fontFace: FONT, fontSize: 7.5, italic: true, color: 'A0B8D0',
      align: 'center', valign: 'middle',
    });
  }
  // Right: timestamp
  const ts = new Date().toLocaleString(lang === 'de' ? 'de-DE' : 'en-GB');
  slide.addText(ts, {
    x: W - 3.4, y: footerY,
    w: 3.2, h: 0.42,
    fontFace: FONT, fontSize: 7.5, color: C.silver,
    align: 'right', valign: 'middle',
  });

  return prs;
}

// ─── ENTRY POINT ─────────────────────────────────────────────────────────────
async function main() {
  if (process.argv.length < 3) {
    console.error('Usage: node slide_generator.js <path_to_json_file>');
    process.exit(1);
  }

  const jsonPath = process.argv[2];
  if (!fs.existsSync(jsonPath)) {
    console.error(`Error: File not found: ${jsonPath}`);
    process.exit(1);
  }

  let slideData;
  try {
    slideData = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
    console.log(`[OK] JSON loaded. Keys: ${Object.keys(slideData).join(', ')}`);
  } catch (e) {
    console.error(`Error reading JSON: ${e.message}`);
    process.exit(1);
  }

  if (!slideData.project_name) {
    console.error('Error: Missing required field: project_name');
    process.exit(1);
  }

  try {
    console.log('[OK] Generating slide...');
    const prs = generateSlide(slideData);

    let outputPath = slideData.output_path ||
      path.join('exports', `${slideData.project_name}_${slideData.week_number || 'update'}.pptx`);

    const absOut = path.isAbsolute(outputPath) ? outputPath : path.resolve(process.cwd(), outputPath);
    const relOut = path.isAbsolute(outputPath) ? path.relative(process.cwd(), outputPath) : outputPath;

    const dir = path.dirname(absOut);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    await prs.writeFile({ fileName: relOut });
    await new Promise(r => setTimeout(r, 100));

    if (fs.existsSync(absOut)) {
      console.log(`✓ Slide generated: ${absOut} (${fs.statSync(absOut).size} bytes)`);
      process.exit(0);
    } else {
      console.error(`Error: File not created at ${absOut}`);
      process.exit(1);
    }
  } catch (e) {
    console.error(`Error generating slide: ${e.message}\n${e.stack}`);
    process.exit(1);
  }
}

if (require.main === module) main();
module.exports = { generateSlide };
