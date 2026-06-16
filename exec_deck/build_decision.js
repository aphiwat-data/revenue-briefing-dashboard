// Executive Decision Deck — "Own Your Data"
// Sell the value of a company-owned database / data layer.
// Navy + Gold premium executive theme.

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const {
  FaDatabase, FaLayerGroup, FaInfinity, FaSlidersH, FaRobot, FaChartLine,
  FaCheckCircle, FaTimesCircle, FaExclamationTriangle, FaLock, FaKey,
  FaCloud, FaCogs, FaUsers, FaSitemap, FaBalanceScale, FaCoins,
  FaBuilding, FaServer, FaArrowRight, FaEye, FaHandshake, FaBolt,
  FaSearchDollar, FaFileExcel, FaProjectDiagram,
} = require("react-icons/fa");

// ============================================================
// PALETTE — Midnight + Gold
// ============================================================
const C = {
  bgDark:   "0A1428",
  bgMid:    "13243F",
  bgLight:  "F7F8FC",
  card:     "FFFFFF",
  ink:      "0A1428",
  inkSoft:  "475569",
  inkMute:  "94A3B8",
  textLight:"F7F8FC",
  textDim:  "B8C5D6",
  gold:     "D4A437",
  teal:     "14B8A6",
  coral:    "F97066",
  divider:  "1E3A5F",
  ruleLt:   "E2E8F0",
};

function renderSvg(IconComponent, color = "#FFFFFF", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}
async function icon(IconComponent, color = "#FFFFFF", size = 256) {
  const svg = renderSvg(IconComponent, color, size);
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Revenue Intelligence Team";
  pres.title  = "Own Your Data — Executive Decision Brief";

  const W = 13.3, H = 7.5;
  const TOTAL = 12;

  const shadowSoft = () => ({ type: "outer", color: "000000", blur: 18, offset: 4, angle: 90, opacity: 0.18 });
  const shadowCard = () => ({ type: "outer", color: "0A1428", blur: 12, offset: 3, angle: 90, opacity: 0.10 });

  const goldChip = (s, x, y, label) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.18, h: 0.32, fill: { color: C.gold }, line: { type: "none" } });
    s.addText(label, { x: x + 0.28, y: y - 0.02, w: 11, h: 0.36, margin: 0,
      fontFace: "Calibri", fontSize: 12, bold: true, color: C.gold, charSpacing: 4 });
  };
  const footer = (s, n) => {
    s.addText("OWN YOUR DATA  ·  EXECUTIVE DECISION BRIEF", {
      x: 0.5, y: H - 0.32, w: 9, h: 0.28, fontFace: "Calibri", fontSize: 9,
      color: C.inkMute, charSpacing: 4, margin: 0 });
    s.addText(`${n} / ${TOTAL}`, { x: W - 1.5, y: H - 0.32, w: 1, h: 0.28, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0 });
  };
  const footerDark = (s, n) => {
    s.addText("OWN YOUR DATA  ·  EXECUTIVE DECISION BRIEF", {
      x: 0.5, y: H - 0.32, w: 9, h: 0.28, fontFace: "Calibri", fontSize: 9,
      color: C.inkMute, charSpacing: 4, margin: 0 });
    s.addText(`${n} / ${TOTAL}`, { x: W - 1.5, y: H - 0.32, w: 1, h: 0.28, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0 });
  };

  // ============================================================
  // SLIDE 1 — TITLE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: H, fill: { color: C.gold }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: W - 5.0, y: 0, w: 5.0, h: H, fill: { color: C.bgMid, transparency: 35 }, line: { type: "none" } });

    s.addText("EXECUTIVE DECISION BRIEF", {
      x: 1, y: 1.6, w: 10, h: 0.4, fontFace: "Calibri", fontSize: 13, bold: true,
      color: C.gold, charSpacing: 8, margin: 0 });
    s.addText("Own Your Data.", {
      x: 1, y: 2.15, w: 11, h: 1.5, fontFace: "Georgia", fontSize: 64, bold: true,
      color: C.textLight, margin: 0 });
    s.addText("The case for a data platform the company controls —\nnot another tool we rent.", {
      x: 1, y: 3.85, w: 11, h: 1.1, fontFace: "Calibri", fontSize: 22, italic: true,
      color: C.textDim, margin: 0, lineSpacingMultiple: 1.15 });

    s.addShape(pres.shapes.LINE, { x: 1, y: 5.55, w: 4, h: 0, line: { color: C.gold, width: 2 } });
    s.addText("Prepared for leadership review", {
      x: 1, y: 5.7, w: 8, h: 0.4, fontFace: "Calibri", fontSize: 14, color: C.textLight, bold: true, margin: 0 });
    s.addText("Revenue Intelligence Team  ·  2026", {
      x: 1, y: 6.1, w: 8, h: 0.35, fontFace: "Calibri", fontSize: 12, color: C.inkMute, charSpacing: 2, margin: 0 });
  }

  // ============================================================
  // SLIDE 2 — THE DECISION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "THE DECISION ON THE TABLE");

    s.addText("Do we keep renting our insight —\nor start owning it?", {
      x: 0.5, y: 0.95, w: 12.3, h: 1.7, fontFace: "Georgia", fontSize: 38, bold: true,
      color: C.ink, margin: 0, lineSpacingMultiple: 1.05 });

    s.addText("Today every number we trust lives inside a vendor's tool. This brief lays out the problem, the options, and a clear ask.", {
      x: 0.5, y: 2.75, w: 12.3, h: 0.5, fontFace: "Calibri", fontSize: 16, italic: true, color: C.inkSoft, margin: 0 });

    const items = [
      { ic: FaExclamationTriangle, t: "The Problem", d: "Our data is split across five platforms — and reconciled by hand." },
      { ic: FaBalanceScale,        t: "The Options", d: "Buy a big BI tool · stay as we are · build our own data layer." },
      { ic: FaKey,                 t: "The Ask",     d: "A small, owned foundation that pays back from month one." },
    ];
    const cW = 3.95, cGap = 0.2, startX = (W - (3 * cW + 2 * cGap)) / 2;
    for (let i = 0; i < items.length; i++) {
      const x = startX + i * (cW + cGap);
      const it = items[i];
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.6, w: cW, h: 2.7, fill: { color: C.card }, line: { color: C.ruleLt, width: 0.75 }, shadow: shadowCard() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.6, w: cW, h: 0.08, fill: { color: C.gold }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.4, y: 3.95, w: 0.9, h: 0.9, fill: { color: C.bgDark }, line: { type: "none" } });
      s.addImage({ data: await icon(it.ic, "#D4A437", 256), x: x + 0.63, y: 4.18, w: 0.44, h: 0.44 });
      s.addText(it.t, { x: x + 0.4, y: 5.0, w: cW - 0.8, h: 0.5, fontFace: "Georgia", fontSize: 22, bold: true, color: C.ink, margin: 0 });
      s.addText(it.d, { x: x + 0.4, y: 5.55, w: cW - 0.8, h: 0.7, fontFace: "Calibri", fontSize: 13, color: C.inkSoft, margin: 0 });
    }
    footer(s, 2);
  }

  // ============================================================
  // SLIDE 3 — ROOT PROBLEM
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "THE ROOT PROBLEM");

    s.addText("Our data lives in five places.\nThe truth lives in none.", {
      x: 0.5, y: 0.95, w: 7.6, h: 1.6, fontFace: "Georgia", fontSize: 33, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.05 });

    // Left: scattered tools
    const tools = [
      { n: "Duetto",    d: "Pricing & forecast" },
      { n: "SiteMinder",d: "Channel & distribution" },
      { n: "Comanche",  d: "Operations & billing" },
      { n: "Power BI",  d: "Reporting (no owner)" },
      { n: "Excel",     d: "Manual gap-filling" },
    ];
    tools.forEach((t, i) => {
      const y = 2.75 + i * 0.78;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 6.0, h: 0.66, fill: { color: C.card }, line: { color: C.ruleLt, width: 0.75 }, shadow: shadowCard() });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.09, h: 0.66, fill: { color: C.coral }, line: { type: "none" } });
      s.addText(t.n, { x: 0.75, y, w: 2.4, h: 0.66, valign: "middle", fontFace: "Calibri", fontSize: 15, bold: true, color: C.ink, margin: 0 });
      s.addText(t.d, { x: 3.1, y, w: 3.3, h: 0.66, valign: "middle", fontFace: "Calibri", fontSize: 12, italic: true, color: C.inkSoft, margin: 0 });
    });
    s.addText("Each sees half the picture. None talks to the others.", {
      x: 0.5, y: 6.7, w: 6.5, h: 0.35, fontFace: "Calibri", fontSize: 12, italic: true, color: C.coral, bold: true, margin: 0 });

    // Right: research stats
    s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 2.75, w: 5.8, h: 1.75, fill: { color: C.bgDark }, line: { type: "none" }, shadow: shadowSoft() });
    s.addText("< 1 in 4", { x: 7.3, y: 2.9, w: 5.2, h: 0.9, fontFace: "Georgia", fontSize: 52, bold: true, color: C.gold, margin: 0 });
    s.addText("hotels have their core systems fully integrated.", {
      x: 7.3, y: 3.85, w: 5.2, h: 0.55, fontFace: "Calibri", fontSize: 13, color: C.textDim, margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: 7.0, y: 4.65, w: 5.8, h: 1.75, fill: { color: C.card }, line: { color: C.ruleLt, width: 1 }, shadow: shadowCard() });
    s.addImage({ data: await icon(FaFileExcel, "#F97066", 256), x: 7.3, y: 4.9, w: 0.55, h: 0.55 });
    s.addText("“Disconnected systems force teams to extract reports, compare datasets, and chase discrepancies — a daily manual task.”", {
      x: 7.3, y: 5.55, w: 5.2, h: 0.8, fontFace: "Calibri", fontSize: 12.5, italic: true, color: C.inkSoft, margin: 0 });
    s.addText("— Hospitality Net, 2026", {
      x: 7.3, y: 6.32, w: 5.2, h: 0.3, fontFace: "Calibri", fontSize: 10, bold: true, color: C.inkMute, margin: 0 });

    footer(s, 3);
  }

  // ============================================================
  // SLIDE 4 — HIDDEN COST TODAY
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "WHAT IT COSTS US TODAY");

    s.addText("Fragmentation isn't free. We pay for it every day.", {
      x: 0.5, y: 0.95, w: 12.3, h: 0.9, fontFace: "Georgia", fontSize: 31, bold: true, color: C.ink, margin: 0 });

    const rows = [
      { ic: FaTimesCircle,  t: "Conflicting numbers",   d: "Duetto says one occupancy, Power BI another. Which do we price on?" },
      { ic: FaFileExcel,    t: "Manual reconciliation",  d: "Hours each week spent copying, comparing, and patching spreadsheets." },
      { ic: FaBolt,         t: "Slow, late decisions",   d: "By the time numbers agree, the pricing window has closed." },
      { ic: FaUsers,        t: "Knowledge trapped",      d: "The logic lives in one person's file. They leave — it leaves with them." },
      { ic: FaChartLine,    t: "Stuck at room revenue",  d: "No unified data means no path to total-profit (TRevPAR) management." },
    ];
    const startY = 2.15, rowH = 0.82;
    rows.forEach((r, i) => {
      const y = startY + i * (rowH + 0.06);
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: W - 1.0, h: rowH, fill: { color: i % 2 ? "EEF2F7" : C.card }, line: { color: C.ruleLt, width: 0.5 } });
      s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.18, w: 0.46, h: 0.46, fill: { color: C.coral }, line: { type: "none" } });
    });
    for (let i = 0; i < rows.length; i++) {
      const y = startY + i * (rowH + 0.06);
      s.addImage({ data: await icon(rows[i].ic, "#FFFFFF", 256), x: 0.88, y: y + 0.26, w: 0.3, h: 0.3 });
      s.addText(rows[i].t, { x: 1.5, y, w: 3.6, h: rowH, valign: "middle", fontFace: "Calibri", fontSize: 16, bold: true, color: C.ink, margin: 0 });
      s.addText(rows[i].d, { x: 5.2, y, w: 7.6, h: rowH, valign: "middle", fontFace: "Calibri", fontSize: 13, color: C.inkSoft, margin: 0 });
    }
    footer(s, 4);
  }

  // ============================================================
  // SLIDE 5 — THE ANSWER (transition, dark)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    goldChip(s, 0.7, 0.6, "THE ANSWER");

    s.addText("The answer isn't another tool.\nIt's a foundation we own.", {
      x: 0.7, y: 1.5, w: 12, h: 2.0, fontFace: "Georgia", fontSize: 44, bold: true, color: C.textLight, margin: 0, lineSpacingMultiple: 1.05 });

    s.addText("A database the company controls — one place where every platform reports in, and every number reconciles.", {
      x: 0.7, y: 3.7, w: 11.5, h: 0.9, fontFace: "Calibri", fontSize: 19, italic: true, color: C.textDim, margin: 0 });

    // central icon
    s.addShape(pres.shapes.OVAL, { x: W/2 - 0.85, y: 4.85, w: 1.7, h: 1.7, fill: { color: C.bgMid }, line: { color: C.gold, width: 2 } });
    s.addImage({ data: await icon(FaDatabase, "#D4A437", 256), x: W/2 - 0.45, y: 5.25, w: 0.9, h: 0.9 });

    footerDark(s, 5);
  }

  // ============================================================
  // SLIDE 6 — 6 BENEFITS OF OWNING A DATABASE (money slide)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "WHY A DATABASE CHANGES EVERYTHING");

    s.addText("Six things only an owned database gives us.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.75, fontFace: "Georgia", fontSize: 30, bold: true, color: C.ink, margin: 0 });

    const benefits = [
      { ic: FaLayerGroup, t: "One source of truth", d: "Every team prices, reports, and reconciles off the same numbers — no more debates." },
      { ic: FaLock,       t: "We own the asset",    d: "Data lives with us, not a vendor. Cancel any tool — the history stays." },
      { ic: FaInfinity,   t: "Unlimited history",   d: "Vendors keep 1–2 years. We keep forever — real year-over-year comparison." },
      { ic: FaSlidersH,   t: "Metrics our way",     d: "Build any KPI the business wants. No waiting on a vendor roadmap." },
      { ic: FaRobot,      t: "Zero manual work",    d: "Data flows in automatically. No copy-paste, no reconciliation by hand." },
      { ic: FaChartLine,  t: "TRevPAR-ready",       d: "Joins rooms + F&B + channel into one view — the foundation for total-profit." },
    ];
    const gW = 3.95, gH = 2.15, gx = 0.18, gy = 0.18;
    const sx = (W - (3 * gW + 2 * gx)) / 2;
    for (let i = 0; i < benefits.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = sx + col * (gW + gx);
      const y = 1.75 + row * (gH + gy);
      const b = benefits[i];
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: gW, h: gH, fill: { color: C.card }, line: { color: C.ruleLt, width: 0.75 }, shadow: shadowCard() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: gH, fill: { color: C.gold }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: C.bgDark }, line: { type: "none" } });
      s.addImage({ data: await icon(b.ic, "#D4A437", 256), x: x + 0.5, y: y + 0.45, w: 0.4, h: 0.4 });
      s.addText(b.t, { x: x + 1.25, y: y + 0.32, w: gW - 1.4, h: 0.65, valign: "middle", fontFace: "Georgia", fontSize: 18, bold: true, color: C.ink, margin: 0 });
      s.addText(b.d, { x: x + 0.35, y: y + 1.15, w: gW - 0.65, h: 0.9, fontFace: "Calibri", fontSize: 12, color: C.inkSoft, margin: 0 });
    }
    footer(s, 6);
  }

  // ============================================================
  // SLIDE 7 — SINGLE SOURCE OF TRUTH (before/after)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "THE BIG SHIFT");

    s.addText("From four answers to one.", {
      x: 0.5, y: 0.95, w: 12.3, h: 0.85, fontFace: "Georgia", fontSize: 32, bold: true, color: C.ink, margin: 0 });

    // BEFORE
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.1, w: 6.0, h: 4.3, fill: { color: C.card }, line: { color: "FECACA", width: 1.5 }, shadow: shadowCard() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.1, w: 6.0, h: 0.55, fill: { color: C.coral }, line: { type: "none" } });
    s.addText("TODAY  ·  FOUR SOURCES", { x: 0.5, y: 2.1, w: 6.0, h: 0.55, align: "center", valign: "middle", fontFace: "Calibri", fontSize: 14, bold: true, color: C.card, charSpacing: 4, margin: 0 });
    const before = [
      "Duetto:  occupancy 85%",
      "SiteMinder:  18 rooms left (92%)",
      "Power BI:  on-book 80%",
      "Excel:  hand-typed 88%",
    ];
    before.forEach((t, i) => {
      const y = 2.95 + i * 0.72;
      s.addImage({ data: i === 3 ? null : null, x: 0, y: 0, w: 0.01, h: 0.01 });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.85, y, w: 5.3, h: 0.55, fill: { color: "FEF2F2" }, line: { color: "FECACA", width: 0.75 } });
      s.addText(t, { x: 1.05, y, w: 5.0, h: 0.55, valign: "middle", fontFace: "Consolas", fontSize: 13.5, color: C.ink, margin: 0 });
    });
    s.addText("Four numbers. Endless arguments.", { x: 0.5, y: 5.95, w: 6.0, h: 0.4, align: "center", fontFace: "Calibri", fontSize: 13, italic: true, bold: true, color: C.coral, margin: 0 });

    // arrow
    s.addImage({ data: await icon(FaArrowRight, "#D4A437", 256), x: 6.62, y: 4.0, w: 0.55, h: 0.55 });

    // AFTER
    s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 2.1, w: 6.0, h: 4.3, fill: { color: C.card }, line: { color: "99F6E4", width: 1.5 }, shadow: shadowCard() });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 2.1, w: 6.0, h: 0.55, fill: { color: C.teal }, line: { type: "none" } });
    s.addText("WITH OUR DATABASE  ·  ONE SOURCE", { x: 6.8, y: 2.1, w: 6.0, h: 0.55, align: "center", valign: "middle", fontFace: "Calibri", fontSize: 14, bold: true, color: C.card, charSpacing: 3, margin: 0 });

    s.addShape(pres.shapes.OVAL, { x: 9.05, y: 3.05, w: 1.5, h: 1.5, fill: { color: C.bgDark }, line: { color: C.teal, width: 2 } });
    s.addImage({ data: await icon(FaCheckCircle, "#14B8A6", 256), x: 9.45, y: 3.45, w: 0.7, h: 0.7 });
    s.addText("Occupancy: 84%", { x: 6.8, y: 4.75, w: 6.0, h: 0.6, align: "center", fontFace: "Georgia", fontSize: 26, bold: true, color: C.ink, margin: 0 });
    s.addText("Reconciled across every platform. Auditable to the source. Same for everyone.", {
      x: 7.3, y: 5.45, w: 5.0, h: 0.85, align: "center", fontFace: "Calibri", fontSize: 13, color: C.inkSoft, margin: 0 });

    footer(s, 7);
  }

  // ============================================================
  // SLIDE 8 — STRATEGIC UPSIDE: TRevPAR (dark)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    goldChip(s, 0.7, 0.55, "THE STRATEGIC UPSIDE");

    s.addText("From reporting the past\nto optimizing profit.", {
      x: 0.7, y: 1.1, w: 7.5, h: 1.8, fontFace: "Georgia", fontSize: 38, bold: true, color: C.textLight, margin: 0, lineSpacingMultiple: 1.05 });
    s.addText("A unified database is the entry ticket to Total Revenue Management — pricing on total profit, not just room rate.", {
      x: 0.7, y: 3.05, w: 7.6, h: 1.0, fontFace: "Calibri", fontSize: 16, italic: true, color: C.textDim, margin: 0 });

    // progression
    const steps = [
      { t: "RevPAR", d: "Room revenue only" },
      { t: "TRevPAR", d: "+ F&B, spa, ancillary" },
      { t: "GOPPAR", d: "+ profit after cost" },
    ];
    steps.forEach((st, i) => {
      const y = 4.4 + i * 0.85;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 7.4, h: 0.7, fill: { color: C.bgMid }, line: { color: C.divider, width: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.09, h: 0.7, fill: { color: C.gold }, line: { type: "none" } });
      s.addText(st.t, { x: 0.95, y, w: 2.3, h: 0.7, valign: "middle", fontFace: "Calibri", fontSize: 16, bold: true, color: C.gold, margin: 0 });
      s.addText(st.d, { x: 3.2, y, w: 4.7, h: 0.7, valign: "middle", fontFace: "Calibri", fontSize: 13, color: C.textDim, margin: 0 });
    });

    // big stat right
    s.addShape(pres.shapes.RECTANGLE, { x: 8.6, y: 1.85, w: 4.2, h: 4.4, fill: { color: C.bgMid }, line: { color: C.gold, width: 1.5 }, shadow: shadowSoft() });
    s.addShape(pres.shapes.RECTANGLE, { x: 8.6, y: 1.85, w: 4.2, h: 0.07, fill: { color: C.gold }, line: { type: "none" } });
    s.addText("25–40%", { x: 8.6, y: 2.6, w: 4.2, h: 1.3, align: "center", fontFace: "Georgia", fontSize: 58, bold: true, color: C.gold, margin: 0 });
    s.addText("MORE PROFITABLE", { x: 8.6, y: 3.95, w: 4.2, h: 0.4, align: "center", fontFace: "Calibri", fontSize: 14, bold: true, color: C.textLight, charSpacing: 4, margin: 0 });
    s.addText("Hotels that manage on total revenue out-earn those that watch room rate alone.", {
      x: 8.95, y: 4.5, w: 3.5, h: 1.2, align: "center", fontFace: "Calibri", fontSize: 13, italic: true, color: C.textDim, margin: 0 });
    s.addText("Source: Total RM industry studies, 2026", {
      x: 8.6, y: 5.85, w: 4.2, h: 0.3, align: "center", fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0 });

    footerDark(s, 8);
  }

  // ============================================================
  // SLIDE 9 — THREE WAYS (comparison)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "THREE WAYS FORWARD");

    s.addText("Three options. One clear winner for us.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.75, fontFace: "Georgia", fontSize: 30, bold: true, color: C.ink, margin: 0 });

    const opts = [
      { tag: "OPTION A", t: "Buy enterprise BI", sub: "Duetto ScoreBoard, Avalon, Lighthouse",
        pros: ["Fast to switch on", "Vendor support + ML"],
        cons: ["6-figure license / year", "Data stays the vendor's", "Can't build custom KPIs"],
        verdict: "Powerful, but rented", win: false },
      { tag: "OPTION B", t: "Stay as we are", sub: "Tools + orphaned Power BI",
        pros: ["No new spend", "Team already familiar"],
        cons: ["Manual reconciliation forever", "Numbers keep conflicting", "No path to TRevPAR"],
        verdict: "Cheapest — until it breaks", win: false },
      { tag: "OPTION C", t: "Build our own layer", sub: "Owned database + automation",
        pros: ["We own the data + history", "One source of truth", "Custom metrics, TRevPAR-ready", "Fraction of license cost"],
        cons: ["We build & maintain it"],
        verdict: "Owned, flexible, affordable", win: true },
    ];
    const cW = 4.0, cGap = 0.18, sx = (W - (3 * cW + 2 * cGap)) / 2;
    for (let i = 0; i < opts.length; i++) {
      const x = sx + i * (cW + cGap);
      const o = opts[i];
      const cardH = 4.95;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: cW, h: cardH, fill: { color: C.card },
        line: { color: o.win ? C.gold : C.ruleLt, width: o.win ? 2.5 : 0.75 }, shadow: shadowCard() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: cW, h: 1.0, fill: { color: o.win ? C.bgDark : "EEF2F7" }, line: { type: "none" } });
      s.addText(o.tag, { x: x + 0.3, y: 1.82, w: cW - 0.6, h: 0.3, fontFace: "Calibri", fontSize: 10, bold: true, color: o.win ? C.gold : C.inkMute, charSpacing: 4, margin: 0 });
      s.addText(o.t, { x: x + 0.3, y: 2.08, w: cW - 0.6, h: 0.4, fontFace: "Georgia", fontSize: 19, bold: true, color: o.win ? C.textLight : C.ink, margin: 0 });
      s.addText(o.sub, { x: x + 0.3, y: 2.46, w: cW - 0.6, h: 0.25, fontFace: "Calibri", fontSize: 10, italic: true, color: o.win ? C.textDim : C.inkSoft, margin: 0 });

      const pros = o.pros.map((p, idx) => ({ text: p, options: { bullet: { code: "2713" }, color: "0F766E", breakLine: true, paraSpaceAfter: 3 } }));
      const cons = o.cons.map((cc, idx) => ({ text: cc, options: { bullet: { code: "2715" }, color: "B91C1C", breakLine: idx < o.cons.length - 1, paraSpaceAfter: 3 } }));
      s.addText([...pros, ...cons], { x: x + 0.35, y: 2.95, w: cW - 0.65, h: 2.45, fontFace: "Calibri", fontSize: 11.5, color: C.ink, margin: 0 });

      // verdict bar
      s.addShape(pres.shapes.RECTANGLE, { x, y: 5.95, w: cW, h: 0.7, fill: { color: o.win ? C.gold : "EEF2F7" }, line: { type: "none" } });
      s.addText(o.verdict, { x: x + 0.2, y: 5.95, w: cW - 0.4, h: 0.7, align: "center", valign: "middle", fontFace: "Calibri", fontSize: 12.5, bold: true, italic: true, color: o.win ? C.bgDark : C.inkSoft, margin: 0 });
    }
    footer(s, 9);
  }

  // ============================================================
  // SLIDE 10 — COST REALITY
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "THE COST REALITY");

    s.addText("Enterprise insight, at a fraction of enterprise cost.", {
      x: 0.5, y: 0.95, w: 12.3, h: 0.85, fontFace: "Georgia", fontSize: 30, bold: true, color: C.ink, margin: 0 });

    // Two big cards
    s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.2, w: 5.6, h: 3.6, fill: { color: C.card }, line: { color: "FECACA", width: 1.5 }, shadow: shadowCard() });
    s.addText("RENT IT  (Enterprise BI)", { x: 1.1, y: 2.45, w: 5.0, h: 0.4, fontFace: "Calibri", fontSize: 13, bold: true, color: C.coral, charSpacing: 3, margin: 0 });
    s.addText("฿100K+", { x: 1.1, y: 2.95, w: 5.0, h: 1.0, fontFace: "Georgia", fontSize: 54, bold: true, color: C.ink, margin: 0 });
    s.addText("per year, every year", { x: 1.1, y: 4.0, w: 5.0, h: 0.4, fontFace: "Calibri", fontSize: 14, italic: true, color: C.inkSoft, margin: 0 });
    const rentCons = ["License renews forever", "Data stays with vendor", "Stop paying — lose access"];
    s.addText(rentCons.map((c, i) => ({ text: c, options: { bullet: { code: "2715" }, color: "B91C1C", breakLine: i < rentCons.length - 1, paraSpaceAfter: 5 } })), {
      x: 1.2, y: 4.55, w: 5.0, h: 1.1, fontFace: "Calibri", fontSize: 13, color: C.ink, margin: 0 });

    s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 2.2, w: 5.6, h: 3.6, fill: { color: C.bgDark }, line: { color: C.gold, width: 2 }, shadow: shadowSoft() });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.9, y: 2.2, w: 5.6, h: 0.07, fill: { color: C.gold }, line: { type: "none" } });
    s.addText("OWN IT  (Our database)", { x: 7.2, y: 2.45, w: 5.0, h: 0.4, fontFace: "Calibri", fontSize: 13, bold: true, color: C.gold, charSpacing: 3, margin: 0 });
    s.addText("~฿1,500", { x: 7.2, y: 2.95, w: 5.0, h: 1.0, fontFace: "Georgia", fontSize: 54, bold: true, color: C.gold, margin: 0 });
    s.addText("per month — infrastructure only", { x: 7.2, y: 4.0, w: 5.0, h: 0.4, fontFace: "Calibri", fontSize: 14, italic: true, color: C.textDim, margin: 0 });
    const ownPros = ["The data asset is ours", "Keeps every year of history", "Reuses tools we already pay for"];
    s.addText(ownPros.map((c, i) => ({ text: c, options: { bullet: { code: "2713" }, color: "5EEAD4", breakLine: i < ownPros.length - 1, paraSpaceAfter: 5 } })), {
      x: 7.3, y: 4.55, w: 5.0, h: 1.1, fontFace: "Calibri", fontSize: 13, color: C.textLight, margin: 0 });

    s.addText("Roughly the cost of one team lunch a month — for an asset the company keeps forever.", {
      x: 0.5, y: 6.1, w: 12.3, h: 0.5, align: "center", fontFace: "Georgia", fontSize: 16, italic: true, bold: true, color: C.ink, margin: 0 });

    footer(s, 10);
  }

  // ============================================================
  // SLIDE 11 — WHAT WE BUILD (architecture)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.5, "WHAT WE BUILD");

    s.addText("Simple by design. Built on what we already have.", {
      x: 0.5, y: 0.95, w: 12.3, h: 0.8, fontFace: "Georgia", fontSize: 29, bold: true, color: C.ink, margin: 0 });

    // Layer 1: sources
    s.addText("YOUR TOOLS (unchanged)", { x: 0.5, y: 2.0, w: 12.3, h: 0.3, align: "center", fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkMute, charSpacing: 4, margin: 0 });
    const srcs = ["Duetto", "SiteMinder", "Comanche", "Vendor email", "OTA reports"];
    const srcW = 2.2, srcGap = 0.2, srcStartX = (W - (5 * srcW + 4 * srcGap)) / 2;
    srcs.forEach((t, i) => {
      const x = srcStartX + i * (srcW + srcGap);
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.35, w: srcW, h: 0.7, fill: { color: C.card }, line: { color: C.ruleLt, width: 1 }, shadow: shadowCard() });
      s.addText(t, { x, y: 2.35, w: srcW, h: 0.7, align: "center", valign: "middle", fontFace: "Calibri", fontSize: 13, bold: true, color: C.ink, margin: 0 });
    });

    // arrows down
    s.addImage({ data: await icon(FaArrowRight, "#94A3B8", 256), x: W/2 - 0.2, y: 3.2, w: 0.4, h: 0.4, rotate: 90 });

    // Layer 2: the 3 components
    const comp = [
      { ic: FaCogs,     t: "Automation", d: "Pulls every source on schedule. Retries on failure. Alerts the team." , sub: "(Airflow)" },
      { ic: FaDatabase, t: "The Database", d: "One cloud warehouse. Every source, reconciled. History kept forever.", sub: "(Cloud Warehouse)" },
      { ic: FaSitemap,  t: "Logic Layer", d: "Business rules — net revenue, occupancy, TRevPAR — versioned & auditable.", sub: "(dbt)" },
    ];
    const compW = 3.95, compGap = 0.2, compStartX = (W - (3 * compW + 2 * compGap)) / 2;
    for (let i = 0; i < comp.length; i++) {
      const x = compStartX + i * (compW + compGap);
      const c = comp[i];
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.85, w: compW, h: 1.95, fill: { color: C.bgDark }, line: { type: "none" }, shadow: shadowSoft() });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.85, w: compW, h: 0.06, fill: { color: C.gold }, line: { type: "none" } });
      s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: 4.15, w: 0.65, h: 0.65, fill: { color: C.bgMid }, line: { color: C.gold, width: 1.5 } });
      s.addImage({ data: await icon(c.ic, "#D4A437", 256), x: x + 0.46, y: 4.26, w: 0.43, h: 0.43 });
      s.addText(c.t, { x: x + 1.15, y: 4.12, w: compW - 1.3, h: 0.4, fontFace: "Georgia", fontSize: 17, bold: true, color: C.textLight, margin: 0 });
      s.addText(c.sub, { x: x + 1.15, y: 4.5, w: compW - 1.3, h: 0.3, fontFace: "Calibri", fontSize: 10, italic: true, color: C.gold, margin: 0 });
      s.addText(c.d, { x: x + 0.35, y: 4.9, w: compW - 0.6, h: 0.85, fontFace: "Calibri", fontSize: 11.5, color: C.textDim, margin: 0 });
    }

    // arrow down
    s.addImage({ data: await icon(FaArrowRight, "#94A3B8", 256), x: W/2 - 0.2, y: 5.95, w: 0.4, h: 0.4, rotate: 90 });

    // Layer 3: consumers
    s.addText("OPEN IT ANYWHERE:   Power BI  ·  Looker  ·  Excel  ·  Mobile  ·  API", {
      x: 0.5, y: 6.45, w: 12.3, h: 0.4, align: "center", fontFace: "Calibri", fontSize: 14, bold: true, color: C.ink, charSpacing: 2, margin: 0 });

    footer(s, 11);
  }

  // ============================================================
  // SLIDE 12 — THE ASK
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.3, h: H, fill: { color: C.gold }, line: { type: "none" } });
    goldChip(s, 0.7, 0.5, "THE ASK");

    s.addText("Approve the foundation.", {
      x: 0.7, y: 0.95, w: 11.6, h: 1.0, fontFace: "Georgia", fontSize: 40, bold: true, color: C.textLight, margin: 0 });
    s.addText("A small, owned investment — and the company stops renting its truth.", {
      x: 0.7, y: 1.95, w: 11.6, h: 0.5, fontFace: "Calibri", fontSize: 17, italic: true, color: C.textDim, margin: 0 });

    const nums = [
      { v: "~฿1,500", l: "PER MONTH", sub: "infrastructure only" },
      { v: "6–8", l: "WEEKS", sub: "to production" },
      { v: "Forever", l: "WE OWN IT", sub: "data + history stay ours" },
    ];
    const nW = 3.6, nGap = 0.35, nStartX = (W - (3 * nW + 2 * nGap)) / 2;
    nums.forEach((n, i) => {
      const x = nStartX + i * (nW + nGap);
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.75, w: nW, h: 1.85, fill: { color: C.bgMid }, line: { color: C.gold, width: 1 }, shadow: shadowSoft() });
      s.addText(n.v, { x, y: 2.85, w: nW, h: 0.95, align: "center", fontFace: "Georgia", fontSize: 44, bold: true, color: C.gold, margin: 0 });
      s.addText(n.l, { x, y: 3.8, w: nW, h: 0.35, align: "center", fontFace: "Calibri", fontSize: 12, bold: true, color: C.textLight, charSpacing: 4, margin: 0 });
      s.addText(n.sub, { x, y: 4.15, w: nW, h: 0.35, align: "center", fontFace: "Calibri", fontSize: 11, italic: true, color: C.textDim, margin: 0 });
    });

    // what we need
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 4.95, w: 11.9, h: 1.5, fill: { color: C.bgMid }, line: { type: "none" } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 4.95, w: 0.1, h: 1.5, fill: { color: C.gold }, line: { type: "none" } });
    s.addText("WHAT WE NEED TO START", { x: 1.0, y: 5.05, w: 11.5, h: 0.35, fontFace: "Calibri", fontSize: 11, bold: true, color: C.gold, charSpacing: 5, margin: 0 });
    s.addText([
      { text: "Approval of the monthly infrastructure  ·  API access to Duetto, SiteMinder & Comanche  ·  one team member for hand-off.",
        options: { color: C.textLight, bold: true, breakLine: true } },
      { text: "Implementation and rollout are already scoped end-to-end.", options: { color: C.textDim, italic: true } },
    ], { x: 1.0, y: 5.45, w: 11.4, h: 0.9, fontFace: "Calibri", fontSize: 14.5, margin: 0 });

    s.addText("Stop renting the truth. Start owning it.", {
      x: 0.7, y: 6.6, w: 11.9, h: 0.4, align: "center", fontFace: "Georgia", fontSize: 17, italic: true, bold: true, color: C.gold, margin: 0 });

    s.addText(`${TOTAL} / ${TOTAL}`, { x: W - 1.5, y: H - 0.32, w: 1, h: 0.28, align: "right", fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0 });
  }

  await pres.writeFile({ fileName: "Own_Your_Data_Decision_Brief.pptx" });
  console.log("Built: Own_Your_Data_Decision_Brief.pptx");
}

build().catch(err => { console.error(err); process.exit(1); });
