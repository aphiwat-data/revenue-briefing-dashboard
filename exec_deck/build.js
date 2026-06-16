// Executive Pitch Deck — Hotel Revenue Intelligence Platform
// Hard-sell: Cloud + Airflow as MUST-HAVE for the team.

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icons
const {
  FaCloud, FaBolt, FaBell, FaShieldAlt, FaChartLine, FaMobileAlt,
  FaUsers, FaCheckCircle, FaTimesCircle, FaExclamationTriangle,
  FaServer, FaLaptopCode, FaSyncAlt, FaCogs, FaDatabase, FaRocket,
  FaEye, FaClock, FaHotel, FaDollarSign, FaCalendarCheck, FaArrowRight,
  FaLock, FaInfinity, FaSearchDollar, FaPlay, FaHistory, FaEnvelope,
} = require("react-icons/fa");

// ============================================================
// PALETTE — "Midnight Hospitality" — premium, executive, hotel
// ============================================================
const C = {
  bgDark:   "0A1428",   // deep midnight
  bgMid:    "13243F",   // navy panel
  bgLight:  "F7F8FC",   // soft white
  card:     "FFFFFF",
  ink:      "0A1428",   // primary text on light
  inkSoft: "475569",
  inkMute: "94A3B8",
  textLight: "F7F8FC",  // text on dark
  textDim:   "B8C5D6",
  gold:     "D4A437",   // accent — premium hospitality gold
  goldDim:  "8C6B14",
  teal:     "14B8A6",   // success / positive
  coral:    "F97066",   // problem / urgency
  amber:    "F59E0B",   // warn
  divider:  "1E3A5F",
  chipBg:   "1B2D4D",
};

// ============================================================
// ICON HELPER
// ============================================================
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

// ============================================================
// PRESENTATION
// ============================================================
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";  // 13.3" x 7.5"
  pres.author = "Revenue Intelligence Team";
  pres.title  = "Cloud + Airflow Migration Proposal";

  const W = 13.3, H = 7.5;

  // ---------- helpers ----------
  const shadowSoft = () => ({ type: "outer", color: "000000", blur: 18, offset: 4, angle: 90, opacity: 0.18 });
  const shadowCard = () => ({ type: "outer", color: "0A1428", blur: 12, offset: 3, angle: 90, opacity: 0.10 });
  const goldChip = (slide, x, y, label) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.18, h: 0.32, fill: { color: C.gold }, line: { type: "none" },
    });
    slide.addText(label, {
      x: x + 0.28, y: y - 0.02, w: 10, h: 0.36, margin: 0,
      fontFace: "Calibri", fontSize: 12, bold: true, color: C.gold, charSpacing: 4,
    });
  };
  const footer = (slide, n) => {
    slide.addText("HOTEL REVENUE INTELLIGENCE  |  CLOUD MIGRATION PROPOSAL", {
      x: 0.5, y: H - 0.32, w: 9, h: 0.28, fontFace: "Calibri", fontSize: 9,
      color: C.inkMute, charSpacing: 4, margin: 0,
    });
    slide.addText(`${n} / 12`, {
      x: W - 1.5, y: H - 0.32, w: 1, h: 0.28, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0,
    });
  };

  // ============================================================
  // SLIDE 1 — TITLE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };

    // Gold corner accent
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.25, h: H, fill: { color: C.gold }, line: { type: "none" },
    });
    // Faint diagonal panel
    s.addShape(pres.shapes.RECTANGLE, {
      x: W - 5.2, y: 0, w: 5.2, h: H, fill: { color: C.bgMid, transparency: 35 }, line: { type: "none" },
    });

    // Eyebrow
    s.addText("PROPOSAL  ·  REVENUE INTELLIGENCE PLATFORM", {
      x: 1, y: 1.7, w: 10, h: 0.4, fontFace: "Calibri", fontSize: 13,
      bold: true, color: C.gold, charSpacing: 8, margin: 0,
    });

    // Title
    s.addText("From One Laptop to\nAlways-On Intelligence", {
      x: 1, y: 2.25, w: 11, h: 2.4, fontFace: "Georgia", fontSize: 54,
      bold: true, color: C.textLight, margin: 0, lineSpacingMultiple: 1.05,
    });

    // Subtitle
    s.addText("Scaling the Hotel Revenue Pipeline to the Cloud with Airflow", {
      x: 1, y: 4.8, w: 11, h: 0.5, fontFace: "Calibri", fontSize: 22,
      color: C.textDim, italic: true, margin: 0,
    });

    // Bottom strip
    s.addShape(pres.shapes.LINE, {
      x: 1, y: 6.05, w: 4, h: 0, line: { color: C.gold, width: 2 },
    });
    s.addText("Revenue Intelligence Team", {
      x: 1, y: 6.15, w: 8, h: 0.4, fontFace: "Calibri", fontSize: 14,
      color: C.textLight, bold: true, margin: 0,
    });
    s.addText("Executive Review  ·  2026", {
      x: 1, y: 6.55, w: 8, h: 0.35, fontFace: "Calibri", fontSize: 12,
      color: C.inkMute, margin: 0, charSpacing: 2,
    });
  }

  // ============================================================
  // SLIDE 2 — WHAT WE'VE BUILT (success foundation)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "WHERE WE STAND TODAY");

    s.addText("We turned 7 hotels of scattered PDFs into\none source of truth.", {
      x: 0.5, y: 0.85, w: 12.3, h: 1.4, fontFace: "Georgia",
      fontSize: 32, bold: true, color: C.ink, margin: 0, lineSpacingMultiple: 1.1,
    });

    // KPI stat row
    const stats = [
      { v: "7",        l: "HOTELS LIVE",              c: C.bgDark },
      { v: "53,740",   l: "BOOKINGS TRACKED",         c: C.bgDark },
      { v: "฿340M",    l: "REVENUE UNDER MANAGEMENT", c: C.gold   },
      { v: "฿50.9M",   l: "OTA COMMISSION VISIBLE",   c: C.bgDark },
      { v: "100%",     l: "COMMISSION COVERAGE",      c: C.teal   },
    ];
    const sW = 2.4, gap = 0.18, totalW = stats.length * sW + (stats.length - 1) * gap;
    const startX = (W - totalW) / 2;
    stats.forEach((st, i) => {
      const x = startX + i * (sW + gap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.55, w: sW, h: 1.6, fill: { color: C.card },
        line: { color: "E2E8F0", width: 0.75 }, shadow: shadowCard(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.55, w: sW, h: 0.08, fill: { color: st.c }, line: { type: "none" },
      });
      s.addText(st.v, {
        x, y: 2.78, w: sW, h: 0.8, align: "center", fontFace: "Georgia",
        fontSize: 30, bold: true, color: st.c, margin: 0,
      });
      s.addText(st.l, {
        x: x + 0.1, y: 3.7, w: sW - 0.2, h: 0.4, align: "center",
        fontFace: "Calibri", fontSize: 9, bold: true, color: C.inkSoft,
        charSpacing: 3, margin: 0,
      });
    });

    // "What works" capability strip
    const caps = [
      { ic: await icon(FaDatabase,    "#" + C.bgDark, 256), t: "Automated PDF parsing",    d: "RS02 · RS03 · ST14 — daily" },
      { ic: await icon(FaChartLine,   "#" + C.bgDark, 256), t: "Layered analytics (dbt)", d: "Staging → Intermediate → Mart" },
      { ic: await icon(FaSearchDollar,"#" + C.bgDark, 256), t: "Net revenue + commission", d: "Per OTA, per hotel, daily" },
      { ic: await icon(FaCalendarCheck,"#" + C.bgDark,256), t: "Daily snapshot model",     d: "Booking pace and pickup tracked" },
    ];
    const cW = 2.95, cGap = 0.18, cTotalW = caps.length * cW + (caps.length - 1) * cGap;
    const cStartX = (W - cTotalW) / 2;
    caps.forEach((c, i) => {
      const x = cStartX + i * (cW + cGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 4.55, w: cW, h: 1.6, fill: { color: C.bgDark }, line: { type: "none" },
      });
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.25, y: 4.78, w: 0.55, h: 0.55, fill: { color: C.gold }, line: { type: "none" },
      });
      s.addImage({ data: c.ic, x: x + 0.34, y: 4.87, w: 0.37, h: 0.37 });
      s.addText(c.t, {
        x: x + 0.95, y: 4.75, w: cW - 1.1, h: 0.45, fontFace: "Calibri",
        fontSize: 14, bold: true, color: C.textLight, margin: 0,
      });
      s.addText(c.d, {
        x: x + 0.95, y: 5.18, w: cW - 1.1, h: 0.45, fontFace: "Calibri",
        fontSize: 11, color: C.textDim, margin: 0,
      });
      s.addShape(pres.shapes.LINE, {
        x: x + 0.95, y: 5.65, w: 1, h: 0,
        line: { color: C.gold, width: 1.5 },
      });
      s.addText("PROVEN · IN PRODUCTION", {
        x: x + 0.95, y: 5.72, w: cW - 1.1, h: 0.3, fontFace: "Calibri",
        fontSize: 8.5, bold: true, color: C.gold, charSpacing: 3, margin: 0,
      });
    });

    s.addText("The foundation is real. The question now is reach.", {
      x: 0.5, y: 6.45, w: 12.3, h: 0.5, align: "center",
      fontFace: "Georgia", fontSize: 18, italic: true, color: C.inkSoft, margin: 0,
    });

    footer(s, 2);
  }

  // ============================================================
  // SLIDE 3 — THE BOTTLENECK (problem statement)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "THE BOTTLENECK");

    s.addText("All that intelligence lives on one laptop.", {
      x: 0.5, y: 0.85, w: 12.3, h: 1.2, fontFace: "Georgia",
      fontSize: 32, bold: true, color: C.ink, margin: 0,
    });
    s.addText("Every morning, one machine must be awake, online, and unlocked — or the team flies blind.", {
      x: 0.5, y: 1.85, w: 12.3, h: 0.5, fontFace: "Calibri",
      fontSize: 16, color: C.inkSoft, italic: true, margin: 0,
    });

    // Center "single point of failure" panel
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.15, y: 2.7, w: 3, h: 2.4, fill: { color: C.bgDark }, line: { type: "none" }, shadow: shadowSoft(),
    });
    s.addShape(pres.shapes.OVAL, {
      x: 6.25, y: 2.85, w: 0.8, h: 0.8, fill: { color: C.coral }, line: { type: "none" },
    });
    s.addImage({ data: await icon(FaLaptopCode, "#FFFFFF", 256), x: 6.4, y: 3.0, w: 0.5, h: 0.5 });
    s.addText("ONE LAPTOP", {
      x: 5.15, y: 3.78, w: 3, h: 0.4, align: "center",
      fontFace: "Calibri", fontSize: 14, bold: true, color: C.gold, charSpacing: 5, margin: 0,
    });
    s.addText("Single point of failure", {
      x: 5.15, y: 4.2, w: 3, h: 0.4, align: "center",
      fontFace: "Georgia", fontSize: 16, bold: true, color: C.textLight, margin: 0,
    });
    s.addText("for the entire revenue function", {
      x: 5.15, y: 4.6, w: 3, h: 0.4, align: "center",
      fontFace: "Calibri", fontSize: 12, color: C.textDim, italic: true, margin: 0,
    });

    // 4 pain points around it
    const pains = [
      { x: 0.5,  y: 2.55, ic: await icon(FaTimesCircle, "#F97066", 256),
        t: "If the laptop sleeps, the pipeline dies",
        d: "No team member can recover it remotely. Numbers go stale.",
      },
      { x: 9.0,  y: 2.55, ic: await icon(FaExclamationTriangle, "#F97066", 256),
        t: "No remote access for revenue managers",
        d: "Screenshots over chat. No live drill-down. No mobile.",
      },
      { x: 0.5,  y: 4.55, ic: await icon(FaClock, "#F97066", 256),
        t: "Manual run every morning",
        d: "Someone clicks a .bat file. If they forget — no report.",
      },
      { x: 9.0,  y: 4.55, ic: await icon(FaHistory, "#F97066", 256),
        t: "No alerts, no audit trail",
        d: "Failures are silent. Yesterday's run isn't reproducible.",
      },
    ];
    pains.forEach(p => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: p.x, y: p.y, w: 3.8, h: 1.85, fill: { color: C.card },
        line: { color: "FECACA", width: 1 }, shadow: shadowCard(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: p.x, y: p.y, w: 0.1, h: 1.85, fill: { color: C.coral }, line: { type: "none" },
      });
      s.addImage({ data: p.ic, x: p.x + 0.3, y: p.y + 0.22, w: 0.45, h: 0.45 });
      s.addText(p.t, {
        x: p.x + 0.9, y: p.y + 0.18, w: 2.8, h: 0.55, fontFace: "Calibri",
        fontSize: 13.5, bold: true, color: C.ink, margin: 0,
      });
      s.addText(p.d, {
        x: p.x + 0.3, y: p.y + 0.9, w: 3.3, h: 0.85, fontFace: "Calibri",
        fontSize: 11.5, color: C.inkSoft, margin: 0,
      });
    });

    s.addText("Right now, our revenue intelligence is one cup of coffee away from disaster.", {
      x: 0.5, y: 6.55, w: 12.3, h: 0.4, align: "center",
      fontFace: "Georgia", fontSize: 15, italic: true, color: C.coral, bold: true, margin: 0,
    });

    footer(s, 3);
  }

  // ============================================================
  // SLIDE 4 — THE HIDDEN COST (urgency)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "THE HIDDEN COST");

    s.addText("What 'good enough today' actually costs the team.", {
      x: 0.5, y: 0.85, w: 12.3, h: 1.0, fontFace: "Georgia",
      fontSize: 30, bold: true, color: C.ink, margin: 0,
    });

    // Cost rows — table-like
    const rows = [
      { area: "Revenue meetings",   today: "Wait for screenshots from one person",   future: "Everyone opens the same live view",  ic: FaUsers },
      { area: "Out-of-hours alerts", today: "Nobody knows pickup dropped overnight", future: "Auto email if pickup < threshold",   ic: FaBell  },
      { area: "Pipeline failures",  today: "Discovered next morning, manually",      future: "Slack ping within 60 seconds",        ic: FaShieldAlt },
      { area: "Scaling to 10 hotels", today: "Same one laptop. Same one person.",    future: "Same architecture. Zero extra effort.", ic: FaHotel },
      { area: "Backfilling history",today: "Run 30 .bat files by hand",              future: "One click. 30 days re-processed.",   ic: FaSyncAlt },
    ];
    const rowH = 0.7, startY = 2.0;
    // Header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: startY, w: W - 1.0, h: 0.5, fill: { color: C.bgDark }, line: { type: "none" },
    });
    s.addText("BUSINESS AREA", {
      x: 0.75, y: startY, w: 2.7, h: 0.5, fontFace: "Calibri", fontSize: 11,
      bold: true, color: C.gold, charSpacing: 4, valign: "middle", margin: 0,
    });
    s.addText("TODAY", {
      x: 3.6, y: startY, w: 4.5, h: 0.5, fontFace: "Calibri", fontSize: 11,
      bold: true, color: C.coral, charSpacing: 4, valign: "middle", margin: 0,
    });
    s.addText("AFTER MIGRATION", {
      x: 8.3, y: startY, w: 4.5, h: 0.5, fontFace: "Calibri", fontSize: 11,
      bold: true, color: C.teal, charSpacing: 4, valign: "middle", margin: 0,
    });

    for (let i = 0; i < rows.length; i++) {
      const y = startY + 0.5 + i * (rowH + 0.08);
      const bg = i % 2 === 0 ? C.card : "EEF2F7";
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y, w: W - 1.0, h: rowH, fill: { color: bg }, line: { color: "E2E8F0", width: 0.5 },
      });
      // icon
      s.addShape(pres.shapes.OVAL, {
        x: 0.7, y: y + 0.16, w: 0.45, h: 0.45, fill: { color: C.bgDark }, line: { type: "none" },
      });
      s.addImage({ data: await icon(rows[i].ic, "#D4A437", 256), x: 0.78, y: y + 0.24, w: 0.3, h: 0.3 });
      s.addText(rows[i].area, {
        x: 1.3, y, w: 2.2, h: rowH, fontFace: "Calibri", fontSize: 13.5, bold: true,
        color: C.ink, valign: "middle", margin: 0,
      });
      // today
      s.addImage({ data: await icon(FaTimesCircle, "#F97066", 256), x: 3.6, y: y + 0.27, w: 0.28, h: 0.28 });
      s.addText(rows[i].today, {
        x: 3.95, y, w: 4.2, h: rowH, fontFace: "Calibri", fontSize: 12,
        color: C.inkSoft, valign: "middle", margin: 0,
      });
      // future
      s.addImage({ data: await icon(FaCheckCircle, "#14B8A6", 256), x: 8.3, y: y + 0.27, w: 0.28, h: 0.28 });
      s.addText(rows[i].future, {
        x: 8.65, y, w: 4.2, h: rowH, fontFace: "Calibri", fontSize: 12, bold: true,
        color: C.ink, valign: "middle", margin: 0,
      });
    }

    s.addText("Every day on the old setup is a day the team is one click away from a blackout.", {
      x: 0.5, y: 6.6, w: 12.3, h: 0.4, align: "center",
      fontFace: "Georgia", fontSize: 14, italic: true, color: C.inkSoft, margin: 0,
    });

    footer(s, 4);
  }

  // ============================================================
  // SLIDE 5 — THE VISION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    goldChip(s, 0.5, 0.45, "THE VISION");

    s.addText("One platform. Every hotel. Always on.", {
      x: 0.5, y: 0.95, w: 12.3, h: 1.3, fontFace: "Georgia",
      fontSize: 40, bold: true, color: C.textLight, margin: 0,
    });
    s.addText("Revenue managers, GMs, and executives — same numbers, same minute, anywhere.", {
      x: 0.5, y: 2.05, w: 12.3, h: 0.5, fontFace: "Calibri",
      fontSize: 17, color: C.textDim, italic: true, margin: 0,
    });

    // 3 big pillars
    const pillars = [
      { ic: await icon(FaInfinity, "#D4A437", 256),
        t: "Always On",
        l: "No laptop dependency",
        body: "Pipeline runs in the cloud, on a schedule. If our office loses power, the data still updates." },
      { ic: await icon(FaMobileAlt, "#D4A437", 256),
        t: "Always Accessible",
        l: "Phone, tablet, anywhere",
        body: "Any authorized user opens a URL. No VPN. No screenshots. Same numbers everywhere." },
      { ic: await icon(FaShieldAlt, "#D4A437", 256),
        t: "Always Accountable",
        l: "Every run logged, every error alerted",
        body: "Failures ping a channel within 60 seconds. Every metric is reproducible." },
    ];
    const pW = 3.9, pGap = 0.3, pTotalW = pillars.length * pW + (pillars.length - 1) * pGap;
    const pStartX = (W - pTotalW) / 2;
    pillars.forEach((p, i) => {
      const x = pStartX + i * (pW + pGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 3.0, w: pW, h: 3.5, fill: { color: C.bgMid }, line: { color: C.divider, width: 1 },
        shadow: shadowSoft(),
      });
      // gold top stripe
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 3.0, w: pW, h: 0.06, fill: { color: C.gold }, line: { type: "none" },
      });
      // icon circle
      s.addShape(pres.shapes.OVAL, {
        x: x + pW/2 - 0.5, y: 3.3, w: 1.0, h: 1.0, fill: { color: C.bgDark },
        line: { color: C.gold, width: 2 },
      });
      s.addImage({ data: p.ic, x: x + pW/2 - 0.3, y: 3.5, w: 0.6, h: 0.6 });
      s.addText(p.t, {
        x, y: 4.45, w: pW, h: 0.55, align: "center",
        fontFace: "Georgia", fontSize: 24, bold: true, color: C.textLight, margin: 0,
      });
      s.addText(p.l, {
        x, y: 5.0, w: pW, h: 0.4, align: "center",
        fontFace: "Calibri", fontSize: 12, bold: true, color: C.gold, charSpacing: 3, margin: 0,
      });
      s.addText(p.body, {
        x: x + 0.4, y: 5.45, w: pW - 0.8, h: 1.0, align: "center",
        fontFace: "Calibri", fontSize: 12.5, color: C.textDim, margin: 0,
      });
    });

    s.addText("This isn't a 'nice to have.' It's the difference between a side project and a business platform.", {
      x: 0.5, y: 6.7, w: 12.3, h: 0.4, align: "center",
      fontFace: "Georgia", fontSize: 14, italic: true, color: C.gold, bold: true, margin: 0,
    });

    footer(s, 5);
  }

  // ============================================================
  // SLIDE 6 — SOLUTION PART 1: CLOUD DATA WAREHOUSE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "PILLAR 1  ·  CLOUD DATA WAREHOUSE");

    s.addText("Move the warehouse to the cloud.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.9, fontFace: "Georgia",
      fontSize: 32, bold: true, color: C.ink, margin: 0,
    });
    s.addText("Same SQL. Same dbt models. New superpower: every authorized user shares one live database.", {
      x: 0.5, y: 1.75, w: 12.3, h: 0.5, fontFace: "Calibri",
      fontSize: 16, color: C.inkSoft, italic: true, margin: 0,
    });

    // Left: big cloud icon panel
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.55, w: 4.6, h: 4.4, fill: { color: C.bgDark }, line: { type: "none" }, shadow: shadowSoft(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.55, w: 4.6, h: 0.08, fill: { color: C.gold }, line: { type: "none" },
    });
    s.addImage({ data: await icon(FaCloud, "#D4A437", 256), x: 1.95, y: 2.85, w: 1.7, h: 1.7 });
    s.addText("Cloud Warehouse", {
      x: 0.5, y: 4.7, w: 4.6, h: 0.5, align: "center",
      fontFace: "Georgia", fontSize: 22, bold: true, color: C.textLight, margin: 0,
    });
    s.addText("MotherDuck  ·  BigQuery  ·  AWS Athena", {
      x: 0.5, y: 5.2, w: 4.6, h: 0.4, align: "center",
      fontFace: "Calibri", fontSize: 12, bold: true, color: C.gold, charSpacing: 3, margin: 0,
    });
    s.addText("Battle-tested platforms used by\nNetflix, Airbnb, Booking.com,\nand every major hotel group.", {
      x: 0.7, y: 5.7, w: 4.2, h: 1.1, align: "center",
      fontFace: "Calibri", fontSize: 12.5, color: C.textDim, italic: true, margin: 0,
    });

    // Right: 5 benefits
    const benefits = [
      { ic: FaUsers,        t: "Multi-user simultaneous access",
        d: "Revenue manager, GM, and accountant can query at the same time — no more 'close DBeaver, I need to update.'" },
      { ic: FaShieldAlt,    t: "Built-in backup + disaster recovery",
        d: "Automatic snapshots. Six months of history preserved, even if a hotel laptop is stolen." },
      { ic: FaMobileAlt,    t: "Works on phones and tablets",
        d: "Glance at occupancy on the way to a meeting. No software install, just a URL." },
      { ic: FaLock,         t: "Enterprise-grade security",
        d: "Encrypted at rest and in transit. Role-based access — finance sees commission, GM sees occupancy." },
      { ic: FaRocket,       t: "Zero maintenance overhead",
        d: "No server to patch. No disk to clean. The provider handles uptime — typically 99.9%+." },
    ];
    benefits.forEach((b, i) => {
      const y = 2.55 + i * 0.9;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 5.4, y, w: 7.4, h: 0.78, fill: { color: C.card },
        line: { color: "E2E8F0", width: 0.5 }, shadow: shadowCard(),
      });
      s.addShape(pres.shapes.OVAL, {
        x: 5.55, y: y + 0.16, w: 0.46, h: 0.46, fill: { color: C.gold }, line: { type: "none" },
      });
    });
    for (let i = 0; i < benefits.length; i++) {
      const y = 2.55 + i * 0.9;
      const ic = await icon(benefits[i].ic, "#FFFFFF", 256);
      s.addImage({ data: ic, x: 5.62, y: y + 0.23, w: 0.32, h: 0.32 });
      s.addText(benefits[i].t, {
        x: 6.15, y: y + 0.04, w: 6.55, h: 0.36, fontFace: "Calibri",
        fontSize: 13.5, bold: true, color: C.ink, margin: 0,
      });
      s.addText(benefits[i].d, {
        x: 6.15, y: y + 0.38, w: 6.55, h: 0.4, fontFace: "Calibri",
        fontSize: 10.5, color: C.inkSoft, margin: 0,
      });
    }

    footer(s, 6);
  }

  // ============================================================
  // SLIDE 7 — SOLUTION PART 2: AIRFLOW ORCHESTRATION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "PILLAR 2  ·  AIRFLOW ORCHESTRATION");

    s.addText("Stop running the pipeline. Let Airflow run it.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.9, fontFace: "Georgia",
      fontSize: 30, bold: true, color: C.ink, margin: 0,
    });
    s.addText("The industry-standard scheduler used by 80%+ of data teams worldwide.", {
      x: 0.5, y: 1.75, w: 12.3, h: 0.5, fontFace: "Calibri",
      fontSize: 16, color: C.inkSoft, italic: true, margin: 0,
    });

    // Top callout: Airflow at marquee companies
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.4, w: 12.3, h: 0.65, fill: { color: C.bgDark }, line: { type: "none" },
    });
    s.addImage({ data: await icon(FaCheckCircle, "#D4A437", 256), x: 0.7, y: 2.55, w: 0.35, h: 0.35 });
    s.addText("Already running production pipelines at:  ", {
      x: 1.15, y: 2.45, w: 4.5, h: 0.55, fontFace: "Calibri",
      fontSize: 13, color: C.gold, bold: true, valign: "middle", margin: 0,
    });
    s.addText("AIRBNB · NETFLIX · BOOKING.COM · LYFT · ADOBE · WALMART", {
      x: 5.5, y: 2.45, w: 7.2, h: 0.55, fontFace: "Calibri",
      fontSize: 12, bold: true, color: C.textLight, valign: "middle",
      charSpacing: 3, margin: 0,
    });

    // 6 benefit cards (3x2 grid)
    const grid = [
      { ic: FaClock,        t: "Self-running on a schedule",
        d: "Runs every morning at 06:00. Holidays included. Time-zone aware." },
      { ic: FaSyncAlt,      t: "Auto-retry on failure",
        d: "Transient network blip? Airflow retries 3 times with backoff. We sleep through it." },
      { ic: FaBell,         t: "Real-time alerts",
        d: "Slack or email the moment anything breaks. No more 'oh, the data didn't update.'" },
      { ic: FaPlay,         t: "One-click backfill",
        d: "Need to reprocess last month after a fix? One button. Done in minutes." },
      { ic: FaEye,          t: "Visual pipeline map",
        d: "See every step: PDF → parse → load → dbt → mart. Spot the slow link in seconds." },
      { ic: FaHistory,      t: "Full audit trail",
        d: "Every run, every log, every output — kept forever. Compliance-ready." },
    ];
    const gW = 3.95, gH = 1.65, gGapX = 0.18, gGapY = 0.18;
    const gridStartX = (W - (3 * gW + 2 * gGapX)) / 2;
    for (let i = 0; i < grid.length; i++) {
      const col = i % 3, row = Math.floor(i / 3);
      const x = gridStartX + col * (gW + gGapX);
      const y = 3.25 + row * (gH + gGapY);
      const b = grid[i];
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: gW, h: gH, fill: { color: C.card },
        line: { color: "E2E8F0", width: 0.5 }, shadow: shadowCard(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 0.1, h: gH, fill: { color: C.gold }, line: { type: "none" },
      });
      const ic = await icon(b.ic, "#D4A437", 256);
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.32, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: C.bgDark }, line: { type: "none" },
      });
      s.addImage({ data: ic, x: x + 0.4, y: y + 0.38, w: 0.4, h: 0.4 });
      s.addText(b.t, {
        x: x + 1.05, y: y + 0.25, w: gW - 1.2, h: 0.5, fontFace: "Calibri",
        fontSize: 14, bold: true, color: C.ink, margin: 0,
      });
      s.addText(b.d, {
        x: x + 0.32, y: y + 0.95, w: gW - 0.5, h: 0.65, fontFace: "Calibri",
        fontSize: 11.5, color: C.inkSoft, margin: 0,
      });
    }

    s.addText("Airflow turns our pipeline from 'someone has to run this' into 'it just runs.'", {
      x: 0.5, y: 6.78, w: 12.3, h: 0.35, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.gold, bold: true, margin: 0,
    });

    footer(s, 7);
  }

  // ============================================================
  // SLIDE 8 — ARCHITECTURE: TODAY vs TOMORROW
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "ARCHITECTURE");

    s.addText("Today vs. Tomorrow — at a glance.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.9, fontFace: "Georgia",
      fontSize: 30, bold: true, color: C.ink, margin: 0,
    });

    // Two big panels
    // LEFT — TODAY
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.0, w: 6.0, h: 4.5, fill: { color: C.card },
      line: { color: "FECACA", width: 1.5 }, shadow: shadowCard(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 2.0, w: 6.0, h: 0.55, fill: { color: C.coral }, line: { type: "none" },
    });
    s.addText("TODAY", {
      x: 0.5, y: 2.0, w: 6.0, h: 0.55, align: "center",
      fontFace: "Calibri", fontSize: 16, bold: true, color: C.card,
      charSpacing: 6, valign: "middle", margin: 0,
    });

    // Today components — vertical flow
    const todayFlow = [
      { ic: FaLaptopCode,  l: "One laptop",        d: "Manual .bat file every morning" },
      { ic: FaDatabase,    l: "DuckDB on G:\\ drive", d: "File-locked, single-writer" },
      { ic: FaUsers,       l: "Team",              d: "Asks one person for screenshots" },
    ];
    todayFlow.forEach((t, i) => {
      const y = 2.85 + i * 1.35;
      s.addShape(pres.shapes.OVAL, {
        x: 0.85, y, w: 0.7, h: 0.7, fill: { color: C.coral }, line: { type: "none" },
      });
    });
    for (let i = 0; i < todayFlow.length; i++) {
      const y = 2.85 + i * 1.35;
      const t = todayFlow[i];
      const ic = await icon(t.ic, "#FFFFFF", 256);
      s.addImage({ data: ic, x: 0.96, y: y + 0.11, w: 0.48, h: 0.48 });
      s.addText(t.l, {
        x: 1.75, y: y - 0.05, w: 4.5, h: 0.45,
        fontFace: "Calibri", fontSize: 16, bold: true, color: C.ink, margin: 0,
      });
      s.addText(t.d, {
        x: 1.75, y: y + 0.4, w: 4.5, h: 0.4,
        fontFace: "Calibri", fontSize: 12, color: C.inkSoft, margin: 0,
      });
      if (i < todayFlow.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: 1.2, y: y + 0.78, w: 0, h: 0.45, line: { color: C.coral, width: 2 },
        });
      }
    }

    // RIGHT — TOMORROW
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.8, y: 2.0, w: 6.0, h: 4.5, fill: { color: C.card },
      line: { color: "99F6E4", width: 1.5 }, shadow: shadowCard(),
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.8, y: 2.0, w: 6.0, h: 0.55, fill: { color: C.teal }, line: { type: "none" },
    });
    s.addText("TOMORROW", {
      x: 6.8, y: 2.0, w: 6.0, h: 0.55, align: "center",
      fontFace: "Calibri", fontSize: 16, bold: true, color: C.card,
      charSpacing: 6, valign: "middle", margin: 0,
    });

    const futureFlow = [
      { ic: FaCogs,      l: "Airflow (cloud)",     d: "Runs itself daily · alerts on failure" },
      { ic: FaCloud,     l: "Cloud warehouse",     d: "Multi-user · auto-backup · always on" },
      { ic: FaUsers,     l: "Team",                d: "Opens one URL · live data from any device" },
    ];
    futureFlow.forEach((t, i) => {
      const y = 2.85 + i * 1.35;
      s.addShape(pres.shapes.OVAL, {
        x: 7.15, y, w: 0.7, h: 0.7, fill: { color: C.teal }, line: { type: "none" },
      });
    });
    for (let i = 0; i < futureFlow.length; i++) {
      const y = 2.85 + i * 1.35;
      const t = futureFlow[i];
      const ic = await icon(t.ic, "#FFFFFF", 256);
      s.addImage({ data: ic, x: 7.26, y: y + 0.11, w: 0.48, h: 0.48 });
      s.addText(t.l, {
        x: 8.05, y: y - 0.05, w: 4.5, h: 0.45,
        fontFace: "Calibri", fontSize: 16, bold: true, color: C.ink, margin: 0,
      });
      s.addText(t.d, {
        x: 8.05, y: y + 0.4, w: 4.5, h: 0.4,
        fontFace: "Calibri", fontSize: 12, color: C.inkSoft, margin: 0,
      });
      if (i < futureFlow.length - 1) {
        s.addShape(pres.shapes.LINE, {
          x: 7.5, y: y + 0.78, w: 0, h: 0.45, line: { color: C.teal, width: 2 },
        });
      }
    }

    s.addText("Same data. Same business logic. Different reliability class.", {
      x: 0.5, y: 6.78, w: 12.3, h: 0.3, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.inkSoft, margin: 0,
    });

    footer(s, 8);
  }

  // ============================================================
  // SLIDE 9 — BUSINESS IMPACT (specific to team)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "BUSINESS IMPACT");

    s.addText("What each role actually gets.", {
      x: 0.5, y: 0.85, w: 12.3, h: 0.9, fontFace: "Georgia",
      fontSize: 32, bold: true, color: C.ink, margin: 0,
    });

    const roles = [
      { who: "REVENUE MANAGER",
        ic: FaChartLine,
        wins: [
          "Live pickup + occupancy on a phone every morning",
          "Alert when occupancy drops below threshold — react same-day",
          "OTA mix breakdown by hotel, ready for the 10:00 huddle",
        ],
      },
      { who: "GENERAL MANAGER",
        ic: FaHotel,
        wins: [
          "Property-level dashboard — no need to ask anyone",
          "Compare own hotel vs. group benchmark live",
          "Forecast vs. on-book gap visible by stay-date",
        ],
      },
      { who: "FINANCE / ACCOUNTING",
        ic: FaDollarSign,
        wins: [
          "Net revenue after commission — calculated, not estimated",
          "OTA payable reconciliation on-demand, audit-ready",
          "Month-end close: hours instead of days",
        ],
      },
      { who: "EXECUTIVE TEAM",
        ic: FaRocket,
        wins: [
          "Group KPIs without waiting for a refreshed deck",
          "Same numbers in every meeting — one source of truth",
          "Scales to new properties with zero added overhead",
        ],
      },
    ];

    // 2x2 grid
    const rW = 6.05, rH = 2.25, rGap = 0.18;
    const rStartX = (W - (2 * rW + rGap)) / 2;
    for (let i = 0; i < roles.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = rStartX + col * (rW + rGap);
      const y = 1.85 + row * (rH + rGap);
      const r = roles[i];

      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: rW, h: rH, fill: { color: C.card },
        line: { color: "E2E8F0", width: 0.75 }, shadow: shadowCard(),
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 0.12, h: rH, fill: { color: C.gold }, line: { type: "none" },
      });

      // Header band
      s.addShape(pres.shapes.OVAL, {
        x: x + 0.35, y: y + 0.3, w: 0.75, h: 0.75, fill: { color: C.bgDark }, line: { type: "none" },
      });
      const ic = await icon(r.ic, "#D4A437", 256);
      s.addImage({ data: ic, x: x + 0.49, y: y + 0.44, w: 0.47, h: 0.47 });

      s.addText(r.who, {
        x: x + 1.25, y: y + 0.4, w: rW - 1.4, h: 0.45,
        fontFace: "Calibri", fontSize: 15, bold: true, color: C.bgDark,
        charSpacing: 4, margin: 0,
      });
      s.addText("WHAT CHANGES", {
        x: x + 1.25, y: y + 0.82, w: rW - 1.4, h: 0.3,
        fontFace: "Calibri", fontSize: 9, bold: true, color: C.gold,
        charSpacing: 3, margin: 0,
      });

      // Wins list
      const wins = r.wins.map((w, idx) => ({
        text: w,
        options: { bullet: { code: "25A0" }, breakLine: idx < r.wins.length - 1, color: C.ink, paraSpaceAfter: 4 },
      }));
      s.addText(wins, {
        x: x + 0.4, y: y + 1.25, w: rW - 0.6, h: rH - 1.35,
        fontFace: "Calibri", fontSize: 12.5, color: C.ink, margin: 0,
      });
    }

    s.addText("Every team gets faster, sharper, and freer of single-person dependencies.", {
      x: 0.5, y: 6.65, w: 12.3, h: 0.3, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.inkSoft, margin: 0,
    });

    footer(s, 9);
  }

  // ============================================================
  // SLIDE 10 — USE CASES UNLOCKED (concrete scenarios)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };
    goldChip(s, 0.5, 0.45, "USE CASES UNLOCKED");

    s.addText("Scenarios that are impossible today — routine tomorrow.", {
      x: 0.5, y: 0.95, w: 12.3, h: 1.1, fontFace: "Georgia",
      fontSize: 28, bold: true, color: C.textLight, margin: 0,
    });

    const cases = [
      { time: "06:00", title: "Pipeline runs itself",
        d: "Airflow pulls all 7 hotels, parses every PDF, refreshes the warehouse before anyone opens email.",
        ic: FaSyncAlt },
      { time: "07:30", title: "Pickup-drop alert fires",
        d: "Arbour occupancy is forecast to slip below 80% next weekend. Revenue Manager gets a Slack ping with the exact stay-dates.",
        ic: FaBell },
      { time: "09:00", title: "GM huddle on a phone",
        d: "Property GM opens the dashboard on iPad on the way to the lobby. Live occupancy, OTA mix, ADR vs forecast.",
        ic: FaMobileAlt },
      { time: "11:00", title: "Finance pulls payable",
        d: "Accountant queries OTA commission by hotel for the previous month — same warehouse, no laptop, no PDF.",
        ic: FaDollarSign },
      { time: "15:00", title: "Adding the 8th hotel",
        d: "New property drops PDFs in the same folder shape. Zero code change. Tomorrow's dashboard includes it automatically.",
        ic: FaHotel },
      { time: "22:00", title: "Quiet hours — alert fires",
        d: "PDF format change at one property breaks parsing. On-call gets emailed. Fix lands by morning. Team never noticed.",
        ic: FaShieldAlt },
    ];

    // Two columns of 3 timeline items
    const colW = 6.1, colGap = 0.3;
    const colStartX = (W - (2 * colW + colGap)) / 2;
    for (let i = 0; i < cases.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = colStartX + col * (colW + colGap);
      const y = 2.1 + row * 1.45;
      const c = cases[i];

      // Time chip
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: y + 0.1, w: 1.0, h: 0.55, fill: { color: C.gold }, line: { type: "none" },
      });
      s.addText(c.time, {
        x, y: y + 0.1, w: 1.0, h: 0.55, align: "center", valign: "middle",
        fontFace: "Georgia", fontSize: 18, bold: true, color: C.bgDark, margin: 0,
      });

      // Icon
      s.addShape(pres.shapes.OVAL, {
        x: x + 1.2, y: y + 0.1, w: 0.55, h: 0.55, fill: { color: C.bgMid }, line: { color: C.gold, width: 1 },
      });
      const ic = await icon(c.ic, "#D4A437", 256);
      s.addImage({ data: ic, x: x + 1.29, y: y + 0.19, w: 0.37, h: 0.37 });

      // Title
      s.addText(c.title, {
        x: x + 1.95, y: y + 0.05, w: colW - 2, h: 0.5,
        fontFace: "Calibri", fontSize: 15, bold: true, color: C.textLight, margin: 0,
      });

      // Description
      s.addText(c.d, {
        x: x + 1.95, y: y + 0.55, w: colW - 2, h: 0.9,
        fontFace: "Calibri", fontSize: 11.5, color: C.textDim, margin: 0,
      });
    }

    s.addText("A day in the life of an always-on platform.", {
      x: 0.5, y: 6.78, w: 12.3, h: 0.35, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.gold, bold: true, margin: 0,
    });

    footer(s, 10);
  }

  // ============================================================
  // SLIDE 11 — ROADMAP
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgLight };
    goldChip(s, 0.5, 0.45, "ROLLOUT ROADMAP");

    s.addText("Six weeks. Three phases. Production-grade at the end of week 6.", {
      x: 0.5, y: 0.85, w: 12.3, h: 1.0, fontFace: "Georgia",
      fontSize: 26, bold: true, color: C.ink, margin: 0,
    });

    // 3 phases as horizontal cards with progress
    const phases = [
      { num: "01", title: "CLOUD WAREHOUSE",  weeks: "Week 1–2",
        body: "Lift DuckDB to the cloud platform. Re-point dbt. Team starts using shared URL.",
        deliverables: ["Cloud warehouse live", "All historical data migrated", "Team access provisioned"],
      },
      { num: "02", title: "AIRFLOW ON CLOUD",  weeks: "Week 3–4",
        body: "Deploy Airflow. Convert daily batch to scheduled DAG. Wire failure alerts to Slack.",
        deliverables: ["DAGs running daily at 06:00", "Slack alerts on failure", "One-click backfill enabled"],
      },
      { num: "03", title: "DASHBOARDS + ALERTS",  weeks: "Week 5–6",
        body: "Build role-specific dashboards. Wire automated KPI alerts. Hand over to the team.",
        deliverables: ["Revenue Manager dashboard", "GM dashboard per property", "Pickup-drop alert active"],
      },
    ];

    const phW = 4.1, phH = 4.6, phGap = 0.18;
    const phStartX = (W - (3 * phW + 2 * phGap)) / 2;
    for (let i = 0; i < phases.length; i++) {
      const x = phStartX + i * (phW + phGap);
      const p = phases[i];

      // Card
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.2, w: phW, h: phH, fill: { color: C.card },
        line: { color: "E2E8F0", width: 0.75 }, shadow: shadowCard(),
      });
      // Top dark header
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.2, w: phW, h: 1.4, fill: { color: C.bgDark }, line: { type: "none" },
      });
      // Big number
      s.addText(p.num, {
        x: x + 0.3, y: 2.3, w: 1.5, h: 1.3,
        fontFace: "Georgia", fontSize: 60, bold: true, color: C.gold, margin: 0,
      });
      // Title
      s.addText(p.title, {
        x: x + 1.6, y: 2.4, w: phW - 1.7, h: 0.5,
        fontFace: "Calibri", fontSize: 13.5, bold: true, color: C.textLight,
        charSpacing: 4, margin: 0,
      });
      s.addText(p.weeks, {
        x: x + 1.6, y: 2.9, w: phW - 1.7, h: 0.4,
        fontFace: "Calibri", fontSize: 12, color: C.gold, italic: true, bold: true, margin: 0,
      });

      // Body
      s.addText(p.body, {
        x: x + 0.3, y: 3.8, w: phW - 0.6, h: 1.2,
        fontFace: "Calibri", fontSize: 13, color: C.ink, margin: 0,
      });

      // Divider
      s.addShape(pres.shapes.LINE, {
        x: x + 0.3, y: 5.05, w: phW - 0.6, h: 0,
        line: { color: C.gold, width: 1 },
      });
      s.addText("KEY DELIVERABLES", {
        x: x + 0.3, y: 5.1, w: phW - 0.6, h: 0.3,
        fontFace: "Calibri", fontSize: 10, bold: true, color: C.gold, charSpacing: 3, margin: 0,
      });

      // Deliverables
      const items = p.deliverables.map((d, idx) => ({
        text: d,
        options: { bullet: { code: "25A0" }, breakLine: idx < p.deliverables.length - 1, paraSpaceAfter: 4 },
      }));
      s.addText(items, {
        x: x + 0.3, y: 5.5, w: phW - 0.5, h: 1.3,
        fontFace: "Calibri", fontSize: 11.5, color: C.ink, margin: 0,
      });
    }

    // Bottom timeline arrow (above footer, fits in narrow space)
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 6.95, w: 12.3, h: 0.04, fill: { color: C.gold }, line: { type: "none" },
    });
    s.addShape(pres.shapes.OVAL, {
      x: 0.42, y: 6.91, w: 0.16, h: 0.16, fill: { color: C.gold }, line: { type: "none" },
    });
    s.addShape(pres.shapes.OVAL, {
      x: W - 0.58, y: 6.91, w: 0.16, h: 0.16, fill: { color: C.gold }, line: { type: "none" },
    });

    footer(s, 11);
  }

  // ============================================================
  // SLIDE 12 — THE ASK
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgDark };

    // Gold left band
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.3, h: H, fill: { color: C.gold }, line: { type: "none" },
    });

    goldChip(s, 0.7, 0.5, "THE DECISION");

    s.addText("Approve cloud migration. Unlock the platform.", {
      x: 0.7, y: 0.95, w: 11.6, h: 1.2, fontFace: "Georgia",
      fontSize: 36, bold: true, color: C.textLight, margin: 0,
    });
    s.addText("A small, time-boxed investment for a platform the team can rely on for years.", {
      x: 0.7, y: 2.15, w: 11.6, h: 0.5, fontFace: "Calibri",
      fontSize: 16, color: C.textDim, italic: true, margin: 0,
    });

    // Three big numbers
    const nums = [
      { v: "6",        l: "WEEKS",                sub: "from approval to production" },
      { v: "1×",       l: "MIGRATION EFFORT",     sub: "set up once, runs forever" },
      { v: "10+",      l: "HOTELS SUPPORTED",     sub: "without architectural change" },
    ];
    const nW = 3.6, nGap = 0.35, nTotalW = nums.length * nW + (nums.length - 1) * nGap;
    const nStartX = (W - nTotalW) / 2;
    nums.forEach((n, i) => {
      const x = nStartX + i * (nW + nGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 3.0, w: nW, h: 1.9, fill: { color: C.bgMid },
        line: { color: C.gold, width: 1 }, shadow: shadowSoft(),
      });
      s.addText(n.v, {
        x, y: 3.05, w: nW, h: 1.0, align: "center",
        fontFace: "Georgia", fontSize: 60, bold: true, color: C.gold, margin: 0,
      });
      s.addText(n.l, {
        x, y: 4.05, w: nW, h: 0.4, align: "center",
        fontFace: "Calibri", fontSize: 12, bold: true, color: C.textLight,
        charSpacing: 5, margin: 0,
      });
      s.addText(n.sub, {
        x, y: 4.45, w: nW, h: 0.4, align: "center",
        fontFace: "Calibri", fontSize: 11, italic: true, color: C.textDim, margin: 0,
      });
    });

    // What we need
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 5.25, w: 11.9, h: 1.55, fill: { color: C.bgMid }, line: { type: "none" },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 5.25, w: 0.1, h: 1.55, fill: { color: C.gold }, line: { type: "none" },
    });
    s.addText("WHAT WE NEED FROM YOU TODAY", {
      x: 1.0, y: 5.35, w: 11.5, h: 0.35, fontFace: "Calibri",
      fontSize: 11, bold: true, color: C.gold, charSpacing: 5, margin: 0,
    });
    s.addText([
      { text: "Approval to migrate to a cloud data warehouse and adopt Airflow as the orchestrator. ",
        options: { color: C.textLight, bold: true, breakLine: true } },
      { text: "That's it — implementation, deployment, and team rollout are already scoped end-to-end.",
        options: { color: C.textDim, italic: true } },
    ], {
      x: 1.0, y: 5.75, w: 11.5, h: 1.0, fontFace: "Calibri", fontSize: 15, margin: 0,
    });

    // Closing line
    s.addText("Today the data is brilliant — and trapped. Let's set it free.", {
      x: 0.7, y: 6.95, w: 11.9, h: 0.4, align: "center",
      fontFace: "Georgia", fontSize: 16, italic: true, bold: true, color: C.gold, margin: 0,
    });

    s.addText("12 / 12", {
      x: W - 1.5, y: H - 0.42, w: 1, h: 0.3, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0,
    });
  }

  await pres.writeFile({ fileName: "Cloud_Airflow_Migration_Proposal.pptx" });
  console.log("Built: Cloud_Airflow_Migration_Proposal.pptx");
}

build().catch(err => { console.error(err); process.exit(1); });
