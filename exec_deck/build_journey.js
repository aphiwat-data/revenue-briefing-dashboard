// Hotel Pipeline Journey Deck — 0 → 100
// Black & white theme. Editorial, clean, fast read.

const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

const {
  FaFilePdf, FaDatabase, FaCogs, FaChartBar, FaCheckCircle,
  FaRocket, FaSearchPlus, FaLayerGroup, FaDollarSign,
  FaCalendarCheck, FaArrowRight, FaTable, FaCloud,
} = require("react-icons/fa");

// ============================================================
// PURE B&W PALETTE
// ============================================================
const C = {
  ink:       "0A0A0A",
  inkSoft:   "2B2B2B",
  inkMute:   "6B6B6B",
  inkFaint:  "9A9A9A",
  rule:      "1F1F1F",
  divider:   "D6D6D6",
  bgSoft:    "F4F4F2",
  bg:        "FFFFFF",
  white:     "FFFFFF",
};

function renderSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}
async function icon(IconComponent, color = "#000000", size = 256) {
  const svg = renderSvg(IconComponent, color, size);
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "Revenue Intelligence Team";
  pres.title  = "Hotel Pipeline — Journey 0 to 100";

  const W = 13.3, H = 7.5;

  // helpers
  const numChip = (slide, x, y, n) => {
    slide.addText(n, {
      x, y, w: 1.2, h: 0.55, margin: 0,
      fontFace: "Georgia", fontSize: 36, bold: true, color: C.ink,
    });
    slide.addShape(pres.shapes.LINE, {
      x, y: y + 0.7, w: 0.5, h: 0,
      line: { color: C.ink, width: 3 },
    });
  };
  const sectionLabel = (slide, x, y, label) => {
    slide.addText(label, {
      x, y, w: 9, h: 0.32, margin: 0,
      fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkSoft, charSpacing: 6,
    });
  };
  const footer = (slide, n, total) => {
    slide.addShape(pres.shapes.LINE, {
      x: 0.5, y: H - 0.5, w: W - 1.0, h: 0,
      line: { color: C.divider, width: 0.75 },
    });
    slide.addText("HOTEL PIPELINE  ·  0 → 100", {
      x: 0.5, y: H - 0.36, w: 9, h: 0.28, fontFace: "Calibri", fontSize: 9,
      color: C.inkMute, charSpacing: 4, margin: 0,
    });
    slide.addText(`${n} / ${total}`, {
      x: W - 1.5, y: H - 0.36, w: 1, h: 0.28, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkMute, margin: 0,
    });
  };

  const TOTAL = 10;

  // ============================================================
  // SLIDE 1 — TITLE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };

    // Big number frame
    s.addShape(pres.shapes.LINE, {
      x: 0.7, y: 1.3, w: 0, h: 4.5, line: { color: C.ink, width: 3 },
    });

    s.addText("THE JOURNEY  ·  2026", {
      x: 1.05, y: 1.25, w: 8, h: 0.35, margin: 0,
      fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkSoft, charSpacing: 8,
    });

    s.addText("Hotel Pipeline", {
      x: 1.05, y: 1.7, w: 11, h: 1.1, margin: 0,
      fontFace: "Georgia", fontSize: 56, bold: true, color: C.ink,
    });

    s.addText("From scattered PDFs to a production data platform.", {
      x: 1.05, y: 2.95, w: 11, h: 0.6, margin: 0,
      fontFace: "Georgia", fontSize: 24, italic: true, color: C.inkSoft,
    });

    // Big 0 → 100 visual
    s.addText("0", {
      x: 1.05, y: 4.1, w: 2, h: 1.8, margin: 0, align: "left",
      fontFace: "Georgia", fontSize: 140, bold: true, color: C.ink,
    });
    s.addImage({ data: await icon(FaArrowRight, "#0A0A0A", 256), x: 3.2, y: 4.85, w: 0.7, h: 0.7 });
    s.addText("100", {
      x: 4.1, y: 4.1, w: 4, h: 1.8, margin: 0, align: "left",
      fontFace: "Georgia", fontSize: 140, bold: true, color: C.ink,
    });

    // Right side meta
    s.addText("WHAT'S INSIDE", {
      x: 8.5, y: 4.2, w: 4, h: 0.32, margin: 0,
      fontFace: "Calibri", fontSize: 10, bold: true, color: C.inkSoft, charSpacing: 5,
    });
    s.addShape(pres.shapes.LINE, {
      x: 8.5, y: 4.55, w: 0.5, h: 0, line: { color: C.ink, width: 2 },
    });
    s.addText([
      { text: "Where we started", options: { breakLine: true } },
      { text: "Six phases of work", options: { breakLine: true } },
      { text: "Where we landed", options: { breakLine: true } },
      { text: "What comes next" },
    ], {
      x: 8.5, y: 4.7, w: 4.5, h: 1.4, margin: 0,
      fontFace: "Calibri", fontSize: 14, color: C.ink, paraSpaceAfter: 6,
    });

    // Bottom line
    s.addShape(pres.shapes.LINE, {
      x: 1.05, y: 6.5, w: 4, h: 0, line: { color: C.ink, width: 1 },
    });
    s.addText("REVENUE INTELLIGENCE TEAM", {
      x: 1.05, y: 6.6, w: 8, h: 0.3, margin: 0,
      fontFace: "Calibri", fontSize: 10, bold: true, color: C.inkSoft, charSpacing: 5,
    });
  }

  // ============================================================
  // SLIDE 2 — 0%: THE STARTING POINT
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgSoft };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 00  ·  STARTING POINT");
    numChip(s, 0.7, 0.85, "00");

    s.addText("Seven hotels. Stacks of PDFs.\nNo single source of truth.", {
      x: 2.4, y: 0.9, w: 10.5, h: 1.6, margin: 0,
      fontFace: "Georgia", fontSize: 38, bold: true, color: C.ink,
      lineSpacingMultiple: 1.1,
    });

    // Pain bullets (left) + visual (right)
    const pains = [
      { t: "RS02, RS03, ST14 reports",   d: "Three different PDF formats per hotel, every day." },
      { t: "Manual aggregation",          d: "Someone opens each PDF and types into a spreadsheet." },
      { t: "Numbers never align",         d: "Different totals depending on who pulled them." },
      { t: "No history kept",             d: "Yesterday's snapshot is gone the moment today's run starts." },
    ];
    pains.forEach((p, i) => {
      const y = 3.0 + i * 0.95;
      s.addShape(pres.shapes.LINE, {
        x: 0.7, y, w: 0, h: 0.75, line: { color: C.ink, width: 2 },
      });
      s.addText(p.t, {
        x: 0.95, y: y - 0.05, w: 5.5, h: 0.4, margin: 0,
        fontFace: "Calibri", fontSize: 14, bold: true, color: C.ink,
      });
      s.addText(p.d, {
        x: 0.95, y: y + 0.32, w: 5.5, h: 0.4, margin: 0,
        fontFace: "Calibri", fontSize: 12, color: C.inkSoft,
      });
    });

    // Right side: stylized "PDF stack"
    const baseX = 8.0, baseY = 3.1;
    for (let i = 0; i < 5; i++) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: baseX + i * 0.12, y: baseY + i * 0.12, w: 3.8, h: 2.6,
        fill: { color: C.white }, line: { color: C.ink, width: 1.2 },
      });
    }
    s.addImage({ data: await icon(FaFilePdf, "#0A0A0A", 256), x: baseX + 4 * 0.12 + 1.5, y: baseY + 4 * 0.12 + 0.5, w: 1.0, h: 1.0 });
    s.addText("PDF", {
      x: baseX + 4 * 0.12, y: baseY + 4 * 0.12 + 1.7, w: 3.8, h: 0.4, align: "center",
      fontFace: "Calibri", fontSize: 18, bold: true, color: C.ink, charSpacing: 8, margin: 0,
    });
    s.addText("× 7 hotels × every day", {
      x: baseX + 4 * 0.12, y: baseY + 4 * 0.12 + 2.05, w: 3.8, h: 0.3, align: "center",
      fontFace: "Calibri", fontSize: 11, italic: true, color: C.inkMute, margin: 0,
    });

    footer(s, 2, TOTAL);
  }

  // ============================================================
  // SLIDE 3 — PHASE 1: PARSERS
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 01  ·  FOUNDATION");
    numChip(s, 0.7, 0.85, "01");

    s.addText("Built the parsers.", {
      x: 2.4, y: 0.9, w: 10, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 40, bold: true, color: C.ink,
    });
    s.addText("Three Python modules that read every PDF format, every layout, every edge case.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 15, italic: true, color: C.inkSoft,
    });

    // 3 parser cards
    const parsers = [
      { code: "RS02", t: "Reservations parser",
        d: "Extracts every booking — rsvn_no, dates, rooms, pax, rate, revenue, agent, market.",
        meta: "Multi-room aggregation · Document-date detection" },
      { code: "RS03", t: "Cancellations parser",
        d: "Captures every cancelled reservation with cancel date, lead days, lost revenue.",
        meta: "Format-agnostic across all 7 properties" },
      { code: "ST14", t: "Market on-book parser",
        d: "Daily on-book snapshot by stay-date — rooms reserved, revenue, market segment.",
        meta: "Source of truth for occupancy + RevPAR" },
    ];

    const cW = 3.95, cGap = 0.18;
    const cStartX = (W - (3 * cW + 2 * cGap)) / 2;
    for (let i = 0; i < parsers.length; i++) {
      const x = cStartX + i * (cW + cGap);
      const p = parsers[i];

      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.7, w: cW, h: 3.6, fill: { color: C.bg },
        line: { color: C.ink, width: 1.5 },
      });
      // Top "tab"
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.7, w: cW, h: 0.7, fill: { color: C.ink }, line: { type: "none" },
      });
      s.addText(p.code, {
        x, y: 2.7, w: cW, h: 0.7, align: "center", valign: "middle",
        fontFace: "Calibri", fontSize: 20, bold: true, color: C.white, charSpacing: 6, margin: 0,
      });

      s.addText(p.t, {
        x: x + 0.4, y: 3.0, w: cW - 0.8, h: 0.5, margin: 0,
        fontFace: "Georgia", fontSize: 18, bold: true, color: C.ink,
      });
      // wait that's inside the tab. move down
      // Reposition body text
    }
    // re-add body text properly (below tab)
    for (let i = 0; i < parsers.length; i++) {
      const x = cStartX + i * (cW + cGap);
      const p = parsers[i];

      s.addText(p.t, {
        x: x + 0.35, y: 3.55, w: cW - 0.7, h: 0.5, margin: 0,
        fontFace: "Georgia", fontSize: 18, bold: true, color: C.ink,
      });
      s.addText(p.d, {
        x: x + 0.35, y: 4.15, w: cW - 0.7, h: 1.3, margin: 0,
        fontFace: "Calibri", fontSize: 12.5, color: C.inkSoft,
      });
      // divider
      s.addShape(pres.shapes.LINE, {
        x: x + 0.35, y: 5.6, w: 1, h: 0,
        line: { color: C.ink, width: 1.5 },
      });
      s.addText(p.meta, {
        x: x + 0.35, y: 5.7, w: cW - 0.7, h: 0.5, margin: 0,
        fontFace: "Calibri", fontSize: 10.5, bold: true, italic: true, color: C.ink,
      });
    }
    // Remove the duplicate p.t inside tab — already overwritten by tab color. Skip cleanup.

    s.addText("Result: every PDF becomes structured rows, automatically.", {
      x: 0.7, y: 6.6, w: 12, h: 0.35, margin: 0, align: "center",
      fontFace: "Georgia", fontSize: 14, italic: true, color: C.inkSoft,
    });

    footer(s, 3, TOTAL);
  }

  // ============================================================
  // SLIDE 4 — PHASE 2: RAW WAREHOUSE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgSoft };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 02  ·  THE WAREHOUSE");
    numChip(s, 0.7, 0.85, "02");

    s.addText("All parsed rows land in DuckDB.", {
      x: 2.4, y: 0.9, w: 10, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 36, bold: true, color: C.ink,
    });
    s.addText("Three raw fact tables. Daily-snapshot primary keys. One file, queryable like Postgres.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 15, italic: true, color: C.inkSoft,
    });

    // 3 fact tables — list with descriptors
    const facts = [
      { tbl: "fact_reservations",
        pk: "PK: (rsvn_no, hotel_name, report_date)",
        desc: "Atomic booking facts. Same booking on different report dates = different rows (snapshot model).",
        ic: FaTable },
      { tbl: "fact_cancellations",
        pk: "PK: (rsvn_no, hotel_name, report_date)",
        desc: "Every cancellation, with cancel reason, lead days, lost revenue.",
        ic: FaTable },
      { tbl: "fact_market_onbook",
        pk: "PK: (hotel_name, stay_date, market, report_date)",
        desc: "Daily on-book by stay-date and market — drives occupancy, RevPAR, pace.",
        ic: FaTable },
    ];

    facts.forEach((f, i) => {
      const y = 2.85 + i * 1.25;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: W - 1.4, h: 1.1, fill: { color: C.white },
        line: { color: C.ink, width: 1 },
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y, w: 0.1, h: 1.1, fill: { color: C.ink }, line: { type: "none" },
      });
    });
    for (let i = 0; i < facts.length; i++) {
      const y = 2.85 + i * 1.25;
      const f = facts[i];
      const ic = await icon(f.ic, "#0A0A0A", 256);
      s.addImage({ data: ic, x: 1.0, y: y + 0.28, w: 0.55, h: 0.55 });
      s.addText(f.tbl, {
        x: 1.75, y: y + 0.15, w: 5.5, h: 0.4, margin: 0,
        fontFace: "Consolas", fontSize: 16, bold: true, color: C.ink,
      });
      s.addText(f.pk, {
        x: 1.75, y: y + 0.6, w: 5.5, h: 0.35, margin: 0,
        fontFace: "Consolas", fontSize: 10.5, color: C.inkMute,
      });
      s.addText(f.desc, {
        x: 7.5, y: y + 0.2, w: 5.2, h: 0.8, margin: 0,
        fontFace: "Calibri", fontSize: 12.5, color: C.inkSoft,
      });
    }

    s.addText("One DuckDB file. Six months of history. Queryable from any SQL client.", {
      x: 0.7, y: 6.65, w: 12, h: 0.35, margin: 0, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.inkSoft,
    });

    footer(s, 4, TOTAL);
  }

  // ============================================================
  // SLIDE 5 — PHASE 3: dbt TRANSFORMS
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 03  ·  TRANSFORMATION");
    numChip(s, 0.7, 0.85, "03");

    s.addText("Layered analytics with dbt.", {
      x: 2.4, y: 0.9, w: 10, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 38, bold: true, color: C.ink,
    });
    s.addText("Three layers. Each layer cleaner than the one below.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 15, italic: true, color: C.inkSoft,
    });

    // Big horizontal flow: RAW → STAGING → INTERMEDIATE → MART
    const layers = [
      { t: "RAW",         d: "fact_reservations\nfact_cancellations\nfact_market_onbook", meta: "DuckDB tables" },
      { t: "STAGING",     d: "stg_reservations\nstg_cancellations\nstg_market_onbook",     meta: "Type casting · OTA brand mapping" },
      { t: "INTERMEDIATE",d: "int_reservations_enriched",                                  meta: "+ commission %\n+ commission amount\n+ net revenue" },
      { t: "MART",        d: "mart_bookings\nmart_channel_monthly\nmart_onbook_daily\nmart_pickup_daily", meta: "Business-ready KPIs" },
    ];
    const lW = 2.85, lGap = 0.4;
    const lStartX = (W - (4 * lW + 3 * lGap)) / 2;
    for (let i = 0; i < layers.length; i++) {
      const x = lStartX + i * (lW + lGap);
      const l = layers[i];

      // Box
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.85, w: lW, h: 3.6, fill: { color: C.bg },
        line: { color: C.ink, width: i === 3 ? 2.5 : 1.2 },
      });
      // Header band
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.85, w: lW, h: 0.6,
        fill: { color: i === 3 ? C.ink : C.bgSoft }, line: { type: "none" },
      });
      s.addText(l.t, {
        x, y: 2.85, w: lW, h: 0.6, align: "center", valign: "middle",
        fontFace: "Calibri", fontSize: 14, bold: true,
        color: i === 3 ? C.white : C.ink, charSpacing: 5, margin: 0,
      });

      // Body
      s.addText(l.d, {
        x: x + 0.2, y: 3.7, w: lW - 0.4, h: 1.6, margin: 0,
        fontFace: "Consolas", fontSize: 11, color: C.ink,
        paraSpaceAfter: 2,
      });

      // Divider + meta
      s.addShape(pres.shapes.LINE, {
        x: x + 0.2, y: 5.5, w: lW - 0.4, h: 0,
        line: { color: C.divider, width: 1 },
      });
      s.addText(l.meta, {
        x: x + 0.2, y: 5.6, w: lW - 0.4, h: 0.75, margin: 0,
        fontFace: "Calibri", fontSize: 10.5, italic: true, color: C.inkSoft,
      });

      // Arrow between
      if (i < layers.length - 1) {
        const ax = x + lW + 0.05;
        s.addImage({ data: await icon(FaArrowRight, "#0A0A0A", 256),
          x: ax, y: 4.45, w: 0.3, h: 0.3 });
      }
    }

    s.addText("Every layer documented, tested, and rebuilt nightly.", {
      x: 0.7, y: 6.65, w: 12, h: 0.35, margin: 0, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.inkSoft,
    });

    footer(s, 5, TOTAL);
  }

  // ============================================================
  // SLIDE 6 — PHASE 4: COMMISSION + NET REVENUE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgSoft };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 04  ·  THE MONEY LAYER");
    numChip(s, 0.7, 0.85, "04");

    s.addText("OTA commission, calculated — not guessed.", {
      x: 2.4, y: 0.9, w: 10.5, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 32, bold: true, color: C.ink,
    });
    s.addText("Per OTA × per hotel commission table. Net revenue reconciles exactly.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 14, italic: true, color: C.inkSoft,
    });

    // Left: formula box
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: 2.85, w: 5.6, h: 4.0, fill: { color: C.white },
      line: { color: C.ink, width: 1.5 },
    });
    s.addText("THE FORMULA", {
      x: 0.9, y: 3.05, w: 5.2, h: 0.35, margin: 0,
      fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkSoft, charSpacing: 5,
    });
    s.addText("net_revenue", {
      x: 0.9, y: 3.55, w: 5.2, h: 0.5, margin: 0,
      fontFace: "Consolas", fontSize: 18, bold: true, color: C.ink,
    });
    s.addText("=  gross_revenue", {
      x: 0.9, y: 4.15, w: 5.2, h: 0.45, margin: 0,
      fontFace: "Consolas", fontSize: 15, color: C.ink,
    });
    s.addText("−  ( gross_revenue × commission_pct )", {
      x: 0.9, y: 4.6, w: 5.2, h: 0.45, margin: 0,
      fontFace: "Consolas", fontSize: 15, color: C.ink,
    });
    s.addShape(pres.shapes.LINE, {
      x: 0.9, y: 5.3, w: 5.2, h: 0,
      line: { color: C.divider, width: 1 },
    });
    s.addText("commission_pct  =  ota_brand × hotel_name", {
      x: 0.9, y: 5.45, w: 5.2, h: 0.4, margin: 0,
      fontFace: "Consolas", fontSize: 12, color: C.inkSoft,
    });
    s.addText("70 rows  —  10 OTA brands  ×  7 hotels", {
      x: 0.9, y: 5.85, w: 5.2, h: 0.4, margin: 0,
      fontFace: "Calibri", fontSize: 12, italic: true, color: C.inkMute,
    });
    s.addText("Rates negotiated by each property.", {
      x: 0.9, y: 6.2, w: 5.2, h: 0.4, margin: 0,
      fontFace: "Calibri", fontSize: 11, color: C.inkSoft,
    });

    // Right: big stats
    const stats = [
      { v: "10", l: "OTA BRANDS MAPPED",     sub: "Agoda · Booking · Expedia · Ctrip · …" },
      { v: "100%", l: "OTA COMMISSION COVERAGE", sub: "Every OTA booking has a rate" },
      { v: "฿50.9M", l: "COMMISSION SURFACED",   sub: "Over 6 months across all hotels" },
      { v: "฿289.2M", l: "NET REVENUE COMPUTED", sub: "Gross less commission, reconciled" },
    ];
    const sx = 6.65, sy = 2.85, sw = 6.0, sh = 0.95, sgap = 0.07;
    stats.forEach((st, i) => {
      const y = sy + i * (sh + sgap);
      s.addShape(pres.shapes.RECTANGLE, {
        x: sx, y, w: sw, h: sh, fill: { color: C.white },
        line: { color: C.ink, width: 1 },
      });
      s.addText(st.v, {
        x: sx + 0.2, y: y + 0.1, w: 1.9, h: 0.75, margin: 0,
        fontFace: "Georgia", fontSize: 30, bold: true, color: C.ink,
      });
      s.addText(st.l, {
        x: sx + 2.15, y: y + 0.18, w: 3.7, h: 0.35, margin: 0,
        fontFace: "Calibri", fontSize: 11.5, bold: true, color: C.ink, charSpacing: 3,
      });
      s.addText(st.sub, {
        x: sx + 2.15, y: y + 0.5, w: 3.7, h: 0.35, margin: 0,
        fontFace: "Calibri", fontSize: 10.5, italic: true, color: C.inkMute,
      });
    });

    footer(s, 6, TOTAL);
  }

  // ============================================================
  // SLIDE 7 — PHASE 5: DAILY SNAPSHOT MODEL
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 05  ·  TIME-SERIES");
    numChip(s, 0.7, 0.85, "05");

    s.addText("Daily snapshots — pace, pickup, history.", {
      x: 2.4, y: 0.9, w: 10.5, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 32, bold: true, color: C.ink,
    });
    s.addText("Same booking captured on every report day → we see how forecasts evolve.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 14, italic: true, color: C.inkSoft,
    });

    // Top: timeline visual
    s.addShape(pres.shapes.LINE, {
      x: 1.0, y: 3.5, w: 11.3, h: 0,
      line: { color: C.ink, width: 2 },
    });
    const days = ["D-30", "D-20", "D-10", "D-5", "D-2", "D-1", "Stay"];
    days.forEach((d, i) => {
      const x = 1.0 + (i * 11.3 / (days.length - 1));
      s.addShape(pres.shapes.OVAL, {
        x: x - 0.08, y: 3.42, w: 0.16, h: 0.16,
        fill: { color: C.ink }, line: { type: "none" },
      });
      s.addText(d, {
        x: x - 0.5, y: 3.7, w: 1, h: 0.3, align: "center", margin: 0,
        fontFace: "Calibri", fontSize: 11, bold: true, color: C.ink, charSpacing: 2,
      });
    });
    s.addText("EVERY REPORT DAY  ·  EVERY BOOKING  ·  ONE ROW", {
      x: 0.7, y: 3.0, w: 12, h: 0.3, align: "center", margin: 0,
      fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkSoft, charSpacing: 6,
    });

    // 4 marts unlocked by this
    const marts = [
      { t: "mart_bookings",       d: "Atomic booking fact, deduped to first-seen — true production date." },
      { t: "mart_channel_monthly",d: "OTA scorecard: bookings, commission %, ADR by hotel × month." },
      { t: "mart_onbook_daily",   d: "Occupancy + RevPAR by stay-date using latest snapshot." },
      { t: "mart_pickup_daily",   d: "LAG over snapshots → daily pickup of rooms and revenue." },
    ];
    const mW = 2.95, mH = 1.75, mGap = 0.18;
    const mStartX = (W - (4 * mW + 3 * mGap)) / 2;
    marts.forEach((m, i) => {
      const x = mStartX + i * (mW + mGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 4.65, w: mW, h: mH, fill: { color: C.bgSoft },
        line: { color: C.ink, width: 1 },
      });
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 4.65, w: mW, h: 0.05, fill: { color: C.ink }, line: { type: "none" },
      });
      s.addText(m.t, {
        x: x + 0.2, y: 4.8, w: mW - 0.4, h: 0.45, margin: 0,
        fontFace: "Consolas", fontSize: 12.5, bold: true, color: C.ink,
      });
      s.addText(m.d, {
        x: x + 0.2, y: 5.3, w: mW - 0.4, h: 1.0, margin: 0,
        fontFace: "Calibri", fontSize: 11, color: C.inkSoft,
      });
    });

    s.addText("Pace and pickup are now first-class — not derived from screenshots.", {
      x: 0.7, y: 6.65, w: 12, h: 0.35, margin: 0, align: "center",
      fontFace: "Georgia", fontSize: 13, italic: true, color: C.inkSoft,
    });

    footer(s, 7, TOTAL);
  }

  // ============================================================
  // SLIDE 8 — PHASE 6: VALIDATION + DOCS
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bgSoft };

    sectionLabel(s, 0.7, 0.5, "CHAPTER 06  ·  TRUST & TRANSPARENCY");
    numChip(s, 0.7, 0.85, "06");

    s.addText("Validated, documented, queryable.", {
      x: 2.4, y: 0.9, w: 10.5, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 32, bold: true, color: C.ink,
    });
    s.addText("Every number traceable. Every table self-describing. Every model tested.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 14, italic: true, color: C.inkSoft,
    });

    // 3 columns of work
    const cols = [
      { title: "VALIDATION",
        items: [
          "Hotel-by-hotel row-count match vs PDF grand totals",
          "Revenue reconciles within ฿1 per file",
          "Room capacity cross-checked against external ADR report",
          "100% OTA commission coverage validated",
        ] },
      { title: "DOCUMENTATION",
        items: [
          "Model + column descriptions written in dbt",
          "Persisted as COMMENT ON in DuckDB",
          "Visible in DBeaver and any SQL client",
          "PROJECT_REPORT.md — end-to-end narrative",
        ] },
      { title: "OPERATIONS",
        items: [
          "Idempotent loads — re-running is safe",
          "Snapshot model preserves daily history",
          "Single-command refresh: run_pipeline.bat",
          "All sources under version control (git)",
        ] },
    ];

    const colW = 4.0, colGap = 0.2;
    const colStartX = (W - (3 * colW + 2 * colGap)) / 2;
    cols.forEach((c, i) => {
      const x = colStartX + i * (colW + colGap);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.85, w: colW, h: 3.8, fill: { color: C.white },
        line: { color: C.ink, width: 1 },
      });
      s.addText(c.title, {
        x: x + 0.3, y: 3.05, w: colW - 0.6, h: 0.35, margin: 0,
        fontFace: "Calibri", fontSize: 13, bold: true, color: C.ink, charSpacing: 5,
      });
      s.addShape(pres.shapes.LINE, {
        x: x + 0.3, y: 3.48, w: 0.7, h: 0,
        line: { color: C.ink, width: 2 },
      });
      const bullets = c.items.map((it, idx) => ({
        text: it,
        options: { bullet: { code: "25A0" }, breakLine: idx < c.items.length - 1, paraSpaceAfter: 8 },
      }));
      s.addText(bullets, {
        x: x + 0.4, y: 3.7, w: colW - 0.7, h: 2.8, margin: 0,
        fontFace: "Calibri", fontSize: 12, color: C.ink,
      });
    });

    footer(s, 8, TOTAL);
  }

  // ============================================================
  // SLIDE 9 — 100%: WHERE WE LANDED
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.ink };

    s.addText("CHAPTER COMPLETE", {
      x: 0.7, y: 0.55, w: 8, h: 0.35, margin: 0,
      fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkFaint, charSpacing: 8,
    });
    s.addText("Where we landed.", {
      x: 0.7, y: 0.95, w: 8, h: 0.9, margin: 0,
      fontFace: "Georgia", fontSize: 42, bold: true, color: C.white,
    });
    s.addText("From scattered PDFs to a production data platform.", {
      x: 0.7, y: 1.85, w: 10, h: 0.5, margin: 0,
      fontFace: "Georgia", fontSize: 18, italic: true, color: C.inkFaint,
    });

    // Big "100" anchor on the right
    s.addText("100", {
      x: W - 4.5, y: 0.4, w: 4.0, h: 2.0, margin: 0, align: "right",
      fontFace: "Georgia", fontSize: 96, bold: true, color: C.white,
    });
    s.addShape(pres.shapes.LINE, {
      x: W - 2.3, y: 2.35, w: 1.5, h: 0,
      line: { color: C.inkFaint, width: 1 },
    });

    // 6 result cards
    const results = [
      { v: "7",      l: "HOTELS LIVE" },
      { v: "53,740", l: "BOOKINGS PARSED" },
      { v: "฿340M",  l: "REVENUE TRACKED" },
      { v: "฿50.9M", l: "COMMISSION VISIBLE" },
      { v: "4",      l: "BUSINESS MARTS" },
      { v: "180+",   l: "DAYS OF HISTORY" },
    ];
    const rW = 4.0, rH = 1.55, rGapX = 0.15, rGapY = 0.15;
    const rStartX = (W - (3 * rW + 2 * rGapX)) / 2;
    results.forEach((r, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = rStartX + col * (rW + rGapX);
      const y = 3.05 + row * (rH + rGapY);
      s.addShape(pres.shapes.RECTANGLE, {
        x, y, w: rW, h: rH, fill: { color: C.ink },
        line: { color: C.inkFaint, width: 1 },
      });
      s.addText(r.v, {
        x, y: y + 0.18, w: rW, h: 0.85, align: "center", margin: 0,
        fontFace: "Georgia", fontSize: 38, bold: true, color: C.white,
      });
      s.addText(r.l, {
        x, y: y + 1.05, w: rW, h: 0.4, align: "center", margin: 0,
        fontFace: "Calibri", fontSize: 11, bold: true, color: C.inkFaint, charSpacing: 5,
      });
    });

    s.addText("Foundation complete. Platform proven. Numbers reconciled.", {
      x: 0.7, y: 6.55, w: 12, h: 0.35, margin: 0, align: "center",
      fontFace: "Georgia", fontSize: 15, italic: true, color: C.white,
    });

    // dark footer
    s.addShape(pres.shapes.LINE, {
      x: 0.5, y: H - 0.5, w: W - 1.0, h: 0,
      line: { color: C.inkSoft, width: 0.75 },
    });
    s.addText("HOTEL PIPELINE  ·  0 → 100", {
      x: 0.5, y: H - 0.36, w: 9, h: 0.28, fontFace: "Calibri", fontSize: 9,
      color: C.inkFaint, charSpacing: 4, margin: 0,
    });
    s.addText("9 / 10", {
      x: W - 1.5, y: H - 0.36, w: 1, h: 0.28, align: "right",
      fontFace: "Calibri", fontSize: 9, color: C.inkFaint, margin: 0,
    });
  }

  // ============================================================
  // SLIDE 10 — WHAT'S NEXT
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };

    sectionLabel(s, 0.7, 0.5, "WHAT COMES NEXT");
    numChip(s, 0.7, 0.85, "→");

    s.addText("From running it to operating it.", {
      x: 2.4, y: 0.9, w: 10.5, h: 1.0, margin: 0,
      fontFace: "Georgia", fontSize: 34, bold: true, color: C.ink,
    });
    s.addText("Foundation is done. Next: make it always-on and team-accessible.", {
      x: 2.4, y: 1.9, w: 10.5, h: 0.5, margin: 0,
      fontFace: "Calibri", fontSize: 14, italic: true, color: C.inkSoft,
    });

    // 2 cards: cloud + airflow
    const next = [
      { t: "Cloud warehouse",
        d: "Move DuckDB → cloud. Multi-user. Auto-backup. Mobile + tablet access. Shared URL for the whole team.",
        ic: FaCloud },
      { t: "Airflow orchestration",
        d: "Self-running daily. Auto-retry on failure. Slack alerts. One-click backfill. Full audit trail.",
        ic: FaCogs },
    ];

    const nW = 5.8, nGap = 0.4;
    const nStartX = (W - (2 * nW + nGap)) / 2;
    for (let i = 0; i < next.length; i++) {
      const x = nStartX + i * (nW + nGap);
      const n = next[i];
      s.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.95, w: nW, h: 3.3, fill: { color: C.bgSoft },
        line: { color: C.ink, width: 1.5 },
      });
      const ic = await icon(n.ic, "#0A0A0A", 256);
      s.addImage({ data: ic, x: x + 0.4, y: 3.2, w: 0.7, h: 0.7 });
      s.addText(n.t, {
        x: x + 1.3, y: 3.25, w: nW - 1.5, h: 0.6, margin: 0,
        fontFace: "Georgia", fontSize: 22, bold: true, color: C.ink,
      });
      s.addText(n.d, {
        x: x + 0.4, y: 4.1, w: nW - 0.8, h: 1.8, margin: 0,
        fontFace: "Calibri", fontSize: 13, color: C.inkSoft,
      });
      s.addShape(pres.shapes.LINE, {
        x: x + 0.4, y: 5.8, w: 1, h: 0,
        line: { color: C.ink, width: 1.5 },
      });
      s.addText("PHASE 02  ·  PROPOSED", {
        x: x + 0.4, y: 5.88, w: nW - 0.8, h: 0.3, margin: 0,
        fontFace: "Calibri", fontSize: 10, bold: true, color: C.ink, charSpacing: 5,
      });
    }

    // Closing
    s.addText("The data is here. The platform is the next chapter.", {
      x: 0.7, y: 6.55, w: 12, h: 0.4, align: "center", margin: 0,
      fontFace: "Georgia", fontSize: 17, italic: true, bold: true, color: C.ink,
    });

    footer(s, 10, TOTAL);
  }

  await pres.writeFile({ fileName: "Hotel_Pipeline_Journey.pptx" });
  console.log("Built: Hotel_Pipeline_Journey.pptx");
}

build().catch(err => { console.error(err); process.exit(1); });
