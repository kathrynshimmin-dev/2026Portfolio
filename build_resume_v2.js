const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
  VerticalAlign, TabStopType, TabStopPosition, PageBreak
} = require('docx');
const fs = require('fs');

// ── Palette ──────────────────────────────────────────────────────────────────
const ACCENT = "DB6763";
const BLACK  = "1F2328";
const MUTED  = "57606A";
const BORDER = "E5E7EB";
const SURF   = "F7F8FA";

// ── Page geometry (Letter, 0.6" margins) ─────────────────────────────────────
// Total width = 12240 DXA. Margins 864 each side → usable = 10512 DXA
// Left col = 6900, gap col = 180, right col = 3432
const TOTAL   = 10512;
const LEFT_W  = 6900;
const GAP_W   = 180;
const RIGHT_W = TOTAL - LEFT_W - GAP_W; // 3432

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helpers ───────────────────────────────────────────────────────────────────
function secHead(text) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 16, color: MUTED, font: "Arial", characterSpacing: 80 })]
  });
}

function sideHead(text) {
  return new Paragraph({
    spacing: { before: 140, after: 50 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 14, color: MUTED, font: "Arial", characterSpacing: 80 })]
  });
}

function bullet(runs, ref = "main-bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 30, after: 30 },
    children: runs
  });
}

function subBullet(text) {
  return new Paragraph({
    numbering: { reference: "sub-bullets", level: 0 },
    spacing: { before: 20, after: 20 },
    children: [new TextRun({ text, size: 20, font: "Arial", color: MUTED })]
  });
}

function t(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: opts.size || 22, color: opts.color || BLACK, bold: opts.bold || false, italics: opts.italics || false });
}

function tb(text, size = 22) { return t(text, { bold: true, size }); }

function spacer(before = 80) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [t("")] });
}

function companyRow(name, tenure) {
  return new Paragraph({
    spacing: { before: 160, after: 30 },
    tabStops: [{ type: TabStopType.RIGHT, position: LEFT_W - 100 }],
    children: [
      new TextRun({ text: name, bold: true, size: 26, font: "Arial", color: BLACK }),
      new TextRun({ text: "\t" + tenure, size: 22, font: "Arial", color: BLACK })
    ]
  });
}

function subroleRow(title, dates) {
  return new Paragraph({
    spacing: { before: 20, after: 20 },
    tabStops: [{ type: TabStopType.RIGHT, position: LEFT_W - 100 }],
    children: [
      new TextRun({ text: title, bold: true, size: 22, font: "Arial", color: ACCENT }),
      new TextRun({ text: "\t" + dates, size: 20, font: "Arial", color: ACCENT })
    ]
  });
}

function accomplishmentsLabel() {
  return new Paragraph({
    spacing: { before: 60, after: 40 },
    children: [new TextRun({ text: "SELECTED ACCOMPLISHMENTS", bold: true, size: 17, font: "Arial", color: MUTED, characterSpacing: 60 })]
  });
}

function metricBlock(value, label) {
  return [
    new Paragraph({
      spacing: { before: 60, after: 0 },
      shading: { fill: SURF, type: ShadingType.CLEAR },
      children: [new TextRun({ text: value, bold: true, size: 28, font: "Arial", color: BLACK })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      shading: { fill: SURF, type: ShadingType.CLEAR },
      children: [new TextRun({ text: label, size: 17, font: "Arial", color: MUTED })]
    })
  ];
}

function awardBlock(name, detail) {
  return [
    new Paragraph({
      spacing: { before: 60, after: 10 },
      children: [new TextRun({ text: name, bold: true, size: 20, font: "Arial", color: BLACK })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 50 },
      children: [new TextRun({ text: detail, size: 18, font: "Arial", color: MUTED })]
    })
  ];
}

function skillGroup(label, tags) {
  return [
    new Paragraph({
      spacing: { before: 60, after: 20 },
      children: [new TextRun({ text: label, bold: true, size: 19, font: "Arial", color: BLACK })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: tags.join("  ·  "), size: 18, font: "Arial", color: MUTED })]
    })
  ];
}

function testimony(quote, cite) {
  return [
    new Paragraph({
      spacing: { before: 60, after: 10 },
      border: { left: { style: BorderStyle.SINGLE, size: 6, color: BORDER } },
      indent: { left: 120 },
      children: [new TextRun({ text: `"${quote}"`, size: 18, font: "Arial", color: MUTED, italics: true })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 50 },
      indent: { left: 120 },
      children: [new TextRun({ text: cite, bold: true, size: 17, font: "Arial", color: BLACK })]
    })
  ];
}

// ── Left column content ───────────────────────────────────────────────────────
const leftCol = [
  // Summary
  secHead("Summary"),
  new Paragraph({
    spacing: { before: 80, after: 120 },
    children: [new TextRun({
      text: "I turn big ideas into experiences people don't forget. Whether it's an executive briefing, an innovation showcase, or an immersive Experience Center, I bring creative vision, strategic clarity, and the energy to make it land. I've spent my career at the intersection of design, technology, and storytelling—directing teams, agencies, and cross-functional partners to translate complex innovation themes into compelling visual narratives and customer experiences. I'm fluent in AI and emerging tech, and I know how to make those ideas feel exciting and accessible to a C-suite audience. I also bring a strong background in Financial Services—understanding the stakes, the audiences, and what it takes to earn trust in that space. Fast-paced, matrixed environments are where I do my best work.",
      size: 21, font: "Arial", color: BLACK
    })]
  }),

  // Experience
  secHead("Experience"),

  // IBM
  companyRow("IBM", "4.5 years"),
  subroleRow("AI Forward Deployed Engineer, Financial Services", "2026 – Present"),
  subroleRow("Head of Strategy Design, Americas Client Engineering", "2025 – 2026"),
  subroleRow("Advisory Innovation Designer", "2021 – 2025"),
  accomplishmentsLabel(),
  bullet([tb("Built and scaled a client-centered account-planning operating model"), t(" for IBM FSM across 63 accounts in one quarter, producing "), tb("$50.6M in won revenue"), t(" and 1,000+ surfaced opportunities; framework adopted by two additional IBM markets.")]),
  bullet([tb("Designed and directed a State Street AI client experience"), t(" (60+ participants, 9 qualified AI use cases) codified into a playbook mandated across 7 IBM markets by a General Manager the following year.")]),
  bullet([t("Led design strategy, program design, and facilitation enablement across IBM's Financial Services, Commercial, and Social Impact markets, coordinating "), tb("89 designers"), t(" across six Americas markets.")]),
  bullet([t("Originated a two-year "), tb("nonprofit AI adoption framework"), t(" for IBM Social Impact, partnering with funders to generate "), tb("$750K+ in opportunity value"), t(" and 20+ AI use cases across housing, healthcare, education, and economic mobility sectors.")]),
  bullet([t("Created "), tb("Career Pathways & Progression"), t(" program artifacts scaled to 200+ Innovation Designers worldwide in 2025, mapping IBM core and adjacent design roles with skills gaps for long-term career guidance.")]),
  bullet([t("Enabled "), tb("15 designers across 6 markets"), t(" to attend Adobe Creative Jam and Figma Config through business cases and budget advocacy.")]),
  bullet([t("Created and led a role-based AI upskilling program for ~80 IBM practitioners: "), tb("93.8% post-event confidence"), t(", 111.73% increase in AI-generated code two months post-event, and 25% of participants advancing a full proficiency level.")]),
  bullet([t("Named "), tb("Top 7 of 2,000"), t(" IBM Client Engineering practitioners worldwide (2024) and selected for the "), tb("Artemis Program"), t(" cohort tasked with transforming how Client Engineering operates as a business.")]),

  spacer(60),

  // Newell Brands
  companyRow("Newell Brands", "7.5 years"),
  subroleRow("Senior Art Director, Experience Design", "2019 – 2021"),
  subroleRow("Senior Graphic Designer", "2015 – 2019"),
  subroleRow("Graphic Designer", "2014 – 2015"),
  accomplishmentsLabel(),
  bullet([t("Identified a gap in consumer-driven eCommerce design, secured executive buy-in for a new Experience Design function, and recruited, developed, and scaled the team from 1 to 3 members ("), tb("200% growth"), t(") in under a year, tripling delivery capacity. Integrated user research as part of a new pre-launch process for CPG brands across outdoor, food, cookware, and baby markets.")]),
  bullet([t("Co-designed the "), tb("Candle Power NYC pop-up"), t(" for Yankee Candle using the 5E experience framework, exceeding impressions targets by "), tb("400%"), t(" and winning the "), tb("PR Platinum Award (2018)"), t(". Featured in a WGSN trend report and expanded to a Massachusetts retail location.")]),
  bullet([t("Drove digital commerce innovation for Yankee Candle, launching the company's "), tb("first subscription offering"), t("—a sold-out launch in under two months that prompted a 3× increase in projected demand.")]),
  bullet([t("Shaped brand experiences across the Newell portfolio: led the "), tb("Paper Mate packaging refresh"), t(", contributed to the "), tb("Sharpie redesign"), t(", and supported the global "), tb("Newell Brands corporate rebrand"), t(".")]),
  bullet([t("Led vision strategy across "), tb("5 business divisions"), t(", influencing "), tb("40+ brand experiences"), t("; presented eCommerce vision to the CEO, who presented it to the Board of Directors.")]),

  spacer(60),

  // Michigan Fitness Foundation
  companyRow("Michigan Fitness Foundation", "1 year"),
  subroleRow("Digital Designer", "2013 – 2014"),
  bullet([t("Designed and maintained seven CMS-powered websites, implementing current web standards and SEO guidelines.")]),

  spacer(80),

  // Speaking
  secHead("Speaking & External Presence"),
  new Paragraph({
    spacing: { before: 60, after: 20 },
    children: [tb("New York University")]
  }),
  subBullet("Wagner Graduate School — Panelist, AI & Data Strategy for Social Impact; coached graduate students on stakeholder confidence and decision-enabling strategy."),
  subBullet("Guest lecturer, Fall 2026 (repeat engagement)."),
  subBullet("eSchool 4 Girls — Volunteer, speaker, and judge for student entrepreneurship pitch presentations."),
  bullet([tb("IBM / Princeton University TigerTrek"), t(" — Panelist for Princeton's top 20 entrepreneurial students on interdisciplinary careers and entrepreneurial thinking inside enterprise.")]),
  bullet([tb("Ferris State University, College of Business"), t(" — Alumni panelist, Design program graduates session (Class of 2020 & 2023).")]),
  bullet([tb("ATL Onboarding, Austin TX"), t(" — Featured speaker (invitation from IBM VP of Tech Sales); delivered practical \"art of storytelling\" guidance to Account Technical Leaders.")]),
  bullet([tb("AIGA New York"), t(" — Member since 2016.")]),
];

// ── Right column (sidebar) content ───────────────────────────────────────────
const rightCol = [
  // Career Highlights
  sideHead("Career Highlights"),
  ...metricBlock("$50.6M", "Won revenue, IBM FSM account-planning program (Q1 2025)"),
  ...metricBlock("$750K+", "Opportunity value, IBM Social Impact AI program (2022–2024)"),
  ...metricBlock("Top 7 / 2,000", "IBM Client Engineering practitioners worldwide (2024)"),
  ...metricBlock("+111%", "Increase in AI-generated code post-upskilling program"),
  ...metricBlock("7 markets", "Adopted the State Street AI innovation playbook"),
  ...metricBlock("400%", "Exceeded impressions target — Candle Power NYC pop-up"),

  // Awards
  sideHead("Awards"),
  ...awardBlock("IBM Artemis Program", "Top 7 of 2,000 CE practitioners — selected by worldwide CE VP (2024)"),
  ...awardBlock("PR Platinum Award", "Candle Power / Yankee Candle NYC pop-up — innovative storytelling & results (2018)"),
  ...awardBlock("AIGA West Michigan Archives", "CMF storytelling & experience design — accepted for archival recognition"),

  // Core Skills
  sideHead("Core Skills"),
  ...skillGroup("Creative & Experience Design", ["Visual storytelling", "Creative direction", "Immersive experience design", "Customer journey design", "Interactive demonstrations", "Data visualization", "Experience frameworks"]),
  ...skillGroup("Innovation & Executive Engagement", ["Executive briefing programs", "Innovation showcases", "Experience Centers", "AI opportunity discovery", "Agentic AI", "Emerging technology"]),
  ...skillGroup("Strategy & Leadership", ["Creative strategy", "Cross-functional leadership", "Agency & partner direction", "Stakeholder alignment", "Executive presence"]),
  ...skillGroup("Brand & Communication", ["Brand strategy", "Keynote & presentation design", "Thought leadership", "Visual narrative"]),
  ...skillGroup("Tools & Platforms", ["Figma", "Mural", "Salesforce", "Jira", "Monday.com"]),

  // Testimonials
  sideHead("Testimonials"),
  ...testimony("You were such a motivated leader during this last quarter of Artemis. With many headwinds and low participation, you still pushed us forward. Your leadership did not go unnoticed by me!", "Takashi Ohno — Artemis Peer, IBM"),
  ...testimony("I appreciate you consistently leading with empathy and collecting perspectives from across the market. The momentum and focus on meaningful impact in our initiatives is exactly what we need right now.", "Amy Lewis — Canada Design Market Lead, IBM"),
  ...testimony("The Career Progression initiative is a rigorous, long-term guide with lasting value for the CE design community.", "Misti Peinado — Strategy & Transformation Design Lead, EMEA, IBM"),
];

// ── Pad shorter column so table rows balance ─────────────────────────────────
function makeCells(leftParagraphs, rightParagraphs) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: LEFT_W, type: WidthType.DXA },
        borders: noBorders,
        children: leftParagraphs
      }),
      new TableCell({
        width: { size: GAP_W, type: WidthType.DXA },
        borders: noBorders,
        children: [new Paragraph({ children: [new TextRun("")] })]
      }),
      new TableCell({
        width: { size: RIGHT_W, type: WidthType.DXA },
        borders: noBorders,
        children: rightParagraphs
      })
    ]
  });
}

const bodyTable = new Table({
  columnWidths: [LEFT_W, GAP_W, RIGHT_W],
  rows: [makeCells(leftCol, rightCol)]
});

// ── Document ──────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "main-bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 300, hanging: 200 } } } }]
      },
      {
        reference: "sub-bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 200 } } } }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: BLACK } } }
  },
  sections: [{
    properties: {
      page: { margin: { top: 864, right: 864, bottom: 864, left: 864 } }
    },
    children: [
      // ── HEADER ──
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [new TextRun({ text: "Katie Shimmin", bold: true, size: 56, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLACK } },
        children: [new TextRun({ text: "Enterprise Design Strategy  ·  AI Innovation  ·  Organizational Transformation  ·  Experience Design", size: 20, color: ACCENT, font: "Arial" })]
      }),
      new Paragraph({
        spacing: { before: 50, after: 160 },
        children: [
          new TextRun({ text: "Kathrynshimmin@gmail.com", size: 19, color: MUTED, font: "Arial" }),
          new TextRun({ text: "   ·   917-667-5716", size: 19, color: MUTED, font: "Arial" }),
          new TextRun({ text: "   ·   New York, NY", size: 19, color: MUTED, font: "Arial" })
        ]
      }),
      // ── BODY TABLE ──
      bodyTable
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("katie_shimmin_resume_v2.docx", buffer);
  console.log("Done: katie_shimmin_resume_v2.docx");
});
