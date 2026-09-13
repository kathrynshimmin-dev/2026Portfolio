const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Colours ──────────────────────────────────────────────────────────────────
const BLUE   = "3B82D4";
const BLACK  = "1F2328";
const MUTED  = "57606A";
const BORDER = "E5E7EB";
const SURF   = "F7F8FA";

// ── Borders ──────────────────────────────────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const ruledBorder = { style: BorderStyle.SINGLE, size: 4, color: BORDER };

// ── Helpers ───────────────────────────────────────────────────────────────────
function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER } },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, size: 18, color: MUTED, font: "Arial", characterSpacing: 80 })
    ]
  });
}

function roleHeader(title, dates) {
  return new Paragraph({
    spacing: { before: 200, after: 40 },
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    children: [
      new TextRun({ text: title, bold: true, size: 26, color: BLACK, font: "Arial" }),
      new TextRun({ text: "\t" + dates, size: 22, color: MUTED, font: "Arial" })
    ]
  });
}

function roleCompany(text) {
  return new Paragraph({
    spacing: { before: 0, after: 100 },
    children: [new TextRun({ text, size: 24, color: BLUE, bold: true, font: "Arial" })]
  });
}

function bullet(text, boldPrefix) {
  const children = [];
  if (boldPrefix) {
    children.push(new TextRun({ text: boldPrefix, bold: true, size: 22, font: "Arial", color: BLACK }));
    children.push(new TextRun({ text, size: 22, font: "Arial", color: BLACK }));
  } else {
    children.push(new TextRun({ text, size: 22, font: "Arial", color: BLACK }));
  }
  return new Paragraph({
    numbering: { reference: "bullet-list", level: 0 },
    spacing: { before: 40, after: 40 },
    children
  });
}

function spacer(pt = 120) {
  return new Paragraph({ spacing: { before: pt, after: 0 }, children: [new TextRun("")] });
}

// ── Document ──────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 240 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 24, color: BLACK } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 28, bold: true, color: BLACK, font: "Arial" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 0 }
      }
    ]
  },
  sections: [{
    properties: {
      page: { margin: { top: 1008, right: 1008, bottom: 1008, left: 1008 } }
    },
    children: [

      // ── NAME & TITLE ─────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Katie Shimmin", bold: true, size: 52, font: "Arial", color: BLACK })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLACK } },
        children: [
          new TextRun({ text: "Design Director  ·  Enterprise Design Strategy  ·  AI Innovation  ·  Organizational Transformation", size: 22, color: BLUE, font: "Arial" })
        ]
      }),
      new Paragraph({
        spacing: { before: 60, after: 200 },
        children: [
          new TextRun({ text: "katieshimmin@ibm.com  ·  Americas Design Lead, IBM Client Engineering", size: 20, color: MUTED, font: "Arial" })
        ]
      }),

      // ── SUMMARY ──────────────────────────────────────────────────────────
      sectionHeading("Summary"),
      new Paragraph({
        spacing: { before: 100, after: 200 },
        children: [new TextRun({
          text: "Design Director with a record of scaling complex programs across IBM's largest enterprise accounts, nonprofits, and internal organizations. Leads through service and trust—building operating models, frameworks, and human-centered experiences that produce measurable pipeline, capability growth, and lasting organizational change. Proven at the intersection of strategic vision, AI innovation, and cross-functional alignment, with a background spanning consumer brand experience, digital strategy, and enterprise transformation at IBM scale.",
          size: 22, font: "Arial", color: BLACK
        })]
      }),

      // ── EXPERIENCE ───────────────────────────────────────────────────────
      sectionHeading("Experience"),

      // IBM
      roleHeader("Americas Design Lead", "2022 – Present"),
      roleCompany("IBM Client Engineering"),
      bullet("Named ", "Top 7 of 2,000"),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: " IBM Client Engineering practitioners worldwide (2024) and selected for the Artemis Program cohort tasked with transforming how Client Engineering operates as a business.", size: 22, font: "Arial", color: BLACK })]
      }),

      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Lead design strategy, program design, and facilitation enablement across IBM's Financial Services, Commercial, and Social Impact markets, coordinating hundreds of designers across six global markets.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Built and scaled a ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "client-centered account-planning operating model", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " for IBM FSM across 63 accounts in one quarter, producing ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "$50.6M in won revenue", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " and 1,000+ surfaced opportunities; framework adopted by two additional IBM markets.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Designed and directed a ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "State Street AI innovation program", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " (60+ participants, 9 qualified AI use cases) codified into a playbook mandated across 7 IBM markets by a General Manager the following year.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Created and led the ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Bobathon", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: ", a role-based AI upskilling program for ~80 IBM practitioners: 93.8% post-event confidence, 111.73% increase in AI-generated code two months post-event, and 25% of participants advancing a full proficiency level.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Originated a two-year ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "nonprofit AI adoption framework", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " for IBM Social Impact, partnering with data.org and Echoing Green to generate ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "$750K+ in opportunity value", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " and 36+ AI use cases across housing, healthcare, education, and economic mobility sectors.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Created ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Career Pathways & Progression", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " research scaled to 200+ Innovation Designers worldwide in 2025, mapping IBM core and adjacent design roles with skills gaps for long-term career guidance.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Launched IBM Client Engineering's ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "first-ever Women in Tech event", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " and facilitated a Design Thinking 101 workshop for State Street's Professional Women's Network, directly enabling the State Street hackathon relationship.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Enabled ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "15 designers across 6 markets", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " to attend Adobe Creative Jam and Figma Config through business cases and budget advocacy.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 200 },
        children: [
          new TextRun({ text: "Editorial Lead for the IBM FSM Client Engineering newsletter (", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "1,000+ practitioners", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: ") since 2023; founded and hosted an internal podcast connecting four nationwide design hubs.", size: 22, font: "Arial", color: BLACK })
        ]
      }),

      // Newell Brands
      roleHeader("Senior Experience Designer & Vision Strategist", "2015 – 2022"),
      roleCompany("Newell Brands  (Yankee Candle, Sunbeam, Crock-Pot, Parker Pens)"),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Led vision strategy across ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "5 business divisions", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: ", influencing ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "40+ brand experiences", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "; presented eCommerce vision to the CEO, who presented it to the Board of Directors.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Co-designed the ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Candle Power NYC pop-up", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " for Yankee Candle using the 5E experience framework, exceeding social and media impressions targets by ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "400%", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " and winning the ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "PR Platinum Award (2018)", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: ".", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Experience featured in a ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "WGSN trend report", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " and expanded to an additional Massachusetts retail location.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Created a CMF (Color, Material, Finish) business-impact exhibit and publication accepted into the ", size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "AIGA West Michigan Archives", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: ".", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 200 },
        children: [
          new TextRun({ text: "Produced future-state vision strategies for IoT, connected wellness, eCommerce, premium brand, and meal-planning platforms—repositioning hardware as platforms and informing connected-product roadmap conversations across brands.", size: 22, font: "Arial", color: BLACK })
        ]
      }),

      // ── SPEAKING ─────────────────────────────────────────────────────────
      sectionHeading("Speaking & External Presence"),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "NYU Wagner Graduate School", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Panelist, AI & Data Strategy for Social Impact; coached graduate students on stakeholder confidence and decision-enabling strategy.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "IBM / Princeton University TigerTrek", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Panelist for Princeton's top 20 entrepreneurial students on interdisciplinary careers and entrepreneurial thinking inside enterprise.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "Ferris State University, College of Business", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Alumni panelist, Design program graduation (Class of 2023).", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "New York University", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Invited guest lecturer, Fall 2026 (repeat engagement).", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "ATL Onboarding, Austin TX", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Featured speaker (invitation from IBM VP of Technology); delivered practical storytelling guidance to Account Technical Leaders.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "eSchool 4 Girls", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Volunteer, speaker, and judge for student entrepreneurship pitch presentations at NYU.", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 200 },
        children: [
          new TextRun({ text: "AIGA New York", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: " — Member since 2016.", size: 22, font: "Arial", color: BLACK })
        ]
      }),

      // ── AWARDS ───────────────────────────────────────────────────────────
      sectionHeading("Awards & Recognition"),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 60, after: 40 },
        children: [
          new TextRun({ text: "IBM Artemis Program — Top 7 of 2,000", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "  |  Selected by worldwide Client Engineering VP; tasked with transforming how Client Engineering operates as a business (2024).", size: 22, font: "Arial", color: MUTED })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: "PR Platinum Award", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "  |  Candle Power / Yankee Candle NYC pop-up — innovative storytelling & measurable results (September 2018).", size: 22, font: "Arial", color: MUTED })
        ]
      }),
      new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { before: 40, after: 200 },
        children: [
          new TextRun({ text: "AIGA West Michigan Archives", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "  |  CMF storytelling & experience design — accepted for archival recognition.", size: 22, font: "Arial", color: MUTED })
        ]
      }),

      // ── SKILLS ───────────────────────────────────────────────────────────
      sectionHeading("Core Skills"),
      new Paragraph({
        spacing: { before: 100, after: 60 },
        children: [
          new TextRun({ text: "Design Leadership:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Strategic program design · Organizational transformation · Career development · Change management · Facilitator enablement", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 60 },
        children: [
          new TextRun({ text: "AI & Innovation:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "AI opportunity discovery · watsonx.ai · Watson Discovery · MVP scoping · Use-case prioritization · Agentic AI", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 60 },
        children: [
          new TextRun({ text: "Experience & Service Design:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Design thinking · Workshop design · Experience frameworks · Journey mapping · Retail & pop-up design", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 60 },
        children: [
          new TextRun({ text: "Strategy & Operations:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Executive storytelling · Pipeline generation · Stakeholder alignment · Operating model design · Analytics & measurement", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 60 },
        children: [
          new TextRun({ text: "Brand & Digital:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Brand strategy · eCommerce vision · IoT strategy · CMF · Environmental graphics", size: 22, font: "Arial", color: BLACK })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 200 },
        children: [
          new TextRun({ text: "Tools:  ", bold: true, size: 22, font: "Arial", color: BLACK }),
          new TextRun({ text: "Mural · Figma · IBM Bob · watsonx", size: 22, font: "Arial", color: BLACK })
        ]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("katie_shimmin_resume.docx", buffer);
  console.log("Done: katie_shimmin_resume.docx");
});
