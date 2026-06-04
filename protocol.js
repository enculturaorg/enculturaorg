const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, VerticalAlign, PageNumber, PageBreak, Header, Footer,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Helpers ──────────────────────────────────────────────────────────────────

const FONT = "Georgia";
const MONO = "Courier New";
const COLOR_INK   = "1A1A2E";
const COLOR_DEEP  = "0F3460";
const COLOR_MID   = "16213E";
const COLOR_RULE  = "CCCCCC";
const COLOR_GOLD  = "8B6914";
const COLOR_TEAL  = "1A6B6B";
const COLOR_AMBER = "7A5200";
const COLOR_RED_W = "7A1A1A";
const COLOR_GREEN_W = "1A5C2A";
const COLOR_ORANGE_W = "7A3A00";

function sp(txt, opts = {}) {
  return new TextRun({ text: txt, font: FONT, size: 22, color: COLOR_INK, ...opts });
}

function mono(txt) {
  return new TextRun({ text: txt, font: MONO, size: 20, color: COLOR_TEAL });
}

function para(children, opts = {}) {
  if (typeof children === 'string') {
    children = [sp(children)];
  }
  return new Paragraph({
    spacing: { before: 100, after: 200, line: 320, lineRule: "auto" },
    ...opts,
    children
  });
}

function headPara(text, level, color = COLOR_INK) {
  const sizes = { 1: 40, 2: 30, 3: 26 };
  return new Paragraph({
    heading: level,
    spacing: { before: 480, after: 200 },
    children: [new TextRun({ text, font: FONT, size: sizes[level === HeadingLevel.HEADING_1 ? 1 : level === HeadingLevel.HEADING_2 ? 2 : 3], bold: true, color })]
  });
}

function rule() {
  return new Paragraph({
    spacing: { before: 240, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_RULE, space: 1 } },
    children: [sp(" ")]
  });
}

function blockquote(text) {
  return new Paragraph({
    spacing: { before: 200, after: 200, line: 320, lineRule: "auto" },
    indent: { left: 720, right: 720 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: COLOR_TEAL, space: 10 } },
    children: [sp(text, { italics: true, color: "444444" })]
  });
}

function articlePara(number, title, body) {
  return [
    new Paragraph({
      spacing: { before: 300, after: 80 },
      children: [
        sp(`Article ${number}. `, { bold: true, color: COLOR_DEEP }),
        sp(title, { bold: true, italics: true, color: COLOR_DEEP })
      ]
    }),
    para(body)
  ];
}

function sectionLabel(text) {
  return new Paragraph({
    spacing: { before: 360, after: 120 },
    children: [sp(text, { bold: true, allCaps: true, size: 18, color: "888888" })]
  });
}

function badge(label, color) {
  // A colored inline label — rendered as a small bold run
  return new TextRun({ text: ` [${label}] `, font: FONT, size: 20, bold: true, color });
}

function bullet(text, level = 0) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 300, lineRule: "auto" },
    numbering: { reference: "bullets", level },
    children: typeof text === 'string' ? [sp(text)] : text
  });
}

function numItem(text, level = 0) {
  return new Paragraph({
    spacing: { before: 80, after: 80, line: 300, lineRule: "auto" },
    numbering: { reference: "numbers", level },
    children: typeof text === 'string' ? [sp(text)] : text
  });
}

// ── Shared border ─────────────────────────────────────────────────────────────
const cellBorder = {
  top:    { style: BorderStyle.SINGLE, size: 1, color: COLOR_RULE },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: COLOR_RULE },
  left:   { style: BorderStyle.SINGLE, size: 1, color: COLOR_RULE },
  right:  { style: BorderStyle.SINGLE, size: 1, color: COLOR_RULE },
};

function twoCol(left, right, leftW = 3000, rightW = 6360) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [leftW, rightW],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: cellBorder,
        width: { size: leftW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        shading: { fill: "F5F5F0", type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [sp(left, { bold: true, size: 20 })] })]
      }),
      new TableCell({
        borders: cellBorder,
        width: { size: rightW, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [sp(right, { size: 20 })] })]
      })
    ]})]
  });
}

function worldspaceRow(code, name, color, drive, shadow, protocol) {
  const w1 = 900, w2 = 1200, w3 = 2420, w4 = 2420, w5 = 2420;
  return new TableRow({ children: [
    new TableCell({
      borders: cellBorder, width: { size: w1, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      shading: { fill: color + "22", type: ShadingType.CLEAR },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [sp(code, { bold: true, color, size: 20 })] })]
    }),
    new TableCell({
      borders: cellBorder, width: { size: w2, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [sp(name, { bold: true, size: 20 })] })]
    }),
    new TableCell({
      borders: cellBorder, width: { size: w3, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [sp(drive, { size: 19 })] })]
    }),
    new TableCell({
      borders: cellBorder, width: { size: w4, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [sp(shadow, { size: 19, italics: true, color: "666666" })] })]
    }),
    new TableCell({
      borders: cellBorder, width: { size: w5, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [sp(protocol, { size: 19 })] })]
    }),
  ]});
}

// ── DOCUMENT ──────────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 40, bold: true, font: FONT, color: COLOR_INK },
        paragraph: { spacing: { before: 600, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: FONT, color: COLOR_DEEP },
        paragraph: { spacing: { before: 400, after: 180 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, italics: true, font: FONT, color: COLOR_MID },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]},
      { reference: "numbers", levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]},
      { reference: "articles", levels: [
          { level: 0, format: LevelFormat.DECIMAL, text: "§%1", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 0, hanging: 0 } } } },
        ]},
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_RULE, space: 4 } },
          children: [sp("Encultura Protocol Specification v0.1 — did:web:encultura.org", { size: 18, color: "888888" })]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_RULE, space: 4 } },
          children: [
            sp("© Encultura Guild  ·  Released under CC BY-SA 4.0  ·  Page ", { size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "888888" }),
          ]
        })
      ]})
    },
    children: [

      // ── COVER ──────────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 960, after: 120 },
        children: [sp("ENCULTURA", { size: 52, bold: true, color: COLOR_TEAL, allCaps: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [sp("Protocol Specification", { size: 34, italics: true, color: COLOR_INK })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [sp("Version 0.1 — Genesis Document", { size: 22, color: "666666" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [mono("did:web:encultura.org")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 480 },
        children: [sp("A Navigational Ecology for the Evolution of Human Worldspaces", { size: 22, italics: true, color: "555555" })]
      }),

      rule(),

      blockquote(
        "This document is a Protocol Specification, not a product launch. It establishes the verifiable grammar " +
        "of a long-horizon project — a 5 to 10 year pedagogical architecture intended to steward the development " +
        "of navigational capacity across a network of participants. It supersedes no prior authority. It creates " +
        "a container. What grows inside the container is the work."
      ),

      rule(),

      // ── PREAMBLE ───────────────────────────────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Preamble", { size: 40, bold: true, color: COLOR_INK })] }),

      para(
        "We are writing at the end of the statistical age. The infrastructure of collective life — digital platforms, " +
        "algorithmic governance, artificial intelligence — was built on a first-order logic: that truth is measurable, " +
        "that humans are classifiable, and that a system's health can be read from its efficiency metrics. This logic " +
        "has produced, at the second order, the systematic degradation of the cognitive and relational substrate upon " +
        "which genuine human participation depends. At the third order, it has eroded the deliberative commons itself."
      ),
      para(
        "The Encultura Protocol is a response to this triple failure. It does not seek to replace the statistical layer " +
        "with a new extraction engine dressed in ecological language. It seeks to hard-code a different orientation: " +
        "one in which the flourishing of life — of persons, communities, and the relationships between them — is the " +
        "terminal value from which all protocol mechanics derive their legitimacy."
      ),
      para(
        "This specification is the first verifiable artifact produced under the identity anchor " +
        "did:web:encultura.org. It is signed, in the sense that it is linked to that identity and may be " +
        "cryptographically verified against it. It is not a manifesto. It is a grammar — the syntax that " +
        "participants opt into when they choose to engage with the Encultura network."
      ),

      // ── PART I: FOUNDATIONS ────────────────────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Part I: Foundational Commitments", { size: 40, bold: true, color: COLOR_INK })] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§1  The Three Axioms", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "All protocol mechanics described in this specification derive from three foundational axioms. " +
        "These axioms are not values statements — they are structural constraints on the design space."
      ),

      new Paragraph({ spacing: { before: 200, after: 60 }, children: [sp("Axiom I — Life Priority", { bold: true, color: COLOR_TEAL })] }),
      para(
        "The protocol prioritizes the flourishing of life above the optimization of measurable outputs. " +
        "Where statistical truth (market signal, ledger, efficiency metric) conflicts with network truth " +
        "(emergent relational health, developmental trajectory, systemic vitality), network truth governs. " +
        "This axiom must be encoded in mechanism, not merely stated in language."
      ),

      new Paragraph({ spacing: { before: 200, after: 60 }, children: [sp("Axiom II — Navigational Plurality", { bold: true, color: COLOR_TEAL })] }),
      para(
        "No single worldspace is the terminal destination of human development. The memeplexes designated " +
        "Red, Amber, Orange, Green, and Teal in this specification are navigational registers, not ranks. " +
        "The protocol is designed to be legible and functional for participants primarily inhabiting any " +
        "of these registers, while providing developmental affordances toward higher-order integration."
      ),

      new Paragraph({ spacing: { before: 200, after: 60 }, children: [sp("Axiom III — Amber Verification Before Teal Aspiration", { bold: true, color: COLOR_TEAL })] }),
      para(
        "A system cannot achieve genuine Teal integration — networked, life-centered, integral awareness — " +
        "without first establishing a foundation of Amber verification: rule-based commitment, demonstrable " +
        "accountability, and the subordination of individual sovereignty to shared governance structures. " +
        "The protocol does not pretend to have achieved Teal. It works toward Teal by building Amber containers " +
        "within which the developmental journey can safely occur."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§2  Worldspace Taxonomy", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The following table establishes the operative taxonomy of worldspaces used throughout this " +
        "specification. Each register is described by its primary drive, its characteristic failure mode " +
        "(shadow), and the protocol's intended response to participants inhabiting it. This is a navigational " +
        "tool, not a judgment taxonomy."
      ),

      // Worldspace table
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [900, 1200, 2420, 2420, 2420],
        rows: [
          // Header row
          new TableRow({ tableHeader: true, children: [
            ...[["Code", 900], ["Name", 1200], ["Primary Drive", 2420], ["Shadow / Failure Mode", 2420], ["Protocol Response", 2420]].map(([txt, w]) =>
              new TableCell({
                borders: cellBorder, width: { size: w, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                shading: { fill: "1A1A2E", type: ShadingType.CLEAR },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [sp(txt, { bold: true, color: "FFFFFF", size: 19 })] })]
              })
            )
          ]}),
          worldspaceRow("RED",    "Sovereignty",  COLOR_RED_W,    "Power, immediacy, individual will, vitality",                         "Predation, exploitation, refusal of shared-world obligation",               "Channel through Amber commitment; subordinate to Logos without suppressing drive"),
          worldspaceRow("AMBER",  "Order",        COLOR_AMBER,    "Rule, belonging, hierarchy, loyalty to shared norms",                 "Rigidity, exclusion, inability to metabolize change",                       "Use as foundational container grammar; the baseline of verification"),
          worldspaceRow("ORANGE", "Achievement",  COLOR_ORANGE_W, "Competence, optimization, meritocratic advancement",                 "Extraction disguised as excellence; metrics colonizing meaning",             "Acknowledge statistical outputs as one signal among several; discipline with network health metrics"),
          worldspaceRow("GREEN",  "Care",         COLOR_GREEN_W,  "Empathy, inclusion, systemic critique, relational awareness",         "Paralytic non-discrimination; vulnerability to predatory Red capture",       "Integrate the discriminating function; Green must learn to name harm without abandoning care"),
          worldspaceRow("TEAL",   "Integration",  COLOR_TEAL,     "Ecological awareness, developmental stewardship, integral complexity","Broken Teal: ecosystemic rhetoric over Orange protocol mechanics",          "Hard-code the tension between Statistical and Network Truth into protocol syntax"),
        ]
      }),

      para(" "),

      // ── PART II: GOVERNANCE ARCHITECTURE ──────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Part II: Governance Architecture", { size: 40, bold: true, color: COLOR_INK })] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§3  The Anjuman (Guild Council)", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The Anjuman — from the Persian tradition of an assembly of cultivated peers — is the foundational " +
        "governance node of the Encultura network. It is not a board of directors, a DAO, or a community of " +
        "practice in the conventional sense. It is the Architect node: the initial testbed for Amber-verified " +
        "interaction, the steward of the protocol's developmental integrity, and the gardener of the conditions " +
        "under which navigational capacity can grow."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("3.1  Composition and Entry", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),
      para(
        "The founding Anjuman consists of a minimum of three and a maximum of twelve Members. Membership " +
        "is not determined by technical skill, social status, or prior achievement — though competence " +
        "is necessary and will be evidenced through the entry process. Membership is determined by demonstrated " +
        "navigational capacity: the ability to hold multiple worldspace registers simultaneously, to act from " +
        "the higher without suppressing the lower, and to distinguish authentic development from its simulation."
      ),
      para("Entry to the Anjuman requires:"),
      bullet("Endorsement by two existing Members who attest to the candidate's navigational capacity through direct co-working experience."),
      bullet("Submission of a Protocol Reflection: a written document (minimum 1,500 words) in which the candidate articulates their understanding of the three Axioms, identifies the worldspace register from which they most frequently operate, and describes a specific instance in which they failed to transcend that register — and what they learned from the failure."),
      bullet("A one-cycle probationary period (see §6: Temporal Architecture) during which the candidate participates in Anjuman deliberations without voting rights."),
      bullet("Ratification by consensus of the full Anjuman at the close of the probationary cycle."),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("3.2  Anjuman Responsibilities", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),
      para("The Anjuman bears collective responsibility for:"),
      bullet("Maintaining the integrity of the Protocol Specification, including proposing and ratifying amendments (see §5: Amendment Procedure)."),
      bullet("Stewarding the Identity Anchor — did:web:encultura.org — and authorizing the signing of documents, attestations, and protocol artifacts produced under that identity."),
      bullet("Monitoring and publishing Network Health Metrics (see §4) on a cadence no less frequent than one full cycle."),
      bullet("Facilitating the pedagogical architecture described in §7, including the design of productive friction that advances navigational development without producing developmental collapse."),
      bullet("Naming harm when harm is present. The Anjuman's Green care function does not supersede its discriminating function. The failure to name predatory behavior is a protocol violation equivalent in severity to an act of predation."),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("3.3  Decision Grammar", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),
      para("The Anjuman operates with a tiered decision grammar:"),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2400, 2400, 2400, 2160],
        rows: [
          new TableRow({ tableHeader: true, children: [
            ...[["Decision Type", 2400], ["Threshold", 2400], ["Notice Required", 2400], ["Reversibility", 2160]].map(([txt, w]) =>
              new TableCell({
                borders: cellBorder, width: { size: w, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                shading: { fill: "16213E", type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [sp(txt, { bold: true, color: "FFFFFF", size: 19 })] })]
              })
            )
          ]}),
          ...[
            ["Operational (day-to-day)", "Simple majority", "None", "Immediately reversible"],
            ["Pedagogical design change", "Supermajority (⅔)", "One full week", "Reversible within one cycle"],
            ["Protocol amendment", "Consensus", "One full cycle", "Requires new amendment process"],
            ["Identity Anchor action", "Consensus + external attestation", "One full cycle", "Irreversible; append-only log"],
            ["Membership change", "Consensus", "Two weeks", "Membership exit is permanent"],
          ].map(([type, threshold, notice, rev]) =>
            new TableRow({ children: [
              ...[type, threshold, notice, rev].map((txt, i) =>
                new TableCell({
                  borders: cellBorder,
                  width: { size: [2400, 2400, 2400, 2160][i], type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 120, right: 120 },
                  children: [new Paragraph({ children: [sp(txt, { size: 19 })] })]
                })
              )
            ]})
          )
        ]
      }),

      para(" "),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§4  Network Health Metrics", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "This section operationalizes the hard-coded tension between Statistical Truth and Network Truth " +
        "that distinguishes Functional Teal from Broken Teal. The following metrics are not aspirational " +
        "indicators — they are governance inputs. If a Network Health Metric falls below its floor threshold, " +
        "the relevant Statistical optimization function is suspended until the metric recovers."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("4.1  The Seven Network Health Indicators (NHI)", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1800, 3600, 1980, 1980],
        rows: [
          new TableRow({ tableHeader: true, children: [
            ...[["Indicator", 1800], ["Definition", 3600], ["Floor Threshold", 1980], ["Suspension Trigger", 1980]].map(([txt, w]) =>
              new TableCell({
                borders: cellBorder, width: { size: w, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                shading: { fill: "1A6B6B", type: ShadingType.CLEAR },
                children: [new Paragraph({ children: [sp(txt, { bold: true, color: "FFFFFF", size: 19 })] })]
              })
            )
          ]}),
          ...[
            ["NHI-1: Developmental Velocity", "Rate at which participants demonstrate measurable worldspace expansion over a given cycle, assessed through Protocol Reflections and peer attestation.", "≥ 40% of active participants showing movement", "Statistical output metrics suspended if < 20% for two consecutive cycles"],
            ["NHI-2: Predation Index", "Incidence of identified predatory-Red behavior (exploitation, bad-faith participation, deliberate disruption of developmental containers) normalized to network size.", "< 5% of interactions flagged", "New participant onboarding suspended if > 10%"],
            ["NHI-3: Discrimination Capacity", "The network's demonstrated ability to name and respond to harm without collapsing into either punitive Red enforcement or paralytic Green non-discrimination.", "Qualitative; assessed by Anjuman each cycle", "Protocol review mandatory if Green paralysis or Red escalation persists"],
            ["NHI-4: Relational Depth", "Average duration and developmental complexity of peer relationships within the network, measured through co-production artifacts and attested interactions.", "≥ 60% of participants with at least one deep-relational tie", "Expansion paused if < 40%"],
            ["NHI-5: Protocol Legibility", "Degree to which participants across all worldspace registers can articulate their understanding of the protocol grammar.", "Assessed through onboarding reflections and cycle reviews", "Protocol simplification process triggered if < 50% legibility across any register"],
            ["NHI-6: Architect Succession Health", "Degree to which navigational capacity is developing among non-Anjuman participants, assessed as readiness for eventual Anjuman membership.", "≥ 2 candidates in probationary pipeline per cycle", "Pedagogical redesign triggered if succession pipeline empty for two cycles"],
            ["NHI-7: Axiom Coherence", "Degree to which network outputs, decisions, and artifacts can be traced to the three foundational axioms without contradiction.", "Qualitative; reviewed by Anjuman annually", "Full protocol audit triggered if coherence failure identified"],
          ].map(([ind, def, floor, trig]) =>
            new TableRow({ children: [
              ...[ind, def, floor, trig].map((txt, i) =>
                new TableCell({
                  borders: cellBorder,
                  width: { size: [1800, 3600, 1980, 1980][i], type: WidthType.DXA },
                  margins: { top: 80, bottom: 80, left: 120, right: 120 },
                  shading: i === 0 ? { fill: "F0F8F8", type: ShadingType.CLEAR } : undefined,
                  children: [new Paragraph({ children: [sp(txt, { size: 19 })] })]
                })
              )
            ]})
          )
        ]
      }),

      para(" "),

      // ── PART III: PROTOCOL MECHANICS ───────────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Part III: Protocol Mechanics", { size: 40, bold: true, color: COLOR_INK })] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§5  Amendment Procedure", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "This specification is a living document. Its amendment procedure is itself a demonstration of the " +
        "Amber verification principle: the rules for changing the rules are fixed, transparent, and " +
        "independent of the preferences of any individual participant — including the founding Anjuman members."
      ),

      numItem([sp("Proposal. ", { bold: true }), sp("Any Member may propose an amendment by submitting a written Amendment Proposal to the Anjuman. The proposal must specify: (a) the text to be changed; (b) the proposed replacement text; (c) the Axiom or NHI rationale for the change; and (d) the anticipated second- and third-order consequences.")]),
      numItem([sp("Review Cycle. ", { bold: true }), sp("The proposal is held open for one full cycle (see §6) for community commentary. Commentary must be substantive — it must engage with the Axiom rationale and must identify any second-order consequences the proposal has not addressed.")]),
      numItem([sp("Deliberation. ", { bold: true }), sp("The Anjuman deliberates on the proposal and all commentary in an open session. Non-Anjuman participants may observe and submit written questions; they may not vote.")]),
      numItem([sp("Ratification. ", { bold: true }), sp("Amendments require Anjuman consensus (unanimous agreement with right of abstention; abstention does not block consensus). A single objection blocks the amendment and requires re-proposal.")]),
      numItem([sp("Identity Anchor Signing. ", { bold: true }), sp("Ratified amendments are published as a signed artifact under did:web:encultura.org, creating an append-only, publicly verifiable record of the protocol's evolution.")]),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§6  Temporal Architecture", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The protocol operates on three interlocking temporal scales, each corresponding to a different " +
        "register of developmental work. Confusing these scales — treating the developmental timescale as " +
        "if it were the operational timescale — is one of the primary failure modes of Broken Teal systems."
      ),

      twoCol("Operational Cycle", "Duration: 4 weeks. The basic unit of coordination. Includes: task completion, peer attestation, NHI measurement, and operational decision-making."),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [sp(" ")] }),
      twoCol("Developmental Cycle", "Duration: 6 months (13 operational cycles). The unit of pedagogical assessment. Includes: Protocol Reflections, navigational capacity assessment, NHI trend review, and curriculum revision."),
      new Paragraph({ spacing: { before: 80, after: 0 }, children: [sp(" ")] }),
      twoCol("Generational Cycle", "Duration: 5 years (10 developmental cycles). The unit of memetic evolution. Includes: full protocol audit, Axiom coherence review, Anjuman succession, and the assessment of whether the navigational ecology has achieved sufficient distributed capacity to govern itself without the Architect node."),

      para(" "),

      para(
        "The Generational Cycle horizon is not arbitrary. It reflects the empirical timescale required for " +
        "genuine shifts in dominant worldspace register within a community of practice. Protocols that expect " +
        "developmental transformation on shorter timescales are either measuring something other than development " +
        "or are confusing individual learning events with collective memetic evolution."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§7  Pedagogical Architecture", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The Encultura network is, at its core, a pedagogical institution. Its output is not a product, " +
        "a platform, or a proof-of-concept. Its output is navigational capacity — the distributed capacity " +
        "of participants to hold and act from increasingly integrated worldspace awareness. All other " +
        "protocol mechanics serve this terminal function."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("7.1  Productive Friction Design", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),

      para(
        "The protocol is designed to introduce productive friction: structured challenges that exceed a " +
        "participant's current navigational register, creating the disequilibrium necessary for development, " +
        "without producing the overwhelming disorientation that forecloses it. The design of productive " +
        "friction is the Anjuman's primary craft. It requires simultaneous attention to:"
      ),
      bullet([sp("Zone of proximal development", { italics: true }), sp(": The challenge must be beyond the participant's current capacity but within reach of their developing capacity with scaffolding.")]),
      bullet([sp("Worldspace translation", { italics: true }), sp(": The same challenge must be legible in multiple registers — a participant primarily in Orange must be able to engage with it through achievement language, even as the challenge is designed to invite Green perspective-taking.")]),
      bullet([sp("Failure protocols", { italics: true }), sp(": The architecture must include explicit, non-punitive responses to developmental failure. Failure is a learning event, not a verdict on navigational capacity.")]),
      bullet([sp("The shadow integration imperative", { italics: true }), sp(": At least one element of every developmental challenge must require participants to work with the shadow of their dominant register — the Orange achiever must encounter the human cost of optimization; the Green empath must encounter the network health consequences of non-discrimination.")]),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("7.2  The Protocol Reflection as Core Practice", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),

      para(
        "The Protocol Reflection — written, substantive, published within the network — is the primary " +
        "instrument of pedagogical assessment. It is not a performance review. It is a navigational log: " +
        "a participant's honest account of where they have been, what they have encountered, and how they " +
        "have been changed by it. A Protocol Reflection that demonstrates only competence and success is, " +
        "by definition, incomplete. The Anjuman will specifically attend to:"
      ),
      bullet("Evidence of genuine developmental friction: instances where the participant was genuinely destabilized, not merely intellectually challenged."),
      bullet("Quality of shadow engagement: the degree to which the participant can name and work with the failure modes of their own dominant register."),
      bullet("Worldspace range: the ability to articulate the perspectives of registers other than one's own, without collapsing them into caricature or projecting one's own register onto them."),
      bullet("Protocol coherence: the ability to trace one's own actions and decisions back to the three Axioms, including — especially — instances where one failed to act in accordance with them."),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§8  The Role of Artificial Intelligence", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "Artificial intelligence tools may be used within the Encultura network subject to the following " +
        "constraints, which derive directly from the framework articulated in the foundational essay."
      ),

      new Paragraph({ spacing: { before: 200, after: 80 }, children: [sp("8.1  AI as Strategic Mimicry Tool — Not Oracle", { bold: true, color: COLOR_TEAL })] }),
      para(
        "No AI output shall be treated as authoritative truth within any Anjuman deliberation or network " +
        "governance process. AI outputs are explicitly designated as statistical approximations — sophisticated " +
        "pattern-matching that can expand the range of perspectives available to human deliberation, but " +
        "that cannot substitute for the authentic developmental work that deliberation requires. " +
        "The authority claim of any AI system is hereby structurally refused."
      ),

      new Paragraph({ spacing: { before: 200, after: 80 }, children: [sp("8.2  Permitted Uses", { bold: true, color: COLOR_TEAL })] }),
      bullet("Worldspace translation: using AI to articulate how a proposed decision would read to participants primarily inhabiting different registers."),
      bullet("Second- and third-order consequence modeling: using AI to surface systemic effects that human deliberation may not have considered."),
      bullet("Pedagogical material generation: using AI to draft challenges, reflections prompts, and translation texts, subject to Anjuman review and approval."),
      bullet("Protocol legibility testing: using AI to identify sections of the specification that are unclear or inaccessible to participants in specific worldspace registers."),

      new Paragraph({ spacing: { before: 200, after: 80 }, children: [sp("8.3  Prohibited Uses", { bold: true, color: COLOR_RED_W })] }),
      bullet("Using AI outputs as the primary basis for any governance decision without explicit human deliberation."),
      bullet("Delegating navigational capacity assessment — of any participant, including oneself — to an AI system."),
      bullet("Using AI to simulate Protocol Reflections or any other first-person developmental account."),
      bullet("Treating AI-generated content as signed artifacts under the did:web:encultura.org identity anchor without full Anjuman review and explicit consensus."),

      // ── PART IV: IDENTITY ANCHOR ───────────────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Part IV: Identity Anchor and Verifiable Grammar", { size: 40, bold: true, color: COLOR_INK })] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§9  The did:web:encultura.org Anchor", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The Decentralized Identifier (DID) did:web:encultura.org serves as the protocol's identity anchor: " +
        "the cryptographic root to which all official Encultura artifacts, amendments, and attestations are " +
        "linked. This anchor is not a brand identity or a web presence. It is a verifiable commitment device — " +
        "a mechanism by which the authenticity and integrity of any document produced under this identity " +
        "can be independently verified."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("9.1  What the Anchor Signs", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),
      para("The following classes of artifact are eligible for signing under the identity anchor:"),
      bullet("Protocol Specification documents (this document and its successors)."),
      bullet("Ratified amendments to the Protocol Specification."),
      bullet("Guild Constitution documents establishing new node Anjumans within the network."),
      bullet("Formal Attestations of navigational capacity for Anjuman membership transitions."),
      bullet("Network Health Metric reports published at the close of each Developmental Cycle."),

      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [sp("9.2  What the Anchor Does Not Sign", { size: 25, bold: true, italics: true, color: COLOR_MID })] }),
      para(
        "The identity anchor does not sign individual participants' Protocol Reflections, operational " +
        "artifacts, or AI-generated content. It does not endorse commercial ventures, partnerships, or " +
        "products unless those have been explicitly ratified through the Amendment Procedure (§5). The " +
        "anchor's authority derives entirely from the discipline with which it is withheld."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("§10  Node Expansion Protocol", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para(
        "The Encultura network is designed to expand through node proliferation rather than centralized " +
        "scaling. A new node — an Anjuman operating in a different community, institution, or context — " +
        "may be established when the following conditions are met:"
      ),
      numItem([sp("Founding Capacity. ", { bold: true }), sp("The founding members of the new node must include at least two individuals who have completed a full Developmental Cycle within the parent Anjuman and received formal Attestation of navigational capacity.")]),
      numItem([sp("Protocol Ratification. ", { bold: true }), sp("The new node must formally ratify this Protocol Specification (or a version amended through the Amendment Procedure) before beginning operations.")]),
      numItem([sp("Contextual Adaptation. ", { bold: true }), sp("The new node may develop a Contextual Supplement — a document specifying how the protocol mechanics are adapted to their specific community context. The Supplement is signed under the new node's own DID and does not require signing under the parent anchor. However, any Supplement that contradicts the three Axioms is invalid.")]),
      numItem([sp("Network Health Commitment. ", { bold: true }), sp("The new node commits to publishing Network Health Metrics on the standard cadence and to participating in inter-node NHI review at the close of each Generational Cycle.")]),

      // ── PART V: RATIFICATION ───────────────────────────────────────────────
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [sp("Part V: Ratification and Effective Date", { size: 40, bold: true, color: COLOR_INK })] }),

      para(
        "This document is designated Version 0.1 — Genesis Document. It becomes operative as the protocol " +
        "grammar of the Encultura network upon signature by the founding Anjuman members and publication " +
        "under the did:web:encultura.org identity anchor."
      ),

      para(
        "Version 0.1 is intentionally incomplete. It establishes the Axioms, the worldspace taxonomy, the " +
        "governance architecture, and the identity anchor framework. It does not specify the content of " +
        "the first pedagogical curriculum, the initial Network Health Metric baselines, or the full " +
        "technical implementation of the DID signing infrastructure. These will be developed through " +
        "the Amendment Procedure within the first Developmental Cycle."
      ),

      blockquote(
        "The incompleteness of Version 0.1 is a feature, not a deficiency. A protocol that presents itself " +
        "as complete before it has been tested is either lying about its origins or foreclosing the " +
        "developmental work it claims to enable. This specification is a container, not a conclusion."
      ),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [sp("Signature Block", { size: 30, bold: true, color: COLOR_DEEP })] }),

      para("The undersigned founding Anjuman members attest that they have read and understood this specification, accept its Axioms as binding constraints on their participation, and commit to the governance mechanics described herein."),

      para(" "),

      // Signature table
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({ children: [
            ...[1,2,3].map(() => new TableCell({
              borders: cellBorder,
              width: { size: 3120, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              children: [
                new Paragraph({ spacing: { before: 60, after: 400 }, children: [sp("Name:", { bold: true, size: 20 })] }),
                new Paragraph({ spacing: { before: 0, after: 400 }, children: [sp("DID / Handle:", { bold: true, size: 20 })] }),
                new Paragraph({ spacing: { before: 0, after: 200 }, children: [sp("Date:", { bold: true, size: 20 })] }),
              ]
            }))
          ]})
        ]
      }),

      para(" "),
      rule(),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120 },
        children: [
          mono("did:web:encultura.org"),
          sp("  ·  Version 0.1  ·  Released under CC BY-SA 4.0", { size: 19, color: "888888" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [sp("\"The full effect comes when the protocol begins to govern the interaction of participants.\"", { size: 20, italics: true, color: "666666" })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/encultura_protocol_v0.1.docx', buffer);
  console.log('Done');
});
