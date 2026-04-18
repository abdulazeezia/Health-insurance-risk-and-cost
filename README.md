const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// ── Colours ──────────────────────────────────────────────────────────────────
const NAVY   = "1F3864";
const BLUE   = "2E75B6";
const PALE   = "DEEAF1";
const WHITE  = "FFFFFF";
const GRAY   = "F2F2F2";
const DARK   = "1A1A2E";
const GREEN  = "375623";
const GREEN_BG = "E2EFDA";
const YELLOW_BG = "FFF2CC";
const RED    = "C00000";

// ── Border helpers ────────────────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "BDC7D6" };
const borders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const thickBorder = { style: BorderStyle.SINGLE, size: 4, color: NAVY };
const thickBorders = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Cell helper ───────────────────────────────────────────────────────────────
function cell(text, opts = {}) {
  const {
    bg = WHITE, bold = false, color = DARK, sz = 20,
    w = 2340, align = AlignmentType.LEFT,
    vAlign = VerticalAlign.CENTER, italic = false,
    brd = borders, colspan = 1
  } = opts;
  return new TableCell({
    borders: brd,
    width: { size: w, type: WidthType.DXA },
    columnSpan: colspan,
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: vAlign,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, color, size: sz, font: "Arial", italics: italic })]
    })]
  });
}

// ── Heading helper ────────────────────────────────────────────────────────────
function heading1(text) {
  return new Paragraph({
    spacing: { before: 360, after: 160 },
    children: [new TextRun({
      text, bold: true, color: NAVY, size: 32, font: "Arial",
      allCaps: true
    })]
  });
}

function heading2(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, color: BLUE, size: 26, font: "Arial" })]
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, color: DARK, size: 22, font: "Arial" })]
  });
}

function para(text, opts = {}) {
  const { bold = false, color = DARK, sz = 20, italic = false, spacing = { before: 60, after: 60 } } = opts;
  return new Paragraph({
    spacing,
    children: [new TextRun({ text, bold, color, size: sz, font: "Arial", italics: italic })]
  });
}

function bullet(text, opts = {}) {
  const { bold = false, color = DARK } = opts;
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, bold, color, size: 20, font: "Arial" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] });
}

function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: BLUE } },
    children: [new TextRun("")]
  });
}

// ── Title banner table ────────────────────────────────────────────────────────
function titleBanner(title, subtitle) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({ children: [new TableCell({
        borders: noBorders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 100, left: 300, right: 300 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: title, bold: true, color: WHITE, size: 40, font: "Arial", allCaps: true })]
        })]
      })] }),
      new TableRow({ children: [new TableCell({
        borders: noBorders,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 300, right: 300 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: subtitle, bold: false, color: WHITE, size: 20, font: "Arial", italics: true })]
        })]
      })] }),
    ]
  });
}

// ── Section banner ────────────────────────────────────────────────────────────
function sectionBanner(text, bg = BLUE) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: noBorders,
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: bg, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text, bold: true, color: WHITE, size: 24, font: "Arial" })]
      })]
    })]})],
  });
}

// ── KPI card row ──────────────────────────────────────────────────────────────
function kpiRow(items) {
  // items = [{label, value, sub}]
  const cellW = Math.floor(9360 / items.length);
  const cells = items.map(({ label, value, sub }) =>
    new TableCell({
      borders: thickBorders,
      width: { size: cellW, type: WidthType.DXA },
      shading: { fill: WHITE, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
          children: [new TextRun({ text: label, bold: true, color: NAVY, size: 18, font: "Arial", allCaps: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
          children: [new TextRun({ text: value, bold: true, color: BLUE, size: 36, font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: sub, color: "666666", size: 16, font: "Arial", italics: true })] }),
      ]
    })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: items.map(() => cellW),
    rows: [new TableRow({ children: cells })]
  });
}

// ── Regression results table ──────────────────────────────────────────────────
function regressionTable(bqTitle, equation, r2, r2pct, pvalue, tstat, intercept, slope, xVar, yVar, interpretation) {
  const rows = [
    // Header
    new TableRow({ children: [
      new TableCell({
        borders: noBorders, columnSpan: 4,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        children: [new Paragraph({ children: [new TextRun({ text: bqTitle, bold: true, color: WHITE, size: 22, font: "Arial" })] })]
      })
    ]}),
    // Variable row
    new TableRow({ children: [
      cell("Dependent Variable (Y):", { bg: YELLOW_BG, bold: true, w: 2340 }),
      cell(yVar, { bg: YELLOW_BG, w: 2340 }),
      cell("Independent Variable (X):", { bg: YELLOW_BG, bold: true, w: 2340 }),
      cell(xVar, { bg: YELLOW_BG, w: 2340 }),
    ]}),
    // Column headers
    new TableRow({ children: [
      cell("Statistic", { bg: NAVY, bold: true, color: WHITE, w: 2340 }),
      cell("Value", { bg: NAVY, bold: true, color: WHITE, w: 2340 }),
      cell("Statistic", { bg: NAVY, bold: true, color: WHITE, w: 2340 }),
      cell("Value", { bg: NAVY, bold: true, color: WHITE, w: 2340 }),
    ]}),
    new TableRow({ children: [
      cell("Regression Equation", { bold: true, w: 2340 }),
      cell(equation, { bg: PALE, bold: true, color: NAVY, w: 2340 }),
      cell("R-Squared (R²)", { bold: true, w: 2340 }),
      cell(`${r2} (${r2pct})`, { bg: PALE, bold: true, color: GREEN, w: 2340 }),
    ]}),
    new TableRow({ children: [
      cell("Intercept (a)", { bold: true, bg: GRAY, w: 2340 }),
      cell(intercept, { bg: GRAY, w: 2340 }),
      cell("Slope (b)", { bold: true, bg: GRAY, w: 2340 }),
      cell(slope, { bg: GRAY, w: 2340 }),
    ]}),
    new TableRow({ children: [
      cell("P-value", { bold: true, w: 2340 }),
      cell(pvalue, { bold: true, color: RED, w: 2340 }),
      cell("T-Statistic", { bold: true, w: 2340 }),
      cell(tstat, { w: 2340 }),
    ]}),
    new TableRow({ children: [
      cell("Significance", { bold: true, bg: GREEN_BG, w: 2340 }),
      cell("p < 0.001 — Highly Significant ✓", { bg: GREEN_BG, bold: true, color: GREEN, w: 2340 }),
      cell("Null Hypothesis", { bold: true, bg: GREEN_BG, w: 2340 }),
      cell("Rejected — Strong evidence of relationship", { bg: GREEN_BG, color: GREEN, w: 2340 }),
    ]}),
    // Interpretation
    new TableRow({ children: [
      new TableCell({
        borders, columnSpan: 4,
        width: { size: 9360, type: WidthType.DXA },
        shading: { fill: PALE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [
          new Paragraph({ children: [new TextRun({ text: "Interpretation:  ", bold: true, color: NAVY, size: 20, font: "Arial" }),
            new TextRun({ text: interpretation, color: DARK, size: 20, font: "Arial" })] })
        ]
      })
    ]}),
  ];
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340,2340,2340,2340], rows });
}

// ── Pivot table summary ───────────────────────────────────────────────────────
function pivotTable(bqTitle, headers, rows_data, finding) {
  const colW = Math.floor(9360 / headers.length);
  const tableRows = [
    // Title
    new TableRow({ children: [new TableCell({
      borders: noBorders, columnSpan: headers.length,
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: bqTitle, bold: true, color: WHITE, size: 22, font: "Arial" })] })]
    })] }),
    // Headers
    new TableRow({ children: headers.map(h => cell(h, { bg: BLUE, bold: true, color: WHITE, w: colW })) }),
    // Data rows
    ...rows_data.map((row, i) => new TableRow({
      children: row.map(v => cell(String(v), { bg: i % 2 === 0 ? "F5F8FB" : WHITE, w: colW }))
    })),
    // Finding
    new TableRow({ children: [new TableCell({
      borders, columnSpan: headers.length,
      width: { size: 9360, type: WidthType.DXA },
      shading: { fill: GREEN_BG, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({ children: [
        new TextRun({ text: "Key Finding:  ", bold: true, color: GREEN, size: 20, font: "Arial" }),
        new TextRun({ text: finding, color: DARK, size: 20, font: "Arial" })
      ]})]
    })] }),
  ];
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: headers.map(() => colW), rows: tableRows });
}

// ─────────────────────────────────────────────────────────────────────────────
// BUILD DOCUMENT
// ─────────────────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }, {
      reference: "numbers",
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // ── TITLE PAGE ──────────────────────────────────────────────────────────
      titleBanner(
        "Health Insurance Cost & Risk Analysis",
        "Data Analysis Project  |  1,338 Policyholders  |  12 Variables  |  5 Business Questions"
      ),
      spacer(),
      kpiRow([
        { label: "Total Records", value: "1,338", sub: "Clean policyholders analysed" },
        { label: "Variables", value: "12", sub: "Demographics, lifestyle & health" },
        { label: "Business Questions", value: "5", sub: "BQ1–BQ5 fully answered" },
        { label: "Multiple R²", value: "0.7504", sub: "75% variance explained" },
      ]),
      spacer(),

      // ── 1. PROJECT OVERVIEW ─────────────────────────────────────────────────
      divider(),
      sectionBanner("1.  PROJECT OVERVIEW"),
      spacer(),
      para("Health insurance companies face significant challenges in accurately pricing premiums and managing financial risk. Charges vary widely across individuals based on lifestyle, health status, and demographic factors. This project analyses 1,338 insured individuals across 12 variables to identify the key cost drivers, quantify their impact through regression and pivot analysis, build a composite risk-scoring framework, and produce clear data-driven recommendations to support smarter underwriting, wellness programme design, and regional pricing strategy."),
      spacer(),

      // Dataset
      heading2("1.1  Dataset Description"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2000, 3680, 3680],
        rows: [
          new TableRow({ children: [
            cell("Column", { bg: NAVY, bold: true, color: WHITE, w: 2000 }),
            cell("Description", { bg: NAVY, bold: true, color: WHITE, w: 3680 }),
            cell("Type / Values", { bg: NAVY, bold: true, color: WHITE, w: 3680 }),
          ]}),
          ...[ ["Age", "Policyholder age", "Integer (18–64)"],
               ["Sex", "Gender of policyholder", "Male / Female"],
               ["BMI", "Body Mass Index", "Float (15.96–53.13)"],
               ["Children", "Number of dependants", "Integer (0–5)"],
               ["Smoker", "Smoking status", "Yes / No"],
               ["Region", "Geographic region (US)", "Northeast / Northwest / Southeast / Southwest"],
               ["Charges ($)", "Annual insurance charge", "Float — Dependent variable (Y)"],
               ["Blood Pressure", "Blood pressure reading", "Float (110.1–140.0)"],
               ["Exercise Freq.", "Exercise frequency", "Never / Rarely / Weekly / Daily"],
               ["Pre-existing Cond.", "Pre-existing health condition", "Yes / No"],
               ["Occupation Risk", "Job risk level", "Low / Moderate / High"],
               ["Annual Income ($)", "Annual income", "Float (30,109–199,966)"],
          ].map((r, i) => new TableRow({ children: r.map((v, j) =>
            cell(v, { bg: i % 2 === 0 ? "F5F8FB" : WHITE, w: [2000, 3680, 3680][j] })
          )}))
        ]
      }),
      spacer(),

      // Derived columns
      heading2("1.2  Derived Columns Added"),
      para("Four additional analytical columns were created to enable deeper segmentation:"),
      bullet("BMI Category — Underweight (< 18.5) / Normal (18.5–25) / Overweight (25–30) / Obese (≥ 30)"),
      bullet("Age Group — Banded into five groups: 18–25 / 26–35 / 36–45 / 46–55 / 56–65"),
      bullet("Risk Score — Composite score (0–9) built from five risk factors: Smoker (+3), Obese BMI (+2), Pre-existing Condition (+2), High Occupation Risk (+1), Never Exercises (+1)"),
      bullet("Risk Tier — Low (0–1) / Moderate (2–3) / High (4–5) / Very High (6–9) based on risk score"),
      spacer(),

      // ── 2. DATA CLEANING ────────────────────────────────────────────────────
      divider(),
      sectionBanner("2.  DATA CLEANING & PREPARATION"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({ children: [
            cell("Issue", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
            cell("Finding", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
            cell("Action Taken", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
          ]}),
          new TableRow({ children: [
            cell("Missing Values", { w: 3120 }),
            cell("3 missing values — age (1), bmi (1), children (1)", { w: 3120 }),
            cell("Median imputation applied. No rows dropped.", { bg: GREEN_BG, w: 3120 }),
          ]}),
          new TableRow({ children: [
            cell("Duplicate Rows", { bg: "F5F8FB", w: 3120 }),
            cell("Zero duplicate records found", { bg: "F5F8FB", w: 3120 }),
            cell("No action required.", { bg: GREEN_BG, w: 3120 }),
          ]}),
          new TableRow({ children: [
            cell("Text Standardisation", { w: 3120 }),
            cell("Inconsistent capitalisation in categorical fields", { w: 3120 }),
            cell("Title Case applied to all categorical columns.", { bg: GREEN_BG, w: 3120 }),
          ]}),
          new TableRow({ children: [
            cell("Final Clean Dataset", { bg: PALE, bold: true, w: 3120 }),
            cell("1,338 records × 16 columns (12 original + 4 derived)", { bg: PALE, bold: true, w: 3120 }),
            cell("Formatted as Excel Table — CleanDataTable", { bg: GREEN_BG, bold: true, w: 3120 }),
          ]}),
        ]
      }),
      spacer(),

      // ── 3. DESCRIPTIVE STATISTICS ───────────────────────────────────────────
      divider(),
      sectionBanner("3.  DESCRIPTIVE STATISTICS (EXPLORATORY DATA ANALYSIS)"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1680, 1280, 1280, 1280, 1280, 1280, 1280],
        rows: [
          new TableRow({ children: [
            cell("Statistic", { bg: NAVY, bold: true, color: WHITE, w: 1680 }),
            cell("Age", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
            cell("BMI", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
            cell("Children", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
            cell("Charges ($)", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
            cell("Risk Score", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
            cell("Annual Income ($)", { bg: NAVY, bold: true, color: WHITE, w: 1280, align: AlignmentType.CENTER }),
          ]}),
          ...[
            ["Mean",    "39.21", "30.66", "1.09", "13,270.42", "2.86", "113,942"],
            ["Median",  "39",    "30.40", "1",    "9,382.03",  "3",    "113,294"],
            ["Std Dev", "13.98", "6.10",  "1.21", "12,110.01", "1.87", "47,967"],
            ["Min",     "18",    "15.96", "0",    "1,121.87",  "0",    "30,109"],
            ["Max",     "64",    "53.13", "5",    "63,770.43", "9",    "199,966"],
          ].map((row, i) => new TableRow({ children: row.map((v, j) =>
            cell(v, { bg: i % 2 === 0 ? "F5F8FB" : WHITE, w: [1680,1280,1280,1280,1280,1280,1280][j],
                      align: j === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })
          )}))
        ]
      }),
      spacer(),
      para("Key observation: Average charges of $13,270 with a wide standard deviation of $12,110 indicate high variability — suggesting that certain high-risk groups are pulling the mean significantly upward. The maximum charge of $63,770 versus minimum of $1,122 confirms extreme spread driven by identifiable risk factors."),
      spacer(),

      // ── 4. BUSINESS QUESTIONS & ANALYSIS ───────────────────────────────────
      divider(),
      sectionBanner("4.  BUSINESS QUESTIONS & ANALYSIS RESULTS"),
      spacer(),
      para("Five business questions were investigated using two methods: Simple Linear Regression for BQ1, BQ3 and BQ5; and Pivot Table analysis for BQ2 and BQ4. The Excel Data Analysis ToolPak SUMMARY OUTPUT format was used for all regression results."),
      spacer(),

      // ── BQ1 ─────────────────────────────────────────────────────────────────
      heading2("BQ1  —  What is the impact of smoking on insurance charges?"),
      para("Analysis Type: Simple Linear Regression", { italic: true, color: BLUE }),
      spacer(),
      regressionTable(
        "BQ1: Regression Analysis — Smoking vs Insurance Charges (Dependent Y = Charges, Independent X = Smoker)",
        "Charges = $8,434.27 + $23,615.96 × Smoker",
        "0.6198", "62.0% variance explained",
        "8.27e-283", "46.66",
        "$8,434.27 (Non-smoker baseline predicted charge)",
        "+$23,615.96 (additional charge if smoker = 1)",
        "Smoker Status (0 = Non-Smoker, 1 = Smoker)",
        "Insurance Charges ($)",
        "Smoking alone explains 62% of all charge variation — an exceptionally high R² for a single binary variable. For every policyholder who smokes, predicted charges increase by $23,616. Non-smokers have a predicted baseline of $8,434 while smokers are predicted at $32,050 — a 3.8× multiplier. With p = 8.27e-283, this result is overwhelmingly statistically significant."
      ),
      spacer(),

      // ── BQ2 ─────────────────────────────────────────────────────────────────
      heading2("BQ2  —  How does BMI category affect insurance charges?"),
      para("Analysis Type: Pivot Table Analysis", { italic: true, color: BLUE }),
      spacer(),
      pivotTable(
        "BQ2: Pivot Table — BMI Category vs Average Insurance Charges",
        ["BMI Category", "Avg Charges ($)", "Count", "% of Portfolio", "vs Normal BMI"],
        [
          ["Underweight (< 18.5)", "$8,657.62",  "21",  "1.6%",  "Base (Lowest)"],
          ["Normal (18.5–25)",     "$10,435.44", "226", "16.9%", "Base"],
          ["Overweight (25–30)",   "$10,950.68", "385", "28.8%", "+5% above Normal"],
          ["Obese (≥ 30)",         "$15,580.16", "706", "52.8%", "+49% above Normal"],
          ["Grand Total",          "$13,270.42", "1,338","100%", "—"],
        ],
        "Obese policyholders pay $15,580 on average — 49% MORE than the Normal BMI group ($10,435). With 52.8% of the entire portfolio classified as Obese, this represents the single largest volume risk segment."
      ),
      spacer(),

      // ── BQ3 ─────────────────────────────────────────────────────────────────
      heading2("BQ3  —  Does age significantly drive insurance charges?"),
      para("Analysis Type: Simple Linear Regression", { italic: true, color: BLUE }),
      spacer(),
      regressionTable(
        "BQ3: Regression Analysis — Age vs Insurance Charges (Dependent Y = Charges, Independent X = Age)",
        "Charges = $3,179.96 + $257.29 × Age",
        "0.0890", "8.9% variance explained",
        "6.53e-29", "11.43",
        "$3,179.96 (predicted charge at age zero — mathematical base)",
        "+$257.29 per additional year of age",
        "Age (Years — continuous variable)",
        "Insurance Charges ($)",
        "Every additional year of age adds $257.29 to predicted charges. Age 18 predicts $7,811; Age 35 predicts $12,185; Age 50 predicts $16,044; Age 64 predicts $19,646. Although R² of 8.9% is lower than smoking, age is a continuous universal variable affecting every policyholder and the slope provides a direct actuarial basis for age-banded premium structures."
      ),
      spacer(),

      // ── BQ4 ─────────────────────────────────────────────────────────────────
      heading2("BQ4  —  Which geographic region records the highest average insurance charges?"),
      para("Analysis Type: Pivot Table Analysis", { italic: true, color: BLUE }),
      spacer(),
      pivotTable(
        "BQ4: Pivot Table — Region vs Average Insurance Charges",
        ["Region", "Avg Charges ($)", "Count", "% of Portfolio", "Rank"],
        [
          ["Southeast", "$14,735.41", "364", "27.2%", "1st (Highest)"],
          ["Northeast", "$13,406.38", "324", "24.2%", "2nd"],
          ["Northwest", "$12,417.58", "325", "24.3%", "3rd"],
          ["Southwest", "$12,346.94", "325", "24.3%", "4th (Lowest)"],
          ["Grand Total","$13,270.42","1,338","100%", "—"],
        ],
        "The Southeast records the highest average charges at $14,735 — $2,388 more than the Southwest ($12,347). Regional differences may reflect variations in the mix of smokers, BMI levels, and age distributions rather than purely geographic pricing differences."
      ),
      spacer(),

      // ── BQ5 ─────────────────────────────────────────────────────────────────
      heading2("BQ5  —  How effectively does the composite risk score predict insurance charges?"),
      para("Analysis Type: Simple Linear Regression + Multiple Linear Regression", { italic: true, color: BLUE }),
      spacer(),
      regressionTable(
        "BQ5: Regression Analysis — Composite Risk Score vs Insurance Charges (Dependent Y = Charges, Independent X = Risk Score)",
        "Charges = $1,953.66 + $3,542.78 × Risk Score",
        "0.3145", "31.5% variance explained",
        "1.08e-111", "24.76",
        "$1,953.66 (predicted charge at risk score 0 — lowest risk)",
        "+$3,542.78 per additional risk point",
        "Composite Risk Score (0–9 points)",
        "Insurance Charges ($)",
        "Each 1-point increase in the composite risk score adds $3,543 to predicted charges. A Very High risk policyholder (score 9) is predicted at $33,839 versus $1,954 for a zero-risk individual. Multiple regression combining all 6 variables (age, BMI, smoker, children, pre-existing condition, risk score) achieves R² = 0.7504 — meaning the full model explains 75% of all charge variation, making it highly suitable for underwriting use."
      ),
      spacer(),

      // Multiple regression
      heading3("BQ5 Multiple Regression — All 6 Variables Combined"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [
          new TableRow({ children: [
            cell("Variable", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
            cell("Coefficient", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
            cell("Interpretation", { bg: NAVY, bold: true, color: WHITE, w: 3120 }),
          ]}),
          ...[
            ["Intercept (Base)",      "$1,953.66",  "Base charge — lowest risk policyholder"],
            ["Smoker (0/1)",          "+$22,619.00","Largest single contributor to charges"],
            ["Age (per year)",        "+$258.52",   "Each year of age adds $258.52"],
            ["BMI (per unit)",        "+$268.28",   "Each BMI unit increase adds $268.28"],
            ["Risk Score (per point)","+$400.17",   "Each composite risk point adds $400.17"],
            ["Multiple R²",           "0.7504",     "75.0% of all charge variation explained"],
          ].map((row, i) => new TableRow({ children: row.map((v, j) =>
            cell(v, { bg: i % 2 === 0 ? "F5F8FB" : WHITE, w: 3120 })
          )}))
        ]
      }),
      spacer(),

      // ── 5. KEY FINDINGS ─────────────────────────────────────────────────────
      new Paragraph({ children: [new PageBreak()] }),
      divider(),
      sectionBanner("5.  KEY FINDINGS SUMMARY"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [520, 2340, 3500, 3000],
        rows: [
          new TableRow({ children: [
            cell("#", { bg: NAVY, bold: true, color: WHITE, w: 520, align: AlignmentType.CENTER }),
            cell("Finding", { bg: NAVY, bold: true, color: WHITE, w: 2340 }),
            cell("Evidence", { bg: NAVY, bold: true, color: WHITE, w: 3500 }),
            cell("Business Impact", { bg: NAVY, bold: true, color: WHITE, w: 3000 }),
          ]}),
          ...[
            ["1","Smoking is the #1 cost driver",
             "Smokers average $32,050 vs $8,434 for non-smokers — 3.8× higher. R² = 0.62, p < 0.001.",
             "Only 20.5% of policyholders smoke yet they drive disproportionate claims costs."],
            ["2","Obesity significantly raises charges",
             "Obese policyholders average $15,580 — 49% more than Normal BMI ($10,435).",
             "52.8% of portfolio is Obese — this is a portfolio-wide risk concern."],
            ["3","Charges rise consistently with age",
             "Age 56–65 averages $18,796 — more than double the 18–25 group ($9,111). R² = 0.089, p < 0.001.",
             "Linear slope of $257.29/year provides direct basis for age-banded pricing."],
            ["4","Southeast is the highest-cost region",
             "Southeast ($14,735) is $2,388 above Southwest ($12,347) — a 19% regional gap.",
             "Regional premium calibration and provider negotiations are warranted."],
            ["5","Risk score effectively predicts charges",
             "Simple R² = 0.3145. Multiple R² = 0.7504 — model explains 75% of variation.",
             "Risk score provides a transparent, explainable framework for underwriting."],
          ].map((row, i) => new TableRow({ children: row.map((v, j) =>
            cell(v, { bg: i % 2 === 0 ? "F5F8FB" : WHITE, w: [520,2340,3500,3000][j],
                      align: j === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })
          )}))
        ]
      }),
      spacer(),

      // ── 6. RECOMMENDATIONS ──────────────────────────────────────────────────
      divider(),
      sectionBanner("6.  RECOMMENDATIONS"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [520, 2000, 4640, 2200],
        rows: [
          new TableRow({ children: [
            cell("ID", { bg: NAVY, bold: true, color: WHITE, w: 520, align: AlignmentType.CENTER }),
            cell("Recommendation", { bg: NAVY, bold: true, color: WHITE, w: 2000 }),
            cell("Action Detail", { bg: NAVY, bold: true, color: WHITE, w: 4640 }),
            cell("Priority", { bg: NAVY, bold: true, color: WHITE, w: 2200, align: AlignmentType.CENTER }),
          ]}),
          ...[
            ["R1","Implement Smoker Premium Surcharge",
             "Regression confirms smoking adds $23,616 to charges (R²=0.62, p<0.001). Apply surcharge of $21,254–$23,616 above base premium. Offer cessation programmes with 12-month verified abstinence reducing premiums by 30%.",
             "CRITICAL"],
            ["R2","Launch BMI Wellness Programmes",
             "Offer subsidised weight management and nutrition coaching for policyholders with BMI ≥ 30. Track outcomes annually and link sustained weight loss to measurable premium reductions.",
             "HIGH"],
            ["R3","Adopt Age-Banded Premium Structures",
             "Use regression slope of $257.29/year as actuarial basis. Every 5-year age band increases premium by ~$1,286. Invest in preventive care for the 46–65 cohort to flatten the cost curve.",
             "HIGH"],
            ["R4","Calibrate Regional Premiums",
             "Adjust base premiums by region. Partner with Southeast healthcare providers for better-negotiated rates. Target wellness initiatives specifically at the Southeast portfolio.",
             "MEDIUM"],
            ["R5","Deploy Risk Score in Underwriting",
             "Each risk point adds $3,543 to charges. Use tiers: Low (0–1) = base rate | Moderate (2–3) = +15% | High (4–5) = +40% | Very High (6–9) = +80%. Multiple R²=0.75 validates the model for production use.",
             "CRITICAL"],
          ].map((row, i) => {
            const priorityColors = { CRITICAL: RED, HIGH: "B8860B", MEDIUM: GREEN };
            const pColor = priorityColors[row[3]] || DARK;
            return new TableRow({ children: [
              cell(row[0], { bg: i%2===0?"F5F8FB":WHITE, w: 520, align: AlignmentType.CENTER, bold: true }),
              cell(row[1], { bg: i%2===0?"F5F8FB":WHITE, w: 2000, bold: true }),
              cell(row[2], { bg: i%2===0?"F5F8FB":WHITE, w: 4640 }),
              cell(row[3], { bg: i%2===0?"F5F8FB":WHITE, w: 2200, bold: true, color: pColor, align: AlignmentType.CENTER }),
            ]});
          })
        ]
      }),
      spacer(),

      // ── 7. WORKBOOK STRUCTURE ────────────────────────────────────────────────
      divider(),
      sectionBanner("7.  EXCEL WORKBOOK STRUCTURE"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2000, 3680, 3680],
        rows: [
          new TableRow({ children: [
            cell("Sheet", { bg: NAVY, bold: true, color: WHITE, w: 2000 }),
            cell("Contents", { bg: NAVY, bold: true, color: WHITE, w: 3680 }),
            cell("Key Features", { bg: NAVY, bold: true, color: WHITE, w: 3680 }),
          ]}),
          ...[
            ["1. Data","1,338 clean records × 16 columns","Formatted as Excel Table (CleanDataTable). Sorted by Region. Includes 4 derived columns."],
            ["2. Business Questions","5 BQs with findings and recommendations","BQ1/3/5 = Regression results. BQ2/4 = Pivot results."],
            ["3. Analysis","EDA, Pivot Tables, Regression SUMMARY OUTPUT","Descriptive stats, BQ2 pivot + chart, BQ4 pivot + chart, BQ1/3/5 full regression output."],
            ["Pivot Tables","7 pivot tables and 5 pivot charts","Source for the interactive dashboard."],
            ["EXCEL DASH BOARD","5 pivot charts displayed together","Interactive dashboard — slicers for Smoker, Region, BMI, Risk Tier."],
            ["5. Summary","Executive summary with 5 findings & 5 recommendations","High-level overview of the entire analysis."],
            ["6. Interpretation","In-depth interpretation of all 5 BQ results","Regression interpretation, pivot interpretation, revised recommendations."],
          ].map((row, i) => new TableRow({ children: row.map((v, j) =>
            cell(v, { bg: i%2===0?"F5F8FB":WHITE, w: [2000,3680,3680][j] })
          )}))
        ]
      }),
      spacer(),

      // ── 8. METHODOLOGY ──────────────────────────────────────────────────────
      divider(),
      sectionBanner("8.  METHODOLOGY NOTES"),
      spacer(),
      heading3("Regression Analysis (BQ1, BQ3, BQ5)"),
      para("All three regression analyses used the Excel Data Analysis ToolPak producing a full SUMMARY OUTPUT including: Regression Statistics (Multiple R, R Square, Adjusted R Square, Standard Error, Observations), ANOVA table (df, SS, MS, F, Significance F), and Coefficients table (Coefficients, Standard Error, t Stat, P-value, Lower 95%, Upper 95%)."),
      spacer(),
      heading3("Pivot Table Analysis (BQ2, BQ4)"),
      para("Pivot tables were built using Excel's native PivotTable feature from the clean data. Each pivot table summarises Average of Charges ($) by category, alongside counts. All pivot tables are connected to pivot charts on the dashboard sheet."),
      spacer(),
      heading3("Composite Risk Score"),
      para("The five-factor risk model assigns points as follows: Smoker = 3 points, Obese BMI (≥ 30) = 2 points, Pre-existing Condition = 2 points, High Occupation Risk = 1 point, Never Exercises = 1 point. Maximum possible score = 9. Risk tiers: Low (0–1), Moderate (2–3), High (4–5), Very High (6–9)."),
      spacer(),
      heading3("Statistical Significance"),
      para("All regression results in this project are highly statistically significant (p < 0.001). The null hypothesis — that the independent variable has no effect on charges — is rejected in all three regression analyses. Results are reliable for underwriting and premium pricing decisions."),
      spacer(),

      // ── FOOTER NOTE ─────────────────────────────────────────────────────────
      divider(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders: noBorders,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: NAVY, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [
              new TextRun({ text: "Health Insurance Cost & Risk Analysis  |  Data Analysis Class Project", bold: true, color: WHITE, size: 18, font: "Arial" })
            ]}),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [
              new TextRun({ text: "Dataset: 1,338 policyholders  |  Analysis: Regression + Pivot Tables  |  Tools: Microsoft Excel", color: PALE, size: 16, font: "Arial" })
            ]}),
          ]
        })] })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/README_Health_Insurance_Project.docx', buffer);
  console.log('Done');
});
