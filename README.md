# 🏥 Health Insurance Cost & Risk Analysis

<img width="739" height="415" alt="Image" src="https://github.com/user-attachments/assets/a9559625-4989-4367-a1be-03c2ab1fab22" />

> **Data Analysis Project** | 1,338 Policyholders | 12 Variables | 5 Business Questions | Multiple R² = 0.7504

---

## 📋 Table of Contents

- [Project Overview](#project-overview)
- [Dataset Description](#dataset-description)
- [Data Cleaning & Preparation](#data-cleaning--preparation)
- [Descriptive Statistics](#descriptive-statistics)
- [Business Questions & Results](#business-questions--results)
- [Key Findings](#key-findings)
- [Recommendations](#recommendations)
- [Workbook Structure](#workbook-structure)
- [Methodology](#methodology)

---

## Project Overview

Health insurance companies face significant challenges in accurately pricing premiums and managing financial risk. Charges vary widely across individuals based on lifestyle, health status, and demographic factors.

This project analyses **1,338 insured individuals** across **12 variables** to:

- Identify the key cost drivers of insurance charges
- Quantify their impact through **regression** and **pivot table** analysis
- Build a **composite risk-scoring framework**
- Produce clear, data-driven recommendations to support smarter underwriting, wellness programme design, and regional pricing strategy

---

## Dataset Description

| Column | Description | Type / Values |
|---|---|---|
| Age | Policyholder age | Integer (18–64) |
| Sex | Gender of policyholder | Male / Female |
| BMI | Body Mass Index | Float (15.96–53.13) |
| Children | Number of dependants | Integer (0–5) |
| Smoker | Smoking status | Yes / No |
| Region | Geographic region (US) | Northeast / Northwest / Southeast / Southwest |
| Charges ($) | Annual insurance charge | Float — Dependent variable (Y) |
| Blood Pressure | Blood pressure reading | Float (110.1–140.0) |
| Exercise Freq. | Exercise frequency | Never / Rarely / Weekly / Daily |
| Pre-existing Cond. | Pre-existing health condition | Yes / No |
| Occupation Risk | Job risk level | Low / Moderate / High |
| Annual Income ($) | Annual income | Float (30,109–199,966) |

### Derived Columns Added

Four additional analytical columns were created to enable deeper segmentation:

| Derived Column | Description |
|---|---|
| **BMI Category** | Underweight (< 18.5) / Normal (18.5–25) / Overweight (25–30) / Obese (≥ 30) |
| **Age Group** | Banded into five groups: 18–25 / 26–35 / 36–45 / 46–55 / 56–65 |
| **Risk Score** | Composite score (0–9) — Smoker (+3), Obese BMI (+2), Pre-existing Condition (+2), High Occupation Risk (+1), Never Exercises (+1) |
| **Risk Tier** | Low (0–1) / Moderate (2–3) / High (4–5) / Very High (6–9) |

---

## Data Cleaning & Preparation

| Issue | Finding | Action Taken |
|---|---|---|
| Missing Values | 3 missing values — age (1), bmi (1), children (1) | Median imputation applied. No rows dropped. |
| Duplicate Rows | Zero duplicate records found | No action required. |
| Text Standardisation | Inconsistent capitalisation in categorical fields | Title Case applied to all categorical columns. |
| **Final Clean Dataset** | **1,338 records × 16 columns (12 original + 4 derived)** | Formatted as Excel Table — `CleanDataTable` |

---

## Descriptive Statistics

| Statistic | Age | BMI | Children | Charges ($) | Risk Score | Annual Income ($) |
|---|---|---|---|---|---|---|
| Mean | 39.21 | 30.66 | 1.09 | 13,270.42 | 2.86 | 113,942 |
| Median | 39 | 30.40 | 1 | 9,382.03 | 3 | 113,294 |
| Std Dev | 13.98 | 6.10 | 1.21 | 12,110.01 | 1.87 | 47,967 |
| Min | 18 | 15.96 | 0 | 1,121.87 | 0 | 30,109 |
| Max | 64 | 53.13 | 5 | 63,770.43 | 9 | 199,966 |

> **Key Observation:** Average charges of $13,270 with a wide standard deviation of $12,110 indicate high variability — certain high-risk groups are pulling the mean significantly upward. The maximum charge of $63,770 versus minimum of $1,122 confirms extreme spread driven by identifiable risk factors.

---

## Business Questions & Results

Five business questions were investigated using two methods:
- **Simple Linear Regression** for BQ1, BQ3, and BQ5
- **Pivot Table Analysis** for BQ2 and BQ4

---

### BQ1 — What is the impact of smoking on insurance charges?

**Analysis Type: Simple Linear Regression**

| Statistic | Value |
|---|---|
| Regression Equation | **Charges = $8,434.27 + $23,615.96 × Smoker** |
| R-Squared (R²) | **0.6198 (62.0% variance explained)** |
| Intercept (a) | $8,434.27 (Non-smoker baseline predicted charge) |
| Slope (b) | +$23,615.96 (additional charge if smoker = 1) |
| P-value | 8.27e-283 |
| T-Statistic | 46.66 |
| Significance | **p < 0.001 — Highly Significant ✓** |

> **Interpretation:** Smoking alone explains 62% of all charge variation. Smokers are predicted at $32,050 vs non-smokers at $8,434 — a **3.8× multiplier**. With p = 8.27e-283, this result is overwhelmingly statistically significant.

---

### BQ2 — How does BMI category affect insurance charges?

**Analysis Type: Pivot Table**

| BMI Category | Avg Charges ($) | Count | % of Portfolio | vs Normal BMI |
|---|---|---|---|---|
| Underweight (< 18.5) | $8,657.62 | 21 | 1.6% | Base (Lowest) |
| Normal (18.5–25) | $10,435.44 | 226 | 16.9% | Base |
| Overweight (25–30) | $10,950.68 | 385 | 28.8% | +5% above Normal |
| Obese (≥ 30) | $15,580.16 | 706 | 52.8% | **+49% above Normal** |
| Grand Total | $13,270.42 | 1,338 | 100% | — |

> **Key Finding:** Obese policyholders pay $15,580 on average — **49% more** than the Normal BMI group. With 52.8% of the entire portfolio classified as Obese, this is the single largest volume risk segment.

---

### BQ3 — Does age significantly drive insurance charges?

**Analysis Type: Simple Linear Regression**

| Statistic | Value |
|---|---|
| Regression Equation | **Charges = $3,179.96 + $257.29 × Age** |
| R-Squared (R²) | **0.0890 (8.9% variance explained)** |
| Intercept (a) | $3,179.96 |
| Slope (b) | +$257.29 per additional year of age |
| P-value | 6.53e-29 |
| T-Statistic | 11.43 |
| Significance | **p < 0.001 — Highly Significant ✓** |

> **Interpretation:** Every additional year of age adds $257.29 to predicted charges. Age 18 → $7,811 | Age 35 → $12,185 | Age 50 → $16,044 | Age 64 → $19,646. The slope provides a direct actuarial basis for age-banded premium structures.

---

### BQ4 — Which geographic region records the highest average insurance charges?

**Analysis Type: Pivot Table**

| Region | Avg Charges ($) | Count | % of Portfolio | Rank |
|---|---|---|---|---|
| Southeast | $14,735.41 | 364 | 27.2% | 1st (Highest) |
| Northeast | $13,406.38 | 324 | 24.2% | 2nd |
| Northwest | $12,417.58 | 325 | 24.3% | 3rd |
| Southwest | $12,346.94 | 325 | 24.3% | 4th (Lowest) |
| Grand Total | $13,270.42 | 1,338 | 100% | — |

> **Key Finding:** The Southeast records the highest average charges at $14,735 — **$2,388 more** than the Southwest ($12,347).

---

### BQ5 — How effectively does the composite risk score predict insurance charges?

**Analysis Type: Simple Linear Regression + Multiple Linear Regression**

| Statistic | Value |
|---|---|
| Regression Equation | **Charges = $1,953.66 + $3,542.78 × Risk Score** |
| R-Squared (R²) | **0.3145 (31.5% variance explained)** |
| P-value | 1.08e-111 |
| T-Statistic | 24.76 |
| Significance | **p < 0.001 — Highly Significant ✓** |

**Multiple Regression — All 6 Variables Combined (R² = 0.7504)**

| Variable | Coefficient | Interpretation |
|---|---|---|
| Intercept (Base) | $1,953.66 | Base charge — lowest risk policyholder |
| Smoker (0/1) | +$22,619.00 | Largest single contributor to charges |
| Age (per year) | +$258.52 | Each year of age adds $258.52 |
| BMI (per unit) | +$268.28 | Each BMI unit increase adds $268.28 |
| Risk Score (per point) | +$400.17 | Each composite risk point adds $400.17 |
| **Multiple R²** | **0.7504** | **75.0% of all charge variation explained** |

> **Interpretation:** The full model explains **75% of all charge variation**, making it highly suitable for underwriting use.

---

## Key Findings

| # | Finding | Evidence | Business Impact |
|---|---|---|---|
| 1 | **Smoking is the #1 cost driver** | Smokers average $32,050 vs $8,434 for non-smokers — 3.8× higher. R² = 0.62, p < 0.001. | Only 20.5% of policyholders smoke yet they drive disproportionate claims costs. |
| 2 | **Obesity significantly raises charges** | Obese policyholders average $15,580 — 49% more than Normal BMI ($10,435). | 52.8% of portfolio is Obese — a portfolio-wide risk concern. |
| 3 | **Charges rise consistently with age** | Age 56–65 averages $18,796 — more than double the 18–25 group ($9,111). | Linear slope of $257.29/year provides direct basis for age-banded pricing. |
| 4 | **Southeast is the highest-cost region** | Southeast ($14,735) is $2,388 above Southwest ($12,347) — a 19% regional gap. | Regional premium calibration and provider negotiations are warranted. |
| 5 | **Risk score effectively predicts charges** | Simple R² = 0.3145. Multiple R² = 0.7504 — model explains 75% of variation. | Provides a transparent, explainable framework for underwriting. |

---

## Recommendations

| ID | Recommendation | Action Detail | Priority |
|---|---|---|---|
| **R1** | **Implement Smoker Premium Surcharge** | Apply surcharge of $21,254–$23,616 above base premium. Offer cessation programmes with 12-month verified abstinence reducing premiums by 30%. | 🔴 CRITICAL |
| **R2** | **Launch BMI Wellness Programmes** | Offer subsidised weight management and nutrition coaching for policyholders with BMI ≥ 30. Link sustained weight loss to measurable premium reductions. | 🟠 HIGH |
| **R3** | **Adopt Age-Banded Premium Structures** | Use regression slope of $257.29/year as actuarial basis. Every 5-year age band increases premium by ~$1,286. Invest in preventive care for the 46–65 cohort. | 🟠 HIGH |
| **R4** | **Calibrate Regional Premiums** | Adjust base premiums by region. Partner with Southeast healthcare providers for better-negotiated rates. | 🟡 MEDIUM |
| **R5** | **Deploy Risk Score in Underwriting** | Each risk point adds $3,543 to charges. Use tiers: Low (0–1) = base rate | Moderate (2–3) = +15% | High (4–5) = +40% | Very High (6–9) = +80%. | 🔴 CRITICAL |

---

## Workbook Structure

| Sheet | Contents | Key Features |
|---|---|---|
| 1. Data | 1,338 clean records × 16 columns | Formatted as Excel Table (`CleanDataTable`). Sorted by Region. Includes 4 derived columns. |
| 2. Business Questions | 5 BQs with findings and recommendations | BQ1/3/5 = Regression results. BQ2/4 = Pivot results. |
| 3. Analysis | EDA, Pivot Tables, Regression SUMMARY OUTPUT | Descriptive stats, BQ2 pivot + chart, BQ4 pivot + chart, BQ1/3/5 full regression output. |
| Pivot Tables | 7 pivot tables and 5 pivot charts | Source for the interactive dashboard. |
| EXCEL DASHBOARD | 5 pivot charts displayed together | Interactive dashboard — slicers for Smoker, Region, BMI, Risk Tier. |
| 5. Summary | Executive summary — 5 findings & 5 recommendations | High-level overview of the entire analysis. |
| 6. Interpretation | In-depth interpretation of all 5 BQ results | Regression interpretation, pivot interpretation, revised recommendations. |

---

## Methodology

### Regression Analysis (BQ1, BQ3, BQ5)

All three regression analyses used the **Excel Data Analysis ToolPak**, producing a full SUMMARY OUTPUT including:
- Regression Statistics (Multiple R, R Square, Adjusted R Square, Standard Error, Observations)
- ANOVA table (df, SS, MS, F, Significance F)
- Coefficients table (Coefficients, Standard Error, t Stat, P-value, Lower 95%, Upper 95%)

### Pivot Table Analysis (BQ2, BQ4)

Pivot tables were built using Excel's native PivotTable feature from the clean dataset. Each pivot table summarises **Average of Charges ($)** by category alongside counts. All pivot tables are connected to pivot charts on the dashboard sheet.

### Composite Risk Score

The five-factor risk model assigns points as follows:

| Risk Factor | Points |
|---|---|
| Smoker | +3 |
| Obese BMI (≥ 30) | +2 |
| Pre-existing Condition | +2 |
| High Occupation Risk | +1 |
| Never Exercises | +1 |
| **Maximum Score** | **9** |

Risk Tiers: **Low** (0–1) | **Moderate** (2–3) | **High** (4–5) | **Very High** (6–9)

### Statistical Significance

All regression results are **highly statistically significant (p < 0.001)**. The null hypothesis — that the independent variable has no effect on charges — is rejected in all three regression analyses. Results are reliable for underwriting and premium pricing decisions.

---

*Health Insurance Cost & Risk Analysis | Data Analysis Class Project*  
*Dataset: 1,338 policyholders | Analysis: Regression + Pivot Tables | Tools: Microsoft Excel*
