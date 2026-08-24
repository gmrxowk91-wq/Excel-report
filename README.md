# Excel Data Analysis — Data Technician Bootcamp (Week 1)

Spreadsheet analysis completed during Week 1 of a Data Technician bootcamp, working with **retail sales**, **bike sales** and **regional product sales** data. The week covered turning raw spreadsheets into structured, queryable tables, then summarising them into PivotTables and charts that answer a specific business question.

---

## Overview

| | |
|---|---|
| **Tool** | Microsoft Excel |
| **Datasets** | Retail sales transactions, bike sales (multi-country, 2017–2021), product sales by English county, student grades |
| **Deliverables** | Structured tables, PivotTables, PivotCharts, calculated categorisation columns |
| **Duration** | 1 week (4 days of practical tasks) |

---

## Skills demonstrated

### Formulas and functions

| Function | Purpose | Applied to |
|---|---|---|
| `SUM` | Total commission across all transactions | Retail sales dataset |
| `AVERAGE` | Mean commission; per-student average across three subjects | Retail sales; student grades |
| `MAX` | Highest score per student across subjects | Student grades |
| `AVERAGEIF` | Conditional mean — average sale value by customer gender | Retail sales dataset |
| `SWITCH` | Banding sales volume into High / Medium / Low | Regional product sales |

```excel
' Totals and averages
=SUM(P2:P9)
=AVERAGE(B2:D2)
=MAX(B2:D2)

' Conditional average — mean sale value for male customers only
=AVERAGEIF(D:D, "Male", I:I)

' Conditional banding with SWITCH
=SWITCH(TRUE, C2 > 600, "High", C2 >= 300, "Medium", "Low")
```

`SWITCH(TRUE, …)` is the idiom worth noting: `SWITCH` normally matches an exact value, but passing `TRUE` as the expression turns it into a sequential condition test — a flatter, more readable alternative to nested `IF` statements. The order of conditions carries the logic: `> 600` must be tested before `>= 300`, or every high value would be caught by the medium branch first.

### Data preparation

- Converting flat ranges into **structured Excel Tables** (`Ctrl + T`) with named references, so formulas and PivotTables stay accurate as data grows
- Cleaning **hidden trailing spaces** from numeric fields and confirming the data type before aggregating — values that look numeric but are stored as text break a PivotTable silently, with no error raised
- Working on **copies** of source datasets to preserve the originals
- Spotting inconsistent category values (the same country appearing under several near-identical spellings), which splits what should be one PivotTable column into three

### Sorting and filtering

- Filtering the retail dataset to isolate subsets, and sorting `Age` from largest to smallest to inspect the customer distribution
- Sorting the student grades table descending by calculated average, and separately by highest score
- Applying filters and sorts to surface the top performer per subject
- Using a filtered view as a sanity check against formula results

### Conditional formatting

- Highlighting the **highest and lowest** average scores in the grades table (green / red fill)
- Applying colour to PivotTable grand-total rows so the leading and trailing values in a country comparison read at a glance
- Flagging the higher of two conditional averages when comparing sale value by gender

### PivotTables

| PivotTable | Rows | Columns | Values |
|---|---|---|---|
| Profit by demographic and market | Customer gender → age group | Country | Sum of Profit |
| Order volume by demographic | Age group | Country | Sum of Order Quantity |
| Revenue by market and category | Country | Product category | Sum of Revenue |
| Revenue by demographic | Age group | — | Sum of Revenue |
| Annual performance | Year | — | Annual Profit, Annual Revenue |
| Regional sales | County | Product | Sum of Sales Volume |

Techniques applied: nested row fields for multi-level breakdowns, **grouping transaction dates up to Year** for the annual view, and masking blank cells as `0` via *PivotTable Analyse → Options → Layout & Format* so gaps read as genuine zeroes rather than missing data.

### Charts

Built as PivotCharts driven directly by the PivotTables above:

| Chart | Type | What it shows |
|---|---|---|
| Product Revenue by Country | Stacked column | Revenue split by category within each country, axis titled in US dollars |
| Sales Summary | Stacked column | Order quantity by age group, segmented by country |
| Revenue Comparison by Age Group | Pie with percentage labels | Adults 50%, Young Adults 36%, Youth 14%, Seniors 0% |
| Revenue vs Profits | Line, dual series | Annual revenue against annual profit, 2017–2021 |

---

## Task breakdown

### Retail sales analysis

Converted columns A–H into a named Excel Table, filtered and sorted the `Age` column, calculated total and average commission with `SUM` and `AVERAGE`, then extended the analysis independently with `AVERAGEIF` to compare average sale value by customer gender — **£455.43 for male customers against £456.55 for female**, a difference small enough to be worth reporting as "no meaningful difference" rather than as a finding.

### Student grades

Converted the grades range to an Excel Table, then:

- `=AVERAGE(B2:D2)` for each student's mean across English, Mathematics and Science
- `=MAX(B2:D2)` for each student's highest single subject score
- Sorted descending by average, and separately by highest score
- Filtered and sorted to identify the strongest student per subject
- Conditional formatting to mark the highest average (82) and lowest (50)

### Bike sales PivotTable lab

Built PivotTables over a multi-country bike sales dataset and answered analytical questions from them:

- **Germany** has customers across all three markets — Accessories, Bikes and Clothing
- **All six countries** (Australia, Canada, France, Germany, United Kingdom, United States) record sales in all three markets
- The **United States is the most profitable market** — £13.9m of £42.1m total profit, roughly a third of the whole, and the leading market in every age group and both genders
- Revenue is heavily concentrated in **Adults (35–64) at 50%** and **Young Adults (25–34) at 36%**, with the Seniors segment effectively absent at under 1%

### Regional sales categorisation

Summarised product sales volume across six English counties (Cornwall, Durham, Essex, Greater Manchester, Lancashire, Yorkshire) against three products (Laptops, Printers, Smartphones) in a PivotTable totalling 5,200 units, then added a calculated column using `SWITCH` to band each row:

| Sales volume | Category |
|---|---|
| Greater than 600 | High |
| 300 – 600 | Medium |
| Less than 300 | Low |

### Bike sales visualisation lab

Created PivotCharts from pre-built PivotTables — stacked columns for revenue by country and category, a pie chart for revenue share by age group, and a dual-series line chart tracking revenue against profit across 2017–2021.

---

## Wider context covered in Week 1

Alongside the Excel work, the week covered two areas that frame how the analysis is handled and communicated:

- **Data governance** — the Data Protection Act, UK GDPR and the Freedom of Information Act, and what each means in practice for storing, using and deleting personal data
- **Communicating findings** — preparing an analysis of customer churn at the 12-month renewal point for delivery to a board of directors: structuring the narrative as problem → evidence → business impact → recommended action, leading with what matters to cost, revenue and risk, and choosing Excel for the analysis with PowerPoint for delivery

---

## Repository contents

```
.
├── README.md
├── retail-sales_dataset.xlsx
├── Bike_Sales_Pivot_and_Visualisations.xlsx
├── regional_sales_switch.xlsx
└── screenshots/
```

---

## Key takeaways

- Structure the data before analysing it — Excel Tables and correct data types remove most downstream errors before they happen
- Hidden trailing spaces and inconsistent category spellings are the two faults that break a PivotTable without ever raising an error
- Condition order is the logic in `SWITCH(TRUE, …)` — reversing two thresholds produces a formula that runs and returns the wrong band
- A PivotTable answers "how does this break down by…" faster than any formula; reach for a formula when the answer is a single number
- Report a difference of £1.12 as no difference — a result that is technically calculable is not automatically a finding
- Charts exist to make one point clearly, not to display every column available
