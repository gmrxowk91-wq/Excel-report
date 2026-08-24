# Excel Data Analysis — Data Technician Bootcamp (Week 1)

A collection of Excel exercises completed during Week 1 of a Data Technician bootcamp, working with **retail sales**, **bike sales** and **regional product sales** datasets. The focus of the week was turning raw spreadsheets into structured, queryable data and then summarising it into insights that a non-technical audience can act on.

---

## Overview

| | |
|---|---|
| **Topic** | Spreadsheet analysis and reporting with Microsoft Excel |
| **Datasets** | Retail sales transactions, bike sales (multi-country), product sales by English county, student grades |
| **Deliverables** | Formatted Excel tables, PivotTables, PivotCharts, calculated categorisation columns |
| **Duration** | 1 week (4 days of practical tasks) |

---

## Skills demonstrated

### Formulas and functions

| Function | Purpose | Applied to |
|---|---|---|
| `SUM` | Total commission across all transactions | Retail sales dataset |
| `SUMIF` | Conditional totals for a single category or region | Retail / regional sales |
| `AVERAGE` | Mean commission per transaction | Retail sales dataset |
| `AVERAGEIF` | Mean value filtered by a criterion | Retail / regional sales |
| `DATE`, `MONTH`, `YEAR` | Extracting and rebuilding date parts to group transactions by period | Sales date columns |
| `UNIQUE` | Producing a distinct list of products, counties and customer segments | Lookup and validation lists |
| `VLOOKUP` | Joining descriptive fields to transaction rows via a key | Product / customer reference tables |
| `SWITCH` | Banding sales volume into categories | Regional product sales |

```excel
' Totals and averages
=SUM(P2:P9)
=AVERAGE(P2:P9)
=SUMIF($B$2:$B$13, "Laptops", $C$2:$C$13)
=AVERAGEIF($B$2:$B$13, "Laptops", $C$2:$C$13)

' Date handling — normalise a transaction date to the first of its month
=DATE(YEAR(A2), MONTH(A2), 1)

' Distinct value list
=UNIQUE(B2:B13)

' Lookup against a reference table
=VLOOKUP(A2, ProductTable, 3, FALSE)

' Conditional banding with SWITCH
=SWITCH(TRUE, C2 > 600, "High", C2 >= 300, "Medium", "Low")
```

### Data preparation

- Converting flat ranges into **structured Excel Tables** (`Ctrl + T`) with named references, so formulas and PivotTables stay accurate as data grows
- Cleaning **hidden trailing spaces** from numeric fields and confirming data types before aggregation — values that look numeric but are stored as text will silently break a PivotTable
- Working on **copies** of source datasets to preserve the original data

### Sorting and filtering

- Applying filters to isolate subsets of the retail dataset
- Sorting `Age` from largest to smallest to inspect the distribution of customers
- Using filters as a sanity check on totals produced by formulas

### PivotTables

- Summarising bike sales by **country, market, age group and gender** to identify where revenue concentrates
- Summarising regional sales with **County in rows and Product in columns**, using Sales Volume as the aggregated value
- Refining layout — masking blank cells as `0` via *PivotTable Analyse → Options → Layout & Format* so gaps read as genuine zeroes rather than missing data

### Charts and visualisation

- Building **PivotCharts** from existing PivotTables to compare performance across countries and product categories
- Selecting chart types that suit the question being asked, and keeping visuals simple enough to be read at a glance in a presentation

---

## Task breakdown

### Retail sales analysis
Converted columns A–H into a named Excel Table, filtered and sorted the `Age` column, then calculated total and average commission with `SUM` and `AVERAGE`.

### Student grades table
Converted the grades range into an Excel Table and derived per-student **average** and **highest score** columns across English, Mathematics and Science.

### Bike sales PivotTable lab
Built a PivotTable over a multi-country bike sales dataset and answered analytical questions from it:

- **Germany** has customers across all three markets — Accessories, Bikes and Clothing
- **All six countries** in the dataset (Australia, Canada, France, Germany, the United Kingdom, the United States) record sales in all three markets
- The **United States** is the most profitable market overall when broken down by country, age group and gender

### Regional sales categorisation
Summarised product sales volume across six English counties with a PivotTable, then added a calculated column using `SWITCH` to band each row:

| Sales volume | Category |
|---|---|
| Greater than 600 | High |
| 300 – 600 | Medium |
| Less than 300 | Low |

### Bike sales visualisation lab
Created PivotCharts from pre-built PivotTables to visualise sales performance by country and product category.

---

## Repository structure

```
.
├── README.md
├── data/           # source datasets (retail sales, bike sales, regional sales)
├── workbooks/      # completed Excel workbooks with tables, PivotTables and charts
└── screenshots/    # evidence of each completed task
```

> Adjust the folder names above to match the files actually committed.

---

## Wider context covered in Week 1

Alongside the Excel work, the week covered two areas that frame how the analysis is handled and communicated:

- **Data governance** — the Data Protection Act, UK GDPR and the Freedom of Information Act, and what each means in practice for storing, using and deleting personal data
- **Communicating findings** — preparing an analysis of customer churn at the 12-month renewal point for delivery to a board of directors: structuring the narrative as problem → evidence → business impact → recommended action, leading with the findings that matter to cost, revenue and risk, and choosing Excel for the analysis and charts with PowerPoint for delivery

---

## Tools

- Microsoft Excel — Tables, formulas, PivotTables, PivotCharts
- Microsoft PowerPoint — presenting findings to stakeholders

---

## Key takeaways

- Structure the data before analysing it — Excel Tables and correct data types remove most downstream errors
- Validate aggregations against a filtered view before trusting them
- A PivotTable answers questions faster than a formula when the question is "how does this break down by…"
- Charts exist to make one point clearly, not to display every column available
EOF
