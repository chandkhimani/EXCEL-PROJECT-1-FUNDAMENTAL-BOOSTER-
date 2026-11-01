<div align="center">

# 📊 PR. 1 — FUNDAMENTAL BOOSTER

### An Excel-Based Data Analysis & Formula Implementation Project

**Microsoft Excel • Data Analysis • Formula Automation • Dynamic Reporting**

---

**Prepared by:** Chand Khimani  
**Program:** Bachelor of Computer Applications (BCA) — Final Year  
**Course Focus:** Data Analysis  
**Faculty Guide:** Prof. Girish Gondaliya

**Project Type:** Academic Practical Project  
**Deliverables:** Excel Workbook · Video Walkthrough · Project Documentation

[▶️ **WATCH PROJECT DEMONSTRATION**](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)

> **Video access:** Google Drive sharing is set to **Anyone with the link → Viewer**.

</div>

---

## 01 · Executive Overview

**PR. 1 — Fundamental Booster** is a structured, multi-worksheet Excel project designed to demonstrate how spreadsheet functions can turn raw records into useful, searchable, and dynamically updated information.

The workbook applies arithmetic calculations, conditional logic, criteria-based aggregation, lookup methods, text processing, date calculations, and dynamic references to practical datasets covering student performance, product sales, employee records, and salesperson performance.

### Project at a Glance

| Metric | Scope |
|:--|:--|
| Workbook architecture | 7 worksheets |
| Core domains | Academic, sales, employee, and performance analytics |
| Student dataset | 20 student records |
| Sales dataset | 10 transaction records |
| Employee dataset | 10 employee records |
| Salesperson dataset | 10 salespersons |
| Key techniques | Conditional logic, multi-criteria analysis, lookups, text processing, dynamic references |
| Presentation | Recorded practical walkthrough via Google Drive |

> **Note:** Counts and examples describe the prepared project structure. The submitted workbook is the authoritative source for live formula outputs.

---

## 02 · Objectives

- Apply fundamental and advanced Excel functions to structured datasets.
- Automate repetitive calculations and rule-based decisions.
- Retrieve records accurately using lookup functions.
- Analyze data using multiple conditions rather than manual filtering.
- Transform text and calculate date-based metrics.
- Build dynamic calculations that respond to changed inputs.
- Present the solution clearly through a practical video demonstration.

---

## 03 · Workbook Architecture

| # | Worksheet | Purpose | Main Techniques |
|---:|:--|:--|:--|
| 01 | **Project Dashboard** | Project introduction, scope, and navigation | Structured project overview |
| 02 | **Student Data** | Marks, grades, pass criteria, and student analysis | `SUM`, `AVERAGE`, `IF`, `AND`, `COUNTIFS`, `AVERAGEIFS`, `VLOOKUP`, `FILTER` |
| 03 | **Sales Data** | Transaction totals, discount rules, and sales analysis | Arithmetic, `IF`, `OR`, `SUMIFS`, `VLOOKUP`, `XMATCH`, `ROUND`, `CEILING`, `FLOOR` |
| 04 | **Employee Data** | Age, tenure, salary lookup, and employee search | `TODAY`, `DATEDIF`, `XLOOKUP`, `INDEX`, `MATCH` |
| 05 | **Salesperson Performance** | Monthly totals, performance categories, and month-specific lookup | `SUM`, `AVERAGE`, nested `IF`, `XLOOKUP`, `INDEX` + `MATCH` |
| 06 | **Text Functions** | Name extraction and text standardization | `LEFT`, `RIGHT`, `FIND`, `LEN`, `UPPER`, `LOWER` |
| 07 | **Dynamic References** | Relative/absolute references and rolling-period analysis | `$` references, `INDIRECT`, `OFFSET`, `SUM`, `AVERAGE` |

---

## 04 · Functional Modules

### 4.1 Student Performance Analytics

The **Student Data** worksheet processes subject marks to calculate totals, averages, grades, and eligibility conditions. Criteria-based formulas provide summary insights without manual counting. Student-ID lookup and dynamic filtering demonstrate fast record retrieval.

**Demonstrated analysis includes:** students scoring above a Math threshold; students exceeding an average threshold; students meeting combined Math and Science criteria; conditional average scores; and students exceeding thresholds in all three subjects.

**Skills demonstrated:** nested decision logic, multi-condition evaluation, aggregation, exact-match lookup, and dynamic arrays.

### 4.2 Sales & Discount Analysis

The **Sales Data** worksheet calculates sales amounts from quantities and unit prices, then applies a tiered discount rule. It also supports multi-criteria sales summaries and exact-match product retrieval.

**Business rule demonstrated:** transactions with total sales of **₹40,000 or more** receive a **10% discount**; otherwise, the discount is **5%**.

**Verified lookup examples from the project walkthrough:**

| Function | Search Input | Expected Demonstration |
|:--|:--|:--|
| `VLOOKUP` | `P103` | Unit price: **₹1,200** |
| `XMATCH` | `P104` | First matching position: **4** |

### 4.3 Employee Information Analysis

The **Employee Data** worksheet uses dates to calculate age and experience and employs exact-match lookups to retrieve salary and employee details. This demonstrates the difference between manually scanning records and automated retrieval.

### 4.4 Salesperson Performance Monitoring

Monthly sales are summarized into total and average values. Nested conditions classify performance, while a two-dimensional lookup retrieves sales for a selected salesperson and month.

**Demonstration scenario:** selecting **Neha + March** displays **₹62,000**; changing the month to **April** displays **₹65,000**.

### 4.5 Text Cleaning & Standardization

Text functions separate first and last names, convert text case, and count characters. These are foundational data-preparation techniques used before analysis and reporting.

### 4.6 Dynamic Reference Modeling

The **Dynamic References** worksheet demonstrates how relative references adjust when copied, how absolute references lock a parameter such as a **5% bonus rate**, and how `INDIRECT` and `OFFSET` make calculations responsive to user-selected cells and time windows.

**Live demonstration:** change the recent-month input from **3** to **6** and observe the updated sales total and average.

---

## 05 · Formula Reference Guide

| Technique | Formula Pattern | Why It Matters |
|:--|:--|:--|
| Total | `=SUM(range)` | Adds numerical values |
| Average | `=AVERAGE(range)` | Calculates the arithmetic mean |
| Conditional classification | `=IF(condition,result_if_true,result_if_false)` | Automates rule-based decisions |
| Multiple required conditions | `=AND(condition1,condition2)` | Returns TRUE only if all conditions pass |
| Alternative conditions | `=OR(condition1,condition2)` | Returns TRUE if at least one condition passes |
| Criteria-based count | `=COUNTIFS(criteria_range,criteria)` | Counts records matching conditions |
| Criteria-based sum | `=SUMIFS(sum_range,criteria_range,criteria)` | Aggregates matching records |
| Criteria-based average | `=AVERAGEIFS(average_range,criteria_range,criteria)` | Calculates a conditional average |
| Exact lookup | `=VLOOKUP(key,table,column,FALSE)` | Retrieves a value by matching a key |
| Flexible lookup | `=XLOOKUP(key,lookup_range,return_range)` | Retrieves data from a chosen return column |
| Position lookup | `=XMATCH(key,lookup_range,0)` | Returns an exact-match position |
| Position + value retrieval | `=INDEX(return_range,MATCH(key,lookup_range,0))` | Combines matching and retrieval |
| Conditional rows | `=FILTER(data_range,condition,"No Match")` | Returns all matching records |
| Fixed reference | `=$B$2` | Locks row and column during copying |
| Indirect reference | `=INDIRECT(cell_with_address)` | Resolves an address stored as text |
| Offset reference | `=OFFSET(start,rows,cols,height,width)` | Creates a shifted or variable range |

*Formula patterns are illustrative. Actual cell references and conditions are available in the Excel workbook.*

---

## 06 · Video Demonstration

<div align="center">

### ▶️ [OPEN THE FULL PROJECT WALKTHROUGH](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)

**Video format:** MP4 · **Suggested resolution:** 1080p · **Narration:** Hinglish

</div>

**Walkthrough order:** Project Dashboard → Student Data → Sales Data → Employee Data → Salesperson Performance → Text Functions → Dynamic References → Conclusion.

The video demonstrates the worksheet purpose, selected formulas, formula-bar logic, and changes in output when an input is updated.

**Access instructions for the evaluator:** Click the link above to open the video in Google Drive. If access is blocked, the owner must change Google Drive sharing to **Anyone with the link → Viewer**.

---

## 07 · How to Review the Project

1. Open the submitted Excel workbook in a version of Excel that supports the formulas used.
2. Start with **Project Dashboard** for the project overview.
3. Navigate through the worksheets in the order listed above.
4. Select calculated cells to inspect the formulas in the formula bar.
5. Test lookup inputs or dynamic month selection to verify that outputs update.
6. Open the **Video Demonstration** link for a guided explanation.

**Compatibility note:** `XLOOKUP`, `XMATCH`, and `FILTER` require a compatible Excel version, such as Microsoft 365 or supported recent Excel releases. Older Excel versions may not support these functions.

---

## 08 · Quality & Validation Checklist

- [x] Structured multi-sheet workbook design
- [x] Arithmetic and conditional formulas demonstrated
- [x] Multi-criteria calculations included
- [x] Lookup and search functions included
- [x] Text and date functions included
- [x] Relative, absolute, and dynamic references included
- [ ] Final workbook rechecked before submission
- [ ] Final MP4 video uploaded to Google Drive
- [ ] Google Drive video sharing set to **Anyone with the link → Viewer**
- [ ] Both video links in this README replaced with the actual sharing URL
- [ ] Video link tested in an incognito/private browser window

---

## 09 · Learning Outcomes

By completing this project, I strengthened my understanding of:

- Spreadsheet-based data organization and analysis.
- Conditional decision-making and multi-criteria calculations.
- Accurate record retrieval using lookup formulas.
- Text preparation and date-based analysis.
- Dynamic models that respond to input changes.
- Explaining technical work through a practical demonstration.

---

## 10 · Academic Information

| Field | Details |
|:--|:--|
| Student | **Chand Khimani** |
| Academic Program | **BCA — Final Year** |
| Learning Area | **Data Analysis** |
| Project | **PR. 1 — Fundamental Booster** |
| Faculty Guide | **Prof. Girish Gondaliya** |
| Submission Components | **Excel workbook, README documentation, video demonstration** |

---

<div align="center">

### Thank You

*Prepared with a focus on clarity, practical implementation, and reproducible Excel analysis.*

**Chand Khimani** · **PR. 1 — Fundamental Booster**

</div>
