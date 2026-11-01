<div align="center">

# 📊 PR. 1 — FUNDAMENTAL BOOSTER

### ✦ Excel Data Analysis · Formula Automation · Dynamic Reporting ✦

**An Academic Practical Project in Microsoft Excel**

<br>

![Project](https://img.shields.io/badge/PROJECT-FUNDAMENTAL%20BOOSTER-1d4ed8?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/STATUS-PRESENTATION%20READY-16a34a?style=for-the-badge)
![Course](https://img.shields.io/badge/FOCUS-DATA%20ANALYSIS-7c3aed?style=for-the-badge)

<br>

**👩‍🎓 Chand Khimani** &nbsp; • &nbsp; **🎓 BCA — Final Year**  
**📚 Data Analysis Course** &nbsp; • &nbsp; **👨‍🏫 Faculty Guide: Prof. Girish Gondaliya**

<br>

### [▶️ WATCH THE PROJECT VIDEO PRESENTATION](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)

*Click above to open the full video demonstration in Google Drive.*

---

**📘 Documentation** · **🧮 Formula Implementation** · **📈 Practical Analysis** · **🎥 Video Walkthrough**

</div>

> [!IMPORTANT]
> **Evaluator quick access:** [Open the project demonstration video](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing). The video owner has configured **Anyone with the link → Viewer**. The Excel workbook should be submitted alongside this README.

## 🧭 Table of Contents

- [✨ Project Overview](#-project-overview)
- [🎯 Project Objectives](#-project-objectives)
- [🗂️ Workbook Architecture](#️-workbook-architecture)
- [🔍 Worksheet-by-Worksheet Breakdown](#-worksheet-by-worksheet-breakdown)
- [🧠 Excel Formula Toolkit](#-excel-formula-toolkit)
- [🧪 Practical Demonstrations](#-practical-demonstrations)
- [🎬 Video Presentation](#-video-presentation)
- [🚀 How to Review](#-how-to-review)
- [✅ Submission Checklist](#-submission-checklist)
- [🏆 Learning Outcomes](#-learning-outcomes)
- [👩‍🎓 Academic Details](#-academic-details)

---

## ✨ Project Overview

**PR. 1 — Fundamental Booster** is a structured Excel project demonstrating how raw data can be transformed into meaningful insights using **calculations, logical conditions, lookups, text processing, date functions, and dynamic references**.

The project combines realistic academic and business-style datasets with interactive formula demonstrations. Its focus is not just to display results, but to show **how those results are calculated and how they respond when inputs change**.

### 📌 Project Snapshot

| 📍 Metric | 📊 Project Scope |
|:--|:--|
| 🗂️ Total worksheets | **7** |
| 👨‍🎓 Student records | **20** |
| 🛍️ Sales transactions | **10** |
| 👥 Employee records | **10** |
| 📈 Salespersons analyzed | **10** |
| 🧩 Main capabilities | Conditional logic, multi-criteria analysis, lookups, dynamic references |
| 🎥 Presentation | [Google Drive video walkthrough](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing) |

### 🔄 Analysis Workflow

```text
   DATA ENTRY          EXCEL FORMULAS           INSIGHTS & OUTPUTS
       │                     │                         │
       ▼                     ▼                         ▼
 Student / Sales  ──►  IF · SUMIFS · XLOOKUP  ──►  Grades / Totals
 Employee / Text  ──►  INDEX · MATCH · TEXT   ──►  Search / Cleaning
 Monthly Trends   ──►  INDIRECT · OFFSET     ──►  Dynamic Analysis
```

---

## 🎯 Project Objectives

- 🧮 **Automate calculations** instead of calculating results manually.
- 🧠 **Apply logical rules** using `IF`, nested `IF`, `AND`, and `OR`.
- 📊 **Analyze data with multiple criteria** using `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
- 🔎 **Retrieve accurate records** using `VLOOKUP`, `XLOOKUP`, `XMATCH`, and `INDEX` + `MATCH`.
- 🧹 **Clean and standardize text** with Excel text functions.
- 📅 **Calculate age and experience** using date functions.
- 🔁 **Build responsive calculations** using relative/absolute references, `INDIRECT`, and `OFFSET`.

---

## 🗂️ Workbook Architecture

| # | 📑 Worksheet | 🎯 Purpose | ⚙️ Core Functions |
|:--:|:--|:--|:--|
| 01 | **Project Dashboard** | Introduction and workbook overview | Structured navigation and project scope |
| 02 | **Student Data** | Student marks, grades, eligibility and analysis | `SUM` · `AVERAGE` · `IF` · `COUNTIFS` · `AVERAGEIFS` · `VLOOKUP` · `FILTER` |
| 03 | **Sales Data** | Sales, discounts and product search | `IF` · `OR` · `SUMIFS` · `VLOOKUP` · `XMATCH` · rounding functions |
| 04 | **Employee Data** | Employee age, experience and search | `TODAY` · `DATEDIF` · `XLOOKUP` · `INDEX` · `MATCH` |
| 05 | **Salesperson Performance** | Monthly sales and performance comparison | `SUM` · `AVERAGE` · nested `IF` · `XLOOKUP` · `INDEX` + `MATCH` |
| 06 | **Text Functions** | Text extraction and standardization | `LEFT` · `RIGHT` · `FIND` · `LEN` · `UPPER` · `LOWER` |
| 07 | **Dynamic References** | Fixed references and rolling sales analysis | `$` references · `INDIRECT` · `OFFSET` |

---

## 🔍 Worksheet-by-Worksheet Breakdown

### 01 · 🏠 Project Dashboard

The starting point of the workbook. Introduces the project, the practical tasks and the Excel concepts covered.

### 02 · 🎓 Student Data — Academic Performance Analysis

**Dataset:** 20 students, with Math, Science and English marks.

**Implemented features:** Total and average marks, nested-IF grading, AND-based eligibility checks, conditional counts, conditional averages, student-ID lookup and dynamic filtering.

| 📌 Analysis | 📈 Result |
|:--|--:|
| Math score above 50 | **15 students** |
| Average score above 60 | **13 students** |
| Math **and** Science above 60 | **12 students** |
| Average of qualifying scores above 60 | **≈ 81.21** |
| Above 60 in all three subjects | **12 students** |

**Key idea:** Criteria-based formulas generate summaries automatically when source marks are updated.

### 03 · 🛒 Sales Data — Transaction & Discount Analysis

**Dataset:** 10 sales records with product codes, regions, salespersons, quantities and unit prices.

**Business rule:** Sales of **₹40,000 or more → 10% discount**; otherwise **5% discount**.

**Implemented features:** Gross sales, discount, final amount, eligibility, region-and-product `SUMIFS`, product-price `VLOOKUP`, product-position `XMATCH`, and numeric rounding.

| 🔎 Lookup | ⌨️ Input | ✅ Expected Result |
|:--|:--|:--|
| `VLOOKUP` | `P103` | **₹1,200** unit price |
| `XMATCH` | `P104` | **4** (first match within lookup range) |

**Actual formulas demonstrated:**

```excel
=VLOOKUP(N14,$C$4:$H$13,6,FALSE)
=XMATCH(N17,$C$4:$C$13,0)
```

> **Why this matters:** The product code is matched against column **C**, and the exact-match lookup returns the corresponding value or position. `$` locks the lookup range when copying a formula.

### 04 · 👥 Employee Data — Workforce Information

**Dataset:** 10 employees, with department, designation, salary, date of birth and date of joining.

**Implemented features:** Age, years of experience, total service days, salary retrieval by employee ID and dynamic name/department/designation lookup.

**Key idea:** `MATCH` identifies a record's position; `INDEX` returns the required field from that position.

### 05 · 📈 Salesperson Performance — Monthly Analysis

**Dataset:** 10 salespersons across January, February, March and April.

**Implemented features:** Total sales, monthly average, performance labels, salesperson lookup and month-specific sales retrieval.

| 👤 Salesperson | 🗓️ Month | 💰 Sales |
|:--|:--|--:|
| Neha | March | **₹62,000** |
| Neha | April | **₹65,000** |

**Key idea:** Changing the selected month updates the matching result without rewriting the lookup formula.

### 06 · 🔤 Text Functions — Data Preparation

**Implemented features:** Extracting first/last names, converting text to uppercase/lowercase and counting characters.

| ⚙️ Function | 📝 Application |
|:--|:--|
| `LEFT` + `FIND` | Extract first name |
| `RIGHT` + `LEN` + `FIND` | Extract last name |
| `UPPER` | Convert text to uppercase |
| `LOWER` | Convert text to lowercase |
| `LEN` | Count characters, including spaces |

### 07 · 🔁 Dynamic References — Responsive Calculations

**Dataset:** January–December monthly sales.

**Implemented features:** Relative references, absolute references for a fixed **5% bonus rate**, `INDIRECT` cell lookup and `OFFSET`-based recent-month calculations.

**Live scenario:** Change the recent-month input from **3** to **6** to update the total and average for the selected rolling period.

---

## 🧠 Excel Formula Toolkit

| 🧩 Category | ⚙️ Functions | 💡 What They Do |
|:--|:--|:--|
| Arithmetic | `SUM`, `AVERAGE` | Calculate totals and means |
| Logical | `IF`, nested `IF`, `AND`, `OR` | Evaluate conditions and make decisions |
| Multi-criteria | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Analyze matching records |
| Lookup | `VLOOKUP`, `XLOOKUP`, `XMATCH` | Find values and positions |
| Dynamic lookup | `INDEX` + `MATCH` | Return data from matched positions |
| Text | `LEFT`, `RIGHT`, `FIND`, `LEN`, `UPPER`, `LOWER` | Clean and transform strings |
| Date | `TODAY`, `DATEDIF` | Calculate age and tenure |
| Dynamic arrays | `FILTER` | Return rows that satisfy a condition |
| References | Relative, absolute, `INDIRECT`, `OFFSET` | Create reusable and responsive formulas |
| Math | `ROUND`, `CEILING`, `FLOOR` | Round numeric results |

### 💻 Selected Formula Patterns

```excel
=SUM(F4:F23)
=COUNTIFS(C4:C23,">50")
=SUMIFS(sum_range,region_range,selected_region,product_range,selected_product)
=XLOOKUP(employee_id,id_range,salary_range)
=INDEX(return_range,MATCH(lookup_value,lookup_range,0))
=FILTER(data_range,condition,"No matching records")
=INDIRECT(cell_containing_address)
=SUM(OFFSET(start_cell,rows_offset,0,number_of_rows,1))
```

*These are learning examples, not a claim that every sample cell reference is used verbatim in the submitted workbook. The workbook itself contains the actual formulas.*

---

## 🧪 Practical Demonstrations

| 🎯 Demo | 🖱️ Action | 👀 Expected Observation |
|:--|:--|:--|
| Student analysis | Select conditional-summary cells | Counts and averages reflect criteria |
| Product lookup | Search `P103` | Price **₹1,200** |
| Product position | Search `P104` | First matching position **4** |
| Employee lookup | Change employee ID | Corresponding employee details update |
| Salesperson analysis | Change Neha's month from March to April | **₹62,000 → ₹65,000** |
| Dynamic period | Change recent months from 3 to 6 | Total and average recalculate |

> [!TIP]
> While reviewing, select a calculated cell and inspect the **Formula Bar** to understand the underlying logic.

---

## 🎬 Video Presentation

<div align="center">

### ▶️ [CLICK HERE TO WATCH THE COMPLETE VIDEO](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)

**📹 Format:** MP4 &nbsp; | &nbsp; **🗣️ Explanation:** Hinglish &nbsp; | &nbsp; **📚 Content:** Workbook walkthrough

[![Watch Video](https://img.shields.io/badge/▶%20WATCH%20PROJECT%20DEMO-GOOGLE%20DRIVE-e11d48?style=for-the-badge&logo=googledrive&logoColor=white)](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)

</div>

**Presentation sequence:** Introduction → Student Data → Sales Data → Employee Data → Salesperson Performance → Text Functions → Dynamic References → Conclusion.

> [!NOTE]
> The video is hosted externally on Google Drive. If the link cannot be opened, check that the file's **General access** is **Anyone with the link** and its role is **Viewer**.

---

## 🚀 How to Review

1. 📂 Open the submitted **Excel workbook**.
2. 🏠 Start with **Project Dashboard** to understand the workbook structure.
3. 📑 Open each worksheet in the order shown above.
4. 🔍 Select formula cells and inspect the **Formula Bar**.
5. 🔁 Change a lookup input or selected month to observe the dynamic result.
6. ▶️ Open the **[video presentation](https://drive.google.com/file/d/1hf6CtBq7yKUORy_Hf-WQAHEzr0wNBCVn/view?usp=sharing)** for the guided walkthrough.

**💻 Compatibility:** `XLOOKUP`, `XMATCH`, and `FILTER` require Microsoft 365 or a compatible modern Excel version. Some older versions do not support these functions.

---

## ✅ Submission Checklist

- [x] Seven-sheet workbook structure documented
- [x] Core formulas and their purposes explained
- [x] Practical lookup examples included
- [x] Video presentation link added to README
- [x] Google Drive sharing configured as **Anyone with the link → Viewer**
- [ ] Confirm the video opens in an incognito/private window
- [ ] Recheck the final Excel workbook before submission
- [ ] Submit the **Excel workbook + README.md** together

---

## 🏆 Learning Outcomes

Through this project, I developed practical skills in **spreadsheet organization, formula-based automation, conditional analysis, exact-match searching, text cleaning, date calculations and dynamic reporting**.

The project also strengthened my ability to **explain technical work clearly**, verify formula outputs and present structured data analysis in an academic setting.

---

## 👩‍🎓 Academic Details

| 🏷️ Field | 📋 Information |
|:--|:--|
| **Student Name** | **Chand Khimani** |
| **Academic Program** | Bachelor of Computer Applications (BCA) — Final Year |
| **Course Focus** | Data Analysis |
| **Project Title** | PR. 1 — Fundamental Booster |
| **Faculty Guide** | **Prof. Girish Gondaliya** |
| **Tools** | Microsoft Excel · Google Drive · Markdown |
| **Deliverables** | Excel Workbook · README Documentation · Video Demonstration |

---

<div align="center">

### 🌟 Thank You for Reviewing My Project!

**Designed, documented and presented by Chand Khimani**  
*PR. 1 — Fundamental Booster | Excel Data Analysis*

[⬆️ Back to Top](#-pr-1--fundamental-booster)

</div>
