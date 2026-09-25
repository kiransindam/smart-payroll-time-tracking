# 💼 Smart Payroll & Time Tracking Automation — Excel 365

> A professional Excel 365 payroll and workforce time-tracking solution that transforms daily employee hours into automated payroll calculations, deductions, net pay, and management-ready reports.

![Excel](https://img.shields.io/badge/Microsoft%20Excel-365-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-ETL-blue?style=for-the-badge)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-Excel-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Portfolio%20Project-success?style=for-the-badge)

---

## 📌 Project Overview

The **Smart Payroll & Time Tracking Automation** project is an Excel 365-based payroll workflow designed to help organizations manage employee information, record daily working hours, import time-clock data, and automatically calculate payroll.

The workbook combines:

* Employee master data
* Daily timesheets
* CSV/time-clock imports
* Power Query data transformation
* Automated payroll calculations
* Overtime-ready formulas
* Tax and deduction placeholders
* Payroll summaries
* Management dashboard

The goal is to create a single, structured Excel solution where both **manual time entry** and **time-clock CSV data** can flow into the same payroll calculation system.

---

# 🎯 Business Problem

Organizations often receive employee attendance and working-hour information from multiple sources.

For example:

```text
Manual Timesheet
       +
Time-Clock CSV
       ↓
Different Data Formats
       ↓
Manual Processing
       ↓
Payroll Calculation
       ↓
Payroll Report
```

Manual processing can introduce:

* Data-entry errors
* Incorrect employee mapping
* Calculation mistakes
* Duplicate records
* Difficult payroll reconciliation
* Repetitive reporting work

This project creates a structured workflow to reduce repetitive calculations and improve consistency.

---

# 💡 Solution

The solution follows this architecture:

```text
                         ┌──────────────────┐
                         │ Employee Master  │
                         │    Employees     │
                         └────────┬─────────┘
                                  │
                    ┌─────────────┴─────────────┐
                    │                           │
             Manual Entry                 Time-Clock CSV
                    │                           │
                    ▼                           ▼
             ┌──────────────┐          ┌────────────────┐
             │  Timesheet   │          │  Power Query   │
             └──────┬───────┘          └───────┬────────┘
                    │                          │
                    └────────────┬─────────────┘
                                 ▼
                       ┌──────────────────┐
                       │   Payroll Calc   │
                       └────────┬─────────┘
                                │
                                ▼
                       ┌──────────────────┐
                       │ Payroll Summary  │
                       └────────┬─────────┘
                                │
                                ▼
                       ┌──────────────────┐
                       │    Dashboard     │
                       └──────────────────┘
```

---

# ✨ Key Features

## 👥 Employee Management

The `Employees` sheet stores:

* Employee ID
* Employee name
* Department
* Job title
* Employment status
* Pay type
* Hourly rate
* Tax percentage
* Other deduction percentage
* Active/inactive status

---

## ⏱️ Time Tracking

The `Timesheet` sheet allows users to enter:

* Date
* Employee ID
* Regular hours
* Overtime hours
* Notes
* Data source

Employee names and hourly rates are automatically retrieved from the employee database.

---

## 📥 CSV Time-Clock Import

The project is designed to accept time-clock CSV files containing fields such as:

```text
Employee ID
Date
Clock In
Clock Out
```

Power Query can transform the incoming data into the standardized structure used by the payroll calculation system.

Example:

```text
CSV
 ↓
Power Query
 ↓
Clean / Transform
 ↓
Import_Raw
 ↓
Payroll Calculation
```

---

# 🧮 Payroll Calculation

The workbook calculates payroll using formula-driven logic.

### Regular Pay

```text
Regular Hours × Hourly Rate
```

### Overtime Pay

```text
Overtime Hours × Hourly Rate × Overtime Multiplier
```

### Gross Pay

```text
Regular Pay + Overtime Pay
```

### Tax

```text
Gross Pay × Tax %
```

### Net Pay

```text
Gross Pay − Tax − Other Deductions
```

The overtime multiplier is stored in the `Settings` sheet so the calculation structure can be changed without rebuilding the workbook.

---

# 📊 Dashboard

The workbook includes a management dashboard showing key payroll information such as:

* Active employees
* Total hours
* Gross payroll
* Net payroll
* Pay period
* Employee payroll snapshot
* Gross pay visualization

The dashboard is designed to give management a quick overview without requiring them to inspect calculation sheets.

---

# 📁 Workbook Structure

| Sheet             | Purpose                               |
| ----------------- | ------------------------------------- |
| `Dashboard`       | Management KPIs and payroll overview  |
| `README`          | Workbook instructions                 |
| `Employees`       | Employee master database              |
| `Settings`        | Pay period and payroll assumptions    |
| `Timesheet`       | Manual time entry                     |
| `Import_Raw`      | Standardized imported time-clock data |
| `Payroll_Calc`    | Payroll calculation engine            |
| `Payroll_Summary` | Final payroll report                  |

---

# 🛠️ Technologies Used

### Microsoft Excel 365

Used for:

* Data management
* Tables
* Formulas
* Data validation
* Reporting
* Dashboard

### Power Query

Used for:

* CSV ingestion
* Data transformation
* Column standardization
* Time-clock data preparation
* Refreshable ETL workflow

### Excel Functions

Important functions include:

```text
XLOOKUP
SUMIFS
IFERROR
IF
MOD
```

---

# 🏗️ How to Build the Project From Scratch

## Step 1 — Create the workbook

Open Microsoft Excel 365 and create a blank workbook.

Create these sheets:

```text
Dashboard
README
Employees
Settings
Timesheet
Import_Raw
Payroll_Calc
Payroll_Summary
```

---

# Step 2 — Create Employee Database

Create an Excel Table called:

```text
tblEmployees
```

Recommended columns:

```text
Employee ID
Employee Name
Department
Job Title
Employment Status
Pay Type
Hourly Rate
Tax %
Other Deduction %
Active
```

Example:

| Employee ID | Employee Name | Status    | Hourly Rate |
| ----------- | ------------- | --------- | ----------: |
| E001        | John Smith    | Full-Time |          20 |
| E002        | Sarah Lee     | Full-Time |          25 |
| E003        | David Brown   | Part-Time |          22 |

---

# Step 3 — Create Settings

Add:

```text
Pay Period Start
Pay Period End
Standard Hours / Day
Overtime Multiplier
Default Tax Rate
```

Example:

```text
Pay Period Start     09/01/2026
Pay Period End       09/15/2026
Standard Hours/Day   8
Overtime Multiplier  1.5
Default Tax Rate     10%
```

---

# Step 4 — Create Timesheet

Create an Excel Table called:

```text
tblTimesheet
```

Columns:

```text
Date
Employee ID
Employee Name
Regular Hours
Overtime Hours
Hourly Rate
Regular Pay
OT Pay
Gross Pay
Source
Notes
```

---

# Step 5 — Connect Employee Information

Use `XLOOKUP`.

### Employee Name

```excel
=XLOOKUP(
    [@[Employee ID]],
    tblEmployees[Employee ID],
    tblEmployees[Employee Name],
    ""
)
```

### Hourly Rate

```excel
=XLOOKUP(
    [@[Employee ID]],
    tblEmployees[Employee ID],
    tblEmployees[Hourly Rate],
    0
)
```

This prevents the user from repeatedly typing employee names and rates.

---

# Step 6 — Calculate Regular Pay

```excel
=[@[Regular Hours]]*[@[Hourly Rate]]
```

---

# Step 7 — Calculate Overtime Pay

```excel
=[@[Overtime Hours]]
*[@[Hourly Rate]]
*Settings!$B$5
```

The overtime multiplier is controlled from the Settings sheet.

---

# Step 8 — Calculate Gross Pay

```excel
=[@[Regular Pay]]+[@[OT Pay]]
```

---

# Step 9 — Build Payroll Calculation

Create the `Payroll_Calc` sheet.

Recommended columns:

```text
Employee ID
Employee Name
Regular Hours
OT Hours
Hourly Rate
Regular Pay
OT Pay
Gross Pay
Tax %
Tax
Other Deduction
Net Pay
```

Use `SUMIFS` to aggregate employee hours.

Example:

```excel
=SUMIFS(
    tblTimesheet[Regular Hours],
    tblTimesheet[Employee ID],[@[Employee ID]],
    tblTimesheet[Date],">="&Settings!$B$2,
    tblTimesheet[Date],"<="&Settings!$B$3
)
```

---

# Step 10 — Calculate Tax

```excel
=[@[Gross Pay]]*[@[Tax %]]
```

---

# Step 11 — Calculate Net Pay

```excel
=[@[Gross Pay]]
-[@Tax]
-[@[Other Deduction]]
```

---

# Step 12 — Build Payroll Summary

Create a clean report containing:

```text
Employee
Regular Hours
Overtime Hours
Gross Pay
Tax
Other Deductions
Net Pay
```

Add a total row for the entire pay period.

---

# Step 13 — Create CSV Import

Create a sample CSV:

```csv
Employee ID,Date,Clock In,Clock Out
E001,09/01/2026,08:00,17:00
E001,09/02/2026,08:00,18:00
E002,09/01/2026,09:00,17:30
E003,09/01/2026,08:30,17:00
```

---

# Step 14 — Configure Power Query

In Excel:

```text
Data
 ↓
Get Data
 ↓
From File
 ↓
From Text/CSV
```

Select the CSV.

Transform:

```text
Employee ID → Text
Date        → Date
Clock In    → Time
Clock Out   → Time
```

Calculate:

```text
Hours Worked
```

Then load the standardized data into:

```text
Import_Raw
```

A reusable Power Query M template is included in:

```text
power_query/TimeClock_Import_Template.pq
```

Detailed instructions are available in:

```text
docs/POWER_QUERY_SETUP.md
```

---

# 🔄 End-to-End Workflow

Once configured, the workflow becomes:

```text
Employee Added
       ↓
Employee Master
       ↓
┌───────────────────────────┐
│                           │
▼                           ▼
Manual Timesheet       Time-Clock CSV
│                           │
│                           ▼
│                     Power Query
│                           │
│                           ▼
│                      Import_Raw
│                           │
└──────────────┬────────────┘
               ▼
        Payroll Calculation
               ↓
          Gross Pay
               ↓
          Tax/Deductions
               ↓
            Net Pay
               ↓
        Payroll Summary
               ↓
           Dashboard
```

---

# 🧪 Testing the Project

Testing is essential before using the workbook for real payroll.

## Test 1 — Employee Lookup

Enter:

```text
Employee ID = E001
```

Verify that:

```text
Employee Name = John Smith
Hourly Rate = 20
```

---

## Test 2 — Regular Pay

Enter:

```text
Regular Hours = 8
Hourly Rate = 20
```

Expected:

```text
Regular Pay = 160
```

---

## Test 3 — Overtime

Enter:

```text
OT Hours = 2
Hourly Rate = 20
OT Multiplier = 1.5
```

Expected:

```text
OT Pay = 60
```

---

## Test 4 — Gross Pay

Expected:

```text
160 + 60 = 220
```

---

## Test 5 — Net Pay

If:

```text
Gross Pay = 220
Tax = 22
Other Deduction = 10
```

Expected:

```text
Net Pay = 188
```

---

## Test 6 — CSV Import

Import:

```text
sample_data/sample_timeclock.csv
```

Verify:

```text
CSV
 ↓
Power Query
 ↓
Import_Raw
 ↓
Payroll_Calc
 ↓
Payroll_Summary
```

---

# 🔐 Data Protection

The workbook separates:

### Input

```text
Employees
Timesheet
Settings
```

### Calculation

```text
Payroll_Calc
```

### Reporting

```text
Payroll_Summary
Dashboard
```

Formula/report sheets can be protected to reduce accidental changes.

> Worksheet protection is not the same as file encryption or enterprise-grade security.

---

# 📈 Future Enhancements

The architecture can be extended with:

* Automatic daily overtime rules
* Weekly overtime rules
* Holiday pay
* Weekend rules
* Break-rule calculations
* Multiple pay rates
* Salaried employees
* Department-level payroll reports
* Payroll audit logs
* Exception reports
* Automated payslip generation
* CSV folder-based refresh
* Power BI reporting
* Advanced access control

---

# ⚠️ Payroll Disclaimer

This project is a portfolio/automation implementation.

Tax rates, statutory deductions, overtime requirements, employment regulations, and payroll rules vary by country, state, and jurisdiction.

The tax and deduction calculations included here are **configurable placeholders** and should not be treated as legal or tax advice.

Before using the workbook for actual payroll processing, validate the calculations against the organization's payroll/accounting requirements and applicable laws.

---

# 👨‍💻 Skills Demonstrated

This project demonstrates practical experience with:

* Microsoft Excel 365
* Advanced Excel formulas
* XLOOKUP
* SUMIFS
* Excel Tables
* Data Validation
* Conditional Formatting
* Data Management
* Payroll Calculations
* Financial Analysis Concepts
* Power Query
* ETL/Data Transformation
* Dashboard Development
* Business Reporting
* Workflow Automation

---

# 📂 Repository Structure

```text
smart-payroll-time-tracking/
│
├── README.md
├── LICENSE
├── .gitignore
│
├── Excel/
│   └── Payroll_Time_Tracking_Workbook.xlsx
│
├── sample_data/
│   └── sample_timeclock.csv
│
├── power_query/
│   └── TimeClock_Import_Template.pq
│
├── docs/
│   ├── POWER_QUERY_SETUP.md
│   ├── WORKBOOK_GUIDE.md
│   └── TESTING_GUIDE.md
│
└── screenshots/
    ├── dashboard.png
    ├── employees.png
    ├── timesheet.png
    ├── payroll-summary.png
    └── power-query.png
```

---

# 🚀 Quick Start

1. Download the Excel workbook.
2. Open it using Microsoft Excel 365.
3. Review the `README` sheet.
4. Add employees in `Employees`.
5. Set the pay period in `Settings`.
6. Enter sample hours in `Timesheet`.
7. Review `Payroll_Calc`.
8. Review `Payroll_Summary`.
9. Open the sample CSV.
10. Configure Power Query using the supplied template.
11. Refresh the query.
12. Verify the Dashboard and Payroll Summary.

---

# ⭐ Project Goal

The goal of this project is to demonstrate how Excel can be used as a structured business automation tool rather than simply as a spreadsheet.

The design separates:

**Data → Transformation → Calculation → Reporting**

This makes the solution easier to understand, test, maintain, and extend.
