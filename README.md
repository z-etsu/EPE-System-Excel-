# 📝 Employee Performance Evaluation System (Excel/VBA-Based)

This system evaluates the performance of an employee based on **four criteria**, each with **five sub-criteria**.  
The evaluation scale ranges from:

> **Excellent (5)** to **Extremely Poor (1)**

There are **four evaluators** with different weightings:
- **Superior:** 40%
- **Colleague:** 30%
- **Subordinate:** 20%
- **Self-Assessment:** 10%

---

## 📁 Workbooks Overview

The system consists of three main Excel workbooks:
1. `Employee List`
2. `Performance Evaluation`
3. `Employees`

---

## 📘 Employee List Workbook

### Sheets:
- **Employee List**  
  Contains a list of all employees. Each entry is hyperlinked to the corresponding record in the `Employees` workbook.

- **Add Employee**  
  A form-style sheet to add new employees.

### Buttons:
- **Upload Picture**  
  - Inserts a picture into a specific cell range and autofits it.

- **Save**  
  Performs multiple actions:
  - Copies the current sheet to the `Employees` workbook.
  - Submits the values to the Employee List Sheet.
  - Hyperlinks the Employee ID on the saved sheet in the `Employees` workbook.
  - Renames the sheet based on Employee ID.
  - Protects the sheet and removes buttons.
  - Optionally clears current input fields.
  - Implements a **dependent dropdown list** (e.g., `Title` depends on selected `Department`).

---

## 📗 Employees Workbook

A centralized workbook containing all employee records and evaluation results.

### Button:
- **Evaluate**
  - Opens the `Performance Evaluation` workbook.
  - Transfers the selected employee's details to it.

---

## 📙 Performance Evaluation Workbook

Contains three sheets:
- `Employee`
- `Overview`
- `Dashboard`

---

### 🧾 Employee Sheet

Used to evaluate the employee.

- Evaluation scales are managed using `VLOOKUP` to determine:
  - Final Score
  - Remarks
- Uses `SUMPRODUCT` and `SUM` functions for weighted scoring.
- Displays a **summary dashboard** per criterion and evaluator.

#### Buttons:
1. **Save to PDF**
   - Exports the evaluation as a PDF, named after the employee.

2. **Submit**
   - Submits employee data and scores to the `Overview` sheet.
   - Hyperlinks the Employee ID to the `Employees` workbook.
   - Copies and formats the final evaluation score back to the `Employees` workbook.

---

### 📄 Overview Sheet

- Displays all evaluated employees and their scores.
- Each Employee ID is hyperlinked to the `Employees` workbook.

---

### 📊 Dashboard Sheet

- Highlights:
  - Top-performing employees
  - Leading departments
- Includes a **Refresh** button to update the dashboard.
