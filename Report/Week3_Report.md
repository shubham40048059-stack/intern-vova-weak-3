# Week 3 Assignment

## SQL & Excel for Data Analytics

**Name:**
**College:**
**Course:**
**Semester:**
**Subject:**
**Submission Date:**

---

## Introduction
A database is an organized collection of data that is stored, managed, and retrieved efficiently. SQL is a structured query language used to communicate with databases. Excel is a spreadsheet program used for tabular data processing, reporting, and analysis.

SQL and Excel are both important in data analytics because they help us clean data, transform raw records, analyze business patterns, and communicate insights clearly. SQL is used to fetch data from structured databases; Excel is used for quick calculations, summarization, charts, and business reporting.

---

## Task 1
**Objective:** Understand how to display records and use aliases.
**Concept:** SELECT, aliasing, retrieving selected columns.
**SQL Query:**
```sql
SELECT * FROM Customers;
SELECT CustomerName, City FROM Customers;
SELECT CustomerName AS Customer_Name, City AS Customer_City FROM Customers;
```
**Output:** All customer records; selected columns; renamed columns in output.
**Explanation:** The SELECT statement retrieves data from a database table. Aliases make the output easier to read.

[INSERT SCREENSHOT — TASK 1 SQL OUTPUT]

---

## Task 2
**Objective:** Use filtering, sorting, and aggregate functions to answer business questions.
**Queries:**
- `WHERE UnitPrice > 5000`
- `ORDER BY UnitPrice DESC`
- `COUNT(*)`, `SUM()`, `AVG()`, `MIN()`, `MAX()`
**Outputs:** Useful sales and transaction summaries.
**Explanation:** WHERE filters records, ORDER BY sorts them, and aggregate functions summarize data mathematically.

[INSERT SCREENSHOT — TASK 2 AGGREGATE FUNCTIONS]

---

## Task 3
**Objective:** Group sales data and filter grouped results.
**Queries:**
```sql
SELECT p.ProductName, SUM(s.Quantity * p.UnitPrice) AS Total_Sales
FROM Sales s
JOIN Products p ON s.ProductID = p.ProductID
GROUP BY p.ProductName;
```
**Outputs:** Sales summaries by product and city.
**Explanation:** GROUP BY combines records into categories, while HAVING filters the grouped results.

[INSERT SCREENSHOT — TASK 3 GROUP BY AND HAVING]

---

## Task 4
**Objective:** Understand relational joins.
**INNER JOIN:** Returns matching rows from both tables.
**LEFT JOIN:** Returns all rows from the left table and matching rows from the right table.
**RIGHT JOIN:** Returns all rows from the right table and matching rows from the left table.
**Outputs:** Combined customer, sales, and product records.
**Explanation:** Joins are used to combine related tables and answer business questions using multiple datasets.

[INSERT SCREENSHOT — TASK 4 INNER JOIN]
[INSERT SCREENSHOT — TASK 4 LEFT JOIN]
[INSERT SCREENSHOT — TASK 4 RIGHT JOIN]

---

## Task 5
**Objective:** Use subqueries to filter records based on calculated values.
**Subquery 1:** Employees earning more than the average salary.
**Subquery 2:** Products priced above the average product price.
**Output:** High-earning employees and above-average products.
**Explanation:** Subqueries are nested queries that are used inside another query for comparison.

[INSERT SCREENSHOT — TASK 5 SUBQUERIES]

---

## Task 6
**Objective:** Format and clean data in Excel.
**Formatting:** Headers bold, currency format, date format, borders, freeze panes.
**Sorting:** SalesAmount descending.
**Filtering:** City, Category, SalesAmount > 5000.
**Screenshot placeholder:**

[INSERT SCREENSHOT — TASK 6 EXCEL FORMATTING]

---

## Task 7
**Objective:** Use conditional formatting and formulas in Excel.
**Conditional Formatting:** SalesAmount > 5000 highlighted, < 2000 highlighted.
**IF():** `=IF(K2>5000,"High","Low")`
**COUNTIF():** `=COUNTIF(N:N,"High")`
**SUMIF():** `=SUMIF(N:N,"High",K:K)`
**Screenshot placeholder:**

[INSERT SCREENSHOT — TASK 7 CONDITIONAL FORMATTING]
[INSERT SCREENSHOT — TASK 7 EXCEL FUNCTIONS]

---

## Task 8
**Objective:** Retrieve product information using lookups.
**VLOOKUP:** `=VLOOKUP(F2,Products!A:E,2,FALSE)`
**XLOOKUP:** `=XLOOKUP(F2,Products!A:A,Products!B:B)`
**Comparison:** VLOOKUP requires column index; XLOOKUP is more flexible and easier to use.
**Screenshot placeholder:**

[INSERT SCREENSHOT — TASK 8 VLOOKUP]
[INSERT SCREENSHOT — TASK 8 XLOOKUP]

---

## Task 9
**Objective:** Build pivot tables and charts from sales data.
**Pivot Table:** Rows `ProductName`, Values `Sum of SalesAmount`.
**Second Pivot Table:** Rows `City`, Values `Sum of SalesAmount`.
**Analysis:** Product sales and city sales patterns can be visualized.
**Screenshot placeholder:**

[INSERT SCREENSHOT — TASK 9 PIVOT TABLE]
[INSERT SCREENSHOT — TASK 9 CHART]

---

## Conclusion
This project demonstrates how SQL and Excel can be used together for business data analytics. SQL is used for structured querying and aggregation, while Excel is used for formatting, formulas, lookups, and reporting. Together they provide a complete analytical workflow suitable for real-world decision-making.
