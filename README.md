# Week 3 SQL & Excel Data Analytics Assignment

## Project Objective
This project demonstrates practical SQL and Excel analytics using a realistic retail sales dataset for a college assignment. It includes SQL scripts for database creation, inserts, SELECT, WHERE, ORDER BY, aggregate functions, GROUP BY, HAVING, joins, and subqueries along with Excel formatting, filtering, formulas, lookup functions, and pivot-table analysis.

## Dataset Description
The project uses a sales management dataset with the following tables:
- Customers
- Products
- Employees
- Sales

The dataset contains realistic Indian names, cities, products, prices, quantities, and sales values suitable for academic practice.

## Technologies Used
- MySQL-compatible SQL
- CSV dataset files
- Microsoft Excel compatible workbook
- Python for dataset generation and validation

## SQL Setup Instructions
1. Open MySQL Workbench or another MySQL client.
2. Run `01_create_database.sql` to create the database.
3. Run `02_create_tables.sql` to create tables.
4. Run `03_insert_data.sql` to load sample data.
5. Use `04_task1_select.sql` to `08_task5_subqueries.sql` for task-wise practice.

## Excel Instructions
1. Open the workbook in Excel or LibreOffice Calc.
2. Work on the `Sales_Data` sheet for raw sales records.
3. Use `Analysis` sheet for formulas and `Pivot_Table` sheet for pivot design.
4. Apply conditional formatting, filters, sorting, VLOOKUP/XLOOKUP, and charts.

## Folder Structure
- SQL/
- Data/
- Excel/
- Report/
- Screenshots/
- Documentation/

## Assignment Tasks
- Task 1: Database and SELECT
- Task 2: WHERE, ORDER BY, and aggregates
- Task 3: GROUP BY and HAVING
- Task 4: SQL joins
- Task 5: Subqueries
- Task 6: Excel formatting and filtering
- Task 7: Conditional formatting and formulas
- Task 8: VLOOKUP and XLOOKUP
- Task 9: Pivot tables and charts

## How to Run SQL
```sql
CREATE DATABASE Week3_Sales;
USE Week3_Sales;
SOURCE SQL/02_create_tables.sql;
SOURCE SQL/03_insert_data.sql;
```

## How to Complete Screenshots
- Take SQL screenshots from query output tabs.
- Capture Excel screenshots after formatting and lookups.
- Store them in the `Screenshots/` folder.

## How to Prepare PDF
1. Compile the report in Markdown or Word.
2. Insert screenshots in the marked placeholders.
3. Export to PDF.
4. Upload the PDF and provide the Google Drive link.

## Submission Checklist
- SQL files completed
- CSV datasets populated
- Excel workbook generated
- Report filled with required screenshots
- PDF and Google Drive link ready
