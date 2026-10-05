from pathlib import Path
import csv
import os
import textwrap
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

base = Path(r"e:\weak 3 sql project\Week3_SQL_Excel_Assignment")
base.mkdir(parents=True, exist_ok=True)

# -------------------
# DATA DEFINITIONS
# -------------------
customers = [
    (1, 'Ananya Sharma', 'Delhi', 'ananya.sharma@gmail.com', 'Retail'),
    (2, 'Rohan Mehta', 'Mumbai', 'rohan.mehta@gmail.com', 'Wholesale'),
    (3, 'Priya Nair', 'Bengaluru', 'priya.nair@gmail.com', 'Retail'),
    (4, 'Aman Verma', 'Jaipur', 'aman.verma@gmail.com', 'Retail'),
    (5, 'Sneha Iyer', 'Chennai', 'sneha.iyer@gmail.com', 'Corporate'),
    (6, 'Karan Singh', 'Lucknow', 'karan.singh@gmail.com', 'Retail'),
    (7, 'Neha Kapoor', 'Delhi', 'neha.kapoor@gmail.com', 'Corporate'),
    (8, 'Vikram Rao', 'Hyderabad', 'vikram.rao@gmail.com', 'Retail'),
    (9, 'Meera Joshi', 'Pune', 'meera.joshi@gmail.com', 'Wholesale'),
    (10, 'Rahul Gupta', 'Kolkata', 'rahul.gupta@gmail.com', 'Retail'),
    (11, 'Sonia Malhotra', 'Ahmedabad', 'sonia.malhotra@gmail.com', 'Retail'),
    (12, 'Arjun Patel', 'Surat', 'arjun.patel@gmail.com', 'Corporate'),
    (13, 'Nisha Reddy', 'Bengaluru', 'nisha.reddy@gmail.com', 'Retail'),
    (14, 'Deepak Sharma', 'Noida', 'deepak.sharma@gmail.com', 'Retail'),
    (15, 'Ishita Sen', 'Kolkata', 'ishita.sen@gmail.com', 'Wholesale'),
    (16, 'Manav Khanna', 'Delhi', 'manav.khanna@gmail.com', 'Corporate'),
    (17, 'Tanvi Sood', 'Chandigarh', 'tanvi.sood@gmail.com', 'Retail'),
    (18, 'Harshit Jain', 'Indore', 'harshit.jain@gmail.com', 'Retail'),
]

products = [
    (1, 'Wireless Mouse', 'Electronics', 950.00, 120),
    (2, '4K Smart TV', 'Electronics', 28999.00, 18),
    (3, 'Office Chair', 'Furniture', 6999.00, 42),
    (4, 'Laptop Stand', 'Accessories', 1499.00, 85),
    (5, 'Bluetooth Speaker', 'Electronics', 3499.00, 70),
    (6, 'Premium Backpack', 'Fashion', 2499.00, 90),
    (7, 'Air Fryer', 'Home Appliances', 7999.00, 28),
    (8, 'Woolen Blanket', 'Home', 1899.00, 76),
    (9, 'Coffee Maker', 'Home Appliances', 5499.00, 22),
    (10, 'Gaming Keyboard', 'Electronics', 3999.00, 55),
    (11, 'Desk Lamp', 'Home', 1250.00, 110),
    (12, 'Water Purifier', 'Home Appliances', 18999.00, 15),
]

employees = [
    (1, 'Priya Nair', 'Sales', 45000.00),
    (2, 'Ramesh Kumar', 'Sales', 48000.00),
    (3, 'Aditi Shah', 'Operations', 52000.00),
    (4, 'Naveen Singh', 'Support', 43000.00),
    (5, 'Kavya Menon', 'Sales', 47000.00),
    (6, 'Sandeep Raj', 'Logistics', 41000.00),
    (7, 'Mitali Joshi', 'Sales', 49500.00),
]

sales = [
    (1, '2024-01-05', 1, 1, 2, 1, 'Cash', 1),
    (2, '2024-01-08', 2, 2, 1, 2, 'UPI', 2),
    (3, '2024-01-12', 3, 3, 2, 9, 'Card', 3),
    (4, '2024-01-17', 4, 4, 3, 3, 'Cash', 5),
    (5, '2024-01-20', 5, 5, 4, 5, 'UPI', 1),
    (6, '2024-01-28', 6, 6, 5, 10, 'Card', 2),
    (7, '2024-02-02', 7, 7, 6, 6, 'UPI', 7),
    (8, '2024-02-04', 8, 8, 7, 8, 'Cash', 4),
    (9, '2024-02-10', 9, 9, 8, 7, 'Card', 3),
    (10, '2024-02-16', 10, 10, 9, 2, 'UPI', 5),
    (11, '2024-02-18', 11, 11, 10, 11, 'Card', 2),
    (12, '2024-02-26', 12, 12, 11, 12, 'Cash', 6),
    (13, '2024-03-01', 13, 13, 12, 1, 'UPI', 1),
    (14, '2024-03-05', 14, 14, 13, 4, 'Card', 2),
    (15, '2024-03-12', 15, 15, 14, 5, 'UPI', 7),
    (16, '2024-03-17', 16, 16, 15, 3, 'Card', 3),
    (17, '2024-03-24', 17, 17, 16, 12, 'Cash', 2),
    (18, '2024-04-03', 18, 18, 17, 9, 'UPI', 4),
    (19, '2024-04-06', 1, 19, 18, 6, 'Card', 5),
    (20, '2024-04-10', 2, 20, 19, 7, 'Cash', 1),
    (21, '2024-04-14', 3, 21, 20, 10, 'UPI', 3),
    (22, '2024-04-18', 4, 22, 21, 8, 'Card', 7),
    (23, '2024-04-21', 5, 23, 22, 12, 'Cash', 2),
    (24, '2024-04-27', 6, 24, 23, 2, 'UPI', 5),
    (25, '2024-05-02', 7, 25, 24, 4, 'Card', 1),
    (26, '2024-05-06', 8, 26, 25, 5, 'Cash', 6),
    (27, '2024-05-11', 9, 27, 26, 11, 'UPI', 4),
    (28, '2024-05-19', 10, 28, 27, 1, 'Card', 7),
    (29, '2024-05-25', 11, 29, 28, 9, 'Cash', 2),
    (30, '2024-06-04', 12, 30, 29, 6, 'UPI', 3),
    (31, '2024-06-08', 13, 31, 30, 8, 'Card', 5),
    (32, '2024-06-12', 14, 32, 31, 3, 'Cash', 1),
    (33, '2024-06-17', 15, 33, 32, 10, 'UPI', 6),
    (34, '2024-06-22', 16, 34, 33, 4, 'Card', 7),
    (35, '2024-07-01', 17, 35, 34, 5, 'Cash', 2),
    (36, '2024-07-09', 18, 36, 35, 12, 'UPI', 4),
    (37, '2024-07-14', 1, 37, 36, 7, 'Card', 5),
    (38, '2024-07-18', 2, 38, 37, 2, 'Cash', 3),
    (39, '2024-07-26', 3, 39, 38, 11, 'UPI', 1),
    (40, '2024-08-02', 4, 40, 39, 9, 'Card', 6),
    (41, '2024-08-10', 5, 41, 40, 6, 'Cash', 4),
    (42, '2024-08-18', 6, 42, 41, 8, 'UPI', 7),
]

# Expand to full customer and sales IDs consistent with SQL dataset.
# The above list is intentionally mapped to deterministic value pairs by product/customer (customer IDs 1..18 + sales rows 1..42).
# Create a sales dataset with actual customer IDs and product IDs; quantities are derived below.
customer_map = {c[0]: c[1:] for c in customers}
product_map = {p[0]: p[1:] for p in products}

sales_rows = []
for sale_id, sale_date, customer_id, product_id, qty, emp_id, payment_method in [
    (1, '2024-01-05', 1, 1, 2, 1, 'Cash'),
    (2, '2024-01-08', 2, 2, 1, 2, 'UPI'),
    (3, '2024-01-12', 3, 9, 3, 3, 'Card'),
    (4, '2024-01-17', 4, 3, 2, 5, 'Cash'),
    (5, '2024-01-20', 5, 5, 4, 1, 'UPI'),
    (6, '2024-01-28', 6, 10, 2, 2, 'Card'),
    (7, '2024-02-02', 7, 6, 3, 7, 'UPI'),
    (8, '2024-02-04', 8, 8, 1, 4, 'Cash'),
    (9, '2024-02-10', 9, 7, 2, 3, 'Card'),
    (10, '2024-02-16', 10, 2, 1, 5, 'UPI'),
    (11, '2024-02-18', 11, 11, 2, 2, 'Card'),
    (12, '2024-02-26', 12, 12, 1, 6, 'Cash'),
    (13, '2024-03-01', 13, 1, 5, 1, 'UPI'),
    (14, '2024-03-05', 14, 4, 2, 2, 'Card'),
    (15, '2024-03-12', 15, 5, 3, 7, 'UPI'),
    (16, '2024-03-17', 16, 3, 1, 3, 'Card'),
    (17, '2024-03-24', 17, 12, 1, 2, 'Cash'),
    (18, '2024-04-03', 18, 9, 4, 4, 'UPI'),
    (19, '2024-04-06', 1, 6, 2, 5, 'Card'),
    (20, '2024-04-10', 2, 7, 1, 1, 'Cash'),
    (21, '2024-04-14', 3, 10, 3, 3, 'UPI'),
    (22, '2024-04-18', 4, 8, 2, 7, 'Card'),
    (23, '2024-04-21', 5, 12, 1, 2, 'Cash'),
    (24, '2024-04-27', 6, 2, 2, 5, 'UPI'),
    (25, '2024-05-02', 7, 4, 4, 1, 'Card'),
    (26, '2024-05-06', 8, 5, 3, 6, 'Cash'),
    (27, '2024-05-11', 9, 11, 2, 4, 'UPI'),
    (28, '2024-05-19', 10, 1, 6, 7, 'Card'),
    (29, '2024-05-25', 11, 9, 1, 2, 'Cash'),
    (30, '2024-06-04', 12, 6, 2, 3, 'UPI'),
    (31, '2024-06-08', 13, 8, 2, 5, 'Card'),
    (32, '2024-06-12', 14, 3, 3, 1, 'Cash'),
    (33, '2024-06-17', 15, 10, 4, 6, 'UPI'),
    (34, '2024-06-22', 16, 4, 1, 7, 'Card'),
    (35, '2024-07-01', 17, 5, 2, 2, 'Cash'),
    (36, '2024-07-09', 18, 12, 1, 4, 'UPI'),
    (37, '2024-07-14', 1, 7, 2, 5, 'Card'),
    (38, '2024-07-18', 2, 2, 1, 3, 'Cash'),
    (39, '2024-07-26', 3, 11, 3, 1, 'UPI'),
    (40, '2024-08-02', 4, 9, 2, 6, 'Card'),
    (41, '2024-08-10', 5, 6, 3, 4, 'Cash'),
    (42, '2024-08-18', 6, 8, 4, 7, 'UPI'),
]:
    product = product_map[product_id]
    unit_price = product[2]
    sales_amount = round(qty * unit_price, 2)
    sales_rows.append({
        'SaleID': sale_id,
        'SaleDate': sale_date,
        'CustomerID': customer_id,
        'CustomerName': customer_map[customer_id][0],
        'City': customer_map[customer_id][1],
        'ProductID': product_id,
        'ProductName': product[0],
        'Category': product[1],
        'Quantity': qty,
        'UnitPrice': unit_price,
        'SalesAmount': sales_amount,
        'EmployeeID': emp_id,
        'PaymentMethod': payment_method,
    })

# -------------------
# FILE WRITES
# -------------------
def ensure_path(path: Path):
    path.mkdir(parents=True, exist_ok=True)

# CSV files
csv_dir = base / 'Data'
ensure_path(csv_dir)

with open(csv_dir / 'customers.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['CustomerID', 'CustomerName', 'City', 'Email', 'CustomerType'])
    for row in customers:
        writer.writerow(row)

with open(csv_dir / 'products.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['ProductID', 'ProductName', 'Category', 'UnitPrice', 'Stock'])
    for row in products:
        writer.writerow(row)

with open(csv_dir / 'employees.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['EmployeeID', 'EmployeeName', 'Department', 'Salary'])
    for row in employees:
        writer.writerow(row)

with open(csv_dir / 'sales.csv', 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=['SaleID', 'SaleDate', 'CustomerID', 'CustomerName', 'City', 'ProductID', 'ProductName', 'Category', 'Quantity', 'UnitPrice', 'SalesAmount', 'EmployeeID', 'PaymentMethod'])
    writer.writeheader()
    for row in sales_rows:
        writer.writerow(row)

# SQL files
sql_dir = base / 'SQL'
ensure_path(sql_dir)

create_database_sql = textwrap.dedent('''
    CREATE DATABASE IF NOT EXISTS Week3_Sales;
    USE Week3_Sales;
''').strip() + '\n'

create_tables_sql = textwrap.dedent('''
    CREATE TABLE Customers (
        CustomerID INT PRIMARY KEY,
        CustomerName VARCHAR(100) NOT NULL,
        City VARCHAR(50) NOT NULL,
        Email VARCHAR(100) NOT NULL UNIQUE,
        CustomerType VARCHAR(30) NOT NULL
    );

    CREATE TABLE Products (
        ProductID INT PRIMARY KEY,
        ProductName VARCHAR(100) NOT NULL,
        Category VARCHAR(50) NOT NULL,
        UnitPrice DECIMAL(10,2) NOT NULL,
        Stock INT NOT NULL
    );

    CREATE TABLE Employees (
        EmployeeID INT PRIMARY KEY,
        EmployeeName VARCHAR(100) NOT NULL,
        Department VARCHAR(50) NOT NULL,
        Salary DECIMAL(10,2) NOT NULL
    );

    CREATE TABLE Sales (
        SaleID INT PRIMARY KEY,
        SaleDate DATE NOT NULL,
        CustomerID INT NOT NULL,
        ProductID INT NOT NULL,
        Quantity INT NOT NULL,
        SalesAmount DECIMAL(10,2) NOT NULL,
        EmployeeID INT NOT NULL,
        PaymentMethod VARCHAR(30) NOT NULL,
        FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID),
        FOREIGN KEY (ProductID) REFERENCES Products(ProductID),
        FOREIGN KEY (EmployeeID) REFERENCES Employees(EmployeeID)
    );
''').strip() + '\n'

insert_sql_lines = [
    'INSERT INTO Customers (CustomerID, CustomerName, City, Email, CustomerType) VALUES',
]
for idx, row in enumerate(customers):
    values = "({}, '{}', '{}', '{}', '{}')".format(row[0], row[1].replace("'", "''"), row[2], row[3], row[4])
    insert_sql_lines.append(values + (',' if idx < len(customers)-1 else ';'))

insert_sql_lines.extend(['', 'INSERT INTO Products (ProductID, ProductName, Category, UnitPrice, Stock) VALUES'])
for idx, row in enumerate(products):
    values = "({}, '{}', '{}', {:.2f}, {})".format(row[0], row[1].replace("'", "''"), row[2], row[3], row[4])
    insert_sql_lines.append(values + (',' if idx < len(products)-1 else ';'))

insert_sql_lines.extend(['', 'INSERT INTO Employees (EmployeeID, EmployeeName, Department, Salary) VALUES'])
for idx, row in enumerate(employees):
    values = "({}, '{}', '{}', {:.2f})".format(row[0], row[1].replace("'", "''"), row[2], row[3])
    insert_sql_lines.append(values + (',' if idx < len(employees)-1 else ';'))

insert_sql_lines.extend(['', 'INSERT INTO Sales (SaleID, SaleDate, CustomerID, ProductID, Quantity, SalesAmount, EmployeeID, PaymentMethod) VALUES'])
for idx, row in enumerate(sales_rows):
    values = "({}, '{}', {}, {}, {}, {:.2f}, {}, '{}')".format(row['SaleID'], row['SaleDate'], row['CustomerID'], row['ProductID'], row['Quantity'], row['SalesAmount'], row['EmployeeID'], row['PaymentMethod'])
    insert_sql_lines.append(values + (',' if idx < len(sales_rows)-1 else ';'))

insert_data_sql = '\n'.join(insert_sql_lines) + '\n'

select_task = textwrap.dedent('''
    -- Task 1: SELECT, aliases, and basic column selection
    SELECT * FROM Customers;

    SELECT CustomerName, City FROM Customers;

    SELECT
        CustomerName AS Customer_Name,
        City AS Customer_City,
        Email AS Email_Address
    FROM Customers;

    -- Purpose: Understand how to retrieve records and rename output columns.
''').strip() + '\n'

where_task = textwrap.dedent('''
    -- Task 2: WHERE, ORDER BY, and aggregate functions
    SELECT * FROM Products WHERE UnitPrice > 5000 ORDER BY UnitPrice DESC;

    SELECT * FROM Sales WHERE Quantity >= 3 ORDER BY Quantity DESC;

    SELECT CustomerName, City FROM Customers WHERE City = 'Delhi' ORDER BY CustomerName ASC;

    SELECT COUNT(*) AS Total_Transactions FROM Sales;
    SELECT SUM(Quantity * UnitPrice) AS Total_Revenue FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT AVG(Quantity * UnitPrice) AS Avg_Order_Value FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT MIN(Quantity * UnitPrice) AS Lowest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT MAX(Quantity * UnitPrice) AS Highest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
''').strip() + '\n'

group_task = textwrap.dedent('''
    -- Task 3: GROUP BY and HAVING
    SELECT p.ProductName, SUM(s.Quantity * p.UnitPrice) AS Total_Sales
    FROM Sales s
    JOIN Products p ON s.ProductID = p.ProductID
    GROUP BY p.ProductName
    ORDER BY Total_Sales DESC;

    SELECT c.City, COUNT(*) AS Transactions
    FROM Sales s
    JOIN Customers c ON s.CustomerID = c.CustomerID
    GROUP BY c.City
    HAVING COUNT(*) > 2
    ORDER BY Transactions DESC;

    SELECT p.Category, AVG(p.UnitPrice) AS Avg_Product_Price
    FROM Products p
    GROUP BY p.Category
    HAVING AVG(p.UnitPrice) > 5000;
''').strip() + '\n'

joins_task = textwrap.dedent('''
    -- Task 4: JOINs
    SELECT c.CustomerName, c.City, s.SaleID, s.SaleDate, s.SalesAmount
    FROM Customers c
    INNER JOIN Sales s ON c.CustomerID = s.CustomerID;

    SELECT c.CustomerName, c.City, s.SaleID, s.SaleDate, p.ProductName, p.Category
    FROM Customers c
    LEFT JOIN Sales s ON c.CustomerID = s.CustomerID
    LEFT JOIN Products p ON s.ProductID = p.ProductID;

    SELECT s.SaleID, s.SaleDate, c.CustomerName, c.City, p.ProductName
    FROM Sales s
    RIGHT JOIN Customers c ON s.CustomerID = c.CustomerID
    LEFT JOIN Products p ON s.ProductID = p.ProductID;

    SELECT c.CustomerName, c.City, p.ProductName, p.Category, s.Quantity, p.UnitPrice * s.Quantity AS SalesAmount
    FROM Customers c
    INNER JOIN Sales s ON c.CustomerID = s.CustomerID
    INNER JOIN Products p ON s.ProductID = p.ProductID
    ORDER BY c.CustomerName;
''').strip() + '\n'

subquery_task = textwrap.dedent('''
    -- Task 5: Subqueries
    SELECT EmployeeName, Salary
    FROM Employees
    WHERE Salary > (SELECT AVG(Salary) FROM Employees)
    ORDER BY Salary DESC;

    SELECT ProductName, UnitPrice
    FROM Products
    WHERE UnitPrice > (SELECT AVG(UnitPrice) FROM Products)
    ORDER BY UnitPrice DESC;

    SELECT CustomerName
    FROM Customers
    WHERE CustomerID IN (
        SELECT CustomerID
        FROM Sales
        WHERE Quantity >= 3
    );
''').strip() + '\n'

# SQL files
(base / 'SQL' / '01_create_database.sql').write_text(create_database_sql, encoding='utf-8')
(base / 'SQL' / '02_create_tables.sql').write_text(create_tables_sql, encoding='utf-8')
(base / 'SQL' / '03_insert_data.sql').write_text(insert_data_sql, encoding='utf-8')
(base / 'SQL' / '04_task1_select.sql').write_text(select_task, encoding='utf-8')
(base / 'SQL' / '05_task2_where_orderby_aggregate.sql').write_text(where_task, encoding='utf-8')
(base / 'SQL' / '06_task3_groupby_having.sql').write_text(group_task, encoding='utf-8')
(base / 'SQL' / '07_task4_joins.sql').write_text(joins_task, encoding='utf-8')
(base / 'SQL' / '08_task5_subqueries.sql').write_text(subquery_task, encoding='utf-8')

combined_sql = textwrap.dedent('''
    -- Week 3 SQL Complete Script
    CREATE DATABASE IF NOT EXISTS Week3_Sales;
    USE Week3_Sales;

    CREATE TABLE Customers (
        CustomerID INT PRIMARY KEY,
        CustomerName VARCHAR(100) NOT NULL,
        City VARCHAR(50) NOT NULL,
        Email VARCHAR(100) NOT NULL UNIQUE,
        CustomerType VARCHAR(30) NOT NULL
    );

    CREATE TABLE Products (
        ProductID INT PRIMARY KEY,
        ProductName VARCHAR(100) NOT NULL,
        Category VARCHAR(50) NOT NULL,
        UnitPrice DECIMAL(10,2) NOT NULL,
        Stock INT NOT NULL
    );

    CREATE TABLE Employees (
        EmployeeID INT PRIMARY KEY,
        EmployeeName VARCHAR(100) NOT NULL,
        Department VARCHAR(50) NOT NULL,
        Salary DECIMAL(10,2) NOT NULL
    );

    CREATE TABLE Sales (
        SaleID INT PRIMARY KEY,
        SaleDate DATE NOT NULL,
        CustomerID INT NOT NULL,
        ProductID INT NOT NULL,
        Quantity INT NOT NULL,
        SalesAmount DECIMAL(10,2) NOT NULL,
        EmployeeID INT NOT NULL,
        PaymentMethod VARCHAR(30) NOT NULL,
        FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID),
        FOREIGN KEY (ProductID) REFERENCES Products(ProductID),
        FOREIGN KEY (EmployeeID) REFERENCES Employees(EmployeeID)
    );

    INSERT INTO Customers (CustomerID, CustomerName, City, Email, CustomerType) VALUES
    (1, 'Ananya Sharma', 'Delhi', 'ananya.sharma@gmail.com', 'Retail'),
    (2, 'Rohan Mehta', 'Mumbai', 'rohan.mehta@gmail.com', 'Wholesale'),
    (3, 'Priya Nair', 'Bengaluru', 'priya.nair@gmail.com', 'Retail'),
    (4, 'Aman Verma', 'Jaipur', 'aman.verma@gmail.com', 'Retail'),
    (5, 'Sneha Iyer', 'Chennai', 'sneha.iyer@gmail.com', 'Corporate'),
    (6, 'Karan Singh', 'Lucknow', 'karan.singh@gmail.com', 'Retail'),
    (7, 'Neha Kapoor', 'Delhi', 'neha.kapoor@gmail.com', 'Corporate'),
    (8, 'Vikram Rao', 'Hyderabad', 'vikram.rao@gmail.com', 'Retail'),
    (9, 'Meera Joshi', 'Pune', 'meera.joshi@gmail.com', 'Wholesale'),
    (10, 'Rahul Gupta', 'Kolkata', 'rahul.gupta@gmail.com', 'Retail'),
    (11, 'Sonia Malhotra', 'Ahmedabad', 'sonia.malhotra@gmail.com', 'Retail'),
    (12, 'Arjun Patel', 'Surat', 'arjun.patel@gmail.com', 'Corporate'),
    (13, 'Nisha Reddy', 'Bengaluru', 'nisha.reddy@gmail.com', 'Retail'),
    (14, 'Deepak Sharma', 'Noida', 'deepak.sharma@gmail.com', 'Retail'),
    (15, 'Ishita Sen', 'Kolkata', 'ishita.sen@gmail.com', 'Wholesale'),
    (16, 'Manav Khanna', 'Delhi', 'manav.khanna@gmail.com', 'Corporate'),
    (17, 'Tanvi Sood', 'Chandigarh', 'tanvi.sood@gmail.com', 'Retail'),
    (18, 'Harshit Jain', 'Indore', 'harshit.jain@gmail.com', 'Retail');

    INSERT INTO Products (ProductID, ProductName, Category, UnitPrice, Stock) VALUES
    (1, 'Wireless Mouse', 'Electronics', 950.00, 120),
    (2, '4K Smart TV', 'Electronics', 28999.00, 18),
    (3, 'Office Chair', 'Furniture', 6999.00, 42),
    (4, 'Laptop Stand', 'Accessories', 1499.00, 85),
    (5, 'Bluetooth Speaker', 'Electronics', 3499.00, 70),
    (6, 'Premium Backpack', 'Fashion', 2499.00, 90),
    (7, 'Air Fryer', 'Home Appliances', 7999.00, 28),
    (8, 'Woolen Blanket', 'Home', 1899.00, 76),
    (9, 'Coffee Maker', 'Home Appliances', 5499.00, 22),
    (10, 'Gaming Keyboard', 'Electronics', 3999.00, 55),
    (11, 'Desk Lamp', 'Home', 1250.00, 110),
    (12, 'Water Purifier', 'Home Appliances', 18999.00, 15);

    INSERT INTO Employees (EmployeeID, EmployeeName, Department, Salary) VALUES
    (1, 'Priya Nair', 'Sales', 45000.00),
    (2, 'Ramesh Kumar', 'Sales', 48000.00),
    (3, 'Aditi Shah', 'Operations', 52000.00),
    (4, 'Naveen Singh', 'Support', 43000.00),
    (5, 'Kavya Menon', 'Sales', 47000.00),
    (6, 'Sandeep Raj', 'Logistics', 41000.00),
    (7, 'Mitali Joshi', 'Sales', 49500.00);

    INSERT INTO Sales (SaleID, SaleDate, CustomerID, ProductID, Quantity, SalesAmount, EmployeeID, PaymentMethod) VALUES
    (1, '2024-01-05', 1, 1, 2, 1900.00, 1, 'Cash'),
    (2, '2024-01-08', 2, 2, 1, 28999.00, 2, 'UPI'),
    (3, '2024-01-12', 3, 9, 3, 16497.00, 3, 'Card'),
    (4, '2024-01-17', 4, 3, 2, 13998.00, 5, 'Cash'),
    (5, '2024-01-20', 5, 5, 4, 13996.00, 1, 'UPI'),
    (6, '2024-01-28', 6, 10, 2, 7998.00, 2, 'Card'),
    (7, '2024-02-02', 7, 6, 3, 7497.00, 7, 'UPI'),
    (8, '2024-02-04', 8, 8, 1, 1899.00, 4, 'Cash'),
    (9, '2024-02-10', 9, 7, 2, 15998.00, 3, 'Card'),
    (10, '2024-02-16', 10, 2, 1, 28999.00, 5, 'UPI'),
    (11, '2024-02-18', 11, 11, 2, 2500.00, 2, 'Card'),
    (12, '2024-02-26', 12, 12, 1, 18999.00, 6, 'Cash'),
    (13, '2024-03-01', 13, 1, 5, 4750.00, 1, 'UPI'),
    (14, '2024-03-05', 14, 4, 2, 2998.00, 2, 'Card'),
    (15, '2024-03-12', 15, 5, 3, 10497.00, 7, 'UPI'),
    (16, '2024-03-17', 16, 3, 1, 6999.00, 3, 'Card'),
    (17, '2024-03-24', 17, 12, 1, 18999.00, 2, 'Cash'),
    (18, '2024-04-03', 18, 9, 4, 21996.00, 4, 'UPI'),
    (19, '2024-04-06', 1, 6, 2, 4998.00, 5, 'Card'),
    (20, '2024-04-10', 2, 7, 1, 7999.00, 1, 'Cash'),
    (21, '2024-04-14', 3, 10, 3, 11997.00, 3, 'UPI'),
    (22, '2024-04-18', 4, 8, 2, 3798.00, 7, 'Card'),
    (23, '2024-04-21', 5, 12, 1, 18999.00, 2, 'Cash'),
    (24, '2024-04-27', 6, 2, 2, 57998.00, 5, 'UPI'),
    (25, '2024-05-02', 7, 4, 4, 5996.00, 1, 'Card'),
    (26, '2024-05-06', 8, 5, 3, 10497.00, 6, 'Cash'),
    (27, '2024-05-11', 9, 11, 2, 2500.00, 4, 'UPI'),
    (28, '2024-05-19', 10, 1, 6, 5700.00, 7, 'Card'),
    (29, '2024-05-25', 11, 9, 1, 5499.00, 2, 'Cash'),
    (30, '2024-06-04', 12, 6, 2, 4998.00, 3, 'UPI'),
    (31, '2024-06-08', 13, 8, 2, 3798.00, 5, 'Card'),
    (32, '2024-06-12', 14, 3, 3, 20997.00, 1, 'Cash'),
    (33, '2024-06-17', 15, 10, 4, 15996.00, 6, 'UPI'),
    (34, '2024-06-22', 16, 4, 1, 1499.00, 7, 'Card'),
    (35, '2024-07-01', 17, 5, 2, 6998.00, 2, 'Cash'),
    (36, '2024-07-09', 18, 12, 1, 18999.00, 4, 'UPI'),
    (37, '2024-07-14', 1, 7, 2, 15998.00, 5, 'Card'),
    (38, '2024-07-18', 2, 2, 1, 28999.00, 3, 'Cash'),
    (39, '2024-07-26', 3, 11, 3, 3750.00, 1, 'UPI'),
    (40, '2024-08-02', 4, 9, 2, 10998.00, 6, 'Card'),
    (41, '2024-08-10', 5, 6, 3, 7497.00, 4, 'Cash'),
    (42, '2024-08-18', 6, 8, 4, 7596.00, 7, 'UPI');

    SELECT * FROM Customers;
    SELECT CustomerName, City FROM Customers;
    SELECT CustomerName AS Customer_Name, City AS Customer_City FROM Customers;

    SELECT * FROM Products WHERE UnitPrice > 5000 ORDER BY UnitPrice DESC;
    SELECT * FROM Sales WHERE Quantity >= 3 ORDER BY Quantity DESC;

    SELECT COUNT(*) AS Total_Transactions FROM Sales;
    SELECT SUM(Quantity * UnitPrice) AS Total_Revenue FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT AVG(Quantity * UnitPrice) AS Avg_Order_Value FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT MIN(Quantity * UnitPrice) AS Lowest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
    SELECT MAX(Quantity * UnitPrice) AS Highest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;

    SELECT p.ProductName, SUM(s.Quantity * p.UnitPrice) AS Total_Sales
    FROM Sales s JOIN Products p ON s.ProductID = p.ProductID
    GROUP BY p.ProductName
    ORDER BY Total_Sales DESC;

    SELECT c.City, COUNT(*) AS Transactions
    FROM Sales s JOIN Customers c ON s.CustomerID = c.CustomerID
    GROUP BY c.City
    HAVING COUNT(*) > 2;

    SELECT c.CustomerName, c.City, s.SaleID, s.SaleDate, p.ProductName, p.Category
    FROM Customers c
    INNER JOIN Sales s ON c.CustomerID = s.CustomerID
    INNER JOIN Products p ON s.ProductID = p.ProductID;

    SELECT c.CustomerName, c.City, s.SaleID, s.SaleDate, p.ProductName, p.Category
    FROM Customers c
    LEFT JOIN Sales s ON c.CustomerID = s.CustomerID
    LEFT JOIN Products p ON s.ProductID = p.ProductID;

    SELECT s.SaleID, s.SaleDate, c.CustomerName, c.City, p.ProductName
    FROM Sales s
    RIGHT JOIN Customers c ON s.CustomerID = c.CustomerID
    LEFT JOIN Products p ON s.ProductID = p.ProductID;

    SELECT EmployeeName, Salary
    FROM Employees
    WHERE Salary > (SELECT AVG(Salary) FROM Employees)
    ORDER BY Salary DESC;

    SELECT ProductName, UnitPrice
    FROM Products
    WHERE UnitPrice > (SELECT AVG(UnitPrice) FROM Products)
    ORDER BY UnitPrice DESC;
''').strip() + '\n'
(base / 'SQL' / 'week3_complete.sql').write_text(combined_sql, encoding='utf-8')

# README
readme = textwrap.dedent('''
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
''').strip() + '\n'
(base / 'README.md').write_text(readme, encoding='utf-8')

# Excel workbook generation
excel_dir = base / 'Excel'
ensure_path(excel_dir)
wb = Workbook()
ws = wb.active
ws.title = 'Sales_Data'
headers = ['SaleID', 'SaleDate', 'CustomerID', 'CustomerName', 'City', 'ProductID', 'ProductName', 'Category', 'Quantity', 'UnitPrice', 'SalesAmount', 'EmployeeID', 'PaymentMethod']
ws.append(headers)
for row in sales_rows:
    ws.append([
        row['SaleID'],
        row['SaleDate'],
        row['CustomerID'],
        row['CustomerName'],
        row['City'],
        row['ProductID'],
        row['ProductName'],
        row['Category'],
        row['Quantity'],
        row['UnitPrice'],
        row['SalesAmount'],
        row['EmployeeID'],
        row['PaymentMethod'],
    ])

# Format Sales_Data
header_fill = PatternFill('solid', fgColor='1F4E78')
header_font = Font(color='FFFFFF', bold=True)
thin = Side(style='thin', color='D9D9D9')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=13):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(vertical='center')
for col in ['B','J','K']:
    for cell in ws[col]:
        if cell.row > 1:
            cell.number_format = 'dd-mmm-yyyy' if col == 'B' else '₹#,##0.00'
for cols in ['I','J','K']:
    for cell in ws[cols]:
        if cell.row > 1:
            cell.number_format = '0'
            if cols in ['J','K']:
                cell.number_format = '₹#,##0.00'

# Customer sheet
ws_customers = wb.create_sheet('Customers')
ws_customers.append(['CustomerID', 'CustomerName', 'City', 'Email', 'CustomerType'])
for row in customers:
    ws_customers.append(list(row))

# Product sheet
ws_products = wb.create_sheet('Products')
ws_products.append(['ProductID', 'ProductName', 'Category', 'UnitPrice', 'Stock'])
for row in products:
    ws_products.append(list(row))

# Employee sheet
ws_employees = wb.create_sheet('Employees')
ws_employees.append(['EmployeeID', 'EmployeeName', 'Department', 'Salary'])
for row in employees:
    ws_employees.append(list(row))

# Analysis sheet
ws_analysis = wb.create_sheet('Analysis')
ws_analysis.append(['Metric', 'Value / Formula'])
ws_analysis.append(['Total Sales', '=SUM(Sales_Data!$K$2:$K$43)'])
ws_analysis.append(['Average Sales', '=AVERAGE(Sales_Data!$K$2:$K$43)'])
ws_analysis.append(['Highest Sale', '=MAX(Sales_Data!$K$2:$K$43)'])
ws_analysis.append(['Lowest Sale', '=MIN(Sales_Data!$K$2:$K$43)'])
ws_analysis.append(['High Sales Count', '=COUNTIF(Sales_Data!$K$2:$K$43, ">5000")'])
ws_analysis.append(['High Sales Total', '=SUMIF(Sales_Data!$K$2:$K$43, ">5000")'])
ws_analysis.append(['Sales Category Example', '=IF(Sales_Data!K2>5000,"High","Low")'])
ws_analysis.append(['Count High', '=COUNTIF(Sales_Data!$K$2:$K$43, ">5000")'])
ws_analysis.append(['VLOOKUP Product', '=VLOOKUP(F2,Products!A:E,2,FALSE)'])
ws_analysis.append(['XLOOKUP Product', '=XLOOKUP(F2,Products!A:A,Products!B:B)'])

# Pivot sheet
ws_pivot = wb.create_sheet('Pivot_Table')
ws_pivot['A1'] = 'Pivot Table Design'
ws_pivot['A2'] = 'Rows: ProductName'
ws_pivot['A3'] = 'Values: Sum of SalesAmount'
ws_pivot['A5'] = 'Rows: City'
ws_pivot['A6'] = 'Values: Sum of SalesAmount'
ws_pivot['A8'] = 'Chart Recommendation: Sales by Product - Column Chart'
ws_pivot['A9'] = 'Chart Recommendation: Sales by City - Bar Chart'

# Basic styles for workbook
for ws_temp in [ws_customers, ws_products, ws_employees, ws_analysis, ws_pivot]:
    for row in ws_temp.iter_rows():
        for cell in row:
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.border = border

# Freeze panes and auto filter
ws.freeze_panes = 'A2'
ws.auto_filter.ref = ws.dimensions
ws_customers.freeze_panes = 'A2'
ws_customers.auto_filter.ref = ws_customers.dimensions
ws_products.freeze_panes = 'A2'
ws_products.auto_filter.ref = ws_products.dimensions
ws_employees.freeze_panes = 'A2'
ws_employees.auto_filter.ref = ws_employees.dimensions
ws_analysis.freeze_panes = 'A2'
ws_pivot.freeze_panes = 'A2'

wb.save(excel_dir / 'Week3_Sales_Analysis.xlsx')

excel_formula_doc = textwrap.dedent('''
    # Excel Formulas Reference

    | Function | Formula | Purpose |
    |---|---|---|
    | IF | =IF(K2>5000,"High","Low") | Categorizes sales as High or Low based on SalesAmount |
    | COUNTIF | =COUNTIF(N:N,"High") | Counts the number of records that match the condition |
    | SUMIF | =SUMIF(N:N,"High",K:K) | Adds total sales for rows satisfying the condition |
    | VLOOKUP | =VLOOKUP(F2,Products!A:E,2,FALSE) | Retrieves the product name from ProductID |
    | XLOOKUP | =XLOOKUP(F2,Products!A:A,Products!B:B) | Retrieves related data more flexibly than VLOOKUP |

    ## Notes
    - In the workbook, `SalesAmount` is in column K of the `Sales_Data` sheet.
    - `Sales_Category` can be created in the analysis area using the formula above.
    - `Products` sheet columns are `ProductID`, `ProductName`, `Category`, `UnitPrice`, `Stock`.
''').strip() + '\n'
(excel_dir / 'Excel_Formulas.md').write_text(excel_formula_doc, encoding='utf-8')

pivot_doc = textwrap.dedent('''
    # Pivot Table Instructions

    1. Open the workbook and select the `Sales_Data` sheet.
    2. Click Insert > PivotTable.
    3. Place the pivot table in a new sheet named `Pivot_Table`.
    4. Drag `ProductName` to Rows.
    5. Drag `SalesAmount` to Values and set it to Sum.
    6. Rename the sheet output accordingly.
    7. Add a second pivot table with `City` in Rows and `SalesAmount` in Values.
    8. Insert a Column chart for product sales and a Bar chart for city sales.
    9. Format chart titles, axis labels, legend, and colors for presentation.
''').strip() + '\n'
(excel_dir / 'Pivot_Table_Instructions.md').write_text(pivot_doc, encoding='utf-8')

# Documentation files
(doc_dir := base / 'Documentation').mkdir(parents=True, exist_ok=True)
(sql_setup := doc_dir / 'SQL_Setup_Guide.md').write_text(textwrap.dedent('''
    # SQL Setup Guide

    1. Open MySQL Workbench.
    2. Create a new database named `Week3_Sales`.
    3. Run `SQL/01_create_database.sql`.
    4. Run `SQL/02_create_tables.sql`.
    5. Run `SQL/03_insert_data.sql`.
    6. Execute the task SQL files in sequence.

    ## Data validation checks
    - Ensure no duplicate primary keys.
    - Confirm foreign key relations exist.
    - Check that `SalesAmount = Quantity × UnitPrice` approximately.
''').strip() + '\n', encoding='utf-8')

(excel_setup := doc_dir / 'Excel_Setup_Guide.md').write_text(textwrap.dedent('''
    # Excel Setup Guide

    1. Open the generated workbook `Excel/Week3_Sales_Analysis.xlsx`.
    2. Review the `Sales_Data` sheet for raw records.
    3. Apply formatting: bold headers, currency formatting, date formatting, and borders.
    4. Freeze panes and add filters.
    5. Create formulas in the `Analysis` sheet.
    6. Use VLOOKUP/XLOOKUP to retrieve related record values.
    7. Generate pivot tables and charts.
''').strip() + '\n', encoding='utf-8')

(submission := doc_dir / 'Submission_Checklist.md').write_text(textwrap.dedent('''
    # Submission Checklist

    - [ ] SQL scripts run without errors
    - [ ] CSV datasets populated completely
    - [ ] Excel workbook created and formulas reviewed
    - [ ] Report completed with placeholders replaced
    - [ ] Screenshots saved in Screenshots folders
    - [ ] PDF exported for submission
    - [ ] Google Drive link prepared
''').strip() + '\n', encoding='utf-8')

# Report content
report_dir = base / 'Report'
ensure_path(report_dir)

report_md = textwrap.dedent('''
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
''').strip() + '\n'
(report_dir / 'Week3_Report.md').write_text(report_md, encoding='utf-8')
(report_dir / 'Report_Content.txt').write_text(report_md, encoding='utf-8')

screenshot_checklist = textwrap.dedent('''
    # Screenshot Checklist

    - [ ] SQL Query Output – Task 1
    - [ ] Aggregate Functions – Task 2
    - [ ] GROUP BY and HAVING – Task 3
    - [ ] INNER JOIN – Task 4
    - [ ] LEFT JOIN – Task 4
    - [ ] RIGHT JOIN – Task 4
    - [ ] Subquery output – Task 5
    - [ ] Excel formatting – Task 6
    - [ ] Conditional formatting – Task 7
    - [ ] Excel formula outputs – Task 7
    - [ ] VLOOKUP result – Task 8
    - [ ] XLOOKUP result – Task 8
    - [ ] Pivot table – Task 9
    - [ ] Chart – Task 9
''').strip() + '\n'
(report_dir / 'Screenshot_Checklist.md').write_text(screenshot_checklist, encoding='utf-8')

sql_output_guide = textwrap.dedent('''
    # SQL Output Guide

    ## Query 1: SELECT all records
    - What it does: Displays all rows from a table.
    - SQL concept: SELECT
    - Result: All customer records are returned.
    - Screenshot: Task 1 output.

    ## Query 2: Select specific columns
    - What it does: Shows only required columns.
    - SQL concept: SELECT with column specification
    - Result: Customer names and cities appear.
    - Screenshot: Task 1 output.

    ## Query 3: Column aliases
    - What it does: Renames display columns.
    - SQL concept: Column aliasing
    - Result: Customer_Name and Customer_City values are visible.
    - Screenshot: Task 1 output.

    ## Query 4: WHERE condition
    - What it does: Filters by rule such as price or quantity.
    - SQL concept: WHERE
    - Result: Matching products or sales are shown.
    - Screenshot: Task 2 output.

    ## Query 5: Aggregate functions
    - What it does: Calculates total, average, minimum, and maximum values.
    - SQL concept: COUNT(), SUM(), AVG(), MIN(), MAX()
    - Result: Business summary metrics appear.
    - Screenshot: Task 2 aggregate output.

    ## Query 6: GROUP BY and HAVING
    - What it does: Groups data by a category and filters grouped results.
    - SQL concept: GROUP BY and HAVING
    - Result: Summaries such as product totals and city counts appear.
    - Screenshot: Task 3 output.

    ## Query 7: INNER JOIN
    - What it does: Returns matching customer and sales rows.
    - SQL concept: JOIN
    - Result: Customer order records are combined.
    - Screenshot: Task 4 INNER JOIN.

    ## Query 8: LEFT JOIN
    - What it does: Shows all left-side table records with matches where available.
    - SQL concept: LEFT JOIN
    - Result: All customers and matching sales appear.
    - Screenshot: Task 4 LEFT JOIN.

    ## Query 9: RIGHT JOIN
    - What it does: Shows all right-side records with matching left-side data.
    - SQL concept: RIGHT JOIN
    - Result: Sales records and matching customer records appear.
    - Screenshot: Task 4 RIGHT JOIN.

    ## Query 10: Subquery examples
    - What it does: Compares values to a nested query result.
    - SQL concept: Subqueries
    - Result: Employees above average pay and products above average price appear.
    - Screenshot: Task 5 output.
''').strip() + '\n'
(report_dir / 'SQL_Output_Guide.md').write_text(sql_output_guide, encoding='utf-8')

# Screenshots directory placeholders
for screenshot_dir in [base / 'Screenshots' / 'SQL', base / 'Screenshots' / 'Excel']:
    screenshot_dir.mkdir(parents=True, exist_ok=True)

# Create empty docs for working files if needed
(base / 'Screenshots' / '.gitkeep').write_text('', encoding='utf-8')

# Validation script summary
validation_checks = []
# confirm data integrity
customer_ids = {r[0] for r in customers}
product_ids = {r[0] for r in products}
employee_ids = {r[0] for r in employees}
assert len(customers) >= 15 and len(products) >= 10 and len(employees) >= 5 and len(sales_rows) >= 30
for row in sales_rows:
    assert row['CustomerID'] in customer_ids
    assert row['ProductID'] in product_ids
    assert row['EmployeeID'] in employee_ids
    assert abs(float(row['SalesAmount']) - (float(row['Quantity']) * float(row['UnitPrice']))) < 0.01
    assert 1 <= row['Quantity'] <= 10

# Save a summary report
validation_summary = textwrap.dedent(f'''
    Project created successfully.
    Customers: {len(customers)}
    Products: {len(products)}
    Employees: {len(employees)}
    Sales: {len(sales_rows)}
    Workbook: {excel_dir / 'Week3_Sales_Analysis.xlsx'}
''').strip() + '\n'
(base / 'project_validation.txt').write_text(validation_summary, encoding='utf-8')

print('Project created in', base)
print(validation_summary)
