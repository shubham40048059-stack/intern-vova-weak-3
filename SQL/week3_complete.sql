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
