-- Task 2: WHERE, ORDER BY, and aggregate functions
SELECT * FROM Products WHERE UnitPrice > 5000 ORDER BY UnitPrice DESC;

SELECT * FROM Sales WHERE Quantity >= 3 ORDER BY Quantity DESC;

SELECT CustomerName, City FROM Customers WHERE City = 'Delhi' ORDER BY CustomerName ASC;

SELECT COUNT(*) AS Total_Transactions FROM Sales;
SELECT SUM(Quantity * UnitPrice) AS Total_Revenue FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
SELECT AVG(Quantity * UnitPrice) AS Avg_Order_Value FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
SELECT MIN(Quantity * UnitPrice) AS Lowest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
SELECT MAX(Quantity * UnitPrice) AS Highest_Sale FROM Sales s JOIN Products p ON s.ProductID = p.ProductID;
