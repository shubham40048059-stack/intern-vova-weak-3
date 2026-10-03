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
