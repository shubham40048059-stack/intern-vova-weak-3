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
