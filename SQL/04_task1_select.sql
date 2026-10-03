-- Task 1: SELECT, aliases, and basic column selection
SELECT * FROM Customers;

SELECT CustomerName, City FROM Customers;

SELECT
    CustomerName AS Customer_Name,
    City AS Customer_City,
    Email AS Email_Address
FROM Customers;

-- Purpose: Understand how to retrieve records and rename output columns.
