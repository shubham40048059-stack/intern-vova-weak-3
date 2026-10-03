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
