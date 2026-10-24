/*
PROJECT: Aurelia Consumer Products — Profitability Analysis
TABLE: dbo.AureliaSalesData

Purpose: validate the dataset and analyze overall, regional, channel, salesperson,
monthly, product, and return-related performance.

Interpretation notes:
- COUNT(*) counts records; call them orders only if each row represents one order.
- Negative profit values need investigation but are not automatically data errors.
- Return_Value_Rate_Percent is a screening metric; confirm field definitions first.
- Net_Contribution is analyzed as stored. Confirm its business calculation before
  using it for financial decisions.
*/


-- ============================================================
-- 1. DATA QUALITY AND VALIDATION
-- ============================================================
-- Source data was also validated in Excel.
-- Negative financial values were identified for further review.

SELECT COUNT(*) AS TotalRows
FROM dbo.AureliaSalesData;

SELECT
    Order_ID,
    COUNT(*) AS NumberOfOccurrences
FROM dbo.AureliaSalesData
GROUP BY Order_ID
HAVING COUNT(*) > 1;

SELECT COUNT(*) AS MissingOrderIDs
FROM dbo.AureliaSalesData
WHERE Order_ID IS NULL;

SELECT COUNT(*) AS MissingProductNames
FROM dbo.AureliaSalesData
WHERE Product_Name IS NULL;

SELECT COUNT(*) AS MissingGrossRevenue
FROM dbo.AureliaSalesData
WHERE Gross_Revenue IS NULL;

SELECT COUNT(*) AS MissingNetRevenue
FROM dbo.AureliaSalesData
WHERE Net_Revenue IS NULL;

SELECT COUNT(*) AS MissingGrossProfit
FROM dbo.AureliaSalesData
WHERE Gross_Profit IS NULL;

SELECT COUNT(*) AS MissingReturnValue
FROM dbo.AureliaSalesData
WHERE Return_Value IS NULL;

SELECT COUNT(*) AS MissingReturnFrictionCost
FROM dbo.AureliaSalesData
WHERE Return_Friction_Cost IS NULL;

SELECT COUNT(*) AS MissingNetContribution
FROM dbo.AureliaSalesData
WHERE Net_Contribution IS NULL;

SELECT COUNT(*) AS MissingOrderDates
FROM dbo.AureliaSalesData
WHERE Order_Date IS NULL;

SELECT COUNT(*) AS MissingRegions
FROM dbo.AureliaSalesData
WHERE Region IS NULL;

SELECT COUNT(*) AS MissingSalesChannels
FROM dbo.AureliaSalesData
WHERE Sales_Channel IS NULL;

SELECT COUNT(*) AS MissingQuantity
FROM dbo.AureliaSalesData
WHERE Quantity IS NULL;

SELECT COUNT(*) AS Missing_Product_ID
FROM dbo.AureliaSalesData
WHERE Product_ID IS NULL;

SELECT
    MIN(Order_Date) AS First_Order_Date,
    MAX(Order_Date) AS Last_Order_Date
FROM dbo.AureliaSalesData;

SELECT COUNT(*) AS Invalid_Quantity_Rows
FROM dbo.AureliaSalesData
WHERE Quantity <= 0;

SELECT COUNT(*) AS Negative_Financial_Rows
FROM dbo.AureliaSalesData
WHERE Gross_Revenue < 0
   OR Net_Revenue < 0
   OR Gross_Profit < 0
   OR Return_Value < 0
   OR Return_Friction_Cost < 0
   OR Net_Contribution < 0;

SELECT
    SUM(CASE WHEN Gross_Revenue < 0 THEN 1 ELSE 0 END)
        AS Negative_Gross_Revenue,
    SUM(CASE WHEN Net_Revenue < 0 THEN 1 ELSE 0 END)
        AS Negative_Net_Revenue,
    SUM(CASE WHEN Gross_Profit < 0 THEN 1 ELSE 0 END)
        AS Negative_Gross_Profit,
    SUM(CASE WHEN Return_Value < 0 THEN 1 ELSE 0 END)
        AS Negative_Return_Value,
    SUM(CASE WHEN Return_Friction_Cost < 0 THEN 1 ELSE 0 END)
        AS Negative_Return_Friction_Cost,
    SUM(CASE WHEN Net_Contribution < 0 THEN 1 ELSE 0 END)
        AS Negative_Net_Contribution
FROM dbo.AureliaSalesData;

SELECT TOP 10
    Order_ID,
    Product_Name,
    Region,
    Sales_Channel,
    Gross_Revenue,
    Net_Revenue,
    Gross_Profit,
    Return_Value,
    Return_Friction_Cost,
    Net_Contribution
FROM dbo.AureliaSalesData
WHERE Gross_Profit < 0
ORDER BY Gross_Profit ASC;

SELECT
    COUNT(*) AS Negative_Gross_Profit_Rows,
    SUM(CASE WHEN Gross_Revenue < 0 THEN 1 ELSE 0 END)
        AS Negative_Revenue_Rows,
    SUM(CASE WHEN Gross_Revenue = 0 THEN 1 ELSE 0 END)
        AS Zero_Revenue_Rows,
    SUM(CASE WHEN Gross_Revenue > 0 THEN 1 ELSE 0 END)
        AS Positive_Revenue_Rows
FROM dbo.AureliaSalesData
WHERE Gross_Profit < 0;


-- ============================================================
-- 2. OVERALL BUSINESS PERFORMANCE
-- ============================================================

SELECT
    COUNT(*) AS Total_Rows,
    SUM(Gross_Revenue) AS Total_Gross_Revenue,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Gross_Profit) AS Total_Gross_Profit,
    SUM(Return_Value) AS Total_Return_Value,
    SUM(Return_Friction_Cost) AS Total_Return_Friction_Cost,
    SUM(Net_Contribution) AS Total_Net_Contribution,
    CAST(
        100.0 * SUM(Gross_Profit) /
        NULLIF(SUM(Net_Revenue), 0)
        AS DECIMAL(10,2)
    ) AS Gross_Profit_Margin_Percent
FROM dbo.AureliaSalesData;


-- ============================================================
-- 3. REGIONAL PROFITABILITY
-- ============================================================

SELECT
    Region,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Gross_Profit) AS Total_Gross_Profit,
    SUM(Net_Contribution) AS Total_Net_Contribution,
    CAST(
        100.0 * SUM(Gross_Profit) /
        NULLIF(SUM(Net_Revenue), 0)
        AS DECIMAL(10,2)
    ) AS Gross_Profit_Margin_Percent
FROM dbo.AureliaSalesData
GROUP BY Region
ORDER BY Total_Net_Contribution DESC;


-- ============================================================
-- 4. SALES CHANNEL PROFITABILITY AND MARGIN
-- ============================================================

SELECT
    Sales_Channel,
    COUNT(*) AS Total_Rows,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Gross_Profit) AS Total_Gross_Profit,
    SUM(Net_Contribution) AS Total_Net_Contribution,
    CAST(
        100.0 * SUM(Gross_Profit) /
        NULLIF(SUM(Net_Revenue), 0)
        AS DECIMAL(10,2)
    ) AS Gross_Profit_Margin_Percent
FROM dbo.AureliaSalesData
GROUP BY Sales_Channel
ORDER BY Total_Net_Contribution DESC;


-- ============================================================
-- 5. SALESPERSON PERFORMANCE
-- ============================================================

SELECT
    Salesperson_ID,
    Salesperson_Name,
    COUNT(*) AS Total_Rows,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Gross_Profit) AS Total_Gross_Profit,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY
    Salesperson_ID,
    Salesperson_Name
ORDER BY Total_Net_Contribution DESC;


-- ============================================================
-- 6. MONTHLY PERFORMANCE
-- ============================================================

SELECT
    YEAR(Order_Date) AS Sales_Year,
    MONTH(Order_Date) AS Sales_Month,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Gross_Profit) AS Total_Gross_Profit,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY
    YEAR(Order_Date),
    MONTH(Order_Date)
ORDER BY
    Sales_Year,
    Sales_Month;


-- ============================================================
-- 7. RETURNS ANALYSIS BY PRODUCT
-- ============================================================

SELECT
    Product_Name,
    SUM(Return_Value) AS Total_Return_Value,
    SUM(Return_Friction_Cost) AS Total_Return_Friction_Cost,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY Product_Name
ORDER BY Total_Return_Friction_Cost DESC;


-- ============================================================
-- 8. PERFORMANCE BY REGION AND SALES CHANNEL
-- ============================================================

SELECT
    Region,
    Sales_Channel,
    COUNT(*) AS Total_Rows,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY
    Region,
    Sales_Channel
ORDER BY Total_Net_Contribution DESC;


-- ============================================================
-- 9. TOP 10 PRODUCTS BY NET CONTRIBUTION
-- ============================================================

SELECT TOP 10
    Product_ID,
    Product_Name,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY
    Product_ID,
    Product_Name
ORDER BY Total_Net_Contribution DESC;


-- ============================================================
-- 10. BOTTOM 10 PRODUCTS BY NET CONTRIBUTION
-- ============================================================

SELECT TOP 10
    Product_ID,
    Product_Name,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Net_Contribution) AS Total_Net_Contribution
FROM dbo.AureliaSalesData
GROUP BY
    Product_ID,
    Product_Name
ORDER BY Total_Net_Contribution ASC;


-- ============================================================
-- 11. PRODUCT RETURN-VALUE RATE (SCREENING METRIC)
-- ============================================================

SELECT
    Product_ID,
    Product_Name,
    SUM(Net_Revenue) AS Total_Net_Revenue,
    SUM(Return_Value) AS Total_Return_Value,
    CAST(
        100.0 * SUM(Return_Value) /
        NULLIF(SUM(Net_Revenue), 0)
        AS DECIMAL(10,2)
    ) AS Return_Value_Rate_Percent
FROM dbo.AureliaSalesData
GROUP BY
    Product_ID,
    Product_Name
ORDER BY Return_Value_Rate_Percent DESC;