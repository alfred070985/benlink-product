SELECT
    COUNT(*) AS price_rows,
    COUNT(DISTINCT ProductSysNo) AS product_count,
    SUM(CASE WHEN UnitCost IS NULL THEN 1 ELSE 0 END) AS null_unit_cost_rows,
    SUM(CASE WHEN NoTaxUnitCost IS NULL THEN 1 ELSE 0 END) AS null_no_tax_cost_rows,
    MIN(LastUpdateTime) AS earliest_update,
    MAX(LastUpdateTime) AS latest_update
FROM dbo.Product_Price
