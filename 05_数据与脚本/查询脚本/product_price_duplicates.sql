SELECT TOP 20 ProductSysNo, COUNT(*) AS price_rows
FROM dbo.Product_Price
GROUP BY ProductSysNo
HAVING COUNT(*) > 1
ORDER BY COUNT(*) DESC, ProductSysNo
