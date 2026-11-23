SELECT SysNo, ProductSysNo, UnitCost, NoTaxUnitCost, TaxRate, CustomerType, SaleType, IsWholeSale, LastUpdateTime
FROM dbo.Product_Price
WHERE ProductSysNo IN (478672, 297858, 443563, 13151, 15451)
ORDER BY ProductSysNo, SysNo
