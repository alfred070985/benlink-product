SELECT
  (SELECT COUNT_BIG(1) FROM dbo.Platform_Product WHERE PlatformSysNo = 95012) AS PlatformProductRows,
  (SELECT COUNT_BIG(1) FROM dbo.Platform_Product_WebSite WHERE PlatformSysNo = 95012) AS PlatformWebSiteRows,
  (SELECT COUNT_BIG(1) FROM dbo.Platform_Product_Delivery WHERE PlatformSysNo = 95012) AS PlatformDeliveryRows,
  (SELECT COUNT_BIG(1) FROM dbo.Platform_Product_Exclusive WHERE PlatformSysNo = 95012) AS PlatformExclusiveRows,
  (SELECT COUNT_BIG(DISTINCT ProductSysNo) FROM dbo.Platform_Product_Exclusive WHERE PlatformSysNo = 95012 AND IsDelete = 0) AS ActiveExclusiveProducts;
