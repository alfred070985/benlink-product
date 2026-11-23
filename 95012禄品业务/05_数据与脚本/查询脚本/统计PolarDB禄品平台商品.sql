SELECT
  (SELECT COUNT(*) FROM platform_product WHERE PlatformSysNo = 95012) AS platform_product_rows,
  (SELECT COUNT(*) FROM platform_product WHERE PlatformSysNo = 95012 AND IsDelete = 0) AS active_platform_product_rows,
  (SELECT COUNT(DISTINCT ProductSysNo) FROM platform_product WHERE PlatformSysNo = 95012 AND IsDelete = 0) AS active_platform_products,
  (SELECT COUNT(*) FROM platform_product_website WHERE PlatformSysNo = 95012) AS platform_website_rows,
  (SELECT COUNT(*) FROM platform_product_exclusive WHERE PlatformSysNo = 95012) AS platform_exclusive_rows,
  (SELECT COUNT(DISTINCT ProductSysNo) FROM platform_product_exclusive WHERE PlatformSysNo = 95012 AND IsDelete = 0) AS active_exclusive_products;
