SELECT o.name, o.type_desc, s.base_object_name
FROM sys.objects AS o
LEFT JOIN sys.synonyms AS s ON o.object_id = s.object_id
WHERE o.name IN ('Product', 'Product_Basic', 'Products', 'Product_Info')
