SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema LIKE 'wmsx%'
  AND (
    table_name LIKE '%goods%' OR table_name LIKE '%product%' OR
    table_name LIKE '%weight%' OR table_name LIKE '%material%' OR
    table_name LIKE '%warehouse%' OR table_name LIKE '%owner%'
  )
ORDER BY table_schema, table_name
LIMIT 1000
