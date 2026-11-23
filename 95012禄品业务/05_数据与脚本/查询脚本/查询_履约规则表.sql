SELECT table_schema, table_name
FROM information_schema.tables
WHERE table_schema NOT IN ('mysql', 'sys', 'performance_schema', 'information_schema')
  AND (
    LOWER(table_schema) LIKE '%billing%'
    OR LOWER(table_schema) LIKE '%wmsx%'
    OR LOWER(table_name) LIKE '%billing%'
    OR LOWER(table_name) LIKE '%cost%'
    OR LOWER(table_name) LIKE '%carrier%'
    OR LOWER(table_name) LIKE '%logistics%'
  )
ORDER BY table_schema, table_name
LIMIT 1000
