SELECT table_schema, table_name, table_rows
FROM information_schema.tables
WHERE table_schema LIKE 'wmsx%'
  AND (table_name LIKE '%package%' OR table_name LIKE '%pack%')
ORDER BY table_schema, table_name
LIMIT 500
