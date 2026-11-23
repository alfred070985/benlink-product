SELECT table_name, table_rows
FROM information_schema.tables
WHERE table_schema = 'wmsx_basic'
  AND (
    table_name LIKE '%pack%' OR table_name LIKE '%material%' OR
    table_name LIKE '%delivery%' OR table_name LIKE '%consumable%'
  )
ORDER BY table_name
LIMIT 500
