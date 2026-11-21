SELECT table_name, table_rows
FROM information_schema.tables
WHERE table_schema = 'wmsx_billing'
ORDER BY table_name
LIMIT 500
