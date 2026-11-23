SELECT table_name, column_name, column_type, ordinal_position
FROM information_schema.columns
WHERE table_schema = 'wmsx_billing'
ORDER BY table_name, ordinal_position
LIMIT 1000
