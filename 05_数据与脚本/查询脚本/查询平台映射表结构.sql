SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'dbo'
  AND TABLE_NAME IN (
    'Partner_Business_Relation', 'Partner_SubScribe_Relation',
    'PlatformMatchRelation', 'StorePlatform',
    'StorePlatform_SubscribeConfig', 'StorePlatform_SubscribeRelation'
  )
ORDER BY TABLE_NAME, ORDINAL_POSITION