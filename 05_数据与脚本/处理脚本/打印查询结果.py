#!/usr/bin/env python3
import json
import sys
from pathlib import Path
path = Path(sys.argv[1])
result = json.loads(path.read_text(encoding='utf-8'))
data = result.get('data') or {}
columns = data.get('column_list') or []
print('\t'.join(map(str, columns)))
for row in data.get('rows') or []:
    print('\t'.join('NULL' if value is None else str(value) for value in row))
print(f'共 {len(data.get("rows") or [])} 行；error={data.get("error")}；warning={data.get("warning")}')
