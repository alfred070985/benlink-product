#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from archery_runtime import ARCHERY_DIR
sys.path.insert(0, str(ARCHERY_DIR))
from archery_client import ArcheryClient

client = ArcheryClient()
if not client.login():
    raise SystemExit("Archery login failed")
response = client.session.get(
    client.base_url + "/group/user_all_instances/",
    params={"tag_codes": "can_read"},
    timeout=30,
)
response.raise_for_status()
result = response.json()
data = result.get("data", [])
print(json.dumps({
    "status": result.get("status"),
    "msg": result.get("msg"),
    "count": len(data),
    "matching_instances": [x for x in data if x.get("db_type") in ("mysql", "mssql") and str(x.get("instance_name", "")).lower().startswith("[online]")],
}, ensure_ascii=False, indent=2))
