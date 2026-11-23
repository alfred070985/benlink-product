#!/usr/bin/env python3
import argparse
import json
import sys
from pathlib import Path
from archery_runtime import ARCHERY_DIR
INSTANCE = "[online]生产PolarDB-Newretail(prod-newretail.rwlb.rds.aliyuncs.com)"
DB_NAME = "newretail_oms"
sys.path.insert(0, str(ARCHERY_DIR))
from archery_client import ArcheryClient


def execute(client, sql, limit_num=1000, instance=INSTANCE, db_name=DB_NAME):
    referer = client.base_url + "/sqlquery/"
    page = client.session.get(referer, timeout=30)
    page.raise_for_status()
    client._get_csrf_token(page.text)
    headers = client._get_headers(referer)
    headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"
    response = client.session.post(
        client.base_url + "/query/",
        data={
            "instance_name": instance,
            "db_name": db_name,
            "schema_name": "",
            "tb_name": "",
            "sql_content": sql,
            "limit_num": str(limit_num),
        },
        headers=headers,
        timeout=180,
    )
    response.raise_for_status()
    return response.json()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sql-file", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--limit", type=int, default=1000)
    parser.add_argument("--instance", default=INSTANCE)
    parser.add_argument("--db", default=DB_NAME)
    args = parser.parse_args()
    sql = Path(args.sql_file).read_text(encoding="utf-8").strip()
    if not sql.upper().startswith("SELECT"):
        raise SystemExit("Only SELECT queries are permitted")
    client = ArcheryClient()
    if not client.login():
        raise SystemExit("Archery login failed")
    result = execute(client, sql, args.limit, instance=args.instance, db_name=args.db)
    Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    data = result.get("data") or {}
    print(json.dumps({
        "status": result.get("status"),
        "msg": result.get("msg"),
        "data_type": type(data).__name__,
        "data_keys": list(data.keys()) if isinstance(data, dict) else None,
        "rows": len(data.get("rows", [])) if isinstance(data, dict) else None,
        "column_list": data.get("column_list") if isinstance(data, dict) else None,
        "output": args.output,
    }, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
