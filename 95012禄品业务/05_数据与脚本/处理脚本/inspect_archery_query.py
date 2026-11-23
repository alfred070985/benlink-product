#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from archery_runtime import ARCHERY_DIR
sys.path.insert(0, str(ARCHERY_DIR))
from archery_client import ArcheryClient

client = ArcheryClient()
if not client.login():
    raise SystemExit("Archery login failed")
url = client.base_url + "/sqlquery/"
response = client.session.get(url, timeout=30)
response.raise_for_status()
Path(__file__).with_name("archery_sqlquery.html").write_text(response.text, encoding="utf-8")
soup = BeautifulSoup(response.text, "html.parser")
forms = []
for form in soup.find_all("form"):
    forms.append({
        "id": form.get("id"),
        "action": urljoin(url, form.get("action") or ""),
        "method": form.get("method"),
        "inputs": [{"name": x.get("name"), "type": x.get("type")} for x in form.find_all(["input", "textarea"])],
        "selects": [{"name": x.get("name"), "id": x.get("id"), "options": [{"value": o.get("value"), "text": o.get_text(strip=True)} for o in x.find_all("option")]} for x in form.find_all("select")],
    })
print(json.dumps({"url": response.url, "status": response.status_code, "content_type": response.headers.get("content-type"), "length": len(response.content), "body_prefix": response.text[:3000], "title": soup.title.get_text(strip=True) if soup.title else None, "forms": forms}, ensure_ascii=False, indent=2))
