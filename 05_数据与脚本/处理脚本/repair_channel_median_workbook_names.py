#!/usr/bin/env python3
import json
import zipfile
from pathlib import Path

path = Path(__file__).with_name("来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx")
temp = path.with_suffix(".repair-temp.xlsx")
replacement_count = 0
with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename.endswith(".xml"):
            text = data.decode("utf-8")
            replacement_count += text.count("_x005F_x005f_")
            text = text.replace("_x005F_x005f_", "_x005f_")
            data = text.encode("utf-8")
        zout.writestr(info, data)
if replacement_count != 2:
    temp.unlink(missing_ok=True)
    raise RuntimeError(f"Expected 2 escaped names, found {replacement_count}")
temp.replace(path)
print(json.dumps({"repaired_product_names": replacement_count, "file_size_bytes": path.stat().st_size}, ensure_ascii=False))
