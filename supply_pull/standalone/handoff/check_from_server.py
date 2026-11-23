#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
给 IT 同事用：查这台机器「上网时对外显示的 IP」。

- 不含任何账号、密钥；只用 Python 自带功能，不用安装任何东西
- 不写文件、不上传数据，结果只打印在屏幕上
- 用法：python3 check_from_server.py
  把最后那行「结论」发回即可
"""

import re
import urllib.request

WHITELIST = ("39.102.184.125", "39.102.184.126", "39.102.184.127",
             "39.102.184.96", "8.153.82.69")
SERVICES = ("https://api.ipify.org", "https://ifconfig.me",
            "https://myip.ipip.net")


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "curl/8"})
    return urllib.request.urlopen(req, timeout=8).read().decode("utf-8", "ignore").strip()


def main():
    print("正在检查这台机器的上网出口 IP ...\n")
    found = []
    for url in SERVICES:
        try:
            text = fetch(url)
            m = re.search(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", text)
            if m:
                found.append(m.group(0))
                print(f"  {url}  ->  {m.group(0)}")
            else:
                print(f"  {url}  ->  未取到（返回片段：{text[:60]}）")
        except Exception as exc:
            print(f"  {url}  ->  不通（{type(exc).__name__}）")

    ips = set(found)
    hit = sorted(ips & set(WHITELIST))
    print("\n" + "-" * 52)
    if hit:
        print(f"结论：命中白名单 {', '.join(hit)} —— 这台机器可用")
    elif ips:
        print(f"结论：出口 IP 为 {', '.join(sorted(ips))} —— 不在白名单 5 个 IP 内，这台机器不匹配")
    else:
        print("结论：三个查 IP 的服务都不通，请把这台机器能访问的外网地址告诉我们")
    print("白名单：" + " / ".join(WHITELIST))
    print("-" * 52)


if __name__ == "__main__":
    main()
