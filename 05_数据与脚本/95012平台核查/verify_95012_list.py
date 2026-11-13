#!/usr/bin/env python3
"""
95012 平台商品清单精确差集核验（可复现）

方法：
  本地 P 码清单 -> product 表映射为 SysNo / ProductBasicSysNo
  -> 与 Archery platform_product(PlatformSysNo=95012, IsDelete=0) 明细做精确差集。

依赖（本机特定，不提交）：
  - SQL 查询封装 run_sql.py（提供 DEFAULT_INSTANCE / make_client / run_sql），
    所在目录由环境变量 BENLAI_SQL_RUNNER_DIR 指定；旧机当时位于 lupin-sales/ 下。
    本机能力包只提供 archery_client.py，接口不同，不能直接顶替；
    若要在本机重查生产库，须先确认查询封装的接入方式（单独决策，本脚本不自动连库）。
  - benlai-trace-debug 能力包 + 真实 config.local.yaml 凭证

用法：
  python3 verify_95012_list.py [本地清单xlsx路径]
  默认本地清单：仓库内 01_源数据/禄品4940_筛选表_2026-11-09.xlsx（表头在第二行）
  可用环境变量 BENLINK_SOURCE_XLSX 覆盖；重查前须设置 BENLAI_SQL_RUNNER_DIR。
  核准结论与差集结果已随仓库保存，日常续作无需重跑本脚本。
"""
import os, sys, json, openpyxl

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# 源表默认取自仓库内 01_源数据/（已随仓库提交，走 git-lfs）；可用环境变量覆盖
LOCAL_DEFAULT = os.environ.get(
    "BENLINK_SOURCE_XLSX",
    os.path.join(REPO, "01_源数据/禄品4940_筛选表_2026-11-09.xlsx"),
)

# 查询封装只在本机提供，不写死任何机器的绝对路径；未配置时仅在实际重查时给出明确提示
SQL_RUNNER_DIR = os.environ.get("BENLAI_SQL_RUNNER_DIR")
if SQL_RUNNER_DIR:
    sys.path.insert(0, SQL_RUNNER_DIR)
try:
    import run_sql  # noqa: E402
except ModuleNotFoundError:
    run_sql = None

DB = "benlai_business_product"
PLATFORM = 95012


def main():
    if run_sql is None:
        raise SystemExit(
            "未找到 SQL 查询封装 run_sql.py：请设置 BENLAI_SQL_RUNNER_DIR 指向其所在目录。\n"
            "本机能力包只提供 archery_client.py，接口不同，不能直接顶替；接入方式需单独确认。\n"
            "核准结论与差集结果已随仓库保存，日常续作无需重跑本脚本。"
        )
    local_path = sys.argv[1] if len(sys.argv) > 1 else LOCAL_DEFAULT
    client = run_sql.make_client()
    inst = run_sql.DEFAULT_INSTANCE

    # 1) 本地 P 码
    wb = openpyxl.load_workbook(local_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[1])
    pc = hdr.index("商品编号")
    local = [str(r[pc]).strip() for r in rows[2:]
             if r[pc] is not None and str(r[pc]).strip()]
    local_set = set(local)
    print("本地 P 码数:", len(local_set))

    # 2) 平台明细
    sql_pp = (f"SELECT ProductSysNo, ProductBasicSysNo FROM platform_product "
              f"WHERE PlatformSysNo={PLATFORM} AND IsDelete=0")
    c1, r1 = run_sql.run_sql(client, sql_pp, DB, instance=inst, limit=10000)
    pp_sysno = set(int(x[0]) for x in r1)
    pp_basic = set(int(x[1]) for x in r1 if x[1] is not None)
    print("platform_product 数:", len(r1),
          " ProductSysNo集:", len(pp_sysno),
          " ProductBasicSysNo集:", len(pp_basic))

    # 3) 本地 P 码 -> product 映射
    pids_str = ",".join(f"'{p}'" for p in local)
    sql_map = (f"SELECT ProductId, SysNo, ProductBasicSysNo FROM product "
               f"WHERE ProductId IN ({pids_str})")
    c2, r2 = run_sql.run_sql(client, sql_map, DB, instance=inst, limit=10000)
    m_sysno, m_basic = {}, {}
    for x in r2:
        pid = (x[0] or "").strip()
        m_sysno[pid] = int(x[1]) if x[1] is not None else None
        m_basic[pid] = int(x[2]) if x[2] is not None else None
    local_sysno = set(v for v in m_sysno.values() if v is not None)
    local_basic = set(v for v in m_basic.values() if v is not None)
    print("本地 P 码->SysNo 集:", len(local_sysno),
          " 本地 P 码->ProductBasicSysNo 集:", len(local_basic))

    # 4) 差集
    result = {
        "local_p_rows": len(local_set),
        "platform_product_rows": len(r1),
        "local_p_to_product_hit": len(r2),
        "local_p_unmatched_in_product": len(local_set - set(m_sysno.keys())),
    }
    pairs = [
        ("SysNo<->ProductSysNo", local_sysno, pp_sysno),
        ("ProductBasicSysNo<->ProductBasicSysNo", local_basic, pp_basic),
        ("SysNo<->ProductBasicSysNo(错位,仅参考)", local_sysno, pp_basic),
        ("ProductBasicSysNo<->ProductSysNo(错位,仅参考)", local_basic, pp_sysno),
    ]
    result["pairs"] = {}
    for name, ls, ps in pairs:
        result["pairs"][name] = {
            "intersection": len(ls & ps),
            "only_local": len(ls - ps),
            "only_platform": len(ps - ls),
        }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


if __name__ == "__main__":
    main()
