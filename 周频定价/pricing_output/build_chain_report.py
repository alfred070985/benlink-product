#!/usr/bin/env python3
"""生成分销标准池定价的链路说明 HTML（周频复跑，数字全部取自当轮数据文件）。

数据链跑完后执行：
    python build_chain_report.py
数字来源（均取目录下最新一份）：
    源数据_分销池四站上架关系_*.jsonl   → 关系条数、去重商品数
    标准池候选_四站全上架_*.jsonl        → 四站全上架数
    标准池_四站全上架有库存_*.jsonl      → 有库存数
    源数据_标准池ERP站点成本_*.jsonl     → 成本明细条数
    标准池京东相似匹配_*.jsonl           → 采信数、候选数
    定价统计_*.json                      → 成品统计（发货/对标/建议档/空间率/公式校验）
输出：链路说明_分销标准池定价_<当天>.html
"""
import glob
import json
from datetime import datetime
from pathlib import Path
from string import Template

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
OUT = PRICING_DIR / f"链路说明_分销标准池定价_{TODAY}.html"


def latest(pattern):
    files = sorted(glob.glob(str(PRICING_DIR / pattern)))
    if not files:
        raise SystemExit(f"找不到数据文件: {pattern}（请先跑数据链）")
    return Path(files[-1])


def count_lines(p):
    n = 0
    with p.open(encoding="utf-8") as f:
        for _ in f:
            n += 1
    return n


def fmt(n):
    return f"{n:,}" if isinstance(n, int) else str(n)


def collect():
    rel_file = latest("源数据_分销池四站上架关系_*.jsonl")
    rel_total = count_lines(rel_file)
    dedup = set()
    with rel_file.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                dedup.add(json.loads(line)["productSysNo"])

    four_file = latest("标准池候选_四站全上架_*.jsonl")
    pool_file = latest("标准池_四站全上架有库存_*.jsonl")
    cost_file = latest("源数据_标准池ERP站点成本_*.jsonl")

    matched = records = 0
    match_file = latest("标准池京东相似匹配_*.jsonl")
    with match_file.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                records += 1
                if json.loads(line).get("采纳"):
                    matched += 1

    stats_file = latest("定价统计_*.json")
    stats = json.loads(stats_file.read_text(encoding="utf-8"))

    pool_n = stats["总行数"]
    src = stats["对标来源"]
    real_n = sum(v for k, v in src.items() if k not in ("ERP兜底", "无"))
    erp_n = src.get("ERP兜底", 0)
    sug_real = stats["建议档分布"].get("真实对标", {})
    sug_erp = stats["建议档分布"].get("ERP兜底", {})
    med_real = stats["建议档空间率中位数"].get("真实对标")
    med_erp = stats["建议档空间率中位数"].get("ERP兜底")
    chk = stats["公式校验"]

    return {
        "DATE": TODAY,
        "REL_TOTAL": fmt(rel_total),
        "DEDUP": fmt(len(dedup)),
        "FOUR": fmt(count_lines(four_file)),
        "POOL": fmt(count_lines(pool_file)),
        "COST_ROWS": fmt(count_lines(cost_file)),
        "MATCHED": fmt(matched),
        "MATCH_RATE": f"{matched / pool_n:.0%}" if pool_n else "-",
        "CANDIDATES": fmt(records - matched),
        "REL_FILE": rel_file.name,
        "FOUR_FILE": four_file.name,
        "POOL_FILE": pool_file.name,
        "COST_FILE": cost_file.name,
        "MATCH_FILE": match_file.name,
        "FINAL_ROWS": fmt(pool_n),
        "SHIP_DIRECT": fmt(stats["发货类型"].get("直发", 0)),
        "SHIP_WARE": fmt(stats["发货类型"].get("仓发", 0)),
        "SHIP_DIRECT_PCT": f"{stats['发货类型'].get('直发', 0) / pool_n:.1%}" if pool_n else "-",
        "REAL_N": fmt(real_n),
        "REAL_RATE": f"{real_n / pool_n:.0%}" if pool_n else "-",
        "ERP_N": fmt(erp_n),
        "REAL_DETAIL": " · ".join(f"{k} {fmt(v)}" for k, v in src.items() if k not in ("ERP兜底", "无")),
        "DIFF_SPEC": stats.get("天镜特殊", {}).get("diff_spec无价未折算", 0),
        "COST_MISS": stats["成本缺失"],
        "COST_COVER": f"{(pool_n - stats['成本缺失']) / pool_n:.3%}" if pool_n else "-",
        "SUG_JD1": fmt(sug_real.get("档1", 0)), "SUG_JD2": fmt(sug_real.get("档2", 0)), "SUG_JD3": fmt(sug_real.get("档3", 0)),
        "SUG_ERP1": fmt(sug_erp.get("档1", 0)), "SUG_ERP2": fmt(sug_erp.get("档2", 0)), "SUG_ERP3": fmt(sug_erp.get("档3", 0)),
        "MED_JD": f"{med_real:.1%}" if med_real is not None else "-",
        "MED_ERP": f"{med_erp:.1%}" if med_erp is not None else "-",
        "CHK_CELLS": fmt(chk["格数"]),
        "CHK_ERR": chk["错误数"],
        "XLSX": f"分销标准池定价_{stats.get('日期', TODAY)}.xlsx",
    }


TPL = Template(r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>分销标准池定价 · 全链路说明</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
    background: #f5f5f3;
    color: #2c2c2a;
    line-height: 1.65;
    font-size: 14px;
  }
  .page { max-width: 1080px; margin: 0 auto; padding: 32px 24px 64px; }

  .hero {
    background: #ffffff;
    border: 0.5px solid #d3d1c7;
    border-radius: 12px;
    padding: 28px 32px;
    margin-bottom: 24px;
  }
  .hero h1 { font-size: 22px; font-weight: 600; margin-bottom: 8px; }
  .hero .sub { color: #5f5e5a; font-size: 13px; margin-bottom: 16px; }
  .hero .oneline {
    background: #eeedfe;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 13.5px;
    color: #26215c;
  }
  .hero .oneline b { font-weight: 600; }

  .meta-row { display: flex; gap: 12px; margin-top: 14px; flex-wrap: wrap; }
  .meta-chip {
    font-size: 12px; color: #444441;
    background: #f1efe8; border-radius: 6px; padding: 4px 10px;
  }
  .meta-chip b { color: #26215c; font-weight: 600; }

  .funnel {
    display: flex; align-items: stretch; gap: 0;
    background: #ffffff; border: 0.5px solid #d3d1c7; border-radius: 12px;
    padding: 20px 16px; margin-bottom: 24px; overflow-x: auto;
  }
  .fstep { flex: 1; min-width: 130px; text-align: center; padding: 0 6px; }
  .fnum { font-size: 24px; font-weight: 600; color: #185fa5; }
  .flabel { font-size: 12px; color: #5f5e5a; margin-top: 4px; }
  .farrow { align-self: center; color: #b4b2a9; font-size: 18px; padding: 0 2px; flex: 0 0 auto; }
  .fstep.final .fnum { color: #3b6d11; }
  .fstep.warn .fnum { color: #854f0b; }

  .section-title {
    font-size: 16px; font-weight: 600; margin: 32px 0 14px;
    padding-left: 10px; border-left: 3px solid #534ab7;
  }
  .stage {
    background: #ffffff; border: 0.5px solid #d3d1c7; border-radius: 12px;
  }
  .stage-head {
    display: flex; align-items: center; gap: 14px;
    padding: 16px 24px; border-bottom: 0.5px solid #eceae4;
  }
  .stage-no {
    width: 34px; height: 34px; border-radius: 50%;
    background: #e6f1fb; color: #185fa5;
    display: flex; align-items: center; justify-content: center;
    font-weight: 600; font-size: 15px; flex: 0 0 auto;
  }
  .stage-title { font-size: 15.5px; font-weight: 600; }
  .stage-big { margin-left: auto; text-align: right; }
  .stage-big .n { font-size: 22px; font-weight: 600; color: #185fa5; }
  .stage-big .t { font-size: 11.5px; color: #5f5e5a; }
  .stage-body { padding: 14px 24px 18px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px 24px; }
  .kv { font-size: 12.5px; }
  .kv .k { color: #888780; margin-bottom: 2px; font-size: 11.5px; }
  .kv .v { color: #2c2c2a; }
  .kv .v code {
    background: #f1efe8; border-radius: 4px; padding: 1px 6px;
    font-family: "SF Mono", Menlo, monospace; font-size: 11.5px;
  }
  .rule-box {
    grid-column: 1 / -1;
    background: #faf9f6; border: 0.5px solid #e5e3db; border-radius: 8px;
    padding: 10px 14px; font-size: 12.5px; color: #444441;
  }
  .rule-box b { color: #2c2c2a; }
  .tag {
    display: inline-block; font-size: 11px; border-radius: 5px; padding: 1px 8px; margin-right: 6px;
    font-weight: 500;
  }
  .tag.erp { background: #e6f1fb; color: #185fa5; }
  .tag.db { background: #eeedfe; color: #534ab7; }
  .tag.algo { background: #e1f5ee; color: #0f6e56; }
  .tag.user { background: #faeeda; color: #854f0b; }

  .connector { width: 2px; height: 26px; background: #b4b2a9; margin: 0 auto; }

  .parallel { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .parallel .stage { height: 100%; }
  .merge { width: 60%; height: 2px; background: #b4b2a9; margin: 0 auto; }

  .final-card {
    background: #ffffff; border: 1px solid #97c459; border-radius: 12px; overflow: hidden;
  }
  .final-head {
    background: #eaf3de; padding: 18px 24px;
    display: flex; align-items: center; gap: 14px;
  }
  .final-head .n { font-size: 20px; font-weight: 600; color: #27500a; }
  .final-head .file { font-family: "SF Mono", Menlo, monospace; font-size: 12.5px; color: #3b6d11; }
  .final-body { padding: 16px 24px; display: grid; grid-template-columns: 1fr 1fr; gap: 12px 24px; }
  .stat-chip {
    background: #faf9f6; border: 0.5px solid #e5e3db; border-radius: 8px; padding: 10px 14px;
  }
  .stat-chip .n { font-size: 19px; font-weight: 600; color: #2c2c2a; }
  .stat-chip .t { font-size: 11.5px; color: #888780; margin-top: 2px; }
  .stat-chip.hi .n { color: #3b6d11; }
  .stat-chip.warn .n { color: #854f0b; }

  table { width: 100%; border-collapse: collapse; font-size: 12.5px; background: #fff; }
  th, td { text-align: left; padding: 9px 14px; border-bottom: 0.5px solid #eceae4; vertical-align: top; }
  th { background: #f1efe8; color: #444441; font-weight: 600; font-size: 12px; }
  .table-wrap { border: 0.5px solid #d3d1c7; border-radius: 12px; overflow: hidden; }

  .file-list { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .file-item {
    background: #ffffff; border: 0.5px solid #d3d1c7; border-radius: 8px;
    padding: 10px 14px; font-size: 12px;
  }
  .file-item .fn { font-family: "SF Mono", Menlo, monospace; font-size: 11.5px; color: #185fa5; word-break: break-all; }
  .file-item .fd { color: #888780; margin-top: 3px; }

  .notice {
    background: #faeeda; border: 0.5px solid #ef9f27; border-radius: 12px;
    padding: 16px 24px; margin-top: 24px;
  }
  .notice h3 { font-size: 14px; color: #633806; margin-bottom: 8px; }
  .notice p { font-size: 12.5px; color: #633806; margin-bottom: 6px; }
  .notice p:last-child { margin-bottom: 0; }
  .notice b { font-weight: 600; }

  @media (max-width: 760px) {
    .stage-body, .final-body, .parallel, .file-list { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>
<div class="page">

  <!-- ================= 头部 ================= -->
  <div class="hero">
    <h1>分销标准池定价 · 完整链路</h1>
    <div class="sub">技术公司视角 · 供货给所有客户 · 从本来生活全部分销商品底层数据中挖掘定价 · $DATE 本轮数据</div>
    <div class="oneline">
      一条主线：<b>$POOL 个四站全上架且有库存的分销商品</b>，逐个回 ERP 算成本，其中 <b>$REAL_RATE 配到真实京东对标价（统一走天镜智价）</b>、其余用商城基准价兜底，全量算出三档价并给建议档。
    </div>
    <div class="meta-row">
      <span class="meta-chip">起点 traceid <b>9ef0b69a…8fad</b></span>
      <span class="meta-chip">接口 <b>POST /Product/GetPager</b></span>
      <span class="meta-chip">商品池 <b>ProductPoolType=1（分销）</b></span>
      <span class="meta-chip">四站 <b>1 北京 · 2 广州 · 3 上海 · 7 武汉</b></span>
      <span class="meta-chip">全程<b>只读</b>，未写任何库</span>
    </div>
  </div>

  <!-- ================= 漏斗摘要 ================= -->
  <div class="funnel">
    <div class="fstep"><div class="fnum">$REL_TOTAL</div><div class="flabel">四站上架关系</div></div>
    <div class="farrow">→</div>
    <div class="fstep"><div class="fnum">$DEDUP</div><div class="flabel">去重商品</div></div>
    <div class="farrow">→</div>
    <div class="fstep"><div class="fnum">$FOUR</div><div class="flabel">四站全上架</div></div>
    <div class="farrow">→</div>
    <div class="fstep"><div class="fnum">$POOL</div><div class="flabel">四站都有库存</div></div>
    <div class="farrow">→</div>
    <div class="fstep"><div class="fnum">$COST_ROWS</div><div class="flabel">成本明细条数</div></div>
    <div class="farrow">→</div>
    <div class="fstep warn"><div class="fnum">$REAL_N</div><div class="flabel">真实京东对标</div></div>
    <div class="farrow">→</div>
    <div class="fstep final"><div class="fnum">$FINAL_ROWS</div><div class="flabel">定价成品行数</div></div>
  </div>

  <!-- ================= 阶段 0 ================= -->
  <h2 class="section-title">起点：traceid 背后的那次调用</h2>
  <div class="stage">
    <div class="stage-head">
      <div class="stage-no">0</div>
      <div class="stage-title">从 ARMS 调用链还原查询条件</div>
      <div class="stage-big"><div class="n">$REL_TOTAL</div><div class="t">四站上架关系（本轮）</div></div>
    </div>
    <div class="stage-body">
      <div class="kv"><div class="k">做了什么</div><div class="v">用 traceid 拉 ARMS 调用链，从 <code>http.request.body</code> 里还原出查询参数</div></div>
      <div class="kv"><div class="k">还原出的请求</div><div class="v"><code>{"Query":{…"WebSiteSysNoList":[1,2,3,7], "COnlineStatus":1…}, "ProductPoolType":1}</code></div></div>
      <div class="kv"><div class="k">关键发现</div><div class="v">站点编号 1/2/3/7 = 北京/广州/上海/武汉；<code>COnlineStatus=1</code> 即「上架」</div></div>
      <div class="kv"><div class="k">结论</div><div class="v">「所有分销商品」= 分销池 × 四站的上架关系，本轮实拉 $REL_TOTAL 条</div></div>
    </div>
  </div>

  <div class="connector"></div>

  <!-- ================= 阶段 1 ================= -->
  <h2 class="section-title">第一步：拉取四站上架关系</h2>
  <div class="stage">
    <div class="stage-head">
      <div class="stage-no">1</div>
      <div class="stage-title">分页拉全量商品×站点关系</div>
      <div class="stage-big"><div class="n">$REL_TOTAL</div><div class="t">条关系</div></div>
    </div>
    <div class="stage-body">
      <div class="kv"><div class="k">脚本</div><div class="v"><code>pull_pool_sites.py</code></div></div>
      <div class="kv"><div class="k">数据源 <span class="tag erp">ERP 接口</span></div><div class="v"><code>vm-online ProductWebSiteApi/GetProductPager</code>，每页 5000</div></div>
      <div class="kv"><div class="k">输入</div><div class="v">分销池(ProductPoolType=1) × 四站 × 上架(COnlineStatus=1)</div></div>
      <div class="kv"><div class="k">输出</div><div class="v"><code>$REL_FILE</code></div></div>
    </div>
  </div>

  <div class="connector"></div>

  <!-- ================= 阶段 2 ================= -->
  <h2 class="section-title">第二步：聚合出「四站全上架」商品</h2>
  <div class="stage">
    <div class="stage-head">
      <div class="stage-no">2</div>
      <div class="stage-title">一个商品在四站都有上架记录才算入选</div>
      <div class="stage-big"><div class="n">$FOUR</div><div class="t">个商品（从 $DEDUP 去重商品中筛出）</div></div>
    </div>
    <div class="stage-body">
      <div class="kv"><div class="k">脚本</div><div class="v"><code>aggregate_pool_candidates.py</code></div></div>
      <div class="kv"><div class="k">规则 <span class="tag user">用户拍板</span></div><div class="v">北京、上海、广州、武汉<b>四站都要上架</b>——缺一站就淘汰</div></div>
      <div class="kv"><div class="k">输入</div><div class="v">$REL_TOTAL 条关系</div></div>
      <div class="kv"><div class="k">输出</div><div class="v"><code>$FOUR_FILE</code>（含四站 basicPrice）</div></div>
    </div>
  </div>

  <div class="connector"></div>

  <!-- ================= 阶段 3 ================= -->
  <h2 class="section-title">第三步：库存筛选</h2>
  <div class="stage">
    <div class="stage-head">
      <div class="stage-no">3</div>
      <div class="stage-title">四站都必须「有货」</div>
      <div class="stage-big"><div class="n">$POOL</div><div class="t">个商品 · 淘汰 $FOUR_MINUS_POOL 个</div></div>
    </div>
    <div class="stage-body">
      <div class="kv"><div class="k">脚本</div><div class="v"><code>filter_stock.py</code>（分批 1000 个/批）</div></div>
      <div class="kv"><div class="k">数据源 <span class="tag db">生产库</span></div><div class="v">archery 只读 · SQL Server 实例256 · <code>OpenAPI_Product.dbo.Product_Inventory</code></div></div>
      <div class="kv"><div class="k">规则 <span class="tag user">用户拍板</span></div><div class="v">每站 <code>SaleQty&gt;0</code> 且 Status=1 未删除，四站都有库存记录</div></div>
      <div class="kv"><div class="k">输出</div><div class="v"><code>$POOL_FILE</code></div></div>
    </div>
  </div>

  <div class="connector"></div>

  <!-- ================= 阶段 4/5 并行 ================= -->
  <h2 class="section-title">第四步：两路并行补数据（都围绕 $POOL 个商品）</h2>
  <div class="parallel">
    <div class="stage">
      <div class="stage-head">
        <div class="stage-no">4a</div>
        <div class="stage-title">回 ERP 取成本明细</div>
        <div class="stage-big"><div class="n">$COST_ROWS</div><div class="t">条明细</div></div>
      </div>
      <div class="stage-body" style="grid-template-columns:1fr;">
        <div class="kv"><div class="k">脚本 / 数据源</div><div class="v"><code>pull_standard_pool_costs.py</code> <span class="tag erp">ERP 接口</span> <code>ProLhotse/GetListByBasic</code>，每批 200 个 P 码</div></div>
        <div class="kv"><div class="k">输出</div><div class="v"><code>$COST_FILE</code>（每商品多站点多条）</div></div>
        <div class="rule-box"><b>成本口径（沿用禄品池拍板规则）：</b>先「上架展示」后「有效未展示」回退，作废/删除全程排除；<b>直发取采购价、仓发取含税单位成本（无则采购价）</b>，均取单站最高。发货类型判断：成本明细中 <code>OriginDirectDelivery=True</code> → 直发（用禄品池反推标定，直发 99% / 仓发 100% 吻合；IsDirectSend 字段不可靠）。</div>
      </div>
    </div>
    <div class="stage">
      <div class="stage-head">
        <div class="stage-no">4b</div>
        <div class="stage-title">补商品名称与类目</div>
        <div class="stage-big"><div class="n">$POOL</div><div class="t">个全查到</div></div>
      </div>
      <div class="stage-body" style="grid-template-columns:1fr;">
        <div class="kv"><div class="k">脚本 / 数据源</div><div class="v"><code>pull_pool_product_info.py</code> <span class="tag db">生产库</span> archery · MySQL 实例147 · <code>benlai_business_product.product</code></div></div>
        <div class="kv"><div class="k">输出</div><div class="v"><code>源数据_标准池商品信息_$DATE.jsonl</code>（名称 + 一/二/三级类目编码）</div></div>
        <div class="rule-box"><b>为什么需要：</b>相似度匹配要拿商品名称跟京东比，类目编码用于控制比对范围。名称是匹配质量的地基，这步不可省。</div>
      </div>
    </div>
  </div>
  <div class="merge"></div>
  <div class="connector"></div>

  <!-- ================= 阶段 5 ================= -->
  <h2 class="section-title">第五步：同品关系确认 + 天镜智价取对标价</h2>
  <div class="stage">
    <div class="stage-head">
      <div class="stage-no">5</div>
      <div class="stage-title">两条路确认同品，对标价统一走天镜智价</div>
      <div class="stage-big"><div class="n">$REAL_N</div><div class="t">个真实对标 · 占比 $REAL_RATE</div></div>
    </div>
    <div class="stage-body">
      <div class="kv"><div class="k">脚本</div><div class="v"><code>match_pool_vop.py</code>（复用禄品池的 <code>score_pair</code> 评分体系）</div></div>
      <div class="kv"><div class="k">数据源 <span class="tag db">天镜智价</span><span class="tag erp">京东 VOP</span></div><div class="v">天镜历史价库（每 SKU 最新一条）+ 天镜人工核实映射；VOP 商品池仅作匹配候选池与备选价</div></div>
      <div class="kv"><div class="k">同品确认（两路，人工核实优先）</div><div class="v">① 天镜人工核实映射（同款/不同规/相似，按 P 码直接对应）；② 算法相似度 ≥0.75 自动采信</div></div>
      <div class="kv"><div class="k">对标价取值规则</div><div class="v"><b>统一走天镜智价</b>：同款映射 → 天镜历史价；相似采信 → 天镜历史价，天镜无价才用 VOP jd_price</div></div>
      <div class="kv"><div class="k">输出</div><div class="v"><code>$MATCH_FILE</code>（0.6 以上全记录，$CANDIDATES 个候选未采信备用）</div></div>
      <div class="kv"><div class="k">真实对标来源细分</div><div class="v">$REAL_DETAIL</div></div>
      <div class="rule-box"><b>不同规格怎么折算（用户拍板）：</b>天镜比例 a:b = 京东 a 件 : 本来生活 b 件；<b>单件价 = 京东价 ÷ a；本来生活件数直接从商品名解析</b>（与禄品池同一套规格解析），对标价 = 单件价 × 件数。例：挂面 900g×3（比例 1:3）→ 京东单价 × 3。天镜映射有但价库无价的 $DIFF_SPEC 个暂转兜底；「相似」关系按禄品口径仅参考不给价。</div>
      <div class="rule-box"><b>没匹配上的主因：</b>本来生活自有品牌（荷裕、西鲜记、源禾种意等约四成）京东池里根本没有。</div>
    </div>
  </div>

  <div class="connector"></div>

  <!-- ================= 阶段 6：成品 ================= -->
  <h2 class="section-title">第六步：定价成品</h2>
  <div class="final-card">
    <div class="final-head">
      <div class="n">$FINAL_ROWS 行 × 27 列</div>
      <div class="file">$XLSX</div>
    </div>
    <div class="final-body">
      <div class="stat-chip hi"><div class="n">$COST_COVER</div><div class="t">有成本（成本缺失仅 $COST_MISS 个）</div></div>
      <div class="stat-chip"><div class="n">$REAL_N / $ERP_N</div><div class="t">对标价来源：真实京东对标（天镜智价为主）/ ERP基准价兜底（已标注非京东价）</div></div>
      <div class="stat-chip"><div class="n">$SHIP_DIRECT_PCT 直发</div><div class="t">直发 $SHIP_DIRECT（加价 2/5/8%）· 仓发 $SHIP_WARE（加价 5/12/24%）</div></div>
      <div class="stat-chip hi"><div class="n">$CHK_ERR 错误</div><div class="t">输出价公式校验 $CHK_CELLS 格</div></div>
      <div class="stat-chip warn"><div class="n">$MED_JD</div><div class="t">京东采信商品建议档空间率中位数（10% 为渠道底线）</div></div>
      <div class="stat-chip"><div class="n">档②为主</div><div class="t">建议档：真实对标 档①$SUG_JD1 / 档②$SUG_JD2 / 档③$SUG_JD3；兜底 档①$SUG_ERP1 / 档②$SUG_ERP2 / 档③$SUG_ERP3</div></div>
    </div>
  </div>

  <!-- ================= 建议档规则 ================= -->
  <h2 class="section-title">建议档怎么定（用户拍板规则）</h2>
  <div class="stage">
    <div class="stage-body" style="padding-top:18px;">
      <div class="rule-box"><b>默认中间档②</b>——不紧不松，作为标准供货价。</div>
      <div class="rule-box"><b>档②空间率 &lt;10% → 降档①</b>（低利润，让渠道有空间）。本轮真实对标中 $SUG_JD1 个触发。</div>
      <div class="rule-box"><b>档②空间率 ≥30% 且档③仍 ≥10% → 升档③</b>（毛利空间巨大才定高价）。</div>
      <div class="rule-box"><b>兜底价商品不升档</b>——ERP 商城零售价不是市场价，空间率失真（本轮中位 $MED_ERP vs 京东真实 $MED_JD），只降不升、保守中档。</div>
    </div>
  </div>

  <!-- ================= 拍板决策一览 ================= -->
  <h2 class="section-title">拍板决策一览（规则稳定，每周沿用）</h2>
  <div class="table-wrap">
    <table>
      <tr><th style="width:24%;">决策点</th><th style="width:40%;">拍板内容</th><th style="width:36%;">落地位置</th></tr>
      <tr><td>池子范围</td><td>分销池 × 北京/上海/广州/武汉四站全上架 × 四站都有库存</td><td>阶段 1–3 的筛选条件</td></tr>
      <tr><td>对标价来源</td><td>统一走天镜智价：人工核实同款映射优先，其次相似采信；天镜无价才用 VOP jd_price 备选</td><td>阶段 5 取值规则</td></tr>
      <tr><td>同品采信</td><td>相似度 ≥75 分直接算一个品，不再等人工核实</td><td>阶段 5 匹配阈值</td></tr>
      <tr><td>建议档取向</td><td>默认中间档；空间不足降下档；毛利空间巨大（≥30%）才升档</td><td>阶段 6 建议档规则</td></tr>
      <tr><td>交付形态</td><td>母池全量给价、标建议档，不做「选/不选」筛选</td><td>成品结构（无选品列）</td></tr>
      <tr><td>兜底保守</td><td>ERP 基准价兜底商品不升档（空间率失真）</td><td>阶段 6 建议档规则</td></tr>
      <tr><td>匹配阈值维持</td><td>0.6–0.75 的候选暂不采信，留匹配文件备用</td><td>阶段 5 输出文件</td></tr>
    </table>
  </div>

  <!-- ================= 文件清单 ================= -->
  <h2 class="section-title">产出文件清单（pricing_output/，文件名带日期、每周滚动）</h2>
  <div class="file-list">
    <div class="file-item"><div class="fn">$XLSX</div><div class="fd">最终成品 · $FINAL_ROWS 行 · 01_定价结果 + 02_统计汇总</div></div>
    <div class="file-item"><div class="fn">$REL_FILE</div><div class="fd">阶段1 · $REL_TOTAL 条四站上架关系</div></div>
    <div class="file-item"><div class="fn">$FOUR_FILE</div><div class="fd">阶段2 · $FOUR 个四站全上架商品</div></div>
    <div class="file-item"><div class="fn">$POOL_FILE</div><div class="fd">阶段3 · $POOL 个四站全上架有库存商品</div></div>
    <div class="file-item"><div class="fn">$COST_FILE</div><div class="fd">阶段4a · $COST_ROWS 条成本明细</div></div>
    <div class="file-item"><div class="fn">源数据_标准池商品信息_$DATE.jsonl</div><div class="fd">阶段4b · 名称+类目</div></div>
    <div class="file-item"><div class="fn">$MATCH_FILE</div><div class="fd">阶段5 · 0.6 分以上全记录（采信+候选）</div></div>
    <div class="file-item"><div class="fn">定价统计_$DATE.json</div><div class="fd">成品统计 · 报告数据源</div></div>
    <div class="file-item"><div class="fn">pull → aggregate → filter → pull×2 → match → build → report</div><div class="fd">8 个脚本，顺序执行即可复跑刷新</div></div>
  </div>

  <!-- ================= 观察点 ================= -->
  <div class="notice">
    <h3>两个需要持续观察的点</h3>
      <p><b>1. 兜底占比。</b>本轮 $REAL_RATE 的商品有真实京东市场价（对标价统一来自天镜智价），其余靠商城价兜底——这是标准池最大的可信度边界。对外报价优先用有真实对标的商品；兜底商品定价建议实际询价后再校一轮。每周复跑后看真实对标率变化。</p>
      <p><b>2. 真实市场下空间是否偏薄。</b>真实对标商品建议档空间率中位数本轮 $MED_JD（底线 10%）。若持续贴近底线，说明加价体系在市场压力下空间紧张，是谈价和调档的依据。</p>
  </div>

</div>
</body>
</html>
""")


def main():
    v = collect()
    # 阶段3淘汰数
    v["FOUR_MINUS_POOL"] = fmt(int(v["FOUR"].replace(",", "")) - int(v["POOL"].replace(",", "")))
    html = TPL.substitute(**v)
    OUT.write_text(html, encoding="utf-8")
    print(f"产出: {OUT.name}")
    print(f"  关系 {v['REL_TOTAL']} → 去重 {v['DEDUP']} → 四站全上架 {v['FOUR']} → 有库存 {v['POOL']}")
    print(f"  成本明细 {v['COST_ROWS']} | 采信 {v['MATCHED']} ({v['MATCH_RATE']}) | 成品 {v['FINAL_ROWS']} 行")


if __name__ == "__main__":
    main()
