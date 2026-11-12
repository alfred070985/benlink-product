# benlink-product：从这里继续

- **项目 ID**：`benlink-product`
- **当前阶段**：核准开放平台 `95012（禄品供应商模式）` 的当前商品清单，然后生成新版报价表。
- **当前阻塞**：商品范围尚未确认；上次最终核查时 Archery 登录超时。

新环境按顺序读取，不要先遍历整个目录：

1. `AGENTS.md`：项目边界、业务口径、环境和交接规则。
2. `.ai/project.yaml`：项目身份、建议远程仓库和外部能力依赖。
3. `.ai/state.md`：真实进度、唯一下一步、阻塞和验证状态。
4. `00_项目汇总/benlink-product_交付汇总.md`：完整业务口径和历史交付关系。
5. `.ai/decisions/`：后续新增的长期决策。

## 现在要做的事

通过 Archery 只读通道继续核准 `95012` 商品范围。未取得并核验当前商品清单前，不得把现有 43,519 个分销商品直接改名生成 `95012` 报价表。

## 离线自检

```bash
python3 .ai/verify.py
```

需要实际执行 Archery 查询前，再检查外部能力包：

```bash
python3 .ai/verify.py --require-archery
```

跨设备交接以 `CONTINUE_PROMPT.md` 为启动指令。若远程仓库尚未创建或推送，先完成首次 Git 设置，不能声称已经同步。
