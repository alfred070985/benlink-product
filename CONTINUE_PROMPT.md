# benlink-product 跨环境续作指令

> 远程仓库 `git@github.com:alfred070985/benlink-product.git` 已创建并同步。把下面整段发给另一台电脑或另一个具备文件、Git 和终端能力的代理，即可恢复项目。

把下面整段发给另一台电脑或另一个具备文件、Git 和终端能力的代理：

```text
你将成为 benlink-product 在当前设备、当前应用载体上的续作代理。不要假设你能看到其他设备或应用的历史会话；以 Git 仓库为唯一事实源。

项目 ID：benlink-product
项目仓库：git@github.com:alfred070985/benlink-product.git
续作分支：main

请严格执行以下恢复协议：

0. 跨设备续作范式：本项目遵循 `ai-workspace` 控制仓库的 `MULTI_MACHINE_RULES.md`（R1–R12）。clone 本仓库前先 clone 控制仓库 `ai-workspace/control` 打底；脚本已使用仓库相对路径、源数据已入仓（LFS）；新交付物创建即登记到 `00_项目汇总/交付文件清单.csv`；文档随状态实时更新；仓库认证用 SSH 密钥或 PAT（与数据分离）；多机并行用 merge 且清单按路径幂等。

1. 工作范围仅限 benlink-product 项目根目录，不得遍历父目录、兄弟项目或整盘寻找上下文。
2. 如果本机已经登记该项目，使用登记路径；否则克隆到当前系统用户目录下的 ai-workspace/projects/benlink-product。若仓库不存在或无权访问，停止并让我先创建私有仓库或完成授权。
3. 如果目录已存在，先核对 Git 远程地址。存在未提交改动、远程地址不符或项目身份不符时，停止并报告；不得 reset --hard、clean、覆盖或擅自暂存改动。
4. 安全获取远程更新，切到 main，执行等效于 git pull --ff-only 的非破坏性更新。项目使用 Git LFS，确认 Git LFS 可用后执行 git lfs pull；如果只有 LFS 指针而没有实际 Excel/JSONL 内容，不得继续数据分析。
5. 首先只读取 START_HERE.md、AGENTS.md、.ai/project.yaml、.ai/state.md。核对 project_id 必须为 benlink-product，然后按入口指引阅读项目汇总和必要决策，不要默认遍历全部文件。
6. 运行 python3 .ai/verify.py。该检查只离线验证项目身份、必要文件、脚本语法、LFS 和环境，不连接生产数据。
7. 当前续作需要外部 benlai-trace-debug 能力包。不要复制它，也不要复制或提交 config.local.yaml、凭证或 .tools/。优先让我提供当前电脑的能力包路径，然后设置 BENLAI_ARCHERY_DIR，或把 ~/ai-workspace/shared/benlai-trace-debug 软链接到真实能力包根目录。无法确认路径时只问我这一项。
8. 用不超过六行告诉我：项目、分支、当前 HEAD 提交、已完成、唯一下一步、阻塞。随后按 AGENTS.md 和 .ai/state.md 继续，不要重新规划整个项目。
9. 线上查询只能通过 Archery 只读通道。查询前运行 python3 .ai/verify.py --require-archery；登录失败、超时或权限不足时记录真实阻塞并停止，不紧密重试，不直连生产数据库。
10. 95012 商品清单已核准（4940 在售，零差集），新版报价表已生成并通过 QA。重新生成时不得拿现有分销商品底表直接改名；新版报价表不得包含上海履约成本字段，也不得沿用"来福商品池"命名。
11. 每次暂停或完成阶段前，更新 .ai/state.md 和必要的长期决策／交付清单，运行真实验证，提交并推送当前分支。未成功推送必须明确说“尚未同步”。
12. 结束回复时给出实际 HEAD 提交号，并更新这份续作指令的分支（如有变化），使我可以直接转到下一台设备或另一个应用。

确认恢复完成后，直接继续 .ai/state.md 中的下一步。
```
