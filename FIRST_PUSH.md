# benlink-product 首次 Git 同步

当前副本尚未初始化 Git，建议远程仓库也尚未创建或验证。以下步骤需要由项目所有者在确认后执行。

## 前置条件

1. 在 GitHub 创建**私有**仓库：`alfred070985/benlink-product`。
2. 当前电脑完成 GitHub 认证，并安装 Git LFS。
3. 确认项目目录中的业务数据允许进入该私有仓库；不要加入 `.env`、`config.local.yaml` 或其他凭证。
4. 在项目根目录执行命令，先检查 `git status --short`，不要覆盖任何已有改动。

## 初始化与推送

```bash
git init -b main
git lfs install
git lfs track
git status --short
```

确认 `.gitattributes` 显示 `*.xlsx` 和 `*.jsonl` 已由 LFS 跟踪，并检查暂存范围后：

```bash
git add .
git status --short
git commit -m "chore: initialize benlink-product continuation workspace"
git remote add origin git@github.com:alfred070985/benlink-product.git
git push -u origin main
```

如果远程地址已经存在，不要重复添加；先核对 `git remote -v`。不要强制推送。

## 推送成功后

1. 将 `.ai/state.md` 的 `git_remote` 改为 `active`，保留真实的业务阻塞。
2. 删除 `CONTINUE_PROMPT.md` 顶部“尚未创建或验证”的提示。
3. 提交并推送这次状态更新。
4. 在另一台电脑用 `CONTINUE_PROMPT.md` 做首次恢复演练，确认 Excel／JSONL 的实际 LFS 内容已下载，而不只是指针文件。

未完成上述步骤前，项目只是“已具备续作文件结构”，不是“已经可跨设备同步”。
