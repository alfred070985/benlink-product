# benlink-product 首次 Git 同步

本地仓库已在 `main` 分支形成首次提交 `069b6af6da3871be368892719f469d7b65004a8d`，并已核验 14 个 `*.xlsx`／`*.jsonl` 文件进入 Git LFS。建议远程仓库 `git@github.com:alfred070985/benlink-product.git` 尚未创建或验证，因此当前状态是**本地已提交、尚未同步**。

## 推送前置条件

1. 在 GitHub 创建**私有**仓库：`alfred070985/benlink-product`。
2. 当前电脑具有该私有仓库的推送权限，并完成 GitHub SSH 认证。
3. 当前电脑可运行 Git LFS。系统没有全局安装时，可使用本机已有的便携 Git LFS，但不要把这个本机工具路径提交进项目。
4. 确认项目目录中的业务数据允许进入该私有仓库；不要加入 `.env`、`config.local.yaml` 或其他凭证。

## 推送命令

在项目根目录执行：

```bash
git status --short --branch
git remote -v
```

确认没有本地改动，且 `origin` 不存在或为正确地址后：

```bash
git remote add origin git@github.com:alfred070985/benlink-product.git
git push -u origin main
```

如果 `origin` 已经存在但地址不同，先停止并核对，不要重复添加或强制推送。

## 推送后核验

```bash
git ls-remote origin main
git lfs ls-files | wc -l
git lfs ls-files
```

`git ls-files` 一侧应显示 14 个 LFS 文件。随后：

1. 将 `.ai/state.md` 的 `git_remote` 改为 `active`，保留真实的业务阻塞。
2. 删除 `CONTINUE_PROMPT.md` 顶部“尚未创建或验证”的提示。
3. 提交并推送这次状态更新。
4. 在另一台电脑用 `CONTINUE_PROMPT.md` 做首次恢复演练，确认 Excel／JSONL 的实际 LFS 内容已下载，而不只是指针文件。

未完成远程创建和推送前，项目不能声称已经可跨设备同步。
