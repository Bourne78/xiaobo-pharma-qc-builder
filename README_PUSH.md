# 推送本技能到 GitHub

本目录（解压后）即一个可直接放到仓库根目录的 WorkBuddy 技能包。
以下步骤把 `xiaobo-pharma-qc-builder` 技能推到你建的 GitHub 仓库
`Bourne78/xiaobo-pharma-qc-builder`。

> ⚠️ 仓库当前为 **public（公开）**。技能内虽已脱敏（无真实姓名 / 内网 IP / 内网域名），
> 但仍含真实产品名（聚多卡醇注射液）与文件编号（STP-SP-036-5001）等内部 QC 信息。
> 若需保密，请先在 GitHub 网页将仓库改为 Private，再执行下面的推送。

## 方式一：仓库已建好（空仓，或只有 README）

在**能连 GitHub 的机器**上，进入本目录后执行：

```bash
# 1. 初始化并提交
git init
git add -A
git commit -m "feat: xiaobo-pharma-qc-builder skill v2"

# 2. 关联远程（替换成你的仓库地址）
git remote add origin https://github.com/Bourne78/xiaobo-pharma-qc-builder.git

# 3. 推送
git branch -M main
git push -u origin main
```

若建仓时顺手勾了 README / .gitignore 初始化，远程已有一个 `main` 分支，
先拉取再推（避免 rejected）：

```bash
git pull --rebase origin main
git push -u origin main
# 如遇冲突，按提示解决后 git rebase --continue 再 push
```

## 方式二：想让技能位于仓库子目录（可选）

上面方式会把技能直接放在仓库根（`SKILL.md` / `scripts/` / `references/` 同级）。
若你更希望它位于 `xiaobo-pharma-qc-builder/` 子目录，把本目录内容整体移到该子目录再提交即可。

## 验证推送成功

```bash
git ls-files
# 应能看到 SKILL.md、scripts/*（含 patch_existing_manual.py / verify_revision.py）、
# scripts/export_pdf.ps1、references/*（含 revision_playbook.md）等全部文件
```

## 在 WorkBuddy 中使用

技能装好后，对话里用下面任一方式调用：

- 显式点名：`@skill:xiaobo-pharma-qc-builder 把这份 STP/SOP 整理成检验手册`
- 自然语言：直接说"帮我把这批 SOP 和验证报告梳理成一份检验准备手册"
- 触发词：检验手册 / 检验物料清单 / 备料核对清单 / SOP 嵌入手册 / 偏差案例挂接 / 关键点框 /
  页眉页脚 / 成稿修订 / 修订升版 / 列宽重算 / 批号回填 / 供应商回填 / 矩阵全名 /
  版本沿革 / 增量说明 / 零回归验收
