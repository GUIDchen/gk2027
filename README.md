# 2027 国考岗位工具（gk2027）

> 仓库分为两部分：**开源代码/脚本/文档**（本仓库，可分享） + **个人使用目录**（本地，不进入 git）。

| 目录 | 内容 | 是否提交 git |
|---|---|---|
| 根目录 | `gk2027.py`、`gk2027_lib/`、`docs/`、示例配置、启动/安装脚本 | ✅ 提交 |
| `个人使用/` | 个人条件、下载的官方资料、生成的 Excel、抓取数据、个人记录 | ❌ 忽略 |

---

## 功能

| 命令 | 作用 |
|---|---|
| `./run.command init` | 设置本人条件（别人用自己的条件筛选） |
| `./run.command download` | 下载公告/大纲/政策/问答/附件/岗位表 |
| `./run.command build` | 生成带颜色/新 sheet 的筛选工作簿 |
| `./run.command fetch` | 从已登录 Chrome 抓取报名人数 |
| `./run.command snapshot` | 把某天人数合并到 Excel，与上一天并列 |
| `./run.command report` | 生成独立统计表 |
| `./run.command all` | 一键：抓取 → 合并 → 统计 |
| `./run.command doctor` | 检查依赖、文件、Chrome 登录/页面脚本 |
| `./run.command guide` | 使用说明 |
| `./run.command` | 交互菜单 |

---

## 开源用户快速开始

```bash
cd dsh-26gongkao
./setup.sh              # 首次安装依赖
./run.command init      # 填自己的条件
./run.command download  # 下载官方资料/岗位表
./run.command build     # 生成筛选表

# 登录 Chrome，打开“个人中心 → 职位查询 → 按代码选择职位”
# 显示 → 开发者 → 勾选“允许 Apple 事件中的 JavaScript”
./run.command fetch --label 10-23
./run.command snapshot --label 10-23 --counts data/counts/counts_2026-10-23.json
./run.command report
```

> 开源用户的数据默认生成在根目录的 `2027国考资料/`、`data/`，已被 `.gitignore` 排除。

---

## 个人使用（本地）

本机另有一个 `个人使用/` 目录（不提交 git），里面有：

- `个人使用/gk2027_config.json`：个人条件；
- `个人使用/2027国考资料/`：官方资料 + 筛选 Excel + 统计表；
- `个人使用/data/counts/`：每天抓取的原始 JSON；
- `个人使用/docs/个人记录.md`：个人结果；
- `个人使用/考生基本信息.png`：含敏感信息，请勿分享。

工具会自动检测：根目录没有 `gk2027_config.json` 时，读取 `个人使用/gk2027_config.json`。
个人使用的命令与上面相同，只是输出都在 `个人使用/` 下。

详见 `个人使用/README.md`。

---

## 文档导航（开源部分，无个人信息）

| 文档 | 内容 |
|---|---|
| [01 快速开始](docs/01_快速开始.md) | 安装、第一次流程、命令速查 |
| [02 他人使用](docs/02_他人使用.md) | 别人如何用 `init` 改成自己的条件 |
| [03 文档下载](docs/03_文档下载.md) | `download` 下载哪些、放在哪、参数 |
| [04 数据与Excel说明](docs/04_数据与Excel说明.md) | 工作表、列、快照、counts JSON、颜色 |
| [05 抓取报名人数](docs/05_抓取报名人数.md) | Chrome 准备、抓取原理、401/并发/超时 |
| [06 常见问题](docs/06_常见问题.md) | 14 个常见问题 |
| [07 项目结构](docs/07_项目结构.md) | 目录、模块、数据流、git 忽略 |
| [00 工具总览](docs/00_工具总览.md) | 完整版说明（参考） |

---

## Git 与隐私

`.gitignore` 已排除：

- `.venv/`、`__pycache__/`、`.DS_Store`；
- `个人使用/`（个人配置/数据/结果）；
- 根目录的 `gk2027_config.json`、`考生基本信息.png`；
- `2027国考资料/`、`data/`（可重新 download/fetch）。

仓库只保留代码、脚本、文档和 `gk2027_config.example.json`。

---

## 免责

绿色/黄色是辅助筛选，不替代官方资格审查；黄色岗位请拨打职位表内咨询电话确认专业、学位、英语、体能等条件。报名人数会随报名进程变化，最终以系统实时数据为准。
