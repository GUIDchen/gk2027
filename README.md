# 2027 国考岗位工具（gk2027）

一个可复用的 **2027 国考岗位筛选 + 报名人数统计** 工具：

- 按**自己的条件**筛岗位（18—38 岁、党员/团员、学历学位、专业、基层年限、英语、户籍、应届身份、证书等）；
- **一键下载**官方公告/大纲/政策/问答/附件/岗位表；
- 在**自己登录的 Chrome** 里抓取报名人数；
- 把每天报名人数并列写入 Excel，并生成统计表；
- 适合自己用，也适合别人改条件后直接用。

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

## 快速开始

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

---

## 文档导航

| 文档 | 内容 |
|---|---|
| [01 快速开始](docs/01_快速开始.md) | 安装、第一次流程、命令速查 |
| [02 个人使用](docs/02_个人使用.md) | 本项目作者的条件、1,034 个结果、10-22/10-23 数据 |
| [03 他人使用](docs/03_他人使用.md) | 别人如何用 `init` 改成自己的条件、已验证示例 |
| [04 文档下载](docs/04_文档下载.md) | `download` 下载哪些、放在哪、参数 |
| [05 数据与Excel说明](docs/05_数据与Excel说明.md) | 工作表、列、快照、counts JSON、颜色 |
| [06 抓取报名人数](docs/06_抓取报名人数.md) | Chrome 准备、抓取原理、401/并发/超时 |
| [07 常见问题](docs/07_常见问题.md) | 14 个常见问题 |
| [08 项目结构](docs/08_项目结构.md) | 目录、模块、数据流、git 忽略 |
| [00 工具总览](docs/00_工具总览.md) | 早期完整版说明（可作为参考） |

---

## 项目结构（简）

```
gk2027.py                  # 入口
run.command / setup.sh     # 启动/安装
gk2027_lib/                # matcher/downloader/build/fetch/snapshot/report/cli/util
docs/                      # 分主题文档
gk2027_config.example.json # 脱敏示例配置
2027国考资料/              # 官方资料 + 生成结果（git 忽略）
data/                      # 抓取原始 JSON（git 忽略）
```

---

## 数据与隐私

- `gk2027_config.json`（个人条件）、`考生基本信息.png`、`2027国考资料/`、`data/` 均被 `.gitignore` 排除；
- 仓库只保留代码、文档和脱敏示例；
- 官方资料用 `download` 重新下载，报名人数用 `fetch` 重新抓取；
- 抓取需要**自己的** Chrome 登录，工具不共享账号。

---

## 免责

绿色/黄色是辅助筛选，不替代官方资格审查；黄色岗位请拨打职位表内咨询电话确认专业、学位、英语、体能等条件。报名人数会随报名进程变化，最终以系统实时数据为准。
