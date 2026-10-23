# gk2027 · 2027 国考岗位筛选与报名人数统计

> 把 1.9 万个岗位，变成一份可执行的报考清单。

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-macOS-lightgrey)
![Chrome](https://img.shields.io/badge/Chrome-AppleScript-yellow)

`gk2027` 是一个本地运行的 2027 国考岗位工具：

- 按**自己的条件**筛岗位：学历/学位/政治面貌/专业/基层年限/应届身份/英语/户籍/证书/性别等；
- **一键下载**官方公告、大纲、政策、问答、附件和岗位表；
- 在**自己登录的 Chrome** 里抓取报名人数；
- 每天把报名人数并列写入 Excel，自动算“较昨日增加”；
- 生成按人数升序、报录比升序、低竞争筛选的统计表；
- 输出保留官方原表格式：颜色标记 + 新工作表，不改原 4 张表。

---

## 为什么做这个

国考岗位表有近 2 万个职位、35,000+ 人，条件散落在 27 列和一大段备注里：

- “本科及以上 / 仅限硕士 / 与最高学历相对应”要逐条对照；
- 专业代码、学术学位/专业学位、本科阶段/研究生阶段要分开看；
- 应届、基层年限、服务基层项目、英语六级、户籍、证书、性别……任何一条不符都报不了；
- 报名人数每天变化，手工抄表很容易错。

`gk2027` 把这些重复劳动做成可重复执行的脚本，并且把“为什么不符合”保留在表里，方便复核。

---

## 功能

| 命令 | 作用 |
|---|---|
| `./run.command init` | 交互式设置本人条件（别人用自己的条件筛选） |
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

## 工作流

```text
官方专题网站
   │  download
   ▼
2027国考资料/05_岗位表/*.xls
   │  init（本人条件）
   ▼
gk2027_config.json
   │  build
   ▼
06_个人筛选/...筛选.xls（颜色 + 符合条件岗位 + 地区表）
   │  fetch（Chrome 登录态）
   ▼
data/counts/counts_日期.json
   │  snapshot
   ▼
筛选.xls 右侧追加：报名人数 / 报录比 / 较昨日增加 / 更新时间
   │  report
   ▼
2027国考_符合条件岗位_报名人数统计.xls（总表 + 汇总 + 说明）
```

---

## 快速开始

### 环境

- macOS（`fetch` 需要 Chrome + AppleScript；筛选/下载可在 macOS 上跑）
- Python 3.9+
- Google Chrome

### 安装

```bash
git clone https://github.com/GUIDchen/gk2027.git
cd gk2027
./setup.sh
```

`setup.sh` 会创建 `.venv`、安装 `xlrd / xlwt / xlutils`，然后做一次环境检查。

### 第一次

```bash
# 1) 填自己的条件
./run.command init

# 2) 下载官方资料和岗位表
./run.command download

# 3) 生成筛选表
./run.command build

# 4) 抓取报名人数（需要先登录 Chrome）
#    Chrome 打开 https://bm.scs.gov.cn/kl2027
#    个人中心 → 职位查询 → 按代码选择职位，保持页面打开
#    菜单：显示 → 开发者 → 勾选“允许 Apple 事件中的 JavaScript”
./run.command fetch --label 10-23

# 5) 合并 + 统计
./run.command snapshot --label 10-23 --counts data/counts/counts_2026-10-23.json
./run.command report
```

### 每天更新

```bash
./run.command fetch --label 10-24
./run.command snapshot --label 10-24 --counts data/counts/counts_2026-10-24.json
./run.command report
```

---

## 筛选逻辑

| 维度 | 处理方式 |
|---|---|
| 学历 | 大专/本科/硕士/博士分别匹配“仅限/及以上/或” |
| 学位 | 无/学士/硕士/博士，支持“与最高学历相对应” |
| 政治面貌 | 党员/团员/群众/民主党派 |
| 基层年限 | 无限制/1/2/3/5 年，按本人年限匹配 |
| 服务基层项目 | 无限制/大学生村官/三支一扶/西部计划/特岗/退役士兵 |
| 专业 | 专业代码 + 名称 + 关键词；区分研究生/本科阶段；区分学术/专业学位 |
| 年龄 | 18—38 周岁；应届硕士/博士放宽到 43；公安/特警/监狱等单独限制 |
| 应届 | 2027届/2026届/2025届/往届；识别“限应届”“非应届须2年工作经历” |
| 英语 | CET-4/CET-6 分数；识别六级、专八、雅思、托福等 |
| 户籍 | 省/市/县，识别“限本省/本市/本县/生源” |
| 证书 | 法律职业资格、CPA、税务师、软考等 |
| 性别/其他 | 限女性/男性、退役军人、残疾人、少数民族、定向、运动员等 |

输出分三档：

- **绿色**：明确符合，建议优先；
- **黄色**：基本符合但需电话确认（学术学位/专业学位、专业大类、学位学士、英语熟练、体能测评等）；
- **红色**：不符合，保留原因，便于复核。

---

## 下载内容

`download` 会把官方专题网站的资料下载到：

```
2027国考资料/
├── 01_招考公告与考试大纲/   # 公告、报考指南、公共/专业/非通用语/公安大纲（HTML + PDF）
├── 02_政策法规/             # 公务员法、录用规定、体检标准、考察办法等
├── 03_常见问题/             # 实名认证等官方问答
├── 04_相关下载/             # 招考简章 zip、报名表/推荐表、证监会大纲、最新通知附件
└── 05_岗位表/               # 从 zip 解压出的官方 .xls
```

支持 `--skip-pdf`（只存 HTML）、`--only position/articles/resources`、`--force`。网站新发布的“报名人数较少职位”等通知也会自动收录。

---

## 项目结构

```
gk2027/
├── gk2027.py                 # 命令行入口
├── gk2027_lib/
│   ├── cli.py                # 菜单/参数/init 向导
│   ├── matcher.py            # 绿/黄/红规则
│   ├── downloader.py         # 官方资料下载
│   ├── excel_build.py        # 生成筛选工作簿
│   ├── fetch.py              # Chrome 抓取报名人数
│   ├── excel_snapshot.py     # 合并每日快照
│   ├── stats_report.py       # 统计表
│   └── util.py               # 配置/路径/AppleScript
├── docs/                     # 分主题文档
├── data/                     # 抓取原始 JSON（git 忽略）
├── 2027国考资料/             # 下载/生成结果（git 忽略）
└── gk2027_config.example.json
```

---

## 文档

| 文档 | 内容 |
|---|---|
| [01 快速开始](docs/01_快速开始.md) | 安装、第一次流程、命令速查 |
| [02 他人使用](docs/02_他人使用.md) | 别人如何用 `init` 改成自己的条件 |
| [03 文档下载](docs/03_文档下载.md) | `download` 下载哪些、放在哪、参数 |
| [04 数据与Excel说明](docs/04_数据与Excel说明.md) | 工作表、列、快照、counts JSON、颜色 |
| [05 抓取报名人数](docs/05_抓取报名人数.md) | Chrome 准备、抓取原理、401/并发/超时 |
| [06 常见问题](docs/06_常见问题.md) | 常见问题 |
| [07 项目结构](docs/07_项目结构.md) | 目录、模块、数据流、git 忽略 |

---

## 隐私

- `gk2027_config.json`（个人条件）、`考生基本信息.png`、`2027国考资料/`、`data/` 默认被 `.gitignore` 排除；
- 抓取使用**使用者自己的** Chrome 登录，工具不共享账号；
- `fetch` 只读取职位详情接口，不修改报名信息；
- 仓库只保留代码、文档和脱敏示例 `gk2027_config.example.json`。

---

## 免责

本项目是个人辅助工具，不是官方产品。绿色/黄色判断不替代招录机关的资格审查；黄色岗位请拨打职位表内咨询电话确认专业、学位、英语、户籍、体能等条件。报名人数会随报名进程变化，最终以系统实时数据为准。

---

## License

[MIT](LICENSE)

## 贡献

欢迎提 Issue / PR：规则补充、新的专业匹配、下载器适配、文档改进都可以。提交前请确保不包含任何个人报考信息、账号、截图或抓取数据。
