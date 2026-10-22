# 中国企业温室气体排放核算与报告因子库 · 网页平台

> China Corporate GHG Accounting and Reporting Emission Factor Database

一个**零构建的纯静态网页平台**，把一份 Excel 因子表（626 条排放因子）呈现为可检索、可筛选、可下载的在线数据产品，并附带方法说明、常见问题、参考文献与建设团队介绍。桌面端与移动端分别做了原生化的排版设计。

- **在线访问**：<https://china-ghg-factor-platform-20260709105905.netlify.app>
- **当前数据版本**：`20260726`（适用期间 2025-01-01 ~ 2025-12-31，共 626 条因子）

## 开源与使用范围

本项目自行编写的软件代码采用 [MIT 许可证](./202610220805—代码开源许可证.txt)。使用或再发布代码时，请保留该许可证及版权声明。

排放因子数据、Excel 工作簿、参考文献、第三方内容和 CPCD 标识不属于上述代码许可证的授权范围；其使用、引用和再分发应遵循各自适用的权利和来源要求。数据来源详见站内参考文献。

登录和注册仅为本地演示：账户信息保存在浏览器本地，不具备生产环境所需的认证与安全保障，不能用于保存真实密码或保护敏感数据。

---

## 目录

- [它是什么](#它是什么)
- [功能](#功能)
- [技术栈](#技术栈零构建)
- [目录结构](#目录结构)
- [数据模型](#数据模型)
- [本地运行](#本地运行)
- [更新数据（Excel → JSON）](#更新数据excel--json)
- [部署（Netlify）](#部署netlify)
- [缓存与版本号（改样式前必读）](#缓存与版本号改样式前必读)
- [响应式与设计语言](#响应式与设计语言)
- [数据来源与免责声明](#数据来源与免责声明)

---

## 它是什么

平台的全部数据由一份 Excel 工作簿离线抽取成 JSON 和对应的本地脚本数据。前端是一个**没有任何框架、没有打包步骤**的单页应用，可通过本地服务器预览，也可直接用浏览器打开 `index.html`。

因子按 GHG Protocol 的三个范围组织：

| 范围 | 说明 | 因子数 |
|---|---|---:|
| 范围一 Scope 1 | 直接排放（固定源/移动源燃烧、工业过程、逸散、生物质燃烧） | 210 |
| 范围二 Scope 2 | 外购能源（外购电力、外购其他能源） | 50 |
| 范围三 Scope 3 | 价值链间接排放（外购商品与服务、资本商品、上下游运输、废弃物等） | 366 |
| **合计** | 共 19 个因子类别 | **626** |

## 功能

- **数据库入口** — 封面提供产品数据库和企业数据库两个入口；产品数据库跳转到 `https://lca.cityghg.com/`，企业数据库进入本项目首页。
- **因子库工作台** — 626 条因子的表格视图，支持关键词搜索、按范围/类别筛选、范围一活动数据单位切换、分页加载；点开任一条查看数值、单位、各气体分量、系统边界、计算方法与数据来源。
- **五个内容页** — 因子库、方法说明、Q&A（40 条）、参考文献（378 条）、建设团队。
- **首页封面** — 一屏一版、随浏览器高度自适应的编辑风封面，含范围三色数据阵与类别索引。
- **演示账户** — 登录/注册、收藏、偏好设置。**纯前端演示**，账户与收藏只存在当前浏览器的 `localStorage`，不上传、不具备真实账号安全性。
- **数据下载** — 提供源 Excel 文件下载（`downloads/`）。

## 技术栈（零构建）

- 前端：原生 HTML / CSS / JavaScript（ES2020），**无框架、无打包、无 `node_modules`、无 `package.json`**。
- 样式：CSS 自定义属性（`tokens.css` 为唯一色彩/尺度来源）、`oklch` 色彩空间、`color-mix()`、`clamp()`、scroll-snap、`backdrop-filter`、安全区 `env(safe-area-inset-*)`。
- 数据管线：Python 3 + [`openpyxl`](https://pypi.org/project/openpyxl/)（仅在更新数据时用到，运行时不需要）。
- 托管：Netlify（静态托管，无构建命令）。

## 目录结构

```
emissions-platform/
├── index.html              # 页面骨架、头部导航、弹窗、底部 Tab Bar
├── styles.css              # 全部样式（顶部 @import tokens.css）
├── tokens.css              # 设计 Token：色彩 / 间距 / 字号（改配色只动这里）
├── app.js                  # 全部前端逻辑（渲染、筛选、搜索、路由、账户）
├── data/
│   ├── platform-data.json              # 由 Excel 抽取的源数据（见下方数据模型）
│   └── 202609201721-platform-data.js   # 同步生成，供 file:// 直接打开时加载
├── assets/                 # 图片资源（logo、封面大图）
├── downloads/              # 可供下载的源 Excel 文件
├── scripts/
│   └── extract_workbook.py # Excel → JSON / JS 的抽取脚本
├── netlify.toml            # Netlify 托管配置（publish=".", 缓存头)
└── README.md
```

前端渲染逻辑集中在 `app.js`：一个全局 `state` 对象 + 一组 `render*() / show*()` 函数（如 `renderHomePage`、`renderTable`、`renderContentPage`、`renderGroupedQa`、`showDataView`、`showDetail`），入口在文件末尾的 `bindEvents()`。

## 数据模型

`data/platform-data.json` 的顶层结构：

| 键 | 类型 | 内容 |
|---|---|---|
| `metadata` | object | 标题、`version`、`applicablePeriod`、`sourceFile`、`totalFactors`、`scopeCounts`、`downloadFile` |
| `factors` | array(626) | 全部因子记录 |
| `categoryCounts` | object(19) | 各「范围::类别」的因子计数 |
| `contentPages` | object | 基本信息 / 导航目录 / Q&A / 参考文献 的页面文案 |
| `references` | array(378) | 参考文献列表 |
| `qaItems` | array(40) | 常见问题 |

单条因子记录（`factors[i]`）主要字段：

```jsonc
{
  "id": "S1-001",
  "scope": "范围一",
  "scopeCode": "Scope 1",
  "category": "固定源燃烧",
  "sourceName": "电煤",
  "activityUnit": "TJ",
  "factorHeader": "t CO2e",
  "factorValue": 96.8775,
  "factorUnit": "t CO2e/TJ",
  "gasValues": { "t CO2e": 96.8775, "t CO₂": 96.2, "kg CH₄": 10, "kg N₂O": 1.5 },
  "supportingValues": { "系统边界": "…", "计算方法": "…", "数据来源": "…" },
  "fields":   { /* 原始表格所有列，键为中文表头 */ },
  "formulas": { "t CO2e": "=CO2+(CH4*GWP(CH4)+N2O*GWP(N2O))/1000" },
  "searchText": "…用于前端搜索的归一化文本…"
}
```

> `fields` 保留 Excel 原始列（中文表头为键），`gasValues`/`supportingValues` 等是从中解析出来便于前端渲染的派生字段。`data/platform-data.json` 是**生成产物**——请改 Excel 再重新抽取，不要手改 JSON。

## 本地运行

无需安装依赖，任意静态服务器即可：

```bash
python3 -m http.server 5174
```

然后打开 <http://localhost:5174/>。

> 直接 `file://` 打开 `index.html` 会因为 `fetch()` 加载 JSON 受同源限制而失败，请务必用一个本地 HTTP 服务器。

## 更新数据（Excel → JSON）

数据全部来自一份 Excel 工作簿。修改因子请改 Excel，然后重新抽取：

```bash
# 需要 openpyxl：pip install openpyxl
python3 scripts/extract_workbook.py
```

脚本会读取源 Excel，同时重新生成 `data/platform-data.json` 和 `data/202609201721-platform-data.js`，并把源文件复制到 `downloads/`。源路径通过环境变量覆盖：

```bash
SOURCE_XLSX=/path/to/新版因子库.xlsx python3 scripts/extract_workbook.py
```

抽取后记得同步 `metadata.version` 对应的缓存版本号（见下）。

## 部署（Netlify）

`netlify.toml` 已配置：`publish = "."`、无构建命令、给 `/data/*` 与 `/downloads/*` 设了缓存头。推送到已连接的 Netlify 站点即可自动发布；用 CLI 手动发布：

```bash
npx --yes netlify-cli@latest deploy --prod --dir=.
```

> 发布前请确认没有把体积巨大的本地测试产物（截图、缓存）一并上传——这些目录已在 `.gitignore` 中忽略（`output/`、`.playwright-cli/`、`.hallmark/`、`.netlify/`、`node_modules/`）。

## 缓存与版本号（改样式前必读）

站点用查询串 `?v=…` 给静态资源做缓存失效。**有两处版本号，改样式后要一起更新，否则用户浏览器会拿到旧缓存：**

1. `index.html` 里对 `styles.css?v=…` 和 `app.js?v=…` 的引用。
2. `styles.css` **第二行** `@import url("./tokens.css?v=…")` —— 这一处独立存在，**改了 `tokens.css`（配色/尺度）必须同时手动 bump 它**，否则改动到不了已缓存的浏览器。这是本项目最容易踩的坑。

## 响应式与设计语言

- 断点：`min-width: 921px` 为桌面版式，`max-width: 920px` 走移动端。两套版式共享同一份 `tokens.css`。
- 移动端按「把网页当 App 做」的思路设计：系统字体栈、底部固定 Tab Bar、安全区适配、一屏一版滚动吸附、数据即图形、发丝线分段。配色为莫兰迪低饱和主色 + 范围三色（蓝/琥珀/珊瑚）徽章。
- 改配色只需动 `tokens.css` 的 Token（记得按上一节 bump `@import` 版本号）。

## 数据来源与免责声明

因子数据整理自政府及监管机构发布数据、标准与技术规范、学术文献、行业协会与研究机构报告等公开来源，详见站内「参考文献」页。**本平台仅供核算参考，实际监管报送以适用制度与主管部门要求为准。**
