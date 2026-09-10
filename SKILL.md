---
name: "efotile-delay-report"
description: "Pulls 销售摘要 from oec.efotile.net (方太电商协同平台), builds a formula-rich 订单延迟率 Excel and a self-contained HTML dashboard with export button. Invoke when user asks to pull/generate 销拓订单延迟率/发货延迟报表 or 销售摘要导出+可视化看板."
---

# 方太销拓订单延迟率报表（拉取 → 预处理 → 可视化看板）

端到端完成：登录方太电商综合协同平台 → 按时间范围导出「销售摘要」→ 生成带函数的
《订单延迟率》Excel（销售摘要新增 5 列公式 + 内置店铺匹配表 + 发货统计 sheet）→
生成独立可视化 HTML 看板（数据内嵌、双击即开、右上角一键导出带函数 Excel）。

## 触发场景

- 用户要求拉取/生成"销拓订单延迟率""发货延迟""发货统计—延迟发货"报表/看板；
- 用户要求从 oec.efotile.net 导出"销售摘要"并做处理/可视化；
- 用户提到"欣怡那份表""订单延迟率—欣怡"格式的定期报表。

## 前置依赖

- Python 3 + openpyxl（`pip install openpyxl`）；
- 浏览器自动化能力（integrated_code_mode 的 `tools.browser_*`）；
- 本 skill 目录结构：`scripts/`（build_report.py、make_html.py、file_receiver.py）、
  `assets/店铺名称匹配.csv`（内置匹配表，含 17 家销拓店铺 E系统名称→G现有代理商 映射）。

## 执行流程

### 第 0 步：与用户确认时间范围

**必须主动询问**创建时间起止（如 2026-09-01 至 2026-09-10），不要自行假设。

### 第 1 步：启动本地文件接收服务（后台）

```
python <skill目录>/scripts/file_receiver.py <工作区>/data 8765
```

浏览器无法直接落盘下载文件，报表通过页面内 fetch POST 到该服务（端口 8765，
已带 CORS 头）。先 GET `http://127.0.0.1:8765/` 确认返回 "receiver ok"。

### 第 2 步：浏览器导出「销售摘要」

平台：`https://oec.efotile.net/`（AngularJS SPA，hash 路由 `#!/tenant/orders/index`）。

1. 打开平台；未登录则请用户手动登录（登录一次后 cookie 持久化，后续免登）。
2. 菜单：**B2C销售管理 → 订单管理 → 全部订单**。
3. 结算状态下拉从"未归档"改为 **"全部"**。
4. 开始时间/结束时间填用户确认的日期（格式 2026-09-01），点放大镜搜索。
5. 导出："导出"是文本节点，需 **hover** 弹出菜单（直接 click 不触发），点菜单中
   **"销售摘要"** 按钮，确认框"你确定吗?确定导出么?"点"是"。

**关键：浏览器自动化环境（JxBrowser）没有下载处理器**，直接导航下载 URL 会
`net::ERR_ABORTED` 且 fileToken 被一次性消费。必须在页面内用 `browser_evaluate`
执行 XHR 完成"导出 + 下载 + 回传"：

- 导出接口：`POST /api/services/app/order/GetOrdersToExcel`，
  Content-Type `application/json;charset=utf-8`，**必须带 `X-XSRF-TOKEN` 头**，
  值取自 `document.cookie` 中的 `XSRF-TOKEN`（AngularJS 防 CSRF，不带会 400）。
- 请求体 JSON 字段（固定值）：`productIds/isFutures/isTradeIn/startTime/endTime/
  shopId:"0"/orderStatereal/isTransDRP/isPostBack/timeKeyWord:"2"/flagIcon/
  isFrozen:"0"/isOldData:2/orderTaggingJson/giftOrderType/auditorId:"0"/
  creatorUserId:"0"/warehouseId:0/type:0/OrderSourceNo/isLimitedPrice:false/
  accountFlag:false`（startTime/endTime 用用户确认的日期）。
- 响应：`{result:{fileName,fileType,fileToken}}`。
- 立即同页 XHR GET 下载：
  `/File/DownloadTempFile?fileType=application/vnd.openxmlformats-officedocument.
  spreadsheetml.sheet&fileToken=<token>&fileName=<URL编码的"销售摘要.xlsx">`，
  `responseType:"arraybuffer"` → 转 Blob → `fetch('http://127.0.0.1:8765/save?name=
  销售摘要_<起>-<止>.xlsx',{method:'POST',body:blob})` 落盘。
- 注意：大段 JS 直接传 browser_evaluate 易 SyntaxError，拆成多个小 evaluate
  分步执行；每一步先重新 snapshot（页面元素 ref 会失效）。

落盘文件应为 <工作区>/data/销售摘要_YYYYMMDD-YYYYMMDD.xlsx，sheet 名"销售摘要"，
51 列、表头含"原始单号/订单号/店铺名/.../订单状态(F列)/.../登记时间(AG列,第28列)"。

### 第 3 步：生成带函数 Excel + 可视化看板

```
python <skill目录>/scripts/make_html.py <下载报表.xlsx> <输出目录> <开始日期> <结束日期>
```

例：
```
python .trae/skills/efotile-delay-report/scripts/make_html.py ^
  data/销售摘要_20260901-20260910.xlsx out 2026-09-01 2026-09-10
```

产出（输出目录下）：
- `订单延迟率—欣怡_YYYYMMDD-YYYYMMDD.xlsx`（带函数，打开自动重算）：
  - **销售摘要**：F 列起为下载的 51 列原始数据；A–E 为新增公式列：
    - A 门店简称 `=IFERROR(VLOOKUP(H2,店铺名称匹配!$E:$G,3,0),"")`
      （H=店铺名 → 匹配表 E 系统名称 → 返回 G 现有代理商；未命中显空白）
    - B 订单登记日期 `=TEXT(AG2,"yyyy年mm月dd日")`
    - C 5天未发 / D 7天未发 / E 10天未发：IF+AND+OR(待审核/预订单) + TODAY()-B 天数判断
  - **店铺名称匹配**：内置匹配表（E:I），VLOOKUP 引用内部表，无外链弹窗；
  - **发货统计—延迟发货**：渠道(太和会/新平台/其他)×17 店铺，
    COUNTIFS（累计发货=发货在途单+待发货；未发货=预订单+待审核）、
    SUMIF（5/7/10 天未发）、发货率公式；月份表头按时间范围自动生成；
    合计行 SUM + 整体发货率。
- `发货延迟率看板_YYYYMMDD-YYYYMMDD.html`：独立文件（数据+Excel 均内嵌 base64），
  含 KPI 卡片、自动结论文案、发货率柱状图、未发货结构堆叠图、明细表；
  右上角"导出报表（带函数 Excel）"按钮直接下载上述 xlsx。

### 第 4 步：交付与核对

- 向用户报告输出文件路径与核心数字（累计发货/未发货/5/7/10 天发货率）；
- 脚本 stdout 打印 Python 独立统计结果，应与 Excel 公式口径一致，可用于验算；
- 用户如需核对公式：A 列命中行应全部落在 17 家店铺内，非销拓店铺行为空白。

## 重要业务口径（勿改）

- 17 家店铺分渠道：太和会(普林斯顿/上海探影/徐州青葵/杭州松伽/广州邦耐/南京乾璟佳/
  宁波尚博)、新平台(安徽中厨/合肥维森/北京拾光机/天津倾心/宁波紫竹/徐州青匠)、
  其他(同舟益源/长沙福品/中科云商/佛山优优乐)。
- 匹配表已含口径修正：广州市邦耐贸易→"广州邦耐"、天津倾心（新平台）G 列→"天津倾心"。
  匹配表变更时同步更新 `assets/店铺名称匹配.csv`。
- 订单状态口径：累计发货 = 发货在途单 + 待发货；未发货 = 预订单 + 待审核；
  已取消单不计入。延迟天数按登记时间（AG 列，datetime）到 TODAY() 的天数分档
  5–6 天 / 7–9 天 / ≥10 天。
