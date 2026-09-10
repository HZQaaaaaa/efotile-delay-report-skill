# efotile-delay-report-skill

方太销拓**订单延迟率报表**自动化技能：登录方太电商综合协同平台（oec.efotile.net）导出「销售摘要」，
一键生成带函数的《订单延迟率》Excel 和独立可视化 HTML 看板（双击即开、可发他人、右上角一键导出带函数 Excel）。

## 能力

1. 浏览器自动化：B2C销售管理 → 订单管理 → 全部订单，按创建时间范围导出「销售摘要」（含 XSRF 处理、文件回传）；
2. 报表预处理：销售摘要新增 5 列公式（门店简称 VLOOKUP / 登记日期 / 5/7/10 天未发）、内置店铺匹配表、
   发货统计—延迟发货 sheet（17 家销拓店铺 × 3 渠道，COUNTIFS/SUMIF/发货率公式，打开自动重算）；
3. 可视化看板：KPI 卡片、自动结论、发货率柱状图、未发货结构堆叠图、明细表，数据全内嵌。

## 安装

方式一（skills CLI）：

```bash
npx skills add <owner>/efotile-delay-report-skill
```

方式二（手动）：将本仓库内容复制到工作区 `.trae/skills/efotile-delay-report/` 目录下
（保持 SKILL.md 在该目录根部）。

## 依赖

- Python 3 + `openpyxl`
- 浏览器自动化环境（用于平台导出步骤）

## 使用

在对话中说明要生成销拓订单延迟率报表并给出时间范围即可，技能会引导完成
拉取 → 预处理 → 看板生成。脚本也可单独运行：

```bash
python scripts/make_html.py <销售摘要.xlsx> <输出目录> 2026-09-01 2026-09-10
```

## 目录结构

```
SKILL.md                      技能说明与完整操作流程
scripts/build_report.py       报表构建（openpyxl 从零生成带函数 Excel）
scripts/make_html.py          可视化看板生成（内嵌数据 + 导出按钮）
scripts/file_receiver.py      浏览器下载文件本地接收服务（端口 8765）
assets/店铺名称匹配.csv        内置店铺匹配表（系统名称→门店简称/代理商/分类）
```

## 业务口径

- 17 家销拓店铺：太和会 / 新平台 / 其他 三渠道；
- 累计发货 = 发货在途单 + 待发货；未发货 = 预订单 + 待审核；已取消单不计；
- 延迟分档：5–6 天 / 7–9 天 / ≥10 天（按登记时间至 TODAY()）。
