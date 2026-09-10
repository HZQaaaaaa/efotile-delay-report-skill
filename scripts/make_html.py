# -*- coding: utf-8 -*-
"""
生成可视化 HTML 看板：数据内嵌、双击即开，右上角按钮导出带函数 xlsx。
用法：
  python make_html.py <下载报表.xlsx> <输出目录> <开始日期> <结束日期>
"""
import os, sys, json, base64
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_report import build

def pct(v):
    return '-' if v is None else '%.1f%%' % (v * 100)

def make(res, xlsx_path, html_path):
    rows = res['rows']
    total = res['total']
    start, end = res['start'], res['end']
    base = start.replace('-', '') + '-' + end.replace('-', '')
    xlsx_name = '订单延迟率—欣怡_%s.xlsx' % base

    with open(xlsx_path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode('ascii')

    tot_orders = total['C'] + total['D']

    trs = []
    last_ch = None
    for r in rows:
        ch_cell = ''
        if r['channel'] != last_ch:
            span = sum(1 for x in rows if x['channel'] == r['channel'])
            ch_cell = '<td class="ch" rowspan="%d">%s</td>' % (span, r['channel'])
            last_ch = r['channel']
        def cls(v):
            return ' class="warn"' if v is not None and v < 1 else ''
        trs.append(
            '<tr>%s<td class="shop">%s</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td>'
            '<td%s>%s</td><td%s>%s</td><td%s>%s</td></tr>' % (
                ch_cell, r['shop'], r['C'], r['D'], r['E'], r['F'], r['G'],
                cls(r['h']), pct(r['h']), cls(r['i']), pct(r['i']), cls(r['j']), pct(r['j'])))
    trs.append(
        '<tr class="sum"><td colspan="2">销售拓展业务合计</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td><td>%d</td>'
        '<td>%s</td><td>%s</td><td>%s</td></tr>' % (
            total['C'], total['D'], total['E'], total['F'], total['G'],
            pct(total['h']), pct(total['i']), pct(total['j'])))
    table_rows = '\n'.join(trs)

    bars1 = []
    shops1 = []
    for r in rows:
        shops1.append(r['shop'])
        for key, color in (('h', '#c0392b'), ('i', '#e67e22'), ('j', '#2e86de')):
            v = r[key]
            h = 0 if v is None else v * 100
            bars1.append(
                '<div class="bar" style="height:%.1f%%;background:%s" title="%s %d天发货率 %s"></div>'
                % (max(h, 1.5), color, r['shop'], {'h': 5, 'i': 7, 'j': 10}[key], pct(v)))

    bars2 = []
    max_d = max((r['D'] for r in rows), default=1) or 1
    for r in rows:
        if r['D'] == 0:
            continue
        normal = r['D'] - r['E'] - r['F'] - r['G']
        segs = []
        for val, color, name in (
                (normal, '#b2bec3', '5天内未发'), (r['E'], '#f39c12', '5天未发'),
                (r['F'], '#e74c3c', '7天未发'), (r['G'], '#8e44ad', '10天未发')):
            if val:
                w = val / max_d * 100
                segs.append('<div class="seg" style="width:%.2f%%;background:%s" title="%s：%d单">%s</div>'
                            % (w, color, name, val, val if val >= 3 else ''))
        bars2.append('<div class="row2"><span class="lbl">%s</span><div class="track">%s</div>'
                     '<span class="num">%d</span></div>' % (r['shop'], ''.join(segs), r['D']))
    bars2 = '\n'.join(bars2) or '<p class="empty">无未发货订单</p>'

    delayed = [r for r in rows if (r['E'] + r['F'] + r['G']) > 0]
    delayed.sort(key=lambda r: -(r['E'] + r['F'] + r['G']))
    top_delay = '、'.join('%s（延迟%d单：5天%d/7天%d/10天%d）' % (
        r['shop'], r['E'] + r['F'] + r['G'], r['E'], r['F'], r['G']) for r in delayed[:5]) or '无'
    low_h = [r for r in rows if r['h'] is not None and r['h'] < total['h']]
    low_txt = '、'.join('%s（%s）' % (r['shop'], pct(r['h'])) for r in low_h) or '全部店铺均不低于整体水平'
    summary = (
        '本期（%s 至 %s）覆盖 17 家销拓店铺，订单共 %d 单：%s %d 单、%s %d 单。'
        '整体 5 天发货率 <b>%s</b>、7 天发货率 <b>%s</b>、10 天发货率 <b>%s</b>。'
        '未发货订单中 5 天未发 <b>%d</b> 单、7 天未发 <b>%d</b> 单、10 天未发 <b>%d</b> 单。'
        '存在延迟发货的店铺：%s。5 天发货率低于整体水平（%s）的店铺：%s。' % (
            start, end, tot_orders, res['label_c'], total['C'], res['label_d'], total['D'],
            pct(total['h']), pct(total['i']), pct(total['j']),
            total['E'], total['F'], total['G'], top_delay, pct(total['h']), low_txt))

    html = TEMPLATE
    repl = {
        '__TITLE__': '发货统计—延迟发货看板',
        '__RANGE__': '%s 至 %s' % (start, end),
        '__XLSX_NAME__': xlsx_name,
        '__B64__': b64,
        '__LABEL_C__': res['label_c'],
        '__LABEL_D__': res['label_d'],
        '__KPI_C__': str(total['C']),
        '__KPI_D__': str(total['D']),
        '__KPI_H__': pct(total['h']),
        '__KPI_I__': pct(total['i']),
        '__KPI_J__': pct(total['j']),
        '__TABLE_ROWS__': table_rows,
        '__BARS1_JSON__': json.dumps(bars1, ensure_ascii=False),
        '__SHOPS1_JSON__': json.dumps(shops1, ensure_ascii=False),
        '__BARS2__': bars2,
        '__SUMMARY__': summary,
    }
    for k, v in repl.items():
        html = html.replace(k, v)

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return html_path


TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>__TITLE__</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:"Microsoft YaHei","PingFang SC",sans-serif; background:#f4f6f9; color:#2d3436; padding:24px; }
.wrap { max-width:1180px; margin:0 auto; }
.header { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:18px; }
.header h1 { font-size:24px; color:#c00000; }
.header .sub { color:#7f8c8d; font-size:13px; margin-top:6px; }
.btn { background:#c00000; color:#fff; border:none; padding:11px 22px; font-size:15px; border-radius:6px;
       cursor:pointer; box-shadow:0 2px 6px rgba(192,0,0,.3); white-space:nowrap; }
.btn:hover { background:#a00000; }
.kpis { display:grid; grid-template-columns:repeat(5,1fr); gap:14px; margin-bottom:18px; }
.kpi { background:#fff; border-radius:10px; padding:16px 18px; box-shadow:0 1px 4px rgba(0,0,0,.06); }
.kpi .v { font-size:26px; font-weight:700; color:#c00000; }
.kpi .v.blue { color:#2e86de; }
.kpi .l { font-size:13px; color:#7f8c8d; margin-top:4px; }
.card { background:#fff; border-radius:10px; padding:18px 20px; margin-bottom:18px; box-shadow:0 1px 4px rgba(0,0,0,.06); }
.card h2 { font-size:16px; margin-bottom:14px; color:#2d3436; border-left:4px solid #c00000; padding-left:10px; }
.summary { line-height:1.9; font-size:14px; color:#444; }
table { width:100%; border-collapse:collapse; font-size:13px; }
th,td { border:1px solid #e3e6ea; padding:7px 8px; text-align:center; }
thead th { background:#c00000; color:#fff; font-weight:600; }
td.ch { background:#fdf0f0; font-weight:700; writing-mode:vertical-lr; letter-spacing:4px; }
td.shop { text-align:left; font-weight:600; }
tr.sum td { background:#fff5f5; font-weight:700; }
td.warn { color:#e74c3c; font-weight:700; }
.chart1 { display:flex; align-items:flex-end; gap:6px; height:240px; padding:10px 4px 0; border-bottom:1px solid #ddd; }
.chart1 .grp { flex:1; display:flex; align-items:flex-end; gap:2px; height:100%; }
.chart1 .bar { flex:1; border-radius:3px 3px 0 0; min-height:2px; }
.xlabels { display:flex; gap:6px; margin-top:6px; }
.xlabels span { flex:1; text-align:center; font-size:11px; color:#666; transform:rotate(-30deg); white-space:nowrap; }
.legend { font-size:12px; color:#666; margin-bottom:8px; }
.legend i { display:inline-block; width:10px; height:10px; border-radius:2px; margin:0 4px 0 12px; }
.row2 { display:flex; align-items:center; margin:7px 0; font-size:12px; }
.row2 .lbl { width:78px; text-align:right; padding-right:10px; color:#444; }
.row2 .track { flex:1; display:flex; height:20px; background:#f0f2f5; border-radius:3px; overflow:hidden; }
.row2 .seg { height:100%; color:#fff; font-size:11px; line-height:20px; text-align:center; overflow:hidden; }
.row2 .num { width:36px; padding-left:8px; font-weight:700; color:#c00000; }
.empty { color:#999; padding:20px; text-align:center; }
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <div>
      <h1>__TITLE__</h1>
      <div class="sub">数据范围：__RANGE__　|　数据来源：B2C销售管理-全部订单-销售摘要</div>
    </div>
    <button class="btn" onclick="exportXlsx()">⬇ 导出报表（带函数 Excel）</button>
  </div>

  <div class="kpis">
    <div class="kpi"><div class="v">__KPI_C__</div><div class="l">__LABEL_C__（单）</div></div>
    <div class="kpi"><div class="v">__KPI_D__</div><div class="l">__LABEL_D__（单）</div></div>
    <div class="kpi"><div class="v blue">__KPI_H__</div><div class="l">5天发货率</div></div>
    <div class="kpi"><div class="v blue">__KPI_I__</div><div class="l">7天发货率</div></div>
    <div class="kpi"><div class="v blue">__KPI_J__</div><div class="l">10天发货率</div></div>
  </div>

  <div class="card">
    <h2>整体情况</h2>
    <div class="summary">__SUMMARY__</div>
  </div>

  <div class="card">
    <h2>各店铺发货率对比</h2>
    <div class="legend">
      <i style="background:#c0392b"></i>5天发货率<i style="background:#e67e22"></i>7天发货率<i style="background:#2e86de"></i>10天发货率
    </div>
    <div class="chart1" id="chart1"></div>
    <div class="xlabels" id="xlabels"></div>
  </div>

  <div class="card">
    <h2>未发货订单结构（单）</h2>
    <div class="legend">
      <i style="background:#b2bec3"></i>5天内未发<i style="background:#f39c12"></i>5天未发<i style="background:#e74c3c"></i>7天未发<i style="background:#8e44ad"></i>10天未发
    </div>
    __BARS2__
  </div>

  <div class="card">
    <h2>发货统计—延迟发货明细</h2>
    <table>
      <thead>
        <tr><th>渠道</th><th>店铺</th><th>__LABEL_C__</th><th>__LABEL_D__</th><th>5天未发</th><th>7天未发</th><th>10天未发</th><th>5天发货率</th><th>7天发货率</th><th>10天发货率</th></tr>
      </thead>
      <tbody>
        __TABLE_ROWS__
      </tbody>
    </table>
  </div>
</div>

<script>
window.__XLSX_B64__ = "__B64__";
window.__XLSX_NAME = "__XLSX_NAME__";
function exportXlsx(){
  var b64 = window.__XLSX_B64__, bin = atob(b64), len = bin.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  var blob = new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = window.__XLSX_NAME;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
var bars = __BARS1_JSON__;
var shops = __SHOPS1_JSON__;
var c1 = document.getElementById('chart1');
for (var s = 0; s < shops.length; s++){
  var grp = document.createElement('div'); grp.className = 'grp';
  for (var k = 0; k < 3; k++){
    var d = document.createElement('div');
    d.innerHTML = bars[s*3+k];
    grp.appendChild(d.firstChild);
  }
  c1.appendChild(grp);
}
var xl = document.getElementById('xlabels');
for (var s2 = 0; s2 < shops.length; s2++){
  var sp = document.createElement('span');
  sp.textContent = shops[s2];
  xl.appendChild(sp);
}
</script>
</body>
</html>
"""

if __name__ == '__main__':
    dl = sys.argv[1]
    out_dir = sys.argv[2]
    start = sys.argv[3]
    end = sys.argv[4]
    base = start.replace('-', '') + '-' + end.replace('-', '')
    os.makedirs(out_dir, exist_ok=True)
    out_xlsx = os.path.join(out_dir, '订单延迟率—欣怡_%s.xlsx' % base)
    out_html = os.path.join(out_dir, '发货延迟率看板_%s.html' % base)
    res = build(dl, out_xlsx, start, end)
    make(res, out_xlsx, out_html)
    print('HTML:', out_html)
    print('XLSX:', out_xlsx)
    print('summary total:', json.dumps(res['total'], ensure_ascii=False))

