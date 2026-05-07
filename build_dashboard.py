import pandas as pd
import json
from collections import defaultdict

# ========== 数据处理（同 gen_data.py 逻辑）==========
f = '/Users/anyonghasayyo/Desktop/WorkBuddy/月报工时分析看板/项目工时登记对象导出结果.xlsx'
df = pd.read_excel(f, sheet_name='项目工时登记数据')
df['发生时间'] = pd.to_datetime(df['发生时间'], errors='coerce')
df['日期'] = df['发生时间'].dt.strftime('%m-%d')
df['工时'] = pd.to_numeric(df['工作时长'], errors='coerce').fillna(0)

def pct(a, b):
    return round(a/b*100, 1) if b > 0 else 0

total      = df['工时'].sum()
records    = len(df)
projCount  = df['项目名称'].nunique()
workDays   = df['发生时间'].dropna().dt.date.nunique()
maintTotal = df[df['工作类型']=='维护']['工时'].sum()
highValue  = df[df['工作类型'].isin(['数字化升级','项目实施','售前'])]['工时'].sum()
highValuePct = pct(highValue, total)

def classify_maint(row):
    desc = str(row.get('内容描述',''))
    if row['工作类型'] != '维护':
        return '未分类'
    if any(k in desc for k in ['培训','咨询','如何','怎么','FAQ','使用说明','操作指引','讲解','介绍功能','怎么操作']):
        return '业务/功能咨询'
    if any(k in desc for k in ['新开','开店','上线','开通','初始化','新店']):
        return '业务开通'
    if any(k in desc for k in ['bug','BUG','报错','问题','故障','异常','补丁','修复','解决','错误']):
        return '问题处理'
    return '未分类'

df['maint_class'] = df.apply(classify_maint, axis=1)
optimizable = df[df['maint_class'].isin(['业务/功能咨询','业务开通'])]['工时'].sum()

p0_daily = df[df['项目等级']=='P0'].groupby(df['发生时间'].dt.date)['工时'].sum()
p0Daily  = round(p0_daily.mean(), 1) if len(p0_daily) > 0 else 0
allDaily = round(df.groupby(df['发生时间'].dt.date)['工时'].sum().mean(), 1)

typeColors = {'维护':'#F7664A','项目管理':'#FA8C16','数字化升级':'#4A6CF7',
              '学习/内部例会':'#722ED1','项目实施':'#13C2C2','售前':'#52C41A'}
typePie = [[wt, round(g['工时'].sum(),1)] for wt,g in df.groupby('工作类型')]

projMap = {}
for name, g in df.groupby('项目名称'):
    if pd.isna(name): continue
    h  = round(g['工时'].sum(),1)
    m  = round(g[g['工作类型']=='维护']['工时'].sum(),1)
    v  = round(g[g['工作类型'].isin(['数字化升级','项目实施','售前'])]['工时'].sum(),1)
    lv = g['项目等级'].dropna().iloc[0] if g['项目等级'].dropna().shape[0]>0 else '未分级'
    projMap[name] = {'h':h,'m':m,'v':v,'lv':lv}

projSorted   = sorted(projMap.items(), key=lambda x: -x[1]['h'])
projNames    = [x[0] for x in projSorted]
projHours    = [x[1]['h'] for x in projSorted]
projMaint    = [x[1]['m'] for x in projSorted]
projMaintPct = [pct(x[1]['m'],x[1]['h']) for x in projSorted]
projValuePct = [pct(x[1]['v'],x[1]['h']) for x in projSorted]
projLevel    = [x[1]['lv'] for x in projSorted]

creatorNames = [n for n,g in df.groupby('负责人（必填）') if pd.notna(n)]
workTypes   = ['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前']
creatorTotalHours   = [round(df[df['负责人（必填）']==n]['工时'].sum(),1) for n in creatorNames]
creatorTypeMatrix  = []
creatorMaintMatrix = []
creatorLevelMatrix = []
for n in creatorNames:
    g = df[df['负责人（必填）']==n]
    creatorTypeMatrix.append([round(g[g['工作类型']==wt]['工时'].sum(),1) for wt in workTypes])
    g2 = df[(df['负责人（必填）']==n)&(df['工作类型']=='维护')]
    creatorMaintMatrix.append([round(g2[g2['maint_class']==mc]['工时'].sum(),1) for mc in ['问题处理','业务/功能咨询','业务开通','未分类']])
    creatorLevelMatrix.append([round(g[g['项目等级']==lv]['工时'].sum(),1) for lv in ['P0','P1','P2','P3','普通级','未分级']])

creatorValue = []
for i,n in enumerate(creatorNames):
    t = creatorTotalHours[i]
    high = creatorTypeMatrix[i][2]+creatorTypeMatrix[i][3]+creatorTypeMatrix[i][5]
    low  = creatorTypeMatrix[i][0]
    creatorValue.append({'name':n,'total':t,'high':round(high,1),'mid':round(creatorTypeMatrix[i][1]+creatorTypeMatrix[i][4],1),'low':round(low,1),
                        'highPct':pct(high,t),'lowPct':pct(low,t)})

stdHours = 168
creatorStd = []
for i,n in enumerate(creatorNames):
    a = creatorTotalHours[i]
    p = round(a/stdHours*100,1)
    creatorStd.append({'name':n,'actual':a,'std':stdHours,'pct':p,'dif':round(a-stdHours,1),'status':'over' if p>100 else ('ok' if p>=80 else 'low')})

creatorProjTop3 = []
for n in creatorNames:
    g = df[df['负责人（必填）']==n]
    top = g.groupby('项目名称')['工时'].sum().sort_values(ascending=False).head(3)
    creatorProjTop3.append([[nm,round(hh,1)] for nm,hh in top.items()])

maintType = [[mt,round(g['工时'].sum(),1)] for mt,g in df[df['工作类型']=='维护'].groupby('维护类型') if pd.notna(mt)]
unc1 = df[(df['工作类型']=='维护')&(df['维护类型'].isna())]['工时'].sum()
if unc1>0: maintType.append(['未分类',round(unc1,1)])

biz_cols = ['餐饮产品事业部','风行产品事业部','龙决策产品事业部','吾享产品事业部','快餐产品事业部','总部报表']
prodLineMap = defaultdict(float)
for _,row in df[df['工作类型']=='维护'].iterrows():
    h = row['工时']
    for col in biz_cols:
        v = row[col]
        if pd.notna(v) and str(v).strip():
            prodLineMap[str(v).strip()] += h
prodLine = [[k,round(v,1)] for k,v in prodLineMap.items()]

if '维护原因' in df.columns:
    maintReason = [[r,round(g['工时'].sum(),1)] for r,g in df[df['工作类型']=='维护'].groupby('维护原因') if pd.notna(r)]
    uncR = df[(df['工作类型']=='维护')&(df['维护原因'].isna())]['工时'].sum()
    if uncR>0: maintReason.append(['未填写',round(uncR,1)])
else:
    maintReason = [['未填写',round(df[df['工作类型']=='维护']['工时'].sum(),1)]]

maintLevel = [[lv,round(g['工时'].sum(),1)] for lv,g in df[df['工作类型']=='维护'].groupby('项目等级') if pd.notna(lv)]
uncL = df[(df['工作类型']=='维护')&(df['项目等级'].isna())]['工时'].sum()
if uncL>0: maintLevel.append(['未分级',round(uncL,1)])

maintClass  = [[mc, round(df[df['maint_class']==mc]['工时'].sum(),1)] for mc in ['问题处理','业务/功能咨询','业务开通','未分类']]
goutClass  = [[mc, round(df[(df['维护类型']=='沟通答疑')&(df['maint_class']==mc)]['工时'].sum(),1)] for mc in ['问题处理','业务/功能咨询']]

projMaintClass = []
for name,g in df[df['工作类型']=='维护'].groupby('项目名称'):
    if pd.isna(name): continue
    projMaintClass.append({'name':name,
        '问题处理':round(g[g['maint_class']=='问题处理']['工时'].sum(),1),
        '业务/功能咨询':round(g[g['maint_class']=='业务/功能咨询']['工时'].sum(),1),
        '业务开通':round(g[g['maint_class']=='业务开通']['工时'].sum(),1),
        '未分类':round(g[g['maint_class']=='未分类']['工时'].sum(),1)})

hmProjs = [x[0] for x in sorted(zip(projNames,projMaint), key=lambda x:-x[1])[:15]]
hmTypes = list(set([mt[0] for mt in maintType if mt[0]!='未分类']))[:6]
hmData   = []
for pi,pn in enumerate(hmProjs):
    for ti,tn in enumerate(hmTypes):
        h = df[(df['项目名称']==pn)&(df['工作类型']=='维护')&(df['维护类型']==tn)]['工时'].sum()
        if h>0: hmData.append([ti,pi,round(h,1)])

dailyMap = defaultdict(lambda: defaultdict(float))
for _,row in df.iterrows():
    d = row['日期']; wt = row['工作类型']; h = row['工时']
    if pd.notna(d): dailyMap[d][wt] += h
allDates    = sorted(dailyMap.keys())
dailyDates  = allDates
dailyMaint  = [round(dailyMap[d].get('维护',0),1) for d in allDates]
dailyPM     = [round(dailyMap[d].get('项目管理',0),1) for d in allDates]
dailyDigital = [round(dailyMap[d].get('数字化升级',0),1) for d in allDates]
dailyImpl  = [round(dailyMap[d].get('项目实施',0),1) for d in allDates]
dailyLearn = [round(dailyMap[d].get('学习/内部例会',0),1) for d in allDates]
dailySale  = [round(dailyMap[d].get('售前',0),1) for d in allDates]

valueLayer = [
    ['高价值产出(数字化+实施+售前)', round(highValue,1), '#52C41A'],
    ['项目管理+学习', round(df[df['工作类型'].isin(['项目管理','学习/内部例会'])]['工时'].sum(),1), '#4A6CF7'],
    ['必要维护(问题处理)', round(df[df['maint_class']=='问题处理']['工时'].sum(),1), '#FA8C16'],
    ['可优化维护(咨询+开通)', round(df[df['maint_class'].isin(['业务/功能咨询','业务开通'])]['工时'].sum(),1), '#F7664A'],
    ['其他', round(df[~df['工作类型'].isin(['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前'])]['工时'].sum(),1), '#D9D9D9'],
]

top10Value = sorted([[n,round(projValuePct[i],1),round(projMaintPct[i],1),round(100-projValuePct[i]-projMaintPct[i],1)] for i,n in enumerate(projNames)], key=lambda x:-x[1])[:10]

D = dict(
    total=round(total,1), records=records, projCount=projCount, workDays=workDays,
    maintTotal=round(maintTotal,1), p0Daily=p0Daily, allDaily=allDaily,
    highValue=round(highValue,1), highValuePct=highValuePct, optimizable=round(optimizable,1),
    typePie=typePie, typeColors=typeColors,
    projNames=projNames, projHours=projHours, projMaint=projMaint,
    projMaintPct=projMaintPct, projValuePct=projValuePct, projLevel=projLevel,
    maintType=maintType, prodLine=prodLine, maintReason=maintReason, maintLevel=maintLevel,
    maintClass=maintClass, goutClass=goutClass, projMaintClass=projMaintClass,
    valueLayer=valueLayer, top10Value=top10Value,
    hmProjs=hmProjs, hmTypes=hmTypes, hmData=hmData,
    dailyDates=dailyDates, dailyMaint=dailyMaint, dailyPM=dailyPM,
    dailyDigital=dailyDigital, dailyImpl=dailyImpl, dailyLearn=dailyLearn, dailySale=dailySale,
    creatorNames=creatorNames, creatorTotalHours=creatorTotalHours,
    workTypes=workTypes, creatorTypeMatrix=creatorTypeMatrix,
    creatorProjTop3=creatorProjTop3, creatorValue=creatorValue,
    stdHours=stdHours, creatorStd=creatorStd,
    levelOrder=['P0','P1','P2','P3','普通级','未分级'], creatorLevelMatrix=creatorLevelMatrix,
    maintClasses=['问题处理','业务/功能咨询','业务开通','未分类'], creatorMaintMatrix=creatorMaintMatrix,
)

# ========== 生成自包含 HTML ==========
def to_js(obj, indent=0):
    sp  = '  ' * indent
    sp1 = '  ' * (indent+1)
    if isinstance(obj, dict):
        items = [sp1 + json.dumps(k, ensure_ascii=False) + ': ' + to_js(v, indent+1) for k,v in obj.items()]
        return '{\n' + ',\n'.join(items) + '\n' + sp + '}'
    elif isinstance(obj, list):
        if len(obj)==0: return '[]'
        if isinstance(obj[0], (list,tuple)):
            items = [sp1 + to_js(x, indent+1) for x in obj]
            return '[\n' + ',\n'.join(items) + '\n' + sp + ']'
        else:
            items = [sp1 + to_js(x, indent+1) for x in obj]
            return '[\n' + ',\n'.join(items) + '\n' + sp + ']'
    elif isinstance(obj, str):  return json.dumps(obj, ensure_ascii=False)
    elif isinstance(obj, bool):  return 'true' if obj else 'false'
    elif isinstance(obj, float):  return str(obj)
    elif isinstance(obj, int):  return str(obj)
    else:  return json.dumps(obj, ensure_ascii=False)

D_js = to_js(D)
min_date = df['发生时间'].min().strftime('%Y年%m月%d日')
max_date = df['发生时间'].max().strftime('%Y年%m月%d日')
date_range = f"{min_date} - {max_date}"

html = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>项目工时分析报告</title>
<script src="echarts.min.js"></script>
<style>
:root{{--bg:#f5f7fa;--card:#fff;--primary:#4A6CF7;--primary-light:#6B8AFF;--accent:#F7664A;--accent-orange:#FA8C16;--accent-green:#52C41A;--text:#1F2937;--text-sec:#6B7280;--border:#E5E7EB;--shadow:0 1px 3px rgba(0,0,0,.08);--shadow-lg:0 10px 15px rgba(0,0,0,.08)}}
*{margin:0;padding:0;box-sizing:border-box}
body{{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"PingFang SC","Microsoft YaHei",sans-serif;line-height:1.6}}
.header{{background:linear-gradient(135deg,#4A6CF7,#6B8AFF);color:#fff;padding:40px 0 30px;text-align:center}}
.header h1{{font-size:28px;font-weight:700;margin-bottom:8px}}
.header p{{font-size:15px;opacity:.85}}
.container{{max-width:1200px;margin:0 auto;padding:24px 20px 60px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px}}
.kpi-card{{background:var(--card);border-radius:12px;padding:20px 24px;box-shadow:var(--shadow);transition:transform .2s}}
.kpi-card:hover{{transform:translateY(-2px);box-shadow:var(--shadow-lg)}}
.kpi-label{{font-size:13px;color:var(--text-sec);margin-bottom:6px}}
.kpi-value{{font-size:28px;font-weight:700;color:var(--primary)}}
.kpi-sub{{font-size:12px;color:var(--text-sec);margin-top:4px}}
.kpi-card.accent .kpi-value{{color:var(--accent)}}
.kpi-card.orange .kpi-value{{color:var(--accent-orange)}}
.kpi-card.green .kpi-value{{color:var(--accent-green)}}
.section{{background:var(--card);border-radius:12px;padding:28px;margin-bottom:24px;box-shadow:var(--shadow)}}
.section-title{{font-size:18px;font-weight:700;margin-bottom:4px;display:flex;align-items:center;gap:8px}}
.section-title::before{{content:'';width:4px;height:20px;background:var(--primary);border-radius:2px}}
.section-desc{{font-size:13px;color:var(--text-sec);margin-bottom:20px;padding-left:12px}}
.chart-row{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
.chart-box{{width:100%;height:380px}}
.chart-full{{width:100%;height:420px}}
.insight-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-top:20px}}
.insight-card{{background:#FFFBF0;border:1px solid #FFE7BA;border-radius:10px;padding:16px 18px}}
.insight-card.critical{{background:#FFF1F0;border-color:#FFCCC7}}
.insight-card.info{{background:#F0F5FF;border-color:#ADC6FF}}
.insight-title{{font-size:14px;font-weight:700;margin-bottom:8px}}
.insight-body{{font-size:13px;color:var(--text-sec);line-height:1.7}}
.insight-body strong{{color:var(--text)}}
.data-table{{width:100%;border-collapse:collapse;font-size:13px;margin-top:12px}}
.data-table th{{background:#F9FAFB;padding:10px 12px;text-align:left;font-weight:600;border-bottom:2px solid var(--border);color:var(--text-sec)}}
.data-table td{{padding:9px 12px;border-bottom:1px solid var(--border)}}
.data-table tr:hover td{{background:#F9FAFB}}
.data-table .num{{text-align:right;font-variant-numeric:tabular-nums}}
.pct-badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600}}
.pct-high{{background:#FFF1F0;color:#F5222D}}
.pct-mid{{background:#FFF7E6;color:#FA8C16}}
.pct-low{{background:#F6FFED;color:#52C41A}}
.tag{{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;margin:1px 2px}}
.tag-red{{background:#FFF1F0;color:#F5222D}}
.tag-orange{{background:#FFF7E6;color:#FA8C16}}
.tag-blue{{background:#E6F7FF;color:#1890FF}}
.tag-green{{background:#F6FFED;color:#52C41A}}
.val-high{{color:#52C41A;font-weight:700}}
.val-mid{{color:#FA8C16;font-weight:700}}
.val-low{{color:#F5222D;font-weight:700}}
.footer{{text-align:center;padding:20px;color:var(--text-sec);font-size:12px}}
.highlight-box{{background:linear-gradient(135deg,#FFF7E6,#FFE7BA);border:1px solid #FFD591;border-radius:10px;padding:18px 20px;margin:16px 0}}
.highlight-box.blue{{background:linear-gradient(135deg,#E6F7FF,#BAE7FF);border-color:#91D5FF}}
.highlight-box .hl-title{{font-size:15px;font-weight:700;margin-bottom:8px}}
.highlight-box .hl-body{{font-size:13px;line-height:1.8;color:var(--text-sec)}}
.highlight-box .hl-body strong{{color:var(--text)}}
.sub-title{{font-size:15px;font-weight:600;margin:20px 0 12px;padding-left:10px;border-left:3px solid var(--primary)}}
.rank-1{{color:#FFD700;font-weight:700}}
.rank-2{{color:#C0C0C0;font-weight:700}}
.rank-3{{color:#CD7F32;font-weight:700}}
@media(max-width:768px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}.chart-row{{grid-template-columns:1fr}}.insight-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="header">
  <h1>📊 项目工时分析报告</h1>
  <p id="header-sub"></p>
</div>
<div class="container">
<div class="kpi-grid" id="kpi-grid"></div>

<div class="section">
  <div class="section-title">一、工作类型总览</div>
  <div class="section-desc">各类工作的时间投入与占比分布</div>
  <div class="chart-row">
    <div id="chart-type-pie" class="chart-box"></div>
    <div id="chart-type-bar" class="chart-box"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">二、项目维度分析</div>
  <div class="section-desc">各项目工时投入排名，含维护占比与价值产出比</div>
  <div id="chart-project-bar" class="chart-full"></div>
  <div style="margin-top:16px">
    <table class="data-table">
      <thead><tr><th>#</th><th>项目名称</th><th class="num">总工时</th><th class="num">维护</th><th class="num">维护占比</th><th class="num">价值产出比</th><th>等级</th></tr></thead>
      <tbody id="project-table-body"></tbody>
    </table>
  </div>
</div>

<div class="section">
  <div class="section-title">三、人员工时维度（负责人）</div>
  <div class="section-desc">按负责人分析工时投入、价值产出、标准工时对比</div>
  <div class="sub-title">📋 负责人工时总览 & 标准工时对比</div>
  <div id="chart-creator-std" class="chart-full"></div>
  <div style="margin-top:12px">
    <table class="data-table">
      <thead><tr><th>排名</th><th>负责人</th><th class="num">实际工时</th><th class="num">标准工时</th><th class="num">完成率</th><th class="num">差值</th><th>状态</th></tr></thead>
      <tbody id="creator-std-table"></tbody>
    </table>
  </div>
  <div class="sub-title">📊 负责人 × 工作类型分布</div>
  <div id="chart-creator-type" class="chart-full"></div>
  <div class="sub-title">💎 价值产出排名</div>
  <div class="chart-row">
    <div id="chart-creator-value" class="chart-box"></div>
    <div id="chart-creator-value-bar" class="chart-box"></div>
  </div>
  <div style="margin-top:12px">
    <table class="data-table">
      <thead><tr><th>负责人</th><th class="num">总工时</th><th class="num">高价值</th><th class="num">占比</th><th class="num">中价值</th><th class="num">维护</th><th class="num">维护占比</th><th>价值等级</th></tr></thead>
      <tbody id="creator-value-table"></tbody>
    </table>
  </div>
  <div class="sub-title">🔧 负责人 × 维护类型细分</div>
  <div id="chart-creator-maint" class="chart-full"></div>
  <div class="sub-title">🏗️ 负责人 × 项目等级分布</div>
  <div class="chart-row">
    <div id="chart-creator-level" class="chart-box"></div>
    <div id="chart-creator-proj" class="chart-box"></div>
  </div>
  <div class="highlight-box blue" style="margin-top:16px">
    <div class="hl-title">👥 人员投入洞察</div>
    <div class="hl-body" id="person-insight"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">四、产出价值维度</div>
  <div class="section-desc">按工作类型的价值贡献分层</div>
  <div class="chart-row">
    <div id="chart-value-layer" class="chart-box"></div>
    <div id="chart-value-project" class="chart-box"></div>
  </div>
  <div class="highlight-box" style="margin-top:16px">
    <div class="hl-title">💎 产出价值洞察</div>
    <div class="hl-body" id="value-insight"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">五、维护工时深度分析</div>
  <div class="section-desc" id="maint-desc"></div>
  <div class="chart-row">
    <div id="chart-maint-type" class="chart-box"></div>
    <div id="chart-maint-product" class="chart-box"></div>
  </div>
  <div class="chart-row" style="margin-top:20px">
    <div id="chart-maint-reason" class="chart-box"></div>
    <div id="chart-maint-level" class="chart-box"></div>
  </div>
</div>

<div class="section">
  <div class="section-title">六、维护类型细化：问题处理 vs 业务/功能咨询</div>
  <div class="section-desc">通过内容描述关键词智能分类</div>
  <div class="chart-row">
    <div id="chart-maint-class" class="chart-box"></div>
    <div id="chart-gout-class" class="chart-box"></div>
  </div>
  <div id="chart-maint-class-project" class="chart-full" style="margin-top:20px"></div>
  <div class="insight-grid" id="maint-insights"></div>
</div>

<div class="section">
  <div class="section-title">七、维护类型 × 项目 热力图</div>
  <div class="section-desc">可视化各项目在不同维护类型上的工时分布密度</div>
  <div id="chart-maint-heatmap" class="chart-full"></div>
</div>

<div class="section">
  <div class="section-title">八、每日工时趋势</div>
  <div class="section-desc">按日查看各工作类型的工时变化</div>
  <div id="chart-daily-trend" class="chart-full"></div>
</div>

<div class="section">
  <div class="section-title">九、优化建议</div>
  <div class="section-desc">基于以上分析，提出可落地的改进方向</div>
  <div class="insight-grid" id="optimize-insights"></div>
  <div class="highlight-box blue" style="margin-top:16px">
    <div class="hl-title">📈 优化预期</div>
    <div class="hl-body" id="optimize-expect"></div>
  </div>
</div>

<div class="footer" id="footer-text"></div>
</div>

<script>
var D = __D_PLACEHOLDER__;
</script>
</body>
</html>"""

# 修复CSS花括号（Python字符串格式化导致双括号问题）
html = html.replace('{{', '{').replace('}}', '}')

# 替换占位符
html = html.replace('__D_PLACEHOLDER__', D_js)

# 写入完整的chart逻辑（复用参考网站的JS逻辑）
# 在 </body> 前插入图表JS，插入点标记
html = html.replace('</body>', '__CHART_JS_PLACEHOLDER__\n</body>')

chart_js = r"""
var MC = ['#5B8FF9','#5AD8A6','#F6BD16','#E86452','#6DC8EC','#945FB9','#FF9845','#1E9493','#FF6B6B','#48C9B0'];
function pct(v,t){return t>0?(v/t*100).toFixed(1):'0.0'}
function initChart(id){
  var el=document.getElementById(id);
  if(!el) return null;
  return echarts.init(el);
}

// ===== Header & KPI =====
document.getElementById('header-sub').textContent = '数据周期：__DATE_RANGE__ | '+D.records+'条工时记录 · '+D.projCount+'个项目 · 总工时'+D.total+'h';
document.getElementById('footer-text').textContent = '报告生成时间：__GEN_DATE__ · 数据来源：项目工时登记对象导出结果 · 数据周期：__DATE_RANGE__';

var maintPct = pct(D.maintTotal, D.total);
var probPct = (function(){var x=D.maintClass.find(function(i){return i[0]==='问题处理'});return x?pct(x[1],D.maintTotal):'0';})();
document.getElementById('kpi-grid').innerHTML =
  '<div class="kpi-card"><div class="kpi-label">总工时</div><div class="kpi-value">'+D.total+'h</div><div class="kpi-sub">'+D.records+'条 · '+D.projCount+'个项目 · '+D.workDays+'工作日</div></div>'
+ '<div class="kpi-card accent"><div class="kpi-label">维护工时</div><div class="kpi-value">'+D.maintTotal+'h</div><div class="kpi-sub">占比'+maintPct+'% · 其中问题处理'+probPct+'%</div></div>'
+ '<div class="kpi-card orange"><div class="kpi-label">项目管理工时</div><div class="kpi-value">'+(function(){var x=D.typePie.find(function(i){return i[0]==='项目管理'});return x?x[1]:0;}())+'h</div><div class="kpi-sub">占比'+(function(){var x=D.typePie.find(function(i){return i[0]==='项目管理'});return x?pct(x[1],D.total):'0';}())+'%</div></div>'
+ '<div class="kpi-card green"><div class="kpi-label">高产出工时</div><div class="kpi-value">'+D.highValue+'h</div><div class="kpi-sub">数字化升级+项目实施+售前 占'+D.highValuePct+'%</div></div>';

// ===== 1. 工作类型 =====
(function(){
  var c1=initChart('chart-type-pie');
  if(c1) c1.setOption({title:{text:'工作类型工时占比',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:10,textStyle:{fontSize:12}},
    series:[{type:'pie',radius:['40%','70%'],center:['50%','52%'],itemStyle:{borderRadius:6,borderColor:'#fff',borderWidth:2},
      data:D.typePie.map(function(i){return{name:i[0],value:i[1],itemStyle:{color:(D.typeColors||{})[i[0]]||null}}),
      label:{formatter:'{b}\n{d}%',fontSize:11}}]});
  var c2=initChart('chart-type-bar');
  if(c2){var s=D.typePie.slice().sort(function(a,b){return a[1]-b[1]});
    c2.setOption({title:{text:'各工作类型工时(h)',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis'},grid:{left:100,right:40,top:50,bottom:30},xAxis:{type:'value'},yAxis:{type:'category',data:s.map(function(i){return i[0]}),axisLabel:{fontSize:12}},
      series:[{type:'bar',data:s.map(function(i){return{value:i[1],itemStyle:{color:(D.typeColors||{})[i[0]]||'#4A6CF7'}}}),barWidth:24,itemStyle:{borderRadius:[0,4,4,0]},label:{show:true,position:'right',formatter:'{c}h',fontSize:11}}]});
  }
})();

// ===== 2. 项目维度 =====
(function(){
  var c=initChart('chart-project-bar');
  if(c){var n=D.projNames.slice().reverse(),h=D.projHours.slice().reverse();
    c.setOption({title:{text:'各项目总工时排名(h)',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis'},grid:{left:160,right:50,top:50,bottom:30},xAxis:{type:'value'},yAxis:{type:'category',data:n,axisLabel:{fontSize:11}},
      series:[{type:'bar',data:h,barWidth:18,itemStyle:{borderRadius:[0,4,4,0],color:new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#4A6CF7'},{offset:1,color:'#6B8AFF'}])},label:{show:true,position:'right',formatter:'{c}h',fontSize:10}}]});
  }
  var tb=document.getElementById('project-table-body');
  if(tb){var html='';for(var i=0;i<D.projNames.length;i++){
    var mp=D.projMaintPct[i],vp=D.projValuePct[i];
    var mc=mp>60?'pct-high':mp>40?'pct-mid':'pct-low';
    var vc=vp>30?'val-high':vp>10?'val-mid':'val-low';
    var lv=D.projLevel[i]||'',lt=lv==='P0'?'tag-red':lv==='P1'?'tag-orange':lv==='P2'?'tag-blue':'tag-green';
    html+='<tr><td>'+(i+1)+'</td><td>'+D.projNames[i]+'</td><td class="num">'+D.projHours[i]+'h</td><td class="num">'+D.projMaint[i]+'h</td><td class="num"><span class="pct-badge '+mc+'">'+mp+'%</span></td><td class="num"><span class="'+vc+'">'+vp+'%</span></td><td><span class="tag '+lt+'">'+lv+'</span></td></tr>';
  }tb.innerHTML=html;}
})();

// ===== 3. 人员维度 =====
(function(){
  // 3a. 标准工时对比
  var c0=initChart('chart-creator-std');
  if(c0){var n=D.creatorNames,a=D.creatorTotalHours,std=D.stdHours;
    c0.setOption({title:{text:'负责人工时 vs 标准工时（21天×8h='+std+'h）',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',formatter:function(p){var s=p[0].name+'<br/>';p.forEach(function(x){s+=x.marker+x.seriesName+': '+x.value+'h<br/>'});return s}},legend:{data:['实际工时','标准工时'],bottom:5},grid:{left:80,right:40,top:50,bottom:50},xAxis:{type:'category',data:n,axisLabel:{fontSize:13}},yAxis:{type:'value',name:'工时(h)'},
      series:[{name:'实际工时',type:'bar',data:a,barWidth:36,itemStyle:{borderRadius:[4,4,0,0],color:function(p){return p.value/std*100>100?'#F5222D':p.value/std*100>80?'#52C41A':'#FA8C16'}},label:{show:true,position:'top',formatter:function(p){return p.value+'h'},fontSize:12,fontWeight:600}},
               {name:'标准工时',type:'line',data:Array(n.length).fill(std),lineStyle:{color:'#4A6CF7',type:'dashed',width:2},symbol:'none'}]});
  }
  var stb=document.getElementById('creator-std-table');
  if(stb){var html='';for(var i=0;i<D.creatorStd.length;i++){var s=D.creatorStd[i],rc=i===0?'rank-1':i===1?'rank-2':i===2?'rank-3':'',tg=s.status==='over'?'tag-red':s.status==='ok'?'tag-green':'tag-orange',tx=s.status==='over'?'超标':s.status==='ok'?'达标':'不足';
    html+='<tr><td class="'+rc+'">'+(i+1)+'</td><td>'+s.name+'</td><td class="num">'+s.actual+'h</td><td class="num">'+s.std+'h</td><td class="num"><span class="pct-badge '+(s.pct>100?'pct-high':s.pct>80?'pct-low':'pct-mid')+'">'+s.pct+'%</span></td><td class="num" style="color:'+(s.diff>0?'#F5222D':'#52C41A')+'">'+(s.diff>0?'+':'')+s.diff+'h</td><td><span class="tag '+tg+'">'+tx+'</span></td></tr>';
  }stb.innerHTML=html;}
  // 3b. 负责人×工作类型
  var c1=initChart('chart-creator-type');
  if(c1){var wtColors=['#F7664A','#FA8C16','#4A6CF7','#13C2C2','#722ED1','#52C41A'];
    c1.setOption({title:{text:'负责人 × 工作类型工时拆解',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:D.workTypes,bottom:5,textStyle:{fontSize:11}},grid:{left:80,right:30,top:50,bottom:50},xAxis:{type:'category',data:D.creatorNames,axisLabel:{fontSize:13}},yAxis:{type:'value',name:'工时(h)'},
      series:D.workTypes.map(function(wt,idx){return{name:wt,type:'bar',stack:'total',data:D.creatorTypeMatrix.map(function(row){return row[idx]}),itemStyle:{color:wtColors[idx]},label:{show:idx===0,position:'top',formatter:function(p){var t=0;D.creatorTypeMatrix[p.dataIndex].forEach(function(v){t+=v});return t+'h'},fontSize:11,color:'#1F2937',fontWeight:600}}}})});
  }
  // 3c. 价值产出
  var c2=initChart('chart-creator-value');
  if(c2){c2.setOption({title:{text:'负责人高价值产出占比',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'item',formatter:function(p){return p.name+'<br/>高价值占比: '+p.value+'%'}},
      series:[{type:'pie',radius:['35%','68%'],center:['50%','55%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
        data:D.creatorValue.map(function(v,i){var c=v.highPct>25?'#52C41A':v.highPct>15?'#FA8C16':'#F5222D';return{name:v.name,value:v.highPct,itemStyle:{color:c}}}),
        label:{formatter:'{b}\n{c}%',fontSize:11}}]});
  }
  var c3=initChart('chart-creator-value-bar');
  if(c3){var vs=D.creatorValue.slice().sort(function(a,b){return a.highPct-b.highPct});
    c3.setOption({title:{text:'负责人工时价值分层',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:['高价值','中价值','维护'],bottom:5,textStyle:{fontSize:11}},grid:{left:80,right:30,top:50,bottom:50},xAxis:{type:'value',max:100,name:'%'},yAxis:{type:'category',data:vs.map(function(v){return v.name})},
      series:[{name:'高价值',type:'bar',stack:'v',data:vs.map(function(v){return v.highPct}),itemStyle:{color:'#52C41A'},label:{show:true,position:'inside',formatter:function(p){return p.value>8?p.value+'%':''},fontSize:10,color:'#fff'}},
               {name:'中价值',type:'bar',stack:'v',data:vs.map(function(v){return v.midPct}),itemStyle:{color:'#4A6CF7'},label:{show:true,position:'inside',formatter:function(p){return p.value>8?p.value+'%':''},fontSize:10,color:'#fff'}},
               {name:'维护',type:'bar',stack:'v',data:vs.map(function(v){return v.lowPct}),itemStyle:{color:'#F7664A'},label:{show:true,position:'inside',formatter:function(p){return p.value>8?p.value+'%':''},fontSize:10,color:'#fff'}}]});
  }
  var vtb=document.getElementById('creator-value-table');
  if(vtb){var vs2=D.creatorValue.slice().sort(function(a,b){return b.highPct-a.highPct});var html='';for(var i=0;i<vs2.length;i++){var v=vs2[i],rc=i===0?'rank-1':i===1?'rank-2':i===2?'rank-3':'',hp=v.highPct>25?'val-high':v.highPct>15?'val-mid':'val-low',lp=v.lowPct>50?'pct-high':v.lowPct>30?'pct-mid':'pct-low',lv2=v.highPct>25?'tag-green':v.highPct>15?'tag-orange':'tag-red',tx2=v.highPct>25?'高':v.highPct>15?'中':'低';
    html+='<tr><td class="'+rc+'">'+(i+1)+'</td><td>'+v.name+'</td><td class="num">'+v.total+'h</td><td class="num">'+v.high+'h</td><td class="num"><span class="'+hp+'">'+v.highPct+'%</span></td><td class="num">'+v.mid+'h</td><td class="num">'+v.low+'h</td><td class="num"><span class="pct-badge '+lp+'">'+v.lowPct+'%</span></td><td><span class="tag '+lv2+'">'+tx2+'</span></td></tr>';
  }vtb.innerHTML=html;}
  // 3d. 负责人×维护类型
  var c4=initChart('chart-creator-maint');
  if(c4){var mcColors=['#FA8C16','#F7664A','#722ED1','#D9D9D9'];
    c4.setOption({title:{text:'负责人 × 维护类型工时拆解',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:D.maintClasses,bottom:5,textStyle:{fontSize:11}},grid:{left:80,right:30,top:50,bottom:50},xAxis:{type:'category',data:D.creatorNames,axisLabel:{fontSize:13}},yAxis:{type:'value',name:'工时(h)'},
      series:D.maintClasses.map(function(mc,idx){return{name:mc,type:'bar',stack:'m',data:D.creatorMaintMatrix.map(function(row){return row[idx]}),itemStyle:{color:mcColors[idx]},barWidth:36}})});
  }
  // 3e. 负责人×项目等级
  var c5=initChart('chart-creator-level');
  if(c5){var lvColors=['#F5222D','#FA8C16','#4A6CF7','#52C41A','#D9D9D9','#E8E8E8'];
    c5.setOption({title:{text:'负责人 × 项目等级工时',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:D.levelOrder,bottom:5,textStyle:{fontSize:11}},grid:{left:80,right:30,top:50,bottom:50},xAxis:{type:'category',data:D.creatorNames,axisLabel:{fontSize:13}},yAxis:{type:'value',name:'工时(h)'},
      series:D.levelOrder.map(function(lv,idx){return{name:lv,type:'bar',stack:'l',data:D.creatorLevelMatrix.map(function(row){return row[idx]}),itemStyle:{color:lvColors[idx]}}}})});
  }
  // 3f. 负责人×主要项目
  var c6=initChart('chart-creator-proj');
  if(c6){var pSet=[];D.creatorProjTop3.forEach(function(list){list.forEach(function(x){if(pSet.indexOf(x[0])===-1)pSet.push(x[0])})});
    var pColors=['#4A6CF7','#6B8AFF','#F7664A','#FA8C16','#52C41A','#722ED1','#13C2C2','#FF6B6B','#5AD8A6','#F6BD16'];
    c6.setOption({title:{text:'负责人主要项目工时（Top3）',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:pSet,bottom:5,textStyle:{fontSize:10},type:'scroll'},grid:{left:80,right:30,top:50,bottom:50},xAxis:{type:'category',data:D.creatorNames,axisLabel:{fontSize:13}},yAxis:{type:'value',name:'工时(h)'},
      series:pSet.map(function(pj,pi){return{name:pj,type:'bar',data:D.creatorNames.map(function(n){var f=D.creatorProjTop3[D.creatorNames.indexOf(n)];var r=f?f.find(function(x){return x[0]===pj}):null;return r?r[1]:0}),itemStyle:{color:pColors[pi%pColors.length]},barWidth:14}})});
  }
  // 3-insight
  var pi=document.getElementById('person-insight');
  if(pi){var oc=D.creatorStd.filter(function(s){return s.status==='over'}).length,lc=D.creatorStd.filter(function(s){return s.status==='low'}).length;
    var tv=D.creatorValue.reduce(function(a,b){return a.highPct>b.highPct?a:b}),lv2=D.creatorValue.reduce(function(a,b){return a.highPct<b.highPct?a:b});
    pi.innerHTML='<strong>工时负荷：</strong>'+oc+'人超标（>'+std+'h），'+lc+'人不足（<134h）。'+D.creatorStd[0].name+'最高（'+D.creatorStd[0].actual+'h），'+D.creatorStd[D.creatorStd.length-1].name+'最低（'+D.creatorStd[D.creatorStd.length-1].actual+'h）。<br><strong>价值产出：</strong>'+tv.name+'高价值占比最高（'+tv.highPct+'%），'+lv2.name+'最低（'+lv2.highPct+'%）。<br><strong>建议：</strong>①超标人员适度分担；②提升低产出人员的价值占比；③关注项目分散度。';
  }
})();

// ===== 4. 产出价值 =====
(function(){
  var c1=initChart('chart-value-layer');
  if(c1)c1.setOption({title:{text:'工时价值分层',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:5,textStyle:{fontSize:11}},
    series:[{type:'pie',radius:['35%','68%'],center:['50%','50%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
      data:D.valueLayer.map(function(i){return{name:i[0],value:i[1],itemStyle:{color:i[2]}}}),
      label:{formatter:'{b}\n{d}%',fontSize:10}}]});
  var c2=initChart('chart-value-project');
  if(c2){var t10=D.top10Value,n=t10.map(function(i){return i[0]}).reverse(),v=t10.map(function(i){return i[1]}).reverse(),m=t10.map(function(i){return i[2]}).reverse(),o=t10.map(function(i){return i[3]}).reverse();
    c2.setOption({title:{text:'Top项目 价值产出比 vs 维护占比',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis'},grid:{left:140,right:30,top:50,bottom:40},xAxis:{type:'value',max:100,name:'%'},yAxis:{type:'category',data:n,axisLabel:{fontSize:11}},
      series:[{name:'价值产出比',type:'bar',stack:'a',data:v,itemStyle:{color:'#52C41A'},barWidth:18,label:{show:true,position:'inside',formatter:function(p){return p.value>8?p.value+'%':''},fontSize:10,color:'#fff'}},
               {name:'维护占比',type:'bar',stack:'a',data:m,itemStyle:{color:'#F7664A'},label:{show:true,position:'inside',formatter:function(p){return p.value>8?p.value+'%':''},fontSize:10,color:'#fff'}},
               {name:'其他',type:'bar',stack:'a',data:o,itemStyle:{color:'#D9D9D9'}}]});
  }
  var vi=document.getElementById('value-insight');
  if(vi){var op=pct(D.optimizable,D.maintTotal);
    vi.innerHTML='<strong>高价值产出（数字化升级+项目实施+售前）仅占'+D.highValuePct+'%</strong>，团队主要精力被维护和项目管理占据。<br><strong>可优化维护（业务咨询+开通）'+D.optimizable+'h（'+op+'%）</strong>，可通过自助化手段转化。<br><strong>项目价值产出比排名：</strong>'+D.top10Value.slice(0,3).map(function(x){return x[0]+x[1]+'%'}).join(' > ')+'。';
  }
})();

// ===== 5. 维护深度 =====
(function(){
  document.getElementById('maint-desc').textContent = '维护工时'+D.maintTotal+'h占比'+pct(D.maintTotal,D.total)+'%，从类型、产品线、原因、等级四个维度拆解';
  var c1=initChart('chart-maint-type');
  if(c1)c1.setOption({title:{text:'维护工时按类型分布',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:5,textStyle:{fontSize:11},type:'scroll'},
    series:[{type:'pie',radius:['35%','65%'],center:['50%','50%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
      data:D.maintType.map(function(i,idx){return{name:i[0],value:i[1],itemStyle:{color:MC[idx%MC.length]}}}),
      label:{formatter:'{b}\n{d}%',fontSize:10}}]});
  var c2=initChart('chart-maint-product');
  if(c2){var pl=D.prodLine.slice().sort(function(a,b){return a[1]-b[1]});
    c2.setOption({title:{text:'维护工时按产品线分布',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis'},grid:{left:160,right:40,top:50,bottom:30},xAxis:{type:'value'},yAxis:{type:'category',data:pl.map(function(i){return i[0]}),axisLabel:{fontSize:11}},
      series:[{type:'bar',data:pl.map(function(i){return i[1]}),barWidth:20,itemStyle:{borderRadius:[0,4,4,0],color:new echarts.graphic.LinearGradient(0,0,1,0,[{offset:0,color:'#FA8C16'},{offset:1,color:'#FFC53D'}])},label:{show:true,position:'right',formatter:'{c}h',fontSize:11}}]});
  }
  var c3=initChart('chart-maint-reason');
  if(c3){var nr=D.maintReason.find(function(i){return i[0]==='未填写'}),np=nr?pct(nr[1],D.maintTotal):'0';
    c3.setOption({title:{text:'维护原因分析',subtext:np+'%的维护记录未填写原因',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600},subtextStyle:{color:'#F5222D',fontSize:12}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:5,textStyle:{fontSize:11}},
      series:[{type:'pie',radius:['35%','65%'],center:['50%','52%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
        data:D.maintReason.map(function(i){return{name:i[0],value:i[1],itemStyle:{color:({'未填写':'#D9D9D9','客户直接询问':'#F7664A','没有运维团队':'#FA8C16','代理商/分支机构能力不足':'#4A6CF7'})[i[0]]||'#999'}}}),
        label:{formatter:'{b}\n{d}%',fontSize:10}}]});
  }
  var c4=initChart('chart-maint-level');
  if(c4){var lvC={'P0':'#F5222D','P1':'#FA8C16','P2':'#4A6CF7','P3':'#52C41A','普通级':'#D9D9D9','未分级':'#E8E8E8'};
    c4.setOption({title:{text:'维护工时按项目等级分布',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:10,textStyle:{fontSize:11}},
      series:[{type:'pie',radius:['35%','65%'],center:['50%','50%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
        data:D.maintLevel.map(function(i){return{name:i[0],value:i[1],itemStyle:{color:lvC[i[0]]||'#999'}}}),
        label:{formatter:'{b}\n{d}%',fontSize:10}}]});
  }
})();

// ===== 6. 维护细化 =====
(function(){
  var c1=initChart('chart-maint-class');
  if(c1){var cc={'问题处理':'#FA8C16','业务/功能咨询':'#F7664A','业务开通':'#722ED1','未分类':'#D9D9D9'};
    c1.setOption({title:{text:'维护工时价值分类',subtext:'基于内容描述关键词智能分类',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600},subtextStyle:{color:'#6B7280',fontSize:11}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:5,textStyle:{fontSize:11}},
      series:[{type:'pie',radius:['35%','68%'],center:['50%','50%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
        data:D.maintClass.map(function(i){return{name:i[0],value:i[1],itemStyle:{color:cc[i[0]]||'#999'}}}),
        label:{formatter:'{b}\n{d}%',fontSize:10}}]});
  }
  var c2=initChart('chart-gout-class');
  if(c2){var gt=D.goutClass.reduce(function(s,x){return s+x[1]},0);
    c2.setOption({title:{text:'沟通答疑内容细分',subtext:'真正的问题处理 vs 业务/功能咨询',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600},subtextStyle:{color:'#6B7280',fontSize:11}},tooltip:{trigger:'item',formatter:'{b}: {c}h ({d}%)'},legend:{bottom:5,textStyle:{fontSize:11}},
      series:[{type:'pie',radius:['35%','68%'],center:['50%','50%'],itemStyle:{borderRadius:5,borderColor:'#fff',borderWidth:2},
        data:D.goutClass.map(function(i){var p=gt>0?(i[1]/gt*100).toFixed(1):'0';return{name:i[0]+'('+p+'%)',value:i[1],itemStyle:{color:({'问题处理':'#FA8C16','业务/功能咨询':'#F7664A'})[i[0]]||'#999'}}}),
        label:{formatter:'{b}\n{c}h',fontSize:11}}]});
  }
  var c3=initChart('chart-maint-class-project');
  if(c3){var pmc=D.projMaintClass,names=pmc.map(function(i){return i.name}).reverse(),cls=['问题处理','业务/功能咨询','业务开通','未分类'],cc2=['#FA8C16','#F7664A','#722ED1','#D9D9D9'];
    c3.setOption({title:{text:'维护价值分类 × 项目 工时拆解',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{bottom:5,textStyle:{fontSize:11}},grid:{left:160,right:30,top:50,bottom:50},xAxis:{type:'value'},yAxis:{type:'category',data:names,axisLabel:{fontSize:11}},
      series:cls.map(function(c,idx){return{name:c,type:'bar',stack:'a',data:pmc.map(function(p){return p[c]||0}).reverse(),itemStyle:{color:cc2[idx]},barWidth:22}})});
  }
  var ic=document.getElementById('maint-insights');
  if(ic){var gp=D.goutClass.find(function(i){return i[0]==='问题处理'}),gc=D.goutClass.find(function(i){return i[0]==='业务/功能咨询'}),gu=D.maintClass.find(function(i){return i[0]==='未分类'}),gpt=gp?gp[1]:0,gct=gc?gc[1]:0,gut=gu?gu[1]:0,gt2=gpt+gct+gut;
    var mp2=D.maintClass.find(function(i){return i[0]==='问题处理'}),mc2=D.maintClass.find(function(i){return i[0]==='业务/功能咨询'}),mo2=D.maintClass.find(function(i){return i[0]==='业务开通'}),mu2=D.maintClass.find(function(i){return i[0]==='未分类'});
    ic.innerHTML='<div class="insight-card critical"><div class="insight-title">🔴 沟通答疑中'+(gct>0&&gt2>0?pct(gct,gt2):'0')+'%是业务/功能咨询</div><div class="insight-body">沟通答疑'+gt2.toFixed(1)+'h中：问题处理'+gpt.toFixed(1)+'h、业务咨询'+gct.toFixed(1)+'h、未分类'+gut.toFixed(1)+'h。<br><br>业务咨询部分可通过自助化手段压缩。</div></div>'
      +'<div class="insight-card critical"><div class="insight-title">🔴 维护工时中问题处理占'+(mp2?pct(mp2[1],D.maintTotal):'0')+'%</div><div class="insight-body">维护'+D.maintTotal+'h拆解：问题处理'+(mp2?mp2[1]:0)+'h、业务咨询'+(mc2?mc2[1]:0)+'h、业务开通'+(mo2?mo2[1]:0)+'h、未分类'+(mu2?mu2[1]:0)+'h。<br><br>业务咨询+开通='+((mc2?mc2[1]:0)+(mo2?mo2[1]:0))+'h，可通过自助化压缩。</div></div>'
      +'<div class="insight-card info"><div class="insight-title">🔵 维护原因记录不完整</div><div class="insight-body"><strong>维护记录中大量未填写维护原因</strong>，影响归因精准度，<strong>建议强制填写维护原因</strong>。</div></div>';
  }
})();

// ===== 7. 热力图 =====
(function(){
  var c=initChart('chart-maint-heatmap');
  if(c){c.setOption({title:{text:'维护类型 × 项目 工时热力图（h）',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{formatter:function(p){return D.hmProjs[p.data[1]]+' · '+D.hmTypes[p.data[0]]+'<br/>工时: '+p.data[2]+'h'}},grid:{left:160,right:80,top:50,bottom:40},
      xAxis:{type:'category',data:D.hmTypes,axisLabel:{fontSize:11,rotate:15}},yAxis:{type:'category',data:D.hmProjs,axisLabel:{fontSize:11}},
      visualMap:{min:0,max:40,inRange:{color:['#EBF5FF','#4A6CF7','#F7664A']},orient:'vertical',right:10,top:'center',text:['高','低'],textStyle:{fontSize:11}},
      series:[{type:'heatmap',data:D.hmData,label:{show:true,formatter:function(p){return p.data[2]>0?p.data[2]:''},fontSize:10},itemStyle:{borderColor:'#fff',borderWidth:2}}]});
  }
})();

// ===== 8. 每日趋势 =====
(function(){
  var c=initChart('chart-daily-trend');
  if(c){c.setOption({title:{text:'每日工时趋势（按工作类型）',left:'center',top:10,textStyle:{fontSize:14,fontWeight:600}},tooltip:{trigger:'axis'},legend:{data:['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前'],bottom:5,textStyle:{fontSize:11}},grid:{left:50,right:20,top:50,bottom:50},
      xAxis:{type:'category',data:D.dailyDates,axisLabel:{fontSize:10,rotate:30}},yAxis:{type:'value',name:'工时(h)'},
      series:[{name:'维护',type:'bar',stack:'t',data:D.dailyMaint,itemStyle:{color:'#F7664A'}},{name:'项目管理',type:'bar',stack:'t',data:D.dailyPM,itemStyle:{color:'#FA8C16'}},{name:'数字化升级',type:'bar',stack:'t',data:D.dailyDigital,itemStyle:{color:'#4A6CF7'}},{name:'项目实施',type:'bar',stack:'t',data:D.dailyImpl,itemStyle:{color:'#13C2C2'}},{name:'学习/内部例会',type:'bar',stack:'t',data:D.dailyLearn,itemStyle:{color:'#722ED1'}},{name:'售前',type:'bar',stack:'t',data:D.dailySale,itemStyle:{color:'#52C41A'}}]});
  }
})();

// ===== 9. 优化建议 =====
(function(){
  var oi=document.getElementById('optimize-insights');
  if(oi){var gc2=D.goutClass.find(function(i){return i[0]==='业务/功能咨询'});var gct2=gc2?gc2[1]:0;
    oi.innerHTML='<div class="insight-card info"><div class="insight-title">💡 建立客户自助服务体系</div><div class="insight-body">针对<strong>业务/功能咨询（'+gct2.toFixed(1)+'h）</strong>，建议建立FAQ知识库+视频教程，预计<strong>减少60-70%业务咨询工时</strong>。</div></div>'
      +'<div class="insight-card info"><div class="insight-title">💡 加强产品质量管控</div><div class="insight-body">针对<strong>Bug类维护</strong>，建议专项质量复盘，预计<strong>减少40-50%问题处理工时</strong>。</div></div>'
      +'<div class="insight-card info"><div class="insight-title">💡 推进配置标准化</div><div class="insight-body">针对<strong>设置调整+新开门店工时</strong>，建议建立配置模板库，预计<strong>减少50-60%相关工时</strong>。</div></div>';
  }
  var oe=document.getElementById('optimize-expect');
  if(oe){oe.innerHTML='以上优化预计可<strong>释放约'+Math.round(D.optimizable*0.65)+'h/月</strong>的维护工时，使维护占比从'+pct(D.maintTotal,D.total)+'%下降，<strong>高价值产出占比从'+D.highValuePct+'%提升</strong>。';
  }
})();

window.addEventListener('resize',function(){document.querySelectorAll('[echarts_instance_]').forEach(function(el){var i=echarts.getInstanceByDom(el);if(i)i.resize()})});
"""

# 替换日期占位符
import datetime
gen_date = datetime.datetime.now().strftime('%Y年%m月%d日')
chart_js = chart_js.replace('__DATE_RANGE__', date_range).replace('__GEN_DATE__', gen_date)

# 组装最终HTML：将图表JS插入到 </body> 前
final_html = html.replace('__CHART_JS_PLACEHOLDER__', chart_js)

with open('/Users/anyonghasayyo/Desktop/WorkBuddy/月报工时分析看板/dashboard.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print('✅ dashboard.html 生成成功（自包含，无需服务器）')
print('   总工时:', D['total'], 'h')
print('   负责人:', D['creatorNames'])
print('   项目数:', len(D['projNames']))
print('   HTML 大小:', len(final_html), '字符')
