import pandas as pd
import json
from collections import defaultdict

f = '项目工时登记对象导出结果.xlsx'
df = pd.read_excel(f, sheet_name='项目工时登记数据')
df['发生时间'] = pd.to_datetime(df['发生时间'], errors='coerce')
df['日期'] = df['发生时间'].dt.strftime('%m-%d')
df['工时'] = pd.to_numeric(df['工作时长'], errors='coerce').fillna(0)
minDate = df['发生时间'].min().strftime('%Y-%m-%d')
maxDate = df['发生时间'].max().strftime('%Y-%m-%d')

def pct(a, b):
    return round(a/b*100, 1) if b > 0 else 0

# ===== 1. KPI =====
total        = df['工时'].sum()
records       = len(df)
projCount    = df['项目名称'].nunique()
workDays     = df['发生时间'].dropna().dt.date.nunique()
maintTotal   = df[df['工作类型']=='维护']['工时'].sum()
highValue    = df[df['工作类型'].isin(['数字化升级','项目实施','售前','项目管理'])]['工时'].sum()
highValuePct = pct(highValue, total)

# 使用维护类型_fill 直接计算 optimizable
df['维护类型_fill'] = df['维护类型'].fillna('未填写')
optimizable = df[df['维护类型_fill'].isin(['沟通答疑','新开门店/项目'])]['工时'].sum()

p0_daily = df[df['项目等级']=='P0'].groupby(df['发生时间'].dt.date)['工时'].sum()
p0Daily  = round(p0_daily.mean(), 1) if len(p0_daily) > 0 else 0
allDaily = round(df.groupby(df['发生时间'].dt.date)['工时'].sum().mean(), 1)

# ===== 2. 工作类型饼图 =====
typeColors = {'维护':'#F7664A','项目管理':'#FA8C16','数字化升级':'#4A6CF7',
              '学习/内部例会':'#722ED1','项目实施':'#13C2C2','售前':'#52C41A'}
typePie = [[wt, round(g['工时'].sum(),1)] for wt,g in df.groupby('工作类型')]

# ===== 3. 项目维度 =====
projMap = {}
for name, g in df.groupby('项目名称'):
    if pd.isna(name): continue
    h   = round(g['工时'].sum(), 1)
    m   = round(g[g['工作类型']=='维护']['工时'].sum(), 1)
    v   = round(g[g['工作类型'].isin(['数字化升级','项目实施','售前','项目管理'])]['工时'].sum(), 1)
    lv  = g['项目等级'].dropna().iloc[0] if g['项目等级'].dropna().shape[0]>0 else '未分级'
    projMap[name] = {'h':h,'m':m,'v':v,'lv':lv}

projSorted = sorted(projMap.items(), key=lambda x: -x[1]['h'])
projNames    = [x[0] for x in projSorted]
projHours    = [x[1]['h'] for x in projSorted]
projMaint    = [x[1]['m'] for x in projSorted]
projMaintPct = [pct(x[1]['m'],x[1]['h']) for x in projSorted]
projValuePct = [pct(x[1]['v'],x[1]['h']) for x in projSorted]
projLevel    = [x[1]['lv'] for x in projSorted]
projValue    = [x[1]['v'] for x in projSorted]

# ===== 4. 人员维度 =====
creatorNames = []
for name, g in df.groupby('负责人（必填）'):
    if pd.isna(name): continue
    creatorNames.append(name)

workTypes = ['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前']
creatorTotalHours  = [round(df[df['负责人（必填）']==n]['工时'].sum(),1) for n in creatorNames]
creatorTypeMatrix = []
for n in creatorNames:
    g = df[df['负责人（必填）']==n]
    creatorTypeMatrix.append([round(g[g['工作类型']==wt]['工时'].sum(),1) for wt in workTypes])

creatorValue = []
for i,n in enumerate(creatorNames):
    t     = creatorTotalHours[i]
    high  = creatorTypeMatrix[i][1] + creatorTypeMatrix[i][2] + creatorTypeMatrix[i][3] + creatorTypeMatrix[i][5]
    low   = creatorTypeMatrix[i][0]
    creatorValue.append({'name':n,'total':t,'high':round(high,1),'low':round(low,1),
                        'highPct':pct(high,t),'lowPct':pct(low,t)})

stdHours = 168
creatorStd = []
for i,n in enumerate(creatorNames):
    a   = creatorTotalHours[i]
    p   = round(a/stdHours*100,1)
    dif = round(a-stdHours,1)
    creatorStd.append({'name':n,'actual':a,'std':stdHours,'pct':p,'diff':dif,
                       'status':'over' if p>100 else ('ok' if p>=80 else 'low')})

creatorProjTop3 = []
for n in creatorNames:
    g = df[df['负责人（必填）']==n]
    top = g.groupby('项目名称')['工时'].sum().sort_values(ascending=False).head(3)
    creatorProjTop3.append([[nm,round(hh,1)] for nm,hh in top.items()])

# 动态获取维护类型列表（去重）
maintTypeList = []
for mt in df[df['工作类型']=='维护']['维护类型'].fillna('未填写'):
    if mt not in maintTypeList:
        maintTypeList.append(mt)

creatorMaintMatrix = []
for n in creatorNames:
    g = df[(df['负责人（必填）']==n)&(df['工作类型']=='维护')]
    creatorMaintMatrix.append([round(g[g['维护类型'].fillna('未填写')==mt]['工时'].sum(),1) for mt in maintTypeList])

levelOrder = ['P0','P1','P2','P3','普通级']
creatorLevelMatrix = []
for n in creatorNames:
    g = df[df['负责人（必填）']==n]
    creatorLevelMatrix.append([round(g[g['项目等级']==lv]['工时'].sum(),1) for lv in levelOrder])

# ===== 5. 项目管理工时 =====
pmByProject = [[n, round(df[(df['工作类型']=='项目管理')&(df['项目名称']==n)]['工时'].sum(),1)]
               for n in df[df['工作类型']=='项目管理']['项目名称'].dropna().unique()]
pmByProject = sorted(pmByProject, key=lambda x:-x[1])

pmByCreator = [[n, round(df[(df['工作类型']=='项目管理')&(df['负责人（必填）']==n)]['工时'].sum(),1)]
                for n in df[df['工作类型']=='项目管理']['负责人（必填）'].dropna().unique()]
pmByCreator = sorted(pmByCreator, key=lambda x:-x[1])
pmTotal = round(df[df['工作类型']=='项目管理']['工时'].sum(), 1)

# ===== 6. 维护类型 =====
maintType = [[mt,round(g['工时'].sum(),1)] for mt,g in df[df['工作类型']=='维护'].groupby('维护类型_fill')]

biz_cols = ['餐饮产品事业部','风行产品事业部','龙决策产品事业部','吾享产品事业部','快餐产品事业部','总部报表']
prodLineMap = defaultdict(float)
for _,row in df[df['工作类型']=='维护'].iterrows():
    h = row['工时']
    for col in biz_cols:
        v = row[col]
        if pd.notna(v) and str(v).strip():
            prodLineMap[str(v).strip()] += h
prodLine = [[k,round(v,1)] for k,v in prodLineMap.items()]

# 维护原因：直接用维护类型_fill
maintReason = []
for r,g in df[df['工作类型']=='维护'].groupby('维护类型_fill'):
    maintReason.append([r, round(g['工时'].sum(),1)])
# 合并相同类型
mrMap = {}
for k,v in maintReason:
    mrMap[k] = mrMap.get(k,0)+v
maintReason = [[k,round(v,1)] for k,v in mrMap.items()]

maintLevel = [[lv,round(g['工时'].sum(),1)] for lv,g in df[df['工作类型']=='维护'].groupby('项目等级') if pd.notna(lv)]
uncL = df[(df['工作类型']=='维护')&(df['项目等级'].isna())]['工时'].sum()
if uncL>0: maintLevel.append(['未分级',round(uncL,1)])

# ===== 维护细化分类：直接使用维护类型_fill =====
maintClass = []
for name, g in df[df['工作类型']=='维护'].groupby('维护类型_fill'):
    maintClass.append([name, round(g['工时'].sum(),1)])

# 沟通答疑内容细分：直接使用维护类型_fill
goutClass = []
for name, g in df[df['维护类型_fill']=='沟通答疑'].groupby('维护类型_fill'):
    goutClass.append([name, round(g['工时'].sum(),1)])

projMaintClass = []
for name,g in df[df['工作类型']=='维护'].groupby('项目名称'):
    if pd.isna(name): continue
    d = {'name': name}
    for mt in df['维护类型_fill'].drop_duplicates():
        d[mt] = round(g[g['维护类型_fill']==mt]['工时'].sum(),1)
    projMaintClass.append(d)

# ===== 7. 热力图 =====
hmProjs = [x[0] for x in sorted(zip(projNames,projMaint), key=lambda x:-x[1])[:15]]
hmTypes = [mt[0] for mt in maintType if mt[0]!='未填写']
hmTypes = list(set(hmTypes))[:6]
hmData   = []
for pi,pn in enumerate(hmProjs):
    for ti,tn in enumerate(hmTypes):
        h = df[(df['项目名称']==pn)&(df['工作类型']=='维护')&(df['维护类型_fill']==tn)]['工时'].sum()
        if h>0: hmData.append([ti,pi,round(h,1)])

# ===== 8. 每日趋势 =====
dailyMap = defaultdict(lambda: defaultdict(float))
for _,row in df.iterrows():
    d  = row['日期']
    wt = row['工作类型']
    h  = row['工时']
    if pd.notna(d): dailyMap[d][wt] += h

allDates    = sorted(dailyMap.keys())
dailyDates  = allDates
dailyMaint  = [round(dailyMap[d].get('维护',0),1) for d in allDates]
dailyPM     = [round(dailyMap[d].get('项目管理',0),1) for d in allDates]
dailyDigital = [round(dailyMap[d].get('数字化升级',0),1) for d in allDates]
dailyImpl  = [round(dailyMap[d].get('项目实施',0),1) for d in allDates]
dailyLearn = [round(dailyMap[d].get('学习/内部例会',0),1) for d in allDates]
dailySale  = [round(dailyMap[d].get('售前',0),1) for d in allDates]

# ===== 价值分层：使用维护类型_fill =====
valueLayer = [
    ['高价值产出(数字化+实施+售前+项目管理)', round(highValue,1), '#52C41A'],
    ['学习/内部例会', round(df[df['工作类型'].isin(['学习/内部例会'])]['工时'].sum(),1), '#4A6CF7'],
]
# 按维护类型细分
for name, g in df[df['工作类型']=='维护'].groupby('维护类型_fill'):
    vc = '#FA8C16' if 'bug' in str(name).lower() else '#F7664A'
    valueLayer.append([name, round(g['工时'].sum(),1), vc])
# 其他
other = df[~df['工作类型'].isin(['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前'])]['工时'].sum()
valueLayer.append(['其他', round(other,1), '#D9D9D9'])

top10Value = sorted([[n,round(projValuePct[i],1),round(projMaintPct[i],1),round(100-projValuePct[i]-projMaintPct[i],1)] for i,n in enumerate(projNames)], key=lambda x:-x[1])[:10]

# ===== 输出 data.js =====
D = dict(
    total=round(total,1), records=records, projCount=projCount, workDays=workDays,
    maintTotal=round(maintTotal,1), p0Daily=p0Daily, allDaily=allDaily,
    highValue=round(highValue,1), highValuePct=highValuePct, optimizable=round(optimizable,1),
    typePie=typePie, typeColors=typeColors,
    projNames=projNames, projHours=projHours, projMaint=projMaint,
    projMaintPct=projMaintPct, projValuePct=projValuePct, projLevel=projLevel,
    projValue=projValue,
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
    minDate=minDate, maxDate=maxDate, pmTotal=pmTotal,
    pmByProject=pmByProject, pmByCreator=pmByCreator,
    levelOrder=levelOrder, creatorLevelMatrix=creatorLevelMatrix,
    maintTypeList=maintTypeList, creatorMaintMatrix=creatorMaintMatrix,
)

def to_js(obj, indent=0):
    sp  = '  ' * indent
    sp1 = '  ' * (indent+1)
    if isinstance(obj, dict):
        items = []
        for k,v in obj.items():
            items.append(sp1 + json.dumps(k, ensure_ascii=False) + ': ' + to_js(v, indent+1))
        return '{\n' + ',\n'.join(items) + '\n' + sp + '}'
    elif isinstance(obj, list):
        if len(obj)==0: return '[]'
        if isinstance(obj[0], (list,tuple)):
            items = [sp1 + to_js(x, indent+1) for x in obj]
            return '[\n' + ',\n'.join(items) + '\n' + sp + ']'
        else:
            items = [sp1 + to_js(x, indent+1) for x in obj]
            return '[\n' + ',\n'.join(items) + '\n' + sp + ']'
    elif isinstance(obj, str):
        return json.dumps(obj, ensure_ascii=False)
    elif isinstance(obj, bool):
        return 'true' if obj else 'false'
    elif isinstance(obj, float):
        return str(obj)
    elif isinstance(obj, int):
        return str(obj)
    else:
        return json.dumps(obj, ensure_ascii=False)

js_str = 'var D = ' + to_js(D) + ';'
with open('data.js', 'w', encoding='utf-8') as fout:
    fout.write(js_str)

print('data.js 生成成功')
print('  总工时:', round(total,1), 'h')
print('  记录数:', records)
print('  负责人:', creatorNames)
print('  项目数:', len(projNames))
