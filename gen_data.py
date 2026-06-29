import pandas as pd
import json
import glob, os, re, time
from collections import defaultdict

# ===== 通用处理函数：输入Excel路径，输出数据字典 D =====
def process_excel(filepath):
    df = pd.read_excel(filepath, sheet_name='项目工时登记数据')

    # 列名兼容映射
    col_map = {}
    for c in ['工时','工作时长']:
        if c in df.columns: col_map['工时'] = c; break
    for c in ['负责人（必填）','创建人']:
        if c in df.columns: col_map['负责人'] = c; break
    for c in ['维护类型']:
        if c in df.columns: col_map['维护类型'] = c; break
    for c in ['项目等级']:
        if c in df.columns: col_map['项目等级'] = c; break
    for c in ['发生时间']:
        if c in df.columns: col_map['发生时间'] = c; break
    for c in ['项目名称']:
        if c in df.columns: col_map['项目名称'] = c; break
    for c in ['工作类型']:
        if c in df.columns: col_map['工作类型'] = c; break

    dept_col = '负责人主属部门' if '负责人主属部门' in df.columns else None

    # 数据清洗
    df[col_map['发生时间']] = pd.to_datetime(df[col_map['发生时间']], errors='coerce')
    df['日期'] = df[col_map['发生时间']].dt.strftime('%m-%d')
    df['工时'] = pd.to_numeric(df[col_map['工时']], errors='coerce').fillna(0)
    df['维护类型_fill'] = df[col_map['维护类型']].fillna('未填写')

    minDate = df[col_map['发生时间']].min().strftime('%Y-%m-%d')
    maxDate = df[col_map['发生时间']].max().strftime('%Y-%m-%d')

    def pct(a, b):
        return round(a/b*100, 1) if b > 0 else 0

    # 1. KPI
    total      = df['工时'].sum()
    records    = len(df)
    projCount  = df[col_map['项目名称']].dropna().nunique()
    workDates  = df[col_map['发生时间']].dropna()
    workDays   = workDates[workDates.dt.weekday < 5].dt.date.nunique()
    maintTotal = df[df[col_map['工作类型']]=='维护']['工时'].sum()
    highValue  = df[df[col_map['工作类型']].isin(['数字化升级','项目实施','售前','项目管理'])]['工时'].sum()
    highValuePct = pct(highValue, total)
    optimizable  = df[df['维护类型_fill'].isin(['沟通答疑','新开门店'])]['工时'].sum()
    p0_daily = df[df[col_map['项目等级']]=='P0'].groupby(df[col_map['发生时间']].dt.date)['工时'].sum()
    p0Daily  = round(p0_daily.mean(), 1) if len(p0_daily) > 0 else 0
    allDaily = round(df.groupby(df[col_map['发生时间']].dt.date)['工时'].sum().mean(), 1)

    # 2. 工作类型饼图
    typeColors = {'维护':'#F7764A','项目管理':'#FA8C16','数字化升级':'#4A6CF7',
                 '学习/内部例会':'#722ED1','项目实施':'#13C2C2','售前':'#52C41A'}
    typePie = [[wt, round(g['工时'].sum(),1)] for wt,g in df.groupby(col_map['工作类型'])]

    # 3. 项目维度
    projMap = {}
    for name, g in df.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        h  = round(g['工时'].sum(),1)
        m  = round(g[g[col_map['工作类型']]=='维护']['工时'].sum(),1)
        v  = round(g[g[col_map['工作类型']].isin(['数字化升级','项目实施','售前','项目管理'])]['工时'].sum(),1)
        lv = g[col_map['项目等级']].dropna().iloc[0] if g[col_map['项目等级']].dropna().shape[0]>0 else '普通级'
        projMap[name] = {'h':h,'m':m,'v':v,'lv':lv}
    projSorted   = sorted(projMap.items(), key=lambda x: -x[1]['h'])
    projNames    = [x[0] for x in projSorted]
    projHours    = [x[1]['h'] for x in projSorted]
    projMaint    = [x[1]['m'] for x in projSorted]
    projMaintPct = [pct(x[1]['m'],x[1]['h']) for x in projSorted]
    projValuePct = [pct(x[1]['v'],x[1]['h']) for x in projSorted]
    projLevel    = [x[1]['lv'] for x in projSorted]
    projValue    = [x[1]['v'] for x in projSorted]

    # 4. 人员维度
    creatorNames = []
    for name, g in df.groupby(col_map['负责人']):
        if pd.isna(name): continue
        creatorNames.append(name)
    workTypes = ['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前']
    creatorTotalHours  = [round(df[df[col_map['负责人']]==n]['工时'].sum(),1) for n in creatorNames]
    creatorTypeMatrix = []
    for n in creatorNames:
        g = df[df[col_map['负责人']]==n]
        creatorTypeMatrix.append([round(g[g[col_map['工作类型']]==wt]['工时'].sum(),1) for wt in workTypes])
    creatorValue = []
    for i,n in enumerate(creatorNames):
        t    = creatorTotalHours[i]
        high = creatorTypeMatrix[i][1] + creatorTypeMatrix[i][2] + creatorTypeMatrix[i][3] + creatorTypeMatrix[i][5]
        low  = creatorTypeMatrix[i][0]
        creatorValue.append({'name':n,'total':t,'high':round(high,1),'low':round(low,1),
                             'highPct':pct(high,t),'lowPct':pct(low,t)})
    stdHours = workDays * 8
    creatorStd = []
    for i,n in enumerate(creatorNames):
        a = creatorTotalHours[i]
        p = round(a/stdHours*100,1)
        dif = round(a-stdHours,1)
        creatorStd.append({'name':n,'actual':a,'std':stdHours,'pct':p,'diff':dif,
                           'status':'over' if p>100 else ('ok' if p>=80 else 'low')})
    creatorProjTop3 = []
    for n in creatorNames:
        g = df[df[col_map['负责人']]==n]
        top = g.groupby(col_map['项目名称'])['工时'].sum().sort_values(ascending=False).head(3)
        creatorProjTop3.append([[nm,round(hh,1)] for nm,hh in top.items()])
    # 动态维护类型列表
    maintTypeList = []
    for mt in df[df[col_map['工作类型']]=='维护'][col_map['维护类型']].fillna('未填写'):
        if mt not in maintTypeList:
            maintTypeList.append(mt)
    creatorMaintMatrix = []
    for n in creatorNames:
        g = df[(df[col_map['负责人']]==n)&(df[col_map['工作类型']]=='维护')]
        creatorMaintMatrix.append([round(g[g['维护类型_fill']==mt]['工时'].sum(),1) for mt in maintTypeList])
    levelOrder = ['P0','P1','P2','P3','普通级']
    creatorLevelMatrix = []
    for n in creatorNames:
        g = df[df[col_map['负责人']]==n]
        creatorLevelMatrix.append([round(g[g[col_map['项目等级']]==lv]['工时'].sum(),1) for lv in levelOrder])

    # 4b. 部门维度
    if dept_col:
        deptNames = []
        for name, g in df.groupby(dept_col):
            if pd.isna(name): continue
            if name not in deptNames: deptNames.append(name)
        deptTotalHours  = [round(df[df[dept_col]==d]['工时'].sum(),1) for d in deptNames]
        deptTypeMatrix = []
        for d in deptNames:
            g = df[df[dept_col]==d]
            deptTypeMatrix.append([round(g[g[col_map['工作类型']]==wt]['工时'].sum(),1) for wt in workTypes])
        deptValue = []
        for i,d in enumerate(deptNames):
            t    = deptTotalHours[i]
            high = deptTypeMatrix[i][1] + deptTypeMatrix[i][2] + deptTypeMatrix[i][3] + deptTypeMatrix[i][5]
            low  = deptTypeMatrix[i][0]
            deptValue.append({'name':d,'total':t,'high':round(high,1),'low':round(low,1),
                              'highPct':pct(high,t),'lowPct':pct(low,t)})
        deptMaintMatrix = []
        for d in deptNames:
            g = df[(df[dept_col]==d)&(df[col_map['工作类型']]=='维护')]
            deptMaintMatrix.append([round(g[g['维护类型_fill']==mt]['工时'].sum(),1) for mt in maintTypeList])
        deptLevelMatrix = []
        for d in deptNames:
            g = df[df[dept_col]==d]
            deptLevelMatrix.append([round(g[g[col_map['项目等级']]==lv]['工时'].sum(),1) for lv in levelOrder])
        deptProjTop3 = []
        for d in deptNames:
            g = df[df[dept_col]==d]
            top = g.groupby(col_map['项目名称'])['工时'].sum().sort_values(ascending=False).head(3)
            deptProjTop3.append([[nm,round(hh,1)] for nm,hh in top.items()])
    else:
        deptNames = []; deptTotalHours = []; deptTypeMatrix = []; deptValue = []
        deptMaintMatrix = []; deptLevelMatrix = []; deptProjTop3 = []

    # 5. 项目管理工时
    pmByProject = [[n, round(df[(df[col_map['工作类型']]=='项目管理')&(df[col_map['项目名称']]==n)]['工时'].sum(),1)]
                   for n in df[df[col_map['工作类型']]=='项目管理'][col_map['项目名称']].dropna().unique()]
    pmByProject = sorted(pmByProject, key=lambda x:-x[1])
    pmByCreator = [[n, round(df[(df[col_map['工作类型']]=='项目管理')&(df[col_map['负责人']]==n)]['工时'].sum(),1)]
                    for n in df[df[col_map['工作类型']]=='项目管理'][col_map['负责人']].dropna().unique()]
    pmByCreator = sorted(pmByCreator, key=lambda x:-x[1])
    pmTotal = round(df[df[col_map['工作类型']]=='项目管理']['工时'].sum(),1)

    # 6. 维护类型
    maintType   = [[mt,round(g['工时'].sum(),1)] for mt,g in df[df[col_map['工作类型']]=='维护'].groupby('维护类型_fill')]
    biz_cols   = ['餐饮产品事业部','风行产品事业部','龙决策产品事业部','吾享产品事业部','快餐产品事业部','总部报表']
    prodLineMap = defaultdict(float)
    for _,row in df[df[col_map['工作类型']]=='维护'].iterrows():
        h = row['工时']
        for col in biz_cols:
            v = row[col]
            if pd.notna(v) and str(v).strip():
                prodLineMap[str(v).strip()] += h
    prodLine = [[k,round(v,1)] for k,v in prodLineMap.items()]
    maintReason = []
    for r,g in df[df[col_map['工作类型']]=='维护'].groupby('维护类型_fill'):
        maintReason.append([r, round(g['工时'].sum(),1)])
    mrMap = {}
    for k,v in maintReason: mrMap[k] = mrMap.get(k,0)+v
    maintReason = [[k,round(v,1)] for k,v in mrMap.items()]
    maintLevel = [[lv,round(g['工时'].sum(),1)] for lv,g in df[df[col_map['工作类型']]=='维护'].groupby(col_map['项目等级']) if pd.notna(lv)]
    maintClass = []
    for name, g in df[df[col_map['工作类型']]=='维护'].groupby('维护类型_fill'):
        maintClass.append([name, round(g['工时'].sum(),1)])
    goutClass = []
    for name, g in df[df['维护类型_fill']=='沟通答疑'].groupby('维护类型_fill'):
        goutClass.append([name, round(g['工时'].sum(),1)])
    projMaintClass = []
    for name,g in df[df[col_map['工作类型']]=='维护'].groupby(col_map['项目名称']):
        if pd.isna(name): continue
        d = {'name': name}
        for mt in df['维护类型_fill'].drop_duplicates():
            d[mt] = round(g[g['维护类型_fill']==mt]['工时'].sum(),1)
        projMaintClass.append(d)

    # 7. 热力图
    hmProjs  = [x[0] for x in sorted(zip(projNames,projMaint), key=lambda x:-x[1])[:15]]
    hmTypes  = [mt[0] for mt in maintType if mt[0]!='未填写']
    hmTypes  = list(set(hmTypes))[:6]
    hmData    = []
    for pi,pn in enumerate(hmProjs):
        for ti,tn in enumerate(hmTypes):
            h = df[(df[col_map['项目名称']]==pn)&(df[col_map['工作类型']]=='维护')&(df['维护类型_fill']==tn)]['工时'].sum()
            if h>0: hmData.append([ti,pi,round(h,1)])

    # 8. 每日趋势
    dailyMap = defaultdict(lambda: defaultdict(float))
    for _,row in df.iterrows():
        d  = row['日期']
        wt = row[col_map['工作类型']]
        h  = row['工时']
        if pd.notna(d): dailyMap[d][wt] += h
    allDates    = sorted(dailyMap.keys())
    dailyDates  = allDates
    dailyMaint  = [round(dailyMap[d].get('维护',0),1) for d in allDates]
    dailyPM     = [round(dailyMap[d].get('项目管理',0),1) for d in allDates]
    dailyDigital = [round(dailyMap[d].get('数字化升级',0),1) for d in allDates]
    dailyImpl  = [round(dailyMap[d].get('项目实施',0),1) for d in allDates]
    dailyLearn  = [round(dailyMap[d].get('学习/内部例会',0),1) for d in allDates]
    dailySale  = [round(dailyMap[d].get('售前',0),1) for d in allDates]

    # 9. 价值分层
    valueLayer = [
        ['高价值产出(数字化+实施+售前+项目管理)', round(highValue,1), '#52C41A'],
        ['学习/内部例会', round(df[df[col_map['工作类型']].isin(['学习/内部例会'])]['工时'].sum(),1), '#4A6CF7'],
    ]
    for name, g in df[df[col_map['工作类型']]=='维护'].groupby('维护类型_fill'):
        vc = '#FA8C16' if 'bug' in str(name).lower() else '#F7764A'
        valueLayer.append([name, round(g['工时'].sum(),1), vc])
    other = df[~df[col_map['工作类型']].isin(['维护','项目管理','数字化升级','项目实施','学习/内部例会','售前'])]['工时'].sum()
    valueLayer.append(['其他', round(other,1), '#D9D9D9'])
    top10Value = sorted([[n,round(projValuePct[i],1),round(projMaintPct[i],1),round(100-projValuePct[i]-projMaintPct[i],1)] for i,n in enumerate(projNames)], key=lambda x:-x[1])[:10]

    return dict(
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
        deptNames=deptNames, deptTotalHours=deptTotalHours,
        deptTypeMatrix=deptTypeMatrix, deptValue=deptValue,
        deptMaintMatrix=deptMaintMatrix, deptLevelMatrix=deptLevelMatrix,
        deptProjTop3=deptProjTop3,
    )

# ===== 文件路径 =====
CURR_FILE  = '1-5月份项目工时登记对象导出结果.xlsx'   # 本月文件
PREV_FILE  = None                                          # 上月文件（暂无）

if not os.path.exists(CURR_FILE):
    raise FileNotFoundError(f'本月文件不存在: {CURR_FILE}')

print(f'处理本月文件：{CURR_FILE}')
D = process_excel(CURR_FILE)
print(f'  总工时: {round(D["total"],1)}h  记录数: {D["records"]}  工作日: {D["workDays"]}天')

D_prev = None
if PREV_FILE and os.path.exists(PREV_FILE):
    print(f'处理上月文件：{PREV_FILE}')
    D_prev = process_excel(PREV_FILE)
    print(f'  总工时: {round(D_prev["total"],1)}h  记录数: {D_prev["records"]}  工作日: {D_prev["workDays"]}天')
else:
    print('  上月文件未配置，仅生成本月数据')

# ===== 输出 data.js =====
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

with open('data.js', 'w', encoding='utf-8') as fout:
    fout.write('var D = ' + to_js(D) + ';\n')
    if D_prev is not None:
        fout.write('var D_prev = ' + to_js(D_prev) + ';\n')
    else:
        fout.write('var D_prev = null;\n')

print('data.js 生成成功')
if D_prev is not None:
    print('  （含本月+上月数据）')
else:
    print('  （仅本月数据，D_prev=null）')

# ===== 自动更新 dashboard.html 缓存时间戳 =====
TIMESTAMP = int(time.time())
html_path = 'dashboard.html'
if os.path.exists(html_path):
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()
    if 'data.js?v=' in html:
        new_html = re.sub(r'data\.js\?v=\d+', f'data.js?v={TIMESTAMP}', html)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_html)
        print(f'  dashboard.html 缓存时间戳已更新 (v={TIMESTAMP})')
    else:
        print('  警告：未在 dashboard.html 中找到 data.js?v= 引用')
