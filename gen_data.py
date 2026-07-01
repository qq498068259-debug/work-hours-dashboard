import pandas as pd
import json
import glob, os, re, time
from collections import defaultdict

# 价值工时权重
TYPE_WEIGHT = {'数字化升级':5,'项目管理':4,'项目实施':3,'售前':2,'维护':1,'学习/内部例会':0}
LEVEL_WEIGHT = {'P0':4,'P1':3,'P2':2,'P3':1,'普通级':1}

# 2026年中国法定节假日（影响工作日的日期范围）
HOLIDAYS_2026 = {
    '元旦':     ('2026-01-01', '2026-01-03'),
    '春节':     ('2026-02-15', '2026-02-21'),
    '清明节':   ('2026-04-04', '2026-04-06'),
    '劳动节':   ('2026-05-01', '2026-05-05'),
    '端午节':   ('2026-06-19', '2026-06-21'),
}

# 已转出/协助项目（已不再由本团队维护）
TRANSFERRED_PROJECTS = ['八合里牛肉火锅']

def calc_value_work(h, wt, lv):
    tw = TYPE_WEIGHT.get(wt, 1)
    lw = LEVEL_WEIGHT.get(lv, 2)  # 空值默认2
    return round(h * tw * lw, 1)

def value_grade(avg_factor):
    if avg_factor >= 4.0: return 'S'
    if avg_factor >= 3.0: return 'A'
    if avg_factor >= 2.0: return 'B'
    if avg_factor >= 1.5: return 'C'
    return 'D'

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

    # 计算每行的价值工时
    def calc_row(row):
        return calc_value_work(row['工时'], row[col_map['工作类型']], row[col_map['项目等级']])
    df['价值工时'] = df.apply(calc_row, axis=1)
    df['月份'] = df[col_map['发生时间']].dt.month  # 1~12

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
    totalValue = df['价值工时'].sum()

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
        v  = round(g['价值工时'].sum(),1)
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
    # 价值工时（人员）
    creatorValue = []
    creatorValue2 = []
    for i,n in enumerate(creatorNames):
        t    = creatorTotalHours[i]
        vw   = round(df[df[col_map['负责人']]==n]['价值工时'].sum(),1)
        avgf = round(vw/t, 2) if t > 0 else 0
        grade = value_grade(avgf)
        creatorValue.append({'name':n,'total':t,'high':0,'low':0,
                             'highPct':0,'lowPct':0})
        creatorValue2.append({'name':n,'total':t,'valueWork':vw,'avgFactor':avgf,'grade':grade})
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
        creatorProjTop3.append([[nm,round(h,1)] for nm,hh in top.items()])
    # 动态维护类型列表
    maintTypeList = []
    for mt in df[df[col_map['工作类型']]=='维护']['维护类型_fill'].dropna().unique():
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
            vw   = round(df[df[dept_col]==d]['价值工时'].sum(),1)
            avgf = round(vw/t,2) if t > 0 else 0
            grade = value_grade(avgf)
            deptValue.append({'name':d,'total':t,'valueWork':vw,'avgFactor':avgf,'grade':grade})
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
            deptProjTop3.append([[nm,round(h,1)] for nm,hh in top.items()])
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

    # 月度聚合数据（用于月度优化建议）
    from datetime import datetime, timedelta
    monthlyHours = defaultdict(float)
    monthlyUniqueDays = defaultdict(set)
    for _, row in df.iterrows():
        d = row[col_map['发生时间']]
        if pd.isna(d): continue
        m = d.strftime('%Y-%m')
        monthlyHours[m] += row['工时']
        if d.weekday() < 5:
            monthlyUniqueDays[m].add(d.date())

    # 每月节假日天数（只计算落在工作日的）
    monthlyHolidays = defaultdict(int)
    for name, (start_str, end_str) in HOLIDAYS_2026.items():
        start = datetime.strptime(start_str, '%Y-%m-%d')
        end = datetime.strptime(end_str, '%Y-%m-%d')
        current = start
        while current <= end:
            if current.weekday() < 5:
                m = current.strftime('%Y-%m')
                monthlyHolidays[m] += 1
            current += timedelta(days=1)

    # 整理月度数据
    monthLabels = ['1月','2月','3月','4月','5月','6月']
    monthlyData = []
    for i, m in enumerate(['2026-01','2026-02','2026-03','2026-04','2026-05','2026-06']):
        h = round(monthlyHours.get(m, 0), 1)
        wd = len(monthlyUniqueDays.get(m, set()))
        hd = monthlyHolidays.get(m, 0)
        daily_avg = round(h / wd, 1) if wd > 0 else 0
        monthlyData.append({
            'label': monthLabels[i],
            'hours': h,
            'workDays': wd,
            'holidayDays': hd,
            'dailyAvg': daily_avg
        })
    # 项目是否转出标记
    transferredFlags = {}
    for proj_name in projNames:
        transferredFlags[proj_name] = proj_name in TRANSFERRED_PROJECTS

    # 9.5 售前板块（新增）
    preDf = df[df[col_map['工作类型']] == '售前'].copy()
    preSaleTotal = round(preDf['工时'].sum(), 1)
    # 按项目汇总售前工时
    preProjMap = {}
    for name, g in preDf.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        h = round(g['工时'].sum(), 1)
        # 项目经理：该项目售前工时最多的负责人
        creatorHours = g.groupby(col_map['负责人'])['工时'].sum()
        pm = creatorHours.idxmax() if not creatorHours.empty else ''
        preProjMap[name] = {'hours': h, 'pm': pm, 'pct': 0}
    # 计算占比
    if preSaleTotal > 0:
        for k in preProjMap:
            preProjMap[k]['pct'] = round(preProjMap[k]['hours'] / preSaleTotal * 100, 1)
    preSaleProj = sorted([[k, v['pm'], v['hours'], v['pct']] for k,v in preProjMap.items()], key=lambda x:-x[2])
    preSalePie  = [[k, v['hours']] for k,v in preProjMap.items()]
    preSalePie  = sorted(preSalePie, key=lambda x:-x[1])
    # 售前人员分布
    preCreatorMap = {}
    for _,row in preDf.iterrows():
        c = row[col_map['负责人']]
        h = row['工时']
        preCreatorMap[c] = preCreatorMap.get(c, 0) + h
    preSaleCreator = [[k, round(v,1)] for k,v in preCreatorMap.items()]
    preSaleCreator = sorted(preSaleCreator, key=lambda x:-x[1])

    # 9. 价值工时分析（新增）
    # 项目价值排名
    projValueMap = {}
    for name, g in df.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        t = round(g['工时'].sum(),1)
        vw = round(g['价值工时'].sum(),1)
        avgf = round(vw/t,2) if t > 0 else 0
        lv = g[col_map['项目等级']].dropna().iloc[0] if g[col_map['项目等级']].dropna().shape[0]>0 else '普通级'
        projValueMap[name] = {'name':name,'total':t,'valueWork':vw,'avgFactor':avgf,'grade':value_grade(avgf),'lv':lv}
    projValue2 = sorted(projValueMap.values(), key=lambda x:-x['valueWork'])

    # 类型价值分布
    typeValue = [[wt, round(df[df[col_map['工作类型']]==wt]['价值工时'].sum(),1)] for wt in workTypes if df[df[col_map['工作类型']]==wt]['价值工时'].sum() > 0]

    # 等级价值分布
    levelValue = [[lv, round(df[df[col_map['项目等级']]==lv]['价值工时'].sum(),1)] for lv in levelOrder if lv in df[col_map['项目等级']].values]

    # 10. 数字化升级板块
    digitalUpgrade = []
    digital_df = df[df[col_map['工作类型']]=='数字化升级'].copy()
    digitalTypeCol = '数字化升级'  # Excel中数字化升级内容字段
    for name, g in digital_df.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        total_h = round(g['工时'].sum(), 1)
        # 项目经理：该项目数字化升级工时最多的负责人
        creatorHours = g.groupby(col_map['负责人'])['工时'].sum()
        pm = creatorHours.idxmax() if not creatorHours.empty else ''
        digitalUpgrade.append({'project': name, 'pm': pm, 'total': total_h})
    digitalUpgrade.sort(key=lambda x: -x['total'])
    digitalTotal = round(digital_df['工时'].sum(), 1)

    # 数字化升级饼图数据：项目+类型组合
    digitalPieData = []
    for name, g in digital_df.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        for dt, gg in g.groupby(digitalTypeCol):
            dt_label = str(dt) if pd.notna(dt) else '未分类'
            h = round(gg['工时'].sum(), 1)
            digitalPieData.append({'name': str(name) + '-' + dt_label, 'project': str(name), 'type': dt_label, 'hours': h})
    digitalPieData.sort(key=lambda x: -x['hours'])

    # 11. 高价值工时明细（项目+类型+工时）
    highValueTypes = ['数字化升级','项目管理','项目实施','售前']
    highValueDetail = []
    highValueDf = df[df[col_map['工作类型']].isin(highValueTypes)].copy()
    for name, g in highValueDf.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        total_h = round(g['工时'].sum(), 1)
        types = []
        for wt, gg in g.groupby(col_map['工作类型']):
            types.append({'type': wt, 'hours': round(gg['工时'].sum(), 1)})
        types.sort(key=lambda x: -x['hours'])
        highValueDetail.append({'project': name, 'total': total_h, 'types': types})
    highValueDetail.sort(key=lambda x: -x['total'])
    highValueDetailTotal = round(highValueDf['工时'].sum(), 1)

    # ===== 员工成长分析 =====
    months = sorted([m for m in df['月份'].unique() if m <= 6])  # 只取上半年月份
    monthLabels = [f'{m}月' for m in months]
    creatorsGrowth = []
    allCreatorNames = df[col_map['负责人']].dropna().unique()
    for name in allCreatorNames:
        personDf = df[df[col_map['负责人']] == name]
        monthlyHours = []
        monthlyTypes = []
        monthlyValue = []
        for m in months:
            md = personDf[personDf['月份'] == m]
            monthlyHours.append(round(md['工时'].sum(), 1))
            monthlyTypes.append(md[col_map['工作类型']].nunique())
            monthlyValue.append(round(md['价值工时'].sum(), 1))
        # 成长趋势判定
        activeMonths = [i for i, h in enumerate(monthlyHours) if h > 0]
        if len(activeMonths) >= 3:
            first3 = sum(monthlyHours[activeMonths[0]:activeMonths[0]+3]) if activeMonths[0]+3 <= len(monthlyHours) else sum(monthlyHours[activeMonths[0]:])
            recent3 = sum(monthlyHours[activeMonths[-1]-2:activeMonths[-1]+1]) if activeMonths[-1] >= 2 else sum(monthlyHours[-3:])
            if recent3 > first3 * 1.15:
                trend = 'up'
            elif recent3 < first3 * 0.85:
                trend = 'down'
            else:
                trend = 'stable'
        else:
            trend = 'stable'
        # 综合成长评分 (0-100)
        totalH = sum(monthlyHours)
        avgTypes = sum(monthlyTypes) / max(activeMonths.__len__(), 1)
        valPerHour = sum(monthlyValue) / totalH if totalH > 0 else 0
        # 类型多样性得分
        typeScore = min(avgTypes * 15, 30)
        # 价值密度得分
        valScore = min(valPerHour * 5, 30)
        # 工时投入得分
        hourScore = min(totalH / 20, 20)
        # 趋势得分
        trendScore = 20 if trend == 'up' else (15 if trend == 'stable' else 10)
        growthScore = round(typeScore + valScore + hourScore + trendScore)
        # 优势标签
        strengths = []
        if avgTypes >= 4: strengths.append('多面手')
        if valPerHour > 5: strengths.append('价值高效')
        if totalH > 900: strengths.append('高产')
        if trend == 'up': strengths.append('成长中')
        if not strengths: strengths.append('稳步发展')
        creatorsGrowth.append({
            'name': name,
            'monthlyHours': monthlyHours,
            'monthlyTypes': monthlyTypes,
            'monthlyValue': monthlyValue,
            'trend': trend, 'growthScore': growthScore,
            'strengths': strengths, 'totalHours': round(totalH, 1)
        })
    creatorsGrowth.sort(key=lambda x: -x['growthScore'])

    # ===== 上半年总结 =====
    monthlyTotal = [round(df[df['月份'] == m]['工时'].sum(), 1) for m in months]
    monthlyRecords = [len(df[df['月份'] == m]) for m in months]
    monthlyAvgDaily = []
    monthlyWorkDays = []
    for m in months:
        md = df[df['月份'] == m]
        dates = md[col_map['发生时间']].dropna()
        wd = dates[dates.dt.weekday < 5].dt.date.nunique()
        monthlyWorkDays.append(wd)
        th = round(md['工时'].sum(), 1)
        monthlyAvgDaily.append(round(th / wd, 1) if wd > 0 else 0)
    # 每月类型分布
    monthlyTypeBreakdown = []
    for m in months:
        md = df[df['月份'] == m]
        breakdown = {}
        for wt in workTypes:
            breakdown[wt] = round(md[md[col_map['工作类型']] == wt]['工时'].sum(), 1)
        monthlyTypeBreakdown.append(breakdown)
    # 上半年Top项目
    halfYearProjHours = {}
    for name, g in df.groupby(col_map['项目名称']):
        if pd.isna(name): continue
        halfYearProjHours[name] = round(g['工时'].sum(), 1)
    halfYearTopProj = sorted(halfYearProjHours.items(), key=lambda x:-x[1])[:20]
    # 峰值月
    maxMonthIdx = monthlyTotal.index(max(monthlyTotal))
    peakMonth = monthLabels[maxMonthIdx] if maxMonthIdx < len(monthLabels) else ''
    # 最强成长员工
    strongestGrowth = creatorsGrowth[0]['name'] if creatorsGrowth else ''
    # 关键洞察
    avgMonthlyH = round(total / len(months), 1)
    keyInsight = f'上半年月均工时{avgMonthlyH}h，{strongestGrowth}综合成长表现最优。'
    if maintTotal > total * 0.3:
        keyInsight += f'维护占比{pct(maintTotal, total)}%，建议关注维护优化空间。'

    # 12. 个人工作总结（李洋冰、李泽垣）
    def build_person_summary(name):
        pdf = df[df[col_map['负责人']] == name].copy()
        if pdf.empty:
            return None
        totalH = round(pdf['工时'].sum(), 1)
        recs = len(pdf)
        projs = pdf[col_map['项目名称']].nunique()
        wtypes = pdf[col_map['工作类型']].nunique()
        monthly = []
        for m in months:
            mdf = pdf[pdf['月份'] == m]
            mh = round(mdf['工时'].sum(), 1)
            mt = sorted(mdf[col_map['工作类型']].unique())
            monthly.append({
                'month': int(m), 'hours': float(mh), 'typesCount': int(len(mt)),
                'types': [str(t) for t in mt], 'records': int(len(mdf))
            })
        monthlyAvg = round(totalH / max(len([x for x in monthly if x['hours'] > 0]), 1), 1)
        # 工作类型分布
        typeDist = {}
        for t, g in pdf.groupby(col_map['工作类型']):
            if pd.isna(t): continue
            typeDist[t] = round(g['工时'].sum(), 1)
        # 项目实施分布
        implDf = pdf[pdf[col_map['工作类型']] == '项目实施']
        implH = round(implDf['工时'].sum(), 1) if len(implDf) > 0 else 0
        implProjs = {}
        for pn, g in implDf.groupby(col_map['项目名称']):
            if pd.isna(pn): continue
            implProjs[pn] = round(g['工时'].sum(), 1)
        implTop = sorted(implProjs.items(), key=lambda x: -x[1])[:5]
        # 售前分布
        preDf = pdf[pdf[col_map['工作类型']] == '售前']
        preH = round(preDf['工时'].sum(), 1) if len(preDf) > 0 else 0
        preProjs = {}
        for pn, g in preDf.groupby(col_map['项目名称']):
            if pd.isna(pn): continue
            preProjs[pn] = round(g['工时'].sum(), 1)
        preTop = sorted(preProjs.items(), key=lambda x: -x[1])[:5]
        # 项目管理
        pmDf = pdf[pdf[col_map['工作类型']] == '项目管理']
        pmH = round(pmDf['工时'].sum(), 1) if len(pmDf) > 0 else 0
        # 维护分布
        maintDf = pdf[pdf[col_map['工作类型']] == '维护']
        maintH = round(maintDf['工时'].sum(), 1) if len(maintDf) > 0 else 0
        maintProjs = {}
        for pn, g in maintDf.groupby(col_map['项目名称']):
            if pd.isna(pn): continue
            maintProjs[pn] = round(g['工时'].sum(), 1)
        maintTop = sorted(maintProjs.items(), key=lambda x: -x[1])[:5]
        # 数字化升级
        digiDf = pdf[pdf[col_map['工作类型']] == '数字化升级']
        digiH = round(digiDf['工时'].sum(), 1) if len(digiDf) > 0 else 0
        # 学习
        learnDf = pdf[pdf[col_map['工作类型']] == '学习/内部例会']
        learnH = round(learnDf['工时'].sum(), 1) if len(learnDf) > 0 else 0
        # 项目等级分布
        levelDist = {}
        for lv, g in pdf.groupby(col_map['项目等级']):
            if pd.isna(lv): continue
            levelDist[lv] = round(g['工时'].sum(), 1)
        # Top项目
        projHours = {}
        for pn, g in pdf.groupby(col_map['项目名称']):
            if pd.isna(pn): continue
            projHours[pn] = round(g['工时'].sum(), 1)
        projTop = sorted(projHours.items(), key=lambda x: -x[1])[:5]
        # 价值工时
        personValue = round(pdf['价值工时'].sum(), 1)
        # 高峰/低谷
        if monthly:
            peakMonth = max(monthly, key=lambda x: x['hours'])
            troughMonth = min((x for x in monthly if x['hours'] > 0), key=lambda x: x['hours'], default=monthly[0])
        else:
            peakMonth = {'month': 0, 'hours': 0}
            troughMonth = {'month': 0, 'hours': 0}
        # 成长趋势：类型扩展 + 工时稳定性
        typeCounts = [x['typesCount'] for x in monthly]
        typeGrowth = typeCounts[-1] - typeCounts[0] if len(typeCounts) >= 2 else 0
        hasSale = preH > 0
        hasDigi = digiH > 0
        hourTrend = '上升' if len(monthly) >= 2 and monthly[-1]['hours'] > monthly[0]['hours'] * 1.1 else ('平稳' if len(monthly) >= 2 else '--')
        return {
            'name': str(name), 'totalHours': float(totalH), 'records': int(recs), 'projCount': int(projs),
            'workTypeCount': int(wtypes), 'monthlyAvg': float(monthlyAvg),
            'monthly': monthly, 'typeDist': {str(k): float(v) for k,v in typeDist.items()},
            'projTop': [[str(k), float(v)] for k,v in projTop],
            'implHours': float(implH), 'implTop': [[str(k), float(v)] for k,v in implTop],
            'preHours': float(preH), 'preTop': [[str(k), float(v)] for k,v in preTop], 'hasSale': bool(hasSale),
            'pmHours': float(pmH), 'maintHours': float(maintH),
            'maintTop': [[str(k), float(v)] for k,v in maintTop],
            'digiHours': float(digiH), 'hasDigi': bool(hasDigi), 'learnHours': float(learnH),
            'levelDist': {str(k): float(v) for k,v in levelDist.items()},
            'valueHours': float(personValue),
            'peakMonth': {'month': int(peakMonth['month']), 'hours': float(peakMonth['hours'])},
            'troughMonth': {'month': int(troughMonth['month']), 'hours': float(troughMonth['hours'])},
            'typeGrowth': int(typeGrowth), 'hourTrend': str(hourTrend),
        }

    personLyb = build_person_summary('李洋冰')
    personLzy = build_person_summary('李泽垣')
    personLhf = build_person_summary('吕洪飞')

    # ===== 疑似凑数工时分析 =====
    desc_len_col = '内容描述' if '内容描述' in df.columns else None
    paddingData = None
    if desc_len_col:
        df['desc_len'] = df[desc_len_col].astype(str).str.len()
        # 判定阈值
        PAD_LEN_THRESHOLD = int(df['desc_len'].quantile(0.25))  # P25
        PAD_HOUR_THRESHOLD = 2.0
        SEVERE_LEN = 10
        SEVERE_HOUR = 4.0

        # 疑似凑数
        padding_mask = (df['desc_len'] < PAD_LEN_THRESHOLD) & (df['工时'] > PAD_HOUR_THRESHOLD)
        padding_df = df[padding_mask].copy()
        # 严重凑数
        severe_mask = (df['desc_len'] <= SEVERE_LEN) & (df['工时'] > SEVERE_HOUR)
        severe_df = df[severe_mask].copy()
        # 仅"[图片]"类占位
        placeholder_mask = df[desc_len_col].astype(str).str.match(r'^\s*(\[图片\]\s*)+$')
        placeholder_df = df[placeholder_mask & (df['工时'] > 1)].copy()

        # 凑数明细列表
        paddingDetail = []
        for _, r in padding_df.sort_values('工时', ascending=False).iterrows():
            paddingDetail.append({
                'creator': str(r[col_map['负责人']]),
                'hours': float(r['工时']),
                'descLen': int(r['desc_len']),
                'workType': str(r[col_map['工作类型']]),
                'project': str(r[col_map['项目名称']]) if pd.notna(r[col_map['项目名称']]) else '',
                'desc': str(r[desc_len_col])[:80],
                'level': str(r[col_map['项目等级']]) if pd.notna(r[col_map['项目等级']]) else '',
                'severity': 'severe' if r['desc_len'] <= SEVERE_LEN and r['工时'] > SEVERE_HOUR else 'suspicious'
            })

        # 按人员汇总
        paddingByCreator = []
        all_creators = df[col_map['负责人']].dropna().unique()
        for name in all_creators:
            p = padding_df[padding_df[col_map['负责人']] == name]
            s = severe_df[severe_df[col_map['负责人']] == name]
            ph = placeholder_df[placeholder_df[col_map['负责人']] == name]
            person_total = df[df[col_map['负责人']] == name]['工时'].sum()
            paddingByCreator.append({
                'name': str(name),
                'suspectCount': int(len(p)),
                'suspectHours': round(p['工时'].sum(), 1),
                'severeCount': int(len(s)),
                'severeHours': round(s['工时'].sum(), 1),
                'placeholderCount': int(len(ph)),
                'placeholderHours': round(ph['工时'].sum(), 1),
                'totalHours': round(person_total, 1),
                'suspectPct': round(p['工时'].sum() / person_total * 100, 1) if person_total > 0 else 0
            })

        # 按工作类型汇总
        paddingByType = []
        for wt in workTypes:
            p = padding_df[padding_df[col_map['工作类型']] == wt]
            if len(p) > 0:
                paddingByType.append({
                    'type': str(wt),
                    'count': int(len(p)),
                    'hours': round(p['工时'].sum(), 1)
                })

        # 按月份汇总
        paddingByMonth = []
        for m in sorted(df['月份'].unique()):
            if m > 6: continue
            p = padding_df[padding_df['月份'] == m]
            paddingByMonth.append({
                'month': int(m),
                'count': int(len(p)),
                'hours': round(p['工时'].sum(), 1)
            })

        paddingData = {
            'totalSuspectCount': int(len(padding_df)),
            'totalSuspectHours': round(padding_df['工时'].sum(), 1),
            'totalSevereCount': int(len(severe_df)),
            'totalSevereHours': round(severe_df['工时'].sum(), 1),
            'totalPlaceholderCount': int(len(placeholder_df)),
            'totalPlaceholderHours': round(placeholder_df['工时'].sum(), 1),
            'suspectPctRecords': round(len(padding_df) / len(df) * 100, 1),
            'suspectPctHours': round(padding_df['工时'].sum() / df['工时'].sum() * 100, 1),
            'padLenThreshold': PAD_LEN_THRESHOLD,
            'padHourThreshold': PAD_HOUR_THRESHOLD,
            'byCreator': paddingByCreator,
            'byType': paddingByType,
            'byMonth': paddingByMonth,
            'details': paddingDetail[:100],  # Top100明细
        }

    halfYearSummary = {
        'monthLabels': monthLabels,
        'monthlyTotal': monthlyTotal,
        'monthlyRecords': monthlyRecords,
        'monthlyAvgDaily': monthlyAvgDaily,
        'monthlyWorkDays': monthlyWorkDays,
        'monthlyTypeBreakdown': monthlyTypeBreakdown,
        'halfYearTopProj': [[k, v] for k, v in halfYearTopProj],
        'peakMonth': peakMonth,
        'totalHours': round(total, 1),
        'totalRecords': records,
        'projCount': projCount,
        'strongestGrowth': strongestGrowth,
        'keyInsight': keyInsight,
        'monthlyMaintPct': [round(monthlyTypeBreakdown[i].get('维护', 0) / monthlyTotal[i] * 100, 1) if monthlyTotal[i] > 0 else 0 for i in range(len(months))],
    }

    return dict(
        total=round(total,1), records=records, projCount=projCount, workDays=workDays,
        maintTotal=round(maintTotal,1), p0Daily=p0Daily, allDaily=allDaily,
        highValue=round(highValue,1), highValuePct=highValuePct, optimizable=round(optimizable,1),
        totalValue=round(totalValue,1),
        typePie=typePie, typeColors=typeColors,
        projNames=projNames, projHours=projHours, projMaint=projMaint,
        projMaintPct=projMaintPct, projValuePct=projValuePct, projLevel=projLevel,
        projValue=projValue, projValue2=projValue2,
        maintType=maintType, prodLine=prodLine, maintReason=maintReason, maintLevel=maintLevel,
        maintClass=maintClass, goutClass=goutClass, projMaintClass=projMaintClass,
        hmProjs=hmProjs, hmTypes=hmTypes, hmData=hmData,
        dailyDates=dailyDates, dailyMaint=dailyMaint, dailyPM=dailyPM,
        dailyDigital=dailyDigital, dailyImpl=dailyImpl, dailyLearn=dailyLearn, dailySale=dailySale,
        creatorNames=creatorNames, creatorTotalHours=creatorTotalHours,
        workTypes=workTypes, creatorTypeMatrix=creatorTypeMatrix,
        creatorProjTop3=creatorProjTop3, creatorValue=creatorValue, creatorValue2=creatorValue2,
        stdHours=stdHours, creatorStd=creatorStd,
        minDate=minDate, maxDate=maxDate, pmTotal=pmTotal,
        pmByProject=pmByProject, pmByCreator=pmByCreator,
        levelOrder=levelOrder, creatorLevelMatrix=creatorLevelMatrix,
        maintTypeList=maintTypeList, creatorMaintMatrix=creatorMaintMatrix,
        deptNames=deptNames, deptTotalHours=deptTotalHours,
        deptTypeMatrix=deptTypeMatrix, deptValue=deptValue,
        deptMaintMatrix=deptMaintMatrix, deptLevelMatrix=deptLevelMatrix,
        deptProjTop3=deptProjTop3,
        typeValue=typeValue, levelValue=levelValue,
        preSaleTotal=preSaleTotal, preSaleProj=preSaleProj,
        preSalePie=preSalePie, preSaleCreator=preSaleCreator,
        digitalUpgrade=digitalUpgrade, digitalTotal=digitalTotal,
        digitalPieData=digitalPieData,
        highValueDetail=highValueDetail, highValueDetailTotal=highValueDetailTotal,
        monthlyData=monthlyData, transferredFlags=transferredFlags,
        transferredProjects=TRANSFERRED_PROJECTS,
        monthLabels=monthLabels, creatorsGrowth=creatorsGrowth,
        halfYearSummary=halfYearSummary,
        personLyb=personLyb, personLzy=personLzy, personLhf=personLhf,
        paddingData=paddingData,
    )

# ===== 文件路径 =====
CURR_FILE  = '【1-6月份】项目工时登记对象导出结果.xlsx'   # 本月文件
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
