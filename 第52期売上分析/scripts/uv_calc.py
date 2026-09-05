import pandas as pd, numpy as np
from analyze import load, decomp, MAP
df=pd.concat([load('in/2508.xlsx'),load('in/2608.xlsx')],ignore_index=True)
nk_all=df[df['商品群']!='什器備品・機械その他']          # 突合用（P9と同条件）
nk=nk_all[nk_all['売上数量']>0]                            # 平均単価用
rows=[]
for (code,name),g in nk.groupby(['分類','分類名称１（商品）']):
    a=g[g['年']=='2025']; b=g[g['年']=='2026']
    S0,S1,Q0,Q1=a['売上金額'].sum(),b['売上金額'].sum(),a['売上数量'].sum(),b['売上数量'].sum()
    if S0<=0 or S1<=0 or Q0<=0 or Q1<=0: continue
    v=decomp(nk_all[nk_all['分類']==code]); up=g['販売単価']; up=up[up>0]
    rows.append(dict(分類=code,名称=name,商品群=g['商品群'].iloc[0],前年売上=S0,当年売上=S1,金額増減=S1/S0-1,数量増減=Q1/Q0-1,
        平均単価変化=(S1/Q1)/(S0/Q0)-1,突合単価効果=(v['単価効果'] if v else np.nan),突合品目=(v['品目数'] if v else 0),単価ばらつき=np.log10(up.quantile(.9)/up.quantile(.1))))
t=pd.DataFrame(rows).sort_values('前年売上',ascending=False); t.to_csv('uv_by_class.csv',index=False)
out=[]
for gname in list(MAP)+['全体']:
    if gname=='什器備品・機械その他': continue
    s=t if gname=='全体' else t[t['商品群']==gname]
    if len(s)==0: continue
    w=s['前年売上']; idx=(w*(1+s['平均単価変化'])).sum()/w.sum()-1
    sub=nk_all if gname=='全体' else nk_all[nk_all['商品群']==gname]
    v=decomp(sub); a=sub[sub['年']=='2025']; b=sub[sub['年']=='2026']
    pe=v['単価効果'] if v else np.nan
    out.append(dict(商品群=gname,売上増減=b['売上金額'].sum()/a['売上金額'].sum()-1,分類別平均単価指数=idx,突合単価効果=pe,差=idx-pe))
o=pd.DataFrame(out); o.to_csv('uv_by_group.csv',index=False); print(o.to_string(index=False))
