# -*- coding: utf-8 -*-
"""8月（2025 vs 2026）の単価・数量寄与分解。商品群集計ツール.py と同じロジック（ラスパイレス、0.25〜4.0倍除外、10品目未満は非表示）"""
import pandas as pd, numpy as np, json, sys
IN='in'
MAP = {  # 旧分類コード → 新商品群（商品群_新旧対応表_20260822）
 '食品容器':[2,1,32,23,33,21],
 'フィルム・袋・ラミネート':[5,4,34,36,35,38,31,24],
 '紙箱・紙製品・ラベル・段ボール':[3,6,22,18,30,10],
 '木箱':[7,17,15],
 '消耗品':[40,27,20,37,41,39,26,98,28,42,43,25,29,44],
 '荷役':[9,90,19],
 '什器備品・機械その他':[8,14],
 'その他':[99],
}
code2grp={c:g for g,cs in MAP.items() for c in cs}
OFFICE = {
 '石見営業所':[11,12,13,14,19,20],
 '下関営業所':[21,24,25,26,27,28,29],
 '松江営業所':[32,33,34,35,39],
 '広域営業部':[42,45,46,47,550,552],
 '境港営業所':[903,904,909,910,911,914,918,919,930],
 '水産部':[10,17,31,38,101,102,104,106],
}
code2off={c:o for o,cs in OFFICE.items() for c in cs}

def load(fn):
    d=pd.read_excel(fn)
    d['年']=d['売上伝票日付'].astype(str).str[:4]
    d['月']=d['売上伝票日付'].astype(str).str[4:6].astype(int)
    d['分類']=d['分類コード1(商品)'].astype(int)
    d['商品群']=d['分類'].map(code2grp)
    d['担当']=d['担当者コード'].astype(int)
    d['営業所']=d['担当'].map(code2off)
    return d

def decomp(d, lo=0.25, hi=4.0):
    years=sorted(d['年'].unique())
    if len(years)<2: return None
    y0,y1=years[0],years[-1]
    p=d.pivot_table(index='商品コード',columns='年',values=['売上数量','売上金額'],aggfunc='sum')
    if ('売上数量',y0) not in p or ('売上数量',y1) not in p: return None
    p=p.dropna(); p=p[(p[('売上数量',y0)]>0)&(p[('売上数量',y1)]>0)]
    if len(p)==0: return None
    q0=p[('売上数量',y0)]; q1=p[('売上数量',y1)]; s0=p[('売上金額',y0)]; s1=p[('売上金額',y1)]
    pr0=s0/q0; pr1=s1/q1
    keep=((pr1/pr0)>=lo)&((pr1/pr0)<=hi)
    q0,q1,s0,s1,pr0,pr1=[x[keep] for x in (q0,q1,s0,s1,pr0,pr1)]
    if len(q0)<10: return None
    S0=float(s0.sum()); S1=float(s1.sum()); P=float((q0*pr1).sum())/S0
    return dict(品目数=int(len(q0)),除外=int((~keep).sum()),前年売上=S0,当年売上=S1,単価効果=P-1,数量効果=(S1/S0)-P,金額増減=S1/S0-1)

if __name__=='__main__':
    files=sys.argv[1:] or [f'{IN}/2508.xlsx',f'{IN}/2608.xlsx']
    df=pd.concat([load(f) for f in files],ignore_index=True)
    print('未分類コード:',sorted(df[df['商品群'].isna()]['分類'].unique()))
    print('未割当担当:',df[df['営業所'].isna()].groupby(['担当','担当者名'])['売上金額'].sum())
    nk=df[df['商品群']!='什器備品・機械その他']
    out={}
    for m in sorted(nk['月'].unique()):
        dm=nk[nk['月']==m]
        out[f'全社_{m}']=decomp(dm)
        for g in MAP:
            if g=='什器備品・機械その他': continue
            out[f'{g}_{m}']=decomp(dm[dm['商品群']==g])
        for o in OFFICE:
            out[f'{o}_{m}']=decomp(dm[dm['営業所']==o])
    for k,v in out.items():
        if v: print(f"{k:28s} n={v['品目数']:5d} 除外={v['除外']:3d} 金額{v['金額増減']:+.1%} 単価{v['単価効果']:+.1%} 数量{v['数量効果']:+.1%}")
        else: print(f"{k:28s} -")
    # 参考：売上・粗利の単純集計（8月）
    for key in ['商品群','営業所']:
        g=df.groupby([key,'年']).agg(売上=('売上金額','sum'),粗利=('売上粗利額','sum')).unstack('年')
        print(g.to_string())
    json.dump(out,open('aug_result.json','w'),ensure_ascii=False,indent=1)
