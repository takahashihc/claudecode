# -*- coding: utf-8 -*-
# 9月単月: 前年同月(2509)と今期(2609)の売上全明細を商品別に突き合わせ、価格・数量・原価に分解
import openpyxl, json, collections
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
codes={"石見":["0011","0012","0013","0014","0015","0016","0018","0019","0020"],"下関":["0021","0024","0025","0026","0027","0028","0029"],"松江":["0032","0033","0034","0035","0039"],"広域":["0042","0045","0046","0047","0550","0552"],"境港":["0903","0904","0909","0910","0911","0914","0918","0919","0930"],"水産部":["0010","0017","0031","0038","0101","0102","0104","0106"],"キョウワ":["0502"]}
c2o={c:o for o,cs in codes.items() for c in cs}
def load(f):
    agg=collections.defaultdict(lambda:[0.0,0.0,0.0,0.0])  # qty, sales, cost, gp
    ws=openpyxl.load_workbook(S+f,read_only=True,data_only=True).worksheets[0]
    for r in ws.iter_rows(min_row=2,values_only=True):
        if r[0] is None: continue
        o=c2o.get(str(r[3]).zfill(4)); 
        if not o: continue
        kind="機械" if str(r[14])=="08" else "資材"
        a=agg[(o,kind,str(r[5]))]; a[0]+=float(r[8] or 0); a[1]+=float(r[10] or 0); a[2]+=float(r[12] or 0); a[3]+=float(r[13] or 0)
    return agg
P=load("sep/2509_detail.xlsx"); C=load("sep/2609タカハシ包装売上全明細.xlsx")
def analyze(keys):
    r=dict(s0=0,s1=0,g0=0,g1=0,vol_s=0,price_s=0,new_s=0,lost_s=0,vol_g=0,price_g=0,cost_g=0,new_g=0,lost_g=0,m_s0=0,m_s1=0,n_matched=0,n_new=0,n_lost=0)
    for k in keys:
        p=P.get(k,[0,0,0,0]); c=C.get(k,[0,0,0,0])
        r['s0']+=p[1]; r['s1']+=c[1]; r['g0']+=p[3]; r['g1']+=c[3]
        if p[0]>0 and c[0]>0 and p[1]>0:
            p0=p[1]/p[0]; p1=c[1]/c[0]; c0=p[2]/p[0]; c1=c[2]/c[0]
            r['vol_s']+=p0*(c[0]-p[0]); r['price_s']+=c[0]*(p1-p0)
            r['vol_g']+=(p0-c0)*(c[0]-p[0]); r['price_g']+=c[0]*(p1-p0); r['cost_g']+=-c[0]*(c1-c0)
            r['m_s0']+=p[1]; r['m_s1']+=c[1]; r['n_matched']+=1
        else:
            if c[1]!=0 or c[0]!=0: r['new_s']+=c[1]; r['new_g']+=c[3]; r['n_new']+=1
            if p[1]!=0 or p[0]!=0: r['lost_s']-=p[1]; r['lost_g']-=p[3]; r['n_lost']+=1
    # laspeyres数量指数 = 前年価格で評価した今期数量 / 前年売上(マッチ分)
    r['qty_idx']=(r['m_s0']+r['vol_s'])/r['m_s0']*100 if r['m_s0'] else None
    r['price_idx']=(r['m_s1'])/(r['m_s1']-r['price_s'])*100 if r['m_s1'] else None
    return r
allk=set(P)|set(C); out={}
for kind in ("資材","機械","全"):
    for o in list(codes)+["タカハシ包装(6)","合計"]:
        offs=codes.keys() if o=="合計" else (list(codes)[:6] if o=="タカハシ包装(6)" else [o])
        keys=[k for k in allk if k[0] in offs and (kind=="全" or k[1]==kind)]
        out[f"{kind}|{o}"]=analyze(keys)
json.dump(out,open(S+"pricevol_sep.json","w"),ensure_ascii=False,indent=1)
def f(x): return f"{x/1000:>9,.0f}"
for kind in ("資材","機械","全"):
    print(f"\n=== {kind} 9月単月（千円）===")
    print(f"{'':10}{'売上前年':>9}{'売上今期':>9}{'数量効果':>9}{'価格効果':>9}{'新規':>9}{'消滅':>9} | {'粗利前年':>9}{'粗利今期':>9}{'数量効果':>9}{'価格効果':>9}{'原価効果':>9}{'新規消滅':>9} | 数量指数 価格指数")
    for o in list(codes)+["タカハシ包装(6)","合計"]:
        r=out[f"{kind}|{o}"]
        qi=f"{r['qty_idx']:.1f}" if r['qty_idx'] else '-'; pi=f"{r['price_idx']:.1f}" if r['price_idx'] else '-'
        print(f"{o:10}{f(r['s0'])}{f(r['s1'])}{f(r['vol_s'])}{f(r['price_s'])}{f(r['new_s'])}{f(r['lost_s'])} |{f(r['g0'])}{f(r['g1'])}{f(r['vol_g'])}{f(r['price_g'])}{f(r['cost_g'])}{f(r['new_g']+r['lost_g'])} | {qi:>7} {pi:>7}")
