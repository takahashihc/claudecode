# -*- coding: utf-8 -*-
import openpyxl, json, collections
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
codes={"石見":["0011","0012","0013","0014","0015","0016","0018","0019","0020"],"下関":["0021","0024","0025","0026","0027","0028","0029"],"松江":["0032","0033","0034","0035","0039"],"広域":["0042","0045","0046","0047","0550","0552"],"境港":["0903","0904","0909","0910","0911","0914","0918","0919","0930"],"水産部":["0010","0017","0031","0038","0101","0102","0104","0106"]}
c2o={c:o for o,cs in codes.items() for c in cs}
def load(f):
    prod=collections.defaultdict(lambda:[0.0,0.0,0.0]); cust=collections.defaultdict(lambda:[0.0,0.0])
    ws=openpyxl.load_workbook(S+f,read_only=True,data_only=True).worksheets[0]
    for r in ws.iter_rows(min_row=2,values_only=True):
        if r[0] is None: continue
        o=c2o.get(str(r[3]).zfill(4))
        if not o or str(r[14])=="08": continue     # 機械は除く（大口単発のため）
        a=prod[(o,str(r[5]),str(r[6]))]; a[0]+=float(r[8] or 0); a[1]+=float(r[10] or 0); a[2]+=float(r[12] or 0)
        c=cust[(o,str(r[1]),str(r[2]))]; c[0]+=float(r[10] or 0); c[1]+=float(r[13] or 0)
    return prod,cust
(P,PC),(C,CC)=load("sep/2509_detail.xlsx"),load("sep/2609タカハシ包装売上全明細.xlsx")
# 商品キーは(担当営業所, 商品コード)。商品名は表示用に別持ち
def key2(d):
    r=collections.defaultdict(lambda:[0.0,0.0,0.0,''])
    for (o,code,name),v in d.items():
        a=r[(o,code)]; a[0]+=v[0]; a[1]+=v[1]; a[2]+=v[2]; a[3]=name
    return r
P2,C2=key2(P),key2(C)
res={}
def pv(offs):
    s1=sum(v[1] for k,v in C2.items() if k[0] in offs); s0=sum(v[1] for k,v in P2.items() if k[0] in offs)
    m0=m1=lasp=pa=cost_p=0.0; n=0; ex=0
    for k in set(P2)&set(C2):
        if k[0] not in offs: continue
        p,c=P2[k],C2[k]
        if p[0]<=0 or c[0]<=0 or p[1]<=0: continue
        p0=p[1]/p[0]; p1=c[1]/c[0]
        if not (0.5<=p1/p0<=2.0): ex+=c[1]; continue          # 単位変更・特価等で比較不能な品目を除外
        m0+=p[1]; m1+=c[1]; lasp+=p0*c[0]; n+=1
    return dict(s0=s0,s1=s1,m0=m0,m1=m1,cover=m1/s1*100 if s1 else 0,qty_idx=lasp/m0*100 if m0 else None,price_idx=m1/lasp*100 if lasp else None,n=n,excluded=ex)
for o in list(codes)+["タカハシ包装(6)"]:
    res[o]=pv(list(codes) if o=="タカハシ包装(6)" else [o])
# 取引先ベース
def cv(offs):
    P_={k:v for k,v in PC.items() if k[0] in offs}; C_={k:v for k,v in CC.items() if k[0] in offs}
    both=set(P_)&set(C_); 
    return dict(n_prior=len(P_),n_cur=len(C_),n_both=len(both),lost_s=sum(P_[k][0] for k in P_ if k not in C_),new_s=sum(C_[k][0] for k in C_ if k not in P_),
                cont_s0=sum(P_[k][0] for k in both),cont_s1=sum(C_[k][0] for k in both),cont_g0=sum(P_[k][1] for k in both),cont_g1=sum(C_[k][1] for k in both),
                down=sum(1 for k in both if C_[k][0]<P_[k][0]*0.9),down_amt=sum(C_[k][0]-P_[k][0] for k in both if C_[k][0]<P_[k][0]))
for o in list(codes)+["タカハシ包装(6)"]:
    res[o]['cust']=cv(list(codes) if o=="タカハシ包装(6)" else [o])
json.dump(res,open(S+"pricevol2_sep.json","w"),ensure_ascii=False,indent=1)
print(f"{'':10}{'売上前年':>9}{'売上今期':>9} 継続品カバー率 数量指数 単価指数 | 取引先 前年/今期/継続 継続先売上増減 離脱売上 新規売上 10%超減少先数")
for o,r in res.items():
    c=r['cust']
    print(f"{o:10}{r['s0']/1000:9,.0f}{r['s1']/1000:9,.0f} {r['cover']:8.0f}% {r['qty_idx']:8.1f} {r['price_idx']:8.1f} | {c['n_prior']:5}/{c['n_cur']:5}/{c['n_both']:5} {(c['cont_s1']-c['cont_s0'])/1000:9,.0f} {-c['lost_s']/1000:9,.0f} {c['new_s']/1000:9,.0f} {c['down']:6}")
# 石見: 数量減少の大きい継続品
print('\n石見 継続品で数量が減った上位（前年価格×数量差）')
rows=[]
for k in set(P2)&set(C2):
    if k[0]!='石見': continue
    p,c=P2[k],C2[k]
    if p[0]<=0 or c[0]<=0 or p[1]<=0: continue
    p0=p[1]/p[0]; p1=c[1]/c[0]
    if p0<=0 or not (0.5<=p1/p0<=2.0): continue
    rows.append((p0*(c[0]-p[0]),k[1],c[3][:26],p[0],c[0],p0,p1))
for x in sorted(rows)[:8]: print(f"  {x[0]/1000:8,.0f}  {x[1]:10} {x[2]:28} 数量 {x[3]:9,.0f}->{x[4]:9,.0f} 単価 {x[5]:7.1f}->{x[6]:7.1f}")
