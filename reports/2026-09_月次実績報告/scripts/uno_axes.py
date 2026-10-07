# -*- coding: utf-8 -*-
# 年計表: 指定グラフの左右軸の目盛り区切り数をそろえる（横線を一致させる）
import uno, sys, math
from com.sun.star.beans import PropertyValue
SRC,DST=sys.argv[1],sys.argv[2]
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
ctx=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local).resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
doc=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx).loadComponentFromURL(uno.systemPathToFileUrl(SRC),"_blank",0,(prop("Hidden",True),))
S=doc.Sheets; doc.calculateAll()
def rowmax(rep):
    sh,rng=rep.replace("$","").split(".")
    ws=S.getByName(sh); a,b=rng.split(":")
    r=ws.getCellRangeByName(rng)
    return max(v for row in r.getDataArray() for v in row if isinstance(v,float))
def nice(mx,n):
    # n区切りで mx 以上となる切りの良い目盛り幅
    raw=mx/n; e=10**math.floor(math.log10(raw))
    for m in (1,1.25,1.5,2,2.5,3,4,5,6,7.5,8,10):
        if m*e*n>=mx: return m*e
for sheet in sys.argv[3:]:
    cd=S.getByName(sheet).Charts.getByIndex(0).EmbeddedObject; dg=cd.getDiagram()
    reps=[ls.Values.SourceRangeRepresentation for cs in cd.FirstDiagram.CoordinateSystems for ct in cs.ChartTypes for s in ct.DataSeries for ls in s.DataSequences]
    smax=rowmax([r for r in reps if "売上" in r][0]); gmax=rowmax([r for r in reps if "粗利" in r][0])
    best=None
    for n in (5,6,7,8,9,10):
        ss=nice(smax,n); gs=nice(gmax,n)
        waste=(ss*n-smax)/(ss*n)+(gs*n-gmax)/(gs*n)
        if best is None or waste<best[0]-1e-9: best=(waste,n,ss,gs)
    _,n,ss,gs=best
    FIX={"【グラフ】タカハシ包装合計":(8,500000.0,125000.0),"【グラフ】グループ合計":(8,500000.0,125000.0)}   # 読みやすい固定目盛り（両グラフ共通）
    if sheet in FIX:
        n,ss,gs=FIX[sheet]; assert ss*n>=smax and gs*n>=gmax
    for ax,step in ((dg.getYAxis(),ss),(dg.getSecondaryYAxis(),gs)):
        for k,v in (("AutoMin",False),("Min",0.0),("AutoMax",False),("Max",step*n),("AutoStepMain",False),("StepMain",step)):
            ax.setPropertyValue(k,v)
    cd.setModified(True)
    print(f"{sheet}: 売上max {smax:,.0f} → 0〜{ss*n:,.0f}（{ss:,.0f}刻み） / 粗利max {gmax:,.0f} → 0〜{gs*n:,.0f}（{gs:,.0f}刻み） / 区切り {n}")
doc.calculateAll()
doc.storeToURL(uno.systemPathToFileUrl(DST),(prop("FilterName","MS Excel 97"),))
doc.close(True); print("saved",DST)
