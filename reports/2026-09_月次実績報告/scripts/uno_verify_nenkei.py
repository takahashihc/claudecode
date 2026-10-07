import uno, sys
from com.sun.star.beans import PropertyValue
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
ctx=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local).resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
doc=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx).loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[1]),"_blank",0,(prop("Hidden",True),))
S=doc.Sheets; doc.calculateAll()
print(S.ElementNames)
for n in S.ElementNames:
    sh=S.getByName(n)
    if sh.Charts.Count==0: continue
    cd=sh.Charts.getByIndex(0).EmbeddedObject
    reps=[ls.Values.SourceRangeRepresentation for cs in cd.FirstDiagram.CoordinateSystems for ct in cs.ChartTypes for s in ct.DataSeries for ls in s.DataSequences]
    print(f"{n:24} title={cd.Title.String if cd.HasMainTitle else None:24} {reps}")
for sn,rows in [("年計表売上DATA",{"石見":7,"下関":15,"松江":27,"境港":43,"水産部":59,"キョウワ":71,"広域":75,"タカハシ包装営業":79,"タカハシ包装合計":83,"グループ合計":87,"旧 合計（境港含）":67}),
                ("年計表粗利益DATA",{"石見":7,"下関":15,"松江":27,"境港":43,"水産部":55,"キョウワ":67,"広域":71,"タカハシ包装営業":75,"タカハシ包装合計":79,"グループ合計":83,"旧 境港含む":63})]:
    ws=S.getByName(sn)
    print("==",sn,"2026/9 当月 / 年計")
    for k,r in rows.items():
        print(f"   {k:14} {ws.getCellByPosition(173,r-1).getValue():12,.1f} {ws.getCellByPosition(173,r).getValue():14,.1f}  label={ws.getCellByPosition(2,r-3).getString()}")
    # 新合計と旧合計の差（年計行、全月）
    newr=rows["タカハシ包装合計"]; oldr=rows[[k for k in rows if k.startswith("旧")][0]]
    diffs=[(ws.getCellByPosition(5,newr-2).getString() or ws.getCellByPosition(c,newr-2).getString(), c) for c in []]
    d=[]
    for c in range(3,174):
        a=ws.getCellByPosition(c,newr-1).getValue(); b=ws.getCellByPosition(c,oldr-1).getValue()
        if abs(a-b)>0.5: d.append((ws.getCellByPosition(c,5).getString(),round(a-b,1)))
    print("   新合計−旧合計 当月の差がある月:",len(d)); print("    ",d)
doc.close(True)
