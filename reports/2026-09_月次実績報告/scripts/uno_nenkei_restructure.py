# -*- coding: utf-8 -*-
# 年計表.xls をグラフ10枚（石見・下関・松江・広域・境港・水産部・キョウワ・タカハシ包装営業・タカハシ包装合計・グループ合計）に再構成
import uno, sys
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellAddress, CellRangeAddress
SRC,DST=sys.argv[1],sys.argv[2]
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
ctx=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local).resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
doc=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx).loadComponentFromURL(uno.systemPathToFileUrl(SRC),"_blank",0,(prop("Hidden",True),))
S=doc.Sheets
def col(ci):
    s=""; ci+=1
    while ci: ci,rem=divmod(ci-1,26); s=chr(65+rem)+s
    return s
FIRST,LAST=3,173   # D..FR
# 当月行(1始まり)の構成要素
COMP={"年計表売上DATA":{"石見":[7],"下関":[15],"松江":[27,23],"広域":[31,39,47],"境港":[43],"水産部":[59],"キョウワ":[71]},
      "年計表粗利益DATA":{"石見":[7],"下関":[15],"松江":[27,23],"広域":[31,39],"境港":[43],"水産部":[55],"キョウワ":[67]}}
NEW={}
for sn,comp in COMP.items():
    ws=S.getByName(sn); idx=S.getByName(sn).RangeAddress.Sheet
    cur=ws.createCursor(); cur.gotoEndOfUsedArea(False); start=cur.RangeAddress.EndRow+1   # 0始まり
    blocks=[("広域",comp["広域"]),
            ("タカハシ包装営業",None),("タカハシ包装合計",None),("グループ合計",None)]
    pos={}
    for i,(name,_) in enumerate(blocks):
        t0=start+i*4
        src=CellRangeAddress(); src.Sheet=idx; src.StartColumn=0; src.StartRow=4; src.EndColumn=LAST; src.EndRow=7   # 石見ブロック(見出し・日付・当月・年計の4行)を書式ごと複製
        dst=CellAddress(); dst.Sheet=idx; dst.Column=0; dst.Row=t0
        ws.copyRange(dst,src)
        ws.getCellByPosition(2,t0).setString(name)
        pos[name]=t0+3   # 当月行(1始まり)= t0(0始まり)+3
    rows={"広域":comp["広域"]}
    rows["タカハシ包装営業"]=comp["石見"]+comp["下関"]+comp["松江"]+[pos["広域"]]+comp["境港"]
    rows["タカハシ包装合計"]=[pos["タカハシ包装営業"]]+comp["水産部"]
    rows["グループ合計"]=[pos["タカハシ包装合計"]]+comp["キョウワ"]
    for name,_ in blocks:
        r=pos[name]-1
        for c in range(FIRST,LAST+1):
            ws.getCellByPosition(c,r).setFormula("="+"+".join(f"{col(c)}{x}" for x in rows[name]))
    NEW[sn]={k:v+1 for k,v in pos.items()}   # 年計行(1始まり)
    print(sn,"new blocks 年計行:",NEW[sn])
doc.calculateAll()
# グラフの付け替え
def repoint(sheet,title,srow,grow):
    sh=S.getByName(sheet); c=sh.Charts.getByIndex(0); cd=c.EmbeddedObject; dp=cd.getDataProvider()
    for cs in cd.FirstDiagram.CoordinateSystems:
        for ct in cs.ChartTypes:
            for s in ct.DataSeries:
                for ls in s.DataSequences:
                    rep=ls.Values.SourceRangeRepresentation
                    new=f"$年計表売上DATA.$O${srow}:$FR${srow}" if "売上DATA" in rep else f"$年計表粗利益DATA.$O${grow}:$FR${grow}"
                    seq=dp.createDataSequenceByRangeRepresentation(new); seq.Role="values-y"; ls.setValues(seq)
    if cd.HasMainTitle: cd.Title.String=title
    dg=cd.getDiagram()   # 付け替えたグラフは目盛り上限を自動に（元の固定上限では線が切れるため）
    for ax in (dg.getYAxis(),dg.getSecondaryYAxis()):
        try: ax.setPropertyValue("AutoMax",True); ax.setPropertyValue("AutoStepMain",True)
        except Exception as e: print("axis",sheet,e)
    cd.setModified(True)
SN,GN=NEW["年計表売上DATA"],NEW["年計表粗利益DATA"]
plan=[("【グラフ】鳥栖・静岡","【グラフ】広域","広域","広域"),
      ("【グラフ】東京","【グラフ】タカハシ包装営業","タカハシ包装営業","タカハシ包装営業"),
      ("【グラフ】タカハシ包装合計（境港含）","【グラフ】タカハシ包装合計","タカハシ包装合計","タカハシ包装合計"),
      ("【グラフ】タカハシ包装合計（旧フクダ除く）","【グラフ】グループ合計","グループ合計（タカハシ包装＋キョウワ）","グループ合計")]
for old,new,title,key in plan:
    repoint(old,title,SN[key],GN[key]); S.getByName(old).setName(new)
order=["【グラフ】石見","【グラフ】下関","【グラフ】松江","【グラフ】広域","【グラフ】境港","【グラフ】水産部","【グラフ】キョウワ","【グラフ】タカハシ包装営業","【グラフ】タカハシ包装合計","【グラフ】グループ合計"]
for i,n in enumerate(order): S.moveByName(n,i)
doc.calculateAll()
print("sheets:",S.ElementNames)
doc.storeToURL(uno.systemPathToFileUrl(DST),(prop("FilterName","MS Excel 97"),))
doc.close(True); print("saved",DST)
