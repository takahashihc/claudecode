# -*- coding: utf-8 -*-
# Run with: python3 uno_nenkei.py  (needs python3-uno; soffice listening on socket)
import uno, json, sys, time, subprocess, os
from com.sun.star.beans import PropertyValue
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
vals=json.load(open(S+"nenkei_aug_values.json"))  # {"年計表売上DATA":{label:value}, "年計表粗利益DATA":{...}}
def prop(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
resolver=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
ctx=None
for i in range(60):
    try:
        ctx=resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
if ctx is None: sys.exit("no soffice")
smgr=ctx.ServiceManager; desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
doc=desktop.loadComponentFromURL(uno.systemPathToFileUrl(S+"upload_nenkei.xls"),"_blank",0,(prop("Hidden",True),))
log=[]
for shname,d in vals.items():
    sh=doc.Sheets.getByName(shname)
    # header row index 5, find columns
    c7=c8=None
    for c in range(0,300):
        v=sh.getCellByPosition(c,5).getString()
        if v=="2026/7": c7=c
        if v=="2026/8": c8=c
        if c7 is not None and c8 is not None: break
    cur=sh.createCursor(); cur.gotoEndOfUsedArea(False); nrows=cur.RangeAddress.EndRow+1
    block=None
    for r in range(3,nrows):
        lab=sh.getCellByPosition(2,r).getString()
        if lab and lab not in ("当月","年計","売上","粗利益") and not lab.startswith("20"): block=lab
        cell7=sh.getCellByPosition(c7,r); cell8=sh.getCellByPosition(c8,r)
        t=cell7.getType().value  # EMPTY, VALUE, TEXT, FORMULA
        if t=="EMPTY": continue
        if t=="FORMULA":
            src=cell7.getRangeAddress(); dst=cell8.getCellAddress()
            sh.copyRange(dst,src); log.append(f"{shname} r{r} {block}/{lab}: formula {cell7.getFormula()} -> {cell8.getFormula()}")
        elif t=="VALUE":
            if block in d and lab=="当月":
                cell8.setValue(float(d[block])); log.append(f"{shname} r{r} {block}/{lab}: value {d[block]}")
            elif lab=="年計":
                # 年計 as value -> write rolling 12-month SUM formula
                from_col=c8-11
                def colname(ci):
                    s=""; ci+=1
                    while ci: ci,rem=divmod(ci-1,26); s=chr(65+rem)+s
                    return s
                cell8.setFormula(f"=SUM({colname(from_col)}{r+1}:{colname(c8)}{r+1})"); log.append(f"{shname} r{r} {block}/{lab}: 年計 formula set")
            else:
                log.append(f"{shname} r{r} {block}/{lab}: VALUE row not in dict -> left blank")
        else:
            log.append(f"{shname} r{r} {block}/{lab}: text {cell7.getString()} skipped")
doc.calculateAll()
print("\n".join(log))
out="/tmp/nenkei_out.xls"
doc.storeToURL(uno.systemPathToFileUrl(out),(prop("FilterName","MS Excel 97"),))
doc.close(True)
print("saved",out)
