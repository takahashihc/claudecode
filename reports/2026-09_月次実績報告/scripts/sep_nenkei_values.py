# -*- coding: utf-8 -*-
import json, openpyxl
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
D=json.load(open(S+"sep_final.json"))["offices"]
ws=openpyxl.load_workbook(S+"sep/2609担当者別売上粗利益実績.xlsx",data_only=True).worksheets[0]; R={}
for row in ws.iter_rows(min_row=2,values_only=True):
    if row[0] is None: continue
    R[str(row[0]).zfill(4)]={"sales":[float(x or 0) for x in row[2:8]],"gp":[float(x or 0) for x in row[8:14]]}
codes={"石見":["0011","0012","0013","0014","0015","0016","0018","0019","0020"],"下関":["0021","0024","0025","0026","0027","0028","0029"],"松江":["0032","0033","0034","0035","0039"],"広域":["0042","0045","0046","0047","0550","0552"],"境港":["0903","0904","0909","0910","0911","0914","0918","0919","0930"],"水産部":["0010","0017","0031","0038","0101","0102","0104","0106"]}
def g(cs,k,i): return sum(R[c][k][i] for c in cs if c in R)
out={"年計表売上DATA":{},"年計表粗利益DATA":{}}
for i in range(6):
    sales={"石見":g(codes["石見"],"sales",i),"下関":g(codes["下関"],"sales",i),"松江":g(codes["松江"],"sales",i),"鳥栖・静岡":g(["0042","0047","0046"],"sales",i),"境港":g(codes["境港"],"sales",i),"東京：②":g(["0550","0552"],"sales",i),"水産部":g(codes["水産部"],"sales",i),"キョウワ":round(g(["0502"],"sales",i))}
    gp={"石見":g(codes["石見"],"gp",i),"下関":g(codes["下関"],"gp",i),"松江(内JF抜き）":g(codes["松江"],"gp",i),"鳥栖＋東京①":g(["0042","0047","0046"],"gp",i),"東京":g(["0550","0552"],"gp",i),"境港":g(codes["境港"],"gp",i),"水産部":g(codes["水産部"],"gp",i),"キョウワ":round(g(["0502"],"gp",i))}
    out["年計表売上DATA"][f"2026/{i+4}"]={k:round(v,3) for k,v in sales.items()}; out["年計表粗利益DATA"][f"2026/{i+4}"]={k:round(v,3) for k,v in gp.items()}
json.dump(out,open(S+"scripts/nenkei_sep_values.json","w"),ensure_ascii=False,indent=1)
print({k:v["2026/9"] for k,v in out.items()})
