# -*- coding: utf-8 -*-
import json, openpyxl
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
D=json.load(open(S+"sep_final.json"))["offices"]
ws=openpyxl.load_workbook(S+"sep/2609担当者別売上粗利益実績.xlsx",data_only=True).worksheets[0]; R={}
for row in ws.iter_rows(min_row=2,values_only=True):
    if row[0] is None: continue
    R[str(row[0]).zfill(4)]={"sales":float(row[7] or 0),"gp":float(row[13] or 0)}
def m(o,k): return D[o][k]["m"]["act"]
tosu={k:sum(R[c][k] for c in ["0042","0047","0046"]) for k in ["sales","gp"]}
tokyo={k:sum(R[c][k] for c in ["0550","0552"]) for k in ["sales","gp"]}
sales={"石見":m("石見","sales"),"下関":m("下関","sales"),"松江":m("松江","sales"),"鳥栖・静岡":tosu["sales"],"境港":m("境港","sales"),"東京：②":tokyo["sales"],"水産部":m("水産部","sales"),"キョウワ":round(m("キョウワ","sales"))}
sales["境港+東京"]=sales["境港"]+sales["東京：②"]; sales["営業（境港除）"]=sales["石見"]+sales["下関"]+sales["松江"]+sales["鳥栖・静岡"]
sales["合計旧タカハシ包装"]=sales["営業（境港除）"]+sales["水産部"]; sales["合計（境港含）"]=sales["合計旧タカハシ包装"]+sales["境港+東京"]
gp={"石見":m("石見","gp"),"下関":m("下関","gp"),"松江(内JF抜き）":m("松江","gp"),"鳥栖＋東京①":tosu["gp"],"東京":tokyo["gp"],"境港":m("境港","gp"),"水産部":m("水産部","gp"),"キョウワ":round(m("キョウワ","gp"))}
gp["境港＋東京"]=gp["境港"]+gp["東京"]; gp["営業（旧フクダ除）"]=gp["石見"]+gp["下関"]+gp["松江(内JF抜き）"]+gp["鳥栖＋東京①"]+gp["東京"]
gp["旧タカハシ包装合計"]=gp["営業（旧フクダ除）"]+gp["水産部"]; gp["境港含む"]=gp["旧タカハシ包装合計"]+gp["境港"]
json.dump({"年計表売上DATA":{k:round(v,3) for k,v in sales.items()},"年計表粗利益DATA":{k:round(v,3) for k,v in gp.items()}},open(S+"scripts/nenkei_sep_values.json","w"),ensure_ascii=False,indent=1)
print("SALES:",{k:round(v,1) for k,v in sales.items()}); print("GP:",{k:round(v,1) for k,v in gp.items()})
