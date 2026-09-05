import openpyxl, json
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
OFF=["石見","下関","松江","広域","境港","水産部","キョウワ"]
prev=json.load(open(S+"aug_numbers.json"))
data=prev["data"]; h1=prev["h1"]
# Apr-Jul cumulative from 営業会議資料 7月 下期累計 sheet (source of July report)
wb=openpyxl.load_workbook(S+"【営業会議資料】令和8年7月実績.xlsx",data_only=True); ws=wb["下期累計"]
lab={"石見営業所":"石見","下関営業所":"下関","松江営業所":"松江","広域営業部":"広域","境港営業所":"境港","水産部":"水産部","キョウワ（税抜）":"キョウワ"}
h2_47={}
for row in ws.iter_rows(values_only=True):
    labs=[str(v).strip() for v in row[:4] if isinstance(v,str)]
    if labs and labs[0] in lab:
        o=lab[labs[0]]
        if o in h2_47: continue
        h2_47[o]={"sales":{"prior":float(row[6]),"plan":float(row[7]),"act":float(row[8])},"gp":{"prior":float(row[13]),"plan":float(row[14]),"act":float(row[15])}}
print("下期4-7月 from 営業会議資料:",{o:(round(v["sales"]["act"]),round(v["gp"]["act"])) for o,v in h2_47.items()})
# compare with office-file sums
for o in OFF:
    fs=sum(data[o]["act"]["sales"][:4]); fg=sum(data[o]["act"]["gp"][:4])
    print(f"  {o}: file Apr-Jul {fs:.1f}/{fg:.1f}  kaigi {h2_47[o]['sales']['act']:.1f}/{h2_47[o]['gp']['act']:.1f}  plan file {sum(data[o]['plan']['sales'][:4]):.0f} kaigi {h2_47[o]['sales']['plan']:.0f}  prior file {sum(data[o]['prior']['sales'][:4]):.1f} kaigi {h2_47[o]['sales']['prior']:.1f}")
out={"offices":{}}
for o in OFF:
    res={}
    for m in ["sales","gp"]:
        mth={k:data[o][k][m][4] for k in ["act","plan","prior"]}
        h2={k:h2_47[o][m][k]+mth[k] for k in ["act","plan","prior"]}
        ytd={k:h2[k]+h1[o][m][k] for k in ["act","plan","prior"]}
        res[m]={"m":mth,"h2":h2,"ytd":ytd}
    out["offices"][o]=res
tot={m:{p:{k:sum(out["offices"][o][m][p][k] for o in OFF) for k in ["act","plan","prior"]} for p in ["m","h2","ytd"]} for m in ["sales","gp"]}
out["offices"]["合計"]=tot
# excl Sakaiminato for narrative
ex={m:{k:tot[m]["m"][k]-out["offices"]["境港"][m]["m"][k] for k in ["act","plan","prior"]} for m in ["sales","gp"]}
out["excl_sakai_m"]=ex
out["trend"]=prev["trend"]
json.dump(out,open(S+"aug_final.json","w"),ensure_ascii=False,indent=1)
def pct(a,b): return a/b*100
for p,lab_ in [("m","8月単月"),("h2","下期4〜8月"),("ytd","期初来10〜8月")]:
    print(f"\n=== {lab_} ===")
    for o in OFF+["合計"]:
        s=out["offices"][o]["sales"][p]; g=out["offices"][o]["gp"][p]
        print(f"{o:6}{s['act']:11,.0f}{pct(s['act'],s['prior']):7.1f}{pct(s['act'],s['plan']):7.1f} |{g['act']:9,.0f}{pct(g['act'],g['prior']):7.1f}{pct(g['act'],g['plan']):7.1f} | rate {pct(g['act'],s['act']):5.1f} (prior {pct(g['prior'],s['prior']):5.1f}) diff {g['act']-g['prior']:+,.0f}  plan_abs s{s['plan']:,.0f} g{g['plan']:,.0f} prior s{s['prior']:,.0f} g{g['prior']:,.0f}")
print("\n境港除き 8月:",{m:{k:round(v) for k,v in d.items()} for m,d in ex.items()}, "sales yoy",pct(ex["sales"]["act"],ex["sales"]["prior"]),"gp yoy",pct(ex["gp"]["act"],ex["gp"]["prior"]),"gp plan",pct(ex["gp"]["act"],ex["gp"]["plan"]))
