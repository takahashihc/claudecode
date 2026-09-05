import openpyxl, json
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
OFF=["石見","下関","松江","広域","境港","水産部","キョウワ"]
files={"石見":"iwami_shimoki.xlsx","下関":"第51期下期下関営業所担当者別実績.xlsx","松江":"第51期下期松江営業所担当者別実績.xlsx","広域":"第51期下期広域営業部担当者別実績.xlsx","境港":"第51期下期境港営業所担当者別実績.xlsx","水産部":"第51期下期水産部部門別実績.xlsx"}
codes={"石見":["0011","0012","0013","0014","0015","0016","0018","0019","0020"],
       "下関":["0021","0024","0025","0026","0027","0028","0029"],
       "松江":["0032","0033","0034","0035","0039"],
       "広域":["0042","0045","0046","0047","0550","0552"],
       "境港":["0903","0904","0909","0910","0911","0914","0918","0919","0930"],
       "水産部":["0010","0017","0031","0038","0101","0102","0104","0106"]}
# --- raw Aug data
raw={}
ws=openpyxl.load_workbook(S+"担当者別売上粗利益実績202608.xlsx",data_only=True).worksheets[0]
for row in ws.iter_rows(min_row=2,values_only=True):
    if row[0] is None: continue
    code=str(row[0]).zfill(4)
    raw[code]={"name":row[1],"sales":[float(x or 0) for x in row[2:7]],"gp":[float(x or 0) for x in row[7:12]]}
allcodes=set(sum(codes.values(),[]))
unmapped={c:v for c,v in raw.items() if c not in allcodes and c!="0502" and (any(v["sales"]) or any(v["gp"]))}
print("UNMAPPED codes with data:",{c:(v["name"],v["sales"][4],v["gp"][4]) for c,v in unmapped.items()})
aug={}; jul_raw={}
for o,cs in codes.items():
    aug[o]={"sales":sum(raw[c]["sales"][4] for c in cs if c in raw),"gp":sum(raw[c]["gp"][4] for c in cs if c in raw)}
    jul_raw[o]={"sales":sum(raw[c]["sales"][3] for c in cs if c in raw),"gp":sum(raw[c]["gp"][3] for c in cs if c in raw)}
# --- office files
def find_section(rows,prefix):
    for i,r in enumerate(rows):
        if r[2] and str(r[2]).startswith(prefix): return i
    raise Exception(prefix)
def total_row(rows,start,label="合計"):
    for r in rows[start:start+40]:
        if r[2]==label: return r
    raise Exception(label)
data={}
for o,f in files.items():
    wb=openpyxl.load_workbook(S+f,data_only=True); ws=wb["全体"]; rows=list(ws.iter_rows(values_only=True))
    d={}
    for key,pre in [("prior","2．"),("plan","4．"),("act","7．")]:
        st=find_section(rows,pre); tr=total_row(rows,st)
        # columns: C=label(2) D..I months(3..8) J total(9) K..P GP months(10..15)
        d[key]={"sales":[float(tr[4+i] or 0) for i in range(6)],"gp":[float(tr[11+i] or 0) for i in range(6)]}
    data[o]=d
# Kyowa
wb=openpyxl.load_workbook(S+files["広域"],data_only=True); ws=wb["キョウワ"]; rows=list(ws.iter_rows(values_only=True))
def ky(prefix,label="合計(税抜）"):
    st=find_section(rows,prefix)
    for r in rows[st:st+8]:
        if r[2]==label: return {"sales":[float(r[4+i] or 0) for i in range(6)],"gp":[float(r[11+i] or 0) for i in range(6)]}
data["キョウワ"]={"prior":ky("2．"),"plan":ky("4．"),"act":ky("7．")}
# Check July consistency: raw July vs office file July actual
for o in files:
    print(f"JUL check {o}: raw {jul_raw[o]['sales']:.3f}/{jul_raw[o]['gp']:.3f} vs file {data[o]['act']['sales'][3]:.3f}/{data[o]['act']['gp'][3]:.3f}")
# insert Aug actual
for o in files:
    data[o]["act"]["sales"][4]=aug[o]["sales"]; data[o]["act"]["gp"][4]=aug[o]["gp"]
KYOWA_AUG=True
print("Kyowa sheet act Apr-Jul:",data["キョウワ"]["act"]["sales"][:4],"raw 0502:",raw["0502"]["sales"])
for i in range(5):
    data["キョウワ"]["act"]["sales"][i]=raw["0502"]["sales"][i]; data["キョウワ"]["act"]["gp"][i]=raw["0502"]["gp"][i]
# --- H1 from March file
wb=openpyxl.load_workbook(S+"【営業会議資料】令和8年3月実績4月見込と累計.xlsx",data_only=True); ws=wb["累計実績"]
h1rows={"石見":25,"下関":48,"松江":65,"広域":82,"境港":109,"水産部":117,"キョウワ":125}
h1={}
for o,r in h1rows.items():
    g=lambda c: float(ws.cell(r,c).value or 0)
    h1[o]={"sales":{"prior":g(91),"plan":g(92),"act":g(93)},"gp":{"prior":g(98),"plan":g(99),"act":g(100)}}
# --- build periods
out={"offices":{}, "kyowa_aug_provided": bool(KYOWA_AUG)}
def per(o):
    d=data[o]; res={}
    for m in ["sales","gp"]:
        res[m]={
          "m":{k:d[k][m][4] for k in ["act","plan","prior"]},
          "h2":{k:sum(d[k][m][0:5]) for k in ["act","plan","prior"]},
          "ytd":{k:sum(d[k][m][0:5])+h1[o][m][k] for k in ["act","plan","prior"]},
        }
    return res
for o in OFF: out["offices"][o]=per(o)
tot={}
for m in ["sales","gp"]:
    tot[m]={p:{k:sum(out["offices"][o][m][p][k] for o in OFF) for k in ["act","plan","prior"]} for p in ["m","h2","ytd"]}
out["offices"]["合計"]=tot
# also June/July GP yoy for trend
for o in OFF+["合計"]:
    pass
trend={}
for i,mn in enumerate(["4月","5月","6月","7月","8月"]):
    a=sum(data[o]["act"]["gp"][i] for o in OFF); p=sum(data[o]["prior"]["gp"][i] for o in OFF)
    sa=sum(data[o]["act"]["sales"][i] for o in OFF); sp=sum(data[o]["prior"]["sales"][i] for o in OFF)
    trend[mn]={"gp_yoy":a/p*100,"sales_yoy":sa/sp*100,"gp":a,"sales":sa,"gp_rate":a/sa*100,"gp_rate_prior":p/sp*100}
out["trend"]=trend
out["raw_person_aug"]={c:{"name":v["name"],"sales":v["sales"][4],"gp":v["gp"][4]} for c,v in raw.items()}
out["data"]=data; out["h1"]=h1
json.dump(out,open(S+"aug_numbers.json","w"),ensure_ascii=False,indent=1)
# --- print summary table
def pct(a,b): return a/b*100 if b else float('nan')
for p,lab in [("m","8月単月"),("h2","下期4〜8月"),("ytd","期初来10〜8月")]:
    print(f"\n=== {lab} ===")
    print(f"{'営業所':8}{'売上':>10}{'前期比':>8}{'計画比':>8}{'粗利':>10}{'前期比':>8}{'計画比':>8}{'粗利率':>7}{'前期粗利率':>8}  粗利前期差")
    for o in OFF+["合計"]:
        s=out["offices"][o]["sales"][p]; g=out["offices"][o]["gp"][p]
        print(f"{o:8}{s['act']:10,.0f}{pct(s['act'],s['prior']):8.1f}{pct(s['act'],s['plan']):8.1f}{g['act']:10,.0f}{pct(g['act'],g['prior']):8.1f}{pct(g['act'],g['plan']):8.1f}{pct(g['act'],s['act']):7.1f}{pct(g['prior'],s['prior']):8.1f}  {g['act']-g['prior']:+,.0f}")
print("\nTREND:",{k:{kk:round(vv,1) for kk,vv in v.items()} for k,v in trend.items()})
print("\n境港 2025 Aug person breakdown / 8月 prior big:")
