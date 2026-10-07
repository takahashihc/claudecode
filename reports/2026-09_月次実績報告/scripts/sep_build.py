import json, copy, sys
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from lxml import etree
S="/tmp/claude-0/-home-user-claudecode/7f1b344f-da50-502a-9e34-aab355afdd38/scratchpad/"
D=json.load(open(S+"sep_final.json"))
O=D["offices"]; OFF=["石見","下関","松江","広域","境港","水産部","キョウワ"]
NS="{http://schemas.openxmlformats.org/drawingml/2006/main}"
def pct(a,b): return a/b*100
def f0(x): return f"{x:,.0f}"
def p1(x): return f"{x:.1f}%"
def sgn(x): return f"+{x:,.0f}" if x>=0 else f"▲{abs(x):,.0f}"
def get(o,m,p,k): return O[o][m][p][k]
def yoy(o,m,p): return pct(get(o,m,p,"act"),get(o,m,p,"prior"))
def vsp(o,m,p): return pct(get(o,m,p,"act"),get(o,m,p,"plan"))
def rate(o,p,k="act"): return pct(get(o,"gp",p,k),get(o,"sales",p,k))
T="合計"
# ---------- key figures
m_gp=get(T,"gp","m","act"); m_sa=get(T,"sales","m","act")
h_gp=get(T,"gp","h2","act"); h_sa=get(T,"sales","h2","act")
y_gp=get(T,"gp","ytd","act"); y_sa=get(T,"sales","ytd","act")
gpdiff={o:get(o,"gp","m","act")-get(o,"gp","m","prior") for o in OFF}
tot_gpdiff=m_gp-get(T,"gp","m","prior")
ex={m:{k:get(T,m,"m",k)-get("境港",m,"m",k) for k in ["act","plan","prior"]} for m in ["sales","gp"]}
ex_sa_yoy=pct(ex["sales"]["act"],ex["sales"]["prior"]); ex_gp_yoy=pct(ex["gp"]["act"],ex["gp"]["prior"]); ex_gp_plan=pct(ex["gp"]["act"],ex["gp"]["plan"])
plus=sorted([o for o in OFF if gpdiff[o]>0],key=lambda o:-gpdiff[o])
minus=sorted([o for o in OFF if gpdiff[o]<0],key=lambda o:gpdiff[o])
def joinc(lst,n=3): return "・".join(f"{o}{sgn(gpdiff[o])}" for o in lst[:n])
print("KEY:",f0(m_gp),p1(yoy(T,'gp','m')),p1(vsp(T,'gp','m')),"| sales",f0(m_sa),p1(yoy(T,'sales','m')),p1(vsp(T,'sales','m')),"| rate",p1(rate(T,'m')),p1(rate(T,'m','prior')),"| ex-sakai",p1(ex_sa_yoy),p1(ex_gp_yoy),p1(ex_gp_plan))
# ---------- helpers
def set_text(shape,text):
    tf=shape.text_frame
    paras=tf.paragraphs
    p=paras[0]
    runs=p.runs
    if runs:
        runs[0].text=text
        for r in runs[1:]: r._r.getparent().remove(r._r)
    else:
        p.text=text
    for extra in paras[1:]:
        extra._p.getparent().remove(extra._p)
def set_cell(cell,text,color=None):
    p=cell.text_frame.paragraphs[0]
    r=p.runs[0]; r.text=text
    for rr in p.runs[1:]: rr._r.getparent().remove(rr._r)
    if color:
        rPr=r._r.get_or_add_rPr()
        sf=rPr.find(NS+"solidFill")
        if sf is not None: rPr.remove(sf)
        sf=etree.SubElement(rPr,NS+"solidFill"); c=etree.SubElement(sf,NS+"srgbClr"); c.set("val",color)
        # solidFill must come before latin/ea: reorder
        for tag in ["latin","ea","cs"]:
            el=rPr.find(NS+tag)
            if el is not None: rPr.remove(el); rPr.append(el)
def ratio_color(v):
    v=round(v,1)
    return "1A6E1A" if v>=100 else ("C55A11" if v>=90 else "C0392B")
def shapes_by_name(slide): return {sh.name:sh for sh in slide.shapes}
prs=Presentation(S+"work/令和8年8月実績報告.pptx")
sl=prs.slides
FOOT="株式会社タカハシ包装センター | 令和8年9月実績"
tr=D["trend"]
# ---------- slide 1 (cover)
s=shapes_by_name(sl[0])
set_text(s["TextBox 3"],"令和8年9月　月次実績報告")
set_text(s["TextBox 4"],"第51期下期 9月実績（下期6ヶ月目・第51期通期確定）　～粗利益額を主軸とした分析～")
set_text(s["TextBox 7"],f"粗利益額: 9月 {f0(m_gp)}千円（前期比{p1(yoy(T,'gp','m'))}・計画比{p1(vsp(T,'gp','m'))}）")
set_text(s["TextBox 8"],f"・ 第51期通期は売上{p1(yoy(T,'sales','ytd'))}・粗利{p1(yoy(T,'gp','ytd'))}（計画比{p1(vsp(T,'gp','ytd'))}）で増収増益着地。下期粗利は前期比{p1(yoy(T,'gp','h2'))}")
set_text(s["TextBox 9"],f"・ 9月は境港{sgn(gpdiff['境港'])}（前期機械大口の反動）を {joinc(plus,3)} が吸収し増益")
set_text(s["TextBox 11"],"【警戒】境港の通期減益と水産部・キョウワの粗利率低下")
set_text(s["TextBox 12"],f"・ 境港は通期粗利前期比{p1(yoy('境港','gp','ytd'))}・計画比{p1(vsp('境港','gp','ytd'))}と唯一の減益。水産部は9月粗利率{p1(rate('水産部','m'))}（前期{p1(rate('水産部','m','prior'))}）と低下")
set_text(s["TextBox 13"],"分析期間: 単月（9月）／下期（4〜9月）／期初来累計（令和7年10月〜令和8年9月＝第51期通期）")
set_text(s["TextBox 14"],"作成日: 令和8年10月7日　　資料作成者: 高橋将史")
# ---------- slide 2 (KPI cards)
s=shapes_by_name(sl[1])
set_text(s["TextBox 5"],FOOT)
cards=[("TextBox 8","TextBox 9","TextBox 11","9月 粗利益額",f0(m_gp),f"前期比 {p1(yoy(T,'gp','m'))} / 計画比 {p1(vsp(T,'gp','m'))}"),
       ("TextBox 14","TextBox 15","TextBox 17","9月 売上高",f0(m_sa),f"前期比 {p1(yoy(T,'sales','m'))} / 計画比 {p1(vsp(T,'sales','m'))}"),
       ("TextBox 20","TextBox 21","TextBox 23","9月 粗利率",f"{rate(T,'m'):.1f}",f"前期 {p1(rate(T,'m','prior'))} → 今期"),
       ("TextBox 26","TextBox 27","TextBox 29","下期 粗利益額",f0(h_gp),f"前期比 {p1(yoy(T,'gp','h2'))} / 計画比 {p1(vsp(T,'gp','h2'))}"),
       ("TextBox 32","TextBox 33","TextBox 35","下期 売上高",f0(h_sa),f"前期比 {p1(yoy(T,'sales','h2'))} / 計画比 {p1(vsp(T,'sales','h2'))}"),
       ("TextBox 38","TextBox 39","TextBox 41","下期 粗利率",f"{rate(T,'h2'):.1f}",f"前期 {p1(rate(T,'h2','prior'))} → 今期"),
       ("TextBox 44","TextBox 45","TextBox 47","通期 粗利益額",f0(y_gp),f"前期比 {p1(yoy(T,'gp','ytd'))} / 計画比 {p1(vsp(T,'gp','ytd'))}"),
       ("TextBox 50","TextBox 51","TextBox 53","通期 売上高",f0(y_sa),f"前期比 {p1(yoy(T,'sales','ytd'))} / 計画比 {p1(vsp(T,'sales','ytd'))}"),
       ("TextBox 56","TextBox 57","TextBox 59","通期 粗利率",f"{rate(T,'ytd'):.1f}",f"前期 {p1(rate(T,'ytd','prior'))} → 今期")]
for a,b,c,t1,t2,t3 in cards:
    set_text(s[a],t1); set_text(s[b],t2); set_text(s[c],t3)
set_text(s["TextBox 3"],"3期間サマリー（単位: 千円）　※粗利益額を主指標／合計は6営業所＋キョウワ／通期＝第51期（令和7年10月〜令和8年9月）")
# ---------- slides 3-5 (tables)
names={"石見":"石見営業所","下関":"下関営業所","松江":"松江営業所","広域":"広域営業部","境港":"境港営業所","水産部":"水産部（計）","キョウワ":"キョウワ"}
def fill_table(slide,period):
    t=[sh for sh in slide.shapes if sh.has_table][0].table
    for ri,o in enumerate(OFF+[T],start=1):
        total=(o==T)
        sa=O[o]["sales"][period]; gp=O[o]["gp"][period]
        vals=[names.get(o,"合　計"),f0(sa["act"]),pct(sa["act"],sa["prior"]),pct(sa["act"],sa["plan"]),f0(gp["act"]),pct(gp["act"],gp["prior"]),pct(gp["act"],gp["plan"]),pct(gp["act"],sa["act"])]
        set_cell(t.cell(ri,0),vals[0]); set_cell(t.cell(ri,1),vals[1]); set_cell(t.cell(ri,4),vals[4])
        for ci in [2,3,5,6]:
            set_cell(t.cell(ri,ci),p1(vals[ci]),None if total else ratio_color(vals[ci]))
        set_cell(t.cell(ri,7),p1(vals[7]))
s=shapes_by_name(sl[2]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 2"],"営業所別　9月単月実績　～売上・粗利益額～")
fill_table(sl[2],"m")
set_text(s["TextBox 7"],f"※ 粗利益額{f0(m_gp)}千円は前期比{p1(yoy(T,'gp','m'))}・計画比{p1(vsp(T,'gp','m'))}。減益は前期に機械大口（岡野農場向け29,400千円）のあった境港（{sgn(gpdiff['境港'])}）のみで、他6部門は全て増益。境港除きは売上{p1(ex_sa_yoy)}・粗利{p1(ex_gp_yoy)}")
s=shapes_by_name(sl[3]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 2"],"営業所別　下期（4〜9月）累計　～売上・粗利益額～")
set_text(s["TextBox 3"],"単位: 千円　　第51期下期（4〜9月）確定　　金色列＝粗利益額（主指標）")
fill_table(sl[3],"h2")
set_text(s["TextBox 7"],f"※ 下期粗利益額{f0(h_gp)}千円・前期比{p1(yoy(T,'gp','h2'))}・計画比{p1(vsp(T,'gp','h2'))}（売上{p1(yoy(T,'sales','h2'))}・{p1(vsp(T,'sales','h2'))}）。境港は前期比{p1(yoy('境港','gp','h2'))}・計画比{p1(vsp('境港','gp','h2'))}、キョウワは前期比{p1(yoy('キョウワ','gp','h2'))}と前期を下回る")
s=shapes_by_name(sl[4]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 2"],"営業所別　期初来累計（第51期通期）　～売上・粗利益額～")
set_text(s["TextBox 3"],"単位: 千円　　令和7年10月〜令和8年9月（12ヶ月・第51期通期）　　金色列＝粗利益額（主指標）")
fill_table(sl[4],"ytd")
set_text(s["TextBox 7"],f"※ 第51期通期の粗利益額{f0(y_gp)}千円・前期比{p1(yoy(T,'gp','ytd'))}・計画比{p1(vsp(T,'gp','ytd'))}（売上{p1(yoy(T,'sales','ytd'))}・計画比{p1(vsp(T,'sales','ytd'))}）。境港（{p1(yoy('境港','gp','ytd'))}）・キョウワ（{p1(yoy('キョウワ','gp','ytd'))}）以外の5部門が増益")
# ---------- slide 6/7 charts
def replace_chart(slide,series):
    ch=[sh for sh in slide.shapes if sh.has_chart][0].chart
    cd=CategoryChartData(); cd.categories=OFF
    for nm,vals in series: cd.add_series(nm,vals)
    ch.replace_data(cd)
    return ch
ser6=[("単月9月",[round(yoy(o,"gp","m"),1) for o in OFF]),("下期4〜9月",[round(yoy(o,"gp","h2"),1) for o in OFF]),("通期10〜9月",[round(yoy(o,"gp","ytd"),1) for o in OFF])]
ch=replace_chart(sl[5],ser6)
mx=max(v for _,vs in ser6 for v in vs)
axmax=250.0 if mx>200 else 200.0
for el in ch._chartSpace.iter("{http://schemas.openxmlformats.org/drawingml/2006/chart}max"): el.set("val",str(axmax))
s=shapes_by_name(sl[5]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 7"],f"全社の粗利益額前期比は 単月{p1(yoy(T,'gp','m'))}／下期{p1(yoy(T,'gp','h2'))}／通期{p1(yoy(T,'gp','ytd'))}。単月は境港のみ、下期・通期は境港とキョウワが前期を下回る（境港通期{p1(yoy('境港','gp','ytd'))}・キョウワ通期{p1(yoy('キョウワ','gp','ytd'))}）")
ser7=[("単月9月",[round(gpdiff[o]) for o in OFF]),("下期4〜9月",[round(get(o,"gp","h2","act")-get(o,"gp","h2","prior")) for o in OFF]),("通期10〜9月",[round(get(o,"gp","ytd","act")-get(o,"gp","ytd","prior")) for o in OFF])]
ch=replace_chart(sl[6],ser7)
s=shapes_by_name(sl[6]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 7"],f"9月単月の粗利は{sgn(tot_gpdiff)}千円。境港{sgn(gpdiff['境港'])}（前期岡野農場向け機械大口の反動）を {joinc(plus,4)} の増益で吸収。通期では水産部{sgn(get('水産部','gp','ytd','act')-get('水産部','gp','ytd','prior'))}が最大の寄与")
# ---------- slide 8 narrative
s=shapes_by_name(sl[7]); set_text(s["TextBox 5"],FOOT)
set_text(s["TextBox 8"],f"第51期通期は粗利益額 前期比{p1(yoy(T,'gp','ytd'))}・計画比{p1(vsp(T,'gp','ytd'))}で増益着地")
set_text(s["TextBox 9"],f"・ 9月粗利益額 {f0(m_gp)}千円（前期比{p1(yoy(T,'gp','m'))}・計画比{p1(vsp(T,'gp','m'))}）。売上は前期比{p1(yoy(T,'sales','m'))}・計画比{p1(vsp(T,'sales','m'))}と横ばいで、増益は粗利率改善による")
set_text(s["TextBox 10"],f"・ 通期売上{f0(y_sa)}千円（前期比{p1(yoy(T,'sales','ytd'))}）・粗利{f0(y_gp)}千円（{p1(yoy(T,'gp','ytd'))}）。下期粗利は前期比{p1(yoy(T,'gp','h2'))}と上期から加速")
set_text(s["TextBox 11"],f"・ 粗利率は単月{p1(rate(T,'m'))}（前期{p1(rate(T,'m','prior'))}）・下期{p1(rate(T,'h2'))}（{p1(rate(T,'h2','prior'))}）・通期{p1(rate(T,'ytd'))}（{p1(rate(T,'ytd','prior'))}）と全期間で改善")
set_text(s["TextBox 14"],"【警戒】境港の通期減益と、売上増でも粗利率が低下した部門")
set_text(s["TextBox 15"],f"・ 境港は9月粗利前期比{p1(yoy('境港','gp','m'))}（前期の岡野農場向け機械大口の反動）。通期も売上{p1(yoy('境港','sales','ytd'))}・粗利{p1(yoy('境港','gp','ytd'))}・計画比{p1(vsp('境港','gp','ytd'))}と唯一の減収減益")
set_text(s["TextBox 16"],f"・ 全社粗利前期比は7月{tr['7月']['gp_yoy']:.1f}%→8月{tr['8月']['gp_yoy']:.1f}%→9月{yoy(T,'gp','m'):.1f}%。石見は粗利{p1(yoy('石見','gp','m'))}へ回復も売上計画比{p1(vsp('石見','sales','m'))}と未達")
set_text(s["TextBox 17"],f"・ 水産部は9月粗利率{p1(rate('水産部','m'))}（前期{p1(rate('水産部','m','prior'))}）、キョウワは通期粗利{p1(yoy('キョウワ','gp','ytd'))}。売上増でも利幅が落ちる先は値決めを点検")
set_text(s["TextBox 20"],"営業所別の実態　～9月の増益はどこから来たか～")
set_text(s["TextBox 21"],f"・ 広域：粗利前期比{p1(yoy('広域','gp','m'))}（リンガーハット東京本社 粗利7,652千円・前期4,401）、松江：{p1(yoy('松江','gp','m'))}（コクヨー・岡田商店・ヤマダヤ）")
set_text(s["TextBox 22"],f"・ 石見{sgn(gpdiff['石見'])}（キヌヤ+1,667）、下関{sgn(gpdiff['下関'])}（フクシン・もずくセンター）。境港を除く6部門が増益、粗利率も全社で{rate(T,'m')-rate(T,'m','prior'):+.1f}pt")
set_text(s["TextBox 24"],f"第51期は粗利益額{f0(y_gp)}千円で増益着地。第52期は境港の立て直しと、価格改定効果一巡後の粗利率維持が焦点")

import unicodedata
def w(t): return sum(1 if unicodedata.east_asian_width(c) in "FWA" else 0.5 for c in t)
limits={(1,"TextBox 8"):56,(1,"TextBox 9"):56,(1,"TextBox 12"):56,(1,"TextBox 7"):45,(1,"TextBox 4"):48,(1,"TextBox 13"):60,(8,"TextBox 9"):62,(8,"TextBox 10"):62,(8,"TextBox 11"):62,(8,"TextBox 15"):62,(8,"TextBox 16"):62,(8,"TextBox 17"):62,(8,"TextBox 21"):62,(8,"TextBox 22"):62,(8,"TextBox 24"):64,(8,"TextBox 8"):48,(8,"TextBox 14"):48,(3,"TextBox 7"):130,(4,"TextBox 7"):130,(5,"TextBox 7"):130,(6,"TextBox 7"):130,(7,"TextBox 7"):130,(2,"TextBox 3"):80}
for (si,nm),lim in limits.items():
    t=shapes_by_name(sl[si-1])[nm].text_frame.text
    flag="OK " if w(t)<=lim else "OVER"
    print(f"{flag} slide{si} {nm} width={w(t):.1f}/{lim}: {t}")
prs.save(S+"out/令和8年9月実績報告.pptx")
print("saved")
