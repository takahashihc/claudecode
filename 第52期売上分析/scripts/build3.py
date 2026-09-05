# -*- coding: utf-8 -*-
import json, re, copy
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

NAVY=RGBColor(0x1F,0x38,0x64); CHAR=RGBColor(0x40,0x40,0x40); GREY=RGBColor(0xF2,0xF2,0xF2); WHITE=RGBColor(0xFF,0xFF,0xFF)
GOLD=RGBColor(0xB8,0x86,0x0B); BLUE=RGBColor(0x44,0x72,0xC4); RED=RGBColor(0xC0,0x00,0x00); LIGHT=RGBColor(0xF7,0xF9,0xFC); PALE=RGBColor(0xE8,0xEE,0xF7)
br=json.load(open('bridge.json')); aug=json.load(open('aug_result.json'))
V5=Path('in/v5x/ppt/charts')
def ser_vals(n):
    x=(V5/f'chart{n}.xml').read_text(encoding='utf-8'); out=[]
    for s in re.findall(r'<c:ser>.*?</c:ser>',x,re.S):
        vals=re.findall(r'<c:val>.*?</c:val>',s,re.S)[0]; out.append([float(v) for v in re.findall(r'<c:v>(.*?)</c:v>',vals)])
    return out
OFFICES=['石見営業所','下関営業所','松江営業所','広域営業部','境港営業所','水産部']; OCHART=[16,17,18,19,20,21]
price_by_office={o:[round(x,3) for x in ser_vals(c)[0]]+[round(aug[f'{o}_8']['単価効果'],3)] for o,c in zip(OFFICES,OCHART)}

prs=Presentation('stage2.pptx'); S=prs.slides
def by_name(slide): return {sh.name:sh for sh in slide.shapes}
def set_text(shape,text):
    tf=shape.text_frame; p0=tf.paragraphs[0]
    for p in list(tf.paragraphs[1:]): p._p.getparent().remove(p._p)
    runs=p0.runs
    if not runs: p0.text=text; return
    runs[0].text=text
    for r in runs[1:]: r._r.getparent().remove(r._r)
def new_slide(template_idx, chapter, title, sub, source):
    src=S[template_idx]; ns=prs.slides.add_slide(src.slide_layout)
    for ph in list(ns.placeholders): ph._element.getparent().remove(ph._element)
    sh=by_name(src)
    for nm in ['Rectangle 1','Rectangle 2','TextBox 3','TextBox 4','TextBox 5','TextBox 7']:
        ns.shapes._spTree.append(copy.deepcopy(sh[nm]._element))
    srcname='TextBox 9' if 'TextBox 9' in sh else [n for n in sh if n.startswith('TextBox') and '出所' in sh[n].text_frame.text][0]
    ns.shapes._spTree.append(copy.deepcopy(sh[srcname]._element))
    pic=sh['Picture 6']; ns.shapes.add_picture('in/kx/ppt/media/image1.jpg',pic.left,pic.top,pic.width,pic.height)
    n=by_name(ns); set_text(n['TextBox 3'],chapter); set_text(n['TextBox 4'],title); set_text(n['TextBox 5'],sub); set_text(n[srcname],source)
    return ns
def run(p,text,size=10,bold=False,color=CHAR):
    r=p.add_run(); r.text=text; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name='游ゴシック'; return r
def textbox(slide,x,y,w,h,paras,anchor=MSO_ANCHOR.TOP):
    t=slide.shapes.add_textbox(x,y,w,h); tf=t.text_frame; tf.word_wrap=True; tf.margin_left=tf.margin_right=Inches(0.05); tf.margin_top=tf.margin_bottom=Inches(0.03); tf.vertical_anchor=anchor
    first=True
    for para in paras:
        p=tf.paragraphs[0] if first else tf.add_paragraph(); first=False
        if isinstance(para,tuple): txt,kw=para
        else: txt,kw=para,{}
        sp=kw.pop('space_before',3); p.space_before=Pt(sp)
        bullet=kw.pop('bullet',False)
        run(p,('・' if bullet else '')+txt,**kw)
    return t
def card(slide,x,y,w,h,head,paras,accent,fill=LIGHT):
    r=slide.shapes.add_shape(1,x,y,w,h); r.fill.solid(); r.fill.fore_color.rgb=fill; r.line.fill.background(); r.shadow.inherit=False
    t=slide.shapes.add_textbox(x+Inches(0.15),y+Inches(0.08),w-Inches(0.3),h-Inches(0.16)); tf=t.text_frame; tf.word_wrap=True; tf.margin_left=tf.margin_right=0
    run(tf.paragraphs[0],head,12,True,accent)
    for para in paras:
        p=tf.add_paragraph(); p.space_before=Pt(4)
        if isinstance(para,tuple): run(p,para[0],**para[1])
        else: run(p,para,10)
    return r
def pct(v,d=1): return ('＋' if v>=0 else '▲')+f'{abs(v)*100:.{d}f}％'
def cell(c,text,size=9,bold=False,color=CHAR,align=PP_ALIGN.RIGHT,fill=None):
    c.text=''; p=c.text_frame.paragraphs[0]; run(p,text,size,bold,color); p.alignment=align
    c.margin_left=c.margin_right=Inches(0.05); c.margin_top=c.margin_bottom=Inches(0.02); c.vertical_anchor=MSO_ANCHOR.MIDDLE
    if fill is not None: c.fill.solid(); c.fill.fore_color.rgb=fill
    else: c.fill.background()
def table(slide,x,y,w,widths,rows,hdr,row_h=0.28,hdr_h=0.36,size=9,bold_last=False):
    n=len(rows)+1; gs=slide.shapes.add_table(n,len(hdr),x,y,w,Inches(row_h)*n); t=gs.table
    tp=t._tbl.tblPr; tp.set('bandRow','0'); tp.set('firstRow','0'); st=tp.find(qn('a:tableStyleId'))
    if st is not None: tp.remove(st)
    tot=sum(widths)
    for j,cw in enumerate(widths): t.columns[j].width=Emu(int(w*cw/tot))
    for j,h in enumerate(hdr): cell(t.cell(0,j),h,8,True,WHITE,PP_ALIGN.CENTER,NAVY)
    t.rows[0].height=Inches(hdr_h)
    for i,row in enumerate(rows,1):
        last=bold_last and i==len(rows); fill=PALE if last else (GREY if i%2==0 else None)
        for j,v in enumerate(row):
            if isinstance(v,tuple): txt,kw=v
            else: txt,kw=v,{}
            cell(t.cell(i,j),txt,kw.get('size',size),kw.get('bold',last),kw.get('color',CHAR),kw.get('align',PP_ALIGN.LEFT if j==0 else PP_ALIGN.RIGHT),fill)
        t.rows[i].height=Inches(row_h)
    return gs

# ============ P10: 8月の売上増減ブリッジ（第3章）
M=1e6
S0,S1,m0=br['S0']/M,br['S1']/M,br['m0']/M
pe=br['pe']*m0; qe=br['qe']*m0; out0=-(br['only0']+br['ex0'])/M; in1=(br['only1']+br['ex1'])/M
s=new_slide(5,'第3章　当社の実績','8月の売上増減の内訳 ―― 突合の外側にある「品目入替」【機械類を除く】',
            '突合できた品目は単価・数量に分解。突合できない品目（終売・新規・仕様変更で商品コードが変わったもの）は「品目入替」として別建て。単位：百万円',
            '出所：2508・2608タカハシ包装売上全明細（機械類を除く）。突合対象外には単価が前年比0.25〜4.0倍の範囲外の品目（前年0.6・当年0.2百万円）を含む')
labels=[f'{S0:.1f}',f'＋{pe:.1f}',f'▲{-qe:.1f}',f'▲{-out0:.1f}',f'＋{in1:.1f}',f'{S1:.1f}']
cats=[f'前年8月\n{labels[0]}',f'単価効果\n{labels[1]}',f'数量効果\n{labels[2]}',f'前年のみの品目\n（終売・切替前）\n{labels[3]}',f'当年のみの品目\n（新規・切替後）\n{labels[4]}',f'当年8月\n{labels[5]}']
base=[0, S0, S0+pe+qe, S0+pe+qe+out0, S0+pe+qe+out0, 0]
val =[S0, pe, -qe, -out0, in1, S1]
cd=CategoryChartData(); cd.categories=cats; cd.add_series('base',base); cd.add_series('value',val)
gf=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,Inches(0.45),Inches(1.25),Inches(7.6),Inches(5.6),cd); ch=gf.chart
ch.has_legend=False; ch.plots[0].gap_width=45; ch.plots[0].overlap=100
b=ch.series[0]; b.format.fill.background(); b.format.line.fill.background()
v=ch.series[1]; cols=[NAVY,GOLD,RED,RED,BLUE,NAVY]
for i,pt in enumerate(v.points):
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb=cols[i]; pt.format.line.fill.background()
    if i in (0,5):
        dl=pt.data_label; dl.has_text_frame=True; dl.text_frame.text=labels[i]; dl.position=XL_LABEL_POSITION.INSIDE_END
        for p in dl.text_frame.paragraphs:
            for r in p.runs: r.font.size=Pt(12); r.font.bold=True; r.font.color.rgb=WHITE; r.font.name='游ゴシック'
ca=ch.category_axis; ca.tick_labels.font.size=Pt(10); ca.tick_labels.font.bold=True; ca.tick_labels.font.name='游ゴシック'; ca.format.line.color.rgb=RGBColor(0xBF,0xBF,0xBF); ca.has_major_gridlines=False
va=ch.value_axis; va.tick_labels.font.size=Pt(8); va.has_major_gridlines=True; va.major_gridlines.format.line.color.rgb=RGBColor(0xE5,0xE5,0xE5); va.format.line.fill.background(); va.minimum_scale=0; va.maximum_scale=300
# 右側の解説
x=Inches(8.3); w=Inches(4.55)
card(s,x,Inches(1.25),w,Inches(1.75),'「突合」＝同一商品コードどうしの比較',
     [('単価寄与＋19.6％は、同じ商品が値上げされた効果。材質・色数などの仕様変更で商品コードが変わった品目（スペックダウンを含む）は突合できず、単価・数量には入らない。',{'size':10})],NAVY)
card(s,x,Inches(3.1),w,Inches(1.75),'品目入替は差し引き▲9.3百万円',
     [(f'前年のみの品目{-out0:.1f}百万円が消え、当年のみの品目{in1:.1f}百万円が入った。スペックダウン（単価引下げ）や安価品への切替の影響はここに含まれ、単価効果を打ち消す方向に働いている。',{'size':10})],RED)
card(s,x,Inches(4.95),w,Inches(1.9),'計画に使うのは「全体」の伸び率',
     [(f'突合ベースの＋17.0％をそのまま伸び率にすると過大。包装資材全体では＋{(S1/S0-1)*100:.1f}％で、これが計画の出発点。単価寄与は「値上げがどこまで進んだか」を測る指標として使う。',{'size':10})],GOLD)

# ============ P14: 単価上昇分の織り込み方（第6章）
tmpl=[i for i,sl in enumerate(S) if any(sh.has_text_frame and '第52期計画への道筋' in sh.text_frame.text for sh in sl.shapes)][0]
s=new_slide(tmpl,'第6章　第52期計画','単価上昇分の織り込み方 ―― 実績の上げ幅をそのまま持ち込むと不公平になる',
            '突合分析の単価寄与は「同一商品の値上げ効果」。第52期計画では、担当者・営業所ごとの転嫁の進み方を踏まえて単価前提を分けて置く',
            '出所：営業所別の単価寄与は202501～07　202601～07売上明細（4〜7月）と2508・2608タカハシ包装売上全明細（8月）。担当者コードで営業所に突合')
y=Inches(1.25); h=Inches(2.3); w=Inches(4.0); gap=Inches(0.2)
card(s,Inches(0.45),y,w,h,'① 担当金額が大きい人ほど負担が偏る',
     ['同じ「＋○％」でも、必要な増加額は担当金額に比例して大きくなる。',
      '8月は数量寄与▲2.5％と数量が弱く、金額の大きい担当ほど数量減のリスクを大きく抱える。',
      ('→ 増加額ではなく、単価要因と数量要因を分けて目標を置く。',{'size':10,'bold':True,'color':NAVY})],GOLD)
card(s,Inches(0.45)+w+gap,y,w,h,'② 先行して価格改定した人が損をする',
     ['第51期実績にすでに値上げ分が入っている担当に、さらに同じ上げ幅を乗せると二重計上になる。',
      '未改定の担当は値上げ余地が残っているため、同じ％でも達成が楽になる。',
      ('→ 転嫁の進み方（単価寄与の累積）を見て、未転嫁分だけを単価前提に織り込む。',{'size':10,'bold':True,'color':NAVY})],RED)
card(s,Inches(0.45)+2*(w+gap),y,w,h,'③ 対応の方向性',
     ['単価要因：営業所・担当者ごとの転嫁進捗を確認し、残りの転嫁分を計画に入れる。会社共通の値上げ枠として扱う。',
      '数量要因・粗利益額：担当者の努力として評価する軸にする。',
      ('→ 「値上げ」と「数量・粗利益」を分けて配分・評価する。',{'size':10,'bold':True,'color':NAVY})],BLUE)
# 営業所別 単価寄与の推移 表
tb=s.shapes.add_textbox(Inches(0.45),Inches(3.75),Inches(8),Inches(0.3)); run(tb.text_frame.paragraphs[0],'営業所別 単価寄与の推移（前年同月比、機械類を除く）―― 転嫁の進み方は営業所で違う',11,True,NAVY); tb.text_frame.margin_left=0
rows=[]
for o in OFFICES:
    p=price_by_office[o]; rows.append([o]+[(pct(x,1),{'color':GOLD if x>=0.1 else CHAR,'bold':x>=0.15}) for x in p])
table(s,Inches(0.45),Inches(4.1),Inches(8.2),[1.6,1,1,1,1,1],rows,['営業所','4月','5月','6月','7月','8月'],row_h=0.3,hdr_h=0.32)
textbox(s,Inches(8.9),Inches(4.1),Inches(3.95),Inches(2.5),[
    ('読み方',{'size':11,'bold':True,'color':NAVY,'space_before':0}),
    ('松江は5月＋7.8％→6月＋17.2％と早く転嫁が進み、8月は＋23.9％。',{'size':10,'bullet':True}),
    ('下関は6月＋7.6％と遅れて立ち上がり、8月＋16.5％で全営業所の下限。',{'size':10,'bullet':True}),
    ('先に上げた営業所に、さらに同じ上げ幅を求めると②の問題が起きる。転嫁が遅れた営業所には、残りの転嫁を第52期の単価前提として明示する。',{'size':10,'bullet':True}),
])

# ============ P15: 機械予算と包装資材予算（第6章）
s=new_slide(tmpl,'第6章　第52期計画','機械予算と包装資材予算 ―― メンテナンス要員のいる営業所は機械予算を上げる',
            '機械類はフロー（受注）で単価分析の対象外。売上数字を持たないメンテナンス要員の工数は、営業所の機械予算として明確に計上する。単位：千円',
            '出所：第51期着地及び第52期計画策定資料（第51期着地見込C・第52期計画D）。メンテナンス要員を置く営業所は部門長会議で確定する')
y=Inches(1.25); h=Inches(1.9); w=Inches(6.1); gap=Inches(0.2)
card(s,Inches(0.45),y,w,h,'機械予算：メンテナンス要員の工数を「数字」にする',
     ['売上数字を持たないメンテナンス担当がいる営業所は、その稼働を機械予算（保守・部品・更新提案）として明示的に引き上げる。',
      '機械はフローなので、単価分析の上げ幅は使わない。過年度の機械売上と、要員が担当する設置台数・保守契約から積み上げる。',
      ('→ 「機械の数字は誰のものか」を営業所単位で明確にする。',{'size':10,'bold':True,'color':NAVY})],NAVY)
card(s,Inches(0.45)+w+gap,y,w,h,'包装資材予算：機械との相乗効果を同じ営業所に織り込む',
     ['メンテナンスで現場に入る頻度が高い営業所ほど、包装資材（フィルム・トレー・消耗品）の受注機会が増える。',
      'したがって機械予算を上げる営業所は、包装資材の予算も同じ考え方で引き上げる。相乗効果を前提にした一体の予算にする。',
      ('→ 機械と包装資材を別々に積まず、営業所ごとに連動させる。',{'size':10,'bold':True,'color':NAVY})],GOLD)
tb=s.shapes.add_textbox(Inches(0.45),Inches(3.3),Inches(8),Inches(0.3)); run(tb.text_frame.paragraphs[0],'各営業所提出計画の機械・包装資材（第52期計画D）―― 機械予算の水準を確認する',11,True,NAVY); tb.text_frame.margin_left=0
plan=[('石見営業所','871,000','48,000','21,700','＋26,300'),('下関営業所','587,700','17,700','9,200','＋8,500'),('松江営業所','425,000','20,000','29,000','▲9,000'),
      ('広域営業部','557,000','0','2,900','▲2,900'),('境港営業所','662,400','92,000','147,400','▲55,400'),('水産部','525,600','1,500','10,000','▲8,500')]
rows=[[a,b,c,d,(e,{'color':RED if e.startswith('▲') else NAVY,'bold':True}),'']for a,b,c,d,e in plan]
table(s,Inches(0.45),Inches(3.65),Inches(12.4),[1.7,1.6,1.5,1.5,1.6,3.2],rows,['営業所','包装資材\n第52期計画','機械\n第52期計画','機械\n第51期見込','機械\n増減','メンテナンス要員の有無・機械予算の扱い（会議で記入）'],row_h=0.3,hdr_h=0.4)
textbox(s,Inches(0.45),Inches(6.3),Inches(12.4),Inches(0.5),[('※ 機械の第52期計画は6営業所計179,500千円で、第51期見込220,200千円から▲18.5％。境港は前年の大型案件の反動で減少計画だが、メンテナンス要員がいる営業所は保守・更新提案の分を上乗せして「明示的に上げる」対象とする。',{'size':9,'space_before':0})])

# ============ 並び替え：ブリッジ→P9の直後、他2枚は末尾
lst=prs.slides._sldIdLst; ids=list(lst); bridge=ids[-3]; lst.remove(bridge); lst.insert(9,bridge)
for i,sl in enumerate(prs.slides,1):
    for shp in sl.shapes:
        if shp.has_text_frame and '社外秘　P' in shp.text_frame.text: set_text(shp,f'㈱タカハシ包装センター　社外秘　P{i}')
# P9 注記の更新（突合の意味）
for sl in prs.slides:
    for shp in sl.shapes:
        if shp.has_text_frame and shp.text_frame.text.startswith('※ 包装資材計'):
            set_text(shp,'※ 「突合」は同一商品コードどうしの比較。仕様変更（材質・色数など）で商品コードが変わった品目は突合できず、次ページの「品目入替」に含まれる。包装資材計（機械類除く）の8月売上は＋7.5％（240.4→258.5百万円）、突合できた分は＋17.0％。全社売上は機械類（前年8月87.5百万円→8.1百万円）の反動で▲18.7％。')
prs.save('stage3.pptx'); print('stage3', len(prs.slides))
