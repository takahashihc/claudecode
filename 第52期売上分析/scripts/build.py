# -*- coding: utf-8 -*-
import shutil, os, re, json, copy, zipfile, subprocess
from pathlib import Path
IN=Path('in'); SRC=IN/'kx'; V5=IN/'v5x'; WORK=Path('work')
if WORK.exists(): shutil.rmtree(WORK)
shutil.copytree(SRC, WORK)

# ---------- Step1: 構造コピー（v5 slide11,12 → slide10,11 / chart23,24 → chart2,3）
def rd(p): return Path(p).read_text(encoding='utf-8')
def wr(p,s): Path(p).write_text(s,encoding='utf-8')
pairs=[('slide11.xml','slide10.xml','chart23.xml','chart2.xml','Microsoft_Excel_Sheet19.xlsx'),
       ('slide12.xml','slide11.xml','chart24.xml','chart3.xml','Microsoft_Excel_Sheet20.xlsx')]
RELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpg"/><Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/%s"/></Relationships>'
CRELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/%s"/></Relationships>'
ct=rd(WORK/'[Content_Types].xml')
prels=rd(WORK/'ppt/_rels/presentation.xml.rels')
pres=rd(WORK/'ppt/presentation.xml')
new_sld=''
for i,(s_src,s_dst,c_src,c_dst,emb) in enumerate(pairs):
    shutil.copy(V5/'ppt/slides'/s_src, WORK/'ppt/slides'/s_dst)
    wr(WORK/'ppt/slides/_rels'/(s_dst+'.rels'), RELS%c_dst)
    shutil.copy(V5/'ppt/charts'/c_src, WORK/'ppt/charts'/c_dst)
    wr(WORK/'ppt/charts/_rels'/(c_dst+'.rels'), CRELS%emb)
    shutil.copy(V5/'ppt/embeddings'/emb, WORK/'ppt/embeddings'/emb)
    ct=ct.replace('</Types>', f'<Override PartName="/ppt/slides/{s_dst}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/charts/{c_dst}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>')
    rid=f'rId{20+i}'
    prels=prels.replace('</Relationships>', f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{s_dst}"/></Relationships>')
    new_sld+=f'<p:sldId id="{400+i}" r:id="{rid}"/>'
# slide6 (id=348) の直後に挿入
pres=pres.replace('<p:sldId id="348" r:id="rId7"/>','<p:sldId id="348" r:id="rId7"/>'+new_sld)
wr(WORK/'[Content_Types].xml',ct); wr(WORK/'ppt/_rels/presentation.xml.rels',prels); wr(WORK/'ppt/presentation.xml',pres)
os.system(f'cd {WORK} && rm -f ../stage1.pptx && zip -Xrq ../stage1.pptx .')
print('stage1 ok')
