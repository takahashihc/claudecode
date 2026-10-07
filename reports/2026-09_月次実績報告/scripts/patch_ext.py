# 外部リンク数式([1]Sheet!A1)をキャッシュ値で置換した「計算用コピー」を作る（成果物は変更しない）
import openpyxl, re, zipfile, sys
from xml.etree import ElementTree as ET
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); cache={}
for n in z.namelist():
    if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml"):
        root=ET.fromstring(z.read(n)); ns={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        names=[s.get("val") for s in root.iter("{%s}sheetName"%ns["m"])]
        for i,sd in enumerate(root.iter("{%s}sheetData"%ns["m"])):
            for cell in sd.iter("{%s}cell"%ns["m"]):
                v=cell.find("m:v",ns)
                if v is not None: cache[(names[i],cell.get("r"))]=float(v.text)
print("cached ext values:",len(cache))
wb=openpyxl.load_workbook(src); n=0
pat=re.compile(r"\[\d+\]([^!\[\]]+)!\$?([A-Z]+)\$?(\d+)")
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and "[1]" in c.value:
                def rep(m):
                    k=(m.group(1),m.group(2)+m.group(3)); return repr(cache.get(k,0))
                c.value=pat.sub(rep,c.value); n+=1
print("patched",n); wb.save(dst)
