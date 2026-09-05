import zipfile, shutil, io, sys
from PIL import Image
src, dst = sys.argv[1], sys.argv[2]
zin=zipfile.ZipFile(src); zout=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data=zin.read(it.filename)
    if it.filename.startswith('ppt/media/') and it.filename.lower().endswith(('.jpg','.jpeg')):
        im=Image.open(io.BytesIO(data)).convert('RGB')
        if max(im.size)>600: im.thumbnail((600,600), Image.LANCZOS)
        buf=io.BytesIO(); im.save(buf,'JPEG',quality=88,optimize=True); data=buf.getvalue()
        print('resized',it.filename,im.size,len(data))
    if it.filename=='[Content_Types].xml': zout.writestr(it, data)  # keep first
    else: zout.writestr(it.filename, data)
zout.close(); print('ok')
