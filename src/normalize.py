from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import xml.etree.ElementTree as E
b=Path(__file__).parent
ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
for prefix,uri in [('c',ns['c']),('a','http://schemas.openxmlformats.org/drawingml/2006/main'),('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')]:E.register_namespace(prefix,uri)
with ZipFile(b/'candidate.pptx') as z,ZipFile(b/'candidate_normalized.pptx','w',ZIP_DEFLATED) as out:
    names=set(z.namelist())
    for item in z.infolist():
        data=z.read(item.filename)
        if item.filename=='[Content_Types].xml':
            r=E.fromstring(data)
            for e in list(r):
                part=e.get('PartName','').lstrip('/')
                if part.startswith('ppt/slideMasters/slideMaster') and part not in names:r.remove(e)
            E.register_namespace('','http://schemas.openxmlformats.org/package/2006/content-types')
            data=E.tostring(r,encoding='utf-8',xml_declaration=True)
        if item.filename.startswith('ppt/charts/chart') and item.filename.endswith('.xml'):
            r=E.fromstring(data);c='{'+ns['c']+'}'
            for cat in r.findall('.//c:ser/c:cat',ns):
                ref=cat.find('c:multiLvlStrRef',ns)
                if ref is not None:
                    cache=ref.find('c:multiLvlStrCache',ns);levels=cache.findall('c:lvl',ns)
                    assert len(levels)==1
                    ref.tag=c+'strRef';cache.tag=c+'strCache';cache.remove(levels[0]);cache.extend(list(levels[0]))
            data=E.tostring(r,encoding='utf-8',xml_declaration=True)
        out.writestr(item.filename,data)
print('Package normalized; chart caches retain exact source data.')
