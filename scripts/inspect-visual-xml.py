from zipfile import ZipFile
from xml.etree import ElementTree as E
from pathlib import Path
import fitz,json
p=Path('data/validation/visual-audit');ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
with ZipFile('data/validation/tab-fix/论文_Tab定位修正.docx') as z:
 root=E.fromstring(z.read('word/document.xml'));body=root.find('w:body',ns)
 # same leaf paragraph traversal as SDK, excluding paragraph-contained textboxes
 def walk(n):
  if n.tag=='{'+ns['w']+'}p':yield n;return
  for c in n:yield from walk(c)
 ps=list(walk(body)); records=[]
 for i in [0,2,5,6,33,34,45,46,47,48,97,98,99,100,101]:
  n=ps[i];records.append({'index':i,'xml':E.tostring(n,encoding='unicode')[:10000]})
 (p/'front-xml.json').write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps([{'i':r['index'],'xml':r['xml'][:2200]} for r in records if r['index'] in [0,2,45,46,47,48]],ensure_ascii=True))
 print('anchors',len(root.findall('.//wp:anchor',ns)),'inline',len(root.findall('.//wp:inline',ns)))
doc=fitz.open('output/pdf/论文_Tab定位修正.pdf')
for i in [7,18,31,55]:doc[i].get_pixmap(matrix=fitz.Matrix(1.2,1.2),alpha=False).save(str(p/f'page-{i+1}.png'))
