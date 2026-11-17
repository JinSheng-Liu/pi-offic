import fitz,json
from PIL import Image,ImageDraw
from pathlib import Path
out=Path('data/validation/visual-audit');out.mkdir(parents=True,exist_ok=True)
doc=fitz.open('output/pdf/论文_Tab定位修正.pdf');stats=[]
for start in range(0,len(doc),12):
 sheet=Image.new('RGB',(1000,1125),'#cbd0d5');draw=ImageDraw.Draw(sheet)
 for k in range(start,min(start+12,len(doc))):
  page=doc[k];pix=page.get_pixmap(matrix=fitz.Matrix(.38,.38),alpha=False);im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples)
  x=(k-start)%4*250;y=(k-start)//4*375;sheet.paste(im,(x+10,y+22));draw.text((x+12,y+5),f'Page {k+1}',fill='black')
  blocks=page.get_text('blocks');text=page.get_text();stats.append({'page':k+1,'characters':len(text.strip()),'images':len(page.get_images()),'start':text[:100],'bottom':max([b[3] for b in blocks],default=0)})
 sheet.save(out/f'pages-{start+1}-{min(start+12,len(doc))}.png')
(out/'pages.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'pages':len(doc),'lowText':[x for x in stats if x['characters']<60]},ensure_ascii=True))
