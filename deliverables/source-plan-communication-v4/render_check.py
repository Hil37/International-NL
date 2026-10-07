from pathlib import Path
import fitz, json, re, unicodedata
from PIL import Image, ImageDraw
R=Path(__file__).resolve().parent.parent
D=fitz.open(R/'output/Plan_Communication_France_FY2026-27_Leyton_V4.pdf')
manifest=json.loads((R/'output/content_manifest.json').read_text());geos=json.loads((R/'output/text_geometry.json').read_text())
norm=lambda s:re.sub(r'[^a-z0-9]','',unicodedata.normalize('NFKD',s.lower()))
missing=[];overflow=[];bounds=[]
for i,s in enumerate(manifest):
 t=norm(D[i].get_text())
 for block in s['text']:
  for para in block.split('\n'):
   if norm(para) and norm(para) not in t:missing.append((i+1,para))
for g in geos:
 page=D[g['slide']-1];target=norm(g['text'])
 blocks=[b for b in page.get_text('dict')['blocks'] if 'lines' in b]
 for b in blocks:
  bt=norm(' '.join(s['text'] for l in b['lines'] for s in l['spans']))
  if bt==target and target and abs(b['bbox'][0]-g['x'])<8 and abs(b['bbox'][1]-g['y'])<20:
   bb=b['bbox']
   if bb[3]>g['y']+g['h']+4 or bb[2]>g['x']+g['w']+3:overflow.append({'slide':g['slide'],'text':g['text'],'bottom':round(bb[3]-g['y']-g['h'],1),'right':round(bb[2]-g['x']-g['w'],1)})
   break
 for b in blocks:
  if b['bbox'][0]<-2 or b['bbox'][1]<-2 or b['bbox'][2]>1442 or b['bbox'][3]>812:bounds.append((g['slide'],b['bbox']))
prev=R/'output/preview';prev.mkdir(exist_ok=True);qa=R/'qa';qa.mkdir(exist_ok=True)
for i,page in enumerate(D,1):
 pix=page.get_pixmap(matrix=fitz.Matrix(1.25,1.25),alpha=False)
 pix.save(str(prev/f'slide-{i:02d}.png'))
for start in range(0,len(D),6):
 sheet=Image.new('RGB',(1600,1440),'#CAD2D8');dr=ImageDraw.Draw(sheet)
 for i in range(start,min(start+6,len(D))):
  sm=Image.open(prev/f'slide-{i+1:02d}.png');sm.thumbnail((790,444));k=i-start;x=k%2*800;y=k//2*480;sheet.paste(sm,(x,y+30));dr.text((x+16,y+9),f'SLIDE {i+1}',fill='black')
 sheet.save(qa/f'contact-{start//6+1}.png')
cols=4;w=2000;cellw=w//cols;cellh=310;rows=(len(D)+cols-1)//cols
mont=Image.new('RGB',(w,rows*cellh),'#CAD2D8');dr=ImageDraw.Draw(mont)
for i in range(len(D)):
 im=Image.open(prev/f'slide-{i+1:02d}.png');im.thumbnail((490,276));x=i%cols*cellw;y=i//cols*cellh;mont.paste(im,(x,y+25));dr.text((x+10,y+6),f'{i+1:02d} '+('PRINCIPAL' if i<13 else f'ANNEXE A{i-12:02d}'),fill='black')
mont.save(R/'output/Montage_global_Leyton_V4.png')
report={'slides':len(D),'main':13,'appendix':21,'missing_rendered_text':missing,'overflow':overflow,'outside_slide':list(set(bounds))}
(qa/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print(json.dumps(report,ensure_ascii=False,indent=2))

import subprocess, sys
subprocess.run([sys.executable,str(Path(__file__).resolve().parent/"validate_geometry.py")],check=True)
assert not missing and not overflow and not bounds, "Rendering validation failed"
