import fitz,json,re,unicodedata,math
from pathlib import Path
R=Path(__file__).resolve().parent.parent;D=fitz.open(R/'output/Plan_Communication_France_FY2026-27_Leyton_V4.pdf');geos=json.loads((R/'output/text_geometry.json').read_text())
norm=lambda s:re.sub(r'[^a-z0-9]','',unicodedata.normalize('NFKD',s.lower()))
arrays=[]
for page in D:
 ss=[];pos=0
 for b in page.get_text('dict')['blocks']:
  for line in b.get('lines',[]):
   for span in line['spans']:
    t=norm(span['text']);ss.append((pos,pos+len(t),span['bbox']));pos+=len(t)
 arrays.append((''.join(norm(span['text']) for b in page.get_text('dict')['blocks'] for l in b.get('lines',[]) for span in l['spans']),ss))
o=[];missing=[]
for g in geos:
 target=norm(g['text'])
 if not target:continue
 full,spans=arrays[g['slide']-1];idx=0;options=[]
 while True:
  i=full.find(target,idx)
  if i<0:break
  matches=[bbox for a,b,bbox in spans if b>i and a<i+len(target)]
  if matches:
   bb=[min(b[0] for b in matches),min(b[1] for b in matches),max(b[2] for b in matches),max(b[3] for b in matches)]
   options.append(bb)
  idx=i+len(target)
 if not options:
  missing.append((g['slide'],g['text']));continue
 bb=min(options,key=lambda b:abs(b[0]-g['x'])+abs(b[1]-g['y']))
 if abs(bb[0]-g['x'])>15 or abs(bb[1]-g['y'])>25:continue
 if bb[3]>g['y']+g['h']+4 or bb[2]>g['x']+g['w']+3:o.append({'slide':g['slide'],'text':g['text'],'deltaBottom':round(bb[3]-g['y']-g['h'],1),'deltaRight':round(bb[2]-g['x']-g['w'],1),'actual_height':round((bb[3]-g['y'])/72+.05,3),'old_height':g['h']/72,'x':g['x']/72,'y':g['y']/72})
(R/'qa').mkdir(exist_ok=True)
print(json.dumps({'overflow':o,'unmatched':missing},ensure_ascii=False,indent=2));(R/'qa/span_report.json').write_text(json.dumps({'overflow':o,'unmatched':missing},ensure_ascii=False,indent=2))

assert not o, "Rendered text extends beyond its PowerPoint textbox"
assert not missing, "Textboxes could not be matched to the rendered PDF"
