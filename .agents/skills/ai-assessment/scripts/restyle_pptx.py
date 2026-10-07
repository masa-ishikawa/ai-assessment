"""Apply Oracle brand furniture while retaining a generator deck's body compositions.
No content generation, no source mutation. Requires lxml. Inspect native output after running.
"""
import argparse, copy, hashlib, json, posixpath, re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL='http://schemas.openxmlformats.org/package/2006/relationships'; CT='http://schemas.openxmlformats.org/package/2006/content-types'
def xml(e): return E.tostring(e,xml_declaration=True,encoding='UTF-8',standalone=True)
def read(p):
 with ZipFile(p) as z:return {n:z.read(n) for n in z.namelist()}
def text(s):return ''.join(s.xpath('.//a:t/text()',namespaces=N))
def relpath(p):return posixpath.dirname(p)+'/_rels/'+posixpath.basename(p)+'.rels'
def resolve(p,t):return posixpath.normpath(posixpath.join(posixpath.dirname(p),t)) if not t.startswith('/') else t.lstrip('/')
def ordered(parts):
 p=E.fromstring(parts['ppt/presentation.xml']);r=E.fromstring(parts['ppt/_rels/presentation.xml.rels']);m={x.get('Id'):resolve('ppt/presentation.xml',x.get('Target')) for x in r}
 return [m[x.get('{'+N['r']+'}id')] for x in p.find('p:sldIdLst',N)]
def geom(s):
 x=s.find('p:spPr/a:xfrm',N)
 if x is None:x=s.find('p:xfrm',N)
 if x is None:return None
 o=x.find('a:off',N);e=x.find('a:ext',N)
 if o is None or e is None:return None
 return [int(o.get('x'))/914400,int(o.get('y'))/914400,int(e.get('cx'))/914400,int(e.get('cy'))/914400]
def setgeom(s,vals):
 x=s.find('p:spPr/a:xfrm',N)
 if x is None:x=s.find('p:xfrm',N)
 if x is None:
  pr=s.find('p:spPr',N);x=E.SubElement(pr,'{'+N['a']+'}xfrm');E.SubElement(x,'{'+N['a']+'}off');E.SubElement(x,'{'+N['a']+'}ext')
 for e,attrs,values in [(x.find('a:off',N),['x','y'],vals[:2]),(x.find('a:ext',N),['cx','cy'],vals[2:])]:
  for k,v in zip(attrs,values):e.set(k,str(round(v*914400)))
def settext(s,t,size=None):
 b=s.find('p:txBody',N)
 if b is None:raise ValueError('shape is not editable text')
 op=b.find('a:p/a:pPr',N);rp=b.find('.//a:rPr',N)
 for p in list(b.findall('a:p',N)):b.remove(p)
 for line in t.split('\n'):
  p=E.SubElement(b,'{'+N['a']+'}p')
  if op is not None:p.append(copy.deepcopy(op))
  r=E.SubElement(p,'{'+N['a']+'}r');q=copy.deepcopy(rp) if rp is not None else E.Element('{'+N['a']+'}rPr')
  if size:q.set('sz',str(round(size*100)))
  for tag in ['latin','ea','cs']:
   f=q.find('a:'+tag,N)
   if f is None:f=E.SubElement(q,'{'+N['a']+'}'+tag)
   f.set('typeface','Meiryo UI')
  r.append(q);E.SubElement(r,'{'+N['a']+'}t').text=line

def placeholder_type(shape):
 ph=shape.find('.//p:ph',N)
 return ph.get('type','obj') if ph is not None else None

def body_furniture(root):
 tree=root.find('p:cSld/p:spTree',N)
 found={placeholder_type(shape) for shape in tree if placeholder_type(shape)}
 required={'title','sldNum','ftr'}
 missing=required-found
 if missing:raise ValueError('Oracle body base is missing inherited placeholders: '+', '.join(sorted(missing)))
 return found

def set_core_title(parts, value):
 if not value or 'docProps/core.xml' not in parts:return
 root=E.fromstring(parts['docProps/core.xml'])
 title=next((x for x in root if E.QName(x).localname=='title'),None)
 if title is None:
  title=E.SubElement(root,'{http://purl.org/dc/elements/1.1/}title')
 title.text=value
 parts['docProps/core.xml']=xml(root)

def main():
 a=argparse.ArgumentParser(description=__doc__);a.add_argument('source',type=Path);a.add_argument('template',type=Path);a.add_argument('output',type=Path);a.add_argument('--plan',type=Path,required=True);args=a.parse_args()
 if args.output.resolve() in [args.source.resolve(),args.template.resolve()]:raise ValueError('Output must be a separate candidate')
 plan=json.loads(args.plan.read_text());src=read(args.source);dst=read(args.template);ss=ordered(src);bs=ordered(dst)
 sp=E.fromstring(src['ppt/presentation.xml']);p=E.fromstring(dst['ppt/presentation.xml'])
 for attr in ['cx','cy']:
  if sp.find('p:sldSz',N).get(attr)!=p.find('p:sldSz',N).get(attr):raise ValueError('Source and template dimensions differ')
 ct=E.fromstring(dst['[Content_Types].xml']);sct=E.fromstring(src['[Content_Types].xml']);types={e.get('PartName').lstrip('/'):e.get('ContentType') for e in sct if e.get('PartName')};defaults={e.get('Extension'):e.get('ContentType') for e in sct if e.get('Extension')}
 def addtype(name,typ):
  if not any(x.get('PartName')=='/'+name for x in ct):E.SubElement(ct,'{'+CT+'}Override',PartName='/'+name,ContentType=typ)
 imported={}
 def importpart(name):
  if name in imported:return imported[name]
  target='ppt/assessmentSource/'+name;imported[name]=target;dst[target]=src[name]
  typ=types.get(name) or defaults.get(name.rsplit('.',1)[-1])
  if typ:addtype(target,typ)
  rp=relpath(name)
  if rp in src:
   rr=E.fromstring(src[rp])
   for r in rr:
    if r.get('TargetMode')!='External':r.set('Target',posixpath.relpath(importpart(resolve(name,r.get('Target'))),posixpath.dirname(target)))
   dst[relpath(target)]=xml(rr)
  return target
 pr=E.fromstring(dst['ppt/_rels/presentation.xml.rels']);sl=p.find('p:sldIdLst',N)
 for x in list(sl):sl.remove(x)
 for tag in ['custShowLst','extLst']:
  x=p.find('p:'+tag,N)
  if x is not None:p.remove(x)
 report=[]
 for i,sn in enumerate(ss,1):
  spec=plan['slides'][i-1];assert spec['source_slide']==i
  default_base_slide=1 if spec.get('kind')=='cover' else (3 if spec.get('kind')=='closing' else 2)
  base_slide=default_base_slide if len(bs)==3 else spec.get('base_slide',default_base_slide)
  bn=bs[base_slide-1];root=E.fromstring(dst[bn]);root.attrib.pop('show',None);tree=root.find('p:cSld/p:spTree',N);rr=E.fromstring(dst[relpath(bn)])
  for r in list(rr):
   if r.get('Type').endswith('/notesSlide'):rr.remove(r)
  source=E.fromstring(src[sn]);st=source.find('p:cSld/p:spTree',N);kind=spec.get('kind','body');kept=[];edits=[]
  if kind=='cover':
   for s in tree.findall('p:sp',N):
    ph=s.find('.//p:ph',N)
    if ph is None:continue
    k=ph.get('type','obj');idx=ph.get('idx')
    if k=='ctrTitle':settext(s,spec['title'],30)
    elif k=='body' and idx=='33':settext(s,spec['subtitle'],18)
    elif k=='body' and idx=='34':
     if spec.get('show_presenter',False):settext(s,spec.get('presenter',''),15)
     else:tree.remove(s)
  elif kind!='closing':
   furniture=body_furniture(root)
   for s in list(tree):
    local=E.QName(s).localname
    if local in ['nvGrpSpPr','grpSpPr']:continue
    ph=placeholder_type(s)
    if ph not in {'title','sldNum','ftr'}:tree.remove(s)
    elif ph=='title':settext(s,spec['title'],spec.get('title_size',24))
   sr=E.fromstring(src[relpath(sn)]) if relpath(sn) in src else E.Element('{'+REL+'}Relationships');rmap={r.get('Id'):r for r in sr}
   remap={};ids={};children=[]
   for j,s in enumerate(st):
    old=s.find('.//p:cNvPr',N)
    if old is None or E.QName(s).localname in ['nvGrpSpPr','grpSpPr']:continue
    oid=old.get('id');g=geom(s);t=text(s)
    # Remove inspected generator furniture. All body compositions remain in order.
    if oid in [str(v) for v in spec.get('remove_shapes',[])]:continue
    if g and g[1]>=7.0:continue
    if t.upper()=='OCI AI USE CASE ASSESSMENT':continue
    if g and g[1]<.1 and g[3]<.65:continue
    if g and .6<=g[1]<=.9 and t:continue
    s=copy.deepcopy(s);nv=s.find('.//p:cNvPr',N);newid=str(1000+j);ids[oid]=newid;nv.set('id',newid);nv.set('name','Source shape '+oid)
    for el in s.iter():
     for k,v in list(el.attrib.items()):
      if k.startswith('{'+N['r']+'}'):
       if v not in remap:
        r=copy.deepcopy(rmap[v]);rid='rIdSource'+str(len(remap)+1);remap[v]=rid;r.set('Id',rid)
        if r.get('TargetMode')!='External':r.set('Target',posixpath.relpath(importpart(resolve(sn,r.get('Target'))),'ppt/slides'))
        rr.append(r)
       el.set(k,remap[v])
    # Original renderer uses very wide drawString boxes. Bound them to the canvas
    # without moving their actual text. Never globally shrink body text.
    g=geom(s)
    if g and g[0]+g[2]>12.9 and t:
     g[2]=max(.1,12.65-g[0]);setgeom(s,g)
    is_source_note = t.strip().startswith(('参考：', '出典：', '出典:', '出典 '))
    for rp in s.xpath('.//a:rPr|.//a:defRPr|.//a:endParaRPr',namespaces=N):
     is_catalog = spec.get('role') == 'use_case_catalog' or any(len(t.findall('a:tr',N)) == 16 for t in source.findall('.//a:tbl',N))
     minimum = 900 if is_source_note else (1200 if is_catalog else 1400)
     if rp.get('sz') and int(rp.get('sz')) < minimum:rp.set('sz',str(minimum))
     for tag in ['latin','ea','cs']:
      f=rp.find('a:'+tag,N)
      if f is None:f=E.SubElement(rp,'{'+N['a']+'}'+tag)
      f.set('typeface','Meiryo UI')
    patch=spec.get('shape_edits',{}).get(oid,{})
    if 'text' in patch:settext(s,patch['text'],patch.get('size'))
    elif 'size' in patch:
     for rp in s.findall('.//a:rPr',N):rp.set('sz',str(round(patch['size']*100)))
    if 'geometry' in patch:setgeom(s,patch['geometry'])
    if 'line_spacing_pt' in patch:
     for pp in s.findall('.//a:pPr',N):
      oldsp=pp.find('a:lnSpc',N)
      if oldsp is not None:pp.remove(oldsp)
      x=E.SubElement(pp,'{'+N['a']+'}lnSpc');E.SubElement(x,'{'+N['a']+'}spcPts',val=str(round(patch['line_spacing_pt']*100)))
    for rp in s.xpath('.//a:rPr|.//a:defRPr|.//a:endParaRPr',namespaces=N):
     if rp.get('sz') and int(rp.get('sz')) < minimum:rp.set('sz',str(minimum))
    if patch:edits.append({'source_shape':oid,**patch})
    children.append(s);kept.append(oid)
   for s in children:
    for conn in s.xpath('.//a:stCxn|.//a:endCxn',namespaces=N):
     if conn.get('id') in ids:conn.set('id',ids[conn.get('id')])
    tree.append(s)
   output_furniture={placeholder_type(shape) for shape in tree if placeholder_type(shape)}
   missing={'title','sldNum','ftr'}-output_furniture
   if missing:raise ValueError(f'Slide {i} lost inherited Oracle placeholders: '+', '.join(sorted(missing)))
   for shape in tree.findall('p:sp',N):
    if placeholder_type(shape) is not None:continue
    g=geom(shape);value=text(shape).strip()
    if g and g[1]>=6.9 and (value.isdigit() or 'Copyright' in value):
     raise ValueError(f'Slide {i} contains a standalone footer repair: {value}')
  for f in root.findall('.//a:fld[@type="slidenum"]/a:t',N):f.text=str(i)
  outsn=f'ppt/slides/slide{200+i}.xml';dst[outsn]=xml(root)
  # Retain source evidence notes, with new reciprocal slide relationship.
  sr=E.fromstring(src[relpath(sn)]) if relpath(sn) in src else []
  notes=next((resolve(sn,r.get('Target')) for r in sr if r.get('Type').endswith('/notesSlide')),None)
  if notes:
   nn=importpart(notes);nr=E.fromstring(dst[relpath(nn)])
   for r in nr:
    if r.get('Type').endswith('/slide'):r.set('Target',posixpath.relpath(outsn,posixpath.dirname(nn)))
   dst[relpath(nn)]=xml(nr);E.SubElement(rr,'{'+REL+'}Relationship',Id='rIdAssessmentNotes',Type=N['r']+'/notesSlide',Target=posixpath.relpath(nn,'ppt/slides'))
  dst[relpath(outsn)]=xml(rr);rid='rIdAssessment'+str(i);E.SubElement(pr,'{'+REL+'}Relationship',Id=rid,Type=N['r']+'/slide',Target=posixpath.relpath(outsn,'ppt'))
  x=E.SubElement(sl,'{'+N['p']+'}sldId',id=str(2000+i));x.set('{'+N['r']+'}id',rid);addtype(outsn,'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
  report.append({'slide':i,'source_slide':sn,'base_slide':bn,'title':spec.get('title',''),'retained_source_shapes':kept,'edits':edits,'inherited_furniture':sorted(output_furniture) if kind not in {'cover','closing'} else []})
 assert len(plan['slides'])==len(ss)
 dst['ppt/presentation.xml']=xml(p);dst['ppt/_rels/presentation.xml.rels']=xml(pr);dst['[Content_Types].xml']=xml(ct)
 if 'docProps/app.xml' in dst:
  doc=E.fromstring(dst['docProps/app.xml'])
  for el in doc:
   if E.QName(el).localname in ['Slides','Notes']:el.text=str(len(ss))
   if E.QName(el).localname=='HiddenSlides':el.text='0'
  dst['docProps/app.xml']=xml(doc)
 # PDF exports inherit dc:title from the PPTX package. Replace the retained
 # Sales Session template title with the actual cover title and subtitle.
 cover=next((x for x in plan['slides'] if x.get('kind')=='cover'),{})
 document_title=str(plan.get('document_title') or '').strip()
 if not document_title:
  document_title='｜'.join(
   str(cover.get(key) or '').strip() for key in ('title','subtitle')
   if str(cover.get(key) or '').strip()
  )
 set_core_title(dst,document_title)
 args.output.parent.mkdir(parents=True,exist_ok=True)
 with ZipFile(args.output,'w',ZIP_DEFLATED) as z:
  for n,b in dst.items():z.writestr(n,b)
 result={'source':str(args.source.resolve()),'source_sha256':hashlib.sha256(args.source.read_bytes()).hexdigest(),'template':str(args.template.resolve()),'template_sha256':hashlib.sha256(args.template.read_bytes()).hexdigest(),'output':str(args.output.resolve()),'slides':report}
 args.output.with_suffix('.layout-map.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));print(args.output)
if __name__=='__main__':main()
