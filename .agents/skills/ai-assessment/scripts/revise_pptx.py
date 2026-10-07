"""Revise an existing assessment PPTX using a reviewed page/shape plan, preserving package parts."""
import argparse,copy,json,posixpath
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
N={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def xml(x):return E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
def geo(s,v):
 x=s.find('p:spPr/a:xfrm',N)
 if x is None:x=s.find('p:xfrm',N)
 if x is None:
  x=E.SubElement(s.find('p:spPr',N),'{'+N['a']+'}xfrm');E.SubElement(x,'{'+N['a']+'}off');E.SubElement(x,'{'+N['a']+'}ext')
 for e,keys,vals in [(x.find('a:off',N),['x','y'],v[:2]),(x.find('a:ext',N),['cx','cy'],v[2:])]:
  for k,val in zip(keys,vals):e.set(k,str(round(val*914400)))
def txt(s,t):
 b=s.find('p:txBody',N)
 if b is None:b=s.find('a:txBody',N)
 oldp=b.find('a:p/a:pPr',N);oldr=b.find('.//a:rPr',N)
 for p in list(b.findall('a:p',N)):b.remove(p)
 for line in t.split('\n'):
  p=E.SubElement(b,'{'+N['a']+'}p')
  if oldp is not None:p.append(copy.deepcopy(oldp))
  r=E.SubElement(p,'{'+N['a']+'}r');r.append(copy.deepcopy(oldr) if oldr is not None else E.Element('{'+N['a']+'}rPr'));E.SubElement(r,'{'+N['a']+'}t').text=line

def main():
 a=argparse.ArgumentParser(description=__doc__);a.add_argument('source',type=Path);a.add_argument('plan',type=Path);a.add_argument('output',type=Path);args=a.parse_args();plan=json.loads(args.plan.read_text())
 if args.source.resolve()==args.output.resolve():raise ValueError('Use a separate candidate output')
 with ZipFile(args.source) as z:parts={n:z.read(n) for n in z.namelist()}
 pres=E.fromstring(parts['ppt/presentation.xml']);rr=E.fromstring(parts['ppt/_rels/presentation.xml.rels']);m={x.get('Id'):posixpath.normpath('ppt/'+x.get('Target')) for x in rr};pages=[m[x.get('{'+N['r']+'}id')] for x in pres.find('p:sldIdLst',N)];changed=[]
 for i,name in enumerate(pages,1):
  if i in plan.get('preserve_pages',[]):continue
  root=E.fromstring(parts[name]);spec=plan.get('slides',{}).get(str(i),{});floor=spec.get('minimum_font_pt',plan.get('minimum_font_pt',14));tree=root.find('p:cSld/p:spTree',N)
  for item in spec.get('clone_shapes',[]):
   original=next(s for s in tree if s.find('.//p:cNvPr',N) is not None and s.find('.//p:cNvPr',N).get('id')==str(item['source_id']))
   duplicate=copy.deepcopy(original);nv=duplicate.find('.//p:cNvPr',N);nv.set('id',str(item['new_id']));nv.set('name',item.get('name','Added text'))
   tree.append(duplicate)
  for s in tree:
   nv=s.find('.//p:cNvPr',N)
   if nv is None:continue
   sid=nv.get('id');patch=spec.get('shapes',{}).get(sid,{})
   ph=s.find('.//p:ph',N)
   if ph is not None and ph.get('type') in ['ftr','dt','sldNum']:continue
   if patch.get('remove'):tree.remove(s);continue
   if 'geometry' in patch:geo(s,patch['geometry'])
   if 'text' in patch:txt(s,patch['text'])
   if 'cells' in patch:
    for key,value in patch['cells'].items():
     ri,ci=map(int,key.split(','));cell=s.find('.//a:tbl',N).findall('a:tr',N)[ri].findall('a:tc',N)[ci];txt(cell,value)
   if 'row_heights' in patch:
    for tr,h in zip(s.find('.//a:tbl',N).findall('a:tr',N),patch['row_heights']):tr.set('h',str(round(h*914400)))
   if 'column_widths' in patch:
    for col,w in zip(s.find('.//a:tbl/a:tblGrid',N),patch['column_widths']):col.set('w',str(round(w*914400)))
   for rp in s.xpath('.//a:rPr|.//a:defRPr|.//a:endParaRPr',namespaces=N):
    rp.set('sz',str(round(max(floor,patch.get('font_pt',int(rp.get('sz',str(floor*100)))/100))*100)))
    for tag in ['latin','ea','cs']:
     f=rp.find('a:'+tag,N)
     if f is None:f=E.SubElement(rp,'{'+N['a']+'}'+tag)
     f.set('typeface','Meiryo UI')
   # A real minimum: never let the PowerPoint shrink-to-fit setting reduce it.
   for bp in s.findall('.//a:bodyPr',N):
    for x in list(bp):
     if E.QName(x).localname in ['normAutofit','spAutoFit','noAutofit']:bp.remove(x)
    E.SubElement(bp,'{'+N['a']+'}noAutofit')
    if patch.get('wrap'):bp.set('wrap','square')
    if 'vertical_anchor' in patch:bp.set('anchor',patch['vertical_anchor'])
   if 'alignment' in patch:
    for para in s.findall('.//a:p',N):
     pp=para.find('a:pPr',N)
     if pp is None:pp=E.Element('{'+N['a']+'}pPr');para.insert(0,pp)
     pp.set('algn',patch['alignment'])
   if 'line_pt' in patch:
    for p in s.findall('.//a:p',N):
     pp=p.find('a:pPr',N)
     if pp is None:pp=E.Element('{'+N['a']+'}pPr');p.insert(0,pp)
     for old in list(pp):
      if E.QName(old).localname in ['lnSpc','spcBef','spcAft']:pp.remove(old)
     x=E.SubElement(pp,'{'+N['a']+'}lnSpc');E.SubElement(x,'{'+N['a']+'}spcPts',val=str(round(patch['line_pt']*100)))
   if patch.get('behind_title'):
    tree.remove(s);tree.insert(2,s)
  data=xml(root)
  if data!=parts[name]:parts[name]=data;changed.append(i)
 if 'document_title' in plan:
  core=E.fromstring(parts['docProps/core.xml']);tag='{http://purl.org/dc/elements/1.1/}title';title=core.find(tag)
  if title is None:title=E.SubElement(core,tag)
  title.text=plan['document_title'];parts['docProps/core.xml']=xml(core)
 args.output.parent.mkdir(parents=True,exist_ok=True)
 with ZipFile(args.output,'w',ZIP_DEFLATED) as z:
  for n,b in parts.items():z.writestr(n,b)
 print('Changed pages:',changed);print(args.output)
if __name__=='__main__':main()
