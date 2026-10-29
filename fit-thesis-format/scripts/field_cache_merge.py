"""Import bounded layout-field caches into the original package.

Office SaveAs may normalize mathematical symbols and recompress images. Never
use that serialization as the final paper: retain original parts and content,
copy only approved result ranges and the TOC bookmarks those results require.
"""
from __future__ import annotations
import copy
from collections import defaultdict
import hashlib
from pathlib import Path
import re
import zipfile
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14='http://schemas.microsoft.com/office/word/2010/wordml'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS={'w':W}
SAFE={'TOC','SEQ','REF','PAGEREF','STYLEREF','PAGE','NUMPAGES','SECTIONPAGES'}
def q(s):return '{'+W+'}'+s
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def parse(b):return E.fromstring(b,E.XMLParser(resolve_entities=False,no_network=True))
def package(p):
    with zipfile.ZipFile(p) as z:
        if len(set(z.namelist()))!=len(z.namelist()):raise ValueError('Duplicate package entries')
        return {n:z.read(n) for n in z.namelist()}
def parent_p(n):return next((x for x in n.iterancestors() if x.tag==q('p')),None)
def ptext(p):return ''.join(p.xpath('.//w:t/text()',namespaces=NS))
def heading_text(p):
    result=[];depth=0
    for n in p.iter():
        if n.tag==q('fldChar'):
            kind=n.get(q('fldCharType'))
            if kind=='begin':depth+=1
            elif kind=='end':depth-=1
        elif n.tag==q('t') and depth==0 and not any(a.tag==q('fldSimple') for a in n.iterancestors()):result.append(n.text or '')
    return ''.join(result)
def spans(root):
    # Keep implementation shared with native probe; imported lazily to avoid
    # import cycles when office_probe invokes this merger after export.
    from office_probe import _field_spans
    return _field_spans(root)
def pspan(field):
    if 'simple' in field:return parent_p(field['simple']),parent_p(field['simple'])
    return parent_p(field['begin']),parent_p(field['end'])
def style_mapping(source,native):
    def read(parts):
        if 'word/styles.xml' not in parts:return {}
        return {s.get(q('styleId')):s.find(q('name')).get(q('val')) for s in parse(parts['word/styles.xml']).findall(q('style')) if s.find(q('name')) is not None}
    a,b=read(source),read(native);byname={v.casefold():k for k,v in a.items()}
    return {sid:byname[name.casefold()] for sid,name in b.items() if name.casefold() in byname},set(a)
def sanitize_cache(nodes,stylemap,existing_styles):
    for top in nodes:
        for n in top.iter():
            for attr in list(n.attrib):
                # Strip only identity/revision metadata, never required values
                # such as w14:ligatures/@w14:val (Word rejects their removal).
                if attr in ('{'+W14+'}paraId','{'+W14+'}textId') or attr.startswith('{'+W+'}rsid'):n.attrib.pop(attr)
                if attr in ('{'+R+'}id','{'+R+'}embed','{'+R+'}link'):
                    raise ValueError('TOC cache contains a relationship-bearing object')
            if n.tag in (q('pStyle'),q('rStyle')):
                old=n.get(q('val'))
                if old in stylemap:n.set(q('val'),stylemap[old])
                elif old not in existing_styles:raise ValueError('Unknown native cache style: '+str(old))


def toc_style_mapping(source,native,stylemap):
    """Prefer maintained TOC styles over the source's legacy built-in styles.

    Word regenerates TOC1/TOC2 regardless of the cache's custom paragraph
    styles. A name-only round trip can therefore select an old toc 1/toc 2
    definition that still inherits the paper's indented body formatting.
    """
    mapped=dict(stylemap)
    if 'word/styles.xml' not in source or 'word/styles.xml' not in native:return mapped
    source_styles={s.get(q('styleId')):s for s in parse(source['word/styles.xml']).findall(q('style'))}
    for st in parse(native['word/styles.xml']).findall(q('style')):
        if st.get(q('type'))!='paragraph':continue
        name=st.find(q('name'));name=name.get(q('val'),'') if name is not None else ''
        match=re.fullmatch(r'(?:toc|目录)\s*([12])',name,re.I)
        if match:
            desired='FIT_toc'+match.group(1)
            if desired in source_styles:mapped[st.get(q('styleId'))]=desired
    return mapped


def restore_toc_format(nodes,source):
    """Restore the owned TOC gallery style, retaining native field caches/tabs."""
    if 'word/styles.xml' not in source:return
    import thesis_format as T
    styles={s.get(q('styleId')):s for s in parse(source['word/styles.xml']).findall(q('style'))}
    for p in nodes:
        pp=p.find(q('pPr'));style=pp.find(q('pStyle')) if pp is not None else None
        sid=style.get(q('val')) if style is not None else None
        if sid not in ('FIT_toc1','FIT_toc2') or sid not in styles:continue
        # Native generation can materialize old Normal/TOC formatting as
        # paragraph/run overrides. Keep the measured leader position, but let
        # the maintained style control alignment, indent, fonts and spacing.
        T.format_p(p,sid[4:],styles)
        # Keep clickable native links without inheriting the source's visual
        # Hyperlink underline/colour after direct formatting is normalized.
        for run in p.iter(q('r')):
            rp=T.ensure(run,'rPr',True)
            T.remove_names(rp,{'u','color'})
            T.child(rp,'u',val='none');T.child(rp,'color',val='000000')


def merge_field_caches(source,native,output):
    """Create a NEW package; output never aliases source/native."""
    source,native,output=map(lambda p:Path(p).resolve(),(source,native,output))
    if output.exists() or output in (source,native):raise ValueError('Merge output must be a new file')
    before=sha(source);a,b=package(source),package(native);out=dict(a)
    stylemap,existing_styles=style_mapping(a,b)
    toc_styles=toc_style_mapping(a,b,stylemap)
    changes=[];bookmarks=[];retained_hidden=[]
    # Header/footer PAGE fields dynamically display per-page values; source
    # fields remain untouched. The substantive dynamic labels and TOC live in
    # document.xml. No native styles/media/math/section serialization is copied.
    sr=parse(a['word/document.xml']);nr=parse(b['word/document.xml'])
    original_fields=spans(sr);native_fields=spans(nr)
    source_tocs=[f for f in original_fields if f['kind']=='TOC']
    native_tocs=[f for f in native_fields if f['kind']=='TOC']
    if len(source_tocs)!=len(native_tocs):raise ValueError('TOC count changed during native refresh')
    excluded_source=set()
    for f in source_tocs:
        first,last=pspan(f)
        if first is None or first.getparent() is not last.getparent():raise ValueError('TOC crosses incompatible containers')
        parent=first.getparent();l,r=parent.index(first),parent.index(last)
        excluded_source.update(list(parent)[l:r+1])
    source_candidates=[p for p in sr.iter(q('p')) if p not in excluded_source and not any(x in excluded_source for x in p.iterancestors())]
    byid=defaultdict(list);bytext=defaultdict(list)
    for p in source_candidates:
        if p.get('{'+W14+'}paraId'):byid[p.get('{'+W14+'}paraId')].append(p)
        if heading_text(p):bytext[heading_text(p)].append(p)
    existing={x.get(q('name')):x for x in sr.iter(q('bookmarkStart'))}
    native_bm={x.get(q('name')):x for x in nr.iter(q('bookmarkStart'))}
    next_id=max([int(x.get(q('id'),'0')) for x in sr.iter(q('bookmarkStart')) if x.get(q('id'),'').isdigit()]+[0])+1
    for src_toc,native_toc in zip(source_tocs,native_tocs):
        if src_toc['identity']!=native_toc['identity']:raise ValueError('Native changed TOC field instruction')
        first,last=pspan(src_toc);nf,nl=pspan(native_toc)
        if nf is None or nf.getparent() is not nl.getparent():raise ValueError('Native TOC crosses incompatible containers')
        # Entire TOC paragraphs are replaceable only when no non-field content
        # precedes begin or follows end; do not delete a shared heading/paragraph.
        for f,p,edge in ((src_toc,first,'begin'),(src_toc,last,'end')):
            anchor=f[edge].getparent();index=p.index(anchor)
            outer=list(p)[:index] if edge=='begin' else list(p)[index+1:]
            if any(x.tag!=q('pPr') and (''.join(x.itertext()).strip() or x.tag not in (q('r'),q('bookmarkStart'),q('bookmarkEnd'))) for x in outer):
                raise ValueError('TOC shares a paragraph with protected non-cache content')
        np=nf.getparent();lo,hi=np.index(nf),np.index(nl)
        copied=[copy.deepcopy(x) for x in list(np)[lo:hi+1]]
        if any(x.tag!=q('p') for x in copied):raise ValueError('Native TOC contains unexpected blocks')
        targets=set()
        for p in copied:
            for instruction in p.iter(q('instrText')):
                match=re.match(r'^\s*PAGEREF\s+(?:"([^"]+)"|(\S+))',instruction.text or '',re.I)
                if match:targets.add(match.group(1) or match.group(2))
        for name in sorted(targets):
            if name not in native_bm:raise ValueError('Native TOC references missing bookmark '+name)
            native_p=parent_p(native_bm[name])
            if native_p is None:raise ValueError('Native TOC target is not a paragraph')
            found=[]
            if name in existing and heading_text(parent_p(existing[name]))==heading_text(native_p):found=[parent_p(existing[name])]
            if not found:found=byid.get(native_p.get('{'+W14+'}paraId'),[])
            if len(found)!=1 or heading_text(found[0])!=heading_text(native_p):found=bytext.get(heading_text(native_p),[])
            if len(found)>1:
                st=native_p.find('w:pPr/w:pStyle',NS)
                mapped=stylemap.get(st.get(q('val'))) if st is not None else None
                if mapped:found=[p for p in found if p.find('w:pPr/w:pStyle',NS) is not None and p.find('w:pPr/w:pStyle',NS).get(q('val'))==mapped]
            if len(found)!=1:raise ValueError('Ambiguous TOC heading target: '+ptext(native_p))
            target=found[0]
            if name in existing:
                if parent_p(existing[name]) is not target:raise ValueError('Existing bookmark points to a different heading: '+name)
                continue
            start=E.Element(q('bookmarkStart'));start.set(q('id'),str(next_id));start.set(q('name'),name)
            end=E.Element(q('bookmarkEnd'));end.set(q('id'),str(next_id));next_id+=1
            target.insert(1 if target.find(q('pPr')) is not None else 0,start);target.append(end)
            bookmarks.append({'name':name,'heading':ptext(target)})
            existing[name]=start
        sanitize_cache(copied,toc_styles,existing_styles)
        restore_toc_format(copied,a)
        parent=first.getparent();lo,hi=parent.index(first),parent.index(last)
        for p in list(parent)[lo:hi+1]:parent.remove(p)
        for index,p in enumerate(copied):parent.insert(lo+index,p)
        changes.append({'kind':'TOC','paragraphs':len(copied),'targets':len(targets)})
    # Merge visible plain-text layout caches outside the TOC. Missing hidden
    # reset SEQs are intentionally retained: Word may erase them on SaveAs.
    native_groups=defaultdict(list)
    for f in native_fields:
        if f['kind'] in SAFE and f['kind']!='TOC' and not f['inside_layout']:native_groups[f['identity']].append(f)
    for f in original_fields:
        if f['kind'] not in SAFE or f['kind']=='TOC' or f['inside_layout']:continue
        if f['kind']=='SEQ' and re.search(r'\\h\b',f['code']):
            retained_hidden.append(f['identity']);continue
        matching=native_groups.get(f['identity'],[])
        if not matching:raise ValueError('Native omitted visible layout field '+f['identity'])
        nf=matching.pop(0)
        if f['result']==nf['result']:continue
        if any(ch in nf['result'] for ch in '\r\n'):raise ValueError('Non-TOC layout field has multiline result')
        if 'simple' in f:
            texts=f['simple'].findall('.//w:t',NS)
        else:
            p1,p2=pspan(f)
            if p1 is not p2 or f['separate'] is None:raise ValueError('Layout field crosses paragraphs')
            separator=f['separate'].getparent();end=f['end'].getparent()
            between=list(p1)[p1.index(separator)+1:p1.index(end)]
            if any(n.tag not in {q(x) for x in ('r','rPr','rFonts','sz','szCs','t','b','i','color','lang','noProof','rStyle','vertAlign','spacing','kern','position','rtl')} for el in between for n in el.iter()):
                raise ValueError('Layout field cache contains non-text objects')
            texts=[t for el in between for t in el.iter(q('t'))]
        if texts:
            texts[0].text=nf['result']
            for t in texts[1:]:t.text=''
        else:
            r=E.Element(q('r'));t=E.SubElement(r,q('t'));t.text=nf['result']
            if 'simple' in f:f['simple'].append(r)
            else:p1.insert(p1.index(end),r)
        changes.append({'kind':f['kind'],'code':f['identity'],'before':f['result'],'after':nf['result']})
    out['word/document.xml']=E.tostring(sr,encoding='UTF-8',xml_declaration=True,standalone=True)
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in out.items():z.writestr(name,data)
    if sha(source)!=before:raise RuntimeError('Source changed during cache merge')
    return {'source_sha256':before,'native_docx_sha256':sha(native),'output_sha256':sha(output),
            'changed_parts':['word/document.xml'],'cache_changes':changes,'toc_bookmarks_added':bookmarks,
            'retained_hidden_fields':retained_hidden,'original_package_parts_retained':True}
