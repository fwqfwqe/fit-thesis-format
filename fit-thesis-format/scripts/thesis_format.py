#!/usr/bin/env python3
"""Content-preserving DOCX paragraph/table formatting for FIT theses.

No text recognition, text replacement, numbering rewrite, field refresh, or
section/header creation. A reviewed source-hash-bound role map is required.
Only document.xml, styles.xml and explicitly mapped existing story parts are
serialized; all other ZIP members are copied unchanged.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import math
import re
from pathlib import Path
import sys
import tempfile
import zipfile
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'w': W, 'm': M, 'a': A}
SPEC_PATH = Path(__file__).resolve().parent.parent / 'references' / 'style-spec.json'
SPEC = json.loads(SPEC_PATH.read_text(encoding='utf-8'))
ROLES = SPEC['roles']
TEXT_TAGS = {f'{{{n}}}t' for n in (W, M, A)} | {f'{{{W}}}{x}' for x in ('delText', 'instrText', 'delInstrText')}
PROTECTED_TAGS = {f'{{{n}}}{x}' for n, xs in ((M, ('oMath', 'oMathPara')), (W, ('drawing', 'pict', 'object'))) for x in xs}
ANCHOR_TAGS={qname for qname in (f'{{{W}}}{x}' for x in ('fldChar','bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','commentReference','footnoteReference','endnoteReference','tab','br','cr','sym','noBreakHyphen','softHyphen'))}
HEADING_ROLES={'chapter','section','subsection','subsubsection','cn_title','cn_abstract_heading','en_title','en_abstract_heading','en_keywords','toc_heading','reference_heading','conclusion_heading','acknowledgement_heading','appendix_heading'}
FORMAT_CONTAINERS={f'{{{W}}}{x}' for x in ('pPr','rPr','tblPr','tblPrEx','tcPr','trPr','sectPr')}
RUN_CLEAR = {'rFonts','sz','szCs','color','highlight','shd','spacing','w','kern','fitText','rStyle'}
SEMANTIC = {'b','bCs','i','iCs','u','strike','dstrike','caps','smallCaps','vertAlign','vanish','webHidden','rtl','cs'}
PARA_CLEAR = {'pStyle','jc','ind','spacing','outlineLvl','keepNext','keepLines','pageBreakBefore','widowControl','contextualSpacing','adjustRightInd','snapToGrid','textAlignment','suppressAutoHyphens','pBdr','shd'}

def q(name): return '{' + W + '}' + name
def sha(data): return hashlib.sha256(data).hexdigest()
def read_json(path): return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def write_json(path, data):
    p = Path(path); p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
def parse(data): return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))
def xml(root): return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
def child(parent, name, **attrs):
    el = E.SubElement(parent, q(name))
    for k, v in attrs.items(): el.set(q(k), str(v))
    return el
def ensure(parent, name, first=False):
    el = parent.find(q(name))
    if el is None:
        el = E.Element(q(name)); parent.insert(0, el) if first else parent.append(el)
    return el
def remove_names(parent, names):
    for el in list(parent):
        if el.tag in {q(n) for n in names}: parent.remove(el)
def style_id(role): return 'FIT_' + role
def tree_signature(el):
    # Namespace prefixes and XML serialization are not content.
    return (el.tag, tuple(sorted(el.attrib.items())), el.text, el.tail,
            tuple(tree_signature(c) for c in el if isinstance(c.tag, str)))
def content_skeleton(el):
    # Preserve paragraphs/runs/order, breaks/tabs, links, field controls, tracked
    # content, and bookmarks. Only known formatting property containers are
    # excluded; text nodes keep exact spaces while XML pretty-print whitespace
    # between container elements is immaterial.
    if el.tag in FORMAT_CONTAINERS:return None
    value=el.text if el.tag in TEXT_TAGS or (el.text and el.text.strip()) else None
    return (el.tag,tuple(sorted(el.attrib.items())),value,
            tuple(content_skeleton(c) for c in el if isinstance(c.tag,str) and c.tag not in FORMAT_CONTAINERS))
def package(path):
    with zipfile.ZipFile(path) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError('Duplicate ZIP members are ambiguous and cannot be verified safely.')
        return {i.filename:z.read(i.filename) for i in z.infolist()}
def inventory(parts):
    text = {}; protected = {}; anchors={}; counts = {'math':0,'drawing':0,'object':0,'tracked_revisions':0}
    for name, data in parts.items():
        if not name.endswith('.xml') or name == 'word/styles.xml': continue
        root = parse(data)
        vals = [(el.tag, el.text or '') for el in root.iter() if el.tag in TEXT_TAGS]
        if vals: text[name] = vals
        nodes = []
        for el in root.iter():
            if el.tag in PROTECTED_TAGS and not any(a.tag in PROTECTED_TAGS for a in el.iterancestors()):
                nodes.append(tree_signature(el))
            if el.tag == f'{{{M}}}oMath': counts['math'] += 1
            if el.tag == q('drawing'): counts['drawing'] += 1
            if el.tag == q('object'): counts['object'] += 1
            if el.tag in {q('ins'),q('del'),q('moveFrom'),q('moveTo')}: counts['tracked_revisions'] += 1
        if nodes: protected[name] = nodes
        anchors[name]=[(el.tag,tuple(sorted(el.attrib.items())),el.text) for el in root.iter()
                       if el.tag in ANCHOR_TAGS and not any(a.tag in FORMAT_CONTAINERS for a in el.iterancestors())]
    binary = {n:sha(d) for n,d in parts.items() if n.startswith(('word/media/','word/embeddings/','word/charts/'))}
    return {'text':text,'protected':protected,'anchors':anchors,'binary':binary,'counts':counts,
            'structure':{name:content_skeleton(parse(data)) for name,data in parts.items()
                         if name=='word/document.xml' or allowed_story(name)},
            'body_text_sha256':sha(json.dumps(text.get('word/document.xml',[]), ensure_ascii=False).encode('utf-8'))}
def compare_content(before, after):
    errors=[]
    for key in ('text','protected','anchors','binary','structure'):
        if before[key] != after[key]: errors.append('Original content changed: ' + key)
    return errors
def indexed(root):
    return ({f'p{i:06d}':p for i,p in enumerate(root.iter(q('p')),1)},
            {f't{i:04d}':t for i,t in enumerate(root.iter(q('tbl')),1)})
def ptext(p): return ''.join(el.text or '' for el in p.iter() if el.tag in TEXT_TAGS)

def allowed_story(name):
    return re.fullmatch(r'word/(?:header[0-9]+|footer[0-9]+|footnotes|endnotes)\.xml',name) is not None

def selected_stories(parts,mapping):
    selected=mapping.get('story_paragraphs',{})
    if not isinstance(selected,dict):raise ValueError('story_paragraphs must be an object of existing story part -> paragraph map')
    roots={}
    for name,pmap in selected.items():
        if not allowed_story(name):raise ValueError('Story part is not an allowed header/footer/footnotes/endnotes XML part: '+name)
        if name not in parts:raise ValueError('Story part does not exist; creation is not allowed: '+name)
        root=parse(parts[name]);filename=Path(name).name
        expected='hdr' if filename.startswith('header') else 'ftr' if filename.startswith('footer') else filename[:-4]
        if root.tag!=q(expected):raise ValueError('Unexpected story root in '+name)
        if not isinstance(pmap,dict):raise ValueError('Story paragraph map must be an object: '+name)
        roots[name]=root
    return roots

def resolved_numpr(p,style_index,numbering=None,warnings=None,label=''):
    """Resolve inherited numbering before a paragraph style is changed.

    Word can keep numId in a paragraph style and associate ilvl through the
    abstract list's pStyle, so copying direct numPr alone is insufficient.
    """
    pst=p.find('w:pPr/w:pStyle',NS)
    sid=pst.get(q('val')) if pst is not None else next((s for s,e in style_index.items() if e.get(q('type'))=='paragraph' and e.get(q('default'))=='1'),'Normal')
    current=sid;chain=[];seen=set()
    while current and current not in seen and current in style_index:
        seen.add(current);st=style_index[current];chain.append(st)
        base=st.find(q('basedOn'));current=base.get(q('val')) if base is not None else None
    pieces=[];style_num_owner=None
    for st in reversed(chain):
        np=st.find('w:pPr/w:numPr',NS)
        if np is not None:
            pieces.append(np)
            if np.find(q('numId')) is not None:style_num_owner=st.get(q('styleId'))
    direct=p.find('w:pPr/w:numPr',NS)
    if direct is not None:pieces.append(direct)
    if not pieces:return None
    out=E.Element(q('numPr'));by_tag={}
    for np in pieces:
        for e in np:
            # ISO/IEC 29500: ilvl inside a paragraph style is ignored; the
            # numbering level is selected by the abstract definition's pStyle.
            if np is not direct and e.tag==q('ilvl'):continue
            by_tag[e.tag]=copy.deepcopy(e)
    numid=by_tag.get(q('numId'));ilvl=by_tag.get(q('ilvl'))
    has_direct_numbering=direct is not None and (direct.find(q('numId')) is not None or direct.find(q('ilvl')) is not None)
    if numid is not None and ilvl is None and numbering is not None and not has_direct_numbering:
        num=numbering.find(f'w:num[@w:numId="{numid.get(q("val"))}"]',NS)
        aid=num.find(q('abstractNumId')) if num is not None else None
        abstract=numbering.find(f'w:abstractNum[@w:abstractNumId="{aid.get(q("val"))}"]',NS) if aid is not None else None
        if abstract is not None:
            for lvl in abstract.findall(q('lvl')):
                linked=lvl.find(q('pStyle'))
                if linked is not None and linked.get(q('val'))==style_num_owner:
                    ilvl=E.Element(q('ilvl'));ilvl.set(q('val'),lvl.get(q('ilvl'),'0'));by_tag[q('ilvl')]=ilvl;break
    if numid is not None and ilvl is None and not has_direct_numbering and numid.get(q('val'))!='0':
        if warnings is not None:
            warnings.append(f'{label}: inherited numbering owned by style {style_num_owner} has no effective abstractNum/lvl/pStyle binding; left unnumbered instead of inventing level 0. Review the incomplete source numbering in Word.')
        return None
    if numid is not None and ilvl is None:
        ilvl=E.Element(q('ilvl'));ilvl.set(q('val'),'0');by_tag[q('ilvl')]=ilvl
    for name in ('ilvl','numId','numberingChange','ins'):
        if q(name) in by_tag:out.append(by_tag.pop(q(name)))
    for e in by_tag.values():out.append(e)
    return out

def numbering_signatures(parts,part='word/document.xml'):
    root=parse(parts[part]);styles=parse(parts['word/styles.xml'])
    idx={s.get(q('styleId')):s for s in styles.findall(q('style'))}
    numbering=parse(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else None
    result=[]
    for p in root.iter(q('p')):
        np=resolved_numpr(p,idx,numbering)
        if np is None:result.append(None);continue
        result.append(tuple((E.QName(e).localname,e.get(q('val'))) for e in np if E.QName(e).localname in ('numId','ilvl')))
    return result
def inspect_doc(path):
    parts=package(path); root=parse(parts['word/document.xml']); ps,ts=indexed(root)
    inv=inventory(parts); tid_by_node={t:i for i,t in ts.items()}
    paras=[]
    for pid,p in ps.items():
        anc=next((a for a in p.iterancestors() if a.tag==q('tbl')),None)
        st=p.find('w:pPr/w:pStyle',NS)
        paras.append({'id':pid,'text':ptext(p),'style':st.get(q('val')) if st is not None else None,
                      'table_id':tid_by_node.get(anc),'has_drawing':bool(p.findall('.//w:drawing',NS) or p.findall('.//w:pict',NS)),
                      'has_math':bool(p.findall('.//m:oMath',NS)),
                      'has_field':bool(p.findall('.//w:fldChar',NS) or p.findall('.//w:fldSimple',NS)),
                      'in_textbox':any(E.QName(a).localname=='txbxContent' for a in p.iterancestors())})
    tables=[]
    for tid,t in ts.items():
        tables.append({'id':tid,'rows':len(t.findall('w:tr',NS)),
                       'paragraph_ids':[pid for pid,p in ps.items() if next((a for a in p.iterancestors() if a.tag==q('tbl')),None) is t]})
    secs=[]
    for i,s in enumerate(root.iter(q('sectPr')),1):
        size=s.find(q('pgSz')); mar=s.find(q('pgMar'))
        secs.append({'id':f's{i:03d}','size':{E.QName(k).localname:v for k,v in size.attrib.items()} if size is not None else {},
                     'margins':{E.QName(k).localname:v for k,v in mar.attrib.items()} if mar is not None else {}})
    stories={}
    for name,data in parts.items():
        if not allowed_story(name):continue
        story=parse(data);sp,_=indexed(story)
        stories[name]={'paragraphs':[{'id':pid,'text':ptext(p),
            'style':p.find('w:pPr/w:pStyle',NS).get(q('val')) if p.find('w:pPr/w:pStyle',NS) is not None else None,
            'has_math':bool(p.findall('.//m:oMath',NS)),
            'has_drawing':bool(p.findall('.//w:drawing',NS) or p.findall('.//w:pict',NS)),
            'has_field':bool(p.findall('.//w:fldChar',NS) or p.findall('.//w:fldSimple',NS))} for pid,p in sp.items()]}
    return {'schema_version':1,'source':str(Path(path).resolve()),'source_sha256':sha(Path(path).read_bytes()),
            'body_text_sha256':inv['body_text_sha256'],'counts':inv['counts'],'paragraphs':paras,'tables':tables,'sections':secs,'stories':stories,
            'roles':{k:v['name'] for k,v in ROLES.items()},
            'notes':['IDs are stable only while the source document content/structure is unchanged.',
                     'A source_sha256-bound reviewed map is required; unmapped paragraphs are preserved.',
                     'Existing headers/footers/footnotes/endnotes may be explicitly mapped through story_paragraphs. No story or field creation; comments remain unchanged.']}

STYLE_OVERRIDE_KEYS = {'east_asia','western','size_pt','align','before_pt','after_pt',
    'line_pt','line_rule','indent_chars','left_chars','hanging_chars','bold',
    'page_break','keep_next','keep_lines','bottom_border_pt','tabs'}

def resolved_style_specs(mapping=None):
    """Resolve explicit, source-map-reviewed layout exceptions, never text changes.

    style_overrides: {role: {reviewed: true, reason: str, properties: {...}}}.
    Each exception affects that named style, not arbitrary paragraph/run XML.
    The same map must be supplied to apply and verify; otherwise school defaults
    remain the expected values. Names, outline levels and protections are fixed.
    """
    overrides = (mapping or {}).get('style_overrides', {})
    if not isinstance(overrides, dict):raise ValueError('style_overrides must be an object')
    resolved = {role: SPEC['defaults'] | properties for role,properties in ROLES.items()}
    for role, item in overrides.items():
        if role not in ROLES:raise ValueError('Unknown override style role: '+str(role))
        if (not isinstance(item,dict) or set(item) != {'reviewed','reason','properties'} or
                item.get('reviewed') is not True or not isinstance(item.get('reason'),str) or not item['reason'].strip()):
            raise ValueError('Style override requires reviewed=true, a specific reason and properties: '+role)
        properties=item['properties']
        if not isinstance(properties,dict) or not properties or set(properties)-STYLE_OVERRIDE_KEYS:
            raise ValueError('Unsupported or empty style override properties: '+role)
        if role=='equation' and set(properties)&{'east_asia','western','size_pt','bold'}:
            raise ValueError('Equation math font/size is protected; paragraph exceptions only')
        for key,value in properties.items():
            if key in {'east_asia','western'}:
                if not isinstance(value,str) or not value.strip():raise ValueError('Invalid font override: '+role+'.'+key)
            elif key=='align':
                if value not in ('left','center','right','both','distribute'):raise ValueError('Invalid alignment override: '+role)
            elif key=='line_rule':
                if value not in ('exact','atLeast','auto'):raise ValueError('Invalid line rule override: '+role)
            elif key in {'bold','page_break','keep_next','keep_lines'}:
                if type(value) is not bool:raise ValueError('Style flag must be boolean: '+role+'.'+key)
            elif key=='tabs':
                if not isinstance(value,list):raise ValueError('tabs must be a list: '+role)
                positions=[]
                for tab in value:
                    if (not isinstance(tab,dict) or set(tab)-{'align','position_twips','leader'} or
                            tab.get('align') not in ('left','center','right','decimal','bar','clear') or
                            type(tab.get('position_twips')) is not int or tab['position_twips']<0 or
                            tab.get('leader','none') not in ('none','dot','hyphen','underscore','heavy','middleDot')):
                        raise ValueError('Invalid tab stop override: '+role)
                    positions.append(tab['position_twips'])
                if len(set(positions))!=len(positions):raise ValueError('Duplicate tab positions: '+role)
            elif type(value) not in (int,float) or not math.isfinite(value) or value<0 or (key in {'size_pt','line_pt'} and value==0):
                raise ValueError('Style numeric value must be finite and nonnegative (font/line positive): '+role+'.'+key)
        resolved[role].update(properties)
        if resolved[role]['line_rule']=='exact' and resolved[role]['line_pt']<resolved[role]['size_pt']:
            raise ValueError('Exact line height cannot be smaller than the style font size: '+role)
    return resolved

def make_styles(styles,mapping=None):
    resolved=resolved_style_specs(mapping)
    for priority,(role,over) in enumerate(ROLES.items(),1):
        sid=style_id(role); old=styles.find(f'w:style[@w:styleId="{sid}"]',NS)
        if old is not None: styles.remove(old)
        sp=resolved[role]
        st=child(styles,'style',type='paragraph',styleId=sid,customStyle='1')
        child(st,'name',val=sp['name']); child(st,'basedOn',val='Normal')
        child(st,'next',val=style_id('body')); child(st,'uiPriority',val=priority); child(st,'qFormat')
        # New styles have no semiHidden/locked markers; all are in the gallery.
        child(st,'unhideWhenUsed',val='1')
        pp=child(st,'pPr')
        child(pp,'keepNext',val='1' if sp.get('keep_next') else '0')
        child(pp,'keepLines',val='1' if sp.get('keep_lines') else '0')
        child(pp,'pageBreakBefore',val='1' if sp.get('page_break') else '0')
        child(pp,'widowControl',val='1')
        if sp.get('bottom_border_pt'):
            pb=child(pp,'pBdr'); child(pb,'bottom',val='single',sz=round(sp['bottom_border_pt']*8),space=1,color='000000')
        if sp.get('tabs'):
            tabs=child(pp,'tabs')
            for tab in sorted(sp['tabs'],key=lambda item:item['position_twips']):
                child(tabs,'tab',val=tab['align'],pos=tab['position_twips'],leader=tab.get('leader','none'))
        child(pp,'snapToGrid',val='0')
        child(pp,'spacing',before=round(sp['before_pt']*20),after=round(sp['after_pt']*20),line=round(sp['line_pt']*20),lineRule=sp['line_rule'])
        ind=child(pp,'ind',left=0,right=0,firstLine=0,firstLineChars=round(sp['indent_chars']*100))
        if sp.get('left_chars'): ind.set(q('leftChars'),str(round(sp['left_chars']*100)))
        if sp.get('hanging_chars'):
            ind.attrib.pop(q('firstLine'),None); ind.attrib.pop(q('firstLineChars'),None)
            ind.set(q('hangingChars'),str(round(sp['hanging_chars']*100)))
        child(pp,'contextualSpacing',val='0'); child(pp,'jc',val=sp['align']); child(pp,'outlineLvl',val=sp['outline'])
        if role=='equation': continue
        rp=child(st,'rPr'); child(rp,'rFonts',ascii=sp['western'],hAnsi=sp['western'],eastAsia=sp['east_asia'],cs=sp['western'])
        child(rp,'b',val='1' if sp.get('bold') else '0'); child(rp,'bCs',val='1' if sp.get('bold') else '0')
        child(rp,'color',val='000000'); child(rp,'sz',val=round(sp['size_pt']*2)); child(rp,'szCs',val=round(sp['size_pt']*2))

def semantic_char_style(style_name, style_index, seen=None):
    seen=set() if seen is None else seen
    if style_name in seen or style_name not in style_index:return {}
    seen.add(style_name); st=style_index[style_name]; base=st.find(q('basedOn')); out={}
    if base is not None: out.update(semantic_char_style(base.get(q('val')),style_index,seen))
    rp=st.find(q('rPr'))
    if rp is not None:
        for e in rp:
            if E.QName(e).localname in SEMANTIC:out[E.QName(e).localname]=copy.deepcopy(e)
    return out

def format_p(p,role,style_index,numbering=None,warnings=None,label=''):
    effective_numpr=resolved_numpr(p,style_index,numbering,warnings,label)
    pp=ensure(p,'pPr',True)
    clear=PARA_CLEAR if role not in ('cover_title','cover_info') else {'pStyle','outlineLvl'}
    remove_names(pp,clear)
    st=E.Element(q('pStyle'));st.set(q('val'),style_id(role));pp.insert(0,st)
    if effective_numpr is not None:
        remove_names(pp,{'numPr'});pp.insert(1,effective_numpr)
    for r in p.iter(q('r')):
        if any(a.tag.startswith('{'+M+'}') for a in r.iterancestors()):continue
        # Textboxes have separately indexed paragraphs; do not format their runs twice.
        if next((a for a in r.iterancestors() if a.tag==q('p')),None) is not p:continue
        rp=r.find(q('rPr'))
        if rp is None:continue
        rstyle=rp.find(q('rStyle'))
        if rstyle is not None:
            semantic=semantic_char_style(rstyle.get(q('val')),style_index)
            present={E.QName(e).localname for e in rp}
            for name,el in semantic.items():
                if name not in present:rp.append(el)
        remove_names(rp,RUN_CLEAR)
        if role in HEADING_ROLES:
            remove_names(rp,{'b','bCs'})
    # Paragraph mark direct run formatting can override the gallery for newly typed text.
    mark=pp.find(q('rPr'))
    if mark is not None:
        remove_names(mark,RUN_CLEAR)
        if role in HEADING_ROLES:remove_names(mark,{'b','bCs'})

def border(parent,kind,value,size=0):child(parent,kind,val=value,sz=size,space=0,color='000000')
def format_table(t,setting):
    kind=setting['kind']; rows=t.findall('w:tr',NS); nh=setting.get('header_rows',1)
    if not isinstance(nh,int) or nh<0 or nh>len(rows):raise ValueError('Invalid header_rows')
    tp=ensure(t,'tblPr',True); remove_names(tp,{'tblBorders','jc','tblCellSpacing'})
    child(tp,'jc',val='center'); borders=child(tp,'tblBorders')
    for side in ('top','bottom','left','right','insideH','insideV'):
        if side in ('top','bottom'):border(borders,side,'single',12)
        elif side.startswith('inside') and kind=='complex':border(borders,side,'single',6)
        else:border(borders,side,'nil')
    for ri,row in enumerate(rows):
        rp=ensure(row,'trPr',True);remove_names(rp,{'trHeight','tblHeader'})
        if ri<nh:child(rp,'tblHeader')
        for ci,cell in enumerate(row.findall('w:tc',NS)):
            cp=ensure(cell,'tcPr',True);remove_names(cp,{'tcBorders'})
            cb=child(cp,'tcBorders')
            cells=row.findall('w:tc',NS)
            for side in ('top','bottom','left','right'):
                val='nil';sz=0
                if side=='top' and ri==0 or side=='bottom' and ri==len(rows)-1:val='single';sz=12
                elif side=='bottom' and nh and ri==nh-1:val='single';sz=6
                elif kind=='complex' and ((side=='bottom' and ri<len(rows)-1) or (side=='right' and ci<len(cells)-1)):
                    val='single';sz=6
                border(cb,side,val,sz)
    # Remove conditional border overrides while retaining all text/layout elsewhere.
    for ex in t.findall('w:tr/w:tblPrEx',NS):remove_names(ex,{'tblBorders'})

def apply_pages(root,mapping,warnings):
    if not mapping.get('page_setup',True):return
    pg=SPEC['page']
    for i,s in enumerate(root.iter(q('sectPr')),1):
        size=ensure(s,'pgSz'); is_land=size.get(q('orient'))=='landscape' or int(size.get(q('w'),'0'))>int(size.get(q('h'),'99999'))
        if is_land:
            policy=mapping.get('landscape_policy','preserve')
            if policy!='portrait':
                warnings.append(f's{i:03d}: landscape section preserved; requires visual review.');continue
        size.set(q('w'),str(pg['width_twips']));size.set(q('h'),str(pg['height_twips']));size.set(q('orient'),'portrait')
        margin=ensure(s,'pgMar')
        for x in ('top','bottom','left','right','header','footer'):margin.set(q(x),str(pg[x+'_twips']))
        margin.set(q('gutter'),'0')

def validate_paragraph_roles(pmap,ps,context='document'):
    if not isinstance(pmap,dict):raise ValueError('Paragraph map must be an object of paragraph ID -> role or {role, expected_text}: '+context)
    for pid,item in pmap.items():
        role=item if isinstance(item,str) else item['role']
        if pid not in ps:raise ValueError('Unknown paragraph ID '+context+':'+pid)
        if role not in ROLES:raise ValueError('Unknown role '+str(role))
        if any(a.tag in PROTECTED_TAGS for a in ps[pid].iterancestors()):
            raise ValueError('Paragraph '+pid+' is inside a protected drawing/object; leave it unmapped for manual visual review.')
        if isinstance(item,dict) and 'expected_text' in item and item['expected_text']!=ptext(ps[pid]):raise ValueError('Text mismatch '+pid)

def validate_map(mapping,source,ps,ts,parts):
    if mapping.get('source_sha256')!=sha(Path(source).read_bytes()):raise ValueError('Map source_sha256 missing or stale. Inspect the exact source and review the map again.')
    resolved_style_specs(mapping)
    validate_paragraph_roles(mapping.get('paragraphs'),ps)
    for tid,opts in mapping.get('tables',{}).items():
        if tid not in ts:raise ValueError('Unknown table ID '+tid)
        if opts.get('kind') not in ('three-line','complex'):raise ValueError('Table kind must be three-line or complex')
    for name,root in selected_stories(parts,mapping).items():
        validate_paragraph_roles(mapping['story_paragraphs'][name],indexed(root)[0],name)
    if not (mapping.get('paragraphs') or mapping.get('tables') or any(mapping.get('story_paragraphs',{}).values())):
        raise ValueError('Empty formatting map: select at least one reviewed paragraph, table or story paragraph. No DOCX was changed.')

def format_property_signature(element):
    """Compare semantic property values independently of XML property order.

    This is deliberately limited to the package's generated named styles.
    Local run emphasis, superscripts, existing cover paragraph layout and OMML
    formatting are handled by the existing preservation checks, not flattened.
    """
    if element is None:return None
    boolean_names={'b','bCs','keepNext','keepLines','pageBreakBefore','widowControl','snapToGrid','contextualSpacing'}
    attrs=dict(element.attrib)
    if E.QName(element).localname in boolean_names:
        attrs[q('val')]='0' if attrs.get(q('val'),'1') in ('0','false','off') else '1'
    return (element.tag,tuple(sorted(attrs.items())),
            tuple(sorted((format_property_signature(c) for c in element if isinstance(c.tag,str)),key=repr)))

def verify_style_values(styles,mapping=None):
    expected=E.Element(q('styles'));make_styles(expected,mapping)
    errors=[]
    for wanted in expected.findall(q('style')):
        sid=wanted.get(q('styleId'));actual=styles.find(f'w:style[@w:styleId="{sid}"]',NS)
        if actual is None:continue # Existing gallery checks report the missing style.
        if actual.get(q('type'))!='paragraph':errors.append('Wrong style type: '+sid)
        for container in ('name','basedOn','pPr','rPr'):
            if format_property_signature(actual.find(q(container)))!=format_property_signature(wanted.find(q(container))):
                errors.append('Style format differs from school specification/reviewed override: '+sid+'/'+container)
        flag=actual.find(q('autoRedefine'))
        if flag is not None and flag.get(q('val'),'1') not in ('0','false','off'):
            errors.append('Style automatically redefines from local formatting: '+sid)
    return errors

def verify_paragraph_roles(ps,pmap,context='document'):
    errors=[]
    for pid,item in pmap.items():
        role=item if isinstance(item,str) else item['role'];p=ps.get(pid);label=context+':'+pid
        if p is None:errors.append('Paragraph absent: '+label);continue
        pp=p.find(q('pPr'));st=pp.find(q('pStyle')) if pp is not None else None
        if st is None or st.get(q('val'))!=style_id(role):errors.append('Wrong style: '+label)
        if pp is not None and role not in ('cover_title','cover_info'):
            blockers=[E.QName(e).localname for e in pp if E.QName(e).localname in PARA_CLEAR-{'pStyle'}]
            if blockers:errors.append('Paragraph style blocked '+label+': '+','.join(blockers))
        for r in p.iter(q('r')):
            if any(a.tag.startswith('{'+M+'}') for a in r.iterancestors()):continue
            if next((a for a in r.iterancestors() if a.tag==q('p')),None) is not p:continue
            rp=r.find(q('rPr'))
            if rp is not None and any(E.QName(e).localname in RUN_CLEAR for e in rp):errors.append('Run style blocked: '+label)
    return errors

def verify_doc(path,source=None,mapping=None):
    parts=package(path);root=parse(parts['word/document.xml']);styles=parse(parts['word/styles.xml']);ps,ts=indexed(root)
    errors=[];warnings=[]; inv=inventory(parts)
    story_roots=selected_stories(parts,mapping) if mapping else {}
    errors.extend(verify_style_values(styles,mapping))
    for role in ROLES:
        st=styles.find(f'w:style[@w:styleId="{style_id(role)}"]',NS)
        if st is None:errors.append('Missing style '+role);continue
        if st.find(q('qFormat')) is None:errors.append('Style not in gallery: '+role)
        if st.find(q('uiPriority')) is None:errors.append('Style gallery priority missing: '+role)
        outline=st.find('w:pPr/w:outlineLvl',NS)
        expected=str(ROLES[role].get('outline',9))
        if outline is None or outline.get(q('val'))!=expected:errors.append('Wrong outline level: '+role)
        for flag in ('semiHidden','locked'):
            e=st.find(q(flag))
            if e is not None and e.get(q('val'),'1') not in ('0','false','off'):errors.append('Style '+flag+': '+role)
    if source:
        src=package(source);errors.extend(compare_content(inventory(src),inv))
        for name in src:
            if name in parts and (name=='word/document.xml' or allowed_story(name)):
                if numbering_signatures(src,name)!=numbering_signatures(parts,name):errors.append('Effective automatic numbering changed: '+name)
        if mapping and mapping.get('source_sha256')!=sha(Path(source).read_bytes()):errors.append('Verification map source hash does not match source')
        if set(src)!=set(parts):errors.append('ZIP member set changed')
        for name,data in src.items():
            if name not in ('word/document.xml','word/styles.xml') and name not in story_roots and parts.get(name)!=data:errors.append('Unrelated ZIP member changed: '+name)
    if mapping:
        errors.extend(verify_paragraph_roles(ps,mapping['paragraphs']))
        for name,sroot in story_roots.items():
            errors.extend(verify_paragraph_roles(indexed(sroot)[0],mapping['story_paragraphs'][name],name))
        for tid,opts in mapping.get('tables',{}).items():
            t=ts.get(tid)
            if t is None:errors.append('Table absent: '+tid);continue
            b=t.find('w:tblPr/w:tblBorders',NS)
            if b is None:errors.append('Table borders absent: '+tid);continue
            for side in ('top','bottom'):
                e=b.find(q(side))
                if e is None or e.get(q('sz'))!='12' or e.get(q('val'))!='single':errors.append('Wrong outer table border: '+tid)
            for side in ('left','right','insideH','insideV'):
                expected_val='single' if opts['kind']=='complex' and side.startswith('inside') else 'nil'
                e=b.find(q(side))
                if e is None or e.get(q('val'))!=expected_val:errors.append('Wrong '+side+' table border: '+tid)
                elif expected_val=='single' and e.get(q('sz'))!='6':errors.append('Wrong inside table line width: '+tid)
            for i,row in enumerate(t.findall('w:tr',NS)):
                if i<opts.get('header_rows',1) and row.find('w:trPr/w:tblHeader',NS) is None:errors.append('Missing repeated header: '+tid)
                if row.find('w:trPr/w:trHeight',NS) is not None:errors.append('Fixed row height: '+tid)
                if i==opts.get('header_rows',1)-1 and i<len(t.findall('w:tr',NS))-1:
                    for cell in row.findall('w:tc',NS):
                        e=cell.find('w:tcPr/w:tcBorders/w:bottom',NS)
                        if e is None or e.get(q('val'))!='single' or e.get(q('sz'))!='6':errors.append('Wrong header separator: '+tid)
    if inv['counts']['tracked_revisions']:warnings.append('Tracked revisions preserved; inspect both original/final views in Word.')
    return {'ok':not errors,'errors':errors,'warnings':warnings,'counts':inv['counts'],'body_text_sha256':inv['body_text_sha256'],
            'style_count':len(ROLES),'content_preserved':not any('content changed' in e.lower() or 'numbering changed' in e.lower() or 'zip member' in e.lower() for e in errors) if source else None}

def apply_doc(source,destination,mapping,report):
    src=Path(source).resolve();dest=Path(destination).resolve()
    if src==dest:raise ValueError('Refusing to overwrite the input document.')
    if dest.exists():raise ValueError('Output already exists; choose a new path. No files were overwritten.')
    parts=package(src);root=parse(parts['word/document.xml']);styles=parse(parts['word/styles.xml']);ps,ts=indexed(root)
    validate_map(mapping,src,ps,ts,parts);style_index={s.get(q('styleId')):copy.deepcopy(s) for s in styles.findall(q('style'))}
    numbering=parse(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else None
    make_styles(styles,mapping);warnings=[]
    for pid,item in mapping['paragraphs'].items():format_p(ps[pid],item if isinstance(item,str) else item['role'],style_index,numbering,warnings,pid)
    for tid,opts in mapping.get('tables',{}).items():format_table(ts[tid],opts)
    story_roots=selected_stories(parts,mapping)
    for name,sroot in story_roots.items():
        sp,_=indexed(sroot)
        for pid,item in mapping['story_paragraphs'][name].items():
            format_p(sp[pid],item if isinstance(item,str) else item['role'],style_index,numbering,warnings,name+':'+pid)
    apply_pages(root,mapping,warnings)
    dest.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix='.fit-format-',suffix='.docx',dir=dest.parent,delete=False) as handle:temp=Path(handle.name)
    try:
        with zipfile.ZipFile(src) as zin,zipfile.ZipFile(temp,'w') as zout:
            for info in zin.infolist():
                data=xml(root) if info.filename=='word/document.xml' else xml(styles) if info.filename=='word/styles.xml' else xml(story_roots[info.filename]) if info.filename in story_roots else zin.read(info.filename)
                zout.writestr(info,data)
        result=verify_doc(temp,src,mapping);result['warnings'].extend(warnings)
        result.update({'source':str(src),'output':str(dest),'source_sha256':sha(src.read_bytes()),
                       'selected_paragraphs':len(mapping['paragraphs']),
                       'unmapped_paragraphs':[pid for pid in ps if pid not in mapping['paragraphs']],
                       'selected_tables':list(mapping.get('tables',{})),
                       'selected_story_paragraphs':{name:list(pmap) for name,pmap in mapping.get('story_paragraphs',{}).items()},
                       'limitations':['No field/TOC update or numbering rewrite.','No header/footer or section-boundary creation.','Visual pagination must be reviewed in Word.']})
        if report:write_json(report,result)
        if not result['ok']:raise ValueError('Verification failed: '+ '; '.join(result['errors']))
        temp.rename(dest)
        return result
    finally:
        if temp.exists():temp.unlink()

def main():
    if hasattr(sys.stdout,'reconfigure'):sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr,'reconfigure'):sys.stderr.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser(description=__doc__);sub=ap.add_subparsers(dest='command',required=True)
    ip=sub.add_parser('inspect');ip.add_argument('input');ip.add_argument('--out',required=True)
    aa=sub.add_parser('apply');aa.add_argument('input');aa.add_argument('output');aa.add_argument('--map',required=True);aa.add_argument('--report',required=True)
    vp=sub.add_parser('verify');vp.add_argument('output');vp.add_argument('--source');vp.add_argument('--map');vp.add_argument('--report')
    args=ap.parse_args()
    try:
        if args.command=='inspect':
            result=inspect_doc(args.input);write_json(args.out,result)
            print(json.dumps({'ok':True,'manifest':args.out,'paragraphs':len(result['paragraphs']),'tables':len(result['tables']),'roles':result['roles']},ensure_ascii=False));return 0
        if args.command=='apply':result=apply_doc(args.input,args.output,read_json(args.map),args.report)
        else:
            result=verify_doc(args.output,args.source,read_json(args.map) if args.map else None)
            if args.report:write_json(args.report,result)
        display=dict(result)
        if 'unmapped_paragraphs' in display:
            display['unmapped_paragraph_count']=len(display.pop('unmapped_paragraphs'))
            display['full_report']=args.report
        print(json.dumps(display,ensure_ascii=False,indent=2));return 0 if result['ok'] else 2
    except (ValueError,KeyError,zipfile.BadZipFile,E.XMLSyntaxError,OSError) as exc:
        print(json.dumps({'ok':False,'error':str(exc)},ensure_ascii=False),file=sys.stderr);return 2
if __name__=='__main__':sys.exit(main())
