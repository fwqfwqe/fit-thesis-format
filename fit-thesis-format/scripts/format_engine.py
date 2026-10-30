#!/usr/bin/env python3
"""Deterministic existing-DOCX transformer. No model/network/vendor dependency.

The model confirms ambiguous semantics through source-bound decisions. This
module changes only supported structures and reports every unresolved object.
Native Office rendering and actual page review are separate completion gates.
"""
from __future__ import annotations
import copy
import hashlib
import json
import math
import re
import zipfile
from pathlib import Path
from lxml import etree as E
import thesis_format as F
import thesis_layout as L
from format_audit import Audit
import citation_style as C
import narrative_citations as N
import math_transcription as MT

NS=F.NS; W=F.W; M=F.M; q=F.q
R=L.R; PKG=L.PKG; CT=L.CT
DOC='word/document.xml'
WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC='http://schemas.openxmlformats.org/drawingml/2006/picture'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
LAYOUT_BOOL={'page_break_before':'pageBreakBefore','keep_with_next':'keepNext','keep_lines':'keepLines'}
LAYOUT_NUM={'before_pt':'before','after_pt':'after','line_min_pt':'line'}
LEVELS=('chapter','section','subsection','subsubsection')
PREFIX=re.compile(r'^(?:第[一二三四五六七八九十百\d]+章\s*|[一二三四五六七八九十百]+[、．.]\s*|[（(][一二三四五六七八九十百\d]+[）)]\s*|\d+(?:\.\d+){1,3}[.、．]?\s*|\d+[.、．]\s*)')
LABEL=re.compile(r'^(公式)?\s*[（(]\s*(\d+)\s*[-－—.]\s*(\d+)\s*[）)]\s*$')

def text(p): return ''.join(p.xpath('.//w:t/text()',namespaces=NS))
def issue(code,location,message,severity='error'):
    return {'code':code,'severity':severity,'location':location,'message':message}
def plain_run(p,value):
    r=F.child(p,'r');t=F.child(r,'t');t.text=value;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');return r
def setprop(parent,name,**attrs):
    F.remove_names(parent,{name});return F.child(parent,name,**attrs)

def _layout_values(values):
    if not isinstance(values,dict) or not values or set(values)-set(LAYOUT_BOOL)-set(LAYOUT_NUM):raise ValueError('Unknown or empty layout repair properties')
    for key,value in values.items():
        if key in LAYOUT_BOOL:
            if type(value) is not bool:raise ValueError('Layout boolean required: '+key)
        elif type(value) not in (int,float) or not math.isfinite(value) or not (1 if key=='line_min_pt' else 0)<=value<=144 or value*20!=round(value*20):
            raise ValueError('Layout points must be finite, 0..144 (line >=1), in 0.05pt increments: '+key)
    return values

def _repair_sid(role,values):
    return 'FITRepair_'+role+'_'+F.sha(json.dumps(values,sort_keys=True,separators=(',',':')).encode())[:16]

def _read_repair_style(st):
    if st is None or not st.get(q('styleId'),'').startswith('FITRepair_'):return None
    base=st.find(q('basedOn'));role=base.get(q('val'),'')[4:] if base is not None and base.get(q('val'),'').startswith('FIT_') else None
    if role not in F.ROLES or st.find(q('rPr')) is not None:raise ValueError('Invalid managed layout repair style')
    pp=st.find(q('pPr'));values={}
    if pp is None:raise ValueError('Missing managed repair properties')
    allowed=set(LAYOUT_BOOL.values())|{'spacing'}
    if any(E.QName(x).localname not in allowed for x in pp):raise ValueError('Managed repair style cannot set font, indent, or alignment')
    for key,name in LAYOUT_BOOL.items():
        node=pp.find(q(name))
        if node is not None:
            if node.get(q('val')) not in ('0','1'):raise ValueError('Invalid managed layout boolean')
            values[key]=node.get(q('val'))=='1'
    spacing=pp.find(q('spacing'))
    if spacing is not None:
        if set(spacing.attrib)-{q(x) for x in ('before','after','line','lineRule')}:raise ValueError('Unsupported managed spacing')
        for key,name in LAYOUT_NUM.items():
            if spacing.get(q(name)) is not None:values[key]=int(spacing.get(q(name)))/20
        if 'line_min_pt' in values and spacing.get(q('lineRule'))!='atLeast':raise ValueError('Managed line repair must use minimum spacing')
    values=_layout_values(values)
    # Canonicalize numbers so 20 and 20.0 have the same native style identity.
    values={k:int(v) if isinstance(v,float) and v.is_integer() else v for k,v in values.items()}
    if st.get(q('styleId'))!=_repair_sid(role,values):raise ValueError('Managed repair style hash mismatch')
    return role,values

def _bind_layout_repairs(root,styles,ps,ts,decisions):
    bound={};style_index={s.get(q('styleId')):s for s in styles.findall(q('style'))}
    for p in ps.values():
        sid=p.find('w:pPr/w:pStyle',NS);parsed=_read_repair_style(style_index.get(sid.get(q('val')))) if sid is not None else None
        if parsed:bound[p]={'values':parsed[1],'source':{'kind':'retained','style_id':sid.get(q('val'))}}
    records=decisions.get('layout_repairs',[])
    if not isinstance(records,list):raise ValueError('layout_repairs must be a list')
    for item in records:
        if not isinstance(item,dict) or set(item)-{'kind','id','expected_text','expected_sha256','set','reason'}:raise ValueError('Unknown layout repair fields')
        kind=item.get('kind');target=(ts if kind=='table' else ps).get(item.get('id'))
        if kind not in ('paragraph','table','figure') or target is None:raise ValueError('Invalid layout repair target')
        if kind=='figure' and not (target.findall('.//w:drawing',NS) or target.findall('.//w:pict',NS)):raise ValueError('Figure repair target has no native drawing')
        expected=item.get('expected_sha256')
        if expected is not None and expected!=F.sha(F.xml(target)):raise ValueError('Stale layout repair target hash')
        if 'expected_text' in item and item['expected_text']!=text(target):raise ValueError('Stale layout repair target text')
        if expected is None and (kind!='paragraph' or not item.get('expected_text')):raise ValueError('Table/figure/empty targets require expected_sha256')
        if expected is None and 'expected_text' not in item:raise ValueError('Layout repair requires source evidence')
        values={k:int(v) if isinstance(v,float) and v.is_integer() else v for k,v in _layout_values(item.get('set')).items()}
        targets=[p for p in target.iter(q('p')) if next((a for a in p.iterancestors() if a.tag==q('tbl')),None) is target] if kind=='table' else [target]
        if not targets:raise ValueError('Layout repair target has no owned paragraphs')
        for i,p in enumerate(targets):
            if any(a.tag in F.PROTECTED_TAGS for a in p.iterancestors()):raise ValueError('Cannot alter protected object paragraphs')
            local=dict(values)
            if kind=='table' and i>0:local.pop('page_break_before',None)
            if not local:continue
            prior=bound.get(p,{}).get('values',{})
            bound[p]={'values':prior|local,'source':{k:v for k,v in item.items() if k!='set'}}
    return bound

def _apply_layout_repairs(root,styles,roles,bound,ops):
    for p,item in bound.items():
        if p.getparent() is None or p not in roles:raise ValueError('Reviewed layout target removed or lacks confirmed role')
        values=item['values'];role=roles[p];sid=_repair_sid(role,values)
        st=styles.find('w:style[@w:styleId="'+sid+'"]',NS)
        if st is None:
            st=F.child(styles,'style',type='paragraph',styleId=sid,customStyle='1');F.child(st,'name',val='论文-分页修复-'+role+'-'+sid[-8:]);F.child(st,'basedOn',val=F.style_id(role));F.child(st,'qFormat');pp=F.child(st,'pPr')
            for key,name in LAYOUT_BOOL.items():
                if key in values:F.child(pp,name,val=int(values[key]))
            spacing={LAYOUT_NUM[k]:round(v*20) for k,v in values.items() if k in LAYOUT_NUM}
            if 'line_min_pt' in values:spacing['lineRule']='atLeast'
            if spacing:F.child(pp,'spacing',**spacing)
        pp=F.ensure(p,'pPr',True);before=F.xml(pp).decode();setprop(pp,'pStyle',val=sid)
        for key,name in LAYOUT_BOOL.items():
            if key in values:F.remove_names(pp,{name})
        direct=pp.find(q('spacing'))
        if direct is not None:
            for key,name in LAYOUT_NUM.items():
                if key in values:direct.attrib.pop(q(name),None)
            if 'line_min_pt' in values:direct.attrib.pop(q('lineRule'),None)
            if not direct.attrib:pp.remove(direct)
        ops.append({'kind':'layout_repair','source_target':item['source'],'paragraph_text':text(p),'role':role,'set':values,'style_id':sid,'before_ppr':before,'after_ppr':F.xml(pp).decode()})
    # Stable order prevents second-run style movement after make_styles.
    repair_styles=sorted((s for s in styles.findall(q('style')) if s.get(q('styleId'),'').startswith('FITRepair_')),key=lambda s:s.get(q('styleId')))
    for st in repair_styles:styles.remove(st);styles.append(st)

def _protect_code_styles(styles,code_nodes):
    for p in code_nodes:
        pp=F.ensure(p,'pPr',True);old=pp.find(q('pStyle'));base=old.get(q('val')) if old is not None else 'Normal'
        if base.startswith('FITCode_'):continue
        sid='FITCode_'+F.sha(base.encode())[:16]
        if styles.find('w:style[@w:styleId="'+sid+'"]',NS) is None:
            st=F.child(styles,'style',type='paragraph',styleId=sid,customStyle='1');F.child(st,'name',val='论文-附录代码-'+sid[-8:]);F.child(st,'basedOn',val=base);F.child(st,'qFormat')
        setprop(pp,'pStyle',val=sid)
    for st in sorted((s for s in styles.findall(q('style')) if s.get(q('styleId'),'').startswith('FITCode_')),key=lambda s:s.get(q('styleId'))):styles.remove(st);styles.append(st)

def _picture_width(node,default):
    body=next((x for x in node.iterancestors() if x.tag==q('body')),None);width=default
    if body is not None:
        block=next(x for x in node.iterancestors() if x.getparent() is body)
        for candidate in list(body)[body.index(block):]:
            section=candidate if candidate.tag==q('sectPr') else candidate.find('w:pPr/w:sectPr',NS)
            if section is not None:
                size=section.find(q('pgSz'));margin=section.find(q('pgMar'))
                if size is not None and margin is not None:width=int(size.get(q('w')))-int(margin.get(q('left'),'0'))-int(margin.get(q('right'),'0'))-int(margin.get(q('gutter'),'0'))
                break
    for cell in (x for x in node.iterancestors() if x.tag==q('tc')):
        tw=cell.find('w:tcPr/w:tcW',NS)
        if tw is None or tw.get(q('type'))!='dxa':continue
        table=next((x for x in cell.iterancestors() if x.tag==q('tbl')),None);margins=[]
        for choices in (('start','left'),('end','right')):
            value=None
            for parent in (cell.find('w:tcPr/w:tcMar',NS),table.find('w:tblPr/w:tblCellMar',NS) if table is not None else None):
                if parent is None:continue
                found=next((parent.find(q(k)) for k in choices if parent.find(q(k)) is not None),None)
                if found is not None and found.get(q('type'),'dxa')=='dxa':value=int(found.get(q('w'),'0'));break
            margins.append(108 if value is None else value)
        width=min(width,int(tw.get(q('w')))-sum(margins))
    if width<=0:raise ValueError('Picture container has non-positive width')
    return width

def _fit_pictures(root,width,ops):
    for index,drawing in enumerate((x for x in root.iter() if x.tag in ('{'+WP+'}inline','{'+WP+'}anchor')),1):
        pics=drawing.findall('.//{'+PIC+'}pic');extent=drawing.find('{'+WP+'}extent')
        if len(pics)!=1 or extent is None:continue  # Charts, grouped drawings, and OLE remain intact.
        cx=int(extent.get('cx','0'));cy=int(extent.get('cy','0'));limit=_picture_width(drawing,width)*635
        if cx<=0 or cy<=0:raise ValueError('Native picture has non-positive extent')
        if cx<=limit:continue
        before=F.sha(F.xml(drawing));newcy=max(1,round(cy*limit/cx));updates=[]
        for ext in [extent]+pics[0].findall('./{'+PIC+'}spPr/{'+A+'}xfrm/{'+A+'}ext'):
            old={k:int(ext.get(k,'0')) for k in ('cx','cy')}
            if ext is extent:new={'cx':limit,'cy':newcy}
            else:new={k:max(1,round(v*limit/cx)) for k,v in old.items()}
            for key,value in new.items():ext.set(key,str(value))
            updates.append({'element':E.QName(ext).localname,'before':old,'after':new})
        p=next((x for x in drawing.iterancestors() if x.tag==q('p')),None)
        ops.append({'kind':'picture_fit','drawing_index':index,'paragraph_text':text(p) if p is not None else '', 'before_sha256':before,'after_sha256':F.sha(F.xml(drawing)),'available_width_twips':limit//635,'extent_updates':updates,'media_and_crop':'unchanged'})

def _validate_continuations(ts,settings):
    plans=[]
    for tid,item in settings.items():
        if 'continuation_breaks' not in item:continue
        table=ts[tid];rows=table.findall(q('tr'));breaks=item['continuation_breaks'];headers=item.get('header_rows',1)
        if not isinstance(breaks,list) or not breaks or any(type(b)!=int for b in breaks) or breaks!=sorted(set(breaks)) or type(headers)!=int or headers<1 or any(b<=headers or b>=len(rows) for b in breaks):raise ValueError('Invalid continuation row boundaries')
        if item.get('expected_sha256')!=F.sha(F.xml(table)):raise ValueError('Continuation requires exact source table expected_sha256')
        if 'expected_text' in item and item['expected_text']!=text(table):raise ValueError('Stale continuation table text')
        if item.get('kind') not in ('three-line','complex','panel') or table.getparent().tag!=q('body'):raise ValueError('Only confirmed top-level data tables can be split')
        prior=table.getprevious()
        if prior is not None and prior.tag==q('p') and text(prior).strip()=='续前表':raise ValueError('Table is already a continuation; do not split it again')
        forbidden=F.PROTECTED_TAGS|{q(n) for n in ('fldChar','fldSimple','instrText','bookmarkStart','bookmarkEnd','ins','del','moveFrom','moveTo','commentRangeStart','commentRangeEnd','commentReference','footnoteReference','endnoteReference')}
        for row in rows[:headers]:
            if any(el.tag in forbidden or el.tag==q('tbl') for el in row.iter()):raise ValueError('Continuation headers must contain ordinary text only, without objects, fields, bookmarks, revisions, notes, or nested tables')
        for boundary in [headers]+breaks:
            for cell in rows[boundary].findall(q('tc')):
                merge=cell.find('w:tcPr/w:vMerge',NS)
                if merge is not None and merge.get(q('val'),'continue')!='restart':raise ValueError('Vertical merge crosses continuation boundary')
        plans.append((tid,table,item,rows,breaks,headers))
    return plans

def _bind_caption_targets(ps,ts,roles,records):
    if not isinstance(records,list):raise ValueError('caption_targets must be a list')
    bound=[];seen=set()
    for item in records:
        if not isinstance(item,dict) or set(item)-{'caption_id','expected_text','target_kind','target_id','target_sha256','target_expected_text','reason'}:raise ValueError('Invalid caption target fields')
        p=ps.get(item.get('caption_id'));kind=item.get('target_kind');target=(ts if kind=='table' else ps).get(item.get('target_id'))
        if p is None or target is None or kind not in ('table','figure') or p in seen:raise ValueError('Invalid or repeated caption target')
        if roles.get(p)!=('table_caption' if kind=='table' else 'figure_caption') or text(p)!=item.get('expected_text'):raise ValueError('Caption source role or text mismatch')
        if p.getparent() is not target.getparent() or p.getparent().tag!=q('body'):raise ValueError('Caption and target must be direct document blocks')
        if kind=='figure' and (not target.findall('.//w:drawing',NS) or text(target).strip()):raise ValueError('Figure target must be an unambiguous drawing paragraph')
        if item.get('target_sha256')!=F.sha(F.xml(target)):raise ValueError('Caption target requires exact source target_sha256')
        if 'target_expected_text' in item and item['target_expected_text']!=text(target):raise ValueError('Caption target text mismatch')
        seen.add(p);bound.append((p,target,kind,item))
    return bound

def _position_captions(root,roles,ps,ts,bound,ops):
    body=root.find(q('body'));explicit={p for p,t,k,r in bound};moves=list(bound);source_ids={p:i for i,p in ps.items()}|{t:i for i,t in ts.items()}
    for p,role in list(roles.items()):
        if p in explicit or p.getparent() is not body:continue
        prev=p.getprevious();nxt=p.getnext()
        if role=='table_caption' and prev is not None and prev.tag==q('tbl') and (nxt is None or nxt.tag!=q('tbl')):
            if table_kind(prev)['kind'] in ('three-line','complex','panel'):moves.append((p,prev,'table',None))
        elif role=='figure_caption' and nxt is not None and roles.get(nxt)=='figure' and (prev is None or roles.get(prev)!='figure'):
            moves.append((p,nxt,'figure',None))
    for p,target,kind,evidence in moves:
        already=p.getnext() is target if kind=='table' else p.getprevious() is target
        if already:continue
        before={'caption_index':body.index(p),'target_index':body.index(target)};body.remove(p)
        if kind=='table':target.addprevious(p)
        else:target.addnext(p)
        ops.append({'kind':'caption_position','caption_source_id':source_ids.get(p),'target_source_id':source_ids.get(target),'caption_text':text(p),'target_kind':kind,'position':'before' if kind=='table' else 'after','before':before,'after':{'caption_index':body.index(p),'target_index':body.index(target)},'evidence':evidence or 'unambiguous_adjacent_caption'})

def _split_continuations(plans,roles,table_options,width,ops):
    for tid,table,item,rows,breaks,headers in plans:
        header_nodes=[copy.deepcopy(row) for row in rows[:headers]];header_texts=[text(row) for row in rows[:headers]];tail=table;segments=[]
        clean={k:v for k,v in item.items() if k not in ('continuation_breaks','expected_sha256','expected_text')}
        for start,end in zip(breaks,breaks[1:]+[len(rows)]):
            clone=E.Element(q('tbl'),nsmap=table.nsmap)
            for child in table:
                if child.tag!=q('tr'):clone.append(copy.deepcopy(child))
            for row in header_nodes:clone.append(copy.deepcopy(row))
            for row in rows[start:end]:table.remove(row);clone.append(row)
            label=E.Element(q('p'));plain_run(label,'续前表');F.format_p(label,'continuation_label',{},None,[]);roles[label]='continuation_label'
            tail.addnext(label);label.addnext(clone);tail=clone
            for row in clone.findall(q('tr'))[:headers]:
                for p in row.findall('.//w:p',NS):roles[p]='table_text'
            segment_settings=dict(clean)
            if 'panel_rows' in segment_settings:segment_settings['panel_rows']=[r for r in clean['panel_rows'] if r<headers]+[headers+r-start for r in clean['panel_rows'] if start<=r<end]
            table_options[clone]=segment_settings;_table(clone,segment_settings,width);segments.append({'start_row':start,'end_row_exclusive':end,'data_row_texts':[text(r) for r in rows[start:end]]})
        first_settings=dict(clean)
        if 'panel_rows' in first_settings:first_settings['panel_rows']=[r for r in first_settings['panel_rows'] if r<breaks[0]]
        table_options[table]=first_settings;_table(table,first_settings,width)
        ops.append({'kind':'table_continuation','source_table_id':tid,'source_table_sha256':item['expected_sha256'],'breaks':breaks,'header_rows':headers,'header_row_texts':header_texts,'added_header_texts':[text(p) for row in header_nodes for p in row.iter(q('p'))]*len(breaks),'added_header_xml':[F.xml(r).decode() for r in header_nodes],'label_count':len(breaks),'label':'续前表','segments':segments})
def field(p,code,cached,hidden=False):
    for kind,val in [('fldChar','begin'),('instrText',code),('fldChar','separate'),('t',str(cached)),('fldChar','end')]:
        r=F.child(p,'r')
        if hidden:F.child(F.child(r,'rPr'),'vanish')
        if kind=='fldChar':F.child(r,kind,fldCharType=val)
        else:
            e=F.child(r,kind);e.text=val;e.set('{http://www.w3.org/XML/1998/namespace}space','preserve')

def table_kind(t):
    rows=t.findall(q('tr'));s=text(t)
    if '诚信承诺' in s or '学生承诺' in s:return {'kind':'pledge','header_rows':0}
    if '指导教师' in s and ('学号' in s or '姓名' in s):return {'kind':'cover','header_rows':0}
    if len(rows)==1 and t.findall('.//m:oMath',NS) and len(rows[0].findall(q('tc')))<=3:
        return {'kind':'formula-layout','header_rows':0}
    panels=[i for i,r in enumerate(rows) if re.search(r'Panel\s+[A-Z]',text(r),re.I)]
    if panels:return {'kind':'panel','header_rows':1,'panel_rows':panels}
    cells=[c for row in rows for c in row.findall(q('tc'))]
    lengths=[len(text(c)) for c in cells]
    managed_inner=t.find('w:tblPr/w:tblBorders/w:insideV',NS)
    managed_outer=t.find('w:tblPr/w:tblBorders/w:top',NS)
    if managed_inner is not None and managed_inner.get(q('val'))=='single' and managed_outer is not None and managed_outer.get(q('sz'))=='12':return {'kind':'complex','header_rows':1}
    # Long prose tables are different from compact numerical tables.
    return {'kind':'complex' if max(lengths or [0])>90 else 'three-line','header_rows':1}

def inspect_document(source):
    source=Path(source).resolve();a=Audit(source);a.infer_roles()
    mapping={'source_sha256':a.hash,'reviewed':False,'paragraphs':{},'tables':{},'story_paragraphs':{},'preserve':{}}
    candidates=[];issues=[];body=a.doc.find(q('body'))
    table_settings={t:table_kind(t) for t in a.ts.values()}
    cover_titles=[]
    for t,setting in table_settings.items():
        if setting['kind']=='cover':
            for row in t.findall(q('tr')):
                cells=row.findall(q('tc'))
                if cells and re.match(r'^题\s*目',text(cells[0]).strip()):
                    cover_titles.extend(p for c in cells[1:] for p in c.findall('.//w:p',NS))
    cover_title_text=''.join(text(p).strip() for p in cover_titles)
    abstract_head=next((p for p in a.ps.values() if re.sub(r'\s+','',text(p)) in ('摘要','中文摘要')),None)
    preceding_title=abstract_head.getprevious() if abstract_head is not None else None
    while preceding_title is not None and preceding_title.tag==q('p') and not text(preceding_title).strip():preceding_title=preceding_title.getprevious()
    english_head=next((p for p in a.ps.values() if text(p).strip().casefold()=='abstract'),None)
    english_title=english_head.getprevious() if english_head is not None else None
    while english_title is not None and english_title.tag==q('p') and not text(english_title).strip():english_title=english_title.getprevious()
    for tid,t in a.ts.items():mapping['tables'][tid]=table_settings[t]
    for pid,p in a.ps.items():
        value=text(p);known=a.roles.get(pid);role=known[0] if known else None
        repaired=_read_repair_style(a.styles.get(a.sid(p)))
        if repaired:role=repaired[0]
        if a.sid(p).startswith('FITCode_'):role='appendix_code'
        # Template spacer paragraphs carry real cover-page geometry. Keep their
        # identity on a repeat run instead of treating them as removable body
        # whitespace after the form's isolated styles have been imported.
        if a.sid(p).startswith('FITFrontTpl_') and p.getparent() is body and not value.strip():role='cover_info'
        if any(x.tag in F.PROTECTED_TAGS for x in p.iterancestors()):
            obj=next(x for x in p.iterancestors() if x.tag in F.PROTECTED_TAGS)
            mapping['preserve'][pid]={'expected_text':text(p),'reason':'图形或嵌入对象内部文字保留原对象，仍须逐页视觉复核','object_sha256':F.sha(F.xml(obj)),'pending_visual':True}
            continue
        table=next((x for x in p.iterancestors() if x.tag==q('tbl')),None)
        if table is not None:
            kind=table_settings[table]['kind']
            if kind in ('cover','pledge'):
                # Leave semantic unknowns in school component tables to supplied
                # reviewed mapping; do not turn signatures into normal body.
                if p in cover_titles:role='cover_title'
                elif kind=='cover':role='cover_info'
                elif value.strip() in ('学生承诺','指导教师承诺'):role='commitment_label'
                elif '我承诺' in value:role='commitment_body'
                elif '签名' in value or re.fullmatch(r'(?:\d{4})?年(?:\d{1,2})?月(?:\d{1,2})?日',re.sub(r'\s+','',value)):role='commitment_signature'
                else:role='commitment_info'
            elif kind=='formula-layout':role='equation' if p.findall('.//m:oMath',NS) else 'equation_number'
            else:role='table_text'
        if p is preceding_title and value.strip() and (re.sub(r'\s+','',value)==re.sub(r'\s+','',cover_title_text) or 5<len(value)<160 and not re.search(r'[。；：]',value)):role='cn_title'
        if p is english_title and 8<len(value)<350 and re.search(r'[A-Za-z]{3}',value) and not re.search(r'[\u4e00-\u9fff]',value):role='en_title'
        # Shape text belongs to the drawing, not to the containing body line.
        # Treat a standalone editable flowchart like a picture so fixed body
        # line height cannot let its inline VML group overlap preceding text.
        outside_object_text=''.join(t.text or '' for t in p.iter(q('t')) if not any(a.tag in F.PROTECTED_TAGS for a in t.iterancestors()))
        if p.getparent() is body and (p.findall('.//w:drawing',NS) or p.findall('.//w:pict',NS)) and not outside_object_text.strip():
            role='cover_info' if abstract_head is not None and body.index(p)<body.index(abstract_head) else 'figure'
        if role in ('body_math','table_math'):role='body' if role=='body_math' else 'table_text'
        if p.getparent() is body and p.findall('.//m:oMath',NS) and (not value.strip() or LABEL.match(value.strip())):role='equation'
        if value.strip()=='本科生毕业论文（设计）':role='cover_banner'
        if value.strip()=='（正本）':role='cover_copy'
        if '诚信承诺书' in value and len(value)<45:role='commitment_heading'
        if role not in ('toc1','toc2') and re.match(r'^附\s*录(?:[一二三四五六七八九十A-Z0-9]|\s+)',value) and len(value)<80 and not re.search(r'[。；，]',value):role='appendix_heading'
        if re.match(r'^表\s*\d+[.－—-]\d+(?:\s|$)',value.strip()) and len(value)<100:role='table_caption'
        if re.match(r'^附表\s*\d+\s',value.strip()) and len(value)<100:role='table_caption'
        if re.match(r'^图\s*\d+[.－—-]\d+(?:\s|$)',value.strip()) and len(value)<100:role='figure_caption'
        if role in ('table_caption','figure_caption') and (len(value)>120 or re.search(r'[。；]',value)):role='body'
        if role in LEVELS and len(value)<60:
            if re.match(r'^[一二三四五六七八九十]+、',value):role='chapter'
            elif re.match(r'^（[一二三四五六七八九十]+）',value):role='section'
            elif re.match(r'^\d+[.．](?!\d)',value) and role!='chapter':role='subsection'
            elif re.match(r'^（\d+）',value):role='subsubsection'
        if role in F.ROLES or role=='appendix_code':
            mapping['paragraphs'][pid]={'role':role,'expected_text':text(p)}
            if repaired:mapping['paragraphs'][pid]['layout_overrides']=repaired[1]
            if role in LEVELS and known and known[1]!='high':candidates.append({'id':pid,'role':role,'text':value,'reason':'编号外观不能独自证明标题层级'})
            if role in LEVELS and known and known[1]=='high' and not PREFIX.match(value):
                native_num=a.effective_paragraph_container(p,'numPr')
                native_id=native_num.find(q('numId')) if native_num is not None else None
                canonical=re.sub(r'\s+','',value) in {'绪论','引言','结论','研究结论','总结','结语'}
                if not canonical and (native_id is None or native_id.get(q('val')) in (None,'0')):
                    candidates.append({'id':pid,'role':role,'text':value,'reason':'只有标题样式、缺明确编号或已知语义标题，须结合上下文确认层级'})
        elif value.strip():
            candidates.append({'id':pid,'role':None,'text':value,'reason':'需要结合上下文确认格式角色'})
    for name,data in a.parts.items():
        if not F.allowed_story(name):continue
        sr=F.parse(data);ps,_=F.indexed(sr);sm={}
        for pid,p in ps.items():
            if not (text(p).strip() or p.findall('.//w:fldChar',NS)):continue
            if name.startswith('word/header'):role='header'
            elif name.startswith('word/footer'):role='page_number'
            else:role='reference'
            sm[pid]={'role':role,'expected_text':text(p)}
        if sm:mapping['story_paragraphs'][name]=sm
    for c in candidates:issues.append(issue('role_confirmation',c['id'],c['reason'],'warning'))
    citation_review=N.inspect_parts(a.parts)
    return {'source_sha256':a.hash,'map':mapping,'mapping':mapping,'candidates':candidates,'issues':issues,'citation_candidates':citation_review['candidates'],
            'counts':{'paragraphs':len(a.ps),'tables':len(a.ts)},'title_candidate':a.title}

def _table(t,setting,width):
    kind=setting['kind']
    if kind not in ('three-line','complex','panel','cover'):return
    # The school cover is a fixed-width form, not a page-width data table.
    # Its ruled fields and merged date cells use the template's own grid.
    if kind=='cover':return
    panels=setting.get('panel_rows',[i for i,r in enumerate(t.findall(q('tr'))) if re.search(r'Panel\s+[A-Z]',text(r),re.I)])
    if kind!='cover':F.format_table(t,{**setting,'kind':'three-line' if kind=='panel' else kind})
    tp=F.ensure(t,'tblPr',True);setprop(tp,'tblW',w=width,type='dxa');setprop(tp,'tblLayout',type='fixed')
    if kind=='cover':setprop(tp,'jc',val='center');setprop(tp,'tblInd',w=0,type='dxa')
    grid=t.find(q('tblGrid'))
    if grid is None:return
    cols=grid.findall(q('gridCol'));old=[int(c.get(q('w'),'1')) or 1 for c in cols]
    if not old:return
    new=setting.get('column_widths_twips')
    if new is not None:
        if len(new)!=len(old) or any(type(v)!=int or v<=0 for v in new) or sum(new)>width:raise ValueError('Invalid reviewed column widths')
    else:
        new=[round(x*width/sum(old)) for x in old];new[-1]+=width-sum(new)
    for c,v in zip(cols,new):c.set(q('w'),str(v))
    for ri,row in enumerate(t.findall(q('tr'))):
        rp=F.ensure(row,'trPr',True);setprop(rp,'cantSplit')
        offset=int(row.find('w:trPr/w:gridBefore',NS).get(q('val'),'0')) if row.find('w:trPr/w:gridBefore',NS) is not None else 0
        for cell in row.findall(q('tc')):
            cp=F.ensure(cell,'tcPr',True);span=cell.find('w:tcPr/w:gridSpan',NS);n=int(span.get(q('val'),'1')) if span is not None else 1
            setprop(cp,'tcW',w=sum(new[offset:offset+n]),type='dxa');offset+=n
            if kind=='panel' and ri in panels:
                b=F.ensure(cp,'tcBorders')
                for side in ('top','bottom'):setprop(b,side,val='single',sz=12 if ri==0 and side=='top' else 6,color='000000')
        for p in row.findall('.//w:p',NS):
            pp=F.ensure(p,'pPr',True)
            # Keep panel headings with their following data, not entire tables.
            if kind!='cover' and (ri in panels or ri<setting.get('header_rows',1)):setprop(pp,'keepNext',val=1)

def _numbering(parts,root,roles,ops):
    nr=F.parse(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else E.Element(q('numbering'),nsmap={'w':W})
    existing=next((n for n in nr.findall(q('abstractNum')) if n.find(q('name')) is not None and n.find(q('name')).get(q('val'))=='FITv10Headings'),None)
    if existing is None:
        aid=max([int(x.get(q('abstractNumId'))) for x in nr.findall(q('abstractNum'))]+[-1])+1
        existing=E.Element(q('abstractNum'));existing.set(q('abstractNumId'),str(aid));F.child(existing,'multiLevelType',val='multilevel');F.child(existing,'name',val='FITv10Headings')
        for i,(fmt,label) in enumerate([('chineseCounting','%1、'),('chineseCounting','（%2）'),('decimal','%3.'),('decimal','（%4）')]):
            lv=F.child(existing,'lvl',ilvl=i);F.child(lv,'start',val=1);F.child(lv,'numFmt',val=fmt);F.child(lv,'lvlText',val=label);F.child(lv,'suff',val='nothing');F.child(lv,'pStyle',val=F.style_id(LEVELS[i]))
        nr.insert(next((i for i,x in enumerate(nr) if x.tag==q('num')),len(nr)),existing)
    aid=existing.get(q('abstractNumId'));num=next((n for n in nr.findall(q('num')) if n.find(q('abstractNumId')) is not None and n.find(q('abstractNumId')).get(q('val'))==aid and not n.findall(q('lvlOverride'))),None)
    if num is None:
        nid=max([int(x.get(q('numId'))) for x in nr.findall(q('num'))]+[0])+1;num=F.child(nr,'num',numId=nid);F.child(num,'abstractNumId',val=aid)
    nid=num.get(q('numId'));active_nid=nid;restart_heading=None;claimed_restarts=set()
    # An unnumbered top-level heading does not advance Word's level 0, so the
    # following levels otherwise continue the preceding chapter's counters.
    # Isolate each such block in its own native list. Reuse its existing owned
    # numId on later runs; ordinary numbered chapters retain their shared list.
    restart_roles={'conclusion_heading','reference_heading','appendix_heading','acknowledgement_heading'}
    for p in root.find(q('body')):
        role=roles.get(p)
        if role in restart_roles:
            active_nid=None;restart_heading=text(p)
        if role not in LEVELS:continue
        if role=='chapter':active_nid=nid;restart_heading=None
        elif active_nid is None:
            prior=p.find('w:pPr/w:numPr/w:numId',NS)
            candidate=next((n for n in nr.findall(q('num')) if n.get(q('numId'))==prior.get(q('val'))),None) if prior is not None else None
            candidate_aid=candidate.find(q('abstractNumId')) if candidate is not None else None
            overrides=candidate.findall(q('lvlOverride')) if candidate is not None else []
            reusable=(candidate_aid is not None and candidate_aid.get(q('val'))==aid and
                      candidate.get(q('numId')) not in claimed_restarts and len(overrides)==3 and
                      {x.get(q('ilvl')) for x in overrides}=={'1','2','3'} and
                      all(len(x)==1 and x[0].tag==q('startOverride') and x[0].get(q('val'))=='1' for x in overrides))
            if not reusable:
                next_nid=max([int(x.get(q('numId'))) for x in nr.findall(q('num'))]+[0])+1
                candidate=F.child(nr,'num',numId=next_nid);F.child(candidate,'abstractNumId',val=aid)
                for level in (1,2,3):F.child(F.child(candidate,'lvlOverride',ilvl=level),'startOverride',val=1)
                ops.append({'kind':'heading_number_restart','heading':restart_heading,'num_id':str(next_nid),'levels':[1,2,3],'start':1})
            active_nid=candidate.get(q('numId'));claimed_restarts.add(active_nid)
        old=text(p);match=PREFIX.match(old)
        if match:
            left=match.end()
            for t in p.findall('.//w:t',NS):
                value=t.text or '';take=min(left,len(value));t.text=value[take:];left-=take
                if not left:break
            ops.append({'kind':'heading_prefix','locator':{'text':old},'before':old,'after':text(p),'level':LEVELS.index(role)})
        pp=F.ensure(p,'pPr',True);np=setprop(pp,'numPr');F.child(np,'ilvl',val=LEVELS.index(role));F.child(np,'numId',val=active_nid)
    chapters=[p for p in root.find(q('body')) if roles.get(p)=='chapter']
    for index,p in enumerate(chapters,1):
        existing_codes=' '.join(p.xpath('.//w:instrText/text()',namespaces=NS))
        if 'SEQ FITChapter' not in existing_codes:
            field(p,' SEQ FITChapter '+('\\r 1 ' if index==1 else ''),index,True)
            for counter in ('FITEquation','FITTable','FITFigure'):field(p,' SEQ '+counter+' \\r 0 \\h ',0,True)
            ops.append({'kind':'chapter_counter_fields','locator':{'heading':text(p)},'counters':['FITChapter','FITEquation','FITTable','FITFigure']})
    parts['word/numbering.xml']=F.xml(nr)
    rels=F.parse(parts[L.RELS]);ct=F.parse(parts['[Content_Types].xml'])
    if not any(x.get('Type')==R+'/numbering' for x in rels):E.SubElement(rels,'{'+PKG+'}Relationship',Id='rIdFITv10Numbering',Type=R+'/numbering',Target='numbering.xml')
    if not any(x.get('PartName')=='/word/numbering.xml' for x in ct):E.SubElement(ct,'{'+CT+'}Override',PartName='/word/numbering.xml',ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
    parts[L.RELS]=F.xml(rels);parts['[Content_Types].xml']=F.xml(ct)

def _equation(p,width,ops,issues,assigned=None):
    if not p.findall('.//m:oMath',NS):
        if p.findall('.//w:object',NS):issues.append(issue('legacy_equation_editability',text(p),'旧式方程须在两种办公软件中实测可编辑性。','warning'))
        return
    old=text(p)
    if not old.strip() and assigned is not None:
        plain_run(p,f'公式（{assigned[0]}-{assigned[1]}）');ops.append({'kind':'equation_label_added','before':old,'after':text(p)})
    label=LABEL.match(text(p).strip())
    if old.strip() and not label:
        issues.append(issue('equation_label_ambiguous',old,'公式段混有正文或无法确定的编号，保留原内容。'));return
    pp=F.ensure(p,'pPr',True);setprop(pp,'jc',val='left');setprop(pp,'spacing',before=0,after=0,line=400,lineRule='atLeast');F.remove_names(pp,{'pBdr'})
    tabs=setprop(pp,'tabs');F.child(tabs,'tab',val='center',pos=round(width/2));F.child(tabs,'tab',val='right',pos=width)
    for outer in list(p.findall('{'+M+'}oMathPara')):
        idx=p.index(outer)
        for math in list(outer.findall('{'+M+'}oMath')):outer.remove(math);p.insert(idx,math);idx+=1
        p.remove(outer)
    # Preserve fields/bookmarks; remove only literal spacing tab runs, then
    # introduce exactly one tab before math and one before its label.
    for r in list(p.findall(q('r'))):
        if len(r)>0 and all(x.tag in {q('rPr'),q('tab')} or (x.tag==q('t') and not (x.text or '').strip()) for x in r):p.remove(r)
    math=p.find('{'+M+'}oMath');r=E.Element(q('r'));F.child(r,'tab');p.insert(p.index(math),r)
    direct_text=[r for r in p.findall(q('r')) if r.findall(q('t'))]
    if label:
        wanted=f'公式（{label.group(2)}-{label.group(3)}）'
        # Do not rewrite field cached label pieces. Dynamic labels are already
        # editable and retained; controlled native field refresh follows.
        if not p.findall('.//w:fldChar',NS):
            for r in direct_text:p.remove(r)
            r=F.child(p,'r');F.child(r,'tab');plain_run(p,'公式（');field(p,' SEQ FITChapter \\c ',label.group(2));plain_run(p,'-');field(p,' SEQ FITEquation ',label.group(3));plain_run(p,'）')
            ops.append({'kind':'equation_counter_fields','before':old,'after':wanted})
            if wanted!=old:ops.append({'kind':'equation_label','locator':{'text':old},'before':old,'after':wanted})
        elif direct_text:
            r=E.Element(q('r'));F.child(r,'tab');anchor=direct_text[0]
            if anchor.getprevious() is not None and anchor.getprevious().tag==q('bookmarkStart'):anchor=anchor.getprevious()
            p.insert(p.index(anchor),r)
    else:issues.append(issue('equation_number_missing','equation','缺少公式编号，需要结合章节确定，不猜填。','warning'))

def _toc(root,roles,ops,issues):
    toc_head=next((p for p,r in roles.items() if r=='toc_heading'),None)
    found=False
    instruction=' TOC \\t "论文-中文摘要标题,1,论文-英文摘要标题,1,论文-章标题,1,论文-一级节标题,2,论文-参考文献标题,1,论文-附录标题,1,论文-致谢标题,1,论文-结论标题,1" \\h \\z '
    for e in root.findall('.//w:instrText',NS):
        if re.match(r'\s*TOC\b',e.text or '',re.I):
            old=e.text;found=True;e.text=instruction
            if old!=instruction:ops.append({'kind':'toc_instruction','before':old,'after':instruction})
    if not found and toc_head is not None:
        # Do not delete manually typed TOC entries until their region has been
        # explicitly confirmed, otherwise duplicated entries would mislead.
        siblings=list(toc_head.getparent());i=siblings.index(toc_head)
        manual=[]
        for p in siblings[i+1:]:
            if roles.get(p) in LEVELS:break
            if roles.get(p) in ('toc1','toc2','toc3'):manual.append(p)
            elif text(p).strip():break
        for p in manual:ops.append({'kind':'manual_toc_entry','before':text(p),'after':''});p.getparent().remove(p);roles.pop(p,None)
        p=E.Element(q('p'));F.child(F.child(p,'pPr'),'pStyle',val='FIT_toc1');field(p,instruction,'');toc_head.addnext(p);roles[p]='toc1';ops.append({'kind':'toc_create','before':'','after':instruction})
    elif not found:issues.append(issue('toc_heading_missing','document','没有可确定的目录位置，需在语义决定中补定。','warning'))

def _appendix(root,roles,ops):
    body=root.find(q('body'));ack=next((p for p,r in roles.items() if r=='acknowledgement_heading' and p.getparent() is body),None)
    app=next((p for p,r in roles.items() if r=='appendix_heading' and p.getparent() is body),None)
    if ack is not None and app is not None and body.index(app)>body.index(ack):
        start=body.index(app);end=next((i for i in range(start+1,len(body)) if body[i].tag==q('sectPr')),len(body));block=list(body)[start:end]
        for x in block:body.remove(x)
        index=body.index(ack)
        for i,x in enumerate(block):body.insert(index+i,x)
        ops.append({'kind':'move_appendix','before':'after_acknowledgement','after':'before_acknowledgement','block_count':len(block),'texts':[text(x) for x in block]})

def _trim_terminal_page_breaks(root,roles,ops):
    """Remove only terminal, otherwise empty, manual-page-break paragraphs.

    Run after section reordering: a separator before an appendix can become
    the final paragraph when the appendix is moved before acknowledgements.
    A section property, bookmark, revision, note or any unknown content is a
    hard boundary. Ordinary breaks inside the document are never removed.
    """
    body=root.find(q('body'));pids={p:k for k,p in F.indexed(root)[0].items()}
    formatting={q(n) for n in ('pStyle','rStyle','jc','spacing','ind','keepNext',
        'keepLines','widowControl','pageBreakBefore','snapToGrid','contextualSpacing',
        'rFonts','sz','szCs','color','b','bCs','i','iCs','lang','vertAlign','vanish')}
    allowed={q(n) for n in ('p','pPr','r','rPr','t','br')}|formatting
    pending=[]
    for p in reversed(list(body)):
        if p.tag==q('sectPr') and p is body[-1]:continue
        if p.tag!=q('p') or any(e.tag not in allowed for e in p.iter()):break
        if text(p).strip(' \u3000'):break
        breaks=p.findall('.//w:br',NS)
        if any(e.get(q('type'))!='page' for e in breaks):break
        pending.append((p,breaks))
    # Empty paragraphs may follow the page break. Leave a purely empty tail
    # alone; it may be the mandatory paragraph after a table.
    if not any(breaks for _,breaks in pending):return
    for p,breaks in pending:
        if not breaks:continue
        ops.append({'kind':'remove_terminal_page_break','locator':pids[p],
            'before_sha256':F.sha(F.xml(p)),'page_breaks':len(breaks),
            'before':F.ptext(p),'after':''})
        body.remove(p);roles.pop(p,None)


def _caption_fields(p,ops):
    old=text(p);match=re.match(r'^(表|图)\s*(\d+)\s*[-.－—]\s*(\d+)',old)
    if not match or p.findall('.//w:fldChar',NS):return
    texts=p.findall('.//w:t',NS)
    if not texts or texts[0].getparent().getparent() is not p:return
    target=texts[0].getparent();remaining=match.end()
    for t in texts:
        value=t.text or '';count=min(len(value),remaining);t.text=value[count:];remaining-=count
        if not remaining:break
    temp=E.Element(q('p'));plain_run(temp,match.group(1));field(temp,' SEQ FITChapter \\c ',match.group(2));plain_run(temp,'-');field(temp,' SEQ '+('FITTable' if match.group(1)=='表' else 'FITFigure')+' ',match.group(3))
    for r in list(temp):target.addprevious(r)
    ops.append({'kind':'caption_counter_fields','before':old,'after':text(p)})

def _replace_plain_span(p,start,end,new_nodes):
    """Replace a visible span only in ordinary text runs; retain all siblings."""
    runs=[r for r in p.iter(q('r')) if next((a for a in r.iterancestors() if a.tag==q('p')),None) is p and not any(a.tag.startswith('{'+M+'}') for a in r.iterancestors())]
    offsets=[];offset=0
    for r in runs:
        value=''.join(r.xpath('./w:t/text()',namespaces=NS));offsets.append((r,offset,offset+len(value),value));offset+=len(value)
    touched=[x for x in offsets if x[2]>start and x[1]<end or start==end and x[1]<=start<=x[2]]
    if not touched or any(any(c.tag not in {q('rPr'),q('t')} for c in r) for r,a,b,v in touched):raise ValueError('Requested text span intersects non-text run')
    inserted=False
    for r,a,b,v in touched:
        parent=r.getparent();at=parent.index(r);frags=[]
        if a<start:frags.append(C.sliced_run(r,0,start-a))
        if not inserted:frags.extend(copy.deepcopy(new_nodes));inserted=True
        if b>end:frags.append(C.sliced_run(r,end-a,len(v)))
        parent.remove(r)
        for i,n in enumerate(frags):parent.insert(at+i,n)

def _author_displays(parts,root,styles,ops,issues,protected=()):
    current=dict(parts);current[DOC]=F.xml(root);current['word/styles.xml']=F.xml(styles)
    evidence=N.inspect_parts(current);refs={r['pid']:r for r in evidence['references']};ps,_=F.indexed(root);changes={}
    for ref in refs.values():
        if ref.get('issues') or len(ref['authors'])<=3:continue
        names=[a['exact_name'] for a in ref['authors'][:3]];western=ref['authors'][0]['language']=='en'
        replacement=(', ' if western else ',').join(names)+(', et al.' if western else ',等')
        span=ref['author_list'];changes.setdefault(ref['pid'],[]).append((span['start'],span['end'],replacement,'bibliography_author_display',ref))
    for item in evidence['matches']:
        if 'author_count_needs_review' not in item['issues'] or len(item['reference_pids'])!=1:continue
        ref=refs[item['reference_pids'][0]]
        if ref['issues'] or not ref['authors']:continue
        authors=ref['authors'];count=ref['author_count_min']
        if count==1:new=authors[0]['display_name']
        elif count==2:new='和'.join(a['display_name'] for a in authors)
        else:new=authors[0]['display_name']+(' et al.' if authors[0]['language']=='en' else '等')
        changes.setdefault(item['pid'],[]).append((item['author_start'],item['author_end'],new,'narrative_author_display',ref))
    for pid,items in changes.items():
        p=ps[pid]
        if p in protected:continue
        for start,end,replacement,kind,ref in sorted(items,reverse=True,key=lambda x:x[0]):
            before=text(p)
            # Offsets from the conservative source parser include mathematical
            # instructions; require exact visible source before applying them.
            expected=ref['author_list']['expected_text'] if kind=='bibliography_author_display' else C.ptext(p)[start:end]
            if before[start:end]!=expected:
                issues.append(issue('author_span_complex',pid,'作者范围与可见文本不一致，保留并核对。'));continue
            temp=E.Element(q('p'));plain_run(temp,replacement)
            try:_replace_plain_span(p,start,end,list(temp))
            except ValueError as exc:issues.append(issue('author_span_complex',pid,str(exc)));continue
            ops.append({'kind':kind,'locator':pid,'before':before,'after':text(p),'edit_span':{'start':start,'end':end,'expected_text':before[start:end],'replacement':replacement},'reference':{'number':ref['reference_number'],'source_text':ref['expected_text'],'authors':ref['authors'],'span':ref['author_list']}})

def _citation_insertions(root,records,ops):
    ps,_=F.indexed(root);original={pid:text(p) for pid,p in ps.items()}
    for record in sorted(records,key=lambda r:(r['pid'],-r['position'])):
        pid=record['pid'];p=ps[pid];source=original[pid];claim=record['claim'];pos=record['position'];marker=record['text']
        if source!=record['expected_text'] or source[claim['start']:claim['end']]!=claim['expected_text'] or pos!=claim['end']:raise ValueError('Citation claim/source span mismatch')
        if not re.fullmatch(r'\[\d+(?:,\d+)*\]',marker):raise ValueError('Only numeric citation markers may be inserted')
        numbers={int(n) for n in re.findall(r'\d+',marker)}
        if numbers!={r['reference_number'] for r in record['references']}:raise ValueError('Citation marker/reference numbers mismatch')
        for ref in record['references']:
            if original.get(ref['pid'])!=ref['expected_text'] or not re.match(r'^\s*\['+str(ref['reference_number'])+r'\]',ref['expected_text']):raise ValueError('Bibliography evidence mismatch')
            parsed=N._guess_reference(ref['expected_text'])
            if not parsed or not parsed['authors'] or parsed['authors'][0]['display_name'] not in claim['expected_text'] or not any(y in claim['expected_text'] for y in parsed['years']):raise ValueError('Claim author/year does not match bibliography')
        if marker in claim['expected_text']:raise ValueError('Duplicate citation marker in reviewed claim')
        before=text(p);temp=E.Element(q('p'));run=plain_run(temp,marker);C.style_run(run);_replace_plain_span(p,pos,pos,list(temp))
        ops.append({'kind':'citation_marker_insertion','locator':pid,'before':before,'after':text(p),'edit_span':{'start':pos,'end':pos,'expected_text':'','replacement':marker},'claim':claim,'references':record['references']})

def _equation_crossrefs(root,roles,ops,issues):
    names={e.get(q('name')) for e in root.findall('.//w:bookmarkStart',NS)}
    bid=max([int(e.get(q('id'),'0')) for e in root.findall('.//w:bookmarkStart',NS)]+[0])+1;targets={}
    for p,role in roles.items():
        if role!='equation':continue
        label=LABEL.match(text(p).strip())
        if label:targets.setdefault(label.group(),[]).append(p)
    for label,targets_here in targets.items():
        if len(targets_here)!=1:continue
        p=targets_here[0];name='FITv10Equation'+hashlib.sha256(label.encode()).hexdigest()[:16]
        if name not in names:
            rs=p.findall(q('r'));first=next((r for r in rs if '公式' in text(r)),None);last=next((r for r in reversed(rs) if '）' in text(r)),None)
            if first is None or last is None:continue
            start=E.Element(q('bookmarkStart'));start.set(q('id'),str(bid));start.set(q('name'),name);end=E.Element(q('bookmarkEnd'));end.set(q('id'),str(bid));bid+=1
            first.addprevious(start);last.addnext(end);names.add(name)
            ops.append({'kind':'equation_bookmark','label':label,'bookmark':name})
        for bodyp,bodyrole in list(roles.items()):
            if bodyrole not in ('body','cn_abstract','en_abstract'):continue
            # Skip existing REF fields for this label. Remaining plain mentions
            # are transformed without rebuilding embedded equations or links.
            instructions=' '.join(bodyp.xpath('.//w:instrText/text()',namespaces=NS))
            if name in instructions:continue
            for match in reversed(list(re.finditer(re.escape(label),text(bodyp)))):
                before=text(bodyp);temp=E.Element(q('p'));field(temp,' REF '+name+' \\h ',label)
                try:_replace_plain_span(bodyp,match.start(),match.end(),list(temp))
                except ValueError as exc:issues.append(issue('equation_reference_complex',label,str(exc)));continue
                ops.append({'kind':'equation_cross_reference','before':before,'after':text(bodyp),'bookmark':name,'label':label})

def _numeric_citations(root,styles,roles,ops,issues):
    contexts=C.reference_contexts(root,styles);pids={p:k for k,p in C.indexed(root).items()}
    for p,role in roles.items():
        if role not in ('body','cn_abstract','en_abstract'):continue
        runs=[r for r in p.iter(q('r')) if next((a for a in r.iterancestors() if a.tag==q('p')),None) is p and not any(a.tag.startswith('{'+M+'}') for a in r.iterancestors())]
        visible=''.join(''.join(r.xpath('./w:t/text()',namespaces=NS)) for r in runs)
        matches=list(C.CITATION.finditer(visible))
        if not matches:continue
        pid=pids.get(p)
        if contexts.get(pid,False):continue
        spans=[{'start':m.start(),'end':m.end(),'expected_text':m.group()} for m in matches]
        C.add_style(styles);offset=0
        for run in runs:
            value=''.join(run.xpath('./w:t/text()',namespaces=NS));start=offset;offset+=len(value)
            touching=[s for s in spans if s['end']>start and s['start']<offset]
            if not touching:continue
            if any(x.tag not in {q('rPr'),q('t')} for x in run):
                issues.append(issue('numeric_citation_mixed_run',pid,'引用字符与非文字对象同处单个运行，需确认最小拆分范围。'));continue
            if any(a.tag in {q('ins'),q('del')} for a in run.iterancestors()):
                issues.append(issue('citation_in_revision',pid,'引用处有未接受修订，保留并待作者确认。'));continue
            cuts={0,len(value)}
            for span in touching:cuts.update((max(0,span['start']-start),min(len(value),span['end']-start)))
            cuts=sorted(cuts);parent=run.getparent();at=parent.index(run);fragments=[]
            for lo,hi in zip(cuts,cuts[1:]):
                frag=C.sliced_run(run,lo,hi)
                if any(s['start']<=start+lo and start+hi<=s['end'] for s in touching):C.style_run(frag)
                frag.tail=run.tail if hi==len(value) else None;fragments.append(frag)
            parent.remove(run)
            for i,frag in enumerate(fragments):parent.insert(at+i,frag)
        ops.append({'kind':'numeric_citation_style','locator':pid,'spans':spans,'before':text(p),'after':text(p)})

def _repair_reference_bookmarks(root,roles,ops,issues):
    bookmarks={e.get(q('name')) for e in root.findall('.//w:bookmarkStart',NS)}
    refs={}
    for p,role in roles.items():
        if role!='reference':continue
        m=re.match(r'^\s*\[(\d+)\]',text(p))
        if m:refs.setdefault(m.group(1),[]).append(p)
    next_id=max([int(e.get(q('id'),'0')) for e in root.findall('.//w:bookmarkStart',NS)]+[0])+1
    for p,role in list(roles.items()):
        if role not in ('body','cn_abstract','en_abstract'):continue
        stack=[]
        for e in p.iter():
            if e.tag==q('fldChar'):
                kind=e.get(q('fldCharType'))
                if kind=='begin':stack.append({'instructions':[],'result':[],'separated':False})
                elif kind=='separate' and stack:stack[-1]['separated']=True
                elif kind=='end' and stack:
                    f=stack.pop();code=''.join(x.text or '' for x in f['instructions']);match=re.match(r'\s*REF\s+(\S+)',code,re.I)
                    if not match or match.group(1) in bookmarks:continue
                    visible=''.join(f['result']).strip();number=re.fullmatch(r'\[?(\d+)\]?',visible)
                    # Require a complete [n] token in this paragraph plus a
                    # unique matching bibliography item and an intact cache.
                    if not number or '['+number.group(1)+']' not in text(p) or len(refs.get(number.group(1),[]))!=1:
                        issues.append(issue('broken_ref_ambiguous',match.group(1),'失效 REF 无唯一现有书目证据，不能猜接。'));continue
                    target=refs[number.group(1)][0];full_label=visible.startswith('[');name='FITv10Reference'+number.group(1)+('Full' if full_label else 'Number')
                    if name not in bookmarks:
                        begin=E.Element(q('bookmarkStart'));begin.set(q('id'),str(next_id));begin.set(q('name'),name)
                        end=E.Element(q('bookmarkEnd'));end.set(q('id'),str(next_id));next_id+=1
                        # Anchor only the visible numeric label, not the entire
                        # bibliographic entry, so REF still displays its number.
                        tnodes=target.findall('.//w:t',NS);first=tnodes[0] if tnodes else None
                        if first is None or first.getparent().getparent() is not target or not re.match(r'^\s*\['+re.escape(number.group(1))+r'\]',first.text or ''):
                            issues.append(issue('reference_label_complex',name,'书目编号跨多个对象，需专门定位。'));continue
                        run=first.getparent();label_match=re.match(r'^(\s*\[)('+re.escape(number.group(1))+r')(\])',first.text)
                        lo=label_match.start() if full_label else label_match.start(2);hi=label_match.end() if full_label else label_match.end(2)
                        before=C.sliced_run(run,0,lo);label=C.sliced_run(run,lo,hi);after=C.sliced_run(run,hi,len(C.run_text(run)))
                        pos=target.index(run);target.remove(run)
                        for i,node in enumerate((before,begin,label,end,after)):target.insert(pos+i,node)
                        bookmarks.add(name)
                    # REF copies the reference-list label's formatting unless
                    # CHARFORMAT explicitly anchors it to the instruction run.
                    # Keep citations superscript after native Word refresh as
                    # well as in the original OOXML cache (and in WPS).
                    old=code;new=' REF '+name+' '+chr(92)+'h '+chr(92)+'* CHARFORMAT '
                    f['instructions'][0].text=new
                    setprop(F.ensure(f['instructions'][0].getparent(),'rPr',True),'vertAlign',val='superscript')
                    for extra in f['instructions'][1:]:extra.text=''
                    ops.append({'kind':'repair_reference_target','before':old,'after':new,'reference_number':number.group(1),'bibliography_text':text(target),'cached_result':visible})
            elif e.tag==q('instrText') and stack and not stack[-1]['separated']:stack[-1]['instructions'].append(e)
            elif e.tag==q('t') and stack and stack[-1]['separated']:stack[-1]['result'].append(e.text or '')

def _commitment_blocks(root,roles,ops):
    """Split line-broken pledge blocks into semantic native paragraphs."""
    for p,role in list(roles.items()):
        before=text(p)
        if role!='commitment_body' or not re.match(r'^(学生承诺|指导教师承诺)我承诺',before):continue
        forbidden={q(x) for x in ('fldChar','fldSimple','bookmarkStart','bookmarkEnd','ins','del','commentReference','footnoteReference','endnoteReference')}
        if any(e.tag in forbidden for e in p.iter()):continue
        groups=[];group=[]
        for run in p:
            if run.tag==q('pPr'):continue
            if run.tag!=q('r'):raise ValueError('Unsupported mixed commitment block')
            prop=run.find(q('rPr'));current=E.Element(q('r'))
            if prop is not None:current.append(copy.deepcopy(prop))
            def flush():
                nonlocal current,group
                if any(c.tag!=q('rPr') for c in current):group.append(current)
                current=E.Element(q('r'))
                if prop is not None:current.append(copy.deepcopy(prop))
            for el in run:
                if el.tag==q('rPr'):continue
                if el.tag==q('br'):
                    flush()
                    if group:groups.append(group);group=[]
                else:current.append(copy.deepcopy(el))
            flush()
        if group:groups.append(group)
        parent=p.getparent();at=parent.index(p);newps=[]
        for group in groups:
            np=E.Element(q('p'))
            for r in group:np.append(r)
            value=text(np).strip()
            if not value and not np.findall('.//w:drawing',NS) and not np.findall('.//w:pict',NS):continue
            newrole='commitment_label' if value in ('学生承诺','指导教师承诺') else 'commitment_body' if '我承诺' in value else 'commitment_signature'
            if newrole=='commitment_signature':
                for t in np.findall('.//w:t',NS):t.text=(t.text or '').lstrip(' \u3000')
            F.format_p(np,newrole,{},None,[]);roles[np]=newrole;newps.append(np)
        after=''.join(text(x) for x in newps)
        if re.sub(r'\s+','',before)!=re.sub(r'\s+','',after):raise ValueError('Commitment split changed original wording')
        parent.remove(p);roles.pop(p,None)
        for i,np in enumerate(newps):parent.insert(at+i,np)
        ops.append({'kind':'commitment_split','before':before,'after':after,'paragraphs':[text(x) for x in newps]})


def _front_template():
    parts=F.package(Path(__file__).resolve().parent.parent/'assets/fit-economics-management-template-clean.docx')
    return parts,F.parse(parts[DOC]).find(q('body'))

def _front_pairs(root):
    """Recognize the complete native template scaffold before touching it."""
    template,tb=_front_template();actual=list(root.find(q('body')))[:14];wanted=list(tb)[:14]
    if len(actual)!=14 or [x.tag for x in actual]!=[x.tag for x in wanted]:return None
    if text(actual[2])!='本科生毕业论文（设计）' or '诚信承诺书' not in text(actual[9]):return None
    for n in (6,10):
        rows=actual[n].findall(q('tr'));refrows=wanted[n].findall(q('tr'))
        if len(rows)!=len(refrows):return None
        for row,refrow in zip(rows,refrows):
            cells=row.findall(q('tc'));refcells=refrow.findall(q('tc'))
            if len(cells)!=len(refcells) or [len(c.findall(q('p'))) for c in cells]!=[len(c.findall(q('p'))) for c in refcells]:return None
    # These fixed labels distinguish the front form from arbitrary user tables.
    if [text(r.findall(q('tc'))[0]).strip() for r in actual[6].findall(q('tr'))]!=[text(r.findall(q('tc'))[0]).strip() for r in wanted[6].findall(q('tr'))]:return None
    return template,actual,wanted

def _import_front_styles(styles,template):
    """Isolate template defaults so the thesis Normal style cannot stretch forms."""
    source=F.parse(template['word/styles.xml']);index={x.get(q('styleId')):x for x in source.findall(q('style'))}
    default=next((sid for sid,x in index.items() if x.get(q('type'))=='paragraph' and x.get(q('default'))=='1'),None)
    imported=set()
    def clone(sid):
        target='FITFrontTpl_'+sid
        if sid in imported:return target
        imported.add(sid)
        original=index.get(sid)
        if original is None:raise ValueError('School template references a missing style: '+sid)
        node=copy.deepcopy(original);node.set(q('styleId'),target);node.attrib.pop(q('default'),None)
        F.remove_names(node,{'autoRedefine'})
        name=node.find(q('name'))
        if name is not None:name.set(q('val'),'FIT School Form '+name.get(q('val'),sid))
        for dep in node:
            if dep.tag in (q('basedOn'),q('next'),q('link')) and dep.get(q('val')) in index:dep.set(q('val'),clone(dep.get(q('val'))))
        if sid==default:
            for kind in ('pPr','rPr'):
                defaults=source.find('w:docDefaults/w:'+kind+'Default/w:'+kind,NS)
                dest=F.ensure(node,kind)
                if defaults is not None:
                    for prop in defaults:
                        if dest.find(prop.tag) is None:dest.append(copy.deepcopy(prop))
            # A missing template default means Word's zero paragraph spacing /
            # single line, not permission to inherit the thesis docDefaults.
            # In particular, an after=200 body default adds 10pt to every form
            # paragraph and can spill the pledge's trailing section to a page.
            spacing=F.ensure(F.ensure(node,'pPr'),'spacing')
            for attr,value in (('before','0'),('after','0'),('beforeAutospacing','0'),('afterAutospacing','0'),('line','240'),('lineRule','auto')):
                if spacing.get(q(attr)) is None:spacing.set(q(attr),value)
            ind=F.ensure(F.ensure(node,'pPr'),'ind')
            for attr in ('left','right','firstLine','firstLineChars','leftChars','rightChars','hanging','hangingChars'):
                if ind.get(q(attr)) is None:ind.set(q(attr),'0')
            for attr in ('beforeLines','afterLines'):
                if spacing.get(q(attr)) is None:spacing.set(q(attr),'0')
            for name in ('keepNext','keepLines','pageBreakBefore','contextualSpacing','mirrorIndents'):
                if node.find('w:pPr/w:'+name,NS) is None:F.child(node.find(q('pPr')),name,val=0)
            fonts=F.ensure(F.ensure(node,'rPr'),'rFonts')
            for attr in list(fonts.attrib):
                if 'theme' in attr.lower():del fonts.attrib[attr]
            for attr,value in (('ascii','Times New Roman'),('hAnsi','Times New Roman'),('cs','Times New Roman'),('eastAsia','宋体')):fonts.set(q(attr),value)
        old=styles.find('w:style[@w:styleId="'+target+'"]',NS)
        if old is not None:styles.remove(old)
        styles.append(node)
        return target
    return clone,default

def _replace_property(node,ref,name):
    old=node.find(q(name));position=list(node).index(old) if old is not None else 0
    if old is not None:node.remove(old)
    wanted=ref.find(q(name))
    if wanted is not None:node.insert(position,copy.deepcopy(wanted))

def _front_label_padding(table):
    # Three full-width label characters at 16pt must fit the template's
    # 1071-twip label column. Keep the grid and remove only label-cell padding.
    for row in table.findall(q('tr')):
        cp=F.ensure(row.findall(q('tc'))[0],'tcPr',True)
        F.remove_names(cp,{'tcMar'});m=F.child(cp,'tcMar')
        for side in ('left','right'):F.child(m,side,w=0,type='dxa')

def restore_school_front_geometry(parts,root,styles,roles=None,ops=None,fill_title=False):
    """Restore the recognized school form's properties without rebuilding prose.

    Content, drawings and section boundaries stay in their original nodes. The
    same operation serves the formatter and a bounded repair of an old v10 file.
    """
    pair=_front_pairs(root)
    if pair is None:return {'restored':False,'reason':'front scaffold does not match the school form'}
    template,actual,wanted=pair;clone,default=_import_front_styles(styles,template)
    edits=[];exemplars={}
    for node,refnode in zip(actual,wanted):
        if node.tag==q('tbl'):
            for name in ('tblPr','tblGrid'):_replace_property(node,refnode,name)
            ts=node.find('w:tblPr/w:tblStyle',NS)
            if ts is not None:ts.set(q('val'),clone(ts.get(q('val'))))
            for row,refrow in zip(node.findall(q('tr')),refnode.findall(q('tr'))):
                _replace_property(row,refrow,'trPr')
                for cell,refcell in zip(row.findall(q('tc')),refrow.findall(q('tc'))):_replace_property(cell,refcell,'tcPr')
        ps=[node] if node.tag==q('p') else node.findall('.//w:p',NS)
        rps=[refnode] if refnode.tag==q('p') else refnode.findall('.//w:p',NS)
        for p,rp in zip(ps,rps):
            oldsec=p.find('w:pPr/w:sectPr',NS);oldsec=copy.deepcopy(oldsec) if oldsec is not None else None
            _replace_property(p,rp,'pPr');pp=F.ensure(p,'pPr',True)
            F.remove_names(pp,{'sectPr'})
            sid=pp.find(q('pStyle'));source_sid=sid.get(q('val')) if sid is not None else default
            if sid is None:sid=F.child(pp,'pStyle')
            sid.set(q('val'),clone(source_sid))
            pp.remove(sid);pp.insert(0,sid)
            if oldsec is not None:pp.append(oldsec)
            exemplar=rp.find('w:r/w:rPr',NS)
            if exemplar is None:exemplar=rp.find('w:pPr/w:rPr',NS)
            exemplars[p]=exemplar
            mark=pp.find(q('rPr'))
            if mark is not None:F.remove_names(mark,{'highlight','shd','rStyle'})
            for run in p.findall(q('r')):
                F.remove_names(run,{'rPr'})
                if exemplar is not None:
                    prop=copy.deepcopy(exemplar);F.remove_names(prop,{'highlight','shd','rStyle'});run.insert(0,prop)
            if not text(p).strip(' \u3000') or re.fullmatch(r'\s*年月日\s*',text(p)):
                if re.fullmatch(r'\s*年月日\s*',text(p)):
                    before=text(p)
                    for t in p.findall('.//w:t',NS):t.text=''
                    run=plain_run(p,'                                         年    月    日')
                    if exemplar is not None:
                        prop=copy.deepcopy(exemplar);F.remove_names(prop,{'highlight','shd','rStyle'});run.insert(0,prop)
                    edits.append({'kind':'front_blank_date','before':before,'after':text(p)})
    _front_label_padding(actual[6])
    title=''.join(text(row.findall(q('tc'))[1]).strip() for row in actual[6].findall(q('tr'))[:2])
    if fill_title and title:
        target=actual[10].findall(q('tr'))[2].findall(q('tc'))[1].find(q('p'))
        if text(target).strip()=='中文：':
            before=text(target);run=plain_run(target,title);sample=exemplars.get(target)
            if sample is not None:
                prop=copy.deepcopy(sample);F.remove_names(prop,{'highlight','shd','rStyle'});run.insert(0,prop)
            edits.append({'kind':'front_title_fill','before':before,'after':text(target),'source':'cover_title'})
    # Formatting differences are independently checked against this exact form;
    # any optional text fills are reported as explicit front-matter edits.
    record={'restored':True,'template_sha256':F.sha(template[DOC]),'edits':edits}
    if ops is not None:ops.append({'kind':'school_front_geometry',**record})
    return record

def front_template_geometry_issues(root,styles=None):
    """Lightweight checks for Audit: widths, row/cell grids and native spacing."""
    pair=_front_pairs(root)
    if pair is None:return ['school form scaffold differs from the template']
    _,actual,wanted=pair;errors=[]
    wanted=[copy.deepcopy(n) for n in wanted]
    _front_label_padding(wanted[6])
    if styles is not None:
        base=styles.find('w:style[@w:styleId="FITFrontTpl_1"]/w:pPr/w:spacing',NS)
        if base is None or any(base.get(q(k))!=v for k,v in (('before','0'),('after','0'),('line','240'),('lineRule','auto'))):
            errors.append('isolated form defaults are missing or inherit thesis paragraph spacing')
    def props(node,names):
        return [(name,F.format_property_signature(node.find(q(name)))) for name in names]
    for n in (6,10):
        for label in ('tblGrid',):
            if props(actual[n],[label])!=props(wanted[n],[label]):errors.append('front table '+str(n)+' '+label+' differs')
        for name in ('tblW','tblInd','tblLayout','tblCellMar'):
            if props(actual[n].find(q('tblPr')),[name])!=props(wanted[n].find(q('tblPr')),[name]):errors.append('front table '+str(n)+' '+name+' differs')
        for ri,(row,refrow) in enumerate(zip(actual[n].findall(q('tr')),wanted[n].findall(q('tr')))):
            if props(row,['trPr'])!=props(refrow,['trPr']):errors.append('front row '+str(n)+':'+str(ri)+' geometry differs')
            for ci,(cell,refcell) in enumerate(zip(row.findall(q('tc')),refrow.findall(q('tc')))):
                if props(cell,['tcPr'])!=props(refcell,['tcPr']):errors.append('front cell '+str(n)+':'+str(ri)+':'+str(ci)+' geometry differs')
    for node,refnode in zip(actual,wanted):
        ps=[node] if node.tag==q('p') else node.findall('.//w:p',NS);refs=[refnode] if refnode.tag==q('p') else refnode.findall('.//w:p',NS)
        for p,ref in zip(ps,refs):
            pp=p.find(q('pPr'));rpp=ref.find(q('pPr'))
            for name in ('spacing','jc','tabs','ind','snapToGrid'):
                a=pp.find(q(name)) if pp is not None else None;b=rpp.find(q(name)) if rpp is not None else None
                if name=='ind':
                    if a is not None and all(v=='0' for v in a.attrib.values()):a=None
                    if b is not None and all(v=='0' for v in b.attrib.values()):b=None
                if F.format_property_signature(a)!=F.format_property_signature(b):errors.append('front paragraph spacing/alignment differs: '+text(p)[:30]);break
    return errors

def _school_components(parts,root,roles,metadata,ops,issues):
    """Clone fixed clean-template components; never copy template identities."""
    body=root.find(q('body'));alltext=text(body)
    missing_cover='本科生毕业论文（设计）' not in alltext
    missing_pledge='诚信承诺书' not in alltext
    if not (missing_cover or missing_pledge):return
    template=F.package(Path(__file__).resolve().parent.parent/'assets/fit-economics-management-template-clean.docx')
    tr=F.parse(template[DOC]);tb=tr.find(q('body'));selected=list(tb)[:9] if missing_cover else []
    if missing_pledge:selected+=list(tb)[9:14]
    nodes=[copy.deepcopy(x) for x in selected]
    cover=next((x for x in nodes if x.tag==q('tbl') and '题目：' in text(x)),None)
    if cover is not None:
        # The second row is the title's spare ruled line, not the name row.
        fields={'题目：':'title','姓名：':'name','学号：':'student_id','学院：':'college','专业：':'major','年级：':'grade','指导教师（签名）：':'supervisor'}
        for row in cover.findall(q('tr')):
            cells=row.findall(q('tc'))
            key=fields.get(text(cells[0]).strip())
            if key is None:continue
            value=metadata.get(key,'')
            for c in cells[1:]:
                for p in c.findall('.//w:p',NS):
                    for child in list(p):
                        if child.tag!=q('pPr'):p.remove(child)
                p=c.find(q('p'))
                if p is None:p=F.child(c,'p')
                plain_run(p,value)
        date=re.fullmatch(r'\s*(\d{4})年(?:\s*(\d{1,2})月)?(?:\s*(\d{1,2})日)?\s*',metadata.get('date',''))
        if date:
            cells=cover.findall(q('tr'))[-1].findall(q('tc'))
            for i,value in zip((1,3,5),date.groups()):
                if i<len(cells):
                    p=cells[i].find(q('p'))
                    for run in list(p):
                        if run.tag!=q('pPr'):p.remove(run)
                    plain_run(p,value or '')
    pledge=next((x for x in nodes if x.tag==q('tbl') and '学生承诺' in text(x)),None)
    if pledge is not None:
        # Template's fill-in cells already blank except college and signature
        # dates; remove these values without touching fixed declaration prose.
        for p in pledge.findall('.//w:p',NS):
            s=text(p)
            if '2025' in s and '年' in s:
                for t in p.findall('.//w:t',NS):
                    t.text=re.sub(r'2025|11|3','',t.text or '')
            if s.strip()=='经济管理学院':
                for t in p.findall('.//w:t',NS):t.text=''
                plain_run(p,metadata.get('college',''))
        # Values are placed only next to exact label cells, not guessed dates.
        label_keys={'学生姓名':'name','年级':'grade','学号':'student_id','所在学院':'college','所学专业':'major'}
        for row in pledge.findall(q('tr')):
            cells=row.findall(q('tc'))
            for i,c in enumerate(cells[:-1]):
                key=label_keys.get(re.sub(r'\s+','',text(c)))
                if key:
                    target=cells[i+1]
                    if not text(target).strip():plain_run(target.find(q('p')),metadata.get(key,''))
        cn=pledge.findall(q('tr'))[2].findall(q('tc'))[1].find(q('p'))
        if text(cn).strip()=='中文：' and metadata.get('title'):plain_run(cn,metadata['title'])
    rels=F.parse(parts[L.RELS]);trels=F.parse(template[L.RELS]);ct=F.parse(parts['[Content_Types].xml']);tct=F.parse(template['[Content_Types].xml']);used={x.get('Id') for x in rels};ridmap={};added=[]
    for node in nodes:
        for sec in list(node.findall('.//w:sectPr',NS)):sec.getparent().remove(sec)
        for el in node.iter():
            for attr,value in list(el.attrib.items()):
                if not attr.startswith('{'+R+'}'):continue
                if value not in ridmap:
                    rel=next((r for r in trels if r.get('Id')==value),None)
                    if rel is None or rel.get('TargetMode')=='External':raise ValueError('Unsupported template relationship')
                    oldname='word/'+rel.get('Target');data=template[oldname]
                    newname='word/media/FITCover_'+hashlib.sha256(data).hexdigest()[:16]+Path(oldname).suffix
                    parts[newname]=data;added.append(newname)
                    rid='rIdFITCover'+str(len(ridmap)+1)
                    while rid in used:rid+='x'
                    used.add(rid);ridmap[value]=rid
                    E.SubElement(rels,'{'+PKG+'}Relationship',Id=rid,Type=rel.get('Type'),Target=newname[5:])
                    ext=Path(newname).suffix[1:]
                    if not any(x.get('Extension')==ext for x in ct):
                        src=next((x for x in tct if x.get('Extension')==ext),None)
                        if src is None:raise ValueError('Missing template media content type')
                        ct.append(copy.deepcopy(src))
                el.set(attr,ridmap[value])
    insert=0
    if not missing_cover:
        insert=next((i for i,p in enumerate(body) if roles.get(p) in ('cn_title','cn_abstract_heading')),0)
    for i,node in enumerate(nodes):body.insert(insert+i,node)
    # Assign explicit styles to new components, independent of template styles.
    for node in nodes:
        for p in ([node] if node.tag==q('p') else node.findall('.//w:p',NS)):
            s=text(p);table=next((a for a in p.iterancestors() if a.tag==q('tbl')),None)
            if s=='本科生毕业论文（设计）':role='cover_banner'
            elif s=='（正本）':role='cover_copy'
            elif '诚信承诺书' in s:role='commitment_heading'
            elif table is cover:role='cover_title' if s==metadata.get('title') and s else 'cover_info'
            elif s.strip() in ('学生承诺','指导教师承诺'):role='commitment_label'
            elif '我承诺' in s:role='commitment_body'
            elif table is pledge:
                role='commitment_signature' if '签名' in s or re.fullmatch(r'(?:\d{4})?年(?:\d{1,2})?月(?:\d{1,2})?日',re.sub(r'\s+','',s)) else 'commitment_info'
            else:role='cover_info'
            roles[p]=role
            F.format_p(p,role,{},None,[])
    parts[L.RELS]=F.xml(rels);parts['[Content_Types].xml']=F.xml(ct)
    ops.append({'kind':'school_components','components':(['cover'] if missing_cover else [])+(['commitment'] if missing_pledge else []),'added_assets':added,'relationship_mapping':ridmap,'added_texts':[text(x) for x in nodes]})
    absent=[k for k in ('title','name','student_id','college','major','grade','date') if not metadata.get(k)]
    if absent:issues.append(issue('missing_identity_fields','cover','原稿未能确定，留空且未伪造：'+','.join(absent),'warning'))

def format_document(source,output,decisions=None,profile=None):
    source=Path(source).resolve();output=Path(output).resolve()
    if output==source or output.exists():raise ValueError('Output must be a new file, never the original')
    inspected=inspect_document(source);mapping=copy.deepcopy(inspected['map']);decisions=decisions or {};profile=profile or {}
    if isinstance(decisions,(str,Path)):decisions=F.read_json(decisions)
    if decisions and decisions.get('source_sha256')!=inspected['source_sha256']:raise ValueError('Stale decisions source_sha256')
    for key in ('paragraphs','story_paragraphs'):mapping[key].update(decisions.get(key,{}))
    for tid,item in decisions.get('tables',{}).items():
        if tid not in mapping['tables'] or not isinstance(item,dict):raise ValueError('Invalid reviewed table target')
        mapping['tables'][tid].update(item)
    parts=F.package(source);original_text={k:F.ptext(p) for k,p in F.indexed(F.parse(parts[DOC]))[0].items()};narrative_ops=[]
    if decisions.get('narrative_map'):
        parts=N.expected_parts(source,decisions['narrative_map'])
        for k,p in F.indexed(F.parse(parts[DOC]))[0].items():
            if original_text[k]!=F.ptext(p):narrative_ops.append({'kind':'narrative_citation','locator':k,'before':original_text[k],'after':F.ptext(p)})
    root=F.parse(parts[DOC]);styles=F.parse(parts['word/styles.xml']);ps,ts=F.indexed(root)
    roles={ps[k]:v if isinstance(v,str) else v['role'] for k,v in mapping['paragraphs'].items()}
    unconfirmed_heading_nodes={ps[c['id']] for c in inspected['candidates'] if c.get('role') in LEVELS and c['id'] not in decisions.get('paragraphs',{})}
    for p in unconfirmed_heading_nodes:roles.pop(p,None)
    bound_repairs=_bind_layout_repairs(root,styles,ps,ts,decisions)
    continuation_plans=_validate_continuations(ts,mapping['tables'])
    caption_targets=_bind_caption_targets(ps,ts,roles,decisions.get('caption_targets',[]))
    table_options_by_node={ts[tid]:dict(opts) for tid,opts in mapping['tables'].items()}
    appendix_nodes=set();in_appendix=False
    for block in root.find(q('body')):
        if roles.get(block)=='appendix_heading':in_appendix=True
        elif roles.get(block) in ('chapter','reference_heading','acknowledgement_heading','conclusion_heading'):in_appendix=False
        if in_appendix:appendix_nodes.update(block.iter(q('p')))
    code_nodes={p for p,r in roles.items() if r=='appendix_code'}
    unresolved=[i for i in inspected['issues'] if i['location'] not in decisions.get('paragraphs',{})]
    for pid,item in decisions.get('preserve',{}).items():
        p=ps.get(pid)
        if p is None or not any(a.tag in F.PROTECTED_TAGS for a in p.iterancestors()):raise ValueError('Preserve decision is only for protected nested objects')
        if not isinstance(item,dict) or len(item.get('reason',''))<6 or item.get('expected_text')!=original_text[pid]:raise ValueError('Invalid source-bound preserve decision')
        unresolved=[i for i in unresolved if i['location']!=pid]
    operations=list(narrative_ops);style_index={s.get(q('styleId')):copy.deepcopy(s) for s in styles.findall(q('style'))}
    transcriptions=decisions.get('formula_transcriptions',{})
    if not isinstance(transcriptions,dict):raise ValueError('formula_transcriptions must be a paragraph map')
    for pid,item in transcriptions.items():
        if pid not in ps or roles.get(ps[pid])!='equation' or not isinstance(item,dict) or set(item)-{'expected_text','omml','reason'} or item.get('expected_text')!=text(ps[pid]) or not isinstance(item.get('omml'),str):raise ValueError('Invalid source-bound formula transcription')
    nr=F.parse(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else None
    F.make_styles(styles);warnings=[]
    for pid,item in mapping['paragraphs'].items():
        p=ps.get(pid);role=item if isinstance(item,str) else item['role']
        if p is None or role not in F.ROLES and role!='appendix_code':raise ValueError('Invalid role decision '+pid)
        if isinstance(item,dict) and 'expected_text' in item and item['expected_text'] not in (text(p),F.ptext(p),original_text[pid]):raise ValueError('Decision text mismatch '+pid)
        if p in unconfirmed_heading_nodes:continue
        if role=='appendix_code':continue  # Source code fonts, indentation and whitespace are protected.
        if role=='equation' and not p.findall('.//m:oMath',NS):
            try:
                record=MT.transcribe(p,transcriptions.get(pid,{}).get('omml'))
                operations.append({'kind':'formula_transcription','locator':pid,**record})
            except (ValueError,E.XMLSyntaxError) as exc:unresolved.append(issue('formula_transcription_review',pid,str(exc)))
        F.format_p(p,role,style_index,nr,warnings,pid)
        if p.findall('.//m:oMath',NS) and role!='equation':setprop(F.ensure(p,'pPr',True),'spacing',before=0,after=0,line=400,lineRule='atLeast')
    width=F.SPEC['page']['width_twips']-F.SPEC['page']['left_twips']-F.SPEC['page']['right_twips']
    for tid,setting in mapping['tables'].items():_table(ts[tid],setting,width)
    if not profile.get('strict_text'):_citation_insertions(root,decisions.get('citation_insertions',[]),operations)
    _numeric_citations(root,styles,roles,operations,unresolved)
    _repair_reference_bookmarks(root,roles,operations,unresolved)
    if not profile.get('strict_text'):_author_displays(parts,root,styles,operations,unresolved,code_nodes)
    # Source paragraph IDs remain valid through all citation edits. Pledge
    # splitting changes the index and must run after those source-bound edits.
    _commitment_blocks(root,roles,operations)
    for name,pmap in mapping.get('story_paragraphs',{}).items():
        if name.startswith(('word/header','word/footer')):continue  # Replaced by explicit furniture operation; retain old objects byte-for-byte.
        sr=F.parse(parts[name]);sps,_=F.indexed(sr)
        for pid,item in pmap.items():F.format_p(sps[pid],item if isinstance(item,str) else item['role'],style_index,nr,warnings,name+':'+pid)
        parts[name]=F.xml(sr)
    F.apply_pages(root,{},warnings)
    if profile.get('strict_text',False) is not True:
        for pid,p in list(ps.items()):
            if p.getparent() is not root.find(q('body')):continue
            if p in code_nodes:continue
            # Only genuine spacing paragraphs: no break, bookmark, note,
            # section, field, drawing, revision or paragraph mark semantics.
            if all(x.tag in {q('p'),q('pPr'),q('r'),q('rPr'),q('t')} or any(a.tag in {q('pPr'),q('rPr')} for a in x.iterancestors()) for x in p.iter()) and not F.ptext(p).strip() and p.find('w:pPr/w:sectPr',NS) is None:
                if roles.get(p,'').startswith('cover'):continue
                root.find(q('body')).remove(p);roles.pop(p,None);operations.append({'kind':'remove_empty_paragraph','locator':pid,'before':F.ptext(p),'after':''})
                continue
            if p in appendix_nodes:continue
            if roles.get(p) in ('body','cn_abstract','en_abstract'):
                texts=p.findall('.//w:t',NS)
                if texts:
                    old=text(p)
                    # Remove only padding at paragraph edges, never interior
                    # English/numeric/code spaces or mathematical whitespace.
                    texts[0].text=(texts[0].text or '').lstrip(' \u3000');texts[-1].text=(texts[-1].text or '').rstrip(' \u3000')
                    if text(p)!=old:operations.append({'kind':'paragraph_padding','locator':pid,'before':old,'after':text(p)})
        _numbering(parts,root,roles,operations);_appendix(root,roles,operations);_trim_terminal_page_breaks(root,roles,operations);_toc(root,roles,operations,unresolved)
        for p,role in list(roles.items()):
            if role in ('table_caption','figure_caption'):_caption_fields(p,operations)
        metadata=dict(decisions.get('metadata',{}))
        if not metadata.get('title'):
            titles=[text(p).strip() for p,r in roles.items() if r=='cn_title' and text(p).strip()]
            if len(titles)==1:metadata['title']=titles[0]
        _school_components(parts,root,roles,metadata,operations,unresolved)
        restore_school_front_geometry(parts,root,styles,roles,operations)
        for cover in root.iter(q('tbl')):
            if table_kind(cover)['kind']=='cover':_table(cover,{'kind':'cover','header_rows':0},width)
    chapter=0;equation=0
    for p in root.find(q('body')):
        role=roles.get(p)
        if role=='chapter':chapter+=1;equation=0
        if role=='equation' and not profile.get('strict_text'):
            equation+=1;_equation(p,width,operations,unresolved,(chapter,equation) if chapter else None)
    # Equation-layout tables are migrated only when all non-math content is a
    # single recognizable label. Original mathematical nodes move intact.
    for tid,setting in list(mapping['tables'].items()):
        t=ts[tid]
        if setting['kind']!='formula-layout' or t.getparent() is not root.find(q('body')):continue
        if not LABEL.match(text(t).strip()):unresolved.append(issue('formula_table_ambiguous',tid,'公式布局表含有其他文字或无编号，需确认容器迁移。'));continue
        p=E.Element(q('p'));F.child(F.child(p,'pPr'),'pStyle',val='FIT_equation')
        for oldp in t.findall('.//w:p',NS):
            for child in list(oldp):
                if child.tag!=q('pPr'):oldp.remove(child);p.append(child)
            roles.pop(oldp,None)
        t.addprevious(p);t.getparent().remove(t);roles[p]='equation';_equation(p,width,operations,unresolved);operations.append({'kind':'equation_container','locator':tid,'before':'table','after':'paragraph'})
    if not profile.get('strict_text'):_equation_crossrefs(root,roles,operations,unresolved)
    _position_captions(root,roles,ps,ts,caption_targets,operations)
    _fit_pictures(root,width,operations)
    _split_continuations(continuation_plans,roles,table_options_by_node,width,operations)
    _protect_code_styles(styles,code_nodes)
    _apply_layout_repairs(root,styles,roles,bound_repairs,operations)
    for t in root.findall('.//w:t',NS):
        if t.text=='':t.text=None
    citation_style=styles.find('w:style[@w:styleId="FIT_citation_superscript"]',NS)
    if citation_style is not None:styles.remove(citation_style);styles.append(citation_style)
    parts[DOC]=F.xml(root);parts['word/styles.xml']=F.xml(styles)
    # Use existing furniture transformer and exact original title locator.
    body=root.find(q('body'))
    for p,role in list(roles.items()):
        if role not in ('cn_abstract_heading','chapter') or p.getparent() is not body:continue
        if role=='chapter' and p is not next((x for x in body if roles.get(x)=='chapter'),None):continue
        prev=p.getprevious()
        if prev is not None and prev.tag!=q('p'):
            carrier=E.Element(q('p'));F.child(F.child(carrier,'pPr'),'spacing',before=0,after=0,line=20,lineRule='exact');p.addprevious(carrier)
            operations.append({'kind':'section_carrier','before':'','after':'','boundary':text(p)})
    parts[DOC]=F.xml(root)
    newps,_=F.indexed(root);pids={p:k for k,p in newps.items()}
    title=next((p for p,r in roles.items() if r=='cn_title' and text(p).strip()),None)
    if title is None:title=next((p for p,r in roles.items() if r=='cover_title' and text(p).strip()),None)
    front=next((p for p,r in roles.items() if r=='cn_title' and p.getparent() is body),None)
    if front is None:front=next((p for p,r in roles.items() if r=='cn_abstract_heading' and p.getparent() is body),None)
    first=next((p for p,r in roles.items() if r=='chapter' and p.getparent() is body),None)
    if title is not None and first is not None:
        regions=[{'kind':'unpaged' if front is not None and body.index(front)>0 else 'front' if front is not None else 'body','start':'start'}]
        if front is not None and body.index(front)>0:regions.append({'kind':'front','start':pids[front],'expected_text':F.ptext(front)})
        if body.index(first)>0:regions.append({'kind':'body','start':pids[first],'expected_text':F.ptext(first)})
        try:
            parts,detail=L.transform(parts,{'title':{'paragraph_id':pids[title],'expected_text':text(title)},'regions':regions})
            operations.append({'kind':'running_furniture','details':detail})
        except ValueError as exc:unresolved.append(issue('section_layout_review','document',str(exc)))
    else:unresolved.append(issue('title_or_body_boundary_missing','document','需要确认完整中文题目及正文起点，才能建立正确页眉页码。'))
    if not any('本科生毕业论文' in text(p) for p in root.findall('.//w:p',NS)):
        unresolved.append(issue('cover_missing','document','缺学校封面。须提供封面元数据或确认留空后迁入净化模板。'))
    if not any('诚信承诺书' in text(p) for p in root.findall('.//w:p',NS)):
        unresolved.append(issue('commitment_missing','document','缺学校诚信承诺书，需迁入净化模板并核对个人信息。'))
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in parts.items():
            zi=zipfile.ZipInfo(name,date_time=(2000,1,1,0,0,0));zi.compress_type=zipfile.ZIP_DEFLATED;z.writestr(zi,data)
    final_table_options=[table_options_by_node.get(t,table_kind(t)) for t in root.iter(q('tbl'))]
    after=inspect_document(output)
    for tid,opts in zip(after['map']['tables'],final_table_options):after['map']['tables'][tid]=opts
    # Re-index classifications after moves and deletions by role styles. Table
    # roles are retained by structure, not stale paragraph numbers.
    final_map=after['map'];final_map['reviewed']=True
    operations.insert(0,{'kind':'format_roles','paragraph_count':len(mapping['paragraphs']),'tables':mapping['tables'],'source_mapping':mapping})
    return {'schema_version':1,'source':str(source),'output':str(output),'source_sha256':inspected['source_sha256'],
            'output_sha256':F.sha(output.read_bytes()),'operations':operations,'source_mapping':mapping,'decisions':decisions,'profile':profile,'final_mapping':final_map,
            'unresolved_issues':unresolved,'warnings':warnings,'delivery_ready':False,
            'pending':['controlled_field_refresh','Word_and_WPS_render','all_pages_review']}

def verify_format(source,final,decisions=None,profile=None):
    import tempfile
    with tempfile.TemporaryDirectory(prefix='fit-replay-') as td:
        report=format_document(source,Path(td)/'expected.docx',decisions,profile)
        expected=F.package(Path(td)/'expected.docx');actual=F.package(final)
        errors=[n for n in set(expected)|set(actual) if expected.get(n)!=actual.get(n)]
    return {'ok':not errors,'errors':errors,'expected_sha256':report['output_sha256'],'actual_sha256':F.sha(Path(final).read_bytes())}

if __name__=='__main__':
    import argparse
    ap=argparse.ArgumentParser();ap.add_argument('source');ap.add_argument('--output');ap.add_argument('--decisions');ap.add_argument('--report');a=ap.parse_args()
    result=format_document(a.source,a.output,F.read_json(a.decisions) if a.decisions else None) if a.output else inspect_document(a.source)
    if a.report:F.write_json(a.report,result)
    else:print(json.dumps(result,ensure_ascii=False,indent=2))
