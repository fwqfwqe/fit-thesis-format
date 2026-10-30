"""Deterministic presentation labels for CONFIRMED captions/math paragraphs.

No expressions or research quantities are edited. Every mention change is
derived from a unique, source-bound object label and logged for validation.
"""
from __future__ import annotations
import re
from collections import defaultdict
import thesis_format as F

CAP=re.compile(r'^(表|图|Figure)\s*(\d+(?:[-－—.]\d+)?)(?=\s|[\u3000\u4e00-\u9fff]|$)',re.I)
EQ=re.compile(r'^\s*(?:公式\s*)?[（(]\s*(\d+(?:[-－—.]\d+)?)\s*[）)]\s*$')
def text(p):return ''.join(p.xpath('.//w:t/text()',namespaces=F.NS))

def plan(root,roles):
    chapter=0;counts=defaultdict(int);targets=[]
    for p in root.find(F.q('body')):
        role=roles.get(p)
        if role=='chapter':chapter+=1;counts.clear()
        if not chapter:continue
        if role in ('table_caption','figure_caption'):
            match=CAP.match(text(p))
            if not match or p.findall('.//w:fldChar',F.NS):continue
            kind='表' if role=='table_caption' else '图';counts[kind]+=1
            targets.append((p,match.group(),kind+f'{chapter}-{counts[kind]}',kind,match.group(2)))
        elif role=='equation' and p.findall('.//m:oMath',F.NS):
            counts['公式']+=1;match=EQ.fullmatch(text(p))
            if match and not p.findall('.//w:fldChar',F.NS):
                targets.append((p,match.group(),f'公式（{chapter}-{counts["公式"]}）','公式',match.group(1)))
    return targets

def apply(root,roles,replace_span,ops,issues,source_ids=None):
    ids=source_ids or {p:k for k,p in F.indexed(root)[0].items()}
    targets=plan(root,roles);aliases=defaultdict(list)
    for p,old,new,kind,number in targets:
        aliases[(kind,number.replace('－','-').replace('—','-').replace('.','-'))].append((p,old,new))
    edits=defaultdict(list)
    for p,old,new,kind,number in targets:
        if old!=new:edits[p].append((0,len(old),new,ids.get(p)))
    # Full native fields and revisions are left for a source-bound host repair.
    capmention=re.compile(r'(表|图|Figure)\s*(\d+(?:[-－—.]\d+)?)(?![\d.－—-])',re.I)
    eqmention=re.compile(r'(公式|式)\s*[（(]\s*(\d+(?:[-－—.]\d+)?)\s*[）)]')
    for p,role in roles.items():
        if role not in ('body','cn_abstract','en_abstract','source'):continue
        if p.findall('.//w:fldChar',F.NS) or p.findall('.//w:ins',F.NS) or p.findall('.//w:del',F.NS):continue
        value=text(p)
        for pattern in (capmention,eqmention):
            for m in pattern.finditer(value):
                kind='公式' if pattern is eqmention else '表' if m.group(1)=='表' else '图'
                number=m.group(2).replace('－','-').replace('—','-').replace('.','-')
                matches=aliases.get((kind,number),[])
                if len(matches)!=1:continue
                target,old,new=matches[0]
                replacement=new if kind!='公式' or m.group(1)=='公式' else new.removeprefix('公')
                if replacement!=m.group():edits[p].append((m.start(),m.end(),replacement,ids.get(target)))
    for p,changes in edits.items():
        before=text(p);evidence=[]
        for start,end,replacement,target in sorted(changes,reverse=True):
            from lxml import etree
            temp=etree.Element(F.q('p'));r=F.child(temp,'r');F.child(r,'t').text=replacement
            try:replace_span(p,start,end,list(temp))
            except ValueError as exc:
                issues.append({'code':'academic_label_mixed_run','severity':'error','location':ids.get(p),'message':str(exc)});continue
            evidence.append({'start':start,'end':end,'expected_text':before[start:end], 'replacement':replacement,'target_id':target})
        if text(p)!=before:ops.append({'kind':'academic_label','locator':ids.get(p),'before':before,'after':text(p),'edits':evidence})

def verified(op,source_root,mapping):
    """Independently rederive allowed replacements from confirmed SOURCE roles."""
    ps,_=F.indexed(source_root);roles={ps[k]:(v if isinstance(v,str) else v['role']) for k,v in mapping.get('paragraphs',{}).items() if k in ps}
    targets={next(k for k,v in ps.items() if v is p):(old,new,kind,num) for p,old,new,kind,num in plan(source_root,roles)}
    pid=op.get('locator');p=ps.get(pid)
    if p is None or text(p)!=op.get('before'):return False
    value=op['before'];edits=op.get('edits',[]);last=len(value)+1
    for e in sorted(edits,key=lambda x:x['start'],reverse=True):
        start,end=e['start'],e['end'];target=targets.get(e.get('target_id'))
        if target is None or not 0<=start<=end<last or value[start:end]!=e.get('expected_text'):return False
        old,new,kind,number=target;replacement=e['replacement']
        allowed={new,new.removeprefix('公')} if kind=='公式' else {new}
        if replacement not in allowed:return False
        matched=EQ.fullmatch(e['expected_text']) or re.fullmatch(r'(?:式)\s*[（(]\s*'+re.escape(number)+r'\s*[）)]',e['expected_text']) if kind=='公式' else CAP.fullmatch(e['expected_text'])
        if not matched:return False
        if pid!=e['target_id']:
            nums=re.findall(r'\d+(?:[-－—.]\d+)?',e['expected_text'])
            if nums!=[number]:return False
        value=value[:start]+replacement+value[end:];last=start
    return bool(edits) and value==op.get('after')
