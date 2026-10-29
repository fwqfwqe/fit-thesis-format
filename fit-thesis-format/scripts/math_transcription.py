"""Bounded source-identity conversion of plain formula paragraphs to OMML.

Simple literal equations may be wrapped without inferring algebra. Structured
transcriptions must spell the exact original expression when linearized. Images,
OLE, revisions and unknown structures are not guessed or silently discarded.
"""
import copy
import hashlib
import re
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
NS={'w':W,'m':M}
def q(n):return '{'+W+'}'+n
def m(n):return '{'+M+'}'+n
LABEL=re.compile(r'(?:公式\s*)?[（(]\s*\d+\s*[-.－—]\s*\d+\s*[）)]\s*$')

def canonical(value):
    value=re.sub(r'\s+','',value)
    # Parentheses around an atom do not change grouping; compound groups stay.
    return re.sub(r'\(([A-Za-z0-9_\u0370-\u03ff]+)\)',r'\1',value)

def linear(node):
    if node.tag in (m('rPr'),m('ctrlPr')):return ''
    if node.tag==m('t'):return node.text or ''
    local=E.QName(node).localname
    if node.tag.startswith('{'+M+'}') and local.endswith('Pr'):return ''
    if local in ('oMath','r','e','num','den','sub','sup','deg') and node.tag.startswith('{'+M+'}'):
        return ''.join(linear(c) for c in node)
    def part(name):
        el=node.find(m(name))
        if el is None:raise ValueError('Incomplete mathematical structure: '+local)
        return linear(el)
    if node.tag==m('f'):return '('+part('num')+')/('+part('den')+')'
    if node.tag==m('sSub'):return part('e')+'_('+part('sub')+')'
    if node.tag==m('sSup'):return part('e')+'^('+part('sup')+')'
    if node.tag==m('sSubSup'):return part('e')+'_('+part('sub')+')^('+part('sup')+')'
    if node.tag==m('d'):
        pr=node.find(m('dPr'));beg='(';end=')'
        if pr is not None:
            for name,default in (('begChr','('),('endChr',')')):
                el=pr.find(m(name));value=el.get(m('val'),default) if el is not None else default
                if name=='begChr':beg=value
                else:end=value
        es=node.findall(m('e'))
        if len(es)!=1:raise ValueError('Multiple delimited expressions require separate review')
        return beg+linear(es[0])+end
    raise ValueError('Unsupported transcription structure: '+str(node.tag))

def parse_math(xml):
    node=E.fromstring(xml.encode('utf-8'),E.XMLParser(resolve_entities=False,no_network=True))
    if node.tag!=m('oMath'):raise ValueError('Transcription must be exactly one m:oMath')
    if any(not el.tag.startswith('{'+M+'}') or el.tag==m('rPr') for el in node.iter()):
        raise ValueError('Foreign nodes or math property overrides not allowed in transcription')
    linear(node) # Validate all structures, not only extracted text.
    return node

def source_expression(p):
    if p.findall('.//m:oMath',NS):raise ValueError('Existing editable math must be preserved')
    for el in p.iter():
        if el.tag in (q('p'),q('pPr'),q('r'),q('rPr'),q('t'),q('tab')):continue
        if any(a.tag in (q('pPr'),q('rPr')) for a in el.iterancestors()):continue
        raise ValueError('Formula paragraph contains fields, anchors or protected objects')
    for prop in p.findall('.//w:rPr/w:vertAlign',NS):
        if prop.get(q('val'),'baseline')!='baseline':
            raise ValueError('Plain source contains formatted scripts; review a structural transcription')
    before=''.join(p.xpath('.//w:t/text()',namespaces=NS))
    match=LABEL.search(before)
    if not match:raise ValueError('Plain formula needs an unambiguous source label')
    expression=before[:match.start()].strip(' \t\u3000')
    if not expression:raise ValueError('Plain formula has no expression')
    return before,expression,match.group()

def verified_record(op,p):
    try:
        before,expression,label=source_expression(p)
        if op.get('before')!=before or op.get('source_expression')!=expression:return False
        node=parse_math(op['omml'])
        if canonical(linear(node))!=canonical(expression):return False
        if op.get('after')!=label:return False
        return hashlib.sha256(E.tostring(node)).hexdigest()==op.get('omml_sha256')
    except (ValueError,KeyError,E.XMLSyntaxError):return False

def transcribe(p,supplied_omml=None):
    before,expression,label=source_expression(p)
    if supplied_omml is None:
        if '=' not in expression or re.search(r'[_^/\\\u2211\u221a]',expression):
            raise ValueError('Nonlinear notation needs a source-verified structured OMML transcription')
        node=E.Element(m('oMath'),nsmap={'m':M});r=E.SubElement(node,m('r'));t=E.SubElement(r,m('t'))
        t.text=expression;t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    else:
        node=parse_math(supplied_omml)
        if canonical(linear(node))!=canonical(expression):raise ValueError('Transcription changes source symbols or grouping')
    record={'before':before,'after':label,'source_expression':expression,
            'omml':E.tostring(node,encoding='unicode'),'omml_sha256':hashlib.sha256(E.tostring(node)).hexdigest()}
    for child in list(p):
        if child.tag!=q('pPr'):p.remove(child)
    p.append(copy.deepcopy(node));run=E.SubElement(p,q('r'));text=E.SubElement(run,q('t'));text.text=label
    return record
