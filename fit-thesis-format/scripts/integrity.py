"""Independent DOCX preservation checks for deterministic edits and office refresh.

No baseline replacement and no caller supplied ignore-list.  Layout edits must
replay exactly; native field refresh may change only approved field results and
non-semantic serialization.  A successful comparison is not visual acceptance.
"""
from __future__ import annotations
import collections
import hashlib
import json
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'm': M, 'r': R}
SAFE_FIELDS = {'TOC', 'SEQ', 'REF', 'PAGEREF', 'STYLEREF', 'PAGE', 'NUMPAGES', 'SECTIONPAGES'}
SEMANTIC = ('b', 'i', 'strike', 'dstrike', 'vertAlign', 'vanish', 'webHidden', 'caps', 'smallCaps')

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def read_package(path):
    with zipfile.ZipFile(path) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise ValueError('Duplicate package entries')
        return {n: z.read(n) for n in z.namelist()}

def parse(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))

def q(name):
    return '{' + W + '}' + name

def tree(node, omit=frozenset()):
    if node.tag in omit:
        return None
    attrs = tuple(sorted((k, v) for k, v in node.attrib.items()
                         if not k.startswith('{http://schemas.microsoft.com/office/word/2010/wordml}')))
    children = []
    for c in node:
        if not isinstance(c.tag, str) or c.tag in omit:
            continue
        value = tree(c, omit)
        # Word/WPS insert empty math property containers when saving.  Empty
        # fPr/sSubPr etc. carry no expression semantics; nonempty ones still do.
        if c.tag.startswith('{' + M + '}') and E.QName(c).localname.endswith('Pr') and value and not any(value[1:]):
            continue
        children.append(value)
    return (node.tag, attrs, node.text if node.text and node.text.strip() else '', tuple(children))

def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()

def relationships(parts, name):
    rp = posixpath.join(posixpath.dirname(name), '_rels', posixpath.basename(name) + '.rels')
    if rp not in parts:
        return {}
    result = {}
    for x in parse(parts[rp]):
        target = x.get('Target', '')
        if x.get('TargetMode') != 'External':
            target = posixpath.normpath(posixpath.join(posixpath.dirname(name), target))
        result[x.get('Id')] = (x.get('Type', ''), target, x.get('TargetMode', 'Internal'))
    return result

def styles(parts):
    if 'word/styles.xml' not in parts:
        return {}, {}
    root = parse(parts['word/styles.xml'])
    entries = {x.get(q('styleId')): x for x in root.findall(q('style'))}
    defaults = root.find('w:docDefaults/w:rPrDefault/w:rPr', NS)
    return entries, props(defaults)

def props(rpr):
    result = {}
    if rpr is None:
        return result
    for name in SEMANTIC:
        node = rpr.find(q(name))
        if node is not None:
            val = node.get(q('val'), '1')
            result[name] = val if name == 'vertAlign' else val not in ('0', 'false', 'off')
    color = rpr.find(q('color'))
    if color is not None:
        val = color.get(q('val'), 'auto').lower()
        result['color'] = '000000' if val == 'auto' else val
    return result

def effective(run, entries, defaults):
    out = dict(defaults)
    def inherit(style_id, seen=None):
        seen = seen or set()
        if not style_id or style_id in seen or style_id not in entries:
            return
        seen.add(style_id)
        st = entries[style_id]
        base = st.find(q('basedOn'))
        inherit(base.get(q('val')) if base is not None else None, seen)
        out.update(props(st.find(q('rPr'))))
    p = next((a for a in run.iterancestors() if a.tag == q('p')), None)
    ps = p.find('w:pPr/w:pStyle', NS) if p is not None else None
    inherit(ps.get(q('val')) if ps is not None else 'Normal')
    rs = run.find('w:rPr/w:rStyle', NS)
    inherit(rs.get(q('val')) if rs is not None else None)
    out.update(props(run.find(q('rPr'))))
    return tuple((k, out.get(k, 'baseline' if k == 'vertAlign' else '000000' if k == 'color' else False))
                 for k in (*SEMANTIC, 'color'))

def story_tokens(root, parts, name):
    """Flatten runs without losing visible semantics; suppress approved caches only."""
    entries, defaults = styles(parts)
    rels = relationships(parts, name)
    fields, stack, result = [], [], []
    def suppressed():
        return any(f['result'] and f['safe'] for f in stack)
    def emit(kind, value, semantics=()):
        if suppressed():
            return
        item = (kind, value, semantics)
        if kind == 'text' and result and result[-1][0] == kind and result[-1][2] == semantics:
            old = result.pop()
            result.append((kind, old[1] + value, semantics))
        else:
            result.append(item)
    def walk(node):
        if node.tag == q('fldSimple'):
            code = ' '.join(node.get(q('instr'), '').split())
            safe = code.split(' ', 1)[0].upper() in SAFE_FIELDS if code else False
            if not suppressed():
                fields.append(code)
                if safe:
                    return
        if node.tag == q('fldChar'):
            kind = node.get(q('fldCharType'))
            if kind == 'begin':
                stack.append({'code': '', 'result': False, 'safe': False, 'parent_suppressed': suppressed()})
            elif kind == 'separate' and stack:
                f = stack[-1]
                code = ' '.join(f['code'].split())
                f['safe'] = code.split(' ', 1)[0].upper() in SAFE_FIELDS if code else False
                f['result'] = True
                if not f['parent_suppressed']:
                    fields.append(code)
            elif kind == 'end' and stack:
                stack.pop()
            return
        if node.tag == q('instrText') and stack:
            stack[-1]['code'] += node.text or ''
            return
        if node.tag in (q('rPr'), q('pPr')):
            return
        if node.tag == q('hyperlink'):
            emit('hyperlink', (rels.get(node.get('{' + R + '}id')), node.get(q('anchor'))))
        if node.tag in (q('t'), q('delText')):
            run = next((a for a in node.iterancestors() if a.tag == q('r')), None)
            emit('text', node.text or '', effective(run, entries, defaults) if run is not None else ())
        elif node.tag in (q('tab'), q('br'), q('cr'), q('sym'), q('footnoteReference'), q('endnoteReference')):
            emit('control', (node.tag, tuple(sorted(node.attrib.items()))))
        elif node.tag == '{' + M + '}oMath':
            emit('math', tree(node, frozenset((q('rPr'), '{' + M + '}rPr', '{' + M + '}ctrlPr'))))
            return
        for child in node:
            walk(child)
        if node.tag in (q('p'), q('tc'), q('tr')):
            emit('boundary', node.tag)
    walk(root)
    if stack:
        raise ValueError('Unbalanced field in ' + name)
    return {'tokens': result, 'fields': fields}

def protected(parts):
    """Objects that formatting never has authority to silently remove or change."""
    math, revisions, notes = [], [], {}
    for name, data in parts.items():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        root = parse(data)
        math.extend(digest(tree(x, frozenset((q('rPr'), '{' + M + '}rPr', '{' + M + '}ctrlPr'))))
                    for x in root.iter('{' + M + '}oMath'))
        for x in root.iter():
            if x.tag in {q('ins'), q('del'), q('moveFrom'), q('moveTo')}:
                revisions.append((E.QName(x).localname, x.get(q('author')), x.get(q('date')),
                                  ''.join(x.itertext())))
        if re.fullmatch(r'word/(?:footnotes|endnotes|comments)\.xml', name):
            notes[name] = [((E.QName(x).localname, x.get(q('id'))),
                            ''.join(x.xpath('.//w:t/text()|.//w:delText/text()', namespaces=NS)))
                           for x in root if x.get(q('id')) is not None and x.get(q('type')) not in ('separator', 'continuationSeparator', 'continuationNotice')]
            if not notes[name]:
                del notes[name]
    binaries = collections.Counter(hashlib.sha256(d).hexdigest() for n, d in parts.items()
                                   if n.startswith(('word/media/', 'word/embeddings/', 'word/charts/')))
    return {'math': collections.Counter(math), 'revisions': collections.Counter(revisions),
            'notes': notes, 'binaries': binaries}

def effective_stories(parts):
    """Resolve effective section furniture; native saves may rename its files."""
    result = {n: n for n in ('word/document.xml', 'word/comments.xml') if n in parts}
    for n in ('word/footnotes.xml', 'word/endnotes.xml'):
        if n in parts:
            root = parse(parts[n])
            for x in list(root):
                if x.get(q('type')) in ('separator', 'continuationSeparator', 'continuationNotice'):
                    root.remove(x)
            if len(root): result[n] = E.tostring(root)
    root = parse(parts['word/document.xml'])
    rels = relationships(parts, 'word/document.xml')
    settings = parse(parts['word/settings.xml']) if 'word/settings.xml' in parts else E.Element(q('settings'))
    even = settings.find(q('evenAndOddHeaders')) is not None
    inherited = {}
    for i, sect in enumerate(root.findall('.//w:sectPr', NS)):
        for kind in ('header', 'footer'):
            for ref in sect.findall(q(kind + 'Reference')):
                rel = rels.get(ref.get('{' + R + '}id'))
                if rel: inherited[(kind, ref.get(q('type'), 'default'))] = rel[1]
            active = ['default'] + (['even'] if even else []) + (['first'] if sect.find(q('titlePg')) is not None else [])
            for typ in active:
                path = inherited.get((kind, typ))
                if path in parts:
                    result[f'section:{i}:{kind}:{typ}'] = path
    return result

def research_text(parts, mapping):
    """Independent paragraph payloads; layout counters and metadata are separate."""
    root = parse(parts['word/document.xml'])
    paragraph_map = mapping.get('paragraphs', {})
    excluded = {'toc_heading', 'toc1', 'toc2', 'cover_title', 'cover_info', 'cover_banner', 'cover_copy',
                'commitment_heading', 'commitment_label', 'commitment_body', 'commitment_info', 'commitment_signature', 'equation_number'}
    values = []
    for i, p in enumerate(root.iter(q('p')), 1):
        entry = paragraph_map.get(f'p{i:06d}', {})
        role = entry if isinstance(entry, str) else entry.get('role')
        if role in excluded or any(a.tag in (q('drawing'), q('pict'), q('object')) for a in p.iterancestors()):
            continue
        # Only hidden layout-counter results may be dropped. Hidden ordinary
        # body text remains in this comparison and is checked by replay too.
        texts, stack = [], []
        for el in p.iter():
            if el.tag == q('fldChar'):
                typ = el.get(q('fldCharType'))
                if typ == 'begin': stack.append({'code': '', 'result': False})
                elif typ == 'separate' and stack: stack[-1]['result'] = True
                elif typ == 'end' and stack: stack.pop()
            elif el.tag == q('instrText') and stack: stack[-1]['code'] += el.text or ''
            elif el.tag == q('t'):
                in_counter = any(f['result'] and re.search(r'\bSEQ\s+FIT(?:Chapter|Equation|Table|Figure)\b', f['code'], re.I) for f in stack)
                run = el.getparent()
                hidden = run.find('w:rPr/w:vanish', NS) is not None
                if in_counter and hidden: continue
                texts.append(el.text or '')
        text = ''.join(texts)
        if role in ('chapter', 'section', 'subsection', 'subsubsection'):
            text = re.sub(r'^(?:第[一二三四五六七八九十百\d]+章\s*|[一二三四五六七八九十百]+[、．.]\s*|[（(][一二三四五六七八九十百\d]+[）)]\s*|\d+(?:\.\d+){1,3}[.、．]?\s*|\d+[.、．]\s*)', '', text)
        if role in ('figure_caption', 'table_caption'):
            text = re.sub(r'^([图表]\s*\d+)[.－—](\d+)', r'\1-\2', text)
        if role == 'equation':
            # Only the literal presentation label differs; expressions are
            # compared structurally by protected().
            text = re.sub(r'^公式\s*', '', text).replace('（', '(').replace('）', ')')
        if text.strip(' \u3000'): values.append(text)
    return collections.Counter(values)

def verified_display_edit(op, original_paragraphs):
    """Validate bounded display edits against the ORIGINAL bibliography."""
    before, after = op.get('before'), op.get('after')
    span = op.get('edit_span', {})
    start, end = span.get('start'), span.get('end')
    replacement = span.get('replacement')
    if not isinstance(before, str) or not isinstance(after, str) or type(start) is not int or type(end) is not int or not isinstance(replacement, str):
        return False
    if not 0 <= start <= end <= len(before) or before[start:end] != span.get('expected_text'):
        return False
    if after != before[:start] + replacement + before[end:]:
        return False
    kind = op.get('kind')
    if kind == 'citation_marker_insertion':
        refs = op.get('references', [])
        if start != end or not re.fullmatch(r'\[\d+(?:,\d+)*\]', replacement) or not refs:
            return False
        numbers = []
        for ref in refs:
            source = ref.get('expected_text', ref.get('source_text'))
            number = str(ref.get('reference_number', ref.get('number')))
            if source not in original_paragraphs or not re.match(r'^\s*\[' + re.escape(number) + r'\]', source):
                return False
            numbers.append(number)
        return replacement == '[' + ','.join(numbers) + ']'
    ref = op.get('reference', {})
    source = ref.get('source_text')
    authors = ref.get('authors', [])
    author_span = ref.get('span', {})
    if source not in original_paragraphs or not authors:
        return False
    a, b = author_span.get('start'), author_span.get('end')
    if type(a) is not int or type(b) is not int or source[a:b] != author_span.get('expected_text'):
        return False
    if any(not x.get('exact_name') or x['exact_name'] not in source[a:b] or x.get('display_name', '').lower() not in x['exact_name'].lower() for x in authors):
        return False
    english = authors[0].get('language') == 'en'
    if kind == 'bibliography_author_display':
        if len(authors) <= 3 or before != source or (start, end) != (a, b):
            return False
        expected = (', ' if english else ',').join(x['exact_name'] for x in authors[:3]) + (', et al.' if english else ',等')
    elif kind == 'narrative_author_display':
        if authors[0]['display_name'].lower() not in span['expected_text'].lower():
            return False
        remainder = span['expected_text']
        names = sorted({n for author in authors for n in (author.get('exact_name', ''), author.get('display_name', '')) if n}, key=len, reverse=True)
        for name in names:
            remainder = re.sub(re.escape(name), '', remainder, flags=re.I)
        remainder = re.sub(r'\bet\s+al\.?|\band\b|等|和|与', '', remainder, flags=re.I)
        if re.sub(r'[\s,，、&.（）()]', '', remainder):
            # An author display span cannot include a claim, year or other
            # prose merely because a real surname occurs somewhere in it.
            return False
        if len(authors) == 1: expected = authors[0]['display_name']
        elif len(authors) == 2: expected = '和'.join(x['display_name'] for x in authors)
        else: expected = authors[0]['display_name'] + (' et al.' if english else '等')
    else:
        return False
    return replacement == expected

def verified_continuation_edit(op, parts):
    """Only exact textual source headers may be duplicated by a table split."""
    tid = op.get('source_table_id')
    if not isinstance(tid, str) or not re.fullmatch(r't\d{4}', tid): return False
    tables = list(parse(parts['word/document.xml']).iter(q('tbl')))
    index = int(tid[1:]) - 1
    if not 0 <= index < len(tables): return False
    rows = tables[index].findall(q('tr'))
    header_rows, breaks = op.get('header_rows'), op.get('breaks')
    if type(header_rows) is not int or not 1 <= header_rows < len(rows) or not isinstance(breaks,list) or not breaks:
        return False
    if any(type(b) is not int or not header_rows < b < len(rows) for b in breaks) or breaks != sorted(set(breaks)):
        return False
    forbidden = {q(x) for x in ('fldChar','fldSimple','bookmarkStart','bookmarkEnd','drawing','pict','object','ins','del','moveFrom','moveTo','commentRangeStart','commentRangeEnd','footnoteReference','endnoteReference')}
    forbidden.add('{' + M + '}oMath')
    if any(el.tag in forbidden for row in rows[:header_rows] for el in row.iter()): return False
    header = [''.join(p.xpath('.//w:t/text()', namespaces=NS)) for row in rows[:header_rows] for p in row.iter(q('p'))]
    expected = [t for t in header if t.strip(' \u3000')] * len(breaks)
    actual = [t for t in op.get('added_header_texts', []) if t.strip(' \u3000')]
    return expected == actual and op.get('label_count') == len(breaks)

def verify_transition(source, final, kind, report=None):
    source, final = Path(source), Path(final)
    before_hash, after_hash = sha(source), sha(final)
    errors, checks = [], {}
    report = report or {}
    try:
        a, b = read_package(source), read_package(final)
        pa, pb = protected(a), protected(b)
        approved_math = pa['math'].copy()
        transcriptions=[]
        if kind in ('format', 'format_document'):
            from math_transcription import verified_record, parse_math
            source_ps=list(parse(a['word/document.xml']).iter(q('p')));seen=set()
            for op in report.get('operations', []):
                if op.get('kind')!='formula_transcription':continue
                pid=op.get('locator')
                if isinstance(pid,dict):pid=pid.get('paragraph_id',pid.get('id'))
                index=int(pid[1:])-1 if isinstance(pid,str) and re.fullmatch(r'p\d{6}',pid) else -1
                if not 0<=index<len(source_ps) or pid in seen or not verified_record(op,source_ps[index]):
                    errors.append('Invalid source-identity formula transcription: '+str(pid));continue
                seen.add(pid);transcriptions.append(op)
                approved_math[digest(tree(parse_math(op['omml']),frozenset((q('rPr'),'{'+M+'}rPr','{'+M+'}ctrlPr'))))]+=1
        for key in ('math', 'revisions', 'notes'):
            checks[key] = (approved_math if key=='math' else pa[key]) == pb[key]
            if not checks[key]:
                errors.append('Protected content changed: ' + key)
        # Formatting may import a school logo, but never remove existing assets.
        missing = pa['binaries'] - pb['binaries']
        checks['original_assets_retained'] = not missing
        if missing:
            errors.append('Original media/chart/embedded assets changed or lost')
        if kind in ('format', 'format_document'):
            from format_engine import format_document
            with tempfile.TemporaryDirectory(prefix='fit-replay-') as td:
                expected = Path(td) / 'expected.docx'
                format_document(source, expected, decisions=report.get('decisions', report.get('source_mapping')), profile=report.get('profile'))
                expected_parts = read_package(expected)
                checks['deterministic_replay'] = expected_parts == b
                if not checks['deterministic_replay']:
                    changed = sorted(n for n in expected_parts.keys() | b.keys() if expected_parts.get(n) != b.get(n))
                    errors.append('Output differs from deterministic authorized replay: ' + ', '.join(changed[:12]))
            if report.get('source_mapping') and report.get('final_mapping'):
                old_text = research_text(a, report['source_mapping'])
                new_text = research_text(b, report['final_mapping'])
                original_paragraphs = [''.join(p.xpath('.//w:t/text()', namespaces=NS)) for p in parse(a['word/document.xml']).iter(q('p'))]
                # Citation display changes require their separately verified
                # source-bound map. The engine reports exact old/new paragraphs.
                for op in report.get('operations', []):
                    if op in transcriptions:
                        old=re.sub(r'^公式\s*','',op['before']).replace('（','(').replace('）',')')
                        new=re.sub(r'^公式\s*','',op['after']).replace('（','(').replace('）',')')
                        if old_text[old]<=0:errors.append('Transcribed expression not found in original research payload')
                        else:
                            old_text[old]-=1
                            if new:old_text[new]+=1
                    if op.get('kind') == 'paragraph_padding':
                        before, after = op.get('before', ''), op.get('after', '')
                        if after != before.strip(' \u3000') or old_text[before] <= 0:
                            errors.append('Invalid bounded paragraph-padding change')
                        else:
                            old_text[before] -= 1
                            if after: old_text[after] += 1
                    if op.get('kind') == 'table_continuation':
                        if not verified_continuation_edit(op, a):
                            errors.append('Invalid source-bound continuation headers: '+str(op.get('source_table_id')))
                        else:
                            old_text.update(t for t in op['added_header_texts'] if t.strip(' \u3000'))
                            old_text['续前表'] += op['label_count']
                    if op.get('kind') in ('bibliography_author_display', 'narrative_author_display', 'citation_marker_insertion'):
                        if not verified_display_edit(op, original_paragraphs):
                            errors.append('Invalid source-bound display operation: ' + str(op.get('locator')))
                        elif old_text[op['before']] > 0:
                            old_text[op['before']] -= 1
                            old_text[op['after']] += 1
                        else:
                            errors.append('Display edit not found in original payload chain')
                    if op.get('kind') == 'equation_label' and not op.get('before', '').strip():
                        label = op.get('after', '')
                        if re.fullmatch(r'公式[（(]\d+-\d+[）)]', label) and checks.get('math'):
                            # Explicit newly supplied presentation label only;
                            # no mathematical or ordinary text is exempted.
                            normalized = label.removeprefix('公式').replace('（', '(').replace('）', ')')
                            old_text[normalized] += 1
                lost, added = old_text - new_text, new_text - old_text
                checks['research_payloads'] = not lost and not added
                if lost or added:
                    errors.append('Independent research text mismatch: ' + json.dumps({'lost': list(lost)[:5], 'added': list(added)[:5]}, ensure_ascii=False))
        elif kind in ('office_fields', 'field_refresh'):
            stories_a, stories_b = effective_stories(a), effective_stories(b)
            names_a, names_b = set(stories_a), set(stories_b)
            if names_a != names_b:
                errors.append('Office refresh changed document story set')
            for name in sorted(names_a & names_b):
                sa, sb = stories_a[name], stories_b[name]
                ta = story_tokens(parse(sa if isinstance(sa, bytes) else a[sa]), a, sa if isinstance(sa, str) else name)
                tb = story_tokens(parse(sb if isinstance(sb, bytes) else b[sb]), b, sb if isinstance(sb, str) else name)
                if ta != tb:
                    errors.append('Office changed non-cache text, semantic formatting, fields or controls: ' + name)
            checks['only_approved_field_cache_changes'] = not errors
        elif kind == 'unchanged':
            if a != b:
                errors.append('Unchanged stage contains package changes')
        else:
            errors.append('Unregistered transition kind: ' + str(kind))
    except Exception as exc:
        errors.append(type(exc).__name__ + ': ' + str(exc))
    if sha(source) != before_hash or sha(final) != after_hash:
        errors.append('Input changed during validation')
    return {'passed': not errors, 'errors': errors, 'checks': checks,
            'source_sha256': before_hash, 'final_sha256': after_hash,
            'visual_acceptance': False}
