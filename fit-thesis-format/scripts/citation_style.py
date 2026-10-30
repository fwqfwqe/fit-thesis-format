#!/usr/bin/env python3
"""Style explicitly reviewed body citations, without changing source characters.

Run after the base thesis formatter has passed its own verification and before
header/footer layout. Each stage uses the immediately preceding DOCX as source.
No automatic candidate acceptance, citation relocation or bibliography editing.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
import re
import sys
import tempfile
import zipfile

from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'w': W}
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
STYLE_ID = 'FIT_citation_superscript'
STYLE_NAME = '论文-文献引用上标'
CITATION = re.compile(r'\[\d+(?:\s*[,，、;；\-–—]\s*\d+)*\]')
TEXT_TAGS = {f'{{{n}}}t' for n in (W, M, A)} | {
    f'{{{W}}}{n}' for n in ('delText', 'instrText', 'delInstrText')}
REVISION_NAMES = {'pPrChange', 'rPrChange', 'ins', 'del', 'moveFrom', 'moveTo'}
FONT_POSITION_PROPS = {'rFonts', 'sz', 'szCs', 'position', 'vertAlign', 'rStyle'}


def q(name):
    return '{' + W + '}' + name


def sha(data):
    return hashlib.sha256(data).hexdigest()


def parse(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True,
                                        remove_blank_text=False))


def serialize(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write_json(path, value):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def package(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP member names are not supported')
        if 'word/document.xml' not in names or 'word/styles.xml' not in names:
            raise ValueError('DOCX must contain document.xml and styles.xml')
        return {name: z.read(name) for name in names}


def indexed(root):
    # Same body paragraph IDs as thesis_format.py, including uneditable contexts.
    return {f'p{i:06d}': p for i, p in enumerate(root.iter(q('p')), 1)}


def ptext(p):
    return ''.join(e.text or '' for e in p.iter() if e.tag in TEXT_TAGS)


def reference_contexts(root, styles):
    """Conservative safeguards supplement the mandatory human-reviewed map."""
    style_names = {s.get(q('styleId')): s.find(q('name')).get(q('val'), '')
                   for s in styles.findall(q('style')) if s.find(q('name')) is not None}
    in_references = False
    result = {}
    for pid, p in indexed(root).items():
        t = ptext(p).strip()
        compact = re.sub(r'\s+', '', t).lower()
        ps = p.find('w:pPr/w:pStyle', NS)
        sid = ps.get(q('val'), '') if ps is not None else ''
        name = style_names.get(sid, sid).lower()
        if compact in {'参考文献', 'references', 'bibliography'}:
            in_references = True
        elif re.fullmatch(r'(?:致谢|acknowledg(?:e)?ments?|附录[一二三四五六七八九十a-z0-9]*|appendix[a-z0-9]*)', compact):
            in_references = False
        style_reference = bool(re.search(r'参考文献|bibliograph|reference[_ -]?(?:entry|heading|list|item)', name + ' ' + sid.lower()))
        leading_number = bool(CITATION.match(t))
        result[pid] = in_references or style_reference or leading_number
    return result


def safety_issue(p, is_reference=False):
    if is_reference:
        return 'Reference-list context or leading bracketed number is protected'
    if p.getparent() is None or p.getparent().tag != q('body'):
        return 'Only direct body paragraphs are supported; tables/textboxes/containers are protected'
    for e in p.iter():
        if isinstance(e.tag, str) and E.QName(e).localname in REVISION_NAMES:
            return 'Tracked or formatting revisions are protected'
    for child in p:
        if child.tag == q('pPr'):
            continue
        if child.tag != q('r'):
            return 'Paragraph contains a bookmark, field, link, equation, anchor or other complex structure'
        if any(n.tag not in {q('rPr'), q('t'), q('lastRenderedPageBreak')} for n in child):
            return 'Run contains non-text content such as a field, break, tab, drawing or note marker'
        if len(child.findall(q('rPr'))) > 1:
            return 'Malformed run properties'
        if any(len(t) for t in child.findall(q('t'))):
            return 'Nested text nodes are unsupported'
    return None


def run_text(r):
    return ''.join(t.text or '' for t in r.findall(q('t')))


def minimal_style():
    s = E.Element(q('style'), {q('type'): 'character', q('customStyle'): '1',
                               q('styleId'): STYLE_ID})
    E.SubElement(s, q('name'), {q('val'): STYLE_NAME})
    E.SubElement(s, q('uiPriority'), {q('val'): '30'})
    E.SubElement(s, q('qFormat'))
    rp = E.SubElement(s, q('rPr'))
    E.SubElement(rp, q('vertAlign'), {q('val'): 'superscript'})
    return s


def style_signature(e):
    return (e.tag, tuple(sorted(e.attrib.items())), e.text,
            tuple(style_signature(c) for c in e))


def add_style(styles):
    expected = minimal_style()
    for s in styles.findall(q('style')):
        name = s.find(q('name'))
        if s.get(q('styleId')) == STYLE_ID:
            if style_signature(s) != style_signature(expected):
                raise ValueError('Existing citation style differs; refusing to alter unrelated styled text')
            return
        if name is not None and name.get(q('val')) == STYLE_NAME:
            raise ValueError('A different style already uses the citation style display name')
    styles.append(expected)


def style_run(r):
    rp = r.find(q('rPr'))
    if rp is None:
        rp = E.Element(q('rPr'))
        r.insert(0, rp)
    rs = rp.find(q('rStyle'))
    if rs is not None and rs.get(q('val')) != STYLE_ID:
        raise ValueError('Selected citation has another character style; preserve it for manual review')
    for child in list(rp):
        if child.tag in {q(n) for n in FONT_POSITION_PROPS}:
            rp.remove(child)
    # Font/size and vertical position inherit normally except this style's superscript.
    # All remaining semantic properties, e.g. bold, italic or underline, survive.
    rp.insert(0, E.Element(q('rStyle'), {q('val'): STYLE_ID}))


def sliced_run(r, start, end):
    """Copy the selected character interval and retain original text-node attributes."""
    clone = copy.deepcopy(r)
    for child in list(clone):
        if child.tag == q('t'):
            clone.remove(child)
        elif child.tag == q('lastRenderedPageBreak') and start>0:
            clone.remove(child)
    offset = 0
    for original in r.findall(q('t')):
        value = original.text or ''
        lo, hi = max(start - offset, 0), min(end - offset, len(value))
        if lo < hi:
            t = copy.deepcopy(original)
            t.text = value[lo:hi]
            # Boundary spaces are source characters; xml:space protects them in Word.
            if t.text[:1].isspace() or t.text[-1:].isspace():
                t.set(XMLSPACE, 'preserve')
            clone.append(t)
        offset += len(value)
    return clone


def validate_map(mapping, source, ps, ref_context):
    if not isinstance(mapping, dict):
        raise ValueError('Map must be a JSON object')
    if mapping.get('schema_version') != 1 or mapping.get('reviewed') is not True:
        raise ValueError('Map schema_version=1 and reviewed=true are required')
    if mapping.get('source_sha256') != sha(Path(source).read_bytes()):
        raise ValueError('Map source SHA-256 is missing or stale')
    spans = mapping.get('spans')
    if not isinstance(spans, list) or not spans:
        raise ValueError('A nonempty explicitly reviewed spans list is required')
    grouped = {}
    for s in spans:
        if not isinstance(s, dict) or set(s) != {'pid', 'start', 'end', 'expected_text', 'role'}:
            raise ValueError('Each span requires exactly pid/start/end/expected_text/role')
        if s['role'] != 'body_citation':
            raise ValueError('Only reviewed body_citation spans may be styled')
        pid = s['pid']
        if not isinstance(pid, str) or pid not in ps:
            raise ValueError('Unknown paragraph ID')
        p = ps[pid]
        issue = safety_issue(p, ref_context[pid])
        if issue:
            raise ValueError(pid + ': ' + issue)
        if type(s['start']) is not int or type(s['end']) is not int:
            raise ValueError('Span offsets must be integer Unicode code-point offsets')
        start, end = s['start'], s['end']
        t = ptext(p)
        if not 0 <= start < end <= len(t):
            raise ValueError(pid + ': span out of range')
        token = s['expected_text']
        if not isinstance(token, str) or CITATION.fullmatch(token) is None:
            raise ValueError(pid + ': expected_text must be one complete bracketed numeric citation')
        if t[start:end] != token:
            raise ValueError(pid + ': expected citation text mismatch')
        grouped.setdefault(pid, []).append(s)
    for pid, spans in grouped.items():
        spans.sort(key=lambda s: (s['start'], s['end']))
        if any(b['start'] < a['end'] for a, b in zip(spans, spans[1:])):
            raise ValueError(pid + ': duplicate or overlapping spans')
    return grouped


def transform_paragraph(p, spans):
    offset = 0
    for r in list(p):
        if r.tag != q('r'):
            continue
        t = run_text(r)
        a, b = offset, offset + len(t)
        offset = b
        touching = [s for s in spans if s['end'] > a and s['start'] < b]
        if not touching:
            continue
        cuts = {0, len(t)}
        for s in touching:
            cuts.update({max(0, s['start'] - a), min(len(t), s['end'] - a)})
        cuts = sorted(cuts)
        insert_at = p.index(r)
        fragments = []
        for lo, hi in zip(cuts, cuts[1:]):
            c = sliced_run(r, lo, hi)
            selected = any(s['start'] <= a + lo and a + hi <= s['end'] for s in touching)
            if selected:
                style_run(c)
            # Preserve inter-run XML whitespace only once, on the last fragment.
            c.tail = r.tail if hi == len(t) else None
            fragments.append(c)
        p.remove(r)
        for j, fragment in enumerate(fragments):
            p.insert(insert_at + j, fragment)


def expected_parts(source, mapping):
    parts = package(source)
    root = parse(parts['word/document.xml'])
    styles = parse(parts['word/styles.xml'])
    ps = indexed(root)
    before = [ptext(p) for p in ps.values()]
    grouped = validate_map(mapping, source, ps, reference_contexts(root, styles))
    add_style(styles)
    for pid, spans in grouped.items():
        transform_paragraph(ps[pid], spans)
    after = [ptext(p) for p in indexed(root).values()]
    if before != after:
        raise ValueError('Internal invariant failed: paragraph text changed')
    result = dict(parts)
    result['word/document.xml'] = serialize(root)
    result['word/styles.xml'] = serialize(styles)
    return result


def inspect(source):
    parts = package(source)
    root = parse(parts['word/document.xml'])
    styles = parse(parts['word/styles.xml'])
    refs = reference_contexts(root, styles)
    candidates = []
    for pid, p in indexed(root).items():
        t = ptext(p)
        issue = safety_issue(p, refs[pid])
        for m in CITATION.finditer(t):
            candidates.append({'pid': pid, 'start': m.start(), 'end': m.end(),
                               'expected_text': m.group(), 'context': t[max(0, m.start()-30):m.end()+30],
                               'eligible_structure': issue is None, 'reason': issue,
                               'reviewed': False})
    return {'schema_version': 1, 'source': str(Path(source).resolve()),
            'source_sha256': sha(Path(source).read_bytes()), 'candidates': candidates,
            'notes': ['Candidates are not an approved map. Review roles and exact spans manually.',
                      'Offsets are Unicode code points, start-inclusive/end-exclusive, within paragraph text.',
                      'Run base formatting+verification first, then citation styling, then page layout.']}


def verify(source, output, mapping):
    expected = expected_parts(source, mapping)
    actual = package(output)
    errors = []
    if set(expected) != set(actual):
        errors.append('ZIP member set changed')
    for name in sorted(set(expected) & set(actual)):
        if expected[name] != actual[name]:
            errors.append('Part differs from deterministic approved transform: ' + name)
    return {'ok': not errors, 'errors': errors, 'source_sha256': sha(Path(source).read_bytes()),
            'output_sha256': sha(Path(output).read_bytes()),
            'selected_spans': len(mapping['spans']),
            'comparison': 'every package part byte-for-byte against regeneration from source+reviewed map',
            'style_id': STYLE_ID, 'style_name': STYLE_NAME}


def apply(source, output, mapping):
    src, dst = Path(source).resolve(), Path(output).resolve()
    if src == dst or dst.exists():
        raise ValueError('Output must be a new path; source and existing outputs are preserved')
    result = expected_parts(src, mapping)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix='.docx', dir=dst.parent, delete=False) as tmp:
        temp = Path(tmp.name)
    try:
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temp, 'w') as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                zout.writestr(info, result[info.filename])
        check = verify(src, temp, mapping)
        if not check['ok']:
            raise ValueError('Verification failed: ' + '; '.join(check['errors']))
        # No overwrite even if a file appeared during generation.
        with temp.open('rb') as fi, dst.open('xb') as fo:
            while chunk := fi.read(1024 * 1024):
                fo.write(chunk)
        check['output'] = str(dst)
        return check
    finally:
        temp.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subs = parser.add_subparsers(dest='command', required=True)
    scan = subs.add_parser('inspect')
    scan.add_argument('source')
    scan.add_argument('--out', required=True)
    for cmd in ('apply', 'verify'):
        p = subs.add_parser(cmd)
        p.add_argument('--source', required=True)
        p.add_argument('--output', required=True)
        p.add_argument('--map', required=True)
        p.add_argument('--report')
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect(args.source)
            write_json(args.out, result)
            print(json.dumps({'ok': True, 'manifest': args.out,
                              'candidates': len(result['candidates'])}, ensure_ascii=False))
            return 0
        mapping = read_json(args.map)
        result = (apply if args.command == 'apply' else verify)(args.source, args.output, mapping)
        if args.report:
            write_json(args.report, result)
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result['ok'] else 2
    except (OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, E.XMLSyntaxError) as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
