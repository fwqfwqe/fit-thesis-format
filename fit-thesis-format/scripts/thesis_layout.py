#!/usr/bin/env python3
"""Install reviewed FIT running furniture without editing article text.

Separate from thesis_format.py: source-bound region boundaries and a verbatim
title locator authorize only section layout, new header/footer parts and their
relationships. Existing story parts remain byte-identical. Verify recomputes
the exact permitted transformation and compares every ZIP member.
"""
from __future__ import annotations
import argparse
import copy
import json
from pathlib import Path
import sys
import tempfile
import zipfile
import posixpath
import re
import thesis_format as F

E = F.E
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
SCHOOL = '福州理工学院本科生毕业论文（设计）'
DOC = 'word/document.xml'
RELS = 'word/_rels/document.xml.rels'
SETTINGS = 'word/settings.xml'


def sections(root):
    body = root.find(F.q('body'))
    found = []
    for i, block in enumerate(body):
        sec = block if block.tag == F.q('sectPr') else block.find('w:pPr/w:sectPr', F.NS) if block.tag == F.q('p') else None
        if sec is not None:
            found.append((i, sec))
    if len(found) != len(list(root.iter(F.q('sectPr')))):
        raise ValueError('Nested/revised section properties require manual review.')
    if not found or found[-1][1] is not body[-1]:
        raise ValueError('Expected one final body-level sectPr.')
    return body, found


def title_text(parts, locator):
    part = locator.get('part', DOC)
    if part != DOC and not (part.startswith('word/header') and F.allowed_story(part)):
        raise ValueError('Title must be located in this manuscript or an existing header.')
    if part not in parts:
        raise ValueError('Title part is absent.')
    ps, _ = F.indexed(F.parse(parts[part]))
    p = ps.get(locator.get('paragraph_id'))
    if p is None:
        raise ValueError('Title paragraph is absent.')
    # Do not copy a title with ambiguous field results, objects or revisions.
    permitted = {F.q(x) for x in ('p', 'pPr', 'r', 'rPr', 't','br','bookmarkStart','bookmarkEnd','proofErr')}
    for el in p.iter():
        if any(a.tag in {F.q('pPr'), F.q('rPr')} for a in el.iterancestors()):
            continue
        if el.tag not in permitted:
            raise ValueError('Title locator must select a plain-text paragraph; review complex titles manually.')
    text = ''.join(x.text or '' for x in p.iter(F.q('t')))
    if not text.strip() or text != locator.get('expected_text'):
        raise ValueError('Verbatim title expected_text missing or different.')
    return text


def check_old_furniture(parts, root):
    rels = F.parse(parts[RELS])
    idx = {r.get('Id'): r for r in rels}
    for s in root.iter(F.q('sectPr')):
        for ref in list(s):
            if ref.tag not in {F.q('headerReference'), F.q('footerReference')}:
                continue
            r = idx.get(ref.get('{' + R + '}id'))
            if r is None or r.get('TargetMode') == 'External':
                raise ValueError('Missing or external running-furniture relationship.')
            name = posixpath.normpath(posixpath.join('word', r.get('Target', ''))).lstrip('/')
            if r.get('Target', '').startswith('/'):
                name = r.get('Target').lstrip('/')
            if name not in parts:
                raise ValueError('Missing running-furniture part: ' + name)
            story = F.parse(parts[name])
            forbidden = F.PROTECTED_TAGS | {F.q(x) for x in ('tbl', 'ins', 'del', 'sdt', 'hyperlink', 'altChunk')}
            codes=[(el.text or '').strip() for el in story.iter(F.q('instrText'))]
            cached=''.join(story.xpath('.//w:t/text()',namespaces=F.NS)).strip()
            pure_page_shape=(bool(codes) and all(re.fullmatch(r'PAGE(?:\s+\\\*\s+MERGEFORMAT)?',code,re.I) for code in codes)
                             and (not cached or cached.isdigit())
                             and not any(attr.startswith('{'+R+'}') for el in story.iter() for attr in el.attrib)
                             and not any(el.tag in {F.q('object'),F.q('ins'),F.q('del'),F.q('hyperlink')} for el in story.iter()))
            if any(el.tag in forbidden for el in story.iter()) and not pure_page_shape:
                raise ValueError('Existing header/footer contains protected objects, revisions or structured content: ' + name)


def story(kind, text=None, page=False):
    root = E.Element(F.q('hdr' if kind == 'header' else 'ftr'), nsmap={'w': F.W})
    p = F.child(root, 'p')
    pp = F.child(p, 'pPr')
    if text is not None or page:
        F.child(pp, 'pStyle', val=F.style_id('page_number' if page else 'header'))
    else:
        # Empty front matter header must not inherit a ruled header style.
        F.child(pp, 'spacing', before=0, after=0, line=20, lineRule='exact')
    if text is not None:
        t = F.child(F.child(p, 'r'), 't')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
    if page:
        r = F.child(p, 'r'); F.child(r, 'fldChar', fldCharType='begin')
        r = F.child(p, 'r'); t = F.child(r, 'instrText'); t.text = ' PAGE '
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r = F.child(p, 'r'); F.child(r, 'fldChar', fldCharType='separate')
        r = F.child(p, 'r'); F.child(r, 't').text = '1'
        r = F.child(p, 'r'); F.child(r, 'fldChar', fldCharType='end')
    return F.xml(root)


def inspect(path):
    parts = F.package(path); root = F.parse(parts[DOC]); body, secs = sections(root)
    ps, _ = F.indexed(root); pid = {p: k for k, p in ps.items()}
    rows = []
    start = 0
    for j, (end, s) in enumerate(secs, 1):
        rows.append({'id': f's{j:03d}', 'first_top_level_paragraph': next((pid[b] for b in list(body)[start:end+1] if b in pid), None),
                     'end_paragraph': pid.get(s.getparent().getparent()) if s.getparent().tag == F.q('pPr') else None,
                     'page_number': dict(s.find(F.q('pgNumType')).attrib) if s.find(F.q('pgNumType')) is not None else {},
                     'different_first_page': s.find(F.q('titlePg')) is not None})
        start = end + 1
    return {'source_sha256': F.sha(Path(path).read_bytes()), 'sections': rows,
            'top_level_paragraphs': [{'id': pid[b], 'text': F.ptext(b)} for b in body if b in pid],
            'note': 'Use start for the first region. Other boundaries must be reviewed top-level paragraph IDs. Exact manuscript title locator required.'}


def transform(parts, mapping):
    out = dict(parts); root = F.parse(parts[DOC]); body, original_secs = sections(root)
    check_old_furniture(parts, root)
    title = title_text(parts, mapping['title'])
    ps, _ = F.indexed(root)
    regs = mapping.get('regions')
    if not isinstance(regs, list) or not regs or regs[0].get('start') != 'start':
        raise ValueError('regions must begin with start=start.')
    positions = []
    rank = {'unpaged': 0, 'front': 1, 'body': 2}
    for i, reg in enumerate(regs):
        if reg.get('kind') not in rank:
            raise ValueError('Region kind must be unpaged, front or body.')
        if i == 0:
            pos = 0
        else:
            p = ps.get(reg.get('start'))
            if p is None or p.getparent() is not body:
                raise ValueError('Region boundary must be an existing top-level paragraph.')
            if 'expected_text' not in reg or reg['expected_text'] != F.ptext(p):
                raise ValueError('Region boundary expected_text missing or different.')
            pos = body.index(p)
        if positions and (pos <= positions[-1][0] or rank[reg['kind']] <= rank[positions[-1][1]]):
            raise ValueError('Regions must be ordered, distinct and progress unpaged -> front -> body (optional earlier kinds).')
        positions.append((pos, reg['kind']))
    if positions[-1][1] != 'body':
        raise ValueError('Final region must be body; all back matter continues its page numbers.')
    needs_split = any(body[pos - 1].find('w:pPr/w:sectPr', F.NS) is None for pos, _ in positions[1:])
    if needs_split:
        # Added sections can change visible note numbers or SECTION fields even
        # when every stored character and note part remains byte-identical.
        settings_source = F.parse(parts[SETTINGS]) if SETTINGS in parts else None
        note_roots = [root] + ([settings_source] if settings_source is not None else [])
        if any(el.get(F.q('val')) == 'eachSect' for rr in note_roots for el in rr.findall('.//w:numRestart', F.NS)):
            raise ValueError('New sections could renumber footnotes/endnotes configured to restart each section; requires dedicated review.')
        instructions = [el.text or '' for el in root.iter(F.q('instrText'))]
        instructions += [el.get(F.q('instr'), '') for el in root.iter(F.q('fldSimple'))]
        if any(re.search(r'\bSECTION(?:PAGES)?\b', text, re.I) for text in instructions):
            raise ValueError('New sections could change SECTION/SECTIONPAGES fields in the article; requires dedicated review.')
    # Add only section properties, never a new paragraph or a text break.
    for pos, kind in positions[1:]:
        prev = body[pos - 1]
        if prev.tag != F.q('p'):
            raise ValueError('Boundary follows a table/structured block; use a reviewed existing section or a dedicated layout edit.')
        pp = F.ensure(prev, 'pPr', True)
        if pp.find(F.q('sectPr')) is None:
            inherited = next(s for end, s in original_secs if end >= pos)
            pp.append(copy.deepcopy(inherited))
    body, secs = sections(root)
    styles = F.parse(parts['word/styles.xml'])
    for role in ('header', 'page_number'):
        if styles.find(f'w:style[@w:styleId="{F.style_id(role)}"]', F.NS) is None:
            raise ValueError('Run thesis_format.py first to install editable header/page-number styles.')
    rels = F.parse(parts[RELS]); types = F.parse(parts['[Content_Types].xml'])
    taken_ids = {r.get('Id') for r in rels}
    added = []

    def add_part(kind, data):
        # Reuse only parts already owned by this transformer; never adopt or
        # mutate user headers. Stable byte content means a second run adds none.
        for existing in rels:
            if not (existing.get('Id','').startswith('rIdFITLayout') and existing.get('Type') == R + '/' + kind): continue
            target = existing.get('Target','')
            name = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word',target))
            if out.get(name) == data: return existing.get('Id')
        n = 1
        while f'word/{kind}{n}.xml' in out:
            n += 1
        name = f'word/{kind}{n}.xml'
        out[name] = data
        ri = 1
        while f'rIdFITLayout{ri}' in taken_ids:
            ri += 1
        rid = f'rIdFITLayout{ri}'; taken_ids.add(rid)
        el = E.SubElement(rels, '{' + PKG + '}Relationship')
        el.set('Id', rid); el.set('Type', R + '/' + kind); el.set('Target', name.removeprefix('word/'))
        el = E.SubElement(types, '{' + CT + '}Override')
        el.set('PartName', '/' + name)
        el.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
        added.append(name)
        return rid

    blank_h = add_part('header', story('header'))
    odd_h = add_part('header', story('header', title))
    even_h = add_part('header', story('header', SCHOOL))
    blank_f = add_part('footer', story('footer'))
    page_f = add_part('footer', story('footer', page=True))
    start = 0; seen = set(); section_report = []
    for i, (end, s) in enumerate(secs, 1):
        kind = next(k for pos, k in reversed(positions) if pos <= start)
        first_kind = kind not in seen
        seen.add(kind)
        F.remove_names(s, {'headerReference', 'footerReference', 'titlePg', 'pgNumType'})
        refs = []
        for tag in ('headerReference', 'footerReference'):
            for typ in ('default', 'even', 'first'):
                el = E.Element(F.q(tag)); el.set(F.q('type'), typ)
                rid = (even_h if typ == 'even' else odd_h) if tag == 'headerReference' and kind == 'body' else blank_h if tag == 'headerReference' else blank_f if kind == 'unpaged' else page_f
                el.set('{' + R + '}id', rid); refs.append(el)
        for offset, el in enumerate(refs):
            s.insert(offset, el)
        if kind != 'unpaged':
            pg = E.Element(F.q('pgNumType')); pg.set(F.q('fmt'), 'upperRoman' if kind == 'front' else 'decimal')
            if first_kind:
                pg.set(F.q('start'), '1')
            # Schema order: pgNumType follows pgBorders/lnNumType, before cols.
            following = {'cols', 'formProt', 'vAlign', 'noEndnote', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'}
            idx = next((j for j, e in enumerate(s) if E.QName(e).localname in following), len(s))
            s.insert(idx, pg)
        if first_kind and start > 0:
            typ = s.find(F.q('type'))
            if typ is None:
                typ = E.Element(F.q('type')); s.insert(len(refs), typ)
            typ.set(F.q('val'), 'nextPage')
        section_report.append({'section': i, 'kind': kind, 'restart': first_kind and kind != 'unpaged'})
        start = end + 1
    if SETTINGS not in parts:
        raise ValueError('Missing settings part; review unusual package manually.')
    settings = F.parse(parts[SETTINGS]); F.remove_names(settings, {'evenAndOddHeaders'})
    # Keep settings schema order; common successors to evenAndOddHeaders.
    marker = E.Element(F.q('evenAndOddHeaders'))
    later = {'bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'}
    idx = next((i for i, e in enumerate(settings) if E.QName(e).localname in later), len(settings))
    settings.insert(idx, marker)
    out[DOC] = F.xml(root); out[RELS] = F.xml(rels)
    out['[Content_Types].xml'] = F.xml(types); out[SETTINGS] = F.xml(settings)
    # Independent article guard, ignoring ONLY section/paragraph formatting.
    if F.content_skeleton(F.parse(parts[DOC])) != F.content_skeleton(root):
        raise ValueError('Article structure changed during layout transformation.')
    return out, {'sections': section_report, 'added_parts': added,
                 'article_content_preserved': True, 'original_story_parts_preserved': True,
                 'note': 'Title and fixed school name are running furniture; PAGE fields are layout. Final physical pagination requires Word/PDF review.'}


def checked_transform(source, mapping):
    if mapping.get('source_sha256') != F.sha(Path(source).read_bytes()):
        raise ValueError('Source hash mismatch; inspect exact intermediate DOCX again.')
    return transform(F.package(source), mapping)


def verify(output, source, mapping):
    expected, info = checked_transform(source, mapping)
    actual = F.package(output)
    errors = []
    if set(expected) != set(actual):
        errors.append('ZIP member set differs from permitted layout transformation.')
    errors.extend('Unexpected change: ' + n for n in expected if actual.get(n) != expected[n])
    return {'ok': not errors, 'errors': errors, **info,
            'article_content_preserved': not errors,
            'source_sha256': F.sha(Path(source).read_bytes()), 'output_sha256': F.sha(Path(output).read_bytes())}


def apply(source, output, mapping):
    src = Path(source).resolve(); dest = Path(output).resolve()
    if src == dest or dest.exists():
        raise ValueError('Choose a new output path; original and existing outputs are never overwritten.')
    parts, _ = checked_transform(src, mapping)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix='.fit-layout-', suffix='.docx', dir=dest.parent, delete=False) as handle:
        temp = Path(handle.name)
    try:
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
            existing = set()
            for info in zin.infolist():
                zout.writestr(info, parts[info.filename]); existing.add(info.filename)
            for n in parts.keys() - existing:
                zout.writestr(n, parts[n])
        result = verify(temp, src, mapping)
        if not result['ok']:
            raise ValueError('; '.join(result['errors']))
        temp.rename(dest)
        result.update({'source': str(src), 'output': str(dest)})
        return result
    finally:
        if temp.exists():
            temp.unlink()


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8'); sys.stderr.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description=__doc__); sp = ap.add_subparsers(dest='command', required=True)
    ip = sp.add_parser('inspect'); ip.add_argument('input'); ip.add_argument('--out', required=True)
    aa = sp.add_parser('apply'); aa.add_argument('input'); aa.add_argument('output'); aa.add_argument('--map', required=True); aa.add_argument('--report', required=True)
    vp = sp.add_parser('verify'); vp.add_argument('output'); vp.add_argument('--source', required=True); vp.add_argument('--map', required=True); vp.add_argument('--report')
    args = ap.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect(args.input); F.write_json(args.out, result); print(json.dumps({'ok': True, 'manifest': args.out}, ensure_ascii=False)); return 0
        mapping = F.read_json(args.map)
        result = apply(args.input, args.output, mapping) if args.command == 'apply' else verify(args.output, args.source, mapping)
        if args.report:
            F.write_json(args.report, result)
        print(json.dumps(result, ensure_ascii=False, indent=2)); return 0 if result['ok'] else 2
    except (ValueError, KeyError, OSError, zipfile.BadZipFile, E.XMLSyntaxError) as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}, ensure_ascii=False), file=sys.stderr); return 2


if __name__ == '__main__':
    sys.exit(main())
