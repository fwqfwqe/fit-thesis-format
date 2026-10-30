#!/usr/bin/env python3
"""Normalize explicitly reviewed narrative citations against existing references.

Only author spelling/display, author-year baseline and a confirmed claim's
numeric marker may change. Inspection is advisory and never creates approval.
Every other package part is compared with a deterministic replay on verification.
"""
from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
import re
import tempfile
import zipfile

from lxml import etree as E

import citation_style as C

YEAR = re.compile(r'(?<![\dA-Za-z])\d{4}[a-z]?(?![\dA-Za-z])')
NUMBER = re.compile(r'^\s*\[(\d+)\]\s*')
HAN_NAME = re.compile(r'[\u3400-\u9fff]{2,8}(?:[·・][\u3400-\u9fff]{1,8})*')
EN_NAME = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’\-]*(?:\s+[A-Za-zÀ-ÖØ-öø-ÿ.'’\-]+)*")
SEPARATOR = r'\s*(?:[,，、;；&]|和|与|and)\s*'
EXPECTED_CONNECTOR = '和'
EXPECTED_ABBREVIATIONS = {'zh': '等', 'en': 'et al.'}
END_PUNCTUATION = '。.!?！？;；'
# Recognition is deliberately wider than the supported canonical style token.
# A real but malformed numeric marker is a review issue, never permission to add
# a second marker. Mixed/fullwidth brackets and inner spaces remain unchanged.
NUMERIC_MARKER = re.compile(r'[\[［]\s*\d[\d\s,，、;；\-–—]*[\]］]')


def is_organization_author(value):
    """Conservative Chinese corporate-author hint; 和 is not a name separator."""
    value = (value or '').strip()
    return bool(3 <= len(value) <= 100
                and re.fullmatch(r'[\u3400-\u9fff]+', value)
                and re.search(r'(?:办公厅|办公局|委员会|人民政府|研究院|研究所|科学院|大学|学院|协会|银行|公司|集团|联合会|中心|总局|部|局)$', value))


def publication_years(text, author_end=0):
    """Separate typed-reference publication years from pages/DOI/access dates.

    Unrecognized formats retain the old conservative year hints.  Recognized
    types with no publication year return no year rather than guessing a page.
    Multiple real publication-year candidates remain ambiguous.
    """
    kind = re.search(r'\[(?:J|M|D|N|R|S|C|G|P|EB|DB|CP|DS|A|Z)(?:/(?:OL|MT|CD|DK))?\]', text[author_end:], re.I)
    if kind is None:
        return [m.group() for m in YEAR.finditer(text[author_end:])]
    publication = text[author_end + kind.end():]
    stop = re.search(r'\bDOI\b|https?://|www\.', publication, re.I)
    if stop: publication = publication[:stop.start()]
    publication = re.sub(r'[\[［]\s*\d{4}(?:[-/.年]\d{1,2}){1,2}日?\s*[\]］]', '', publication)
    page = re.search(r'[:：]\s*(?:pp?\.?\s*)?\d|\bpp?\.\s*\d', publication, re.I)
    if page: publication = publication[:page.start()]
    return list(dict.fromkeys(m.group() for m in YEAR.finditer(publication)))


def _span(value, text, label, keys=('start', 'end', 'expected_text')):
    if not isinstance(value, dict) or set(value) != set(keys):
        raise ValueError(label + ': requires exactly ' + '/'.join(keys))
    start, end = value['start'], value['end']
    if type(start) is not int or type(end) is not int or not 0 <= start < end <= len(text):
        raise ValueError(label + ': invalid Unicode code-point offsets')
    if not isinstance(value['expected_text'], str) or text[start:end] != value['expected_text']:
        raise ValueError(label + ': exact expected text mismatch')
    return start, end


def _name_pattern(author):
    names = {author['exact_name'], author['display_name']}
    return '(?:' + '|'.join(re.escape(n) for n in sorted(names, key=len, reverse=True)) + ')'


def _suffix(residual):
    """Accept separators plus one real terminal truncation label, never names."""
    cleaned = re.sub(r'[\s,，、;；&]', '', residual)
    for word in ('and', '和', '与'):
        cleaned = cleaned.replace(word, '')
    if cleaned in {'', '等', '等人', 'etal.'}:
        return {'': None, '等': '等', '等人': '等人', 'etal.': 'et al.'}[cleaned]
    raise ValueError('Reference author list contains an unproved name or unsupported author syntax')


def bibliography_contexts(root, styles):
    """Actual heading/style scope, excluding leading-number protection hints."""
    style_names = {s.get(C.q('styleId')): s.find(C.q('name')).get(C.q('val'), '')
                   for s in styles.findall(C.q('style')) if s.find(C.q('name')) is not None}
    inside, result = False, {}
    for pid, paragraph in C.indexed(root).items():
        compact = re.sub(r'\s+', '', C.ptext(paragraph)).lower()
        style = paragraph.find('w:pPr/w:pStyle', C.NS)
        sid = style.get(C.q('val'), '') if style is not None else ''
        if compact in {'参考文献', 'references', 'bibliography'}:
            inside = True
        elif re.fullmatch(r'(?:致谢|acknowledg(?:e)?ments?|附录[一二三四五六七八九十a-z0-9]*|appendix[a-z0-9]*)', compact):
            inside = False
        label = style_names.get(sid, sid).lower() + ' ' + sid.lower()
        result[pid] = inside or bool(re.search(r'参考文献|bibliograph|reference[_ -]?(?:entry|heading|list|item)', label))
    return result


def _reference(record, paragraphs, contexts):
    allowed = {'pid', 'expected_text', 'reference_number', 'author_list', 'authors', 'year'}
    if not isinstance(record, dict) or set(record) != allowed:
        raise ValueError('Reference evidence requires exactly ' + '/'.join(sorted(allowed)))
    pid = record['pid']
    if not isinstance(pid, str) or pid not in paragraphs or not contexts[pid]:
        raise ValueError('Reference PID must identify an existing bibliography entry')
    text = C.ptext(paragraphs[pid])
    if record['expected_text'] != text:
        raise ValueError('Reference full text mismatch')
    numbered = NUMBER.match(text)
    number = record['reference_number']
    if type(number) is not int or number < 1 or numbered is None or int(numbered.group(1)) != number:
        raise ValueError('Reference number does not match the existing bibliography')
    start, end = _span(record['author_list'], text, 'reference author_list')
    if start != numbered.end():
        raise ValueError('Reference author_list must cover the whole author area immediately after [n]')
    author_list = text[start:end]
    # The final et al. period can also be the bibliography delimiter. Otherwise
    # the delimiter is outside author_list, including a final initial's period.
    tail = text[end:]
    suffix_is_period = bool(re.search(r'et\s+al\.$', author_list))
    if suffix_is_period:
        if not tail or not tail[0].isspace() or not tail.strip():
            raise ValueError('Truncated author_list must end at the actual bibliography delimiter')
    elif not re.match(r'^[.。．](?!\s*[,，、;；&])\s*\S', tail):
        raise ValueError('Reference author_list must end at a bibliography sentence delimiter')
    authors = record['authors']
    if not isinstance(authors, list) or not authors:
        raise ValueError('Reference must provide every listed author in order')
    cursor, residual, languages = start, '', set()
    for author_index, author in enumerate(authors):
        required = {'start', 'end', 'exact_name', 'display_name', 'language'}
        if not isinstance(author, dict) or set(author) != required:
            raise ValueError('Each author requires start/end/exact_name/display_name/language')
        lo, hi = author['start'], author['end']
        name = author['exact_name']
        if type(lo) is not int or type(hi) is not int or not cursor <= lo < hi <= end:
            raise ValueError('Reference authors must have nonoverlapping ordered exact spans')
        if not isinstance(name, str) or text[lo:hi] != name:
            raise ValueError('Reference author exact_name mismatch')
        if author_index and re.fullmatch(SEPARATOR, text[cursor:lo]) is None:
            raise ValueError('Reference author names must be separated by actual bibliography author delimiters')
        language, display = author['language'], author['display_name']
        if not isinstance(display, str) or not display:
            raise ValueError('Author display_name must be an evidenced name')
        if language == 'zh':
            if (HAN_NAME.fullmatch(name) is None and not is_organization_author(name)) or display != name:
                raise ValueError('Chinese author display must preserve the exact bibliography name')
        elif language == 'en':
            if EN_NAME.fullmatch(name) is None or EN_NAME.fullmatch(display) is None:
                raise ValueError('Unsupported foreign name syntax; preserve for manual review')
            if name != display and not name.startswith(display + ' '):
                raise ValueError('Foreign display_name must be an exact surname prefix at a word boundary')
            display_words = display.split()
            particles = {'de', 'del', 'da', 'das', 'dos', 'di', 'du', 'la', 'le', 'van', 'von', 'der', 'den'}
            if any('.' in word or len(word) < 2 for word in display_words) or (
                    len(display_words) > 1 and any(word.lower() not in particles for word in display_words[:-1])):
                raise ValueError('Foreign display_name must be a surname prefix, not given names or initials')
        else:
            raise ValueError('Author language must be zh or en')
        languages.add(language)
        residual += text[cursor:lo]
        cursor = hi
    residual += text[cursor:end]
    if len(languages) != 1:
        raise ValueError('Mixed author languages require manual review')
    truncation = _suffix(residual)
    # A truncation label must really be at the end, not hidden between names.
    if truncation and (len(authors) != 3 or not author_list.rstrip().endswith(truncation)):
        raise ValueError('Only three listed authors followed by a real terminal truncation label are supported')
    if truncation == 'et al.' and languages != {'en'} or truncation in {'等', '等人'} and languages != {'zh'}:
        raise ValueError('Reference truncation language does not match its authors')
    year_start, year_end = _span(record['year'], text, 'reference year')
    year = record['year']['expected_text']
    if not re.fullmatch(r'\d{4}[a-z]?', year) or year_start < end or not any(
            m.start() == year_start and m.end() == year_end for m in YEAR.finditer(text)):
        raise ValueError('Reference year must be an exact existing bibliography year after the author list')
    return {'authors': authors, 'year': year, 'number': number,
            'language': next(iter(languages)), 'truncation': truncation,
            'author_count_min': 4 if truncation else len(authors)}


def _author_display(record, evidence, source_author):
    authors, language = evidence['authors'], evidence['language']
    count = evidence['author_count_min']
    full_pattern = SEPARATOR.join(_name_pattern(a) for a in authors)
    full = re.fullmatch(full_pattern, source_author)
    first = _name_pattern(authors[0])
    short_pattern = first + (r'\s*(?:等人|等)' if language == 'zh' else r'(?:\s*(?:等人|等)|\s+et\s+al\.)')
    short = re.fullmatch(short_pattern, source_author) is not None
    bare_first = re.fullmatch(first, source_author) is not None
    if full is None and not short and not bare_first:
        raise ValueError('Source authors do not match the evidenced author list or its first author abbreviation')
    if count == 1:
        if 'connector' in record or 'abbreviation' in record:
            raise ValueError('One author does not use connector or abbreviation')
        return authors[0]['display_name']
    if count == 2:
        connector = record.get('connector')
        if connector != EXPECTED_CONNECTOR or 'abbreviation' in record:
            raise ValueError('Two authors require the reviewed target connector 和 for both languages')
        return EXPECTED_CONNECTOR.join(a['display_name'] for a in authors)
    abbreviation = record.get('abbreviation')
    if abbreviation != EXPECTED_ABBREVIATIONS[language] or 'connector' in record:
        raise ValueError('Three or more authors require the reviewed target abbreviation: Chinese 等, foreign et al.')
    return authors[0]['display_name'] + (' ' if abbreviation == 'et al.' else '') + abbreviation


def _baseline_run(run):
    rp = run.find(C.q('rPr'))
    if rp is None:
        rp = E.Element(C.q('rPr'))
        run.insert(0, rp)
    for item in list(rp):
        if item.tag in {C.q('vertAlign'), C.q('position')} or (
                item.tag == C.q('rStyle') and item.get(C.q('val')) == C.STYLE_ID):
            rp.remove(item)
    rp.append(E.Element(C.q('vertAlign'), {C.q('val'): 'baseline'}))


def _author_format(run):
    clone = copy.deepcopy(run.find(C.q('rPr')))
    if clone is None:
        return ()
    for item in list(clone):
        if item.tag in {C.q('vertAlign'), C.q('position')} or (
                item.tag == C.q('rStyle') and item.get(C.q('val')) == C.STYLE_ID):
            clone.remove(item)
    return C.style_signature(clone) if len(clone) or clone.attrib else ()


def _uniform_author_format(paragraph, start, end):
    offset, signatures = 0, []
    for run in paragraph.findall(C.q('r')):
        hi = offset + len(C.run_text(run))
        if offset < end and hi > start:
            signatures.append(_author_format(run))
        offset = hi
    if not signatures or any(s != signatures[0] for s in signatures[1:]):
        raise ValueError('Author replacement crosses different source formatting; preserve for manual review')


def _marker(text, claim, author_year_end, number):
    """Locate an existing marker at the reviewed endpoint or after the year."""
    lo, hi = claim['start'], claim['end']
    markers = [m for m in NUMERIC_MARKER.finditer(text) if lo <= m.start() < hi]
    after = re.match(r'\s*(' + NUMERIC_MARKER.pattern + ')', text[hi:])
    if after:
        token_start = hi + after.start(1)
        token_end = hi + after.end(1)
        markers.append((token_start, token_end, after.group(1)))
    elif re.match('[' + re.escape(END_PUNCTUATION) + r'\s]+(' + NUMERIC_MARKER.pattern + ')', text[hi:]):
        raise ValueError('Existing numeric marker after sentence punctuation requires manual review; do not duplicate or relocate it')
    normalized = [(m.start(), m.end(), m.group()) if hasattr(m, 'start') else m for m in markers]
    if len(normalized) > 1:
        raise ValueError('Multiple numeric markers in the reviewed claim require manual review')
    if normalized:
        start, end, token = normalized[0]
        if start < hi < end:
            raise ValueError('Reviewed claim ends inside an existing numeric marker')
        if token != '[' + str(number) + ']':
            raise ValueError('Existing numeric marker differs from the reviewed reference or is ambiguous')
        if re.match(r'\s*' + NUMERIC_MARKER.pattern, text[end:]):
            raise ValueError('Adjacent numeric markers require manual review')
        trailing = text[end:hi]
        if start < hi and start != author_year_end and trailing.strip(END_PUNCTUATION + ' \t\r\n'):
            raise ValueError('Existing marker is not at a supported claim endpoint or directly after the year')
        return {'start': start, 'end': end, 'expected_text': token}, None
    if hi == author_year_end:
        sentence_end = next((i for i in range(hi, len(text)) if text[i] in END_PUNCTUATION), len(text))
        after_sentence = re.match('[' + re.escape(END_PUNCTUATION) + r'\s]+(' + NUMERIC_MARKER.pattern + ')', text[sentence_end:])
        if NUMERIC_MARKER.search(text[hi:sentence_end]) or after_sentence:
            raise ValueError('Existing numeric marker later in the same sentence is outside the reviewed claim; review the complete claim rather than duplicate it')
        if text[hi:].strip() and re.match(r'\s*[' + re.escape(END_PUNCTUATION) + ']', text[hi:]) is None:
            raise ValueError('Missing numeric marker must follow the complete cited claim; review the substantive text after author-year and select the endpoint before sentence punctuation')
    if hi != author_year_end and not (hi < len(text) and text[hi] in END_PUNCTUATION):
        raise ValueError('Missing marker insertion must be directly after author-year or before the sentence punctuation; claim.end excludes punctuation')
    return None, {'position': hi, 'text': '[' + str(number) + ']'}


def validate_map(mapping, source, paragraphs, contexts, bibliography):
    if not isinstance(mapping, dict) or set(mapping) != {'schema_version', 'reviewed', 'source_sha256', 'records'}:
        raise ValueError('Map requires exactly schema_version/reviewed/source_sha256/records; arbitrary changes are prohibited')
    if type(mapping['schema_version']) is not int or mapping['schema_version'] != 1 or mapping['reviewed'] is not True:
        raise ValueError('Map schema_version=1 and reviewed=true are required')
    if mapping['source_sha256'] != C.sha(Path(source).read_bytes()):
        raise ValueError('Map source SHA-256 is missing or stale')
    if not isinstance(mapping['records'], list) or not mapping['records']:
        raise ValueError('A nonempty explicitly reviewed records list is required')
    grouped = {}
    for record in mapping['records']:
        base = {'pid', 'author_year', 'claim', 'reference'}
        if not isinstance(record, dict) or not base <= set(record) or set(record) - base - {'connector', 'abbreviation'}:
            raise ValueError('Record requires pid/author_year/claim/reference and only the applicable connector or abbreviation')
        pid = record['pid']
        if not isinstance(pid, str) or pid not in paragraphs:
            raise ValueError('Unknown body paragraph ID')
        paragraph = paragraphs[pid]
        issue = C.safety_issue(paragraph, contexts[pid])
        if issue:
            raise ValueError(pid + ': ' + issue)
        text = C.ptext(paragraph)
        ay = record['author_year']
        a, b = _span(ay, text, pid + ' author_year', ('start', 'end', 'expected_text', 'author', 'year'))
        author_start, author_end = _span(ay['author'], text, pid + ' source author')
        year_start, year_end = _span(ay['year'], text, pid + ' source year')
        if a != author_start or not author_end <= year_start < year_end <= b:
            raise ValueError('author_year must start with the exact author span and contain its year')
        between, after_year = text[author_end:year_start], text[year_end:b]
        if not ((between.isspace() and not after_year) or (not between and not after_year) or
                re.fullmatch(r'\s*\(\s*', between) and re.fullmatch(r'\s*\)', after_year) or
                re.fullmatch(r'\s*（\s*', between) and re.fullmatch(r'\s*）', after_year)):
            raise ValueError('Only plain author-year text with optional paired parentheses is supported')
        evidence = _reference(record['reference'], paragraphs, bibliography)
        if evidence['language'] == 'en' and a and re.match(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-]", text[a-1]):
            raise ValueError('Foreign source author must start at a word boundary, not inside an unknown name or word')
        if ay['year']['expected_text'] != evidence['year']:
            raise ValueError('Body year does not match the exact bibliography year, including its suffix')
        display = _author_display(record, evidence, ay['author']['expected_text'])
        if display != ay['author']['expected_text']:
            _uniform_author_format(paragraph, author_start, author_end)
        claim_start, claim_end = _span(record['claim'], text, pid + ' claim')
        if not claim_start <= a < b <= claim_end:
            raise ValueError('Reviewed claim must contain the whole author-year span')
        marker, insert = _marker(text, record['claim'], b, evidence['number'])
        grouped.setdefault(pid, []).append({'baseline': {'start': a, 'end': b},
            'author': {'start': author_start, 'end': author_end, 'replacement': display},
            'claim': {'start': claim_start, 'end': claim_end}, 'marker': marker, 'insert': insert})
    for pid, records in grouped.items():
        records.sort(key=lambda r: (r['claim']['start'], r['claim']['end']))
        if any(b['claim']['start'] < a['claim']['end'] for a, b in zip(records, records[1:])):
            raise ValueError(pid + ': duplicate or overlapping reviewed claims')
        intervals = sorted([(r['baseline']['start'], r['baseline']['end']) for r in records] +
                           [(r['marker']['start'], r['marker']['end']) for r in records if r['marker']])
        if any(b[0] < a[1] for a, b in zip(intervals, intervals[1:])):
            raise ValueError(pid + ': overlapping author-year or marker selections')
        if len({r['insert']['position'] for r in records if r['insert']}) != sum(bool(r['insert']) for r in records):
            raise ValueError(pid + ': duplicate insertion points')
    return grouped


def _new_text_run(template, text):
    run = copy.deepcopy(template)
    for child in list(run):
        if child.tag == C.q('t'):
            run.remove(child)
    original = template.find(C.q('t'))
    node = copy.deepcopy(original) if original is not None else E.Element(C.q('t'))
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set(C.XMLSPACE, 'preserve')
    run.append(node)
    run.tail = None
    return run


def transform_paragraph(paragraph, records):
    baselines = [r['baseline'] for r in records]
    replacements = [r['author'] for r in records if r['author']['replacement'] !=
                    C.ptext(paragraph)[r['author']['start']:r['author']['end']]]
    markers = [r['marker'] for r in records if r['marker']]
    inserts = {r['insert']['position']: r['insert']['text'] for r in records if r['insert']}
    emitted, offset = set(), 0
    intervals = baselines + replacements + markers
    for run in list(paragraph.findall(C.q('r'))):
        text = C.run_text(run)
        a, b = offset, offset + len(text)
        offset = b
        selected = [s for s in intervals if s['end'] > a and s['start'] < b]
        pending = [i for i in inserts if a <= i <= b and i not in emitted]
        if not selected and not pending:
            continue
        cuts = {0, len(text)}
        for span in selected:
            cuts.update({max(0, span['start'] - a), min(len(text), span['end'] - a)})
        cuts.update(i - a for i in pending)
        cuts = sorted(cuts)
        fragments = []

        def insert_at(position):
            if position in inserts and position not in emitted:
                inserted = _new_text_run(run, inserts[position])
                rp = inserted.find(C.q('rPr'))
                if rp is not None:
                    for rs in list(rp.findall(C.q('rStyle'))):
                        rp.remove(rs)
                C.style_run(inserted)
                fragments.append(inserted)
                emitted.add(position)

        for lo, hi in zip(cuts, cuts[1:]):
            x, y = a + lo, a + hi
            insert_at(x)
            replacement = next((s for s in replacements if s['start'] <= x and y <= s['end']), None)
            if replacement:
                if x == replacement['start']:
                    fragment = _new_text_run(run, replacement['replacement'])
                    _baseline_run(fragment)
                    fragments.append(fragment)
                continue
            fragment = C.sliced_run(run, lo, hi)
            fragment.tail = None
            if any(s['start'] <= x and y <= s['end'] for s in baselines):
                _baseline_run(fragment)
            if any(s['start'] <= x and y <= s['end'] for s in markers):
                C.style_run(fragment)
            fragments.append(fragment)
        insert_at(b)
        if fragments:
            fragments[-1].tail = run.tail
        index = paragraph.index(run)
        paragraph.remove(run)
        for i, fragment in enumerate(fragments):
            paragraph.insert(index + i, fragment)
    if emitted != set(inserts):
        raise ValueError('Internal invariant failed: numeric marker was not inserted')


def expected_parts(source, mapping):
    parts = C.package(source)
    root, styles = C.parse(parts['word/document.xml']), C.parse(parts['word/styles.xml'])
    paragraphs = C.indexed(root)
    grouped = validate_map(mapping, source, paragraphs, C.reference_contexts(root, styles), bibliography_contexts(root, styles))
    expected_texts = {pid: C.ptext(p) for pid, p in paragraphs.items()}
    # Compute the only allowed text changes before touching any XML.
    for pid, records in grouped.items():
        changes = []
        for record in records:
            author = record['author']
            changes.append((author['start'], author['end'], author['replacement']))
            if record['insert']:
                insert = record['insert']
                changes.append((insert['position'], insert['position'], insert['text']))
        for lo, hi, value in sorted(changes, reverse=True):
            expected_texts[pid] = expected_texts[pid][:lo] + value + expected_texts[pid][hi:]
    C.add_style(styles)
    for pid, records in grouped.items():
        transform_paragraph(paragraphs[pid], records)
    if expected_texts != {pid: C.ptext(p) for pid, p in C.indexed(root).items()}:
        raise ValueError('Internal invariant failed: text changed outside approved author/marker operations')
    result = dict(parts)
    result['word/document.xml'], result['word/styles.xml'] = C.serialize(root), C.serialize(styles)
    return result


def _guess_reference(text):
    """Conservative hints only; authoritative author/year spans are reviewed."""
    number = NUMBER.match(text)
    if number is None:
        return None
    start, end = number.end(), None
    for m in re.finditer(r'[.。．]', text[start:]):
        position = start + m.start()
        if text[position:position+1] == '.' and re.match(r'^\.\s*[,，、;；&]', text[position:]):
            continue  # Dot after a nonfinal initial.
        end = position
        break
    if end is None:
        return {'reference_number': int(number.group(1)), 'authors': [], 'years': [], 'issues': ['unparsed_author_list']}
    area = text[start:end]
    truncation = re.search(r'(?:[,，、;；]\s*)?(等人|等|et\s+al)$', area)
    if truncation:
        label = truncation.group(1)
        body = area[:truncation.start()]
        suffix = 'et al.' if label.startswith('et') else label
        if suffix == 'et al.':
            end += 1
    else:
        body, suffix = area, None
    # Chinese organization names may contain 和 (工业和信息化部).
    organization=is_organization_author(body.strip()) and not suffix
    names = [body.strip()] if organization else [n.strip() for n in re.split(r'\s*(?:[,，、;；&]|\band\b|和|与)\s*', body) if n.strip()]
    authors = []
    for name in names:
        if HAN_NAME.fullmatch(name) or organization:
            authors.append({'exact_name': name, 'display_name': name, 'language': 'zh'})
        elif EN_NAME.fullmatch(name):
            authors.append({'exact_name': name, 'display_name': name.split()[0], 'language': 'en'})
        else:
            authors = []
            break
    issues = []
    if not authors:
        issues.append('unparsed_author_list')
    if suffix and len(authors) != 3:
        issues.append('unsupported_truncated_author_list')
    return {'reference_number': int(number.group(1)), 'author_list': {'start': start, 'end': end, 'expected_text': text[start:end]},
            'authors': authors, 'years': publication_years(text, end),
            'author_kind': 'organization' if organization else 'personal_or_unconfirmed',
            'truncation': suffix, 'author_count_min': 4 if suffix else len(authors), 'issues': issues}


def _missing_year_mentions(text, references, covered, structure_issue):
    """Flag named, numbered citations even when their author-year token is absent.

    Bibliography names and an existing marker in the same sentence are evidence
    for review, never an automatic year choice. Foreign bibliography capitals
    are compared case-insensitively; boundaries still exclude name substrings.
    """
    found = {}
    for reference in references:
        authors = reference['authors']
        if not authors:
            continue
        foreign = authors[0]['language'] == 'en'
        first = _name_pattern(authors[0])
        full = SEPARATOR.join(_name_pattern(a) for a in authors)
        short = first + (r'\s*(?:等人|等)' if not foreign else r'(?:\s*(?:等人|等)|\s+et\s+al\.)')
        pattern = '(?:' + '|'.join((full, short, first)) + ')'
        for match in re.finditer(pattern, text, re.I if foreign else 0):
            start, end = match.span()
            if foreign and ((start and re.match(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-]", text[start-1])) or
                            (end < len(text) and re.match(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-]", text[end]))):
                continue
            if any(start < hi and end > lo for lo, hi in covered):
                continue
            if re.match(r'\s*(?:[（(]\s*)?\d{4}[a-z]?', text[end:]):
                continue  # Existing years belong to the author-year validator.
            sentence_end = next((i for i in range(end, len(text)) if text[i] in END_PUNCTUATION and
                                 not (text[i] == '.' and i > 0 and i+1 < len(text) and
                                      text[i-1].isdigit() and text[i+1].isdigit())), len(text))
            markers = list(NUMERIC_MARKER.finditer(text[end:sentence_end]))
            if not markers:
                continue  # An unnumbered name may just be an ordinary mention.
            numbers = {int(n) for token in markers for n in re.findall(r'\d+', token.group())}
            if reference['reference_number'] not in numbers:
                continue
            hit = found.setdefault((start, end), {'references': [], 'markers': [m.group() for m in markers]})
            hit['references'].append(reference)
    result = []
    for (start, end), hit in sorted(found.items()):
        # Prefer the complete two-author or et al. token over its bare prefix.
        if any(lo == start and hi > end for lo, hi in found):
            continue
        refs = hit['references']
        issues = ['missing_year']
        if len(refs) > 1:
            issues.append('reference_ambiguous')
        if structure_issue:
            issues.append('protected_structure')
        result.append({'start': start, 'end': end, 'expected_text': text[start:end],
                       'author_start': start, 'author_end': end, 'year': None,
                       'year_start': None, 'year_end': None,
                       'status': 'reference_ambiguous' if len(refs) > 1 else 'missing_year',
                       'issues': issues, 'reference_pids': [r['pid'] for r in refs],
                       'possible_years': sorted({y for r in refs for y in r['years']}),
                       'eligible_structure': structure_issue is None, 'reviewed': False,
                       'notes': ['Author is cited with an existing numbered reference but no adjacent year; verify the full bibliography publication year before a source-bound edit.']})
    return result


def inspect_parts(parts):
    """Inspect already loaded ZIP parts; useful for audit without rereading DOCX."""
    root, styles = C.parse(parts['word/document.xml']), C.parse(parts['word/styles.xml'])
    paragraphs = C.indexed(root)
    contexts = C.reference_contexts(root, styles)
    bibliography = bibliography_contexts(root, styles)
    references = []
    for pid, paragraph in paragraphs.items():
        if bibliography[pid]:
            text = C.ptext(paragraph)
            hint = _guess_reference(text)
            if hint:
                references.append({'pid': pid, 'expected_text': text, **hint, 'reviewed': False})
    candidates, matches = [], []
    for pid, paragraph in paragraphs.items():
        if contexts[pid]:
            continue
        text, issue = C.ptext(paragraph), C.safety_issue(paragraph)
        found = {}
        for reference in references:
            if not reference['authors']:
                continue
            authors = reference['authors']
            first = _name_pattern(authors[0])
            full = SEPARATOR.join(_name_pattern(a) for a in authors)
            short = first + (r'\s*(?:等人|等)' if authors[0]['language'] == 'zh' else r'(?:\s*(?:等人|等)|\s+et\s+al\.)')
            # A bare first author is a warning only, never an approved rewrite.
            author_pattern = '(?:' + '|'.join((full, short, first)) + ')'
            pattern = author_pattern + r'\s*(?:[（(]\s*)?(\d{4}[a-z]?)(?:\s*[）)])?'
            for match in re.finditer(pattern, text, re.I if authors[0]['language'] == 'en' else 0):
                if authors[0]['language'] == 'en' and match.start() and re.match(
                        r"[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-]", text[match.start()-1]):
                    continue
                year = match.group(1)
                if year not in reference['years']:
                    continue
                key = match.start(), match.end()
                found.setdefault(key, {'match': match, 'references': []})['references'].append(reference)
        for (start, end), hit in sorted(found.items()):
            refs = hit['references']
            match = hit['match']
            issues = []
            # Limit hints to this sentence. Claims are still explicitly reviewed.
            sentence_end = next((i for i in range(end, len(text)) if text[i] in END_PUNCTUATION and not (text[i]=='.' and i>0 and i+1<len(text) and text[i-1].isdigit() and text[i+1].isdigit())), len(text))
            # A later author-year starts another cited claim even when joined by a comma.
            next_author=min([a for (a,b) in found if a>=end]+[sentence_end])
            if text[end:next_author].strip() not in ('和','与','、','及','and'):sentence_end=min(sentence_end,next_author)
            numeric = list(NUMERIC_MARKER.finditer(text[end:sentence_end]))
            post_punctuation = re.match('[' + re.escape(END_PUNCTUATION) + r'\s]+(' + NUMERIC_MARKER.pattern + ')', text[sentence_end:])
            numeric_tokens = [m.group() for m in numeric]
            if post_punctuation:
                numeric_tokens.append(post_punctuation.group(1))
            if not numeric_tokens:
                issues.append('missing_numeric_marker')
            elif (post_punctuation or len(numeric_tokens) != 1
                  or not re.fullmatch(r'\[\d+(?:,\d+)*\]', numeric_tokens[0])
                  or not {int(n) for n in re.findall(r'\d+',numeric_tokens[0])}.issubset({r['reference_number'] for r in references})
                  or not any(reference['reference_number'] in {int(n) for n in re.findall(r'\d+',numeric_tokens[0])} for reference in refs)):
                issues.append('numeric_marker_needs_review')
            source_author = text[start:match.start(1)].rstrip().rstrip('（(').rstrip()
            for reference in refs:
                authors, count = reference['authors'], reference['author_count_min']
                if count == 1:
                    correct = source_author == authors[0]['display_name']
                elif count == 2:
                    correct = source_author == EXPECTED_CONNECTOR.join(a['display_name'] for a in authors)
                else:
                    abbreviation = EXPECTED_ABBREVIATIONS[authors[0]['language']]
                    expected_author = authors[0]['display_name'] + (' ' if abbreviation == 'et al.' else '') + abbreviation
                    correct = source_author == expected_author
                if not correct and authors[0]['language'] == 'en':
                    correct = source_author.casefold() == (authors[0]['display_name'] if count == 1 else
                        EXPECTED_CONNECTOR.join(a['display_name'] for a in authors) if count == 2 else
                        authors[0]['display_name'] + ' et al.').casefold()
                if not correct:
                    issues.append('author_count_needs_review')
                    break
            if len(refs) > 1:
                issues.append('reference_ambiguous')
            if issue:
                issues.append('protected_structure')
            priority = ('reference_ambiguous', 'protected_structure', 'numeric_marker_needs_review', 'missing_numeric_marker', 'author_count_needs_review')
            status = next((s for s in priority if s in issues), 'matched_reference')
            item = {'pid': pid, 'start': start, 'end': end, 'expected_text': text[start:end],
                'year': match.group(1), 'status': status, 'issues': issues,
                'author_start': start, 'author_end': start + len(source_author),
                'year_start': match.start(1), 'year_end': match.end(1),
                'reference_pids': [r['pid'] for r in refs], 'eligible_structure': issue is None,
                'reviewed': False, 'notes': [issue] if issue else ['Low-confidence author/year hint; manually review the claim and full bibliography evidence.']}
            matches.append(item)
            if issues:
                candidates.append(item)
        for item in _missing_year_mentions(text, references, found, issue):
            item['pid'] = pid
            matches.append(item)
            candidates.append(item)
    return {'schema_version': 1, 'references': references, 'candidates': candidates, 'matches': matches,
            'expected_conventions': {'two_author_connector': EXPECTED_CONNECTOR,
                                     'three_or_more_author_abbreviations': dict(EXPECTED_ABBREVIATIONS)},
            'notes': ['Inspection is advisory, never an approved mapping or an automatic citation assignment.',
                      'Reviewed targets: two authors use 和 in both languages; three or more use Chinese 等 or foreign et al. (not 等人/与/and).',
                      'Offsets are paragraph Unicode code points, start-inclusive/end-exclusive.',
                      'Only listed bibliography names and real three-name truncation labels may be used as evidence.',
                      'Named citations with an existing numeric marker but no adjacent year are review candidates; no publication year is chosen automatically.',
                      'Unmatched or unsupported author syntax can be absent from hints; inspect the complete document manually.']}


def inspect(source):
    result = inspect_parts(C.package(source))
    result.update(source=str(Path(source).resolve()), source_sha256=C.sha(Path(source).read_bytes()))
    return result


def verify(source, output, mapping):
    expected, actual = expected_parts(source, mapping), C.package(output)
    errors = []
    if set(expected) != set(actual):
        errors.append('ZIP member set changed')
    for name in sorted(set(expected) & set(actual)):
        if expected[name] != actual[name]:
            errors.append('Part differs from deterministic approved transform: ' + name)
    return {'ok': not errors, 'errors': errors, 'source_sha256': C.sha(Path(source).read_bytes()),
            'output_sha256': C.sha(Path(output).read_bytes()), 'selected_records': len(mapping['records']),
            'comparison': 'every package part byte-for-byte against regeneration from source+reviewed map',
            'allowed_text_changes': 'evidenced author display and missing confirmed-claim numeric marker only',
            'style_id': C.STYLE_ID, 'style_name': C.STYLE_NAME}


def apply(source, output, mapping):
    src, dst = Path(source).resolve(), Path(output).resolve()
    if src == dst or dst.exists():
        raise ValueError('Output must be a new path; source and existing outputs are preserved')
    expected = expected_parts(src, mapping)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix='.docx', dir=dst.parent, delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temporary_path, 'w') as zout:
            zout.comment = zin.comment
            for info in zin.infolist():
                zout.writestr(info, expected[info.filename])
        result = verify(src, temporary_path, mapping)
        if not result['ok']:
            raise ValueError('Verification failed: ' + '; '.join(result['errors']))
        with temporary_path.open('rb') as reader, dst.open('xb') as writer:
            while chunk := reader.read(1024 * 1024):
                writer.write(chunk)
        result['output'] = str(dst)
        return result
    finally:
        temporary_path.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subcommands = parser.add_subparsers(dest='command', required=True)
    scan = subcommands.add_parser('inspect')
    scan.add_argument('source')
    scan.add_argument('--out', required=True)
    for command in ('apply', 'verify'):
        command_parser = subcommands.add_parser(command)
        command_parser.add_argument('--source', required=True)
        command_parser.add_argument('--output', required=True)
        command_parser.add_argument('--map', required=True)
        command_parser.add_argument('--report')
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect(args.source)
            C.write_json(args.out, result)
            print(json.dumps({'ok': True, 'manifest': args.out, 'candidates': len(result['candidates'])}, ensure_ascii=False))
            return 0
        mapping = C.read_json(args.map)
        result = (apply if args.command == 'apply' else verify)(args.source, args.output, mapping)
        if args.report:
            C.write_json(args.report, result)
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result['ok'] else 2
    except (OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, E.XMLSyntaxError) as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
