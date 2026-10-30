"""Insert only explicitly reviewed, uniquely evidenced bibliography years.

No network, year inference, author replacement or citation renumbering. Records
are bound to whole source paragraphs and exact author/reference spans. Use before
any text normalization or author-display changes in the formatting pipeline.
"""
from __future__ import annotations

import copy
import re

from lxml import etree as E
import citation_style as C
import narrative_citations as N

DOC = 'word/document.xml'
REQUIRED = {'pid', 'expected_text', 'author_start', 'author_end', 'expected_author',
            'reference_pid', 'reference_expected_text', 'year', 'reviewed'}
ALLOWED = REQUIRED | {'reason'}
PROTECTED_ROLES = {'reference', 'reference_heading', 'appendix_code', 'equation',
                   'table_text', 'caption', 'table_caption', 'figure_caption'}


def _plans(parts, records):
    if not isinstance(records, list):
        raise ValueError('year_insertions must be a reviewed record list')
    root = C.parse(parts[DOC])
    paragraphs = C.indexed(root)
    evidence = N.inspect_parts(parts)
    refs = {r['pid']: r for r in evidence['references']}
    plans, seen = {}, set()
    for record in records:
        if not isinstance(record, dict) or not REQUIRED <= set(record) or set(record) - ALLOWED:
            raise ValueError('Year insertion requires exact source paragraph, author span and reference evidence')
        if record['reviewed'] is not True:
            raise ValueError('Year insertion requires an explicitly reviewed record')
        pid, refpid = record['pid'], record['reference_pid']
        if pid not in paragraphs or refpid not in refs:
            raise ValueError('Year insertion source paragraph or bibliography entry is missing')
        paragraph, ref = paragraphs[pid], refs[refpid]
        before = C.ptext(paragraph)
        if before != record['expected_text'] or ref['expected_text'] != record['reference_expected_text']:
            raise ValueError('Year insertion source or reference full text mismatch')
        if C.safety_issue(paragraph):
            raise ValueError('Year insertion cannot edit or rely on a protected/non-plain paragraph')
        start, end = record['author_start'], record['author_end']
        if type(start) is not int or type(end) is not int or not 0 <= start < end <= len(before):
            raise ValueError('Year insertion author offsets are invalid')
        if before[start:end] != record['expected_author']:
            raise ValueError('Year insertion author span mismatch')
        if (pid, end) in seen:
            raise ValueError('Duplicate year insertion at the same author')
        seen.add((pid, end))
        matches = [c for c in evidence['candidates'] if c['pid'] == pid and
                   c.get('author_start') == start and c.get('author_end') == end]
        if len(matches) != 1 or set(matches[0]['issues']) != {'missing_year'} or \
                not matches[0].get('eligible_structure') or matches[0]['reference_pids'] != [refpid]:
            raise ValueError('Author/reference must uniquely match an eligible missing-year candidate')
        year = record['year']
        # _guess_reference finds every year-like token after the author list.
        # More than one occurrence is ambiguous even when the digits repeat.
        if not isinstance(year, str) or not re.fullmatch(r'\d{4}[a-z]?', year) or \
                ref.get('issues') or ref.get('years') != [year]:
            raise ValueError('Publication year must be the single existing year in the exact bibliography entry')
        reftext = ref['expected_text']
        author_end = ref['author_list']['end']
        year_match = next(N.YEAR.finditer(reftext[author_end:]))
        year_start = author_end + year_match.start()
        year_end = author_end + year_match.end()
        # A DOI link AFTER the publication date is immutable evidence, not an
        # edited target. Require the author/date prefix itself to be plain
        # native runs; unrelated trailing links must not block a proven year.
        prefix=''
        for child in paragraphs[refpid]:
            if child.tag==C.q('pPr'):continue
            if child.tag!=C.q('r') or any(n.tag not in {C.q('rPr'),C.q('t'),C.q('lastRenderedPageBreak')} for n in child):break
            prefix+=C.run_text(child)
        if len(prefix)<year_end or prefix!=reftext[:len(prefix)]:
            raise ValueError('Publication author/year evidence intersects a protected reference structure')
        if (year_start and reftext[year_start-1] not in ' ,，(（\t') or \
                (year_end < len(reftext) and reftext[year_end] not in ' ,，:：;.。()（）\t'):
            raise ValueError('The year is not in a supported bibliography publication-date position')
        # A four-digit component in a DOI, URL or identifier is not a date.
        token_start = max(reftext.rfind(' ', 0, year_start), reftext.rfind('\t', 0, year_start)) + 1
        if re.search(r'(?:https?://|doi\s*:|/)', reftext[token_start:year_end], re.I):
            raise ValueError('A DOI or URL component cannot serve as publication-year evidence')
        sentence_end = next((i for i in range(end, len(before))
                             if before[i] in N.END_PUNCTUATION and not (
                                 before[i] == '.' and i > 0 and i + 1 < len(before)
                                 and before[i-1].isdigit() and before[i+1].isdigit())), len(before))
        markers = list(N.NUMERIC_MARKER.finditer(before[end:sentence_end]))
        if len(markers) != 1 or markers[0].group() != '[' + str(ref['reference_number']) + ']':
            raise ValueError('The cited sentence must have exactly one existing, matching numeric marker')
        # A duplicated reference number is itself ambiguous regardless of name.
        if sum(r['reference_number'] == ref['reference_number'] for r in refs.values()) != 1:
            raise ValueError('The bibliography has duplicate reference numbers')
        plans.setdefault(pid, []).append(copy.deepcopy(record))
    return plans


def _after(before, records):
    for record in sorted(records, key=lambda r: r['author_end'], reverse=True):
        end = record['author_end']
        before = before[:end] + '（' + record['year'] + '）' + before[end:]
    return before


def apply(parts, root, roles, records, replace_span, operations, issues):
    """Apply one atomic reviewed batch; append source-bound operations or issue."""
    if not records:
        return
    try:
        plans = _plans(parts, records)
        paragraphs = C.indexed(root)
        # Validate every current target before performing any mutation.
        for pid, selections in plans.items():
            paragraph = paragraphs.get(pid)
            if paragraph is None or C.ptext(paragraph) != selections[0]['expected_text'] or C.safety_issue(paragraph):
                raise ValueError('Current paragraph changed before year insertion: ' + pid)
            if roles.get(paragraph) in PROTECTED_ROLES:
                raise ValueError('Year insertion targets a protected semantic role: ' + pid)
        replacements = []
        for pid, selections in plans.items():
            original = paragraphs[pid]
            clone = copy.deepcopy(original)
            before = C.ptext(original)
            for record in sorted(selections, key=lambda r: r['author_end'], reverse=True):
                run = E.Element(C.q('r'))
                rp = E.SubElement(run, C.q('rPr'))
                E.SubElement(rp, C.q('vertAlign'), {C.q('val'): 'baseline'})
                E.SubElement(rp, C.q('position'), {C.q('val'): '0'})
                E.SubElement(run, C.q('t')).text = '（' + record['year'] + '）'
                replace_span(clone, record['author_end'], record['author_end'], [run])
            expected = _after(before, selections)
            if C.ptext(clone) != expected:
                raise ValueError('Year insertion changed text outside the requested spans')
            replacements.append((original, clone, {'kind': 'narrative_year_insertion',
                'locator': pid, 'before': before, 'after': expected, 'records': selections}))
        # Keep paragraph identity so semantic role maps remain valid.
        for original, clone, operation in replacements:
            for child in list(original):
                original.remove(child)
            for child in list(clone):
                original.append(child)
            operations.append(operation)
    except (ValueError, KeyError, TypeError) as exc:
        issues.append({'code': 'narrative_year_insertion_unverified', 'location': 'year_insertions',
                       'severity': 'warning', 'message': str(exc)})


def verified(operation, sourceparts):
    """Independently bind logged changes to original bibliography and prose."""
    try:
        if operation.get('kind') != 'narrative_year_insertion':
            return False
        plans = _plans(sourceparts, operation.get('records'))
        pid = operation.get('locator')
        if list(plans) != [pid]:
            return False
        before = C.ptext(C.indexed(C.parse(sourceparts[DOC]))[pid])
        return operation.get('before') == before and operation.get('after') == _after(before, plans[pid])
    except (ValueError, KeyError, TypeError):
        return False
