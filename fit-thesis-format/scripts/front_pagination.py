"""Source-preserving front-page typography and mutually exclusive page starts.

Roles must already be reviewed.  This module never invents an English title,
translates an abstract, removes a disclosure, or infers a title from proximity.
"""
import thesis_format as F


FRONT_ROLES = {'cn_title', 'cn_abstract_heading', 'cn_abstract', 'cn_keywords',
               'en_title', 'en_abstract_heading', 'en_abstract', 'en_keywords',
               'toc_heading'}


def _role(value):
    if isinstance(value, str): return value
    if isinstance(value, dict): return value.get('role')
    return value[0] if value else None


def _front(root, roles):
    body = root.find(F.q('body')); result = []
    for p in body:
        role = _role(roles.get(p))
        if role == 'chapter': break
        if p.tag == F.q('p'): result.append((p, role))
    return result


def _english_groups(front):
    """A reviewed title belongs only to the next adjacent abstract block."""
    result = []
    boundaries = {'cn_abstract_heading', 'cn_keywords', 'en_keywords',
                  'toc_heading', 'en_abstract_heading'}
    for index, (heading, role) in enumerate(front):
        if role != 'en_abstract_heading': continue
        title = None
        for p, previous_role in reversed(front[:index]):
            if previous_role == 'en_title': title = p; break
            if previous_role in boundaries: break
        result.append((title, heading))
    return result


def _set_page_break(p, value):
    pp = F.ensure(p, 'pPr', True)
    F.remove_names(pp, {'pageBreakBefore'})
    element = F.child(pp, 'pageBreakBefore', val='1' if value else '0')
    # pageBreakBefore precedes paragraph-mark rPr/sectPr in the OOXML schema.
    pp.remove(element)
    position = next((i for i, child in enumerate(pp)
                     if child.tag not in {F.q(n) for n in ('pStyle', 'keepNext', 'keepLines')}), len(pp))
    pp.insert(position, element)


def front_page_anchors(root, roles):
    """Expected managed anchors for the confirmed front paragraphs."""
    ids = {p: pid for pid, p in F.indexed(root)[0].items()}; counts = {}; result = []
    for p, role in _front(root, roles):
        if role not in FRONT_ROLES: continue
        counts[role] = counts.get(role, 0) + 1
        suffix = '' if counts[role] == 1 else '_' + str(counts[role])
        result.append({'name': 'FITFront_' + role + suffix, 'role': role, 'pid': ids[p]})
    return result


def expected_front_page_breaks(root, roles):
    """Only the confirmed English-block/contents dynamic paragraph overrides."""
    ids = {p: pid for pid, p in F.indexed(root)[0].items()}; front = _front(root, roles); expected = {}
    for title, heading in _english_groups(front):
        expected[ids[title if title is not None else heading]] = True
        if title is not None: expected[ids[heading]] = False
    for p, role in front:
        if role == 'toc_heading': expected[ids[p]] = True
    return expected


def _ensure_page_anchors(root, roles):
    paragraphs = F.indexed(root)[0]; added = []
    existing = {}
    for element in root.findall('.//w:bookmarkStart', F.NS):
        name = element.get(F.q('name')); existing.setdefault(name, []).append(element)
    numeric_ids = [int(x.get(F.q('id'))) for x in root.findall('.//w:bookmarkStart', F.NS)
                   if (x.get(F.q('id')) or '').isdigit()]
    next_id = max(numeric_ids + [0]) + 1
    for anchor in front_page_anchors(root, roles):
        p = paragraphs[anchor['pid']]; found = existing.get(anchor['name'], [])
        if found:
            owners = [next((a for a in e.iterancestors() if a.tag == F.q('p')), None) for e in found]
            if len(found) != 1 or owners != [p]:
                raise ValueError('Managed front bookmark is duplicated or points elsewhere: ' + anchor['name'])
            marker_id = found[0].get(F.q('id'))
            if not any(e.get(F.q('id')) == marker_id for e in p.findall('.//w:bookmarkEnd', F.NS)):
                raise ValueError('Managed front bookmark is not closed in its paragraph: ' + anchor['name'])
            continue
        start = F.child(p, 'bookmarkStart', id=next_id); start.set(F.q('name'), anchor['name'])
        p.remove(start); position = 1 if p.find(F.q('pPr')) is not None else 0; p.insert(position, start)
        F.child(p, 'bookmarkEnd', id=next_id)
        added.append({**anchor, 'bookmark_id': str(next_id)}); next_id += 1
    return added


def _issue(code, pid, actual, expected, kind='format', severity='warning'):
    return {'code': code, 'pid': pid, 'severity': severity, 'kind': kind,
            'actual': actual, 'expected': expected,
            'blocks_formatting': kind != 'missing_fact',
            'requires_user_fact': kind == 'missing_fact'}


def apply_front_pagination(root, roles, styles, operations=None):
    """Format confirmed front roles, page-start the English block and contents.

    Returns missing_facts separately.  A missing title must not block moving
    the existing Abstract to its own page; no text is added to fill that gap.
    Call after make_styles and role confirmation, before user layout overrides.
    """
    paragraphs = F.indexed(root)[0]; ids = {p: pid for pid, p in paragraphs.items()}
    before = [F.ptext(p) for p in paragraphs.values()]
    front = _front(root, roles)
    style_index = {s.get(F.q('styleId')): s for s in styles.findall(F.q('style'))}
    for p, role in front:
        if role not in FRONT_ROLES: continue
        if F.style_id(role) not in style_index:
            raise ValueError('Required front style is missing: ' + F.style_id(role))
        # Gallery styles alone do not defeat direct 7pt/Calibri run formatting.
        F.format_p(p, role, style_index)
    starts, missing = [], []
    for title, heading in _english_groups(front):
        anchor = title if title is not None else heading
        _set_page_break(anchor, True)
        if title is not None: _set_page_break(heading, False)
        else:
            missing.append(_issue('missing_english_title', ids[heading],
                                  'Existing Abstract has no confirmed English title',
                                  'Use the original verified English title; do not translate or invent it',
                                  kind='missing_fact'))
        starts.append({'block': 'english_abstract', 'anchor_pid': ids[anchor],
                       'anchor_role': 'en_title' if title is not None else 'en_abstract_heading',
                       'heading_pid': ids[heading]})
    for p, role in front:
        if role == 'toc_heading':
            _set_page_break(p, True)
            starts.append({'block': 'contents', 'anchor_pid': ids[p], 'anchor_role': role})
    anchors_added = _ensure_page_anchors(root, roles)
    assert [F.ptext(p) for p in paragraphs.values()] == before
    result = {'page_starts': starts, 'missing_facts': missing,
              'anchors': front_page_anchors(root, roles), 'anchors_added': anchors_added,
              'text_unchanged': True, 'actual_page_review_required': True}
    if operations is not None:
        operations.append({'kind': 'front_pagination', **result})
    return result


def _explicit_break_before(p):
    previous = p.getprevious()
    while previous is not None and previous.tag == F.q('p'):
        content = []
        for node in previous.iter():
            if node.tag == F.q('br') and node.get(F.q('type')) == 'page': content = ['page_break']
            elif node.tag == F.q('t') and (node.text or '').strip(): content.append('text')
        if content: return content[-1] == 'page_break'
        if previous.find('w:pPr/w:sectPr', F.NS) is not None: return False
        previous = previous.getprevious()
    return False


def audit_front_pagination(audit):
    """Read-only static checks using actual style/direct-format resolution.

    Returns issue dictionaries for Audit.add; it does not claim rendered pages
    are correct.  Use audit_front_pages with bound native page evidence too.
    """
    roles = {audit.ps[pid]: value for pid, value in audit.roles.items()}
    front = _front(audit.doc, roles); ids = audit.pid; issues = []
    def page_start(p):
        value = audit.effective_p(p).get('pageBreakBefore', {}).get('val', '0')
        return value not in ('0', 'false', 'off') or audit.section_break_before(p) or _explicit_break_before(p)
    for title, heading in _english_groups(front):
        anchor = title if title is not None else heading
        if title is None:
            issues.append(_issue('front.missing_english_title', ids[heading],
                                 'Abstract without a confirmed English title',
                                 'An existing verified English title', kind='missing_fact'))
        if not page_start(anchor):
            issues.append(_issue('front.abstract_page_boundary', ids[anchor],
                                 'No effective page start on the English block', 'English block starts on a new page'))
        if title is not None and page_start(heading):
            issues.append(_issue('front.english_title_abstract_split', ids[heading],
                                 'Abstract starts another page after its title', 'Title and Abstract share the English page'))
    for p, role in front:
        if role == 'toc_heading' and not page_start(p):
            issues.append(_issue('front.contents_page_boundary', ids[p],
                                 'Contents may continue on the abstract page', 'Contents starts on a new page'))
        if role not in FRONT_ROLES: continue
        spec = F.resolved_style_specs()[role]
        for run in p.iter(F.q('r')):
            if next((a for a in run.iterancestors() if a.tag == F.q('p')), None) is not p: continue
            if not F.ptext(run).strip(): continue
            if any(a.tag in F.PROTECTED_TAGS for a in run.iterancestors()): continue
            effective = audit.effective_r(p, run)
            actual_size = effective.get('sz', {}).get('val')
            expected_size = str(round(spec['size_pt'] * 2))
            if actual_size != expected_size:
                issues.append(_issue('front.effective_font_size', ids[p],
                                     {'role': role, 'half_points': actual_size},
                                     {'half_points': expected_size, 'points': spec['size_pt']}))
                break
    return issues


def audit_front_pages(root, roles, page_by_pid):
    """Check actual physical pages already bound to the same final DOCX.

    Keys are indexed paragraph IDs, values are physical page integers.  Do not
    populate this from guessed text offsets or displayed Roman/Arabic numbers.
    """
    front = _front(root, roles); ids = {p: pid for pid, p in F.indexed(root)[0].items()}
    groups = _english_groups(front); issues = []
    by_role = {}
    for p, role in front: by_role.setdefault(role, []).append(p)
    def page(p):
        result = page_by_pid.get(ids[p])
        return result if type(result) is int and result > 0 else None
    for title, heading in groups:
        anchor = title if title is not None else heading
        chinese_end = (by_role.get('cn_keywords') or by_role.get('cn_abstract') or by_role.get('cn_abstract_heading') or [None])[-1]
        required = [anchor, heading] + ([chinese_end] if chinese_end is not None else [])
        if any(page(p) is None for p in required):
            issues.append(_issue('front.native_page_evidence_missing', ids[anchor],
                                 'A front paragraph has no native physical-page evidence',
                                 'Native page positions for the Chinese and English blocks'))
            continue
        if chinese_end is not None and page(anchor) <= page(chinese_end):
            issues.append(_issue('front.abstracts_share_page', ids[anchor],
                                 {'english': page(anchor), 'chinese_end': page(chinese_end)},
                                 'English abstract begins after the Chinese abstract page', severity='error'))
        if title is not None and page(title) != page(heading):
            issues.append(_issue('front.english_title_abstract_split', ids[heading],
                                 {'title': page(title), 'abstract': page(heading)},
                                 'English title and Abstract are on the same physical page', severity='error'))
    for heading in by_role.get('toc_heading', []):
        previous = (by_role.get('en_keywords') or by_role.get('en_abstract') or by_role.get('en_abstract_heading') or [None])[-1]
        if previous is None: continue
        if page(previous) is None or page(heading) is None:
            issues.append(_issue('front.native_page_evidence_missing', ids[heading],
                                 'No native contents/English-page boundary evidence',
                                 'Native page positions for the English block and contents'))
        elif page(heading) <= page(previous):
            issues.append(_issue('front.contents_share_abstract_page', ids[heading],
                                 {'contents': page(heading), 'english_end': page(previous)},
                                 'Contents begins after the English abstract page', severity='error'))
    return issues
