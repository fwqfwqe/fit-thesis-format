import copy
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import thesis_format as F
import front_pagination as P
from format_audit import Audit


class FrontPaginationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.base = Path(self.temp.name)

    def tearDown(self): self.temp.cleanup()

    def sample(self, title=True):
        d = Document(); values = [('论文中文题目', 'cn_title'), ('摘要', 'cn_abstract_heading'),
            ('这是中文摘要。', 'cn_abstract'), ('关键词：数字化，投资', 'cn_keywords')]
        if title: values.append(('Digital Transformation and Investment Efficiency', 'en_title'))
        values += [('Abstract', 'en_abstract_heading'), ('This is the complete English abstract.', 'en_abstract'),
                   ('Keywords: digital transformation, investment', 'en_keywords'), ('目录', 'toc_heading')]
        for text, _ in values:
            p = d.add_paragraph(text); p.runs[0].font.size = Pt(7); p.runs[0].font.name = 'Arial'
        path = self.base / 'source.docx'; d.save(path); parts = F.package(path)
        root = F.parse(parts['word/document.xml']); styles = F.parse(parts['word/styles.xml'])
        paragraphs = list(F.indexed(root)[0].values()); roles = {p: role for p, (_, role) in zip(paragraphs, values)}
        F.make_styles(styles)
        return parts, root, styles, roles

    def byrole(self, roles, role): return next(p for p, value in roles.items() if value == role)
    def flag(self, p): return p.find('w:pPr/w:pageBreakBefore', F.NS).get(F.q('val'))

    def audit(self, parts, root, styles):
        parts = dict(parts); parts['word/document.xml'] = F.xml(root); parts['word/styles.xml'] = F.xml(styles)
        path = self.base / 'output.docx'
        with zipfile.ZipFile(path, 'w') as z:
            for name, value in parts.items(): z.writestr(name, value)
        a = Audit(path); a.infer_roles(); return a

    def test_title_is_page_anchor_and_abstract_is_not_a_second_break(self):
        parts, root, styles, roles = self.sample()
        h = self.byrole(roles, 'en_abstract_heading'); P._set_page_break(h, True)
        result = P.apply_front_pagination(root, roles, styles)
        self.assertEqual(self.flag(self.byrole(roles, 'en_title')), '1')
        self.assertEqual(self.flag(h), '0')
        self.assertEqual(self.flag(self.byrole(roles, 'toc_heading')), '1')
        self.assertEqual(result['missing_facts'], [])
        self.assertEqual(P.audit_front_pagination(self.audit(parts, root, styles)), [])

    def test_missing_title_still_moves_existing_abstract_without_inventing_content(self):
        parts, root, styles, roles = self.sample(title=False)
        before = F.ptext(root)
        result = P.apply_front_pagination(root, roles, styles)
        self.assertEqual(self.flag(self.byrole(roles, 'en_abstract_heading')), '1')
        self.assertEqual(F.ptext(root), before)
        self.assertEqual(result['missing_facts'][0]['code'], 'missing_english_title')
        self.assertFalse(result['missing_facts'][0]['blocks_formatting'])
        self.assertEqual([i['code'] for i in P.audit_front_pagination(self.audit(parts, root, styles))], ['front.missing_english_title'])

    def test_direct_small_font_is_removed_and_effective_sizes_are_standard(self):
        parts, root, styles, roles = self.sample()
        P.apply_front_pagination(root, roles, styles)
        a = self.audit(parts, root, styles)
        self.assertFalse(any(i['code'] == 'front.effective_font_size' for i in P.audit_front_pagination(a)))
        for pid, entry in a.roles.items():
            if entry[0] == 'en_title':
                r = a.ps[pid].find(F.q('r'))
                self.assertEqual(a.effective_r(a.ps[pid], r)['sz']['val'], '30')
                self.assertEqual(a.font(a.effective_r(a.ps[pid], r), 'western'), 'Times New Roman')

    def test_audit_checks_effective_size_after_named_style_was_applied(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        p = self.byrole(roles, 'en_abstract_heading'); rp = F.ensure(p.find(F.q('r')), 'rPr', True)
        F.child(rp, 'sz', val=14)
        issues = P.audit_front_pagination(self.audit(parts, root, styles))
        self.assertTrue(any(i['code'] == 'front.effective_font_size' for i in issues))

    def test_double_page_break_is_rejected_even_with_correct_style_name(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        P._set_page_break(self.byrole(roles, 'en_abstract_heading'), True)
        self.assertIn('front.english_title_abstract_split', [i['code'] for i in P.audit_front_pagination(self.audit(parts, root, styles))])

    def test_contents_page_break_is_checked(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        P._set_page_break(self.byrole(roles, 'toc_heading'), False)
        self.assertIn('front.contents_page_boundary', [i['code'] for i in P.audit_front_pagination(self.audit(parts, root, styles))])

    def test_dynamic_abstract_anchor_is_a_verified_layout_override(self):
        parts, root, styles, roles = self.sample(title=False); P.apply_front_pagination(root, roles, styles)
        a = self.audit(parts, root, styles); a.check_paragraphs()
        heading_pid = next(pid for pid, entry in a.roles.items() if entry[0] == 'en_abstract_heading')
        bad = [i for i in a.issues if i['category'] == 'style.direct_format_override' and i['location'].get('paragraph_id') == heading_pid]
        self.assertEqual(bad, [])
        self.assertTrue(any(i['category'] == 'style.contextual_layout_override' and i['location'].get('paragraph_id') == heading_pid for i in a.issues))

    def test_wrong_abstract_override_is_not_whitelisted(self):
        parts, root, styles, roles = self.sample(title=False); P.apply_front_pagination(root, roles, styles)
        P._set_page_break(self.byrole(roles, 'en_abstract_heading'), False)
        a = self.audit(parts, root, styles); a.check_paragraphs()
        self.assertTrue(any(i['category'] == 'style.direct_format_override' for i in a.issues))

    def test_physical_same_page_is_rejected_despite_correct_static_page_flags(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        ids = {p: pid for pid, p in F.indexed(root)[0].items()}
        pages = {ids[p]: 1 for p in roles}
        issues = P.audit_front_pages(root, roles, pages)
        self.assertIn('front.abstracts_share_page', [i['code'] for i in issues])
        self.assertIn('front.contents_share_abstract_page', [i['code'] for i in issues])

    def test_physical_title_and_abstract_must_share_page(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        ids = {p: pid for pid, p in F.indexed(root)[0].items()}
        pages = {ids[p]: (1 if role.startswith('cn_') else 3 if role == 'toc_heading' else 2) for p, role in roles.items()}
        self.assertEqual(P.audit_front_pages(root, roles, pages), [])
        pages[ids[self.byrole(roles, 'en_title')]] = 1
        self.assertIn('front.english_title_abstract_split', [i['code'] for i in P.audit_front_pages(root, roles, pages)])

    def test_missing_native_page_positions_are_not_a_pass(self):
        parts, root, styles, roles = self.sample(); P.apply_front_pagination(root, roles, styles)
        self.assertTrue(P.audit_front_pages(root, roles, {}))
        self.assertTrue(all(i['code'] == 'front.native_page_evidence_missing' for i in P.audit_front_pages(root, roles, {})))

    def test_idempotence_and_protected_objects_preserved(self):
        parts, root, styles, roles = self.sample()
        p = self.byrole(roles, 'cn_abstract'); m = F.child(p, 'bookmarkStart', id=71); m.set(F.q('name'), 'original-evidence')
        F.child(p, 'bookmarkEnd', id=71)
        baseline = F.content_skeleton(root)
        P.apply_front_pagination(root, roles, styles); first = F.xml(root)
        P.apply_front_pagination(root, roles, styles)
        self.assertEqual(F.xml(root), first)
        # Added managed page anchors are the sole new content-container nodes.
        unmarked = copy.deepcopy(root)
        owned_ids = set()
        for b in list(unmarked.findall('.//w:bookmarkStart', F.NS)):
            if (b.get(F.q('name')) or '').startswith('FITFront_'):
                owned_ids.add(b.get(F.q('id'))); b.getparent().remove(b)
        for b in list(unmarked.findall('.//w:bookmarkEnd', F.NS)):
            if b.get(F.q('id')) in owned_ids: b.getparent().remove(b)
        self.assertEqual(F.content_skeleton(unmarked), baseline)

    def test_managed_page_anchors_cover_titles_and_keyword_tail_roles_once(self):
        parts, root, styles, roles = self.sample()
        result = P.apply_front_pagination(root, roles, styles)
        names = [a['name'] for a in result['anchors']]
        for role in ('cn_title', 'cn_keywords', 'en_title', 'en_keywords', 'toc_heading'):
            self.assertIn('FITFront_' + role, names)
        self.assertEqual(len(names), len(set(names)))
        first = F.xml(root); again = P.apply_front_pagination(root, roles, styles)
        self.assertEqual(again['anchors_added'], []); self.assertEqual(F.xml(root), first)

    def test_original_bookmark_does_not_get_reused_as_a_managed_anchor(self):
        parts, root, styles, roles = self.sample()
        p = self.byrole(roles, 'en_title'); old = F.child(p, 'bookmarkStart', id=700); old.set(F.q('name'), 'StudentOriginalBookmark')
        F.child(p, 'bookmarkEnd', id=700)
        result = P.apply_front_pagination(root, roles, styles)
        ids = [int(a['bookmark_id']) for a in result['anchors_added']]
        self.assertTrue(all(i > 700 for i in ids))
        self.assertIsNotNone(root.find('.//w:bookmarkStart[@w:name="StudentOriginalBookmark"]', F.NS))

    def test_conflicting_managed_anchor_is_not_silently_repointed(self):
        parts, root, styles, roles = self.sample()
        p = self.byrole(roles, 'cn_abstract'); b = F.child(p, 'bookmarkStart', id=901); b.set(F.q('name'), 'FITFront_en_title'); F.child(p, 'bookmarkEnd', id=901)
        with self.assertRaisesRegex(ValueError, 'points elsewhere'):
            P.apply_front_pagination(root, roles, styles)


if __name__ == '__main__': unittest.main()
