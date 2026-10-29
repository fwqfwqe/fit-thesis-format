import copy
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import format_engine as G
import thesis_format as F
from format_audit import Audit


class HeadingRestartAuditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def fixture(self, extra=False):
        entries = [('chapter', '第一章 正文'), ('section', '第一节'),
                   ('section', '第二节'), ('conclusion_heading', '结论与展望'),
                   ('section', '研究结论'), ('section', '研究展望'),
                   ('subsection', '局限'), ('subsubsection', '具体局限')]
        if extra:
            entries += [('appendix_heading', '附录 A'), ('section', '附录小节')]
        d = Document()
        for _, value in entries: d.add_paragraph(value)
        source = self.root / 'source.docx'
        d.save(source)
        parts = F.package(source)
        root = F.parse(parts['word/document.xml'])
        ps = root.find(F.q('body')).findall(F.q('p'))
        roles = {p: role for p, (role, _) in zip(ps, entries)}
        G._numbering(parts, root, roles, [])
        return parts, root, ps, roles

    def audit(self, parts, root, roles):
        path = self.root / 'audit.docx'
        parts = dict(parts); parts['word/document.xml'] = F.xml(root)
        with zipfile.ZipFile(path, 'w') as z:
            for name, value in parts.items(): z.writestr(name, value)
        mapping = {'source_sha256': F.sha(path.read_bytes()), 'paragraphs': {
            pid: {'role': roles[p], 'expected_text': G.text(p)}
            for pid, p in F.indexed(root)[0].items() if p in roles}}
        audit = Audit(path, mapping)
        audit.infer_roles(); audit.check_numbering()
        self.assertEqual(F.sha(path.read_bytes()), mapping['source_sha256'])
        return [i for i in audit.issues if i['category'] == 'heading.disconnected_lists']

    @staticmethod
    def num_node(p):
        return p.find('w:pPr/w:numPr/w:numId', F.NS)

    def test_confirmed_conclusion_and_appendix_restarts_pass(self):
        parts, root, ps, roles = self.fixture(extra=True)
        self.assertEqual(self.audit(parts, root, roles), [])
        self.assertNotEqual(self.num_node(ps[0]).get(F.q('val')), self.num_node(ps[4]).get(F.q('val')))
        self.assertNotEqual(self.num_node(ps[4]).get(F.q('val')), self.num_node(ps[9]).get(F.q('val')))

    def test_real_disconnection_inside_ordinary_chapter_is_reported(self):
        parts, root, ps, roles = self.fixture()
        nr = F.parse(parts['word/numbering.xml'])
        nid = self.num_node(ps[0]).get(F.q('val'))
        original = nr.find(f'w:num[@w:numId="{nid}"]', F.NS)
        cloned = copy.deepcopy(original); cloned.set(F.q('numId'), '999'); nr.append(cloned)
        parts['word/numbering.xml'] = F.xml(nr)
        self.num_node(ps[2]).set(F.q('val'), '999')
        issues = self.audit(parts, root, roles)
        self.assertEqual(len(issues), 1)
        self.assertTrue(any(p.get('reason') == 'ordinary_headings_use_multiple_lists'
                            for p in issues[0]['actual']['violations']))

    def test_inconsistent_lists_within_conclusion_are_reported(self):
        parts, root, ps, roles = self.fixture()
        self.num_node(ps[5]).set(F.q('val'), self.num_node(ps[0]).get(F.q('val')))
        self.assertTrue(self.audit(parts, root, roles))

    def test_restart_definition_cannot_be_reused_by_another_block(self):
        parts, root, ps, roles = self.fixture(extra=True)
        self.num_node(ps[9]).set(F.q('val'), self.num_node(ps[4]).get(F.q('val')))
        issues = self.audit(parts, root, roles)
        self.assertIn('restart_list_reused_by_another_block', issues[0]['actual']['violations'][0]['reasons'])

    def test_same_ordinary_list_does_not_count_as_restart(self):
        parts, root, ps, roles = self.fixture()
        for p in ps[4:]: self.num_node(p).set(F.q('val'), self.num_node(ps[0]).get(F.q('val')))
        issues = self.audit(parts, root, roles)
        self.assertIn('block_reuses_ordinary_list', issues[0]['actual']['violations'][0]['reasons'])

    def test_incomplete_or_wrong_restart_definitions_are_reported(self):
        for mutation in ('missing_level', 'wrong_start', 'unmanaged_abstract'):
            with self.subTest(mutation=mutation):
                parts, root, ps, roles = self.fixture()
                nr = F.parse(parts['word/numbering.xml'])
                nid = self.num_node(ps[4]).get(F.q('val'))
                num = nr.find(f'w:num[@w:numId="{nid}"]', F.NS)
                if mutation == 'missing_level': num.remove(num.find('w:lvlOverride[@w:ilvl="3"]', F.NS))
                elif mutation == 'wrong_start': num.find('w:lvlOverride/w:startOverride', F.NS).set(F.q('val'), '2')
                else:
                    aid = num.find(F.q('abstractNumId')).get(F.q('val'))
                    nr.find(f'w:abstractNum[@w:abstractNumId="{aid}"]/w:name', F.NS).set(F.q('val'), 'UnverifiedList')
                parts['word/numbering.xml'] = F.xml(nr)
                self.assertTrue(self.audit(parts, root, roles))


if __name__ == '__main__':
    unittest.main()
