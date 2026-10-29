import sys
import tempfile
import unittest
from pathlib import Path
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import format_engine as G
import thesis_format as F


class HeadingRestartTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def source(self, entries):
        doc = Document()
        for role, value in entries:
            doc.add_paragraph(value)
        source = self.root / 'source.docx'
        doc.save(source)
        parts = F.package(source)
        root = F.parse(parts['word/document.xml'])
        paragraphs = root.find(F.q('body')).findall(F.q('p'))
        roles = {p: role for p, (role, _) in zip(paragraphs, entries)}
        return source, parts, root, roles, paragraphs

    @staticmethod
    def ids(paragraphs):
        return [p.find('w:pPr/w:numPr/w:numId', F.NS).get(F.q('val'))
                if p.find('w:pPr/w:numPr/w:numId', F.NS) is not None else None
                for p in paragraphs]

    def assert_restart(self, parts, num_id):
        nr = F.parse(parts['word/numbering.xml'])
        num = next(n for n in nr.findall(F.q('num')) if n.get(F.q('numId')) == num_id)
        overrides = num.findall(F.q('lvlOverride'))
        self.assertEqual({n.get(F.q('ilvl')): n.find(F.q('startOverride')).get(F.q('val')) for n in overrides},
                         {'1': '1', '2': '1', '3': '1'})

    def test_unnumbered_conclusion_restarts_all_sublevels_and_preserves_chapters(self):
        _, parts, root, roles, ps = self.source([
            ('chapter', '一、样本讨论'), ('section', '（一）第一节'),
            ('section', '（二）第二节'), ('subsection', '1. 细项'),
            ('conclusion_heading', '结论与展望'), ('section', '（六）研究结论'),
            ('section', '（七）研究展望'), ('subsection', '2. 限制'),
            ('subsubsection', '（3）具体限制'), ('chapter', '二、后续正文'),
            ('section', '（一）后续第一节')])
        ops = []
        G._numbering(parts, root, roles, ops)
        ids = self.ids(ps)
        self.assertEqual(ids[:4], [ids[0]] * 4)
        self.assertIsNone(ids[4])
        self.assertNotEqual(ids[0], ids[5])
        self.assertEqual(ids[5:9], [ids[5]] * 4)
        self.assertEqual(ids[9:], [ids[0]] * 2)
        self.assert_restart(parts, ids[5])
        self.assertEqual(G.text(ps[5]), '研究结论')
        self.assertEqual(G.text(ps[6]), '研究展望')
        self.assertEqual(sum(o['kind'] == 'heading_number_restart' for o in ops), 1)
        numbered, document = parts['word/numbering.xml'], F.xml(root)
        G._numbering(parts, root, roles, [])
        self.assertEqual(parts['word/numbering.xml'], numbered)
        self.assertEqual(F.xml(root), document)

    def test_each_unnumbered_block_is_independent_without_empty_restart_definitions(self):
        _, parts, root, roles, ps = self.source([
            ('chapter', '第一章 绪论'), ('section', '第一节'),
            ('conclusion_heading', '结论'), ('section', '结论一'),
            ('reference_heading', '参考文献'), ('reference', '[1] 文献'),
            ('appendix_heading', '附录 A'), ('section', '附录第一节'),
            ('appendix_heading', '附录 B'), ('subsection', '附录细项'),
            ('acknowledgement_heading', '致谢'), ('body', '致谢正文。')])
        G._numbering(parts, root, roles, [])
        ids = self.ids(ps)
        self.assertEqual(len({ids[i] for i in (0, 3, 7, 9)}), 4)
        for index in (3, 7, 9):
            self.assert_restart(parts, ids[index])
        nr = F.parse(parts['word/numbering.xml'])
        owned = [n for n in nr.findall(F.q('num')) if len(n.findall(F.q('lvlOverride'))) == 3]
        self.assertEqual(len(owned), 3)
        numbering = parts['word/numbering.xml']
        G._numbering(parts, root, roles, [])
        self.assertEqual(parts['word/numbering.xml'], numbering)

    def test_input_map_insertion_order_cannot_change_list_sequence(self):
        _, parts, root, roles, ps = self.source([
            ('chapter', '第一章 正文'), ('section', '章节小节'),
            ('conclusion_heading', '结论'), ('section', '结论小节')])
        G._numbering(parts, root, dict(reversed(list(roles.items()))), [])
        ids = self.ids(ps)
        self.assertEqual(ids[0], ids[1])
        self.assertNotEqual(ids[1], ids[3])
        self.assert_restart(parts, ids[3])

    def test_full_formatter_rerun_keeps_native_restart_and_content(self):
        source, _, _, _, _ = self.source([
            ('cn_title', '独立标题测试'), ('cn_abstract_heading', '摘要'),
            ('cn_abstract', '摘要正文。'), ('toc_heading', '目录'), ('chapter', '第一章 正文'),
            ('section', '（一）样本'), ('body', '研究内容 0.123 保持。'),
            ('section', '（二）结果'), ('conclusion_heading', '结论与展望'),
            ('section', '（六）研究结论'), ('section', '（七）研究展望'),
            ('reference_heading', '参考文献'), ('reference', '[1] 张三. 研究[J].期刊,2020.'),
            ('acknowledgement_heading', '致谢'), ('body', '感谢导师。')])
        inventory = G.inspect_document(source)
        values = {'独立标题测试': 'cn_title', '第一章 正文': 'chapter', '（一）样本': 'section',
                  '（二）结果': 'section', '结论与展望': 'conclusion_heading',
                  '（六）研究结论': 'section', '（七）研究展望': 'section',
                  '研究内容 0.123 保持。': 'body', '感谢导师。': 'body'}
        decisions = {'source_sha256': inventory['source_sha256'],
                     'metadata': {'title': '独立标题测试', 'name': '测试用户', 'student_id': '123',
                                  'college': '经济管理学院', 'major': '投资学', 'grade': '2027', 'date': '2027年'},
                     'paragraphs': {
            pid: {'role': values[G.text(p)], 'expected_text': G.text(p)}
            for pid, p in F.indexed(F.parse(F.package(source)['word/document.xml']))[0].items()
            if G.text(p) in values}}
        a, b = self.root / 'a.docx', self.root / 'b.docx'
        G.format_document(source, a, decisions)
        G.format_document(a, b)
        self.assertEqual(F.package(a), F.package(b))
        self.assertTrue(G.verify_format(source, a, decisions)['ok'])
        final_parts = F.package(a)
        final = F.parse(final_parts['word/document.xml'])
        children = final.find(F.q('body')).findall(F.q('p'))
        conclusion = next(p for p in children if G.text(p) == '研究结论')
        outlook = next(p for p in children if G.text(p) == '研究展望')
        section = next(p for p in children if G.text(p) == '样本')
        ids = self.ids([conclusion, outlook, section])
        self.assertEqual(ids[0], ids[1])
        self.assertNotEqual(ids[0], ids[2])
        self.assert_restart(final_parts, ids[0])
        self.assertIn('研究内容 0.123 保持。', G.text(final))


if __name__ == '__main__':
    unittest.main()
