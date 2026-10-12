"""Citation consistency integration through the portable adapter and delivery chain.

All text is synthetic. Tests exercise review boundaries between tools rather
than importing the narrative-citation unit tests or changing real manuscripts.
"""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from docx import Document

REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / 'fit-thesis-format' / 'scripts'
sys.path.insert(0, str(SCRIPTS))
import citation_style as C
import delivery_check as DELIVERY


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def save_json(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def span(text, start, end):
    return {'start': start, 'end': end, 'expected_text': text[start:end]}


class NarrativeIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='fit 引用集成 ')
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name).resolve()
        self.source = self.work / '原稿 中文.docx'
        self.body = [
            '张甲等（2020）指出测试论述甲。',
            '李乙等（2021）指出测试论述乙。',
            '黄丙、赵丁、钱戊（2022）指出测试论述丙。',
        ]
        self.authors = [['张甲'], ['李乙', '孙己'], ['黄丙', '赵丁', '钱戊']]
        self.references = [
            '[1] 张甲.单作者测试[J].合成期刊,2020,1:1-2.',
            '[2] 李乙,孙己.双作者测试[J].合成期刊,2021,1:1-2.',
            '[3] 黄丙,赵丁,钱戊.三作者测试[J].合成期刊,2022,1:1-2.',
        ]
        document = Document()
        for text in ['引用集成测试', *self.body, '本段研究文字与数据123保持不变。',
                     '参考文献', *self.references, '致谢', '合成致谢文字保持不变。']:
            document.add_paragraph(text)
        document.save(self.source)
        self.source_hash = digest(self.source)
        self.addCleanup(lambda: self.assertEqual(digest(self.source), self.source_hash))
        self.mapping = self.work / '逐条审核.json'
        records = []
        for index, (body, reference, names) in enumerate(zip(self.body, self.references, self.authors)):
            author_end = body.index('（')
            year_start = author_end + 1
            year_end = body.index('）')
            ay = span(body, 0, year_end + 1)
            ay['author'] = span(body, 0, author_end)
            ay['year'] = span(body, year_start, year_end)
            ref_author_start = reference.index(']') + 2
            ref_author_end = reference.index('.')
            ref_year_start = reference.index(body[year_start:year_end])
            author_records = []
            cursor = ref_author_start
            for name in names:
                start = reference.index(name, cursor)
                author_records.append({'start': start, 'end': start + len(name),
                                       'exact_name': name, 'display_name': name, 'language': 'zh'})
                cursor = start + len(name)
            record = {'pid': f'p{index + 2:06d}', 'author_year': ay,
                      'claim': span(body, 0, len(body) - 1),
                      'reference': {'pid': f'p{index + 7:06d}', 'expected_text': reference,
                                    'reference_number': index + 1,
                                    'author_list': span(reference, ref_author_start, ref_author_end),
                                    'authors': author_records,
                                    'year': span(reference, ref_year_start, ref_year_start + 4)}}
            if len(names) == 2:
                record['connector'] = '和'
            elif len(names) >= 3:
                record['abbreviation'] = '等'
            records.append(record)
        save_json(self.mapping, {'schema_version': 1, 'reviewed': True,
                                 'source_sha256': self.source_hash, 'records': records})

    def call(self, request):
        run = subprocess.run([sys.executable, str(SCRIPTS / 'agent_cli.py'), '--request', '-'],
                             input=json.dumps(request, ensure_ascii=False), cwd=self.work,
                             capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(run.stderr, '', run.stderr)
        result = json.loads(run.stdout)
        self.assertEqual(run.returncode, 0 if result['ok'] else 2, run.stdout)
        self.assertEqual(result['operation_succeeded'], result['ok'])
        self.assertFalse(result['delivery_ready'], 'Static adapter work must not imply whole-thesis acceptance')
        return result

    def applied(self):
        output = self.work / '修改稿 中文.docx'
        result = self.call({'action': 'apply_narrative_citations', 'source': str(self.source),
                            'output': str(output), 'mapping': str(self.mapping),
                            'report': str(self.work / '修正记录.json')})
        self.assertTrue(result['ok'], result)
        self.assertEqual(Path(result['outputs']['output']), output)
        self.assertTrue(result['checks']['input_files_unchanged']['source'])
        self.assertTrue(result['checks']['input_files_unchanged']['mapping'])
        return output, result

    def chain(self, final):
        path = self.work / '完整链.json'
        save_json(path, {'original_sha256': self.source_hash, 'stages': [
            {'scope': 'narrative_citations', 'source': str(self.source),
             'output': str(final), 'mapping': str(self.mapping),
             # A stored success assertion has no authority over fresh replay.
             'ok': True, 'report': 'untrusted-old-report.json'}]})
        return path

    def test_json_inspect_is_advisory_and_finds_missing_marks(self):
        report = self.work / '作者年份检查.json'
        result = self.call({'action': 'inspect', 'scope': 'narrative_citations',
                            'source': str(self.source), 'report': str(report)})
        self.assertTrue(result['ok'], result)
        inventory = load_json(result['outputs']['report'])
        self.assertEqual(inventory['source_sha256'], self.source_hash)
        self.assertEqual({item['pid'] for item in inventory['references']},
                         {'p000007', 'p000008', 'p000009'})
        missing = [item for item in inventory['candidates'] if 'missing_numeric_marker' in item['issues']]
        self.assertEqual({item['pid'] for item in missing}, {'p000002', 'p000003', 'p000004'})
        self.assertTrue(all(item['reviewed'] is False for item in inventory['candidates']))
        self.assertFalse((self.work / '修改稿 中文.docx').exists())

    def test_json_apply_and_verify_publish_only_reviewed_author_and_marker_changes(self):
        output, applied = self.applied()
        result = self.call({'action': 'verify', 'scope': 'narrative_citations',
                            'source': str(self.source), 'output': str(output),
                            'mapping': str(self.mapping), 'report': str(self.work / '复核记录.json')})
        self.assertTrue(result['ok'], result)
        self.assertTrue(load_json(result['outputs']['report'])['ok'])
        self.assertEqual(load_json(applied['outputs']['report'])['selected_records'], 3)
        paragraphs = C.indexed(C.parse(C.package(output)['word/document.xml']))
        self.assertEqual([C.ptext(paragraphs[f'p{i:06d}']) for i in range(2, 5)], [
            '张甲（2020）指出测试论述甲[1]。',
            '李乙和孙己（2021）指出测试论述乙[2]。',
            '黄丙等（2022）指出测试论述丙[3]。',
        ])
        self.assertEqual([C.ptext(paragraphs[f'p{i:06d}']) for i in range(7, 10)], self.references)
        self.assertEqual(C.ptext(paragraphs['p000005']), '本段研究文字与数据123保持不变。')

    def test_prepare_creates_bound_inventory_and_unapproved_draft(self):
        prepared = self.work / '准备目录'
        result = self.call({'action': 'prepare', 'source': str(self.source), 'work_dir': str(prepared)})
        self.assertTrue(result['ok'], result)
        self.assertEqual(result['status'], 'awaiting_mapping_review')
        self.assertEqual(digest(prepared / 'original.docx'), self.source_hash)
        inventory = load_json(prepared / 'narrative-citation-inventory.json')
        self.assertEqual(inventory['source_sha256'], self.source_hash)
        self.assertTrue(inventory['candidates'])
        draft = prepared / 'narrative-citation-map.draft.json'
        mapping = load_json(draft)
        self.assertEqual(mapping['source_sha256'], self.source_hash)
        self.assertIs(mapping['reviewed'], False)
        self.assertEqual(mapping['records'], [])
        output = self.work / '未经审核.docx'
        rejected = self.call({'action': 'apply_narrative_citations',
                              'source': str(prepared / 'original.docx'), 'output': str(output),
                              'mapping': str(draft), 'report': str(self.work / '未经审核.json')})
        self.assertFalse(rejected['ok'])
        self.assertFalse(output.exists())

    def test_delivery_chain_accepts_fresh_replay_of_reviewed_citation_stage(self):
        output, _ = self.applied()
        stages = DELIVERY.verify_chain(self.source, output, self.chain(output))
        self.assertEqual(len(stages), 1)
        self.assertEqual(stages[0]['scope'], 'narrative_citations')
        self.assertEqual(stages[0]['source_sha256'], self.source_hash)
        self.assertEqual(stages[0]['output_sha256'], digest(output))
        self.assertEqual(stages[0]['mapping_sha256'], digest(self.mapping))

    def test_delivery_chain_rejects_unapproved_prose_and_bibliography_changes(self):
        output, _ = self.applied()
        chain = self.chain(output)
        clean_parts = C.package(output)
        for pid in ('p000005', 'p000007'):
            with self.subTest(changed_paragraph=pid):
                parts = dict(clean_parts)
                root = C.parse(parts['word/document.xml'])
                paragraph = C.indexed(root)[pid]
                first_text = next(paragraph.iter(C.q('t')))
                first_text.text += '未授权改动'
                parts['word/document.xml'] = C.serialize(root)
                with zipfile.ZipFile(output, 'w') as archive:
                    for name, data in parts.items():
                        archive.writestr(name, data)
                with self.assertRaisesRegex(ValueError, 'fresh verifier'):
                    DELIVERY.verify_chain(self.source, output, chain)


if __name__ == '__main__':
    unittest.main()
