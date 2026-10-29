"""Check distributed template privacy and reproducible comment cleanup."""
import hashlib
import json
from pathlib import Path
import posixpath
import sys
import tempfile
import unittest

from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import template_hygiene as H


class ReleaseAssetsTests(unittest.TestCase):
    def test_anonymous_metadata_and_blank_identity_cells(self):
        for path in (ROOT / 'assets').glob('*.docx'):
            with self.subTest(asset=path.name):
                parts = H.read_package(path)
                self.assertNotIn('word/people.xml', parts)
                properties = E.fromstring(parts['docProps/custom.xml'])
                names = {p.get('name') for p in properties}
                self.assertTrue(names.isdisjoint({'ICV', 'commondata', 'KSOTemplateDocerSaveRecord'}))
                if 'word/comments.xml' in parts:
                    comments = E.fromstring(parts['word/comments.xml'])
                    self.assertEqual({c.get('{' + H.W + '}author') for c in comments}, {'学校模板批注'})
                document = E.fromstring(parts['word/document.xml'])
                first_table = document.find('.//{' + H.W + '}tbl')
                identity_rows = 0
                for row in first_table.findall('{' + H.W + '}tr'):
                    cells = row.findall('{' + H.W + '}tc')
                    label = ''.join(cells[0].itertext()).strip()
                    if label in ('姓名：', '学号：'):
                        self.assertFalse(''.join(cells[1].itertext()).strip())
                        self.assertIsNotNone(cells[1].find('{' + H.W + '}tcPr'))
                        identity_rows += 1
                self.assertEqual(identity_rows, 2)

    def test_source_comment_text_and_pinned_hash(self):
        source = ROOT / 'assets/fit-economics-management-template.docx'
        rules = json.loads((ROOT / 'references/rules-v10.json').read_text(encoding='utf-8'))
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        self.assertEqual(digest, rules['sources']['template']['sha256'])
        self.assertEqual(digest, H.TEMPLATE_SHA256)
        comments = E.fromstring(H.read_package(source)['word/comments.xml'])
        actual = {c.get('{' + H.W + '}id'): ''.join(c.itertext()) for c in comments}
        records = json.loads((ROOT / 'references/source-comments.json').read_text(encoding='utf-8'))
        self.assertEqual(set(actual), {x['id'] for x in records})
        for record in records:
            # JSON uses paragraph separators where OOXML has adjacent w:p nodes.
            self.assertEqual(''.join(actual[record['id']].split()), ''.join(record['text'].split()))

    def test_clean_template_is_exact_reproducible_derivative(self):
        source = ROOT / 'assets/fit-economics-management-template.docx'
        clean = ROOT / 'assets/fit-economics-management-template-clean.docx'
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / 'clean.docx'
            result = H.clean_template(source, output)
            self.assertEqual(output.read_bytes(), clean.read_bytes())
            self.assertEqual(result['after']['comment_count'], 0)
            self.assertFalse(result['after']['comment_parts'])
            self.assertFalse(any(result['after']['anchors'].values()))

    def test_all_internal_relationship_targets_resolve(self):
        for path in (ROOT / 'assets').glob('*.docx'):
            parts = H.read_package(path)
            for name, data in parts.items():
                if not name.endswith('.rels'):
                    continue
                base = '' if name == '_rels/.rels' else name.rsplit('/_rels/', 1)[0]
                for relationship in E.fromstring(data):
                    if relationship.get('TargetMode') == 'External':
                        continue
                    target = relationship.get('Target', '').split('#')[0]
                    resolved = posixpath.normpath(posixpath.join(base, target)).lstrip('/')
                    self.assertIn(resolved, parts, (path.name, name, relationship.get('Id')))


if __name__ == '__main__':
    unittest.main()
