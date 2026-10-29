"""Comment-only cleanup of the retained template; no manuscript edits."""
from pathlib import Path
import sys
import tempfile
import unittest

from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'fit-thesis-format' / 'scripts'))
import template_hygiene as H


class TemplateHygieneTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='fit-comments-')
        self.addCleanup(temporary.cleanup)
        self.work = Path(temporary.name)
        self.source = ROOT / 'fit-thesis-format' / 'assets' / 'fit-economics-management-template.docx'

    def test_real_template_is_cleaned_without_changing_source_or_other_parts(self):
        original = self.source.read_bytes()
        output, report = self.work / 'clean.docx', self.work / 'report.json'
        result = H.clean_template(self.source, output, report)
        second = self.work / 'independent-clean.docx'
        second_result = H.clean_template(self.source, second)
        self.assertEqual(output.read_bytes(), second.read_bytes())
        self.assertEqual(result['output_sha256'], second_result['output_sha256'])
        self.assertEqual(self.source.read_bytes(), original)
        self.assertEqual(result['before']['comment_count'], 83)
        self.assertEqual(len(result['before']['template_comment_matches']), 83)
        self.assertEqual(result['after']['comment_count'], 0)
        self.assertEqual(result['after']['comment_parts'], [])
        self.assertEqual(result['after']['comment_relationships'], [])
        self.assertEqual(result['after']['comment_content_types'], [])
        self.assertFalse(any(result['after']['anchors'].values()))
        self.assertEqual(result['before']['revisions'], result['after']['revisions'])
        before, after = H.read_package(self.source), H.read_package(output)
        changed = set(result['changes']['changed_parts'])
        removed = set(result['changes']['removed_parts'])
        self.assertEqual(set(before) - set(after), removed)
        for name in set(before) - changed - removed:
            self.assertEqual(before[name], after[name], name)
        # Ordinary text, math, drawings, fields, and paragraph formatting survive.
        for name in changed:
            if not name.startswith('word/') or not name.endswith('.xml'):
                continue
            expected = H.parse(before[name])
            for node in tuple(expected.iter()):
                if H.word_name(node) in H.ANCHOR_NAMES:
                    node.getparent().remove(node)
            self.assertEqual(E.tostring(expected, method='c14n'),
                             E.tostring(H.parse(after[name]), method='c14n'), name)
        with self.assertRaisesRegex(ValueError, 'new path'):
            H.clean_template(self.source, output)
        with self.assertRaisesRegex(ValueError, 'pinned anonymized school template'):
            H.clean_template(output, self.work / 'second.docx')
        cleaned_again, changes = H.strip_comment_parts(after)
        self.assertEqual(cleaned_again, after)
        self.assertFalse(changes['removed_parts'])
        self.assertFalse(changes['changed_parts'])

    def test_packaged_clean_template_contains_no_comment_residue(self):
        clean = self.source.with_name('fit-economics-management-template-clean.docx')
        self.assertTrue(clean.is_file(), 'Published skill must contain the clean template asset')
        self.assertEqual(H.digest(self.source.read_bytes()), H.TEMPLATE_SHA256)
        result = H.inspect(clean)
        self.assertEqual(result['comment_count'], 0)
        self.assertFalse(any(result['anchors'].values()))
        for key in ('comment_parts', 'comment_relationships', 'comment_content_types',
                    'dangling_comment_relationships', 'dangling_comment_content_types',
                    'orphan_anchor_ids'):
            self.assertEqual(result[key], [], key)
        self.assertFalse(any(result['revisions'].values()))
        expected = self.work / 'rebuilt.docx'
        H.clean_template(self.source, expected)
        # Runtime zlib versions may produce different compressed bytes. The
        # checked-in asset must contain exactly the same uncompressed parts.
        self.assertEqual(H.read_package(clean), H.read_package(expected))

    def synthetic(self):
        document = f'''<w:document xmlns:w="{H.W}"><w:body><w:p><w:commentRangeStart w:id="4"/>
          <w:r><w:commentReference w:id="4"/><w:drawing/><w:instrText>PAGE</w:instrText><w:t>Keep</w:t></w:r>
          <w:commentRangeEnd w:id="4"/><w:ins w:id="8"><w:r><w:t>Inserted</w:t></w:r></w:ins>
          <w:del w:id="9"><w:r><w:delText>Deleted</w:delText></w:r></w:del></w:p>
          <w:tbl><w:tblPr><w:tblBorders><w:insideH w:val="single"/></w:tblBorders></w:tblPr></w:tbl>
          </w:body></w:document>'''.encode()
        comments = f'<w:comments xmlns:w="{H.W}"><w:comment w:id="4"><w:p><w:r><w:t>Known template note</w:t></w:r></w:p></w:comment></w:comments>'.encode()
        rels = f'''<Relationships xmlns="{H.PKG}"><Relationship Id="c" Target="comments.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"/><Relationship Id="picture" Target="media/picture.bin" Type="image"/></Relationships>'''.encode()
        types = f'''<Types xmlns="{H.CT}"><Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" PartName="/word/comments.xml"/><Override PartName="/word/document.xml" ContentType="document"/></Types>'''.encode()
        return {'word/document.xml': document, 'word/comments.xml': comments,
                'word/header1.xml': f'<w:hdr xmlns:w="{H.W}"><w:p><w:r><w:commentReference w:id="4"/><w:t>Header</w:t></w:r></w:p></w:hdr>'.encode(),
                'word/_rels/document.xml.rels': rels, '[Content_Types].xml': types,
                'word/_rels/comments.xml.rels': f'<Relationships xmlns="{H.PKG}"><Relationship Id="cPicture" Target="media/picture.bin" Type="image"/></Relationships>'.encode(),
                'word/media/picture.bin': b'unchanged media'}

    def test_exact_qnames_and_comment_reference_in_mixed_run(self):
        parts = self.synthetic()
        inventory = H.inspect_parts(parts, {'Known template note'})
        self.assertEqual(inventory['comment_count'], 1)
        self.assertEqual(len(inventory['template_comment_matches']), 1)
        self.assertEqual(inventory['revisions']['ins'], 1)
        self.assertEqual(inventory['revisions']['del'], 1)
        self.assertEqual(sum(inventory['revisions'].values()), 2)
        self.assertEqual(inventory['anchors']['commentReference'], 2)
        cleaned, changes = H.strip_comment_parts(parts)
        self.assertEqual(H.inspect_parts(cleaned, set())['revisions'], inventory['revisions'])
        root = H.parse(cleaned['word/document.xml'])
        for tag in ('drawing', 'instrText', 'insideH', 'ins', 'del'):
            self.assertEqual(len(list(root.iter('{' + H.W + '}' + tag))), 1, tag)
        self.assertEqual(''.join(root.itertext()).replace(' ', '').replace('\n', ''),
                         ''.join(H.parse(parts['word/document.xml']).itertext()).replace(' ', '').replace('\n', ''))
        self.assertEqual(cleaned['word/media/picture.bin'], parts['word/media/picture.bin'])
        self.assertNotIn('word/_rels/comments.xml.rels', cleaned)
        self.assertEqual(len(H.parse(cleaned['word/_rels/document.xml.rels'])), 1)
        self.assertEqual(len(H.parse(cleaned['[Content_Types].xml'])), 1)
        self.assertEqual(changes['removed_anchors'], 4)

    def test_orphans_and_unknown_comments_are_reported_not_deleted(self):
        parts = self.synthetic()
        before = dict(parts)
        result = H.inspect_parts(parts, {'Other historical note'})
        self.assertEqual(result['template_comment_matches'], [])
        self.assertEqual(result['orphan_anchor_ids'], [])
        del parts['word/comments.xml']
        result = H.inspect_parts(parts, set())
        self.assertEqual(result['orphan_anchor_ids'], ['4'])
        self.assertTrue(result['comment_relationships'][0]['missing_target'])
        self.assertEqual(result['dangling_comment_relationships'], result['comment_relationships'])
        self.assertEqual(parts['word/document.xml'], before['word/document.xml'])

    def test_custom_named_metadata_and_relationships_are_removed(self):
        parts = self.synthetic()
        parts['word/notes/metadata.xml'] = b'<metadata/>'
        parts['word/_rels/document.xml.rels'] = f'''<Relationships xmlns="{H.PKG}"><Relationship Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="notes/metadata.xml" Id="extended"/></Relationships>'''.encode()
        cleaned, _ = H.strip_comment_parts(parts)
        self.assertNotIn('word/notes/metadata.xml', cleaned)
        self.assertEqual(len(H.parse(cleaned['word/_rels/document.xml.rels'])), 0)

    def test_stale_content_types_are_detected_without_parts_or_relationships(self):
        # Regression: removing comments.xml alone leaves an invalid package;
        # attributes may occur in either order and ContentType contains a slash.
        for reversed_attributes in (False, True):
            with self.subTest(reversed_attributes=reversed_attributes):
                parts = self.synthetic()
                del parts['word/comments.xml']
                del parts['word/_rels/document.xml.rels']
                del parts['word/_rels/comments.xml.rels']
                attributes = ['PartName="/word/comments.xml"',
                              'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"']
                if reversed_attributes:
                    attributes.reverse()
                parts['[Content_Types].xml'] = (f'<Types xmlns="{H.CT}"><Override ' +
                    ' '.join(attributes) + '/><Override PartName="/word/document.xml" ContentType="document"/></Types>').encode()
                snapshot = dict(parts)
                inventory = H.inspect_parts(parts, set())
                self.assertEqual(inventory['comment_parts'], [])
                self.assertEqual(inventory['comment_relationships'], [])
                expected = [{'part': 'word/comments.xml',
                             'content_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
                             'missing_target': True}]
                self.assertEqual(inventory['comment_content_types'], expected)
                self.assertEqual(inventory['dangling_comment_content_types'], expected)
                self.assertEqual(parts, snapshot)
                cleaned, changes = H.strip_comment_parts(parts)
                self.assertEqual(changes['removed_content_types'], 1)
                self.assertEqual(H.inspect_parts(cleaned, set())['comment_content_types'], [])
                self.assertEqual(len(H.parse(cleaned['[Content_Types].xml'])), 1)

    def test_cli_rejects_non_template_and_keeps_input(self):
        path = self.work / 'other.docx'
        path.write_bytes(b'not the pinned original')
        with self.assertRaisesRegex(ValueError, 'pinned anonymized school template'):
            H.clean_template(path, self.work / 'output.docx')
        self.assertEqual(path.read_bytes(), b'not the pinned original')
        self.assertFalse((self.work / 'output.docx').exists())
        with self.assertRaisesRegex(ValueError, 'new path'):
            H.clean_template(self.source, self.source)

    def test_nonstandard_mime_stale_metadata_is_detected_by_known_part_name(self):
        # Real residue used openxmlformats MIME names for Microsoft extensions.
        # None of the target parts or relationships remain in this package.
        names = ('comments', 'commentsExtended', 'commentsIds', 'people')
        overrides = ''.join(
            f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{name}+xml"/>'
            for name in names)
        types = (f'<Types xmlns="{H.CT}">' + overrides +
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>').encode()
        parts = {'[Content_Types].xml': types,
                 'word/document.xml': f'<w:document xmlns:w="{H.W}"><w:body/></w:document>'.encode()}
        before = dict(parts)
        inventory = H.inspect_parts(parts, set())
        self.assertEqual(inventory['comment_parts'], [])
        self.assertEqual(inventory['comment_relationships'], [])
        self.assertEqual(len(inventory['comment_content_types']), 4)
        self.assertEqual(len(inventory['dangling_comment_content_types']), 4)
        self.assertEqual({item['part'] for item in inventory['comment_content_types']},
                         {f'word/{name}.xml' for name in names})
        self.assertEqual(parts, before)
        cleaned, changes = H.strip_comment_parts(parts)
        self.assertEqual(changes['removed_content_types'], 4)
        self.assertEqual(len(H.parse(cleaned['[Content_Types].xml'])), 1)
        self.assertEqual(H.inspect_parts(cleaned, set())['comment_content_types'], [])
        self.assertEqual(cleaned['word/document.xml'], parts['word/document.xml'])


if __name__ == '__main__':
    unittest.main()
