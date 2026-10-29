import sys
import tempfile
import unittest
import zipfile
import shutil
from unittest.mock import patch
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import integrity as I

class PreservationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
    def tearDown(self):
        self.temp.cleanup()
    def doc(self, name, inner, extra=None):
        p = self.root / name
        with zipfile.ZipFile(p, 'w') as z:
            z.writestr('word/document.xml', f'<w:document xmlns:w="{I.W}" xmlns:m="{I.M}"><w:body>{inner}</w:body></w:document>')
            for n, data in (extra or {}).items(): z.writestr(n, data)
        return p
    def check_change(self, old, new):
        a = self.doc('a.docx', old); b = self.doc('b.docx', new)
        return I.verify_transition(a, b, 'office_fields')
    def test_hidden_text_rejected(self):
        a = '<w:p><w:r><w:t>重要结论0.25</w:t></w:r></w:p>'
        b = a.replace('<w:r>', '<w:r><w:rPr><w:vanish/></w:rPr>')
        self.assertFalse(self.check_change(a,b)['passed'])
    def test_superscript_rejected(self):
        a = '<w:p><w:r><w:t>2</w:t></w:r></w:p>'
        b = a.replace('<w:r>', '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>')
        self.assertFalse(self.check_change(a,b)['passed'])
    def test_run_split_allowed(self):
        a = '<w:p><w:r><w:t>中文ABC</w:t></w:r></w:p>'
        b = '<w:p><w:r><w:t>中文</w:t></w:r><w:r><w:t>ABC</w:t></w:r></w:p>'
        self.assertTrue(self.check_change(a,b)['passed'])
    def test_page_cache_allowed_date_not_allowed(self):
        for code, permitted in [('PAGE', True), ('DATE', False)]:
            a = f'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>{code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            self.assertEqual(self.check_change(a,a.replace('<w:t>1','<w:t>2'))['passed'], permitted)
    def test_formula_subscript_rejected(self):
        a = '<w:p><m:oMath><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:oMath></w:p>'
        self.assertFalse(self.check_change(a,a.replace('<m:t>i','<m:t>j'))['passed'])
    def test_original_image_loss_rejected(self):
        a = self.doc('a.docx','<w:p/>',{'word/media/image1.png':b'original'})
        b = self.doc('b.docx','<w:p/>',{'word/media/image1.png':b'changed'})
        self.assertFalse(I.verify_transition(a,b,'office_fields')['passed'])
    def test_link_target_change_rejected(self):
        body = f'<w:p xmlns:r="{I.R}"><w:hyperlink r:id="rId1"><w:r><w:t>链接</w:t></w:r></w:hyperlink></w:p>'
        rel = f'<Relationships xmlns="{I.PKG}"><Relationship Id="rId1" Type="hyperlink" Target="https://example.com/a" TargetMode="External"/></Relationships>'
        a = self.doc('a.docx',body,{'word/_rels/document.xml.rels':rel})
        b = self.doc('b.docx',body,{'word/_rels/document.xml.rels':rel.replace('/a"','/b"')})
        self.assertFalse(I.verify_transition(a,b,'office_fields')['passed'])
    def test_native_minus_substitution_rejected(self):
        a = '<w:p><m:oMath><m:r><m:t>−1</m:t></m:r></m:oMath></w:p>'
        self.assertFalse(self.check_change(a,a.replace('−','-'))['passed'])
    def test_empty_native_math_properties_allowed(self):
        a = '<w:p><m:oMath><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f></m:oMath></w:p>'
        self.assertTrue(self.check_change(a,a.replace('<m:f>','<m:f><m:fPr><m:ctrlPr><w:rPr/></m:ctrlPr></m:fPr>'))['passed'])
    def test_replay_alone_cannot_hide_engine_text_loss(self):
        a = self.doc('a.docx','<w:p><w:r><w:t>第一段</w:t></w:r></w:p><w:p><w:r><w:t>重要第二段</w:t></w:r></w:p>')
        b = self.doc('b.docx','<w:p><w:r><w:t>第一段</w:t></w:r></w:p>')
        report = {'source_mapping':{'paragraphs':{'p000001':{'role':'body'},'p000002':{'role':'body'}}},
                  'final_mapping':{'paragraphs':{'p000001':{'role':'body'}}}}
        # Simulate a deterministic but faulty transformer: replay reproduces
        # the same loss, independent payload comparison must still reject it.
        with patch('format_engine.format_document',side_effect=lambda source, output, **kw: shutil.copyfile(b,output)):
            result=I.verify_transition(a,b,'format',report)
        self.assertFalse(result['passed'])
        self.assertTrue(result['checks']['deterministic_replay'])
        self.assertFalse(result['checks']['research_payloads'])
    def test_author_display_edit_cannot_delete_claim(self):
        source='[1]Greco S, Smith A, Jones B, Wang C. Study[J].2020.'
        authors=[{'exact_name':n,'display_name':n.split()[0],'language':'en'} for n in ['Greco S','Smith A','Jones B','Wang C']]
        author_text='Greco S, Smith A, Jones B, Wang C'
        reference={'source_text':source,'authors':authors,'span':{'start':3,'end':3+len(author_text),'expected_text':author_text}}
        before='Greco等认为该政策改善了投资效率。'
        valid={'kind':'narrative_author_display','before':before,'after':before.replace('Greco等','Greco et al.'),'edit_span':{'start':0,'end':6,'expected_text':'Greco等','replacement':'Greco et al.'},'reference':reference}
        self.assertTrue(I.verified_display_edit(valid,[source]))
        malicious={**valid,'after':'Greco et al.','edit_span':{'start':0,'end':len(before),'expected_text':before,'replacement':'Greco et al.'}}
        self.assertFalse(I.verified_display_edit(malicious,[source]))

if __name__ == '__main__': unittest.main()
