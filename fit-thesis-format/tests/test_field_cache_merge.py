from pathlib import Path
import sys,tempfile,unittest,zipfile
from lxml import etree as E
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import field_cache_merge as F

class CacheMergeTests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def field(self,code,result):
        return '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>'+code+'</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'+result+'<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    def text(self,text):return '<w:r><w:t>'+text+'</w:t></w:r>'
    def package(self,name,body,media=b'original media',styles=None):
        p=self.root/name
        xml=f'<w:document xmlns:w="{F.W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>{body}</w:body></w:document>'.encode()
        with zipfile.ZipFile(p,'w') as z:
            z.writestr('word/document.xml',xml);z.writestr('word/media/image1.png',media)
            if styles is not None:z.writestr('word/styles.xml',styles)
        return p
    def fixtures(self,ambiguous=False,missing_visible=False):
        oldtoc='<w:p>'+self.field('TOC \\o "1-2"',self.text('old'))+'</w:p>'
        newtoc='<w:p>'+self.field('TOC \\o "1-2"',self.field('PAGEREF _TocNew',self.text('3')))+'</w:p>'
        source=oldtoc+'<w:p>'+self.text('Heading')+'</w:p>'
        if ambiguous:source+='<w:p>'+self.text('Heading')+'</w:p>'
        source+='<w:p>'+self.field('DATE',self.text('2000'))+'</w:p>'
        source+='<w:p>'+self.field('SEQ FITEquation \\r 0 \\h',self.text('0'))+'</w:p>'
        source+='<w:p>'+self.field('SEQ FITTable',self.text('1'))+'</w:p>'
        source+='<w:p><m:oMath><m:r><m:t>−x′</m:t></m:r></m:oMath></w:p>'
        native=newtoc+'<w:p><w:bookmarkStart w:id="20" w:name="_TocNew"/>'+self.text('Heading')+'<w:bookmarkEnd w:id="20"/></w:p>'
        native+='<w:p>'+self.field('DATE',self.text('2026'))+'</w:p>'
        if not missing_visible:native+='<w:p>'+self.field('SEQ FITTable',self.text('2'))+'</w:p>'
        native+='<w:p><m:oMath><m:r><m:t>-x\'</m:t></m:r></m:oMath></w:p>'
        return self.package('source.docx',source),self.package('native.docx',native,b'recompressed media')
    def test_only_caches_merged_original_objects_and_date_survive(self):
        source,native=self.fixtures();before=F.sha(source);out=self.root/'merged.docx'
        report=F.merge_field_caches(source,native,out)
        self.assertEqual(F.sha(source),before)
        a,b=F.package(source),F.package(out)
        self.assertEqual(a['word/media/image1.png'],b['word/media/image1.png'])
        root=F.parse(b['word/document.xml'])
        self.assertIn('−x′',E.tostring(root,encoding='unicode'))
        self.assertIn('2000',E.tostring(root,encoding='unicode'))
        self.assertNotIn('2026',E.tostring(root,encoding='unicode'))
        self.assertEqual(report['retained_hidden_fields'],['SEQ FITEquation \\r 0 \\h'])
        self.assertEqual(len(root.findall('.//w:bookmarkStart',F.NS)),1)
        self.assertEqual(report['cache_changes'][-1]['after'],'2')
        again=self.root/'again.docx';F.merge_field_caches(out,native,again)
        self.assertEqual(F.package(out),F.package(again))
    def test_missing_visible_field_fails(self):
        source,native=self.fixtures(missing_visible=True)
        with self.assertRaisesRegex(ValueError,'omitted visible layout field'):F.merge_field_caches(source,native,self.root/'out.docx')
    def test_ambiguous_heading_fails(self):
        source,native=self.fixtures(ambiguous=True)
        with self.assertRaisesRegex(ValueError,'Ambiguous TOC heading'):F.merge_field_caches(source,native,self.root/'out.docx')
    def test_refuse_existing_output(self):
        source,native=self.fixtures()
        with self.assertRaisesRegex(ValueError,'new file'):F.merge_field_caches(source,native,source)
    def test_cache_sanitizer_preserves_required_extension_attributes(self):
        p=E.fromstring(f'<w:p xmlns:w="{F.W}" xmlns:w14="{F.W14}" w14:paraId="12345678"><w:pPr><w:rPr><w14:ligatures w14:val="standardContextual"/></w:rPr></w:pPr></w:p>')
        F.sanitize_cache([p],{},set())
        self.assertIsNone(p.get('{'+F.W14+'}paraId'))
        self.assertEqual(p.find('.//{'+F.W14+'}ligatures').get('{'+F.W14+'}val'),'standardContextual')

    def test_toc_links_remain_clickable_without_source_hyperlink_decoration(self):
        import thesis_format as T
        styles=E.Element(F.q('styles'));T.make_styles(styles)
        p=E.fromstring(f'<w:p xmlns:w="{F.W}"><w:pPr><w:pStyle w:val="FIT_toc1"/></w:pPr><w:hyperlink w:anchor="_TocHeading"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Heading</w:t></w:r></w:hyperlink></w:p>')
        F.restore_toc_format([p],{'word/styles.xml':E.tostring(styles)})
        link=p.find(F.q('hyperlink'));self.assertEqual(link.get(F.q('anchor')),'_TocHeading')
        rp=link.find('w:r/w:rPr',F.NS)
        self.assertEqual(rp.find(F.q('u')).get(F.q('val')),'none')
        self.assertEqual(rp.find(F.q('color')).get(F.q('val')),'000000')
        self.assertEqual(F.ptext(p),'Heading')

    def test_regenerated_builtin_toc_uses_owned_styles_without_body_overrides(self):
        import thesis_format as T
        styles=E.fromstring(f'<w:styles xmlns:w="{F.W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/><w:ind w:firstLine="480" w:firstLineChars="200"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="old1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="old2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/></w:style></w:styles>')
        T.make_styles(styles)
        source_styles=E.tostring(styles)
        native_styles=E.fromstring(source_styles)
        for st in native_styles.findall(F.q('style')):
            if st.get(F.q('styleId')) in ('old1','old2'):
                st.set(F.q('styleId'),'TOC'+st.get(F.q('styleId'))[-1])
        native_styles=E.tostring(native_styles)
        oldtoc='<w:p>'+self.field('TOC \\o "1-2"',self.text('old'))+'</w:p>'
        newtoc=self.field('TOC \\o "1-2"',
            self.text('Heading')+'<w:r><w:tab/></w:r>'+self.field('PAGEREF _TocNew',self.text('3')))
        # Split the outer field result into two realistic TOC paragraphs.
        newtoc=newtoc.rsplit('<w:r><w:fldChar w:fldCharType="end"/></w:r>',1)
        body=''
        for level,result in ((1,newtoc[0]),(2,self.text('Subheading')+'<w:r><w:tab/></w:r>'+self.text('4')+'<w:r><w:fldChar w:fldCharType="end"/></w:r>')):
            body+=f'<w:p><w:pPr><w:pStyle w:val="TOC{level}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs><w:jc w:val="both"/><w:ind w:firstLine="480" w:firstLineChars="200"/><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:pPr>'+result+'</w:p>'
        source=self.package('source.docx',oldtoc+'<w:p>'+self.text('Heading')+'</w:p>',styles=source_styles)
        native=self.package('native.docx',body+'<w:p><w:bookmarkStart w:id="20" w:name="_TocNew"/>'+self.text('Heading')+'<w:bookmarkEnd w:id="20"/></w:p>',styles=native_styles)
        out=self.root/'merged.docx';F.merge_field_caches(source,native,out)
        parts=F.package(out);root=F.parse(parts['word/document.xml']);ps=root.find(F.q('body')).findall(F.q('p'))
        self.assertEqual(parts['word/styles.xml'],source_styles)
        for level,p in enumerate(ps[:2],1):
            pp=p.find(F.q('pPr'));self.assertEqual(pp.find(F.q('pStyle')).get(F.q('val')),'FIT_toc'+str(level))
            self.assertIsNone(pp.find(F.q('jc')));self.assertIsNone(pp.find(F.q('ind')))
            self.assertIsNone(pp.find('w:rPr/w:rFonts',F.NS));self.assertIsNone(pp.find('w:rPr/w:sz',F.NS))
            self.assertEqual(len(pp.findall(F.q('tabs'))),1)
            tab=pp.find('w:tabs/w:tab',F.NS);self.assertEqual(tab.get(F.q('leader')),'dot');self.assertEqual(tab.get(F.q('pos')),'9345')
            st=styles.find('w:style[@w:styleId="FIT_toc'+str(level)+'"]',F.NS)
            self.assertEqual(st.find('w:pPr/w:jc',F.NS).get(F.q('val')),'left')
            self.assertEqual(st.find('w:pPr/w:ind',F.NS).get(F.q('firstLineChars')),'0')
        self.assertEqual([F.ptext(p) for p in ps],['Heading3','Subheading4','Heading'])
        self.assertTrue(root.findall('.//w:bookmarkStart',F.NS))
        self.assertTrue(any(f['kind']=='PAGEREF' and f['result']=='3' for f in F.spans(root)))
        again=self.root/'again.docx';F.merge_field_caches(out,native,again)
        self.assertEqual(F.package(out),F.package(again))

if __name__=='__main__':unittest.main()
