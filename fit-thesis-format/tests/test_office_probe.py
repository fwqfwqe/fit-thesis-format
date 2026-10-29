import importlib.util
import json
import os,subprocess,ctypes
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import office_probe as O


class OfficeProbeTests(unittest.TestCase):
    def test_registered_font_cannot_mask_actual_glyph_fallback(self):
        glyphs=[]
        for text,font in [('宋体中文正文','ABCDEF+SimSun'),('黑体标题检查','SimHei'),('DML-DID','TimesNewRomanPSMT')]:
            glyphs.extend((char,font) for char in text)
        self.assertTrue(O.validate_smoke_fonts(glyphs)['passed'])
        fallback=[(char,'Arial' if font=='TimesNewRomanPSMT' else font) for char,font in glyphs]
        self.assertFalse(O.validate_smoke_fonts(fallback)['passed'])
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
    def tearDown(self):
        self.tmp.cleanup()
    def docx(self,name,xml):
        path=self.root/name
        with zipfile.ZipFile(path,'w') as z:z.writestr('word/document.xml',xml)
        return path
    def document(self,inner):
        return f'<w:document xmlns:w="{O.W}"><w:body><w:p>{inner}</w:p></w:body></w:document>'.encode()
    def field(self,code,cache):
        return '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>'+code+'</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>'+cache+'</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    def test_roman_and_arabic(self):
        self.assertEqual(O._page_label(14,1),'XIV')
        self.assertEqual(O._page_label(4,2),'iv')
        self.assertEqual(O._page_label(14,0),'14')
        self.assertIsNone(O._page_label(14,None))
    def test_toc_nested_cache_compared_to_adjusted_page(self):
        pageref=self.field('PAGEREF _Toc1 \\h','IV')
        toc=self.field('TOC \\o "1-2"','').replace('<w:r><w:t></w:t></w:r>',pageref)
        source=self.docx('source.docx',self.document(toc))
        entries=O.validate_toc(source,{'_Toc1':{'actual_page':4,'physical_page':6,'number_style':1}})
        self.assertEqual(len(entries),1);self.assertTrue(entries[0]['passed'])
        self.assertFalse(O.validate_toc(source,{'_Toc1':{'actual_page':5,'physical_page':7,'number_style':1}})[0]['passed'])
    def test_missing_bookmark_and_number_style_fail(self):
        source=self.docx('source.docx',self.document(self.field('PAGEREF _Toc1','1')))
        self.assertFalse(O.validate_toc(source,{})[0]['passed'])
        self.assertFalse(O.validate_toc(source,{'_Toc1':{'actual_page':1}})[0]['passed'])
    def test_native_table_spans_must_cover_actual_tables(self):
        source=self.docx('source.docx',self.document('<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'))
        self.assertTrue(O.validate_table_spans(source,[{'table_index':1,'start_page':1,'end_page':2}],2))
        with self.assertRaises(RuntimeError):O.validate_table_spans(source,[],2)
        with self.assertRaises(RuntimeError):O.validate_table_spans(source,[{'table_index':1,'start_page':1,'end_page':3}],2)
    def test_restore_date_cache_does_not_restore_layout_cache(self):
        source=self.docx('source.docx',self.document(self.field('DATE','2000-01-01')+self.field('PAGEREF _Toc1','1')))
        saved=self.docx('saved.docx',self.document(self.field('DATE','2026-10-29')+self.field('PAGEREF _Toc1','4')))
        original_hash=O.sha(source)
        changes=O.restore_protected_field_caches(source,saved)
        self.assertEqual(len(changes),1);self.assertEqual(changes[0]['restored_result'],'2000-01-01')
        with zipfile.ZipFile(saved) as z:text=z.read('word/document.xml').decode()
        self.assertIn('2000-01-01',text);self.assertNotIn('2026-10-29',text)
        self.assertEqual(O._parse_cached_pagerefs(saved)[0]['displayed_page'],'4')
        self.assertEqual(O.sha(source),original_hash)
    def test_protected_field_loss_fails(self):
        source=self.docx('source.docx',self.document(self.field('DATE','2000')))
        saved=self.docx('saved.docx',self.document('<w:r><w:t>lost</w:t></w:r>'))
        with self.assertRaisesRegex(RuntimeError,'lost a protected field'):
            O.restore_protected_field_caches(source,saved)
    def test_simple_date_can_be_restored_after_native_complex_conversion(self):
        source=self.docx('source.docx',self.document('<w:fldSimple w:instr="DATE"><w:r><w:t>2000</w:t></w:r></w:fldSimple>'))
        saved=self.docx('saved.docx',self.document(self.field('DATE','2026')))
        changes=O.restore_protected_field_caches(source,saved)
        self.assertEqual(changes[0]['restored_result'],'2000')
        with zipfile.ZipFile(saved) as z:self.assertIn(b'fldSimple',z.read('word/document.xml'))
    def test_complex_protected_field_fails_closed(self):
        code=self.field('INCLUDEPICTURE "x"','old').replace('<w:t>old</w:t>','<w:drawing/><w:t>old</w:t>')
        source=self.docx('source.docx',self.document(code))
        saved=self.docx('saved.docx',self.document(self.field('INCLUDEPICTURE "x"','new')))
        with self.assertRaisesRegex(RuntimeError,'complex content'):
            O.restore_protected_field_caches(source,saved)
    def test_internal_toc_hyperlinks_not_misclassified_as_protected(self):
        def toc(text):
            return self.field('TOC','').replace('<w:r><w:t></w:t></w:r>',self.field('HYPERLINK _Toc1',text))
        source=self.docx('source.docx',self.document(toc('old')))
        saved=self.docx('saved.docx',self.document(toc('updated')))
        self.assertEqual(O.restore_protected_field_caches(source,saved),[])
    def test_readonly_doctor_never_claims_actual_render(self):
        d=O.doctor()
        self.assertFalse(d['dual_render_capable']);self.assertFalse(d['smoke_executed'])
        self.assertTrue(all(not v['export_tested'] for v in d['engines'].values()))
    def test_reject_overwriting_before_native_call(self):
        source=self.docx('source.docx',self.document(''))
        pdf=self.root/'existing.pdf';pdf.write_bytes(b'protected')
        with self.assertRaises(ValueError):O.render('word',source,pdf)
        with self.assertRaises(ValueError):O.render('wps',source,self.root/'new.pdf',source)
        with self.assertRaises(ValueError):O.render('fake',source,self.root/'new.pdf')
        self.assertEqual(pdf.read_bytes(),b'protected')
    def test_timeout_receipt_excludes_preexisting_unbound_and_reused_process(self):
        state={'schema_version':10,'source_sha256':'abc','closed':False,'preexisting_pids':[1],
               'engine_executable':'C:/Office/wps.exe','owned_processes':[
                   {'id':1,'start_filetime_utc':10,'path':'C:/Office/wps.exe'},
                   {'id':2,'start_filetime_utc':10,'path':'C:/Office/wps.exe'},
                   {'id':3,'start_filetime_utc':10,'path':'C:/Other/wps.exe'},
                   {'id':4,'path':'C:/Office/wps.exe'}]}
        self.assertEqual([p['id'] for p in O._owned_receipt_candidates(state,'abc')],[2])
        self.assertEqual(O._owned_receipt_candidates(state,'wrong'),[])
        state['closed']=True
        self.assertEqual(O._owned_receipt_candidates(state,'abc'),[])
    @unittest.skipUnless(os.name=='nt','Windows native handle identity test')
    def test_timeout_cleanup_checks_live_creation_time_before_terminating_owned_helper(self):
        from ctypes import wintypes
        child=subprocess.Popen([sys.executable,'-c','import time; time.sleep(30)'],creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            kernel=ctypes.WinDLL('kernel32',use_last_error=True)
            kernel.GetProcessTimes.argtypes=[wintypes.HANDLE,*([ctypes.POINTER(wintypes.FILETIME)]*4)]
            values=[wintypes.FILETIME() for _ in range(4)]
            self.assertTrue(kernel.GetProcessTimes(int(child._handle),*(ctypes.byref(v) for v in values)))
            stamp=(values[0].dwHighDateTime<<32)|values[0].dwLowDateTime
            path=str(Path(sys.executable).resolve())
            state={'schema_version':10,'source_sha256':'test','closed':False,'preexisting_pids':[],
                   'engine_executable':path,'owned_processes':[{'id':child.pid,'path':path,'start_filetime_utc':stamp+1}]}
            bad=O.cleanup_owned_timeout(state,'test')
            self.assertFalse(bad[0]['terminated']);self.assertIsNone(child.poll())
            state['owned_processes'][0]['start_filetime_utc']=stamp
            good=O.cleanup_owned_timeout(state,'test')
            self.assertTrue(good[0]['terminated']);self.assertIsNotNone(child.wait(timeout=5))
        finally:
            if child.poll() is None:child.kill();child.wait()


if __name__=='__main__':unittest.main()
