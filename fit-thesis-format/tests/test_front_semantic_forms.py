"""Numeric school typography must survive restoration of the native grid."""
import sys,unittest
from pathlib import Path
from lxml import etree as E
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import thesis_format as F,format_engine as G

class FrontSemanticFormsTests(unittest.TestCase):
    def fixture(self,title=None):
        parts,_=G._front_template();root=F.parse(parts[G.DOC]);styles=F.parse(parts['word/styles.xml'])
        body=root.find(F.q('body'))
        if title is not None:
            rows=body[6].findall(F.q('tr'))
            for ri,value in enumerate((title,'')):
                p=rows[ri].findall(F.q('tc'))[1].find(F.q('p'))
                for node in p.findall('.//w:t',F.NS):node.text=''
                if value:G.plain_run(p,value)
        return parts,root,styles

    def test_signature_uses_real_right_alignment_and_twelve_point_text(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'));table=body[10]
        geometry=[F.xml(x) for x in [table.find(F.q('tblGrid'))]+table.findall('w:tr/w:trPr',F.NS)]
        result=G.restore_school_front_geometry(parts,root,styles)
        self.assertTrue(any(e['kind']=='front_signature_padding' for e in result['edits']))
        signatures=[p for p in table.findall('.//w:p',F.NS) if '签名' in G.text(p) or '2025年' in G.text(p)]
        self.assertEqual(len(signatures),4)
        for p in signatures:
            self.assertEqual(G.text(p),G.text(p).lstrip(' \u3000\t'))
            self.assertEqual(p.find('w:pPr/w:jc',F.NS).get(F.q('val')),'right')
            for r in p.findall(F.q('r')):
                if not G.text(r):continue
                self.assertEqual(r.find('w:rPr/w:sz',F.NS).get(F.q('val')),'24')
                self.assertEqual(r.find('w:rPr/w:rFonts',F.NS).get(F.q('eastAsia')),'宋体')
        self.assertEqual(geometry,[F.xml(x) for x in [table.find(F.q('tblGrid'))]+table.findall('w:tr/w:trPr',F.NS)])
        self.assertEqual(G.front_template_geometry_issues(root,styles),[])
        # A centered signature with the same grid is no longer silently clean.
        signatures[-1].find('w:pPr/w:jc',F.NS).set(F.q('val'),'center')
        self.assertTrue(G.front_template_geometry_issues(root,styles))

    def test_heading_label_and_declaration_keep_distinct_school_roles(self):
        parts,root,styles=self.fixture();G.restore_school_front_geometry(parts,root,styles)
        body=root.find(F.q('body'));heading=body[9]
        self.assertEqual(heading.find('w:r/w:rPr/w:sz',F.NS).get(F.q('val')),'32')
        self.assertEqual(heading.find('w:r/w:rPr/w:b',F.NS).get(F.q('val')),'1')
        for p in body[10].findall('.//w:p',F.NS):
            if G.text(p) in ('学生承诺','指导教师承诺'):
                self.assertEqual(p.find('w:r/w:rPr/w:sz',F.NS).get(F.q('val')),'28')
            if G.text(p).startswith('我承诺'):
                spacing=p.find('w:pPr/w:spacing',F.NS)
                self.assertEqual(spacing.get(F.q('line')),'360')
                self.assertEqual(spacing.get(F.q('lineRule')),'auto')
                self.assertEqual(p.find('w:pPr/w:ind',F.NS).get(F.q('firstLineChars')),'200')
                self.assertEqual(p.find('w:r/w:rPr/w:sz',F.NS).get(F.q('val')),'24')
        declaration=next(p for p in body[10].findall('.//w:p',F.NS) if G.text(p).startswith('我承诺'))
        declaration.find('w:r/w:rPr/w:sz',F.NS).set(F.q('val'),'21')
        self.assertTrue(any('typography' in x for x in G.front_template_geometry_issues(root,styles)))

    def test_long_title_uses_two_balanced_ruled_lines_and_replays(self):
        title='数字化转型对制造业企业投资效率的影响'
        parts,root,styles=self.fixture(title);body=root.find(F.q('body'))
        research=[F.xml(n) for n in list(body)[14:]]
        result=G.restore_school_front_geometry(parts,root,styles)
        texts=[G.text(row.findall(F.q('tc'))[1]) for row in body[6].findall(F.q('tr'))[:2]]
        self.assertEqual(''.join(texts),title);self.assertTrue(all(len(t)>1 for t in texts))
        self.assertLessEqual(abs(len(texts[0])-len(texts[1])),2)
        self.assertTrue(any(e['kind']=='front_title_wrap' for e in result['edits']))
        for row in body[6].findall(F.q('tr'))[:2]:
            for r in row.findall(F.q('tc'))[1].findall('.//w:r',F.NS):
                if G.text(r):self.assertEqual(r.find('w:rPr/w:sz',F.NS).get(F.q('val')),'32')
        first=(F.xml(root),F.xml(styles));G.restore_school_front_geometry(parts,root,styles)
        self.assertEqual(first,(F.xml(root),F.xml(styles)))
        self.assertEqual(research,[F.xml(n) for n in list(body)[14:]])

    def test_object_bound_title_is_preserved_for_review(self):
        parts,root,styles=self.fixture('超长标题'*20);body=root.find(F.q('body'))
        p=body[6].findall(F.q('tr'))[0].findall(F.q('tc'))[1].find(F.q('p'))
        bm=F.child(p,'bookmarkStart',id='987');bm.set(F.q('name'),'ProtectedTitle');end=F.child(p,'bookmarkEnd',id='987')
        before=G.text(p);result=G.restore_school_front_geometry(parts,root,styles)
        self.assertEqual(G.text(p),before);self.assertIs(bm.getparent(),p);self.assertIs(end.getparent(),p)
        self.assertTrue(any(e['kind']=='front_title_wrap_review' for e in result['edits']))

    def test_existing_unique_english_title_fills_pledge_without_translation(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'))
        title='Digital Transformation and Manufacturing Investment Efficiency'
        p=E.Element(F.q('p'));G.plain_run(p,title);body.insert(14,p)
        roles={p:'en_title'};result=G.restore_school_front_geometry(parts,root,styles,roles=roles)
        target=body[10].findall(F.q('tr'))[3].findall(F.q('tc'))[1]
        self.assertEqual(G.text(target),'外文：'+title)
        self.assertEqual(G.text(p),title)
        self.assertTrue(any(e.get('source')=='en_title' for e in result['edits']))
        first=(F.xml(root),F.xml(styles));G.restore_school_front_geometry(parts,root,styles,roles=roles)
        self.assertEqual(first,(F.xml(root),F.xml(styles)))

    def test_ambiguous_english_titles_do_not_fill_pledge(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'));roles={}
        for value in ('First English Title','Another English Title'):
            p=E.Element(F.q('p'));G.plain_run(p,value);body.insert(14,p);roles[p]='en_title'
        G.restore_school_front_geometry(parts,root,styles,roles=roles)
        self.assertEqual(G.text(body[10].findall(F.q('tr'))[3].findall(F.q('tc'))[1]),'外文：')

    def test_imported_front_styles_remain_editable_in_gallery(self):
        parts,root,styles=self.fixture();G.restore_school_front_geometry(parts,root,styles)
        imported=[s for s in styles.findall(F.q('style')) if s.get(F.q('styleId'),'').startswith('FITFrontTpl_')]
        self.assertTrue(imported)
        for s in imported:
            self.assertIsNotNone(s.find(F.q('qFormat')))
            self.assertIsNone(s.find(F.q('semiHidden')));self.assertIsNone(s.find(F.q('locked')))
            self.assertEqual(s.find(F.q('unhideWhenUsed')).get(F.q('val')),'1')

if __name__=='__main__':unittest.main()
