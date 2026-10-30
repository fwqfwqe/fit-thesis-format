import copy,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import thesis_format as F,format_engine as G

class TemplateGeometryTests(unittest.TestCase):
    def fixture(self):
        parts,body=G._front_template();root=F.parse(parts[G.DOC]);styles=F.parse(parts['word/styles.xml'])
        return parts,root,styles

    def test_cover_form_is_never_stretched_to_page_width(self):
        parts,root,styles=self.fixture();cover=root.find(F.q('body'))[6];before=F.xml(cover)
        G._table(cover,{'kind':'cover','header_rows':0},9355)
        self.assertEqual(before,F.xml(cover))

    def test_native_form_geometry_is_restored_and_detected_independently(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'))
        body[6].find('w:tblGrid/w:gridCol',F.NS).set(F.q('w'),'9000')
        p=body[10].find('.//w:p',F.NS);G.setprop(F.ensure(p,'pPr',True),'spacing',line=400,lineRule='exact')
        protected=[F.xml(x) for x in list(body)[14:]]
        self.assertTrue(G.front_template_geometry_issues(root))
        result=G.restore_school_front_geometry(parts,root,styles)
        self.assertTrue(result['restored']);self.assertEqual(G.front_template_geometry_issues(root),[])
        self.assertEqual(protected,[F.xml(x) for x in list(body)[14:]])

    def test_front_repair_is_idempotent(self):
        parts,root,styles=self.fixture();G.restore_school_front_geometry(parts,root,styles)
        first=(F.xml(root),F.xml(styles));G.restore_school_front_geometry(parts,root,styles)
        self.assertEqual(first,(F.xml(root),F.xml(styles)))

    def test_form_default_does_not_inherit_body_space_after(self):
        parts,root,styles=self.fixture()
        defaults=F.ensure(F.ensure(styles,'docDefaults',True),'pPrDefault')
        G.setprop(F.ensure(defaults,'pPr'),'spacing',after=200,line=276,lineRule='auto')
        G.restore_school_front_geometry(parts,root,styles)
        spacing=styles.find('w:style[@w:styleId="FITFrontTpl_1"]/w:pPr/w:spacing',F.NS)
        self.assertEqual(spacing.get(F.q('before')),'0');self.assertEqual(spacing.get(F.q('after')),'0')
        self.assertEqual(spacing.get(F.q('line')),'240');self.assertEqual(spacing.get(F.q('lineRule')),'auto')
        self.assertEqual(G.front_template_geometry_issues(root,styles),[])
        spacing.set(F.q('after'),'200')
        self.assertTrue(G.front_template_geometry_issues(root,styles))

    def test_paragraph_style_precedes_other_paragraph_properties(self):
        parts,root,styles=self.fixture();G.restore_school_front_geometry(parts,root,styles)
        for node in list(root.find(F.q('body')))[:14]:
            ps=[node] if node.tag==F.q('p') else node.findall('.//w:p',F.NS)
            for p in ps:self.assertEqual(p.find(F.q('pPr'))[0].tag,F.q('pStyle'))

    def test_labels_fit_template_grid_without_stretching_form(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'));grid=F.xml(body[6].find(F.q('tblGrid')))
        G.restore_school_front_geometry(parts,root,styles)
        self.assertEqual(grid,F.xml(body[6].find(F.q('tblGrid'))))
        for row in body[6].findall(F.q('tr')):
            margin=row.findall(F.q('tc'))[0].find('w:tcPr/w:tcMar',F.NS)
            self.assertEqual(margin.find(F.q('left')).get(F.q('w')),'0')
            self.assertEqual(margin.find(F.q('right')).get(F.q('w')),'0')

    def test_title_fill_and_unknown_dates_are_not_fabricated(self):
        parts,root,styles=self.fixture();body=root.find(F.q('body'))
        cover=body[6];title=cover.findall(F.q('tr'))[0].findall(F.q('tc'))[1]
        for t in title.findall('.//w:t',F.NS):t.text=''
        G.plain_run(title.find(F.q('p')),'测试论文题目')
        for p in body[10].findall('.//w:p',F.NS):
            if '2025' in G.text(p):
                for t in p.findall('.//w:t',F.NS):t.text=''
                G.plain_run(p,'年月日')
        G.restore_school_front_geometry(parts,root,styles,fill_title=True)
        self.assertEqual(G.text(body[10].findall(F.q('tr'))[2].findall(F.q('tc'))[1]),'中文：测试论文题目')
        self.assertEqual(sum(G.text(p).strip()=='年    月    日' for p in body[10].findall('.//w:p',F.NS)),2)
        first=(F.xml(root),F.xml(styles));G.restore_school_front_geometry(parts,root,styles,fill_title=True)
        self.assertEqual(first,(F.xml(root),F.xml(styles)))

if __name__=='__main__':unittest.main()
