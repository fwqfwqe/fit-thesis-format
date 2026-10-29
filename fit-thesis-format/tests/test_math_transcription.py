import sys,unittest
from pathlib import Path
from lxml import etree as E
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import math_transcription as T

class TranscriptionTests(unittest.TestCase):
    def p(self,text):
        p=E.Element(T.q('p'),nsmap={'w':T.W});r=E.SubElement(p,T.q('r'));E.SubElement(r,T.q('t')).text=text;return p
    def test_simple_exact_source_becomes_editable(self):
        p=self.p('y = x + 1\t公式（1-1）');source=E.fromstring(E.tostring(p))
        op=T.transcribe(p)
        self.assertTrue(T.verified_record(op,source))
        self.assertEqual(p.find('.//'+T.m('t')).text,'y = x + 1')
    def test_changed_operator_rejected(self):
        p=self.p('y=x+1公式（1-1）')
        xml=f'<m:oMath xmlns:m="{T.M}"><m:r><m:t>y=x-1</m:t></m:r></m:oMath>'
        with self.assertRaises(ValueError):T.transcribe(p,xml)
    def test_structured_fraction_source_identity(self):
        p=self.p('x/y公式（1-1）');source=E.fromstring(E.tostring(p))
        xml=f'<m:oMath xmlns:m="{T.M}"><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f></m:oMath>'
        self.assertTrue(T.verified_record(T.transcribe(p,xml),source))
    def test_grouping_must_match(self):
        p=self.p('(a+b)/c公式（1-1）')
        xml=f'<m:oMath xmlns:m="{T.M}"><m:r><m:t>a+b/c</m:t></m:r></m:oMath>'
        with self.assertRaises(ValueError):T.transcribe(p,xml)
    def test_protected_object_not_discarded(self):
        p=self.p('y=x+1公式（1-1）');E.SubElement(p,T.q('bookmarkStart'))
        with self.assertRaises(ValueError):T.transcribe(p)
    def test_nonlinear_not_guessed(self):
        with self.assertRaises(ValueError):T.transcribe(self.p('x_i=1公式（1-1）'))

if __name__=='__main__':unittest.main()
