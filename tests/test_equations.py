import os
import tempfile
import unittest
import zipfile
from unittest.mock import patch

from lxml import etree

from src.converters.docx.cleaner import clean_document
from src.converters.docx.converter import FeishuDocxConverter


MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'m': MATH_NS, 'w': WORD_NS}


class EquationExportTests(unittest.TestCase):
    def _render(self, blocks, clean=True):
        workspace = tempfile.TemporaryDirectory()
        self.addCleanup(workspace.cleanup)
        output_path = os.path.join(workspace.name, 'equations.docx')
        converter = FeishuDocxConverter(blocks, None, workspace.name)
        converter.process(output_path)
        if clean:
            clean_document(output_path)
        with zipfile.ZipFile(output_path) as archive:
            return etree.fromstring(archive.read('word/document.xml'))

    def test_inline_and_standalone_formulas_survive_export_and_cleaning(self):
        first = r'\mathcal{B}_{i,j}=\{k:g_i-\Delta<t_{j,k}\le g_i\},\qquad n_{i,j}=|\mathcal{B}_{i,j}|'
        second = r'\bar{x}_{i,j}=\begin{cases}\displaystyle\frac{1}{n_{i,j}}\sum_{k\in\mathcal{B}_{i,j}}x_{j,k},&n_{i,j}>0,\\\mathrm{NaN},&n_{i,j}=0.\end{cases}'
        blocks = [
            {'block_id': 'page', 'block_type': 1, 'children': ['intro', 'first', 'second']},
            {'block_id': 'intro', 'parent_id': 'page', 'block_type': 2, 'text': {
                'elements': [
                    {'text_run': {'content': '间隔为 '}},
                    {'equation': {'content': r'\Delta=1\,\mathrm{h}'}},
                    {'text_run': {'content': '。'}},
                ],
            }},
            {'block_id': 'first', 'parent_id': 'page', 'block_type': 2, 'text': {
                'elements': [{'equation': {'content': first}}], 'style': {'align': 2},
            }},
            {'block_id': 'second', 'parent_id': 'page', 'block_type': 2, 'text': {
                'elements': [{'equation': {'content': second}}], 'style': {'align': 2},
            }},
        ]

        root = self._render(blocks)
        paragraphs = root.xpath('./w:body/w:p', namespaces=NS)
        self.assertEqual(len(paragraphs), 3)
        self.assertEqual([child.tag.rsplit('}', 1)[-1] for child in paragraphs[0] if child.tag != f'{{{WORD_NS}}}pPr'],
                         ['r', 'oMath', 'r'])
        self.assertEqual(len(root.xpath('.//m:oMath', namespaces=NS)), 3)
        self.assertEqual(len(root.xpath('.//m:oMathPara', namespaces=NS)), 2)
        self.assertEqual(len(root.xpath('.//m:f', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('.//m:m', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('.//m:naryPr/m:supHide[@m:val="1"]', namespaces=NS)), 1)
        self.assertIn('NaN', ''.join(root.xpath('.//m:t/text()', namespaces=NS)))
        for paragraph in paragraphs[1:]:
            self.assertEqual(paragraph.xpath('./w:pPr/w:jc/@w:val', namespaces=NS), ['center'])

    def test_unconvertible_equation_keeps_latex_source(self):
        blocks = [
            {'block_id': 'page', 'block_type': 1, 'children': ['formula']},
            {'block_id': 'formula', 'parent_id': 'page', 'block_type': 2,
             'text': {'elements': [{'equation': {'content': r'\unknown{x}'}}]}},
        ]
        with patch('src.converters.docx.converter.latex_to_omml', side_effect=ValueError('unsupported')):
            root = self._render(blocks, clean=False)
        self.assertEqual(root.xpath('.//w:t/text()', namespaces=NS), [r'\(\unknown{x}\)'])


if __name__ == '__main__':
    unittest.main()
