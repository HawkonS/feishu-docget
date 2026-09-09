import os
import tempfile
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.oxml.ns import qn
from PIL import Image

from src.converters.docx.cleaner import clean_document
from src.converters.docx.converter import (
    FeishuDocxConverter,
    extract_table_cell_background,
    extract_table_backgrounds_from_content,
    normalize_table_background_color,
)
from src.converters.docx.style_manager import TableStyleManager


class TableParagraphStyleTests(unittest.TestCase):
    def _save_table_document(self, path, template_path=None):
        document = Document(template_path) if template_path else Document()
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = 'header'
        table.cell(0, 1).text = 'header'
        table.cell(1, 0).text = 'body'
        table.cell(1, 1).text = 'body'
        document.save(path)

    def test_table_paragraphs_use_separate_presets(self):
        with tempfile.TemporaryDirectory() as workspace:
            output_path = os.path.join(workspace, 'output.docx')
            self._save_table_document(output_path)
            clean_document(output_path)

            document = Document(output_path)
            styles = {style.name: style for style in document.styles}
            self.assertEqual(styles['表格正文'].type, WD_STYLE_TYPE.PARAGRAPH)
            self.assertEqual(styles['表格表头'].type, WD_STYLE_TYPE.PARAGRAPH)
            self.assertIsNot(styles['表格正文'], styles['表格表头'])
            self.assertEqual(document.tables[0].cell(0, 0).paragraphs[0].style.name, '表格表头')
            self.assertEqual(document.tables[0].cell(1, 0).paragraphs[0].style.name, '表格正文')

    def test_existing_paragraph_presets_are_reused(self):
        with tempfile.TemporaryDirectory() as workspace:
            template_path = os.path.join(workspace, 'template.docx')
            output_path = os.path.join(workspace, 'output.docx')
            template = Document()
            body_style = template.styles.add_style('表格正文', WD_STYLE_TYPE.PARAGRAPH)
            header_style = template.styles.add_style('表格表头', WD_STYLE_TYPE.PARAGRAPH)
            body_style.font.name = 'Courier New'
            header_style.font.name = 'Arial'
            template.save(template_path)
            self._save_table_document(output_path, template_path=template_path)

            clean_document(output_path, template_path=template_path)

            document = Document(output_path)
            names = [style.name for style in document.styles]
            self.assertEqual(names.count('表格正文'), 1)
            self.assertEqual(names.count('表格表头'), 1)
            self.assertNotIn('表格正文（导出）', names)
            self.assertNotIn('表格表头（导出）', names)

    def test_non_paragraph_name_collision_gets_reusable_suffix(self):
        with tempfile.TemporaryDirectory() as workspace:
            template_path = os.path.join(workspace, 'template.docx')
            output_path = os.path.join(workspace, 'output.docx')
            template = Document()
            template.styles.add_style('表格正文', WD_STYLE_TYPE.TABLE)
            template.styles.add_style('表格表头', WD_STYLE_TYPE.TABLE)
            template.save(template_path)
            self._save_table_document(output_path, template_path=template_path)

            clean_document(output_path, template_path=template_path)

            document = Document(output_path)
            names = [style.name for style in document.styles]
            self.assertEqual(names.count('表格正文'), 1)
            self.assertEqual(names.count('表格表头'), 1)
            self.assertEqual(names.count('表格正文（导出）'), 1)
            self.assertEqual(names.count('表格表头（导出）'), 1)
            self.assertEqual(
                document.tables[0].cell(1, 0).paragraphs[0].style.name,
                '表格正文（导出）',
            )
            self.assertEqual(
                document.tables[0].cell(0, 0).paragraphs[0].style.name,
                '表格表头（导出）',
            )

    def test_image_paragraph_uses_independent_preset(self):
        with tempfile.TemporaryDirectory() as workspace:
            image_path = os.path.join(workspace, 'image.png')
            output_path = os.path.join(workspace, 'output.docx')
            Image.new('RGB', (12, 12), color='red').save(image_path)

            document = Document()
            image_paragraph = document.add_paragraph()
            image_paragraph.add_run().add_picture(image_path)
            document.add_paragraph('caption')
            document.save(output_path)

            clean_document(
                output_path,
                body_style={
                    'fontSize': 30,
                    'lineSpacing': 2,
                    'spaceBefore': 4,
                    'spaceBeforeUnit': 'lines',
                },
            )

            document = Document(output_path)
            self.assertEqual(document.paragraphs[0].style.name, '图片')
            self.assertEqual(document.paragraphs[1].style.name, 'Normal')
            self.assertEqual(
                document.paragraphs[0]._element.pPr.find(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing'
                ),
                None,
            )
            self.assertEqual(
                document.styles['图片'].type,
                WD_STYLE_TYPE.PARAGRAPH,
            )

    def test_existing_image_paragraph_preset_is_reused(self):
        with tempfile.TemporaryDirectory() as workspace:
            image_path = os.path.join(workspace, 'image.png')
            template_path = os.path.join(workspace, 'template.docx')
            output_path = os.path.join(workspace, 'output.docx')
            Image.new('RGB', (12, 12), color='blue').save(image_path)

            template = Document()
            template.styles.add_style('图片', WD_STYLE_TYPE.PARAGRAPH)
            template.save(template_path)
            document = Document(template_path)
            document.add_paragraph().add_run().add_picture(image_path)
            document.save(output_path)

            clean_document(output_path, template_path=template_path)

            document = Document(output_path)
            names = [style.name for style in document.styles]
            self.assertEqual(names.count('图片'), 1)
            self.assertNotIn('图片（导出）', names)
            self.assertEqual(document.paragraphs[0].style.name, '图片')


class TableBackgroundTests(unittest.TestCase):
    def test_extracts_backgrounds_from_docs_ai_table_xml(self):
        blocks = [
            {'block_id': 'table-1', 'block_type': 31, 'table': {}},
            {'block_id': 'cell-a', 'block_type': 32, 'parent_id': 'table-1', 'children': ['p-a']},
            {'block_id': 'cell-b', 'block_type': 32, 'parent_id': 'table-1', 'children': ['p-b']},
            {'block_id': 'p-a', 'block_type': 2, 'parent_id': 'cell-a'},
            {'block_id': 'p-b', 'block_type': 2, 'parent_id': 'cell-b'},
        ]
        content = (
            '<table id="table-1"><tbody><tr>'
            '<td background-color="rgb(239,240,241)"><p id="p-a">A</p></td>'
            '<td background-color="rgb(240,244,255)"><p id="p-b">B</p></td>'
            '</tr></tbody></table>'
        )
        self.assertEqual(
            extract_table_backgrounds_from_content(content, blocks),
            {'table-1': {'cell-a': 'EFF0F1', 'cell-b': 'F0F4FF'}},
        )

    def test_converter_uses_docs_ai_background_mapping_before_selected_style(self):
        blocks = [
            {'block_id': 'page', 'block_type': 1, 'children': ['table-1']},
            {
                'block_id': 'table-1', 'parent_id': 'page', 'block_type': 31,
                'table': {
                    'cells': ['cell-a', 'cell-b'],
                    'property': {'column_size': 2, 'header_row': True, 'merge_info': [{}, {}]},
                },
            },
            {'block_id': 'cell-a', 'parent_id': 'table-1', 'block_type': 32, 'children': ['p-a'], 'table_cell': {}},
            {'block_id': 'cell-b', 'parent_id': 'table-1', 'block_type': 32, 'children': ['p-b'], 'table_cell': {}},
            {'block_id': 'p-a', 'parent_id': 'cell-a', 'block_type': 2, 'text': {'elements': [{'text_run': {'content': 'A'}}]}},
            {'block_id': 'p-b', 'parent_id': 'cell-b', 'block_type': 2, 'text': {'elements': [{'text_run': {'content': 'B'}}]}},
        ]
        with tempfile.TemporaryDirectory() as workspace:
            output_path = os.path.join(workspace, 'table.docx')
            converter = FeishuDocxConverter(
                blocks,
                client=object(),
                img_dir=workspace,
                table_config={'preserveTableBackground': True},
                table_backgrounds={'table-1': {'cell-a': 'EFF0F1', 'cell-b': 'F0F4FF'}},
            )
            converter.process(output_path)
            document = Document(output_path)
            table = document.tables[0]
            TableStyleManager.apply_style(table, 1, preserve_table_background=True)
            fills = [
                table.cell(0, 0)._tc.get_or_add_tcPr().find(qn('w:shd')).get(qn('w:fill')),
                table.cell(0, 1)._tc.get_or_add_tcPr().find(qn('w:shd')).get(qn('w:fill')),
            ]
            self.assertEqual(fills, ['EFF0F1', 'F0F4FF'])

    def test_normalizes_palette_and_literal_colors(self):
        self.assertEqual(normalize_table_background_color(5), 'DDEBFF')
        self.assertEqual(normalize_table_background_color('#12abEF'), '12ABEF')
        self.assertEqual(
            normalize_table_background_color({'red': 18, 'green': 52, 'blue': 86}),
            '123456',
        )
        self.assertIsNone(normalize_table_background_color('not-a-color'))

    def test_extracts_nested_table_cell_background(self):
        self.assertEqual(
            extract_table_cell_background({
                'table_cell': {'style': {'background_color': '#ABCDEF'}},
            }),
            'ABCDEF',
        )

    def test_selected_style_does_not_override_preserved_fill(self):
        document = Document()
        table = document.add_table(rows=2, cols=1)
        table.cell(0, 0).paragraphs[0].add_run('Header').font.color.rgb = RGBColor(18, 52, 86)
        TableStyleManager._apply_shading(table.cell(0, 0)._tc, '12ABEF')

        TableStyleManager.apply_style(table, 1, preserve_table_background=True)

        header_shading = table.cell(0, 0)._tc.get_or_add_tcPr().find(qn('w:shd'))
        body_shading = table.cell(1, 0)._tc.get_or_add_tcPr().find(qn('w:shd'))
        self.assertEqual(header_shading.get(qn('w:fill')), '12ABEF')
        self.assertEqual(body_shading.get(qn('w:fill')), 'FFFFFF')
        header_color = table.cell(0, 0).paragraphs[0].runs[0].font.color.rgb
        self.assertEqual(str(header_color), '123456')

    def test_converter_preserves_cell_and_header_backgrounds_when_enabled(self):
        blocks = [
            {'block_id': 'page', 'block_type': 1, 'children': ['table']},
            {
                'block_id': 'table',
                'parent_id': 'page',
                'block_type': 31,
                'table': {
                    'cells': ['header', 'body'],
                    'property': {
                        'column_size': 1,
                        'header_row': True,
                        'merge_info': [{}, {}],
                    },
                },
            },
            {
                'block_id': 'header',
                'parent_id': 'table',
                'block_type': 32,
                'table_cell': {'background_color': '#F0A000'},
                'children': ['header-text'],
            },
            {
                'block_id': 'body',
                'parent_id': 'table',
                'block_type': 32,
                'table_cell': {'style': {'backgroundColor': 5}},
                'children': ['body-text'],
            },
            {
                'block_id': 'header-text',
                'parent_id': 'header',
                'block_type': 2,
                'text': {'elements': [{'text_run': {'content': 'Header'}}]},
            },
            {
                'block_id': 'body-text',
                'parent_id': 'body',
                'block_type': 2,
                'text': {'elements': [{'text_run': {'content': 'Body'}}]},
            },
        ]

        with tempfile.TemporaryDirectory() as workspace:
            output_path = os.path.join(workspace, 'table.docx')
            converter = FeishuDocxConverter(
                blocks,
                client=object(),
                img_dir=workspace,
                table_config={'preserveTableBackground': True},
            )
            converter.process(output_path)
            document = Document(output_path)
            table = document.tables[0]
            header_fill = table.cell(0, 0)._tc.get_or_add_tcPr().find(qn('w:shd'))
            body_fill = table.cell(1, 0)._tc.get_or_add_tcPr().find(qn('w:shd'))
            self.assertEqual(header_fill.get(qn('w:fill')), 'F0A000')
            self.assertEqual(body_fill.get(qn('w:fill')), 'DDEBFF')

    def test_converter_ignores_cell_backgrounds_by_default(self):
        blocks = [
            {'block_id': 'page', 'block_type': 1, 'children': ['table']},
            {
                'block_id': 'table', 'parent_id': 'page', 'block_type': 31,
                'table': {
                    'cells': ['cell'],
                    'property': {'column_size': 1, 'merge_info': [{}]},
                },
            },
            {
                'block_id': 'cell', 'parent_id': 'table', 'block_type': 32,
                'table_cell': {'background_color': '#F0A000'},
                'children': [],
            },
        ]
        with tempfile.TemporaryDirectory() as workspace:
            output_path = os.path.join(workspace, 'table.docx')
            FeishuDocxConverter(blocks, object(), workspace).process(output_path)
            cell = Document(output_path).tables[0].cell(0, 0)
            self.assertIsNone(cell._tc.get_or_add_tcPr().find(qn('w:shd')))

if __name__ == '__main__':
    unittest.main()
