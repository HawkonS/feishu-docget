import logging

from docx.enum.style import WD_STYLE_TYPE
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml


logger = logging.getLogger('doc_download')

class TableStyleManager:
    BORDER_1PX = 6
    BORDER_2PX = 12
    # These are paragraph styles, rather than Word table styles.  Keeping the
    # two names separate lets a user update all table text or all table headers
    # from Word's Styles pane after an export.
    TABLE_BODY_STYLE_NAME = '表格正文'
    TABLE_HEADER_STYLE_NAME = '表格表头'
    IMAGE_STYLE_NAME = '图片'
    STYLES = {1: '样式 1: 深蓝表头 + 白字加粗', 2: '样式 2: 浅蓝表头 + 网格边框', 3: '样式 3: 浅灰表头 + 细网格边框', 4: '样式 4: 全黑实线 (2px)', 5: '样式 5: 上下黑边 + 中间灰竖线', 6: '样式 6: 黑表头 + 斑马纹'}

    @staticmethod
    def ensure_table_paragraph_styles(doc):
        """Return the paragraph styles used by generated table content.

        A template can already contain a style named ``表格正文`` or
        ``表格表头``.  Reusing an existing *paragraph* style is intentional:
        it allows the template author to define the appearance once and have
        every exported table inherit it.  A style with the same display name
        but another type (for example a table style) cannot be assigned to a
        paragraph, so a deterministic ``（导出）`` suffix is used in that case
        instead of replacing the template style.
        """
        styles = getattr(doc, 'styles', None)
        if styles is None:
            raise ValueError('文档不包含样式集合')

        body_style = TableStyleManager._get_or_create_paragraph_style(
            styles, TableStyleManager.TABLE_BODY_STYLE_NAME
        )
        header_style = TableStyleManager._get_or_create_paragraph_style(
            styles, TableStyleManager.TABLE_HEADER_STYLE_NAME
        )
        return body_style, header_style

    @staticmethod
    def ensure_image_paragraph_style(doc):
        """Return the paragraph style used by paragraphs containing images."""
        styles = getattr(doc, 'styles', None)
        if styles is None:
            raise ValueError('文档不包含样式集合')
        return TableStyleManager._get_or_create_paragraph_style(
            styles, TableStyleManager.IMAGE_STYLE_NAME
        )

    @staticmethod
    def is_image_paragraph(paragraph):
        """Whether a paragraph contains a DrawingML or legacy image."""
        try:
            xml = paragraph._element.xml
            return 'w:drawing' in xml or 'w:pict' in xml
        except Exception:
            return False

    @staticmethod
    def apply_image_paragraph_style(paragraph, image_style):
        """Assign the independent image paragraph style when applicable."""
        if TableStyleManager.is_image_paragraph(paragraph):
            try:
                paragraph.style = image_style
            except (TypeError, ValueError):
                logger.warning(
                    '设置图片段落样式失败，保留原段落样式: %s',
                    getattr(image_style, 'name', image_style),
                )
            return True
        return False

    @staticmethod
    def _get_or_create_paragraph_style(styles, base_name):
        """Find a usable style by name or create one without name collisions."""
        existing = TableStyleManager._find_style_by_name(styles, base_name)
        if existing is not None:
            if existing.type == WD_STYLE_TYPE.PARAGRAPH:
                return existing
            logger.info(
                '模板中的样式 %s 不是段落样式，将创建独立的导出样式',
                base_name,
            )

        candidate = base_name if existing is None else f'{base_name}（导出）'
        suffix = 2
        while True:
            candidate_style = TableStyleManager._find_style_by_name(styles, candidate)
            if candidate_style is None:
                break
            # This commonly occurs when the converter has already created the
            # suffixed style and the later cleaner pass resolves the same
            # template again.  Reuse that paragraph style rather than creating
            # a second suffix on every pass.
            if candidate_style.type == WD_STYLE_TYPE.PARAGRAPH:
                return candidate_style
            candidate = f'{base_name}（导出 {suffix}）'
            suffix += 1

        style = styles.add_style(candidate, WD_STYLE_TYPE.PARAGRAPH)
        # Inherit the document's Normal style so a newly-created style keeps
        # the template's ordinary font until the user customizes it.
        try:
            style.base_style = styles['Normal']
        except (KeyError, ValueError):
            pass
        try:
            # Put the new presets in Word's styles gallery as well as the
            # full Styles pane, making them easy to discover and edit.
            style.quick_style = True
        except (AttributeError, ValueError):
            pass
        logger.info('已创建表格段落样式: %s', candidate)
        return style

    @staticmethod
    def _find_style_by_name(styles, name):
        # Iterating is more reliable than styles[name] because python-docx
        # resolves both UI names and style IDs, while templates may contain
        # localized aliases or a non-paragraph style with the same display
        # name.
        wanted = str(name).casefold()
        for style in styles:
            try:
                if str(style.name or '').casefold() == wanted:
                    return style
            except Exception:
                continue
        return None

    @staticmethod
    def apply_table_paragraph_styles(table, body_style, header_style):
        """Assign body/header paragraph styles to one generated table."""
        for row_index, row in enumerate(table.rows):
            paragraph_style = header_style if row_index == 0 else body_style
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    try:
                        paragraph.style = paragraph_style
                    except (TypeError, ValueError):
                        # A malformed template style should not make an export
                        # fail.  The generated paragraph remains usable with
                        # its direct formatting in that rare case.
                        logger.warning(
                            '设置表格段落样式失败，保留原段落样式: %s',
                            getattr(paragraph_style, 'name', paragraph_style),
                        )

    @staticmethod
    def list_styles():
        return [{'id': k, 'name': v} for k, v in sorted(TableStyleManager.STYLES.items())]

    @staticmethod
    def get_frontend_css():
        return '\n        /* 样式 1: 深蓝表头 + 白字加粗 */\n        table.style-1 th { background: #445bdc; color: white; font-weight: bold; border: 1px solid #D9D9D9; }\n        table.style-1 td { background: white; color: black; border: 1px solid #D9D9D9; }\n        \n        /* 样式 2: 浅蓝表头 + 网格边框 */\n        table.style-2 th { background: #E6F3FF; color: black; font-weight: bold; border: 1px solid #999; }\n        table.style-2 td { background: white; color: black; border: 1px solid #999; }\n        \n        /* 样式 3: 浅灰表头 + 细网格边框 */\n        table.style-3 th { background: #F2F2F2; color: black; border: 1px solid #D9D9D9; }\n        table.style-3 td { background: white; color: black; border: 1px solid #D9D9D9; }\n        \n        /* 样式 4: 全黑实线 (2px) */\n        table.style-4 th, table.style-4 td { background: white; color: black; border: 2px solid black; }\n        \n        /* 样式 5: 上下黑边 + 中间灰竖线 */\n        table.style-5 th, table.style-5 td { border: 1px solid #D9D9D9; color: black; background: white; }\n        table.style-5 tr:first-child th, table.style-5 tr:first-child td { border-top: 1px solid black; }\n        table.style-5 tr:last-child th, table.style-5 tr:last-child td { border-bottom: 1px solid black; }\n        \n        /* 样式 6: 黑表头 + 斑马纹 */\n        table.style-6 thead tr th, table.style-6 thead tr td { background: black; color: white; font-weight: bold; border: 1px solid #D9D9D9; }\n        table.style-6 tbody tr:nth-child(odd) td, table.style-6 tbody tr:nth-child(odd) th { background: #F2F2F2; }\n        table.style-6 tbody tr:nth-child(even) td, table.style-6 tbody tr:nth-child(even) th { background: white; }\n        table.style-6 td, table.style-6 th { border: 1px solid #D9D9D9; color: black; }\n        '

    @staticmethod
    def apply_style(table, style_id):
        TableStyleManager._clear_table_borders(table)
        try:
            style_id = int(style_id)
        except:
            return
        if style_id == 1:
            TableStyleManager._apply_style_1(table)
        elif style_id == 2:
            TableStyleManager._apply_style_2(table)
        elif style_id == 3:
            TableStyleManager._apply_style_3(table)
        elif style_id == 4:
            TableStyleManager._apply_style_4(table)
        elif style_id == 5:
            TableStyleManager._apply_style_5(table)
        elif style_id == 6:
            TableStyleManager._apply_style_6(table)

    @staticmethod
    def apply_default_sheet_style(table):
        TableStyleManager._clear_table_borders(table)
        border_light = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': 'D9D9D9'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            TableStyleManager._apply_border(tc, top=border_light, bottom=border_light, left=border_light, right=border_light)
            TableStyleManager._apply_shading(tc, 'FFFFFF')

    @staticmethod
    def _clear_table_borders(table):
        tblPr = table._element.tblPr
        if tblPr is None:
            return
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)

    @staticmethod
    def _iter_cells(table):
        for r_idx, tr in enumerate(table._element.tr_lst):
            c_idx = 0
            for tc in tr.tc_lst:
                yield (r_idx, c_idx, tc)
                grid_span = 1
                tcPr = tc.get_or_add_tcPr()
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        val = gs.get(qn('w:val'))
                        if val:
                            grid_span = int(val)
                c_idx += grid_span

    @staticmethod
    def _apply_border(tc, top=None, bottom=None, left=None, right=None):
        tcPr = tc.get_or_add_tcPr()
        tcBorders = tcPr.first_child_found_in('w:tcBorders')
        if tcBorders is None:
            tcBorders = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />')
            tcPr.append(tcBorders)
        vMerge = tcPr.find(qn('w:vMerge'))
        is_restart = False
        is_continue = False
        if vMerge is not None:
            val = vMerge.get(qn('w:val'))
            if val == 'restart':
                is_restart = True
            else:
                is_continue = True
        if is_continue and top:
            top = None
        if is_continue:
            top = {'val': 'nil'}
        if is_restart:
            bottom = {'val': 'nil'}
        for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            tag = f'w:{edge}'
            existing = tcBorders.find(parse_xml(f'<{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag)
            if existing is not None:
                tcBorders.remove(existing)
            if val:
                element = parse_xml(f'<{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />')
                tcBorders.append(element)
                if val.get('val') == 'nil':
                    element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'val'), 'nil')
                else:
                    element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'val'), val.get('val', 'single'))
                    element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'sz'), str(val.get('sz', 4)))
                    element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'space'), '0')
                    element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'color'), val.get('color', 'auto'))

    @staticmethod
    def _apply_shading(tc, color_hex):
        tcPr = tc.get_or_add_tcPr()
        shd = tcPr.first_child_found_in('w:shd')
        if shd is not None:
            tcPr.remove(shd)
        shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color_hex}"/>')
        tcPr.append(shd)

    @staticmethod
    def _set_cell_text_color(tc, color_hex, bold=False):
        for p in tc.p_lst:
            for r in p.r_lst:
                rPr = r.get_or_add_rPr()
                if color_hex:
                    color_el = parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color_hex}"/>')
                    existing = rPr.find(qn('w:color'))
                    if existing is not None:
                        rPr.remove(existing)
                    rPr.append(color_el)
                if bold:
                    b_el = parse_xml(f'<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                    existing = rPr.find(qn('w:b'))
                    if existing is not None:
                        pass
                    else:
                        rPr.append(b_el)

    @staticmethod
    def _apply_style_1(table):
        border_gray = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': 'D9D9D9'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            TableStyleManager._apply_border(tc, top=border_gray, bottom=border_gray, left=border_gray, right=border_gray)
            if r_idx == 0:
                TableStyleManager._apply_shading(tc, '445bdc')
                TableStyleManager._set_cell_text_color(tc, 'FFFFFF', bold=True)
            else:
                TableStyleManager._apply_shading(tc, 'FFFFFF')
                TableStyleManager._set_cell_text_color(tc, '000000')

    @staticmethod
    def _apply_style_2(table):
        border = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': '999999'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            TableStyleManager._apply_border(tc, top=border, bottom=border, left=border, right=border)
            if r_idx == 0:
                TableStyleManager._apply_shading(tc, 'E6F3FF')
                TableStyleManager._set_cell_text_color(tc, '000000', bold=True)
            else:
                TableStyleManager._apply_shading(tc, 'FFFFFF')
                TableStyleManager._set_cell_text_color(tc, '000000')

    @staticmethod
    def _apply_style_3(table):
        border_gray = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': 'D9D9D9'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            TableStyleManager._apply_border(tc, top=border_gray, bottom=border_gray, left=border_gray, right=border_gray)
            if r_idx == 0:
                TableStyleManager._apply_shading(tc, 'F2F2F2')
            else:
                TableStyleManager._apply_shading(tc, 'FFFFFF')
            TableStyleManager._set_cell_text_color(tc, '000000')

    @staticmethod
    def _apply_style_4(table):
        border_black = {'val': 'single', 'sz': TableStyleManager.BORDER_2PX, 'color': '000000'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            TableStyleManager._apply_border(tc, top=border_black, bottom=border_black, left=border_black, right=border_black)
            TableStyleManager._apply_shading(tc, 'FFFFFF')
            TableStyleManager._set_cell_text_color(tc, '000000')

    @staticmethod
    def _apply_style_5(table):
        border_black = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': '000000'}
        border_gray = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': 'D9D9D9'}
        last_row_idx = len(table._element.tr_lst) - 1
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            top = border_gray
            bottom = border_gray
            left = border_gray
            right = border_gray
            if r_idx == 0:
                top = border_black
            if r_idx == last_row_idx:
                bottom = border_black
            TableStyleManager._apply_border(tc, top=top, bottom=bottom, left=left, right=right)
            TableStyleManager._apply_shading(tc, 'FFFFFF')
            TableStyleManager._set_cell_text_color(tc, '000000')

    @staticmethod
    def _apply_style_6(table):
        border = {'val': 'single', 'sz': TableStyleManager.BORDER_1PX, 'color': 'D9D9D9'}
        for r_idx, c_idx, tc in TableStyleManager._iter_cells(table):
            if r_idx == 0:
                color = '000000'
                text_color = 'FFFFFF'
                bold = True
            else:
                color = 'F2F2F2' if (r_idx - 1) % 2 == 0 else 'FFFFFF'
                text_color = '000000'
                bold = False
            TableStyleManager._apply_shading(tc, color)
            TableStyleManager._set_cell_text_color(tc, text_color, bold=bold)
            TableStyleManager._apply_border(tc, top=border, bottom=border, left=border, right=border)
