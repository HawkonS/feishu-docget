from docx.shared import RGBColor
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml

class TableStyleManager:
    BORDER_1PX = 6
    BORDER_2PX = 12
    STYLES = {1: '样式 1: 深蓝表头 + 白字加粗', 2: '样式 2: 浅蓝表头 + 网格边框', 3: '样式 3: 浅灰表头 + 细网格边框', 4: '样式 4: 全黑实线 (2px)', 5: '样式 5: 上下黑边 + 中间灰竖线', 6: '样式 6: 黑表头 + 斑马纹'}

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
