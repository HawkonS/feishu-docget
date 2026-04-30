from datetime import datetime, timezone
import os
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from docx import Document
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.shared import Pt, Mm, RGBColor, Cm
from docx.enum.text import WD_LINE_SPACING
from copy import deepcopy
from src.core.config_loader import ConfigLoader, config
try:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
except ImportError:

    class RT:
        IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
logger = ConfigLoader.get_logger('format_cleaner')

def _align_to_docx(align_value, default=1):
    align_map = {'left': 0, 'center': 1, 'right': 2}
    if align_value is None:
        return default
    return align_map.get(str(align_value).lower(), default)

def _resolve_image_style(style_config, fallback_width, fallback_height, fallback_align=1):
    width = fallback_width
    height = fallback_height
    align = fallback_align
    if isinstance(style_config, dict):
        raw_width = style_config.get('maxWidth')
        raw_height = style_config.get('maxHeight')
        try:
            if raw_width not in (None, ''):
                width = float(raw_width)
        except (ValueError, TypeError):
            pass
        try:
            if raw_height not in (None, ''):
                height = float(raw_height)
        except (ValueError, TypeError):
            pass
        align = _align_to_docx(style_config.get('align'), fallback_align)
    return width, height, align

def _apply_image_style_to_paragraph(paragraph, ns, ns_wp, max_height, max_width, align, clear_space_before=False):
    resized = 0
    aligned = 0
    xml = paragraph._element.xml
    is_image = False
    if 'w:drawing' in xml or 'w:pict' in xml:
        is_image = True
        
    for run in paragraph.runs:
        drawings = run._element.findall(f'.//{{{ns}}}drawing')
        for drawing in drawings:
            inlines = drawing.findall(f'.//{{{ns_wp}}}inline')
            for inline in inlines:
                is_image = True
                if _resize_inline_image(inline, max_height, max_width):
                    resized += 1
            anchors = drawing.findall(f'.//{{{ns_wp}}}anchor')
            for anchor in anchors:
                is_image = True
                if _resize_inline_image(anchor, max_height, max_width):
                    resized += 1
                    
    if is_image:
        _force_clear_indent(paragraph, ns, clear_space_before=clear_space_before)
        paragraph.alignment = align
        aligned += 1
        
    return resized, aligned

def _get_all_tables(parent):
    tables = []
    if hasattr(parent, 'rows'): # parent is a table
        tables.append(parent)
        for row in parent.rows:
            for cell in row.cells:
                for nested_table in getattr(cell, 'tables', []):
                    tables.extend(_get_all_tables(nested_table))
    else: # parent is a document or cell
        for table in getattr(parent, 'tables', []):
            tables.extend(_get_all_tables(table))
    return tables

def _set_table_width(table, width_cm, ns):
    try:
        width_cm = float(width_cm)
    except (ValueError, TypeError):
        return False
    if width_cm <= 0:
        return False
    width_twips = str(int(width_cm / 2.54 * 1440))
    try:
        table.autofit = False
    except Exception:
        pass
    tblPr = table._element.tblPr
    if tblPr is None:
        return False
    tbl_w = tblPr.find(f'{{{ns}}}tblW')
    if tbl_w is None:
        tbl_w = parse_xml(f'<w:tblW xmlns:w="{ns}" w:w="{width_twips}" w:type="dxa"/>')
        tblPr.append(tbl_w)
    else:
        tbl_w.set(f'{{{ns}}}w', width_twips)
        tbl_w.set(f'{{{ns}}}type', 'dxa')
    tbl_layout = tblPr.find(f'{{{ns}}}tblLayout')
    if tbl_layout is None:
        tbl_layout = parse_xml(f'<w:tblLayout xmlns:w="{ns}" w:type="fixed"/>')
        tblPr.append(tbl_layout)
    else:
        tbl_layout.set(f'{{{ns}}}type', 'fixed')
    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(f'{{{ns}}}tcW')
            if tc_w is None:
                tc_w = parse_xml(f'<w:tcW xmlns:w="{ns}" w:w="{width_twips}" w:type="dxa"/>')
                tc_pr.append(tc_w)
            else:
                tc_w.set(f'{{{ns}}}w', width_twips)
                tc_w.set(f'{{{ns}}}type', 'dxa')
                
    # Remove gridCols to ensure fixed width takes effect fully without interference
    tblGrid = tblPr.getparent().find(f'{{{ns}}}tblGrid')
    if tblGrid is not None:
        tblPr.getparent().remove(tblGrid)
        
    return True

def _insert_or_update_tblInd(tblPr, ns, w="0", type="dxa"):
    tblInd = tblPr.find(f'{{{ns}}}tblInd')
    if tblInd is not None:
        tblInd.set(f'{{{ns}}}w', str(w))
        tblInd.set(f'{{{ns}}}type', type)
        return tblInd
    
    tblInd = parse_xml(f'<w:tblInd xmlns:w="{ns}" w:w="{w}" w:type="{type}"/>')
    tags_after = ['tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
    inserted = False
    for tag in tags_after:
        el = tblPr.find(f'{{{ns}}}{tag}')
        if el is not None:
            el.addprevious(tblInd)
            inserted = True
            break
    if not inserted:
        tblPr.append(tblInd)
    return tblInd

def _apply_table_layout(table, width_str, auto_fit, ns, min_col_width=120):
    changed = False
    try:
        table.autofit = False
        changed = True
    except Exception:
        pass

    tblPr = table._element.tblPr
    if tblPr is None:
        return changed

    layout_type = 'fixed'
    tbl_layout = tblPr.find(f'{{{ns}}}tblLayout')
    if tbl_layout is None:
        tbl_layout = parse_xml(f'<w:tblLayout xmlns:w="{ns}" w:type="{layout_type}"/>')
        tags_after = ['tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
        inserted = False
        for tag in tags_after:
            el = tblPr.find(f'{{{ns}}}{tag}')
            if el is not None:
                el.addprevious(tbl_layout)
                inserted = True
                break
        if not inserted:
            tblPr.append(tbl_layout)
    else:
        tbl_layout.set(f'{{{ns}}}type', layout_type)

    w_val = '0'
    w_type = 'auto'
    if width_str:
        width_str = str(width_str).strip().lower()
        if width_str.endswith('%'):
            try:
                pct = float(width_str[:-1])
                w_val = str(int(pct * 50))
                w_type = 'pct'
            except:
                pass
        elif width_str == 'auto':
            w_val = '0'
            w_type = 'auto'
        else:
            try:
                if width_str.endswith('cm'):
                    cm_val = float(width_str[:-2])
                else:
                    cm_val = float(width_str)
                w_val = str(int(cm_val / 2.54 * 1440))
                w_type = 'dxa'
            except:
                pass
    elif auto_fit:
        w_val = '5000'
        w_type = 'pct'

    tbl_w = tblPr.find(f'{{{ns}}}tblW')
    if tbl_w is None:
        tbl_w = parse_xml(f'<w:tblW xmlns:w="{ns}" w:w="{w_val}" w:type="{w_type}"/>')
        tags_after = ['jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
        inserted = False
        for tag in tags_after:
            el = tblPr.find(f'{{{ns}}}{tag}')
            if el is not None:
                el.addprevious(tbl_w)
                inserted = True
                break
        if not inserted:
            tblPr.append(tbl_w)
    else:
        tbl_w.set(f'{{{ns}}}w', w_val)
        tbl_w.set(f'{{{ns}}}type', w_type)

    min_col_dxa = int(min_col_width * 240) if min_col_width is not None else 8 * 240
    is_single_col_fixed = w_type == 'dxa' and all(len(r.cells) == 1 for r in table.rows)

    page_width_dxa = 9070
    if w_type == 'pct':
        actual_table_dxa = int((int(w_val) / 5000) * page_width_dxa)
    elif w_type == 'dxa':
        actual_table_dxa = int(w_val)
    else:
        actual_table_dxa = page_width_dxa
        
    min_pct = int((min_col_dxa / actual_table_dxa) * 5000) if actual_table_dxa > 0 else 5000

    final_pcts = {}
    if not is_single_col_fixed:
        col_widths = {}
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                max_line_w = 0
                for p in cell.paragraphs:
                    text = p.text or ""
                    line_w = sum(240 if ord(c) > 127 else 120 for c in text)
                    max_line_w = max(max_line_w, line_w)
                max_line_w += 200
                col_widths[i] = max(col_widths.get(i, 0), max_line_w)
                
        targets = {i: max(cw, min_col_dxa) for i, cw in col_widths.items()}
        
        total_target = sum(targets.values())
        
        if total_target <= actual_table_dxa:
            final_pcts = {i: int((w / actual_table_dxa) * 5000) for i, w in targets.items()}
            leftover = 5000 - sum(final_pcts.values())
            if leftover > 0 and targets:
                max_w = max(targets.values())
                max_cols = [i for i, w in targets.items() if w == max_w]
                share = leftover // len(max_cols)
                for i in max_cols:
                    final_pcts[i] += share
                final_pcts[max_cols[0]] += leftover % len(max_cols)
        else:
            final_pcts = {i: min_pct for i in targets}
            leftover = 5000 - sum(final_pcts.values())
            
            if leftover < 0:
                total_min = sum(final_pcts.values())
                final_pcts = {i: int((min_pct / total_min) * 5000) for i in targets}
            else:
                extra_w = {i: targets[i] - min_col_dxa for i in targets}
                total_extra = sum(extra_w.values())
                
                if total_extra > 0:
                    for i in targets:
                        share = int((extra_w[i] / total_extra) * leftover)
                        final_pcts[i] += share
                    
                    current_total = sum(final_pcts.values())
                    if current_total < 5000:
                        max_extra = max(extra_w.values())
                        max_col = [i for i, w in extra_w.items() if w == max_extra][0]
                        final_pcts[max_col] += (5000 - current_total)

    tblGrids = tblPr.getparent().findall(f'{{{ns}}}tblGrid')
    for tblGrid in tblGrids:
        gridCols = tblGrid.findall(f'{{{ns}}}gridCol')
        for i, gridCol in enumerate(gridCols):
            if is_single_col_fixed:
                gridCol.set(f'{{{ns}}}w', str(max(int(w_val), min_col_dxa)))
            elif i in final_pcts:
                grid_dxa = int((final_pcts[i] / 5000) * actual_table_dxa)
                gridCol.set(f'{{{ns}}}w', str(grid_dxa))
                
    if is_single_col_fixed:
        for tg in tblGrids:
            tblPr.getparent().remove(tg)

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(f'{{{ns}}}tcW')
            
            if is_single_col_fixed:
                target_w = str(max(int(w_val), min_col_dxa))
                t_type = 'dxa'
            else:
                target_w = str(final_pcts.get(i, min_pct))
                t_type = 'pct'
            
            if tc_w is not None:
                tc_w.set(f'{{{ns}}}w', target_w)
                tc_w.set(f'{{{ns}}}type', t_type)
            else:
                tc_w = parse_xml(f'<w:tcW xmlns:w="{ns}" w:w="{target_w}" w:type="{t_type}"/>')
                tags_after = ['gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']
                inserted = False
                for tag in tags_after:
                    el = tc_pr.find(f'{{{ns}}}{tag}')
                    if el is not None:
                        el.addprevious(tc_w)
                        inserted = True
                        break
                if not inserted:
                    tc_pr.append(tc_w)

    return True


def clean_document(docx_path, progress_cb=None, template_path=None, add_cover=False, body_style=None, image_style=None, table_config=None, margin_config=None, code_block_config=None, document_info=None, ignore_template_heading_num=False):
    logger.debug('开始清理文档...')
    doc = Document(docx_path)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    
    if template_path:
        try:
            tpl_doc = Document(template_path)
            _copy_styles_from_template(template_path, doc, ignore_template_heading_num=ignore_template_heading_num)
            
            # 清理 numbering_part 中关联到标题样式的 pStyle
            if ignore_template_heading_num:
                try:
                    if doc.part.numbering_part:
                        heading_style_ids = set()
                        for s in doc.styles:
                            if _is_heading_style(s._element, ns):
                                heading_style_ids.add(s.style_id)
                                
                        numbering_element = doc.part.numbering_part.element
                        p_styles = numbering_element.findall(f'.//{{{ns}}}pStyle')
                        for ps in p_styles:
                            val = ps.get(f'{{{ns}}}val')
                            if val and val in heading_style_ids:
                                ps.getparent().remove(ps)
                except Exception as e:
                    logger.warning(f"Failed to clear pStyle from numbering: {e}")
            
            inserted_count = 0
            if add_cover:
                inserted_count = _prepend_first_page_from_template(doc, template_path)
            
            _copy_headers_from_template(tpl_doc, doc, add_cover=add_cover)
            
            if progress_cb:
                msg = '已添加封面' if add_cover else '已应用模板样式和页眉页脚'
                progress_cb(92, msg, 'success')
        except Exception as e:
            logger.error(f'应用模板失败: {str(e)}')
            if progress_cb:
                progress_cb(92, f'应用模板失败: {str(e)}')
            inserted_count = 0
    else:
        inserted_count = 0
        if progress_cb:
            progress_cb(92, '已跳过应用模板', 'success')
    
    heading_style_ids = _get_heading_style_ids(doc, ns)
    template_heading_numbering_indents = _get_template_heading_numbering_indents(template_path, heading_style_ids, ns)

    default_max_h = 23.0
    try:
        val = config.get('image.max_height', 23)
        if isinstance(val, str) and (not val.strip()):
            val = 23
        default_max_h = float(val)
    except (ValueError, TypeError):
        pass
        
    default_max_w = 16.0
    try:
        val_w = config.get('image.max_width', 16)
        if isinstance(val_w, str) and (not val_w.strip()):
            val_w = 16
        default_max_w = float(val_w)
    except (ValueError, TypeError):
        pass

    target_max_w, target_max_h, target_align = _resolve_image_style(image_style, default_max_w, default_max_h, 1)
    table_image_style = image_style.get('tableImageStyle') if isinstance(image_style, dict) else None
    table_target_max_w, table_target_max_h, table_target_align = _resolve_image_style(
        table_image_style,
        target_max_w,
        target_max_h,
        target_align
    )
    max_height = Mm(target_max_h * 10)
    max_width = Mm(target_max_w * 10)
    table_max_height = Mm(table_target_max_h * 10)
    table_max_width = Mm(table_target_max_w * 10)

    force_clear_tbl_indent = True
    force_clear_tbl_image_space = True
    header_align = None
    content_align = None
    table_auto_fit = False
    
    if table_config:
        force_clear_tbl_indent = table_config.get('forceClearIndent', True)
        force_clear_tbl_image_space = table_config.get('forceClearImageSpace', True)
        table_auto_fit = bool(table_config.get('autoFit', True))
        table_width_str = table_config.get('width', '100%')
        min_col_width = table_config.get('minColWidth', 8)
        
        table_line_spacing = table_config.get('lineSpacing')
        table_space_before = table_config.get('spaceBefore')
        table_space_after = table_config.get('spaceAfter')
        
        if not table_width_str:
            table_width_str = '100%'
        header_align = _align_to_docx(table_config.get('headerAlign', 'center'), 1)
        content_align = _align_to_docx(table_config.get('contentAlign', 'left'), 0)
        content_image_align = _align_to_docx(table_config.get('contentImageAlign', 'left'), 0)
    else:
        force_clear_tbl_indent = True
        force_clear_tbl_image_space = True
        table_auto_fit = True
        table_width_str = '100%'
        min_col_width = 8
        table_line_spacing = None
        table_space_before = None
        table_space_after = None
        header_align = 1  # 默认居中
        content_align = 0 # 默认靠左
        content_image_align = 0 # 默认靠左

    if progress_cb:
        progress_cb(93, '正在处理表格样式...', 'dynamic')
    cover_table_count = inserted_count[1] if isinstance(inserted_count, tuple) else 0
    cover_para_count = inserted_count[0] if isinstance(inserted_count, tuple) else inserted_count if isinstance(inserted_count, int) else 0
    
    top_tables = doc.tables[cover_table_count:] if cover_table_count < len(doc.tables) else []
    all_tables = []
    for t in top_tables:
        all_tables.extend(_get_all_tables(t))

    # First, handle indentation clearance for ALL tables 
    # to avoid interference from nested table processing later.
    count_indent = 0
    for table in all_tables:
        count_indent += 1
        tblPr = table._element.tblPr
        if tblPr is not None:
            _insert_or_update_tblInd(tblPr, ns, w="0", type="dxa")
            
            # Clear cell margin at table level as well if requested
            if force_clear_tbl_indent:
                tblCellMar = tblPr.find(f'{{{ns}}}tblCellMar')
                if tblCellMar is None:
                    tblCellMar = parse_xml(f'<w:tblCellMar xmlns:w="{ns}"><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>')
                    tags_after = ['tblLook', 'tblCaption', 'tblDescription']
                    inserted = False
                    for tag in tags_after:
                        el = tblPr.find(f'{{{ns}}}{tag}')
                        if el is not None:
                            el.addprevious(tblCellMar)
                            inserted = True
                            break
                    if not inserted:
                        tblPr.append(tblCellMar)
                else:
                    for side in ['left', 'right']:
                        side_elem = tblCellMar.find(f'{{{ns}}}{side}')
                        if side_elem is not None:
                            side_elem.set(f'{{{ns}}}w', '0')
                            side_elem.set(f'{{{ns}}}type', 'dxa')
                        else:
                            tblCellMar.append(parse_xml(f'<w:{side} xmlns:w="{ns}" w:w="0" w:type="dxa"/>'))
        
        # Now process all paragraphs in the table to clear indent
        for r_idx, row in enumerate(table.rows):
            for cell in row.cells:
                # Also apply margin clear to cell level
                if force_clear_tbl_indent:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcMar = tcPr.find(f'{{{ns}}}tcMar')
                    if tcMar is None:
                        tcMar = parse_xml(f'<w:tcMar xmlns:w="{ns}"><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>')
                        tags_after = ['textDirection', 'tcFitText', 'vAlign', 'hideMark']
                        inserted = False
                        for tag in tags_after:
                            el = tcPr.find(f'{{{ns}}}{tag}')
                            if el is not None:
                                el.addprevious(tcMar)
                                inserted = True
                                break
                        if not inserted:
                            tcPr.append(tcMar)
                    else:
                        for side in ['left', 'right']:
                            side_elem = tcMar.find(f'{{{ns}}}{side}')
                            if side_elem is not None:
                                side_elem.set(f'{{{ns}}}w', '0')
                                side_elem.set(f'{{{ns}}}type', 'dxa')
                            else:
                                tcMar.append(parse_xml(f'<w:{side} xmlns:w="{ns}" w:w="0" w:type="dxa"/>'))

                for p in cell.paragraphs:
                    if force_clear_tbl_indent:
                        _force_clear_indent(p, ns, clear_space_before=force_clear_tbl_image_space)
                        # Ensure paragraph indentation properties are fully overridden
                        pPr = p._element.get_or_add_pPr()
                        ind = pPr.get_or_add_ind()
                        ind.set(f'{{{ns}}}left', '0')
                        ind.set(f'{{{ns}}}right', '0')
                        ind.set(f'{{{ns}}}firstLine', '0')
                        ind.set(f'{{{ns}}}hanging', '0')
                    else:
                        _clean_text_indent(p, ns)
                        
    count_content = 0
    count_auto_fit = 0
    count_table_image_resized = 0
    count_table_image_aligned = 0
    
    # 1. 优先处理所有常规表格样式和内容
    for table in all_tables:
        is_code_block = False
        try:
            tblPr = table._element.tblPr
            if tblPr is not None:
                caption = tblPr.find(f'{{{ns}}}tblCaption')
                if caption is not None and caption.get(f'{{{ns}}}val') == 'code_block':
                    is_code_block = True
        except:
            pass
            
        if is_code_block:
            continue
            
        count_content += 1
        if _apply_table_layout(table, table_width_str, table_auto_fit, ns, min_col_width):
            count_auto_fit += 1

        for r_idx, row in enumerate(table.rows):
            is_header = (r_idx == 0)
            for cell in row.cells:
                cell.vertical_alignment = 1
                for p in cell.paragraphs:
                    if table_config:
                        if is_header:
                            p.alignment = header_align
                        else:
                            p.alignment = content_align
                    
                    # 对于表格中的图片，让图片严格跟随单元格内容的图片对齐方式配置
                    current_image_align = header_align if is_header else content_image_align

                    resized, aligned = _apply_image_style_to_paragraph(
                        p,
                        ns,
                        ns_wp,
                        table_max_height,
                        table_max_width,
                        current_image_align,
                        clear_space_before=force_clear_tbl_image_space
                    )
                    count_table_image_resized += resized
                    count_table_image_aligned += aligned
                    
                    # Only apply body style to text paragraphs, not to image paragraphs
                    # if we are meant to clear their spacing
                    if body_style:
                        if aligned > 0 and force_clear_tbl_image_space:
                            # Create a copy of body_style without spaceBefore
                            modified_body_style = body_style.copy()
                            modified_body_style['spaceBefore'] = None
                            modified_body_style['spaceBeforeUnit'] = None
                            _apply_paragraph_style(p, modified_body_style, ns)
                        else:
                            _apply_paragraph_style(p, body_style, ns)
                            
                    # 覆盖表格独有的行间距和段前段后距
                    if table_line_spacing is not None or table_space_before is not None or table_space_after is not None:
                        p_pr = p._element.get_or_add_pPr()
                        spacing = p_pr.find(f'{{{ns}}}spacing')
                        if spacing is None:
                            spacing = parse_xml(f'<w:spacing xmlns:w="{ns}"/>')
                            p_pr.append(spacing)
                            
                        if table_line_spacing is not None:
                            try:
                                spacing.set(f'{{{ns}}}line', str(int(float(table_line_spacing) * 240)))
                                spacing.set(f'{{{ns}}}lineRule', 'auto')
                            except:
                                pass
                                
                        if table_space_before is not None:
                            try:
                                spacing.set(f'{{{ns}}}beforeLines', str(int(float(table_space_before) * 100)))
                            except:
                                pass
                                
                        if table_space_after is not None:
                            try:
                                spacing.set(f'{{{ns}}}afterLines', str(int(float(table_space_after) * 100)))
                            except:
                                pass

    # 2. 最后单独处理所有代码块表格，确保代码块的样式和宽度设置不会被外部常规表格覆盖
    for table in all_tables:
        is_code_block = False
        is_inner_code_block = False
        try:
            tblPr = table._element.tblPr
            if tblPr is not None:
                caption = tblPr.find(f'{{{ns}}}tblCaption')
                if caption is not None and caption.get(f'{{{ns}}}val') == 'code_block':
                    is_code_block = True
                    # Check if it's inside another table
                    parent = table._element.getparent()
                    while parent is not None:
                        if parent.tag == f'{{{ns}}}tbl':
                            is_inner_code_block = True
                            break
                        parent = parent.getparent()
        except:
            pass
            
        if is_code_block:
            count_content += 1
            
            # Determine actual width to use
            actual_table_width = None
            if code_block_config:
                if is_inner_code_block and code_block_config.get('innerTableWidth') is not None:
                    actual_table_width = code_block_config.get('innerTableWidth')
                else:
                    actual_table_width = code_block_config.get('tableWidth')
                
                # Temporarily set the config's tableWidth to actual_table_width
                # so that _apply_custom_code_block_style uses it
                original_table_width = code_block_config.get('tableWidth')
                code_block_config['tableWidth'] = actual_table_width
                
                _apply_custom_code_block_style(table, code_block_config, ns)
                
                # Restore it
                code_block_config['tableWidth'] = original_table_width
            
            # Retrieve table width for code block alignment processing if needed
            table_width = actual_table_width
            
            # Since Word doesn't natively center table easily without specific tag <w:jc>,
            # apply alignment at the table level using <w:jc>
            alignment = _align_to_docx(code_block_config.get('align', 'left') if code_block_config else 'left', 0)
            align_val_map = {0: 'left', 1: 'center', 2: 'right'}
            jc_val = align_val_map.get(alignment, 'left')
            
            tblPr = table._element.tblPr
            if tblPr is not None:
                jc = tblPr.find(f'{{{ns}}}jc')
                if jc is not None:
                    jc.set(f'{{{ns}}}val', jc_val)
                else:
                    jc = parse_xml(f'<w:jc xmlns:w="{ns}" w:val="{jc_val}"/>')
                    tags_after = ['tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
                    inserted = False
                    for tag in tags_after:
                        el = tblPr.find(f'{{{ns}}}{tag}')
                        if el is not None:
                            el.addprevious(jc)
                            inserted = True
                            break
                    if not inserted:
                        tblPr.append(jc)
                    
            for r_idx, row in enumerate(table.rows):
                for cell in row.cells:
                    # Also apply margin to cell
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcMar = tcPr.find(f'{{{ns}}}tcMar')
                    if tcMar is None:
                        tcMar = parse_xml(f'<w:tcMar xmlns:w="{ns}"><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>')
                        tags_after = ['textDirection', 'tcFitText', 'vAlign', 'hideMark']
                        inserted = False
                        for tag in tags_after:
                            el = tcPr.find(f'{{{ns}}}{tag}')
                            if el is not None:
                                el.addprevious(tcMar)
                                inserted = True
                                break
                        if not inserted:
                            tcPr.append(tcMar)
                    else:
                        for side in ['left', 'right']:
                            side_elem = tcMar.find(f'{{{ns}}}{side}')
                            if side_elem is not None:
                                side_elem.set(f'{{{ns}}}w', '0')
                                side_elem.set(f'{{{ns}}}type', 'dxa')
                            else:
                                tcMar.append(parse_xml(f'<w:{side} xmlns:w="{ns}" w:w="0" w:type="dxa"/>'))

                    for p in cell.paragraphs:
                        # Use paragraph alignment as well for content alignment
                        p.alignment = alignment

    if progress_cb and (count_indent > 0 or count_content > 0 or count_auto_fit > 0):
        table_msg = f'已调整表格样式：处理缩进 {count_indent} 个，处理内容 {count_content} 个'
        if count_auto_fit > 0:
            table_msg += f'，自适应 {count_auto_fit} 个'
        progress_cb(95, table_msg, 'success')
    if progress_cb:
        progress_cb(96, '正在处理图片样式...', 'dynamic')
    count_resized = 0
    count_centered = 0
    for i, p in enumerate(doc.paragraphs):
        if i < cover_para_count:
            continue
        resized, aligned = _apply_image_style_to_paragraph(p, ns, ns_wp, max_height, max_width, target_align)
        count_resized += resized
        count_centered += aligned
    if progress_cb and (count_resized > 0 or count_centered > 0 or count_table_image_resized > 0 or count_table_image_aligned > 0):
        align_desc = {0: '左对齐', 1: '居中对齐', 2: '右对齐'}.get(target_align, '居中对齐')
        image_msg = f'已调整图片样式：全文处理大小 {count_resized} 张，{align_desc} {count_centered} 张'
        if count_table_image_resized > 0 or count_table_image_aligned > 0:
            table_align_desc = {0: '左对齐', 1: '居中对齐', 2: '右对齐'}.get(table_target_align, '居中对齐')
            image_msg += f'；表格内处理大小 {count_table_image_resized} 张，{table_align_desc} {count_table_image_aligned} 张'
        progress_cb(97, image_msg, 'success')

    if progress_cb:
        progress_cb(98, '正在清理文本缩进并应用样式...', 'dynamic')
        
    # 处理列表缩进（修改 numbering_part 中的缩进设置）
    try:
        numbering_part = doc.part.numbering_part
        if numbering_part:
            for abstractNum in numbering_part._element.findall('.//w:abstractNum', namespaces=numbering_part._element.nsmap):
                for lvl in abstractNum.findall('w:lvl', namespaces=numbering_part._element.nsmap):
                    p_style = lvl.find('w:pStyle', namespaces=numbering_part._element.nsmap)
                    p_style_id = p_style.get(qn('w:val')) if p_style is not None else None
                    pPr = lvl.find('w:pPr', namespaces=numbering_part._element.nsmap)
                    if pPr is not None:
                        ind = pPr.find('w:ind', namespaces=numbering_part._element.nsmap)
                        if ind is not None:
                            if p_style_id in heading_style_ids:
                                # Heading numbering belongs to the template; do not apply body/list first-line indent to it.
                                template_ind = template_heading_numbering_indents.get(p_style_id)
                                if template_ind is not None:
                                    for attr in list(ind.attrib):
                                        del ind.attrib[attr]
                                    ind.attrib.update(template_ind)
                                continue
                            # 保证 numbering 层级不会有强制缩进干扰段落层的首行缩进
                            for attr in list(ind.attrib):
                                del ind.attrib[attr]
                            # 在 OOXML 规范中，firstLine 和 hanging 是互斥的。如果 hanging 存在（即使为 0），
                            # firstLine 也可能会被忽略或者表现不一致。
                            # 用户明确要求“特殊格式：首行缩进 2 字符”以及“文本之前：0 字符”。
                            # 为此，我们彻底清理 hanging，并且不设置 hanging/hangingChars，只设置 left 和 firstLine
                            ind.set(qn('w:left'), '0')
                            ind.set(qn('w:leftChars'), '0')
                            ind.set(qn('w:firstLine'), '420')
                            ind.set(qn('w:firstLineChars'), '200')
    except Exception as e:
        logger.warning(f'处理列表缩进失败: {e}')
    
    count_style_applied = 0
    for i, p in enumerate(doc.paragraphs):
        if i < cover_para_count:
            continue
        
        style = p.style
        style_name = style.name if style else ""
        is_heading = _is_heading_paragraph(p, ns)
        
        is_list = False
        if not is_heading and p._element.pPr is not None and p._element.pPr.numPr is not None:
            is_list = True
        elif not is_heading:
            curr_style = style
            while curr_style is not None:
                if curr_style.name and curr_style.name.startswith('List'):
                    is_list = True
                    break
                if hasattr(curr_style, '_element') and curr_style._element.pPr is not None and curr_style._element.pPr.numPr is not None:
                    is_list = True
                    break
                curr_style = curr_style.base_style
        
        # 飞书列表默认会有 w:left 和 w:hanging 缩进，导致“文本之前 1.48cm”
        # 按照用户需求，无序列表（或所有列表）应该像正文一样，只保留首行缩进，不要左缩进
        # 但是这里不再手动调用 _clean_text_indent，因为 _clean_text_indent 会把所有缩进相关的 tag 强行置 0，
        # 会影响到后面我们设置 firstLine 和 left 的效果。
        
        if is_heading:
            _clean_text_indent(p, ns)
        elif is_list:
            # 列表缩进已经在 numbering_part 中统一处理，此处不需要再修改段落上的 ind 属性
            # 但为防万一，清除段落上的 left 和 hanging 属性（保留 numbering 的效果）
            pPr = p._element.get_or_add_pPr()
            ind = pPr.find(f'{{{ns}}}ind')
            if ind is None:
                ind = pPr.get_or_add_ind()
                
            # 首行缩进会使得列表符号（项目符号）向右移动，这对于无序列表通常不符合常规。
            # Word 默认处理下，如果不设左侧缩进，项目符号就在页面最左边缘。如果要有缩进，应该是文本和符号整体缩进。
            # 这里我们仍然应用首行缩进，但同时需要处理左缩进（left）来保证符合要求
            # 为了达到首行缩进 2 字符效果并且让符号也缩进，我们将首行缩进 w:firstLineChars 设为 200。
            # 同时也保持悬挂缩进 w:hanging 存在，以保证多行文本对齐。
            for attr in list(ind.attrib):
                del ind.attrib[attr]
            
            # 正文首行缩进2字符效果，对于列表意味着：
            # 符号从左侧缩进2字符（leftChars=200），后续行与符号对齐（hangingChars=0）？
            # 还是说文本缩进2字符，符号在前面？
            # 用户要求“首行缩进2字符”，这通常表示：
            # 左侧无额外缩进(left=0)，首行由于有项目符号，我们使用首行缩进（firstLineChars=200）
            # 注意：在 Word 的机制中，如果要达到“首行缩进 2 字符”并与正文一致：
            # 我们直接将段落的缩进设置与正文一样，让 Word 的样式引擎负责处理项目符号
            # 用户明确要求：
            # 文本之前：0 字符
            # 特殊格式：首行缩进，度量值：2 字符
            # OOXML 中 hanging 和 firstLine 互斥，如果要强制首行缩进生效，绝不能设置 hanging。
            ind.set(f'{{{ns}}}left', '0')
            ind.set(f'{{{ns}}}leftChars', '0')
            ind.set(f'{{{ns}}}firstLine', '420')
            ind.set(f'{{{ns}}}firstLineChars', '200')
        else:
            _clean_text_indent(p, ns)
        
        # 应用正文样式
        if body_style:
            # 排除标题样式 (Heading 1-9, Title, Subtitle)
            # is_heading is computed from the complete paragraph style chain above.
            # 排除列表 (可选，这里暂时不排除列表，让列表也应用字体大小，但缩进可能受影响，需谨慎)
            # 这里的 _clean_text_indent 已经处理了缩进。
            # 飞书列表通常是 List Paragraph。
            if not is_heading:
                _apply_paragraph_style(p, body_style, ns)
                count_style_applied += 1
                
        # 强制移除标题的直接编号属性（防止转换时直接附加了编号）
        if ignore_template_heading_num:
            if is_heading:
                pPr = p._element.get_or_add_pPr()
                numPr = pPr.find(f"{{{ns}}}numPr")
                if numPr is not None:
                    pPr.remove(numPr)

    if progress_cb:
        if body_style:
            progress_cb(98, f'已调整文本样式 (应用正文样式到 {count_style_applied} 段)', 'success')
        else:
            progress_cb(98, '已调整文本样式', 'success')

    # Apply Page Margins
    if margin_config:
        if progress_cb:
            progress_cb(99, '正在调整页边距...', 'dynamic')
        
        # 跳过第一节（通常是封面或首页）
        num_sections = len(doc.sections)
        if num_sections > 1:
            for i in range(1, num_sections):
                section = doc.sections[i]
                if 'top' in margin_config: section.top_margin = Cm(margin_config['top'])
                if 'bottom' in margin_config: section.bottom_margin = Cm(margin_config['bottom'])
                if 'left' in margin_config: section.left_margin = Cm(margin_config['left'])
                if 'right' in margin_config: section.right_margin = Cm(margin_config['right'])
            if progress_cb:
                progress_cb(99, f'已调整页边距 (应用至 {num_sections - 1} 个分节)', 'success')
        else:
            # 如果只有一节且用户没有勾选“添加封面”，可能用户还是希望调整。
            # 但 instruction 明确说“不含封面/首页”，通常意味着从第二页开始。
            # 如果只有一页，则不调整以符合“不含首页”的逻辑。
            if progress_cb:
                progress_cb(99, '已跳过页边距调整 (文档只有一页)', 'info')

    doc.save(docx_path)
    del doc

def apply_document_info(docx_path, document_info):
    core_xml = _update_core_properties_xml(docx_path, document_info)
    app_xml = _update_app_properties_xml(docx_path, document_info)
    settings_xml, settings_rels_xml = _update_settings_template_parts(docx_path, document_info)
    if core_xml is None and app_xml is None and settings_xml is None and settings_rels_xml is None:
        return
    with zipfile.ZipFile(docx_path, 'r') as src:
        entries = [(info, src.read(info.filename)) for info in src.infolist()]
    tmp_path = docx_path + '.metadata.tmp'
    with zipfile.ZipFile(tmp_path, 'w') as dst:
        written_names = set()
        for info, content in entries:
            if info.filename == 'docProps/core.xml' and core_xml is not None:
                dst.writestr(info, core_xml)
                written_names.add(info.filename)
            elif info.filename == 'docProps/app.xml' and app_xml is not None:
                dst.writestr(info, app_xml)
                written_names.add(info.filename)
            elif info.filename == 'word/settings.xml' and settings_xml is not None:
                dst.writestr(info, settings_xml)
                written_names.add(info.filename)
            elif info.filename == 'word/_rels/settings.xml.rels' and settings_rels_xml is not None:
                dst.writestr(info, settings_rels_xml)
                written_names.add(info.filename)
            else:
                dst.writestr(info, content)
                written_names.add(info.filename)
        if settings_xml is not None and 'word/settings.xml' not in written_names:
            dst.writestr('word/settings.xml', settings_xml)
        if settings_rels_xml is not None and 'word/_rels/settings.xml.rels' not in written_names:
            dst.writestr('word/_rels/settings.xml.rels', settings_rels_xml)
    os.replace(tmp_path, docx_path)

def _update_core_properties_xml(docx_path, document_info):
    namespaces = {
        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'dcterms': 'http://purl.org/dc/terms/',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    with zipfile.ZipFile(docx_path, 'r') as zf:
        try:
            root = ET.fromstring(zf.read('docProps/core.xml'))
        except KeyError:
            return None
    text_fields = {
        'author': f"{{{namespaces['dc']}}}creator",
        'lastModifiedBy': f"{{{namespaces['cp']}}}lastModifiedBy",
        'title': f"{{{namespaces['dc']}}}title",
        'category': f"{{{namespaces['cp']}}}category",
        'subject': f"{{{namespaces['dc']}}}subject"
    }
    for key, tag in text_fields.items():
        _set_xml_text(root, tag, document_info.get(key, ''))
    datetime_fields = {
        'created': f"{{{namespaces['dcterms']}}}created",
        'modified': f"{{{namespaces['dcterms']}}}modified",
        'lastPrinted': f"{{{namespaces['cp']}}}lastPrinted"
    }
    for key, tag in datetime_fields.items():
        normalized = _normalize_document_datetime(document_info.get(key))
        if normalized:
            extra_attrib = None
            if key in ['created', 'modified']:
                extra_attrib = {f"{{{namespaces['xsi']}}}type": 'dcterms:W3CDTF'}
            _set_xml_text(root, tag, normalized, extra_attrib)
        else:
            _remove_xml_tag(root, tag)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def _update_app_properties_xml(docx_path, document_info):
    app_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
    vt_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
    ET.register_namespace('', app_ns)
    ET.register_namespace('vt', vt_ns)
    with zipfile.ZipFile(docx_path, 'r') as zf:
        try:
            root = ET.fromstring(zf.read('docProps/app.xml'))
        except KeyError:
            return None
    _set_xml_text(root, f'{{{app_ns}}}Template', document_info.get('template', ''))
    _set_xml_text(root, f'{{{app_ns}}}Company', document_info.get('company', ''))
    total_time = document_info.get('totalTime')
    if total_time in [None, '']:
        _remove_xml_tag(root, f'{{{app_ns}}}TotalTime')
    else:
        try:
            safe_total_time = str(max(0, int(total_time)))
        except (TypeError, ValueError):
            _remove_xml_tag(root, f'{{{app_ns}}}TotalTime')
        else:
            _set_xml_text(root, f'{{{app_ns}}}TotalTime', safe_total_time)
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def _update_settings_template_parts(docx_path, document_info):
    template_value = str(document_info.get('template') or '').strip()
    with zipfile.ZipFile(docx_path, 'r') as zf:
        try:
            settings_xml = zf.read('word/settings.xml').decode('utf-8')
        except KeyError:
            return None, None
        try:
            settings_rels_xml = zf.read('word/_rels/settings.xml.rels').decode('utf-8')
        except KeyError:
            settings_rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    attached_template_pattern = r'<w:attachedTemplate\b[^>]*/>'
    if template_value:
        replacement = '<w:attachedTemplate r:id="rIdAttachedTemplate"/>'
        if re.search(attached_template_pattern, settings_xml):
            settings_xml = re.sub(attached_template_pattern, replacement, settings_xml, count=1)
        else:
            settings_xml = settings_xml.replace('>', f'>{replacement}', 1)
    else:
        settings_xml = re.sub(attached_template_pattern, '', settings_xml)
    rel_pattern = r'<Relationship\b[^>]*Type="http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/attachedTemplate"[^>]*/>'
    settings_rels_xml = re.sub(rel_pattern, '', settings_rels_xml)
    if template_value:
        escaped_template_value = escape(template_value, {'"': '&quot;'})
        relation_xml = f'<Relationship Id="rIdAttachedTemplate" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="{escaped_template_value}" TargetMode="External"/>'
        settings_rels_xml = settings_rels_xml.replace('</Relationships>', relation_xml + '</Relationships>')
    return settings_xml.encode('utf-8'), settings_rels_xml.encode('utf-8')

def _set_xml_text(root, tag, value, extra_attrib=None):
    element = root.find(tag)
    if element is None:
        element = ET.SubElement(root, tag)
    element.text = '' if value is None else str(value)
    if extra_attrib:
        for key, attr_value in extra_attrib.items():
            element.set(key, attr_value)

def _remove_xml_tag(root, tag):
    element = root.find(tag)
    if element is not None:
        root.remove(element)

def _normalize_document_datetime(value):
    value = str(value or '').strip()
    if not value:
        return ''
    try:
        dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
    except ValueError:
        return ''
    if dt.tzinfo is None:
        local_tz = datetime.now().astimezone().tzinfo or timezone.utc
        dt = dt.replace(tzinfo=local_tz)
    return dt.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

def _apply_paragraph_style(paragraph, style_config, ns):
    """
    style_config: {
        'fontSize': float (pt),
        'lineSpacing': float,
        'lineSpacingUnit': 'lines'|'pt',
        'spaceBefore': float,
        'spaceBeforeUnit': 'lines'|'pt',
        'spaceAfter': float,
        'spaceAfterUnit': 'lines'|'pt'
    }
    """
    # 1. Font Size (Apply to all runs to override existing formatting)
    font_size = style_config.get('fontSize')
    if font_size:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
    
    # 2. Line Spacing
    line_spacing = style_config.get('lineSpacing')
    line_spacing_unit = style_config.get('lineSpacingUnit')
    if line_spacing is not None:
        pf = paragraph.paragraph_format
        if line_spacing_unit == 'pt':
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(line_spacing)
        else: # lines (default or multiple)
            # 1.5倍行距等
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            pf.line_spacing = line_spacing

    # 3. Space Before/After (Directly manipulating XML for 'lines' unit support)
    space_before = style_config.get('spaceBefore')
    space_before_unit = style_config.get('spaceBeforeUnit')
    space_after = style_config.get('spaceAfter')
    space_after_unit = style_config.get('spaceAfterUnit')
    
    if space_before is not None or space_after is not None:
        pPr = paragraph._element.get_or_add_pPr()
        spacing = pPr.find(f'{{{ns}}}spacing')
        if spacing is None:
            spacing = parse_xml(f'<w:spacing xmlns:w="{ns}"/>')
            pPr.append(spacing)
            
        if space_before is not None:
            if space_before_unit == 'lines':
                # w:beforeLines is in 100th of a line
                spacing.set(f'{{{ns}}}beforeLines', str(int(space_before * 100)))
                # Remove pt setting if exists to avoid conflict
                if f'{{{ns}}}before' in spacing.attrib:
                    del spacing.attrib[f'{{{ns}}}before']
            else: # pt
                # w:before is in twips (1/20 pt)
                spacing.set(f'{{{ns}}}before', str(int(space_before * 20)))
                if f'{{{ns}}}beforeLines' in spacing.attrib:
                    del spacing.attrib[f'{{{ns}}}beforeLines']
                    
        if space_after is not None:
            if space_after_unit == 'lines':
                spacing.set(f'{{{ns}}}afterLines', str(int(space_after * 100)))
                if f'{{{ns}}}after' in spacing.attrib:
                    del spacing.attrib[f'{{{ns}}}after']
            else: # pt
                spacing.set(f'{{{ns}}}after', str(int(space_after * 20)))
                if f'{{{ns}}}afterLines' in spacing.attrib:
                    del spacing.attrib[f'{{{ns}}}afterLines']

def _is_heading_style(style_element, ns_w):
    style_id = (style_element.get(f"{{{ns_w}}}styleId") or "").lower()
    if style_id.startswith('heading') or 'title' in style_id:
        return True
    name_elem = style_element.find(f"{{{ns_w}}}name")
    if name_elem is not None:
        val = (name_elem.get(f"{{{ns_w}}}val") or "").lower()
        if val.startswith('heading') or val.startswith("\u6807\u9898") or val.startswith("\u526f\u6807\u9898") or 'title' in val or 'subtitle' in val:
            return True
    return False

def _get_heading_style_ids(doc, ns_w):
    heading_style_ids = set()
    try:
        for style in doc.styles:
            style_element = getattr(style, '_element', None)
            if style_element is None:
                continue
            if _is_heading_style(style_element, ns_w):
                style_id = style_element.get(f"{{{ns_w}}}styleId")
                if style_id:
                    heading_style_ids.add(style_id)
    except Exception as e:
        logger.debug(f'Failed to collect heading style ids: {e}')
    return heading_style_ids

def _get_template_heading_numbering_indents(template_path, heading_style_ids, ns_w):
    if not template_path or not heading_style_ids or not os.path.exists(template_path):
        return {}
    try:
        with zipfile.ZipFile(template_path) as z:
            if 'word/numbering.xml' not in z.namelist():
                return {}
            root = ET.fromstring(z.read('word/numbering.xml'))
    except Exception as e:
        logger.debug(f'Failed to read template heading numbering indents: {e}')
        return {}

    result = {}
    for lvl in root.findall(f'.//{{{ns_w}}}lvl'):
        p_style = lvl.find(f'{{{ns_w}}}pStyle')
        p_style_id = p_style.get(f'{{{ns_w}}}val') if p_style is not None else None
        if p_style_id not in heading_style_ids or p_style_id in result:
            continue
        pPr = lvl.find(f'{{{ns_w}}}pPr')
        ind = pPr.find(f'{{{ns_w}}}ind') if pPr is not None else None
        result[p_style_id] = dict(ind.attrib) if ind is not None else {}
    return result

def _is_heading_paragraph(paragraph, ns_w):
    try:
        style = paragraph.style
    except Exception:
        style = None
    visited = set()
    while style is not None:
        style_element = getattr(style, '_element', None)
        if style_element is not None and _is_heading_style(style_element, ns_w):
            return True
        style_key = id(style)
        if style_key in visited:
            break
        visited.add(style_key)
        try:
            style = style.base_style
        except Exception:
            break
    return False

def _copy_styles_from_template(template_path, target_doc, ignore_template_heading_num=False):
    try:
        tpl_doc = Document(template_path)
        _ = tpl_doc.styles
        _ = target_doc.styles
        src_styles = tpl_doc.styles
        dst_styles = target_doc.styles
        if not hasattr(src_styles, '_element') or not hasattr(dst_styles, '_element'):
            return
        src_root = src_styles._element
        dst_root = dst_styles._element
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        count = 0
        
        # If ignore_template_heading_num is True, strip out <w:numPr> from existing Heading styles in target_doc
        if ignore_template_heading_num:
            for dst_style in dst_root:
                if not dst_style.tag.endswith('style'):
                    continue
                if _is_heading_style(dst_style, ns['w']):
                    dst_pPr = dst_style.find(f"{{{ns['w']}}}pPr")
                    if dst_pPr is not None:
                        dst_numPr = dst_pPr.find(f"{{{ns['w']}}}numPr")
                        if dst_numPr is not None:
                            dst_pPr.remove(dst_numPr)

        for style in src_root:
            if not style.tag.endswith('style'):
                continue
            style_id = style.get(f"{{{ns['w']}}}styleId")
            if not style_id:
                continue
                
            # If ignore_template_heading_num is True, strip out <w:numPr> from Heading styles in src before copying
            if ignore_template_heading_num and _is_heading_style(style, ns['w']):
                pPr = style.find(f"{{{ns['w']}}}pPr")
                if pPr is not None:
                    numPr = pPr.find(f"{{{ns['w']}}}numPr")
                    if numPr is not None:
                        pPr.remove(numPr)
                        
            found = False
            for dst_style in dst_root:
                if dst_style.tag.endswith('style') and dst_style.get(f"{{{ns['w']}}}styleId") == style_id:
                    found = True
                    break
            if not found:
                new_style = deepcopy(style)
                dst_root.append(new_style)
                count += 1
    except Exception as e:
        logger.error(f'复制样式失败: {e}')

def _add_image_to_part(target_part, image_stream, filename=None):
    rId = None
    if hasattr(target_part, 'get_or_add_image'):
        try:
            if image_stream.seekable():
                image_stream.seek(0)
            if filename and hasattr(image_stream, 'name'):
                try:
                    image_stream.name = filename
                except:
                    pass
            result = target_part.get_or_add_image(image_stream)
            if isinstance(result, tuple):
                for item in result:
                    if isinstance(item, str) and item.startswith('rId'):
                        rId = item
                        break
                if not rId and len(result) > 0:
                    rId = str(result[0])
            else:
                rId = result
            if rId:
                return str(rId)
        except Exception as e:
            pass
    try:
        package = getattr(target_part, 'package', None)
        if not package:
            if hasattr(target_part, 'part') and hasattr(target_part.part, 'package'):
                package = target_part.part.package
        if not package:
            return None
        if image_stream.seekable():
            image_stream.seek(0)
        image_part = None
        if hasattr(package, 'get_or_add_image_part'):
            try:
                if filename:
                    try:
                        image_stream.name = filename
                    except:
                        pass
                image_part = package.get_or_add_image_part(image_stream)
            except Exception as e_pkg:
                pass
        if image_part:
            rel_type = RT.IMAGE
            if hasattr(target_part, 'relate_to'):
                rId = target_part.relate_to(image_part, rel_type)
                return str(rId)
            else:
                pass
        else:
            pass
    except Exception as e:
        pass
    return None

def _copy_related_parts(element, source_part, target_part):
    for node in element.iter():
        for key, value in list(node.attrib.items()):
            value = value.strip() if value else ''
            if (key.endswith('embed') or key.endswith('id')) and value:
                if value in source_part.rels:
                    rel = source_part.rels[value]
                    target_rId = None
                    try:
                        if 'image' in rel.reltype:
                            if not rel.is_external:
                                try:
                                    blob = rel.target_part.blob
                                    max_depth = 3
                                    while isinstance(blob, tuple) and max_depth > 0:
                                        found_bytes = False
                                        for item in blob:
                                            if isinstance(item, bytes):
                                                blob = item
                                                found_bytes = True
                                                break
                                            elif isinstance(item, tuple):
                                                blob = item
                                                found_bytes = True
                                                break
                                        if not found_bytes and len(blob) > 0:
                                            blob = blob[0]
                                        max_depth -= 1
                                    if blob and isinstance(blob, bytes):
                                        image_stream = io.BytesIO(blob)
                                        filename = None
                                        try:
                                            if hasattr(rel.target_part, 'partname'):
                                                filename = os.path.basename(rel.target_part.partname)
                                            elif hasattr(rel.target_part, 'filename'):
                                                filename = rel.target_part.filename
                                        except:
                                            pass
                                        target_rId = _add_image_to_part(target_part, image_stream, filename)
                                    else:
                                        pass
                                except Exception as e:
                                    pass
                        elif 'hyperlink' in rel.reltype and rel.is_external:
                            target_rId = target_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                    except Exception as e:
                        pass
                    if target_rId:
                        if isinstance(target_rId, tuple):
                            for item in target_rId:
                                if isinstance(item, str) and item.startswith('rId'):
                                    target_rId = item
                                    break
                            if isinstance(target_rId, tuple) and len(target_rId) > 0:
                                target_rId = str(target_rId[0])
                        try:
                            node.set(key, str(target_rId))
                        except Exception as e_set:
                            pass

def _get_header_element(header):
    for attr in ('_element', 'element', '_hdr', '_ftr'):
        el = getattr(header, attr, None)
        if el is not None:
            return el
    for part_attr in ('part', '_part', '_hdr_part', '_ftr_part', '_header_part', '_footer_part'):
        part = getattr(header, part_attr, None)
        if part is None:
            continue
        for el_attr in ('_element', 'element', '_hdr', '_ftr'):
            el = getattr(part, el_attr, None)
            if el is not None:
                return el
    return None

def _copy_header_footer_content(source_header, target_header):
    _ = target_header.is_linked_to_previous
    target_root = _get_header_element(target_header)
    source_root = _get_header_element(source_header)
    if target_root is None or source_root is None:
        source_part = getattr(source_header, 'part', None)
        target_part = getattr(target_header, 'part', None)
        if source_part is not None and target_part is not None and hasattr(source_part, 'blob'):
            try:
                source_el = parse_xml(source_part.blob)
                new_el = deepcopy(source_el)
                _copy_related_parts(new_el, source_part, target_part)
                if hasattr(target_part, '_element'):
                    target_part._element = new_el
                if hasattr(target_part, '_blob'):
                    target_part._blob = new_el.xml
                return
            except Exception as e:
                pass
        return
    for child in list(target_root):
        target_root.remove(child)
    for child in source_root:
        new_child = deepcopy(child)
        s_part = getattr(source_header, 'part', None)
        t_part = getattr(target_header, 'part', None)
        if s_part and t_part:
            _copy_related_parts(new_child, s_part, t_part)
        target_root.append(new_child)

def _copy_headers_from_template(tpl_doc, doc, add_cover=True):
    if not tpl_doc.sections or not doc.sections:
        return

    num_tpl_sections = len(tpl_doc.sections)
    num_doc_sections = len(doc.sections)
    cover_tpl_section = tpl_doc.sections[0]
    body_tpl_section = tpl_doc.sections[1] if num_tpl_sections > 1 else tpl_doc.sections[0]

    if add_cover:
        _copy_section_headers_footers(cover_tpl_section, doc.sections[0])
        if num_doc_sections > 1:
            for i in range(1, num_doc_sections):
                _copy_section_headers_footers(body_tpl_section, doc.sections[i])
    else:
        for i in range(num_doc_sections):
            _copy_section_headers_footers(body_tpl_section, doc.sections[i])

def _copy_section_headers_footers(source_section, target_section):
    if hasattr(target_section.header, 'is_linked_to_previous'):
        target_section.header.is_linked_to_previous = False
    if hasattr(target_section.footer, 'is_linked_to_previous'):
        target_section.footer.is_linked_to_previous = False

    _copy_header_footer_content(source_section.header, target_section.header)
    _copy_header_footer_content(source_section.footer, target_section.footer)

    if getattr(source_section, 'different_first_page_header_footer', False):
        target_section.different_first_page_header_footer = True
        if hasattr(target_section.first_page_header, 'is_linked_to_previous'):
            target_section.first_page_header.is_linked_to_previous = False
        if hasattr(target_section.first_page_footer, 'is_linked_to_previous'):
            target_section.first_page_footer.is_linked_to_previous = False
        _copy_header_footer_content(source_section.first_page_header, target_section.first_page_header)
        _copy_header_footer_content(source_section.first_page_footer, target_section.first_page_footer)

    if getattr(source_section, 'odd_and_even_pages_header_footer', False):
        target_section.odd_and_even_pages_header_footer = True
        if hasattr(target_section.even_page_header, 'is_linked_to_previous'):
            target_section.even_page_header.is_linked_to_previous = False
        if hasattr(target_section.even_page_footer, 'is_linked_to_previous'):
            target_section.even_page_footer.is_linked_to_previous = False
        _copy_header_footer_content(source_section.even_page_header, target_section.even_page_header)
        _copy_header_footer_content(source_section.even_page_footer, target_section.even_page_footer)

def _prepend_first_page_from_template(doc, template_path):
    try:
        tpl_doc = Document(template_path)
    except Exception:
        return
    elements_to_copy = []
    body = tpl_doc._element.body
    found_break = False
    source_part = tpl_doc.part
    target_part = doc.part
    for element in body:
        if element.tag.endswith('sectPr'):
            continue
        if element.tag.endswith('p'):
            pPr = element.find(f"{{{element.nsmap['w']}}}pPr")
            if pPr is not None:
                sectPr = pPr.find(f"{{{element.nsmap['w']}}}sectPr")
                if sectPr is not None:
                    new_elem = deepcopy(element)
                    _copy_related_parts(new_elem, source_part, target_part)
                    elements_to_copy.append(new_elem)
                    found_break = True
                    break
        new_elem = deepcopy(element)
        _copy_related_parts(new_elem, source_part, target_part)
        elements_to_copy.append(new_elem)
    target_body = doc._element.body
    import random
    count_paragraphs = 0
    count_tables = 0
    for i, element in enumerate(reversed(elements_to_copy)):
        for drawing in element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'):
            drawing.set('id', str(random.randint(10000, 99999999)))
        target_body.insert(0, element)
        if element.tag.endswith('p'):
            count_paragraphs += 1
        elif element.tag.endswith('tbl'):
            count_tables += 1
    if not found_break:
        ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        p = parse_xml(f'\n        <w:p xmlns:w="{ns}">\n            <w:pPr>\n                <w:sectPr>\n                    <w:type w:val="nextPage"/>\n                </w:sectPr>\n            </w:pPr>\n        </w:p>\n        ')
        target_body.insert(len(elements_to_copy), p)
        count_paragraphs += 1
    try:
        if tpl_doc.sections and doc.sections:
            src_sect = tpl_doc.sections[0]
            dst_sect = doc.sections[0]
            dst_sect.page_height = src_sect.page_height
            dst_sect.page_width = src_sect.page_width
            dst_sect.left_margin = src_sect.left_margin
            dst_sect.right_margin = src_sect.right_margin
            dst_sect.top_margin = src_sect.top_margin
            dst_sect.bottom_margin = src_sect.bottom_margin
            dst_sect.header_distance = src_sect.header_distance
            dst_sect.footer_distance = src_sect.footer_distance
    except Exception as e:
        logger.debug(f'复制分节属性失败: {e}')
    return (count_paragraphs, count_tables)

def _resize_inline_image(drawing_element, max_height, max_width=None):
    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    extent = drawing_element.find(f'{{{ns_wp}}}extent')
    resized = False
    if extent is not None:
        cy = int(extent.get('cy') or 0)
        cx = int(extent.get('cx') or 0)
        if cy > max_height:
            ratio = max_height / cy
            cy = int(max_height)
            cx = int(cx * ratio)
            extent.set('cy', str(cy))
            extent.set('cx', str(cx))
            resized = True
        if max_width and cx > max_width:
            ratio = max_width / cx
            cx = int(max_width)
            cy = int(cy * ratio)
            extent.set('cx', str(cx))
            extent.set('cy', str(cy))
            resized = True
    return resized

def _force_clear_indent(paragraph, ns, clear_space_before=False):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(f'{{{ns}}}left', '0')
    ind.set(f'{{{ns}}}right', '0')
    ind.set(f'{{{ns}}}hanging', '0')
    ind.set(f'{{{ns}}}firstLine', '0')
    ind.set(f'{{{ns}}}leftChars', '0')
    ind.set(f'{{{ns}}}rightChars', '0')
    ind.set(f'{{{ns}}}firstLineChars', '0')
    ind.set(f'{{{ns}}}hangingChars', '0')
    
    # Check if there is an override in style itself if we are applying inline styles
    # Feishu sometimes passes jc directly or we inherit paragraph properties. 
    # Just to be safe, set justification left if not overriden properly later.
    jc = pPr.find(f'{{{ns}}}jc')
    if jc is not None and jc.get(f'{{{ns}}}val') == 'both':
        jc.set(f'{{{ns}}}val', 'left')
        
    # Clear paragraph margin properties which sometimes override indentation
    spacing = pPr.find(f'{{{ns}}}spacing')
    if spacing is not None:
        if clear_space_before:
            if f'{{{ns}}}before' in spacing.attrib:
                spacing.set(f'{{{ns}}}before', '0')
            if f'{{{ns}}}beforeLines' in spacing.attrib:
                spacing.set(f'{{{ns}}}beforeLines', '0')
        else:
            if spacing.get(f'{{{ns}}}beforeLines'):
                pass # Keep vertical spacing
        
        if spacing.get(f'{{{ns}}}afterLines'):
            pass # Keep vertical spacing
        # Clear horizontal spacing if mistakenly placed here
        pass

def _clean_text_indent(paragraph, ns):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    if ind.get(f'{{{ns}}}hanging'):
        ind.set(f'{{{ns}}}hanging', '0')
    if ind.get(f'{{{ns}}}hangingChars'):
        ind.set(f'{{{ns}}}hangingChars', '0')
    if ind.get(f'{{{ns}}}left'):
        ind.set(f'{{{ns}}}left', '0')
    if ind.get(f'{{{ns}}}leftChars'):
        ind.set(f'{{{ns}}}leftChars', '0')
    if ind.get(f'{{{ns}}}start'):
        ind.set(f'{{{ns}}}start', '0')
    if ind.get(f'{{{ns}}}startChars'):
        ind.set(f'{{{ns}}}startChars', '0')
    if ind.get(f'{{{ns}}}firstLine'):
        ind.set(f'{{{ns}}}firstLine', '0')
    if ind.get(f'{{{ns}}}firstLineChars'):
        ind.set(f'{{{ns}}}firstLineChars', '0')

def _apply_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if not tcBorders:
        tcBorders = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />')
        tcPr.append(tcBorders)
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            tag = f'w:{edge}'
            element = tcBorders.find(parse_xml(f'<{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag)
            if element is None:
                element = parse_xml(f'<{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />')
                tcBorders.append(element)
            element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'val'), val.get('val', 'single'))
            element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'sz'), str(val.get('sz', 4)))
            element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'space'), '0')
            element.set(parse_xml('<w:attr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />').tag.replace('attr', 'color'), val.get('color', 'auto'))

def _apply_shading(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color}"/>')
    existing = tcPr.find(shd.tag)
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(shd)

def _set_cell_text_color(cell, color_hex, bold=False):
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.color.rgb = RGBColor.from_string(color_hex)
            if bold:
                run.font.bold = True

def _apply_custom_code_block_style(table, config, ns):
    def hex_to_docx(h):
        return h.lstrip('#').upper()

    cell = table.cell(0, 0)
    table_width = config.get('tableWidth')
    if table_width not in (None, ''):
        _apply_table_layout(table, f"{table_width}cm", False, ns, min_col_width=0)
    else:
        # 默认不指定宽度时，强制清除表格固定宽度标签，允许自适应或按外部环境流式布局
        tblPr = table._element.tblPr
        if tblPr is not None:
            tbl_w = tblPr.find(f'{{{ns}}}tblW')
            if tbl_w is not None:
                tbl_w.set(f'{{{ns}}}w', '0')
                tbl_w.set(f'{{{ns}}}type', 'auto')
                
        for row in table.rows:
            for c in row.cells:
                tc_pr = c._tc.get_or_add_tcPr()
                tc_w = tc_pr.find(f'{{{ns}}}tcW')
                if tc_w is not None:
                    tc_w.set(f'{{{ns}}}w', '0')
                    tc_w.set(f'{{{ns}}}type', 'auto')
                else:
                    tc_w = parse_xml(f'<w:tcW xmlns:w="{ns}" w:w="0" w:type="auto"/>')
                    tags_after = ['gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']
                    inserted = False
                    for tag in tags_after:
                        el = tc_pr.find(f'{{{ns}}}{tag}')
                        if el is not None:
                            el.addprevious(tc_w)
                            inserted = True
                            break
                    if not inserted:
                        tc_pr.append(tc_w)

    bg_color = hex_to_docx(config.get('bgColor', '#F5F5F5'))
    _apply_shading(cell, bg_color)

    border_color = hex_to_docx(config.get('borderColor', '#D9D9D9'))
    borders = config.get('borders', {})
    
    _apply_border(cell, 
        top={'val': borders.get('top', {}).get('type', 'single'), 'sz': borders.get('top', {}).get('width', 4), 'color': border_color},
        bottom={'val': borders.get('bottom', {}).get('type', 'single'), 'sz': borders.get('bottom', {}).get('width', 4), 'color': border_color},
        left={'val': borders.get('left', {}).get('type', 'single'), 'sz': borders.get('left', {}).get('width', 4), 'color': border_color},
        right={'val': borders.get('right', {}).get('type', 'single'), 'sz': borders.get('right', {}).get('width', 4), 'color': border_color}
    )

    font_color = hex_to_docx(config.get('fontColor', '#000000'))
    font_family = config.get('fontFamily', 'Courier New')
    font_size = config.get('fontSize', 9)
    alignment = _align_to_docx(config.get('align', 'left'), 0)
    force_clear_indent = config.get('forceClearIndent', True)
    
    line_spacing = config.get('lineSpacing')
    space_before = config.get('spaceBefore')
    space_after = config.get('spaceAfter')
    
    for p in cell.paragraphs:
        if force_clear_indent:
            _force_clear_indent(p, ns)
        elif config.get('cleanTextIndent', False):
            _clean_text_indent(p, ns)
        p.alignment = alignment
        
        # 处理行间距、段前段后距
        if line_spacing is not None or space_before is not None or space_after is not None:
            p_pr = p._element.get_or_add_pPr()
            spacing = p_pr.find(f'{{{ns}}}spacing')
            if spacing is None:
                spacing = parse_xml(f'<w:spacing xmlns:w="{ns}"/>')
                p_pr.append(spacing)
                
            if line_spacing is not None:
                try:
                    # w:line 单位为 240 = 1 行
                    spacing.set(f'{{{ns}}}line', str(int(float(line_spacing) * 240)))
                    spacing.set(f'{{{ns}}}lineRule', 'auto')
                except:
                    pass
                    
            if space_before is not None:
                try:
                    # w:beforeLines 单位为 100 = 1 行
                    spacing.set(f'{{{ns}}}beforeLines', str(int(float(space_before) * 100)))
                except:
                    pass
                    
            if space_after is not None:
                try:
                    # w:afterLines 单位为 100 = 1 行
                    spacing.set(f'{{{ns}}}afterLines', str(int(float(space_after) * 100)))
                except:
                    pass
                    
        for run in p.runs:
            run.font.name = font_family
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor.from_string(font_color)
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), font_family)

from src.converters.docx.style_manager import TableStyleManager
from docx.oxml.ns import qn
import logging
logger = logging.getLogger('doc_download')

def apply_custom_styles(doc, style_idx):
    if not style_idx:
        return
    logger.info(f'正在应用自定义样式，索引 {style_idx}')
    count_sheets = 0
    count_tables = 0
    for table in doc.tables:
        try:
            tblPr = table._element.tblPr
            if tblPr is not None:
                caption = tblPr.find(qn('w:tblCaption'))
                if caption is not None and caption.get(qn('w:val')) == 'sheet':
                    count_sheets += 1
        except Exception as e:
            logger.warning(f'检查表格标记错误: {e}')
        try:
            # 通过标签识别代码块，避免误伤 1x1 普通表格
            is_code_block = False
            try:
                tblPr = table._element.tblPr
                if tblPr is not None:
                    caption = tblPr.find(qn('w:tblCaption'))
                    if caption is not None and caption.get(qn('w:val')) == 'code_block':
                        is_code_block = True
            except:
                pass
                
            if is_code_block:
                continue
        except Exception as e:
            logger.warning(f'检查代码块错误: {e}')
            continue
        count_tables += 1
        TableStyleManager.apply_style(table, style_idx)
    logger.info(f'已应用样式于 {count_tables} 个表格和 {count_sheets} 个电子表格')

def list_table_styles():
    return TableStyleManager.list_styles()
