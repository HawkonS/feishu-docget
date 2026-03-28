from datetime import datetime
import os
import io
from docx import Document
from docx.oxml.ns import nsdecls
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

def clean_document(docx_path, progress_cb=None, template_path=None, add_cover=False, body_style=None, image_style=None, table_config=None, margin_config=None):
    logger.debug('开始清理文档...')
    doc = Document(docx_path)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ns_wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    
    if template_path:
        try:
            tpl_doc = Document(template_path)
            _copy_styles_from_template(template_path, doc)
            
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
    
    # Determine Image Config
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

    target_max_h = default_max_h
    target_max_w = default_max_w
    target_align = 1 # Center by default

    if image_style:
        target_max_w = float(image_style.get('maxWidth', default_max_w))
        target_max_h = float(image_style.get('maxHeight', default_max_h))
        align_str = image_style.get('align', 'center').lower()
        if align_str == 'left':
            target_align = 0
        elif align_str == 'right':
            target_align = 2
        else:
            target_align = 1
            
    MAX_HEIGHT = Mm(target_max_h * 10)
    MAX_WIDTH = Mm(target_max_w * 10)

    # Determine Table Config
    force_clear_tbl_indent = True
    header_align = None
    content_align = None
    
    if table_config:
        force_clear_tbl_indent = table_config.get('forceClearIndent', True)
        
        # Mapping frontend align strings to docx alignment values
        align_map = {'left': 0, 'center': 1, 'right': 2}
        header_align_str = table_config.get('headerAlign', 'center').lower()
        content_align_str = table_config.get('contentAlign', 'left').lower()
        
        header_align = align_map.get(header_align_str, 1)
        content_align = align_map.get(content_align_str, 0)

    if progress_cb:
        progress_cb(93, '正在处理表格样式...', 'dynamic')
    cover_table_count = inserted_count[1] if isinstance(inserted_count, tuple) else 0
    cover_para_count = inserted_count[0] if isinstance(inserted_count, tuple) else inserted_count if isinstance(inserted_count, int) else 0
    count_indent = 0
    for i, table in enumerate(doc.tables):
        if i < cover_table_count:
            continue
        count_indent += 1
        tblPr = table._element.tblPr
        if tblPr is not None:
            tblInd = tblPr.find(f'{{{ns}}}tblInd')
            if tblInd is not None:
                tblInd.set(f'{{{ns}}}w', '0')
                tblInd.set(f'{{{ns}}}type', 'dxa')
            else:
                tblInd = parse_xml(f'<w:tblInd xmlns:w="{ns}" w:w="0" w:type="dxa"/>')
                tblPr.append(tblInd)
            jc = tblPr.find(f'{{{ns}}}jc')
            if jc is not None:
                jc.set(f'{{{ns}}}val', 'left')
    count_content = 0
    for i, table in enumerate(doc.tables):
        if i < cover_table_count:
            continue
        count_content += 1
        
        # 通过标签识别代码块
        is_code_block = False
        try:
            tblPr = table._element.tblPr
            if tblPr is not None:
                caption = tblPr.find(f'{{{ns}}}tblCaption')
                if caption is not None and caption.get(f'{{{ns}}}val') == 'code_block':
                    is_code_block = True
        except:
            pass
        
        for r_idx, row in enumerate(table.rows):
            is_header = (r_idx == 0)
            for cell in row.cells:
                if not is_code_block:
                    cell.vertical_alignment = 1
                for p in cell.paragraphs:
                    if force_clear_tbl_indent:
                        # 强制清空所有缩进（包括首行缩进）
                        _force_clear_indent(p, ns)
                    else:
                        # 按照正文样式处理（仅清空悬挂缩进和左缩进，保留首行缩进）
                        _clean_text_indent(p, ns)
                    
                    if is_code_block:
                        p.alignment = 0
                    else:
                        if table_config:
                            # Apply alignment if table_config is provided
                            if is_header:
                                p.alignment = header_align
                            else:
                                p.alignment = content_align
                        
                        # 如果提供了正文样式配置，也应用到表格段落中（使其与正文一致）
                        if body_style:
                            _apply_paragraph_style(p, body_style, ns)

    if progress_cb and (count_indent > 0 or count_content > 0):
        progress_cb(95, f'已调整表格样式：处理缩进 {count_indent} 个，处理内容 {count_content} 个', 'success')
    if progress_cb:
        progress_cb(96, '正在处理图片样式...', 'dynamic')
    count_resized = 0
    count_centered = 0
    for i, p in enumerate(doc.paragraphs):
        if i < cover_para_count:
            continue
        xml = p._element.xml
        if 'w:drawing' in xml or 'w:pict' in xml:
            _force_clear_indent(p, ns)
            p.alignment = target_align
            count_centered += 1
        for run in p.runs:
            drawings = run._element.findall(f'.//{{{ns}}}drawing')
            for drawing in drawings:
                inlines = drawing.findall(f'.//{{{ns_wp}}}inline')
                for inline in inlines:
                    if _resize_inline_image(inline, MAX_HEIGHT, MAX_WIDTH):
                        count_resized += 1
                anchors = drawing.findall(f'.//{{{ns_wp}}}anchor')
                for anchor in anchors:
                    if _resize_inline_image(anchor, MAX_HEIGHT, MAX_WIDTH):
                        count_resized += 1
    if progress_cb and (count_resized > 0 or count_centered > 0):
        align_desc = {0: '左对齐', 1: '居中对齐', 2: '右对齐'}.get(target_align, '居中对齐')
        progress_cb(97, f'已调整图片样式：处理大小 {count_resized} 张，{align_desc} {count_centered} 张', 'success')

    if progress_cb:
        progress_cb(98, '正在清理文本缩进并应用样式...', 'dynamic')
    
    count_style_applied = 0
    for i, p in enumerate(doc.paragraphs):
        if i < cover_para_count:
            continue
        _clean_text_indent(p, ns)
        
        # 应用正文样式
        if body_style:
            # 排除标题样式 (Heading 1-9, Title, Subtitle)
            style_name = p.style.name
            is_heading = style_name.startswith('Heading') or style_name in ['Title', 'Subtitle']
            # 排除列表 (可选，这里暂时不排除列表，让列表也应用字体大小，但缩进可能受影响，需谨慎)
            # 这里的 _clean_text_indent 已经处理了缩进。
            # 飞书列表通常是 List Paragraph。
            if not is_heading:
                _apply_paragraph_style(p, body_style, ns)
                count_style_applied += 1

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

def _copy_styles_from_template(template_path, target_doc):
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
        for style in src_root:
            if not style.tag.endswith('style'):
                continue
            style_id = style.get(f"{{{ns['w']}}}styleId")
            if not style_id:
                continue
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

def _force_clear_indent(paragraph, ns):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(f'{{{ns}}}ind')
    if ind is None:
        ind = parse_xml(f'<w:ind xmlns:w="{ns}"/>')
        pPr.append(ind)
    ind.set(f'{{{ns}}}left', '0')
    ind.set(f'{{{ns}}}right', '0')
    ind.set(f'{{{ns}}}hanging', '0')
    ind.set(f'{{{ns}}}firstLine', '0')
    ind.set(f'{{{ns}}}firstLineChars', '0')

def _clean_text_indent(paragraph, ns):
    pPr = paragraph._element.get_or_add_pPr()
    ind = pPr.find(f'{{{ns}}}ind')
    if ind is not None:
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
    else:
        ind = parse_xml(f'<w:ind xmlns:w="{ns}" w:left="0" w:hanging="0"/>')
        pPr.append(ind)

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
