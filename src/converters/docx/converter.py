import os
import concurrent.futures
import traceback
import json
import math
import re
from urllib.parse import unquote
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
import docx.opc.constants
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from src.core.config_loader import ConfigLoader
from src.core.image_processor import (
    calculate_center_crop,
    get_image_center_crop,
    get_image_dimensions,
    smart_crop,
)
from src.converters.docx.style_manager import TableStyleManager


def _is_svg_file(file_path):
    """检测文件是否为 SVG 格式"""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(512)
        if b'<svg' in header or (b'<?xml' in header and b'<svg' in f.read(2048)):
            return True
    except Exception:
        pass
    return False


_SVG_XML_TEMPLATE = (
    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
    ' xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:drawing>'
    '<wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:docPr id="{doc_pr_id}" name="SVG {doc_pr_id}"/>'
    '<a:graphic>'
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:pic>'
    '<pic:nvPicPr>'
    '<pic:cNvPr id="{doc_pr_id}" name="SVG {doc_pr_id}"/>'
    '<pic:cNvPicPr/>'
    '</pic:nvPicPr>'
    '<pic:blipFill>'
    '<a:blip r:embed="{png_rid}"><a:extLst>'
    '<a:ext uri="{svg_ext_uri}"><asvg:svgBlip r:embed="{svg_rid}"/>'
    '</a:ext></a:extLst></a:blip>'
    '{src_rect}'
    '<a:stretch><a:fillRect/></a:stretch>'
    '</pic:blipFill>'
    '<pic:spPr>'
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '</pic:spPr>'
    '</pic:pic>'
    '</a:graphicData>'
    '</a:graphic>'
    '</wp:inline>'
    '</w:drawing>'
    '</w:r>'
)
_SVG_EXT_URI = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'


def _apply_feishu_image_crop(inline_shape, image_path, image_data):
    """Apply a Feishu crop rectangle as a native DrawingML source rectangle.

    Docs AI exposes the editor crop as an absolute normalized rectangle.  The
    DrawingML ``srcRect`` element instead expects the margins outside that
    rectangle, so convert ``[left, top, right, bottom]`` accordingly.  Older
    block-only responses do not include this field; retain the centered-ratio
    fallback for those responses.
    """
    if not isinstance(image_data, dict):
        return False
    frame_width = image_data.get('width')
    frame_height = image_data.get('height')
    crop_rect = _normalize_feishu_crop(image_data.get('crop'))
    if crop_rect:
        crop = {
            'left': crop_rect[0],
            'top': crop_rect[1],
            'right': 1.0 - crop_rect[2],
            'bottom': 1.0 - crop_rect[3],
        }
    else:
        crop = get_image_center_crop(image_path, frame_width, frame_height)
    if not crop:
        return False

    try:
        frame_width = float(frame_width)
        frame_height = float(frame_height)
        blip_fill = inline_shape._inline.graphic.graphicData.pic.blipFill
        existing = blip_fill.find(qn('a:srcRect'))
        if existing is not None:
            blip_fill.remove(existing)

        src_rect = OxmlElement('a:srcRect')
        drawingml_names = {
            'left': 'l',
            'top': 't',
            'right': 'r',
            'bottom': 'b',
        }
        for name, attribute in drawingml_names.items():
            value = int(round(crop[name] * 100000))
            if value > 0:
                src_rect.set(attribute, str(value))

        stretch = blip_fill.find(qn('a:stretch'))
        if stretch is None:
            blip_fill.append(src_rect)
        else:
            blip_fill.insert(blip_fill.index(stretch), src_rect)

        # A source rectangle can have a different aspect ratio from the
        # original image frame.  Match the shape to the cropped source area so
        # Word does not stretch the remaining pixels into the old frame.
        if crop_rect:
            source_size = get_image_dimensions(image_path)
            source_width, source_height = source_size or (frame_width, frame_height)
            crop_width = crop_rect[2] - crop_rect[0]
            crop_height = crop_rect[3] - crop_rect[1]
            crop_ratio = (source_height * crop_height) / (source_width * crop_width)
            inline_shape.height = max(1, int(round(inline_shape.width * crop_ratio)))
        else:
            inline_shape.height = max(1, int(round(inline_shape.width * frame_height / frame_width)))
        return True
    except Exception as exc:
        logger.warning(f'应用飞书图片裁剪失败: {exc}')
        return False


def _normalize_feishu_crop(value):
    """Return a validated ``(left, top, right, bottom)`` crop rectangle."""
    if isinstance(value, str):
        raw = value.strip()
        try:
            value = json.loads(raw)
        except (TypeError, ValueError, json.JSONDecodeError):
            try:
                value = [float(item) for item in re.split(r'[,\s]+', raw.strip('[]')) if item]
            except (TypeError, ValueError):
                return None
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    try:
        rect = tuple(float(item) for item in value)
    except (TypeError, ValueError):
        return None
    left, top, right, bottom = rect
    if any(not math.isfinite(value) or value < 0 or value > 1 for value in rect):
        return None
    if right <= left or bottom <= top:
        return None
    return rect


def extract_image_crops_from_content(content):
    """Extract Docs AI image crop rectangles keyed by block id and token.

    The public block API omits crop metadata, while ``docs_ai`` rendered XML
    includes it on ``<img crop="[...]" ...>``.  Returning both identifiers
    makes this robust to gateways that omit either the image id or source token.
    """
    if not isinstance(content, str) or not content.strip():
        return {}
    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring('<root>' + content + '</root>')
    except Exception as exc:
        logger.warning(f'解析图片裁剪 XML 失败，跳过图片裁剪: {exc}')
        return {}

    result = {}
    for element in root.iter():
        if element.tag.rsplit('}', 1)[-1].lower() not in ('img', 'image'):
            continue
        crop = _normalize_feishu_crop(element.get('crop'))
        if not crop:
            continue
        identifiers = (element.get('id'), element.get('src'), element.get('token'))
        for identifier in identifiers:
            if identifier:
                result[identifier] = crop
    return result


def _get_svg_dimensions(svg_path, fallback_width_cm=15):
    """从 SVG 文件中提取宽高（像素），失败则返回默认值"""
    import re as _re
    try:
        with open(svg_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read(4096)
        w = h = None
        wm = _re.search(r'<svg[^>]*\bwidth=["\']([\d.]+)', content)
        hm = _re.search(r'<svg[^>]*\bheight=["\']([\d.]+)', content)
        if wm and hm:
            w, h = float(wm.group(1)), float(hm.group(1))
        else:
            vm = _re.search(r'viewBox=["\'][\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)', content)
            if vm:
                w, h = float(vm.group(1)), float(vm.group(2))
        if w and h and w > 0 and h > 0:
            return int(w * 9525), int(h * 9525)
    except Exception:
        pass
    cx = int(fallback_width_cm * 360000)
    cy = int(cx * 0.75)
    return cx, cy


def _create_placeholder_png(path, width=100, height=75):
    """创建最小的占位 PNG 文件（灰色矩形），供旧版 Word 回退显示"""
    try:
        from PIL import Image, ImageDraw
        img = Image.new('RGB', (width, height), (230, 230, 230))
        ImageDraw.Draw(img).text((10, height // 2 - 6), 'SVG', fill=(150, 150, 150))
        img.save(path, 'PNG')
        return True
    except Exception:
        pass
    # 最小有效 PNG（1x1 灰色像素）
    import struct, zlib
    def _chunk(ctype, data):
        c = ctype + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)
    ihdr = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
    raw = zlib.compress(b'\x00\xe6\xe6\xe6')
    png = b'\x89PNG\r\n\x1a\n' + _chunk(b'IHDR', ihdr) + _chunk(b'IDAT', raw) + _chunk(b'IEND', b'')
    with open(path, 'wb') as f:
        f.write(png)
    return True


def _add_svg_to_docx(run, svg_path, width_cm=15, image_data=None):
    """将 SVG 直接嵌入 DOCX（Word 2019+/M365 原生渲染），附带 PNG 回退"""
    from docx.opc.part import Part as OpcPart
    from docx.opc.packuri import PackURI

    part = run.part
    package = part.package

    # 创建 SVG Part 并添加关系
    svg_partname = PackURI('/word/media/' + os.path.basename(svg_path).replace('.png', '.svg'))
    with open(svg_path, 'rb') as f:
        svg_blob = f.read()
    svg_part = OpcPart(svg_partname, 'image/svg+xml', svg_blob, package)
    svg_rid = part.relate_to(svg_part, _SVG_EXT_URI)

    # 生成占位 PNG 并创建 Part
    placeholder = svg_path.rsplit('.', 1)[0] + '.fallback.png'
    if not os.path.exists(placeholder):
        _create_placeholder_png(placeholder)
    png_partname = PackURI('/word/media/' + os.path.basename(placeholder))
    with open(placeholder, 'rb') as f:
        png_blob = f.read()
    png_part = OpcPart(png_partname, 'image/png', png_blob, package)
    png_rid = part.relate_to(png_part, docx.opc.constants.RELATIONSHIP_TYPE.IMAGE)

    # 计算尺寸；裁剪后按源区域比例设置图形框，避免 SVG 被拉伸。
    cx, cy = _get_svg_dimensions(svg_path, width_cm)
    src_rect = ''
    if isinstance(image_data, dict):
        crop_rect = _normalize_feishu_crop(image_data.get('crop'))
        if crop_rect:
            crop = {
                'left': crop_rect[0],
                'top': crop_rect[1],
                'right': 1.0 - crop_rect[2],
                'bottom': 1.0 - crop_rect[3],
            }
        else:
            crop = calculate_center_crop(
                cx,
                cy,
                image_data.get('width'),
                image_data.get('height'),
            )
        if crop:
            src_rect = '<a:srcRect ' + ' '.join(
                f'{attribute}="{int(round(crop[name] * 100000))}"'
                for name, attribute in (
                    ('left', 'l'),
                    ('top', 't'),
                    ('right', 'r'),
                    ('bottom', 'b'),
                )
                if crop[name] > 0
            ) + '/>'
            try:
                if crop_rect:
                    crop_width = crop_rect[2] - crop_rect[0]
                    crop_height = crop_rect[3] - crop_rect[1]
                    cy = max(1, int(round(cx * crop_height / crop_width)))
                else:
                    frame_width = float(image_data.get('width'))
                    frame_height = float(image_data.get('height'))
                    cy = max(1, int(round(cx * frame_height / frame_width)))
            except (TypeError, ValueError, ZeroDivisionError):
                src_rect = ''

    # 构建 Drawing XML 并替换 run 内容
    doc_pr_id = abs(hash(svg_path)) % 100000
    xml_str = _SVG_XML_TEMPLATE.format(
        cx=cx, cy=cy, doc_pr_id=doc_pr_id,
        png_rid=png_rid, svg_rid=svg_rid,
        svg_ext_uri=_SVG_EXT_URI,
        src_rect=src_rect,
    )
    run_elem = parse_xml(xml_str)
    rPr = run._r.find(qn('w:rPr'))
    for child in list(run._r):
        run._r.remove(child)
    if rPr is not None:
        run._r.append(rPr)
    drawing = run_elem.find(qn('w:drawing'))
    if drawing is not None:
        run._r.append(drawing)
logger = ConfigLoader.get_logger('feishu2docx')
BLOCK_TYPES = {1: 'page', 2: 'text', 3: 'heading1', 4: 'heading2', 5: 'heading3', 6: 'heading4', 7: 'heading5', 8: 'heading6', 9: 'heading7', 10: 'heading8', 11: 'heading9', 12: 'bullet', 13: 'ordered', 14: 'code', 15: 'quote', 17: 'todo', 18: 'bitable', 19: 'highlight', 20: 'callout', 21: 'iframe', 22: 'divider', 23: 'file', 24: 'column', 25: 'column', 26: 'iframe', 27: 'image', 28: 'callout', 29: 'mindnote', 30: 'sheet', 31: 'table', 32: 'table_cell', 33: 'view', 34: 'quote_container', 35: 'task', 36: 'okr', 37: 'okr_objective', 38: 'okr_key_result', 39: 'okr_progress', 40: 'callout', 41: 'file', 42: 'callout', 43: 'whiteboard'}
TEXT_COLOR_MAP = {1: 'E85E5E', 2: 'F08C4A', 3: 'F5D450', 4: '7ED321', 5: '4A90E2', 6: '9013FE', 7: '9B9B9B'}
FEISHU_BG_TO_WORD_HIGHLIGHT = {
    1: 'DARK_BLUE',         # 浅红 -> 深蓝
    2: 'TEAL',              # 浅橙 -> 青色
    3: 'YELLOW',            # 浅黄 -> 黄色
    4: 'BRIGHT_GREEN',      # 浅绿 -> 鲜绿
    5: 'TURQUOISE',         # 浅蓝 -> 青绿
    6: 'PINK',              # 浅紫 -> 粉红
    7: 'GRAY_50',           # 中灰 -> 50% 灰
    8: 'RED',               # 红 -> 红
    9: 'DARK_RED',          # 橙 -> 深红
    10: 'DARK_YELLOW',      # 黄 -> 深黄
    11: 'GREEN',            # 绿 -> 绿
    12: 'BLUE',             # 蓝 -> 蓝
    13: 'VIOLET',           # 紫 -> 紫罗兰
    14: 'BLACK',            # 灰 -> 黑色
    15: 'GRAY_25',          # 浅灰 -> 25% 灰
}

# Table-cell fills are returned by some versions of the Feishu document API
# as the same palette indexes used by text highlights.  Keep a concrete RGB
# palette for Word cell shading, while also accepting a literal hex color
# when the API (or a proxy) returns one.  The values intentionally use the
# light palette for the first seven entries so a light Feishu fill remains
# readable in the exported document.
FEISHU_TABLE_BG_COLORS = {
    1: 'FDE2E2',  # 浅红
    2: 'FCE8D5',  # 浅橙
    3: 'FFF4CC',  # 浅黄
    4: 'E4F7D2',  # 浅绿
    5: 'DDEBFF',  # 浅蓝
    6: 'EFE1FF',  # 浅紫
    7: 'D9D9D9',  # 中灰
    8: 'FF7D7D',  # 红
    9: 'FFBA5C',  # 橙
    10: 'FFE66D',  # 黄
    11: '7ED321',  # 绿
    12: '4A90E2',  # 蓝
    13: '9013FE',  # 紫
    14: '8F959E',  # 灰
    15: 'F2F3F5',  # 浅灰
}


def normalize_table_background_color(value):
    """Normalize a Feishu table-cell background value to ``RRGGBB``.

    The public SDK currently leaves ``table_cell`` untyped, and deployments
    have returned both palette indexes and literal colors.  Supporting the
    common representations here keeps the feature backwards compatible and
    makes malformed values a harmless no-op.
    """
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, dict):
        for key in ('hex', 'hex_color', 'rgb', 'rgb_color', 'rgbColor', 'color', 'value', 'background_color', 'backgroundColor'):
            if key in value:
                color = normalize_table_background_color(value.get(key))
                if color:
                    return color
        # Some responses encode RGB as separate channels.
        if all(key in value for key in ('red', 'green', 'blue')):
            try:
                channels = [int(value[key]) for key in ('red', 'green', 'blue')]
                if all(0 <= channel <= 255 for channel in channels):
                    return ''.join(f'{channel:02X}' for channel in channels)
            except (TypeError, ValueError):
                pass
        return None
    if isinstance(value, (int, float)):
        try:
            index = int(value)
        except (TypeError, ValueError):
            return None
        return FEISHU_TABLE_BG_COLORS.get(index)
    raw = str(value).strip()
    if not raw:
        return None
    if raw.startswith('#'):
        raw = raw[1:]
    if len(raw) == 8:
        # Accommodate either AARRGGBB or RRGGBBAA values.  Word shading does
        # not support alpha, so retain the RGB portion only.  Fully opaque or
        # fully transparent alpha prefixes are the unambiguous AARRGGBB form;
        # other 8-digit values are treated as RRGGBBAA.
        if raw[:2].lower() == '0x' or raw[:2].lower() in {'00', 'ff'}:
            raw = raw[2:]
        else:
            raw = raw[:6]
    if len(raw) == 6 and all(ch in '0123456789abcdefABCDEF' for ch in raw):
        return raw.upper()
    # Accept CSS-style rgb()/rgba() values returned by some integrations.
    import re
    match = re.fullmatch(r'rgba?\(\s*(\d{1,3})\s*[, ]\s*(\d{1,3})\s*[, ]\s*(\d{1,3})(?:\s*[,/]\s*[^)]*)?\s*\)', raw, re.IGNORECASE)
    if match:
        channels = [int(match.group(i)) for i in range(1, 4)]
        if all(0 <= channel <= 255 for channel in channels):
            return ''.join(f'{channel:02X}' for channel in channels)
    if raw.isdigit():
        return FEISHU_TABLE_BG_COLORS.get(int(raw))
    return FEISHU_TABLE_BG_COLORS.get(raw.casefold())


def extract_table_cell_background(block):
    """Read a cell fill from the loosely-typed Feishu table-cell payload."""
    if not isinstance(block, dict):
        return None
    candidates = [block]
    payload = block.get('table_cell')
    if isinstance(payload, dict):
        candidates.insert(0, payload)
    elif payload is not None:
        color = normalize_table_background_color(payload)
        if color:
            return color
    for candidate in candidates:
        for key in (
            'background_color', 'backgroundColor', 'background_color_index',
            'backgroundColorIndex', 'bg_color', 'bgColor', 'fill_color',
            'fillColor', 'background', 'color',
        ):
            if key in candidate:
                color = normalize_table_background_color(candidate.get(key))
                if color:
                    return color
        # A nested property/style object is used by a few API gateways.
        for key in ('property', 'style'):
            nested = candidate.get(key)
            if isinstance(nested, dict):
                color = extract_table_cell_background(nested)
                if color:
                    return color
    return None


def extract_table_backgrounds_from_content(content, blocks):
    """Extract native table-cell fills from the Docs AI rendered XML.

    ``docx/v1/.../blocks`` currently returns ``table_cell: {}`` even when a
    cell has a fill.  The rendered Docs AI XML keeps that visual information
    on each ``td`` as ``background-color="rgb(...)"``.  Paragraph IDs inside
    each ``td`` let us join the XML back to the block API's table-cell IDs,
    without trying to reproduce rowspan/colspan layout ourselves.

    The return value is ``{table_block_id: {cell_block_id: 'RRGGBB'}}``.
    Malformed or incomplete content is treated as an empty mapping so this
    optional fidelity feature cannot break an otherwise valid export.
    """
    if not isinstance(content, str) or not content.strip() or not isinstance(blocks, (list, tuple)):
        return {}

    block_map = {
        block.get('block_id'): block
        for block in blocks
        if isinstance(block, dict) and block.get('block_id')
    }
    parent_by_child = {}
    for block in block_map.values():
        parent_id = block.get('block_id')
        for child_id in block.get('children') or []:
            if child_id and parent_id:
                parent_by_child[child_id] = parent_id

    def table_cell_ancestor(block_id, table_id):
        """Find the table-cell ancestor of a paragraph within one table."""
        seen = set()
        current = block_id
        while current and current not in seen:
            seen.add(current)
            block = block_map.get(current) or {}
            if block.get('block_type') == 32:
                if block.get('parent_id') == table_id:
                    return current
                # A nested table may contain a cell block from another table;
                # don't accidentally associate it with the outer table.
                return None
            current = block.get('parent_id') or parent_by_child.get(current)
        return None

    try:
        import xml.etree.ElementTree as ET
        # Fetch returns a sequence of top-level XML blocks rather than one
        # document element, so wrap it before parsing.
        root = ET.fromstring('<root>' + content + '</root>')
    except Exception as exc:
        logger.warning(f'解析文档渲染 XML 失败，跳过表格背景色: {exc}')
        return {}

    result = {}
    for table in root.iter('table'):
        table_id = table.get('id')
        if not table_id:
            continue
        table_result = result.setdefault(table_id, {})
        for td in table.iter('td'):
            raw_color = td.get('background-color') or td.get('backgroundColor')
            color = normalize_table_background_color(raw_color)
            if not color:
                continue
            # The first paragraph ID is sufficient to identify the owning
            # table-cell block.  Empty cells still carry an empty paragraph ID
            # in the current API; cells without one are safely skipped.
            cell_block_id = None
            for descendant in td.iter():
                paragraph_id = descendant.get('id') if descendant.tag == 'p' else None
                if paragraph_id:
                    cell_block_id = table_cell_ancestor(paragraph_id, table_id)
                    if cell_block_id:
                        break
            if cell_block_id:
                table_result[cell_block_id] = color
        if not table_result:
            result.pop(table_id, None)
    return result

def add_hyperlink(paragraph, url, text, color='0000FF', underline=True):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

class NumberingInjector:

    def __init__(self, doc):
        self.doc = doc
        self.bullet_num_id = 9001
        self.ordered_num_id = 9002
        self.next_num_id = self._get_initial_max_num_id()
        self._inject_definitions()

    def _get_initial_max_num_id(self):
        try:
            numbering_part = self.doc.part.numbering_part
            if numbering_part:
                existing_num_ids = [int(n.get(qn('w:numId'))) for n in numbering_part.element.findall(qn('w:num'))]
                return max(existing_num_ids, default=0) + 1
        except:
            pass
        return 1

    def _inject_definitions(self):
        try:
            numbering_part = self.doc.part.numbering_part
        except:
            try:
                numbering_part = self.doc.part.numbering_part
            except:
                return
        if numbering_part is None:
            return
        abstract_ids = [int(an.get(qn('w:abstractNumId'))) for an in numbering_part.element.findall(qn('w:abstractNum'))]
        if self.bullet_num_id not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9001">\n                <w:nsid w:val="FFFFFF01"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="●"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="420" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="○"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="840" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9003 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9003">\n                <w:nsid w:val="FFFFFF03"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="■"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="420" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="□"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="840" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9004 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9004">\n                <w:nsid w:val="FFFFFF04"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="◆"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="420" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="◇"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="840" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9005 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9005">\n                <w:nsid w:val="FFFFFF05"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="➢"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="420" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="➤"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="840" w:hanging="420"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if self.ordered_num_id not in abstract_ids:
            import random
            nsid = f'{random.randint(0, 16777215):06X}'
            
            levels = []
            for i in range(9):
                if i % 3 == 0:
                    numFmt = 'decimal'
                elif i % 3 == 1:
                    numFmt = 'lowerLetter'
                else:
                    numFmt = 'lowerRoman'
                lvl_xml = f'''                <w:lvl w:ilvl="{i}">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="{numFmt}"/>
                    <w:lvlText w:val="%{i+1}."/>
                    <w:lvlJc w:val="left"/>
                    <w:pPr>
                        <w:ind w:left="{420 * (i+1)}" w:hanging="420"/>
                    </w:pPr>
                </w:lvl>'''
                levels.append(lvl_xml)

            abstract_xml = f'''
            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9002">
                <w:nsid w:val="{nsid}"/>
                <w:multiLevelType w:val="hybridMultilevel"/>
''' + '\n'.join(levels) + '''
            </w:abstractNum>
            '''
            numbering_part.element.append(parse_xml(abstract_xml))

    def create_num(self, abstract_num_id, restart=True):
        try:
            numbering_part = self.doc.part.numbering_part
            new_num_id = self.next_num_id
            self.next_num_id += 1
            
            overrides = ""
            if restart:
                for i in range(9):
                    overrides += f'''
                <w:lvlOverride w:ilvl="{i}">
                    <w:startOverride w:val="1"/>
                </w:lvlOverride>'''
                    
            num_xml = f'''
            <w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="{new_num_id}">
                <w:abstractNumId w:val="{abstract_num_id}"/>{overrides}
            </w:num>
            '''
            numbering_part.element.append(parse_xml(num_xml))
            return new_num_id
        except Exception as e:
            logger.error(f'创建编号失败: {e}')
            return None

class FeishuDocxConverter:

    def __init__(self, blocks, client, img_dir, template_path=None, progress_cb=None, check_stop_func=None, unordered_list_style='default', ignore_mention=False, add_title=False, image_style=None, table_config=None, table_backgrounds=None, image_crops=None):
        self.blocks = blocks
        self.client = client
        self.img_dir = img_dir
        self.template_path = template_path
        self.progress_cb = progress_cb
        self.check_stop_func = check_stop_func
        self.unordered_list_style = unordered_list_style
        self.ignore_mention = ignore_mention
        self.add_title = add_title
        self.image_style = image_style if isinstance(image_style, dict) else None
        self.table_config = table_config if isinstance(table_config, dict) else None
        self.preserve_table_background = bool(
            self.table_config and self.table_config.get('preserveTableBackground')
        )
        self.table_backgrounds = table_backgrounds if isinstance(table_backgrounds, dict) else {}
        self.image_crops = image_crops if isinstance(image_crops, dict) else {}
        self.block_map = {b['block_id']: b for b in blocks}
        self.tree = self._build_tree()
        self.doc = None
        self.injector = None
        self.user_cache = {}
        self.processed_count = 0
        self.total_blocks = len(blocks)
        self.fallback_download_count = 0
        self._pre_download_images()

    def _pre_download_images(self):
        media_tasks = []
        for block in self.blocks:
            if self.check_stop_func and self.check_stop_func():
                raise InterruptedError('任务已停止')
            btype = block.get('block_type')
            if btype == 27:
                image_data = block.get('image') or {}
                token = image_data.get('token')
                if token:
                    path = os.path.join(self.img_dir, f'{token}.png')
                    if not os.path.exists(path):
                        media_tasks.append((token, path, 'image'))
            elif btype == 43:
                wb_data = block.get('whiteboard') or {}
                wb_id = block.get('board', {}).get('token') or wb_data.get('token') or wb_data.get('whiteboard_id')
                if wb_id:
                    path = os.path.join(self.img_dir, f'wb_{wb_id}.png')
                    media_tasks.append((wb_id, path, 'whiteboard'))
        if not media_tasks:
            return
        total = len(media_tasks)
        logger.info(f'开始并行下载 {total} 个媒体资源...')
        if self.progress_cb:
            self._update_progress(message=f'正在并行下载 {total} 个图片及画板资源...', log_type='dynamic')
        if self.check_stop_func and self.check_stop_func():
            raise InterruptedError('任务已停止')
        max_workers = ConfigLoader.get_int('download.threads', 10)
        failed_count = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_token = {executor.submit(self._download_task, task): task for task in media_tasks}
            completed = 0
            for future in concurrent.futures.as_completed(future_to_token):
                if self.check_stop_func and self.check_stop_func():
                    executor.shutdown(wait=False, cancel_futures=True)
                    raise InterruptedError('任务已停止')
                token, path, type_ = future_to_token[future]
                completed += 1
                try:
                    success = future.result()
                    if not success:
                        logger.warning(f'下载 {type_} 失败: {token}')
                        failed_count += 1
                except PermissionError:
                    executor.shutdown(wait=False, cancel_futures=True)
                    raise
                except Exception as e:
                    logger.error(f'下载任务异常 {token}: {e}')
                    failed_count += 1
                if self.progress_cb and total > 0:
                    progress = 10 + int(completed / total * 30)
                    self._update_progress(percentage=progress, message=f'正在并行下载 {completed} / {total} 个图片及画板资源', log_type='dynamic')
        if failed_count > 0:
            success_count = total - failed_count
            self._update_progress(percentage=40, message=f'已完成 {success_count} / {total} 个图片及画板资源的并发下载（余下 {failed_count} 个在组织文件块时单线程下载）', log_type='success')
        else:
            self._update_progress(percentage=40, message=f'并行下载 {total} 个图片及画板资源完成', log_type='success')

    def _download_task(self, task_info):
        token, path, type_ = task_info
        import random
        import time
        time.sleep(random.uniform(0.1, 0.4))
        try:
            if type_ == 'image':
                if self.client.download_media(token, path):
                    return True
            elif type_ == 'whiteboard':
                if self.client.download_whiteboard(token, path):
                    try:
                        smart_crop(path, padding=20)
                    except:
                        pass
                    return True
            return False
        except PermissionError:
            raise
        except Exception as e:
            logger.debug(f'预下载 {type_} 异常 {token}: {e}')
            return False

    def _update_progress(self, percentage=None, message=None, log_type='info'):
        if self.progress_cb:
            if percentage is not None:
                try:
                    self.progress_cb(percentage, message, t=log_type)
                except TypeError:
                    self.progress_cb(percentage, message)
            elif self.total_blocks > 0:
                p = 40 + int(self.processed_count / self.total_blocks * 50)
                msg = f'正在组织文件（第 {self.processed_count} / {self.total_blocks} 个块）'
                if message:
                    msg += f': {message}'
                if not log_type or log_type == 'info':
                    log_type = 'dynamic'
                try:
                    self.progress_cb(p, msg, t=log_type)
                except TypeError:
                    self.progress_cb(p, msg)

    def _build_tree(self):
        page_block = next((b for b in self.blocks if b.get('block_type') == 1), None)
        if not page_block:
            known_ids = set(self.block_map.keys())
            roots = [b for b in self.blocks if not b.get('parent_id') or b.get('parent_id') not in known_ids]
            return roots
        return [page_block]

    def process(self, output_path):
        if self.template_path and os.path.exists(self.template_path):
            try:
                self.doc = Document(self.template_path)
                logger.info(f'加载模板于 {self.template_path}')
                if self.doc.element.body is not None:
                    for element in list(self.doc.element.body):
                        if element.tag.endswith('sectPr'):
                            continue
                        self.doc.element.body.remove(element)
            except Exception as e:
                logger.error(f'加载模板失败 {self.template_path}: {e}')
                self.doc = Document()
        else:
            self.doc = Document()
            logger.info('创建了新的空文档')
        if len(self.doc.sections) == 0:
            self.doc.add_section()
        # Register the two paragraph presets before rendering so a DOCX
        # produced directly by the converter (without the cleaner) still has
        # the same table body/header styles as the normal service pipeline.
        self.table_body_style, self.table_header_style = TableStyleManager.ensure_table_paragraph_styles(self.doc)
        self.image_paragraph_style = TableStyleManager.ensure_image_paragraph_style(self.doc)
        try:
            self.injector = NumberingInjector(self.doc)
        except Exception as e:
            logger.error(f'注入列表样式失败: {e}')
        self._update_progress(message='开始渲染文档...')
        for root in self.tree:
            if self.check_stop_func and self.check_stop_func():
                raise InterruptedError('任务已停止')
            self._render_block(root, self.doc, level=0)
        self._apply_table_paragraph_styles()
        self._apply_image_paragraph_styles()
        self.processed_count = self.total_blocks
        self._update_progress(percentage=90, message=f'已组织 {self.total_blocks} / {self.total_blocks} 个文件块（已完成 {self.fallback_download_count} 张补充图片下载）', log_type='success')
        self.doc.save(output_path)
        return output_path

    def _apply_table_paragraph_styles(self):
        """Apply the table body/header presets to generated table paragraphs."""
        ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        def visit(table):
            try:
                tbl_pr = table._element.tblPr
                caption = tbl_pr.find(f'{{{ns}}}tblCaption') if tbl_pr is not None else None
                if caption is not None and caption.get(f'{{{ns}}}val') == 'code_block':
                    return
            except Exception:
                pass
            TableStyleManager.apply_table_paragraph_styles(
                table, self.table_body_style, self.table_header_style
            )
            for row in table.rows:
                for cell in row.cells:
                    for nested_table in cell.tables:
                        visit(nested_table)

        for table in self.doc.tables:
            visit(table)

    def _apply_image_paragraph_styles(self):
        """Keep image paragraphs independent from the document body style."""
        for paragraph in self.doc.paragraphs:
            TableStyleManager.apply_image_paragraph_style(
                paragraph, self.image_paragraph_style
            )

        def visit(table):
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        TableStyleManager.apply_image_paragraph_style(
                            paragraph, self.image_paragraph_style
                        )
                    for nested_table in cell.tables:
                        visit(nested_table)

        for table in self.doc.tables:
            visit(table)

    def _render_children(self, block, container, child_level):
        children_ids = block.get('children') or []
        current_ordered_num_id = None
        for cid in children_ids:
            child = self.block_map.get(cid)
            if not child:
                continue
            btype = child.get('block_type')
            if btype == 13:
                ordered_props = child.get('ordered') or {}
                style_props = ordered_props.get('style') or {}
                sequence = style_props.get('sequence', 'auto')
                if current_ordered_num_id is None:
                    current_ordered_num_id = self.injector.create_num(self.injector.ordered_num_id)
                elif sequence == '1':
                    current_ordered_num_id = self.injector.create_num(self.injector.ordered_num_id)
                child['_num_id'] = current_ordered_num_id
            self._render_block(child, container, level=child_level)

    def _render_block(self, block, container, level=0):
        if self.check_stop_func and self.check_stop_func():
            raise InterruptedError('任务已停止')
        self.processed_count += 1
        if self.processed_count % 10 == 0:
            self._update_progress()
        btype = block.get('block_type')
        handler_name = f"_handle_{BLOCK_TYPES.get(btype, 'unknown')}"
        handler = getattr(self, handler_name, self._handle_unknown)
        block['_level'] = level
        try:
            handler(block, container)
        except PermissionError:
            raise
        except Exception as e:
            logger.error(f"处理块错误 {block.get('block_id')} ({btype}): {str(e)}")

    def _handle_page(self, block, container):
        # 渲染页面标题 (Page Block 的 elements 通常包含标题)
        page_data = block.get('page') or {}
        elements = page_data.get('elements') or []
        if elements and self.add_title:
            try:
                # 尝试使用 Title 样式
                p = container.add_paragraph(style='Title')
            except:
                # 如果模板没有 Title 样式，回退到默认并加粗加大
                p = container.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            self._add_runs(p, elements)
            
            # 如果是回退样式，手动设置格式
            if p.style.name != 'Title':
                for run in p.runs:
                    run.font.bold = True
                    run.font.size = Pt(22)  # 约等于二号字
        
        # 渲染子块
        self._render_children(block, container, child_level=0)

    def _handle_text(self, block, container):
        text_data = block.get('text') or {}
        p = self._add_paragraph(container, text_data)
        
        # 处理缩进 (indentation)
        style = text_data.get('style') or {}
        # indentation_level 通常是 "OneLevelIndent", "TwoLevelIndent" 等
        indent_level = style.get('indentation_level')
        
        if indent_level:
            # 简单映射：OneLevelIndent -> 1级缩进
            level_map = {
                'OneLevelIndent': 1,
                'TwoLevelIndent': 2,
                'ThreeLevelIndent': 3,
                'FourLevelIndent': 4,
                'FiveLevelIndent': 5,
                'SixLevelIndent': 6,
                'SevenLevelIndent': 7,
                'EightLevelIndent': 8,
                'NineLevelIndent': 9
            }
            level = level_map.get(indent_level, 0)
            if level > 0:
                # 每一级缩进 2 个字符宽度 (21磅左右，或者直接用 cm)
                # Word 默认缩进是 0.75cm 或 21pt
                # 这里使用 Pt(21) * level
                try:
                    p.paragraph_format.left_indent = Pt(21 * level)
                except Exception as e:
                    pass
        
        # 渲染子块（Text 块下也可能有子节点，如缩进内容）
        self._render_children(block, container, child_level=block.get('_level', 0) + 1)

    def _handle_heading(self, block, level, container):
        text_data = (block.get(f'heading{level}') or {}).get('elements') or []
        style = f'Heading {level}'
        try:
            p = container.add_paragraph(style=style)
        except:
            p = container.add_paragraph()
        self._add_runs(p, text_data)
        
        # 渲染子块（Heading 下的折叠内容等）
        self._render_children(block, container, child_level=block.get('_level', 0) + 1)

    def _handle_heading1(self, b, c):
        self._handle_heading(b, 1, c)

    def _handle_heading2(self, b, c):
        self._handle_heading(b, 2, c)

    def _handle_heading3(self, b, c):
        self._handle_heading(b, 3, c)

    def _handle_heading4(self, b, c):
        self._handle_heading(b, 4, c)

    def _handle_heading5(self, b, c):
        self._handle_heading(b, 5, c)

    def _handle_heading6(self, b, c):
        self._handle_heading(b, 6, c)

    def _handle_heading7(self, b, c):
        self._handle_heading(b, 7, c)

    def _handle_heading8(self, b, c):
        self._handle_heading(b, 8, c)

    def _handle_heading9(self, b, c):
        self._handle_heading(b, 9, c)

    def _handle_bullet(self, block, container):
        text_data = (block.get('bullet') or {}).get('elements') or []
        level = block.get('_level', 0)
        if self.unordered_list_style == 'none':
            p = container.add_paragraph()
            self._add_runs(p, text_data)
            self._render_children(block, container, child_level=level + 1)
            return
        parent_id = block.get('parent_id') or 'root'
        list_key = f'{parent_id}_bullet'
        if not hasattr(self, '_list_context'):
            self._list_context = {}
        if list_key not in self._list_context:
            abstract_id = self.injector.bullet_num_id
            if self.unordered_list_style == 'square':
                abstract_id = 9003
            elif self.unordered_list_style == 'diamond':
                abstract_id = 9004
            elif self.unordered_list_style == 'arrow':
                abstract_id = 9005
            new_id = self.injector.create_num(abstract_id)
            self._list_context[list_key] = new_id
        num_id = self._list_context[list_key]
        p = container.add_paragraph()
        try:
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = int(num_id)
            numPr.get_or_add_ilvl().val = int(level)
        except Exception as e:
            logger.warning(f'设置无序列表属性失败: {e}')
        self._add_runs(p, text_data)
        self._render_children(block, container, child_level=level + 1)

    def _handle_ordered(self, block, container):
        text_data = (block.get('ordered') or {}).get('elements') or []
        level = block.get('_level', 0)
        num_id = block.get('_num_id')
        if not num_id:
            num_id = self.injector.create_num(self.injector.ordered_num_id)
            logger.warning(f"块 {block.get('block_id')}: 缺少 _num_id，已创建回退值 {num_id}")
        p = container.add_paragraph()
        try:
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = int(num_id)
            numPr.get_or_add_ilvl().val = int(level)
            ind = pPr.get_or_add_ind()
            left_indent = 420 * (level + 1)
            ind.set(qn('w:left'), str(left_indent))
            ind.set(qn('w:hanging'), '420')
        except Exception as e:
            logger.warning(f'设置有序列表属性失败: {e}')
        self._add_runs(p, text_data)
        self._render_children(block, container, child_level=level + 1)

    def _handle_quote(self, block, container):
        text_data = (block.get('quote') or {}).get('elements') or []
        try:
            p = container.add_paragraph(style='Quote')
        except:
            p = container.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
        self._add_runs(p, text_data)
        self._render_children(block, container, child_level=block.get('_level', 0))

    def _handle_code(self, block, container):
        code_data = (block.get('code') or {}).get('elements') or []
        try:
            table = container.add_table(rows=1, cols=1)
            try:
                table.style = 'Table Grid'
            except:
                pass
            
            # Initial code block table property clearing to avoid inherited auto-fit interference
            table.autofit = False
            tblPr = table._element.tblPr
            if tblPr is not None:
                # Add code block marker
                caption = OxmlElement('w:tblCaption')
                caption.set(qn('w:val'), 'code_block')
                tblPr.append(caption)
                
                # Force fixed layout
                tbl_layout = tblPr.find(qn('w:tblLayout'))
                if tbl_layout is None:
                    tbl_layout = OxmlElement('w:tblLayout')
                    tbl_layout.set(qn('w:type'), 'fixed')
                    tblPr.append(tbl_layout)
                else:
                    tbl_layout.set(qn('w:type'), 'fixed')
                    
                # Force default width to auto to avoid 7.33cm issue before cleaner runs
                tbl_w = tblPr.find(qn('w:tblW'))
                if tbl_w is None:
                    tbl_w = OxmlElement('w:tblW')
                    tbl_w.set(qn('w:w'), '0')
                    tbl_w.set(qn('w:type'), 'auto')
                    tblPr.append(tbl_w)
                else:
                    tbl_w.set(qn('w:w'), '0')
                    tbl_w.set(qn('w:type'), 'auto')
                    
            cell = table.cell(0, 0)
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is not None:
                tcW.set(qn('w:w'), '0')
                tcW.set(qn('w:type'), 'auto')
            else:
                tcW = OxmlElement('w:tcW')
                tcW.set(qn('w:w'), '0')
                tcW.set(qn('w:type'), 'auto')
                tcPr.append(tcW)

            self._set_cell_shading(cell, 'F5F5F5')
            if cell.paragraphs:
                p = cell.paragraphs[0]
            else:
                p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            
            # 使用统一的 _add_runs 处理内容，支持加粗、颜色、@、链接等
            self._add_runs(p, code_data)
            
            # 强制设置代码块字体（在 _add_runs 之后设置）
            for run in p.runs:
                run.font.name = 'Courier New'
                run.font.size = Pt(9)
                try:
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), 'Courier New')
                    rFonts.set(qn('w:ascii'), 'Courier New')
                    rFonts.set(qn('w:hAnsi'), 'Courier New')
                except:
                    pass
        except Exception as e:
            logger.error(f"渲染代码块失败 {block.get('block_id')}: {e}")
            logger.error(traceback.format_exc())

    def _handle_image(self, block, container):
        image_data = dict(block.get('image') or {})
        token = image_data.get('token')
        if not token:
            return
        crop = (
            self.image_crops.get(block.get('block_id'))
            or self.image_crops.get(token)
        )
        if crop:
            image_data['crop'] = crop
        file_path = os.path.join(self.img_dir, f'{token}.png')
        if not os.path.exists(file_path):
            self._update_progress(message=f'正在下载图片 ({token[:8]}...)')
            try:
                self.client.download_media(token, file_path)
                self.fallback_download_count += 1
            except PermissionError as e:
                logger.error(str(e))
                self._update_progress(message=f'下载失败(无权限): {token}', log_type='error')
                raise
            except Exception as e:
                logger.error(f'下载图片异常 {token}: {e}')
        if os.path.exists(file_path):
            is_svg = _is_svg_file(file_path)
            try:
                configured_width = self.image_style.get('maxWidth') if self.image_style else None
                try:
                    configured_width = float(configured_width) if configured_width not in (None, '') else None
                except (TypeError, ValueError):
                    configured_width = None
                if configured_width is not None and configured_width > 0:
                    max_w_cm = configured_width
                    width_for_insert = max_w_cm
                else:
                    max_w_cm = ConfigLoader.get_float('image.max_width', 16)
                    width_for_insert = max_w_cm - 1 if max_w_cm > 1 else max_w_cm
                p = container.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                if is_svg:
                    logger.info(f'检测到 SVG 图片，直接嵌入 DOCX: {token}')
                    _add_svg_to_docx(
                        run,
                        file_path,
                        width_cm=width_for_insert,
                        image_data=image_data,
                    )
                    # SVG 内容可能变动，删除缓存确保下次重新下载
                    try:
                        os.remove(file_path)
                        logger.debug(f'SVG 缓存已清除，下次将重新下载: {token}')
                    except Exception as e:
                        logger.warning(f'清除 SVG 缓存失败: {token}: {e}')
                else:
                    inline_shape = run.add_picture(file_path, width=Cm(width_for_insert))
                    if _apply_feishu_image_crop(inline_shape, file_path, image_data):
                        logger.info(
                            f'已按飞书可见区域裁剪图片: {token} '
                            f'({image_data.get("width")}x{image_data.get("height")})'
                        )
            except Exception as e:
                logger.error(f'添加图片失败 {token}: {e}')
        # 渲染图片描述（caption），按正文段落处理
        caption_data = image_data.get('caption') or {}
        caption_text = (caption_data.get('content') or '').strip()
        if caption_text:
            try:
                cp = container.add_paragraph()
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cp.add_run(caption_text)
                logger.info(f'已添加图片描述: {caption_text}')
            except Exception as e:
                logger.error(f'添加图片描述失败: {e}')

    def _handle_whiteboard(self, block, container):
        wb_data = block.get('whiteboard') or {}
        wb_id = block.get('board', {}).get('token') or wb_data.get('token') or wb_data.get('whiteboard_id')
        if not wb_id:
            return
        file_path = os.path.join(self.img_dir, f'wb_{wb_id}.png')
        self._update_progress(message=f'正在下载画板 ({wb_id[:8]}...)')
        if self.client.download_whiteboard(wb_id, file_path):
            self.fallback_download_count += 1
            try:
                smart_crop(file_path, padding=20)
            except Exception as e:
                logger.warning(f'裁剪画板失败 {wb_id}: {e}')
        if os.path.exists(file_path):
            try:
                p = container.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                configured_width = self.image_style.get('maxWidth') if self.image_style else None
                try:
                    configured_width = float(configured_width) if configured_width not in (None, '') else None
                except (TypeError, ValueError):
                    configured_width = None
                run.add_picture(file_path, width=Cm(configured_width if configured_width and configured_width > 0 else 15))
            except Exception as e:
                logger.error(f'添加画板失败 {wb_id}: {e}')

    def _get_col_letter(self, col_idx):
        result = ''
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _handle_sheet(self, block, container):
        sheet_data = block.get('sheet') or {}
        token = sheet_data.get('token')
        if not token or '_' not in token:
            logger.warning(f'无效的表格 Token: {token}')
            return
        try:
            ss_token, sheet_id = token.split('_')
            self._update_progress(message=f'正在获取电子表格元数据 ({token[:8]}...)')
            meta = self.client.get_sheet_meta(ss_token, sheet_id)
            if not meta:
                logger.error(f'获取表格元数据失败 {token}')
                return
            grid = meta.get('grid_properties') or {}
            row_count = grid.get('row_count', 0)
            col_count = grid.get('column_count', 0)
            merges = meta.get('merges', [])
            if row_count == 0 or col_count == 0:
                logger.warning(f'空表格 {token}')
                return
            end_col_char = self._get_col_letter(col_count)
            range_str = f'{sheet_id}!A1:{end_col_char}{row_count}'
            self._update_progress(message=f'正在获取电子表格内容 ({row_count}行 x {col_count}列)')
            values_data = self.client.get_sheet_values(ss_token, range_str)
            values = values_data.get('values', []) if values_data else []
            self._update_progress(message=f'正在渲染电子表格 ({row_count}行 x {col_count}列)')
            table = container.add_table(rows=row_count, cols=col_count)
            try:
                table.style = 'Table Grid'
            except Exception:
                pass
            try:
                tblPr = table._element.tblPr
                if tblPr is None:
                    tblPr = OxmlElement('w:tblPr')
                    table._element.insert(0, tblPr)
                caption = OxmlElement('w:tblCaption')
                caption.set(qn('w:val'), 'sheet')
                tblPr.append(caption)
            except Exception as e:
                logger.warning(f'标记表格为 sheet 失败: {e}')
            for r in range(row_count):
                row_vals = values[r] if r < len(values) else []
                for c in range(col_count):
                    val = row_vals[c] if c < len(row_vals) else ''
                    if val:
                        if isinstance(val, list):
                            text = ''.join([v.get('text', '') for v in val if isinstance(v, dict)])
                            table.cell(r, c).text = text
                        else:
                            table.cell(r, c).text = str(val)
            for merge in merges:
                start_row = merge.get('start_row_index', 0)
                start_col = merge.get('start_column_index', 0)
                end_row = merge.get('end_row_index', start_row)
                end_col = merge.get('end_column_index', start_col)
                row_span = end_row - start_row + 1
                col_span = end_col - start_col + 1
                if row_span > 1 or col_span > 1:
                    try:
                        top_left = table.cell(start_row, start_col)
                        bottom_right = table.cell(start_row + row_span - 1, start_col + col_span - 1)
                        top_left.merge(bottom_right)
                    except Exception as e:
                        logger.warning(f'合并表格单元格失败 {token}: {e}')
            logger.info(f'已渲染表格 {token} ({row_count}x{col_count})')
        except PermissionError:
            raise
        except Exception as e:
            logger.error(f'处理表格失败 {token}: {e}')
            logger.error(traceback.format_exc())

    def _handle_table(self, block, container):
        table_data = block.get('table') or {}
        props = table_data.get('property') or {}
        cols = int(props.get('column_size') or table_data.get('column_size') or 0)
        cells = table_data.get('cells') or []
        if not cells or cols == 0:
            logger.warning(f"表格块 {block.get('block_id')} 没有单元格或列数为 0")
            return
        import math
        rows = math.ceil(len(cells) / cols)
        self._update_progress(message=f'正在渲染表格 ({rows}行 x {cols}列)')
        try:
            doc_table = container.add_table(rows=rows, cols=cols)
            doc_table.autofit = False
            try:
                doc_table.style = 'Table Grid'
            except Exception:
                pass
            merge_info = props.get('merge_info') or []
            for idx, cell_id in enumerate(cells):
                if idx >= len(merge_info):
                    break
                r = idx // cols
                c = idx % cols
                if r >= rows:
                    break
                info = merge_info[idx]
                row_span = info.get('row_span', 1)
                col_span = info.get('col_span', 1)
                if row_span > 1 or col_span > 1:
                    end_row = r + row_span - 1
                    end_col = c + col_span - 1
                    if end_row < rows and end_col < cols:
                        try:
                            cell_tl = doc_table.cell(r, c)
                            cell_br = doc_table.cell(end_row, end_col)
                            if cell_tl._tc != cell_br._tc:
                                cell_tl.merge(cell_br)
                        except Exception as e:
                            logger.warning(f'预合并错误在 {r},{c}: {e}')
            header_row = props.get('header_row', False)
            # Newer Feishu responses may expose cell fills on the cell block;
            # older gateways have also returned a parallel list on the table
            # payload.  Both are supported, with an explicit cell value taking
            # precedence over the fallback list/table-level value.
            table_backgrounds = self._extract_table_backgrounds(
                table_data, props, len(cells), block
            )
            covered_cells = set()
            for idx, cell_id in enumerate(cells):
                r = idx // cols
                c = idx % cols
                if r >= rows:
                    break
                if (r, c) in covered_cells:
                    continue
                if idx < len(merge_info):
                    info = merge_info[idx]
                    row_span = info.get('row_span', 1)
                    col_span = info.get('col_span', 1)
                    for rr in range(r, r + row_span):
                        for cc in range(c, c + col_span):
                            if rr == r and cc == c:
                                continue
                            if rr < rows and cc < cols:
                                covered_cells.add((rr, cc))
                cell_block = self.block_map.get(cell_id)
                if not cell_block:
                    continue
                try:
                    if r >= len(doc_table.rows) or c >= len(doc_table.rows[r].cells):
                        continue
                    doc_cell = doc_table.cell(r, c)
                    doc_cell._element.clear_content()
                    if self.preserve_table_background:
                        table_colors = self.table_backgrounds.get(block.get('block_id'), {})
                        color = table_colors.get(cell_id) if isinstance(table_colors, dict) else None
                        # Keep support for any future/legacy block payload that
                        # does expose a fill directly on ``table_cell``.
                        if not color:
                            color = extract_table_cell_background(cell_block)
                        if not color and idx < len(table_backgrounds):
                            color = table_backgrounds[idx]
                        if color:
                            self._set_cell_shading(doc_cell, color)
                    self._render_children(cell_block, doc_cell, child_level=0)
                    if not doc_cell.paragraphs:
                        doc_cell.add_paragraph()
                    if header_row and r == 0:
                        for p in doc_cell.paragraphs:
                            for run in p.runs:
                                run.font.bold = True
                except PermissionError:
                    raise
                except Exception as e:
                    logger.warning(f'渲染单元格错误 {cell_id} 位于 {r},{c}: {e}')
        except PermissionError:
            raise
        except Exception as e:
            logger.error(f"创建表格失败 {block.get('block_id')}: {e}")
            logger.error(traceback.format_exc())

    @staticmethod
    def _extract_table_backgrounds(table_data, props, cell_count, block=None):
        """Return normalized per-cell fills from alternate API response shapes."""
        if not isinstance(table_data, dict):
            return []
        candidates = []
        payloads = [table_data, props if isinstance(props, dict) else {}]
        if isinstance(block, dict):
            payloads.append(block)
        for payload in payloads:
            for key in (
                'background_colors', 'backgroundColors',
                'cell_background_colors', 'cellBackgroundColors',
                'cell_background', 'cellBackground',
            ):
                value = payload.get(key)
                if isinstance(value, (list, tuple)):
                    candidates = list(value)
                    break
                if isinstance(value, dict):
                    candidates = [value.get(str(i), value.get(i)) for i in range(cell_count)]
                    break
            if candidates:
                break
        if not candidates:
            # A scalar table-level fill is a useful fallback for API versions
            # that expose one color for the whole table.
            for payload in payloads:
                for key in (
                    'background_color', 'backgroundColor', 'background_color_index',
                    'backgroundColorIndex', 'bg_color', 'bgColor', 'fill_color',
                    'fillColor', 'background',
                ):
                    if key in payload:
                        color = normalize_table_background_color(payload.get(key))
                        return [color] * cell_count if color else []
        result = [normalize_table_background_color(value) for value in candidates[:cell_count]]
        if len(result) < cell_count:
            result.extend([None] * (cell_count - len(result)))
        return result

    def _handle_table_cell(self, block, container):
        self._render_children(block, container, child_level=0)

    def _handle_unknown(self, block, container):
        # 对于未知或通用的容器块，递归渲染其子节点
        # 注意：这里需要确保所有可能包含子节点的块类型都调用了 _render_children
        # 飞书文档的嵌套结构可能很深，通过递归调用 _render_block -> _handle_xxx -> _render_children -> _render_block 实现无限层级支持
        self._render_children(block, container, child_level=block.get('_level', 0) + 1)

    def _add_paragraph(self, container, text_data):
        elements = text_data.get('elements') or []
        p = container.add_paragraph()
        self._add_runs(p, elements)
        return p

    def _add_runs(self, paragraph, elements):
        for el in elements:
            if 'text_run' in el:
                tr = el['text_run']
                content = tr.get('content', '')
                style = tr.get('text_element_style', {})
                
                # Check for link in text_element_style (Feishu new format)
                final_link = None
                style_link = style.get('link')
                if style_link and style_link.get('url'):
                    final_link = style_link.get('url')
                
                if final_link:
                    # Handle hyperlink
                    try:
                        if '%3A' in final_link or '%3a' in final_link:
                            from urllib.parse import unquote
                            final_link = unquote(final_link)
                    except:
                        pass
                    # For hyperlinks, we use a helper which adds a run
                    # Note: style application for hyperlinks is limited in python-docx helper
                    add_hyperlink(paragraph, final_link, content)
                else:
                    # Normal text run
                    run = paragraph.add_run(content)
                    if style.get('bold'):
                        run.font.bold = True
                    if style.get('italic'):
                        run.font.italic = True
                    if style.get('underline'):
                        run.font.underline = True
                    if style.get('strikethrough'):
                        run.font.strike = True
                    
                    color_idx = style.get('text_color')
                    if color_idx and color_idx in TEXT_COLOR_MAP:
                        run.font.color.rgb = RGBColor.from_string(TEXT_COLOR_MAP[color_idx])
                    
                    # 背景颜色 (Highlight)
                    bg_color_idx = style.get('background_color')
                    if bg_color_idx and bg_color_idx in FEISHU_BG_TO_WORD_HIGHLIGHT:
                        # 使用 Word 原生 highlight
                        highlight_color = FEISHU_BG_TO_WORD_HIGHLIGHT[bg_color_idx]
                        run.font.highlight_color = getattr(WD_COLOR_INDEX, highlight_color, WD_COLOR_INDEX.YELLOW)
                        
                        # 特殊处理：深色背景下文字自动改为白色，确保可读性
                        # 1=深蓝, 2=青色(Teal), 7=50%灰, 8=红, 9=深红, 10=深黄, 11=绿, 12=蓝, 13=紫罗兰, 14=黑
                        dark_bg_ids = {1, 2, 7, 8, 9, 10, 11, 12, 13, 14}
                        if bg_color_idx in dark_bg_ids:
                            run.font.color.rgb = RGBColor(255, 255, 255)
            elif 'mention_user' in el:
                if self.ignore_mention:
                    continue
                user = el['mention_user']
                user_id = user.get('user_id')
                name = user.get('user_name') or 'User'
                if user_id:
                    if user_id in self.user_cache:
                        name = self.user_cache[user_id]
                    else:
                        try:
                            user_info = self.client.get_user_info(user_id)
                            if user_info and user_info.get('name'):
                                name = user_info.get('name')
                                self.user_cache[user_id] = name
                        except Exception as e:
                            logger.warning(f'获取用户信息失败 {user_id}: {e}')
                run = paragraph.add_run(f'@{name}')
                run.font.color.rgb = RGBColor(0, 0, 255)
            elif 'mention_doc' in el:
                doc = el['mention_doc']
                title = doc.get('title') or 'Doc'
                url = doc.get('url')
                if url:
                    add_hyperlink(paragraph, url, title, color='0000FF', underline=True)
                else:
                    run = paragraph.add_run(title)
                    run.font.color.rgb = RGBColor(0, 0, 255)
                    run.font.underline = True

    def _set_cell_shading(self, cell, color_hex):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        existing = tcPr.find(qn('w:shd'))
        if existing is not None:
            tcPr.remove(existing)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color_hex)
        tcPr.append(shd)
