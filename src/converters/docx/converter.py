import os
import concurrent.futures
import traceback
from urllib.parse import unquote
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
import docx.opc.constants
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from src.core.config_loader import ConfigLoader
from src.core.image_processor import smart_crop
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
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9001">\n                <w:nsid w:val="FFFFFF01"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="●"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="720" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="○"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="1440" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9003 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9003">\n                <w:nsid w:val="FFFFFF03"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="■"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="720" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="□"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="1440" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9004 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9004">\n                <w:nsid w:val="FFFFFF04"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="◆"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="720" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="◇"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="1440" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if 9005 not in abstract_ids:
            abstract_xml = '\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9005">\n                <w:nsid w:val="FFFFFF05"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="➢"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="720" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="bullet"/>\n                    <w:lvlText w:val="➤"/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="1440" w:hanging="360"/>\n                    </w:pPr>\n                    <w:rPr>\n                        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>\n                    </w:rPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))
        if self.ordered_num_id not in abstract_ids:
            import random
            nsid = f'{random.randint(0, 16777215):06X}'
            abstract_xml = f'\n            <w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="9002">\n                <w:nsid w:val="{nsid}"/>\n                <w:multiLevelType w:val="hybridMultilevel"/>\n                <w:lvl w:ilvl="0">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="decimal"/>\n                    <w:lvlText w:val="%1."/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="720" w:hanging="360"/>\n                    </w:pPr>\n                </w:lvl>\n                <w:lvl w:ilvl="1">\n                    <w:start w:val="1"/>\n                    <w:numFmt w:val="decimal"/>\n                    <w:lvlText w:val="%1.%2."/>\n                    <w:lvlJc w:val="left"/>\n                    <w:pPr>\n                        <w:ind w:left="1440" w:hanging="360"/>\n                    </w:pPr>\n                </w:lvl>\n            </w:abstractNum>\n            '
            numbering_part.element.append(parse_xml(abstract_xml))

    def create_num(self, abstract_num_id, restart=True):
        try:
            numbering_part = self.doc.part.numbering_part
            new_num_id = self.next_num_id
            self.next_num_id += 1
            num_xml = f'\n            <w:num xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:numId="{new_num_id}">\n                <w:abstractNumId w:val="{abstract_num_id}"/>\n                <w:lvlOverride w:ilvl="0">\n                    <w:startOverride w:val="1"/>\n                </w:lvlOverride>\n            </w:num>\n            '
            numbering_part.element.append(parse_xml(num_xml))
            return new_num_id
        except Exception as e:
            logger.error(f'创建编号失败: {e}')
            return None

class FeishuDocxConverter:

    def __init__(self, blocks, client, img_dir, template_path=None, progress_cb=None, check_stop_func=None, unordered_list_style='default'):
        self.blocks = blocks
        self.client = client
        self.img_dir = img_dir
        self.template_path = template_path
        self.progress_cb = progress_cb
        self.check_stop_func = check_stop_func
        self.unordered_list_style = unordered_list_style
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
                    if not os.path.exists(path):
                        media_tasks.append((wb_id, path, 'whiteboard'))
        if not media_tasks:
            return
        total = len(media_tasks)
        logger.info(f'开始并行下载 {total} 个媒体资源...')
        if self.progress_cb:
            self._update_progress(message=f'正在并行下载 {total} 个图片及画板资源...', log_type='dynamic')
        if self.check_stop_func and self.check_stop_func():
            raise InterruptedError('任务已停止')
        max_workers = 10
        try:
            val = ConfigLoader.load_config().get('download.threads', 10)
            if isinstance(val, str) and (not val.strip()):
                val = 10
            max_workers = int(val)
        except (ValueError, TypeError):
            max_workers = 10
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
        try:
            self.injector = NumberingInjector(self.doc)
        except Exception as e:
            logger.error(f'注入列表样式失败: {e}')
        self._update_progress(message='开始渲染文档...')
        for root in self.tree:
            if self.check_stop_func and self.check_stop_func():
                raise InterruptedError('任务已停止')
            self._render_block(root, self.doc, level=0)
        self.processed_count = self.total_blocks
        self._update_progress(percentage=90, message=f'已组织 {self.total_blocks} / {self.total_blocks} 个文件块（已完成 {self.fallback_download_count} 张补充图片下载）', log_type='success')
        self.doc.save(output_path)
        return output_path

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
        except Exception as e:
            logger.error(f"处理块错误 {block.get('block_id')} ({btype}): {str(e)}")

    def _handle_page(self, block, container):
        # 渲染页面标题 (Page Block 的 elements 通常包含标题)
        page_data = block.get('page') or {}
        elements = page_data.get('elements') or []
        if elements:
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
            left_indent = 720 * (level + 1)
            ind.set(qn('w:left'), str(left_indent))
            ind.set(qn('w:hanging'), '360')
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
            table.style = 'Table Grid'
            cell = table.cell(0, 0)
            self._set_cell_shading(cell, 'F5F5F5')
            if cell.paragraphs:
                p = cell.paragraphs[0]
            else:
                p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            for el in code_data:
                text_run = el.get('text_run') or {}
                content = text_run.get('content') or ''
                run = p.add_run(content)
                run.font.name = 'Courier New'
                run.font.size = Pt(9)
        except Exception as e:
            logger.error(f"渲染代码块失败 {block.get('block_id')}: {e}")

    def _handle_image(self, block, container):
        image_data = block.get('image') or {}
        token = image_data.get('token')
        if not token:
            return
        file_path = os.path.join(self.img_dir, f'{token}.png')
        if not os.path.exists(file_path):
            self._update_progress(message=f'正在下载图片 ({token[:8]}...)')
            try:
                self.client.download_media(token, file_path)
                self.fallback_download_count += 1
            except PermissionError as e:
                logger.error(str(e))
                self._update_progress(message=f'下载失败(无权限): {token}', log_type='error')
            except Exception as e:
                logger.error(f'下载图片异常 {token}: {e}')
        if os.path.exists(file_path):
            try:
                try:
                    val_w = ConfigLoader.load_config().get('image.max_width', 16)
                    if isinstance(val_w, str) and (not val_w.strip()):
                        val_w = 16
                    max_w_cm = float(val_w)
                except (ValueError, TypeError):
                    max_w_cm = 16.0
                p = container.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(file_path, width=Cm(max_w_cm - 1 if max_w_cm > 1 else max_w_cm))
            except Exception as e:
                logger.error(f'添加图片失败 {token}: {e}')

    def _handle_whiteboard(self, block, container):
        wb_data = block.get('whiteboard') or {}
        wb_id = block.get('board', {}).get('token') or wb_data.get('token') or wb_data.get('whiteboard_id')
        if not wb_id:
            return
        file_path = os.path.join(self.img_dir, f'wb_{wb_id}.png')
        if not os.path.exists(file_path):
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
                run.add_picture(file_path, width=Cm(15))
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
                    self._render_children(cell_block, doc_cell, child_level=0)
                    if not doc_cell.paragraphs:
                        doc_cell.add_paragraph()
                    if header_row and r == 0:
                        for p in doc_cell.paragraphs:
                            for run in p.runs:
                                run.font.bold = True
                except Exception as e:
                    logger.warning(f'渲染单元格错误 {cell_id} 位于 {r},{c}: {e}')
        except Exception as e:
            logger.error(f"创建表格失败 {block.get('block_id')}: {e}")
            logger.error(traceback.format_exc())

    def _handle_table_cell(self, block, container):
        self._render_children(block, container, child_level=0)

    def _handle_unknown(self, block, container):
        self._render_children(block, container, child_level=block.get('_level', 0))

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
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color_hex)
        tcPr.append(shd)