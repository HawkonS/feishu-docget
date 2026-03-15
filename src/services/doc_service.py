import os
from src.core.feishu_client import FeishuClient
from src.converters.docx.converter import FeishuDocxConverter
from src.converters.docx.cleaner import clean_document, apply_custom_styles
from src.core.config_loader import config, ConfigLoader
from src.core.utils import sanitize_name
from docx import Document

def process_document(doc_url, template_path=None, table_style=None, base_dir='.', output_root='output', progress_cb=None, add_cover=False, check_stop_func=None, unordered_list_style='default', body_style=None, image_style=None, ignore_mention=False):
    logger = ConfigLoader.get_logger('service')
    if check_stop_func and check_stop_func():
        raise InterruptedError('任务已停止')
    app_id = config.get('feishu.app_id') or os.getenv('FEISHU_APP_ID', '')
    app_secret = config.get('feishu.app_secret') or os.getenv('FEISHU_APP_SECRET', '')
    if not app_id or not app_secret:
        raise RuntimeError('缺少飞书 App ID 或 Secret')
    client = FeishuClient(app_id, app_secret)
    doc_id = client.extract_doc_id(doc_url)
    if not doc_id:
        raise RuntimeError('无效的文档链接')
    if progress_cb:
        progress_cb(10, '正在获取文档信息', 'dynamic')
    meta = client.get_document_meta(doc_id)
    base_title = sanitize_name(meta.get('title') or meta.get('name') or doc_id)
    folder_base = doc_id
    folder_name = folder_base
    counter = 1
    while True:
        doc_folder = os.path.join(output_root, folder_name)
        if not os.path.exists(doc_folder):
            break
        folder_name = f'{folder_base}-{counter}'
        counter += 1
    try:
        os.makedirs(doc_folder, exist_ok=True)
        master_img_dir = os.path.join(output_root, doc_id, 'img')
        img_dir = os.path.join(doc_folder, 'img')
        if os.path.abspath(img_dir) == os.path.abspath(master_img_dir):
            os.makedirs(img_dir, exist_ok=True)
        else:
            try:
                if not os.path.exists(master_img_dir):
                    os.makedirs(master_img_dir, exist_ok=True)
                if not os.path.exists(img_dir):
                    try:
                        rel_path = os.path.relpath(master_img_dir, os.path.dirname(img_dir))
                        os.symlink(rel_path, img_dir)
                        logger.info(f'创建图片软链接: {img_dir} -> {master_img_dir}')
                    except OSError as e:
                        # Windows: Try Junction if symlink fails (usually due to privileges)
                        if os.name == 'nt':
                            try:
                                import subprocess
                                abs_target = os.path.abspath(master_img_dir)
                                abs_link = os.path.abspath(img_dir)
                                # mklink /J Link Target
                                cmd = f'mklink /J "{abs_link}" "{abs_target}"'
                                subprocess.check_call(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                                logger.info(f'创建图片目录联接(Junction): {img_dir} -> {master_img_dir}')
                            except Exception as je:
                                logger.warning(f'创建Junction失败: {je}')
                                raise e
                        else:
                            raise e
                    
                    if progress_cb:
                        progress_cb(25, '检测到历史下载记录，已关联至历史图片资源', 'success')
                elif os.path.islink(img_dir):
                    if progress_cb:
                        progress_cb(25, '检测到历史下载记录，已关联至历史图片资源', 'success')
                    pass
                else:
                    logger.warning(f'图片目录已存在且不是软链接，将独立下载: {img_dir}')
            except OSError as e:
                logger.warning(f'创建软链接失败，回退到独立目录模式: {e}')
                os.makedirs(img_dir, exist_ok=True)
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        if progress_cb:
            progress_cb(30, '正在获取文档块', 'dynamic')
        blocks = client.get_blocks(doc_id)
        if not blocks:
            raise RuntimeError('未找到内容')
        total_blocks = len(blocks)
        if progress_cb:
            progress_cb(50, f'已成功获取文档信息，共 {total_blocks} 个块', 'success')
        docx_path = os.path.join(doc_folder, f'{base_title}.docx')
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        converter = FeishuDocxConverter(blocks, client, img_dir, template_path, progress_cb, check_stop_func, unordered_list_style, ignore_mention)
        converter.process(docx_path)
        if progress_cb:
            progress_cb(80, '正在应用样式', 'dynamic')
        if table_style:
            try:
                doc = Document(docx_path)
                apply_custom_styles(doc, int(table_style))
                doc.save(docx_path)
            except Exception as e:
                logger.warning(f'应用表格样式失败: {e}')
        if progress_cb:
            progress_cb(80, '已完成样式应用', 'success')
        try:
            clean_document(docx_path, progress_cb=progress_cb, template_path=template_path, add_cover=add_cover, body_style=body_style, image_style=image_style)
        except Exception as e:
            logger.error(f'格式清理错误: {e}')
        logger.info('处理成功: ' + docx_path)
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        return {'docx_path': docx_path, 'folder': doc_folder, 'title': base_title}
    except InterruptedError:
        if os.path.exists(doc_folder):
            try:
                import shutil
                shutil.rmtree(doc_folder)
                logger.info(f'任务已停止，清理临时文件夹: {doc_folder}')
            except Exception as e:
                logger.error(f'清理临时文件夹失败: {e}')
        raise