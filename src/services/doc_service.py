import os
import shutil

from docx import Document

from src.converters.docx.cleaner import apply_custom_styles, apply_document_info, clean_document
from src.converters.docx.converter import FeishuDocxConverter
from src.core.bot_store import normalize_bot_config, save_bot_credentials, validate_bot_credentials
from src.core.config_loader import ConfigLoader, config
from src.core.feishu_client import FeishuClient
from src.core.utils import sanitize_name


def _raise_if_stopped(check_stop_func):
    if check_stop_func and check_stop_func():
        raise InterruptedError('任务已停止')


def _cleanup_doc_folder(doc_folder, logger, reason):
    if not doc_folder or not os.path.exists(doc_folder):
        return
    try:
        shutil.rmtree(doc_folder)
        logger.info(f'{reason}，清理临时文件夹: {doc_folder}')
    except Exception as e:
        logger.error(f'清理临时文件夹失败: {e}')


def _process_document_with_client(client, doc_url, template_path=None, table_style=None, output_root='output', progress_cb=None, add_cover=False, check_stop_func=None, unordered_list_style='default', body_style=None, image_style=None, ignore_mention=False, ignore_template_heading_num=False, table_config=None, margin_config=None, code_block_config=None, document_info=None, add_title=False):
    logger = ConfigLoader.get_logger('service')
    _raise_if_stopped(check_stop_func)

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
                        if os.name == 'nt':
                            try:
                                import subprocess

                                abs_target = os.path.abspath(master_img_dir)
                                abs_link = os.path.abspath(img_dir)
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
                else:
                    logger.warning(f'图片目录已存在且不是软链接，将独立下载: {img_dir}')
            except OSError as e:
                logger.warning(f'创建软链接失败，回退到独立目录模式: {e}')
                os.makedirs(img_dir, exist_ok=True)

        _raise_if_stopped(check_stop_func)
        if progress_cb:
            progress_cb(30, '正在获取文档块', 'dynamic')
        blocks = client.get_blocks(doc_id)
        if not blocks:
            raise RuntimeError('未找到内容')

        total_blocks = len(blocks)
        if progress_cb:
            progress_cb(50, f'已成功获取文档信息，共 {total_blocks} 个块', 'success')

        docx_path = os.path.join(doc_folder, f'{base_title}.docx')
        _raise_if_stopped(check_stop_func)

        converter = FeishuDocxConverter(blocks, client, master_img_dir, template_path=template_path, progress_cb=progress_cb, check_stop_func=check_stop_func, unordered_list_style=unordered_list_style, ignore_mention=ignore_mention, add_title=add_title)
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
            clean_document(docx_path, progress_cb=progress_cb, template_path=template_path, add_cover=add_cover, body_style=body_style, image_style=image_style, table_config=table_config, margin_config=margin_config, code_block_config=code_block_config, document_info=document_info, ignore_template_heading_num=ignore_template_heading_num)
            if document_info is not None:
                apply_document_info(docx_path, document_info)
        except Exception as e:
            logger.error(f'格式清理错误: {e}')

        logger.info('处理成功: ' + docx_path)
        _raise_if_stopped(check_stop_func)
        return {'docx_path': docx_path, 'folder': doc_folder, 'title': base_title}
    except InterruptedError:
        _cleanup_doc_folder(doc_folder, logger, '任务已停止')
        raise
    except PermissionError:
        _cleanup_doc_folder(doc_folder, logger, '机器人权限不足')
        raise


def process_document(doc_url, template_path=None, table_style=None, base_dir='.', output_root='output', progress_cb=None, add_cover=False, check_stop_func=None, unordered_list_style='default', body_style=None, image_style=None, ignore_mention=False, ignore_template_heading_num=False, table_config=None, margin_config=None, code_block_config=None, document_info=None, add_title=False, bot_config=None):
    logger = ConfigLoader.get_logger('service')
    _raise_if_stopped(check_stop_func)

    custom_bot = normalize_bot_config(bot_config)
    had_custom_bot = custom_bot is not None
    app_id = config.get('feishu.app_id') or os.getenv('FEISHU_APP_ID', '')
    app_secret = config.get('feishu.app_secret') or os.getenv('FEISHU_APP_SECRET', '')

    system_client = None
    if app_id and app_secret:
        system_client = FeishuClient(app_id, app_secret)

    if custom_bot:
        if progress_cb:
            progress_cb(6, '正在校验自定义机器人', 'dynamic')
        if validate_bot_credentials(custom_bot['app_id'], custom_bot['app_secret']):
            try:
                save_bot_credentials(base_dir, custom_bot)
            except Exception as e:
                logger.warning(f'自定义机器人可用，但保存记录失败: {e}')
        else:
            logger.warning('自定义机器人身份验证未通过，自动回退至系统默认机器人')
            if progress_cb:
                progress_cb(8, '自定义机器人身份验证未通过，自动回退至系统默认机器人', 'info')
            custom_bot = None

    if not custom_bot and not system_client:
        if had_custom_bot:
            raise RuntimeError('自定义机器人身份验证未通过，且系统默认机器人未配置')
        raise RuntimeError('缺少飞书 App ID 或 Secret')

    process_kwargs = {
        'doc_url': doc_url,
        'template_path': template_path,
        'table_style': table_style,
        'output_root': output_root,
        'progress_cb': progress_cb,
        'add_cover': add_cover,
        'check_stop_func': check_stop_func,
        'unordered_list_style': unordered_list_style,
        'body_style': body_style,
        'image_style': image_style,
        'ignore_mention': ignore_mention,
        'ignore_template_heading_num': ignore_template_heading_num,
        'table_config': table_config,
        'margin_config': margin_config,
        'code_block_config': code_block_config,
        'document_info': document_info,
        'add_title': add_title,
    }

    if custom_bot:
        custom_client = FeishuClient(custom_bot['app_id'], custom_bot['app_secret'])
        try:
            if progress_cb:
                progress_cb(8, '正在使用自定义机器人下载', 'dynamic')
            return _process_document_with_client(custom_client, **process_kwargs)
        except PermissionError as custom_error:
            logger.warning(f'自定义机器人权限不足，准备切换系统默认机器人: {custom_error}')
            if not system_client:
                raise PermissionError('自定义机器人权限不足，且系统默认机器人未配置') from custom_error
            if progress_cb:
                progress_cb(10, '自定义机器人权限不足，正在切换系统默认机器人', 'dynamic')
            try:
                return _process_document_with_client(system_client, **process_kwargs)
            except PermissionError as system_error:
                raise PermissionError(f'自定义机器人和系统默认机器人均无权限。自定义机器人错误: {custom_error}；系统默认机器人错误: {system_error}') from system_error

    return _process_document_with_client(system_client, **process_kwargs)
