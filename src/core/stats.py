import os
import time
from datetime import datetime
from urllib.parse import urlsplit

from src.core import sqlite_store

def _mask_ip(ip):
    """IP 地址脱敏：保留前两段，后两段用 * 替换"""
    if not ip:
        return ''
    parts = ip.split('.')
    if len(parts) == 4:
        return f'{parts[0]}.{parts[1]}.*.*'
    return ip  # IPv6 等不做处理

def _mask_url(url):
    """URL 脱敏：仅保留带主机的 HTTPS URL，去掉查询参数并截断。"""
    if not url:
        return ''
    try:
        parsed = urlsplit(str(url).strip())
        hostname = (parsed.hostname or '').lower()
        allowed_host = (
            hostname == 'feishu.cn' or hostname.endswith('.feishu.cn') or
            hostname == 'larksuite.com' or hostname.endswith('.larksuite.com')
        )
        if parsed.scheme.lower() != 'https' or not parsed.netloc or not allowed_host:
            return ''
        url = parsed._replace(query='', fragment='').geturl()
    except ValueError:
        return ''
    return url[:60] + '...' if len(url) > 60 else url

def get_stats_file(base_dir, config):
    workspace = config.get('workspace.dir', '.')
    log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    path = os.path.join(log_dir, 'download_stats.jsonl')
    if os.path.exists(path):
        try:
            os.chmod(os.path.realpath(path), 0o600)
        except OSError:
            pass
    return path

def update_download_stat(base_dir, config, task_id, status, doc_url='', file_path='', title='', ip_address='', user_name=''):
    entry = {'id': task_id, 'status': status, 'ts': int(time.time()), 'time': datetime.now().isoformat(), 'url': _mask_url(doc_url), 'path': file_path, 'title': title, 'ip': _mask_ip(ip_address), 'user': user_name}
    sqlite_store.upsert_download_stat(base_dir, config, entry)

def get_download_stats(base_dir, config, limit=None, page=None, page_size=None,
                       title='', url='', ip='', date=''):
    """读取下载统计，并可在服务端完成筛选和分页。

    ``limit`` 保留给旧调用方使用；管理后台使用 ``page``/``page_size``，
    避免把整个统计文件发送到浏览器后再做筛选和分页。
    """
    return sqlite_store.list_download_stats(
        base_dir, config, limit=limit, page=page, page_size=page_size,
        title=title, url=url, ip=ip, date=date,
    )
