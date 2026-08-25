import os
import json
import time
from datetime import datetime
from urllib.parse import urlsplit

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
    stats_file = get_stats_file(base_dir, config)
    entry = {'id': task_id, 'status': status, 'ts': int(time.time()), 'time': datetime.now().isoformat(), 'url': _mask_url(doc_url), 'path': file_path, 'title': title, 'ip': _mask_ip(ip_address), 'user': user_name}
    with open(stats_file, 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')

def get_download_stats(base_dir, config, limit=None):
    stats_file = get_stats_file(base_dir, config)
    if not os.path.exists(stats_file):
        return {'total': 0, 'items': []}
    items = []
    try:
        with open(stats_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        item = json.loads(line)
                        # 旧版本可能已经写入非 HTTPS/非飞书协议的恶意值；读取时
                        # 再净化一遍，避免修复部署后仍显示存量 XSS 载荷。
                        if isinstance(item, dict):
                            item['url'] = _mask_url(item.get('url', ''))
                            items.append(item)
                    except json.JSONDecodeError:
                        pass
    except Exception:
        return {'total': 0, 'items': []}
    stats_map = {}
    for item in items:
        tid = item.get('id')
        if tid:
            if tid not in stats_map or item.get('ts', 0) >= stats_map[tid].get('ts', 0):
                if tid in stats_map:
                    old = stats_map[tid]
                    for k, v in old.items():
                        if k not in item or not item[k]:
                            item[k] = v
                stats_map[tid] = item
        else:
            stats_map[f"legacy_{item.get('ts')}"] = item
    final_items = list(stats_map.values())
    final_items.sort(key=lambda x: x.get('ts', 0), reverse=True)
    if limit:
        final_items = final_items[:limit]
    return {'total': len(final_items), 'items': final_items}
