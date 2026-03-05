import os
import json
import time
from datetime import datetime

def get_stats_file(base_dir, config):
    workspace = config.get('workspace.dir', '.')
    log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, 'download_stats.jsonl')

def update_download_stat(base_dir, config, task_id, status, doc_url='', file_path='', title='', ip_address=''):
    stats_file = get_stats_file(base_dir, config)
    entry = {'id': task_id, 'status': status, 'ts': int(time.time()), 'time': datetime.now().isoformat(), 'url': doc_url, 'path': file_path, 'title': title, 'ip': ip_address}
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
                        items.append(json.loads(line))
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