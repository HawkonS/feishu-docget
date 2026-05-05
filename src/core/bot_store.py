import json
import os
import threading
from datetime import datetime

from src.core.config_loader import config
from src.core.feishu_client import FeishuClient


BOT_STORE_FILE = 'custom_bots.json'
_store_lock = threading.Lock()


def normalize_bot_config(bot_config):
    if not bot_config:
        return None

    app_id = str(bot_config.get('app_id') or bot_config.get('appId') or '').strip()
    app_secret = str(bot_config.get('app_secret') or bot_config.get('appSecret') or '').strip()

    if not app_id and not app_secret:
        return None
    if not app_id or not app_secret:
        raise ValueError('请同时填写机器人 App ID 和 App Secret')

    return {'app_id': app_id, 'app_secret': app_secret}


def get_bot_store_path(base_dir='.'):
    base_path = os.path.abspath(base_dir or '.')
    log_dir = str(config.get('log.dir', 'logs') or 'logs').strip() or 'logs'
    if os.path.isabs(log_dir):
        store_dir = log_dir
    else:
        store_dir = os.path.join(base_path, log_dir)
    return os.path.join(store_dir, BOT_STORE_FILE)


def validate_bot_credentials(app_id, app_secret):
    client = FeishuClient(app_id, app_secret)
    return bool(client.get_token())


def _read_store(path):
    if not os.path.exists(path):
        return {'bots': []}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return {'bots': []}

    if isinstance(data, list):
        bots = data
    else:
        bots = data.get('bots') if isinstance(data, dict) else []
    if not isinstance(bots, list):
        bots = []
    return {'bots': [bot for bot in bots if isinstance(bot, dict)]}


def save_bot_credentials(base_dir, bot_config):
    normalized = normalize_bot_config(bot_config)
    if not normalized:
        return None

    path = get_bot_store_path(base_dir)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    now = datetime.now().isoformat(timespec='seconds')

    with _store_lock:
        data = _read_store(path)
        bots = data['bots']
        matched = None
        for item in bots:
            item_app_id = item.get('app_id') or item.get('appId')
            item_app_secret = item.get('app_secret') or item.get('appSecret')
            if item_app_id == normalized['app_id'] and item_app_secret == normalized['app_secret']:
                matched = item
                break

        if matched:
            matched['app_id'] = normalized['app_id']
            matched['app_secret'] = normalized['app_secret']
            matched.pop('appId', None)
            matched.pop('appSecret', None)
            matched['last_used_at'] = now
        else:
            bots.append({
                'app_id': normalized['app_id'],
                'app_secret': normalized['app_secret'],
                'created_at': now,
                'last_used_at': now,
            })

        unique_bots = []
        seen = set()
        for item in bots:
            item_app_id = item.get('app_id') or item.get('appId')
            item_app_secret = item.get('app_secret') or item.get('appSecret')
            if not item_app_id or not item_app_secret:
                continue
            key = (item_app_id, item_app_secret)
            if key in seen:
                continue
            seen.add(key)
            item['app_id'] = item_app_id
            item['app_secret'] = item_app_secret
            item.pop('appId', None)
            item.pop('appSecret', None)
            unique_bots.append(item)
        bots = unique_bots

        tmp_path = path + '.tmp'
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump({'bots': bots}, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass

    return normalized


def validate_and_store_custom_bot(base_dir, bot_config):
    normalized = normalize_bot_config(bot_config)
    if not normalized:
        return None

    if not validate_bot_credentials(normalized['app_id'], normalized['app_secret']):
        raise ValueError('自定义机器人身份验证未通过，请检查 App ID 和 App Secret')

    return save_bot_credentials(base_dir, normalized)
