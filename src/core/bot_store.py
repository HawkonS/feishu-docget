from src.core.feishu_client import FeishuClient


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


def validate_bot_credentials(app_id, app_secret):
    client = FeishuClient(app_id, app_secret)
    return bool(client.get_token())
