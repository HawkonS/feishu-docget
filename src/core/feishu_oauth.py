from urllib.parse import urlencode

import requests

from src.core.config_loader import ConfigLoader, config

AUTHORIZE_URL = 'https://accounts.feishu.cn/oauth/v1/authorize'
TOKEN_URL = 'https://open.feishu.cn/open-apis/authen/v2/oauth/token'
USER_INFO_URL = 'https://open.feishu.cn/open-apis/authen/v2/user_info'
TENANT_TOKEN_URL = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'

logger = ConfigLoader.get_logger('feishu_docget')

_session = requests.Session()
_adapter = requests.adapters.HTTPAdapter(max_retries=3, pool_connections=10, pool_maxsize=10)
_session.mount('http://', _adapter)
_session.mount('https://', _adapter)


def build_authorize_url(app_id, redirect_uri, state):
    """构造飞书 OAuth v2 授权页跳转 URL"""
    query = urlencode({
        'app_id': app_id,
        'redirect_uri': redirect_uri,
        'state': state,
    })
    return f'{AUTHORIZE_URL}?{query}'


def exchange_code(code, redirect_uri):
    """用授权码换取用户 access_token，失败抛出 RuntimeError（含飞书错误描述）"""
    data = {
        'grant_type': 'authorization_code',
        'client_id': config.get('feishu.app_id', ''),
        'client_secret': config.get('feishu.app_secret', ''),
        'code': code,
        'redirect_uri': redirect_uri,
    }
    try:
        res = _session.post(TOKEN_URL, data=data, timeout=10).json()
    except Exception as e:
        logger.error(f'飞书 OAuth 换取 token 请求异常: {e}')
        raise RuntimeError(f'飞书 OAuth 换取 token 请求异常: {e}')
    if 'access_token' in res:
        return res
    desc = res.get('error_description') or res.get('error') or str(res)
    logger.error(f'飞书 OAuth 换取 token 失败: {desc}')
    raise RuntimeError(f'飞书 OAuth 换取 token 失败: {desc}')


def refresh_user_token(refresh_token):
    """刷新用户 access_token，成功返回 token dict，失败返回 None"""
    data = {
        'grant_type': 'refresh_token',
        'client_id': config.get('feishu.app_id', ''),
        'client_secret': config.get('feishu.app_secret', ''),
        'refresh_token': refresh_token,
    }
    try:
        res = _session.post(TOKEN_URL, data=data, timeout=10).json()
    except Exception as e:
        logger.error(f'刷新用户 token 请求异常: {e}')
        return None
    if 'access_token' in res:
        return res
    desc = res.get('error_description') or res.get('error') or str(res)
    logger.error(f'刷新用户 token 失败: {desc}')
    return None


def get_oauth_user_info(user_access_token):
    """用用户 access_token 获取当前登录用户信息（authen/v2 可返回 department_ids），失败返回 {}"""
    headers = {'Authorization': 'Bearer ' + user_access_token}
    try:
        res = _session.get(USER_INFO_URL, headers=headers, timeout=10).json()
    except Exception as e:
        logger.error(f'获取 OAuth 用户信息请求异常: {e}')
        return {}
    if res.get('code') != 0:
        logger.error(f'获取 OAuth 用户信息失败: {res.get("msg", "")}')
        return {}
    # v2 返回体字段在顶层（无 data 包裹），防御性兼容
    data = res.get('data') if isinstance(res.get('data'), dict) else res
    return {
        'name': data.get('name', ''),
        'open_id': data.get('open_id', ''),
        'union_id': data.get('union_id', ''),
        'user_id': data.get('user_id', ''),
        'avatar_url': data.get('avatar_url', ''),
        'department_ids': data.get('department_ids') or [],
    }


def resolve_department_names(department_ids):
    """用 tenant 身份把部门 ID 列表解析为部门名称，"、" 拼接；任何异常均返回 ''"""
    try:
        if not department_ids:
            return ''
        app_id = config.get('feishu.app_id', '')
        app_secret = config.get('feishu.app_secret', '')
        if not app_id or not app_secret:
            return ''
        token_res = _session.post(
            TENANT_TOKEN_URL,
            json={'app_id': app_id, 'app_secret': app_secret},
            timeout=10).json()
        if token_res.get('code') != 0:
            logger.warning(f'解析部门名称: 获取 tenant_access_token 失败: {token_res.get("msg", "")}')
            return ''
        tenant_token = token_res.get('tenant_access_token') or ''
        if not tenant_token:
            return ''
        headers = {'Authorization': 'Bearer ' + tenant_token}
        names = []
        for dept_id in department_ids:
            if not dept_id:
                continue
            url = f'https://open.feishu.cn/open-apis/contact/v3/departments/{dept_id}'
            res = _session.get(url, headers=headers, params={'user_id_type': 'open_id'}, timeout=10).json()
            if res.get('code') != 0:
                logger.warning(f'解析部门名称失败 dept_id={dept_id}: {res.get("msg", "")}')
                continue
            name = (res.get('data') or {}).get('department', {}).get('name') or ''
            if name:
                names.append(name)
        return '、'.join(names)
    except Exception as e:
        logger.warning(f'解析部门名称异常: {e}')
        return ''
