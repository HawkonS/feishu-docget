from urllib.parse import urlencode

import requests

from src.core.config_loader import ConfigLoader, config

AUTHORIZE_URL = 'https://passport.feishu.cn/suite/passport/oauth/authorize'
TOKEN_URL = 'https://open.feishu.cn/open-apis/authen/v2/oauth/token'
USER_INFO_URL = 'https://open.feishu.cn/open-apis/authen/v1/user_info'
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
    """用用户 access_token 获取当前登录用户信息（authen/v1，data 包裹），失败返回 {}

    department_ids 因 v1 接口不返回，改经 tenant 身份通讯录接口补充；任何失败降级为空列表，不阻断登录
    """
    headers = {'Authorization': 'Bearer ' + user_access_token}
    try:
        resp = _session.get(USER_INFO_URL, headers=headers, timeout=10)
    except Exception as e:
        logger.error(f'获取 OAuth 用户信息请求异常: {e}')
        return {}
    content_type = resp.headers.get('Content-Type', '')
    if resp.status_code != 200 or 'application/json' not in content_type:
        logger.error(f'获取 OAuth 用户信息失败: status={resp.status_code}, content-type={content_type}, body={resp.text[:200]!r}')
        return {}
    try:
        res = resp.json()
    except Exception as e:
        logger.error(f'获取 OAuth 用户信息 JSON 解析失败: {e}, body={resp.text[:200]!r}')
        return {}
    if res.get('code') != 0:
        logger.error(f'获取 OAuth 用户信息失败: code={res.get("code")}, msg={res.get("msg", "")}')
        return {}
    # 防御性兼容 data 包裹/顶层字段两种结构
    data = res.get('data') if isinstance(res.get('data'), dict) else res
    open_id = data.get('open_id', '')
    info = {
        'name': data.get('name', ''),
        'open_id': open_id,
        'union_id': data.get('union_id', ''),
        'user_id': data.get('user_id', ''),
        'avatar_url': data.get('avatar_url', ''),
        'department_ids': data.get('department_ids') or fetch_department_ids(open_id),
    }
    return info


def fetch_department_ids(open_id):
    """用 tenant 身份经通讯录接口获取用户部门 ID 列表；任何失败/无权限均降级为空列表，不阻断登录"""
    try:
        if not open_id:
            return []
        app_id = config.get('feishu.app_id', '')
        app_secret = config.get('feishu.app_secret', '')
        if not app_id or not app_secret:
            return []
        token_res = _session.post(
            TENANT_TOKEN_URL,
            json={'app_id': app_id, 'app_secret': app_secret},
            timeout=10).json()
        if token_res.get('code') != 0:
            logger.warning(f'获取部门 ID: 获取 tenant_access_token 失败: {token_res.get("msg", "")}')
            return []
        tenant_token = token_res.get('tenant_access_token') or ''
        if not tenant_token:
            return []
        headers = {'Authorization': 'Bearer ' + tenant_token}
        url = f'https://open.feishu.cn/open-apis/contact/v3/users/{open_id}'
        res = _session.get(url, headers=headers, params={'user_id_type': 'open_id'}, timeout=10).json()
        if res.get('code') != 0:
            logger.warning(f'获取部门 ID 失败 open_id={open_id}: {res.get("msg", "")}')
            return []
        return (res.get('data') or {}).get('user', {}).get('department_ids') or []
    except Exception as e:
        logger.warning(f'获取部门 ID 异常: {e}')
        return []


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
