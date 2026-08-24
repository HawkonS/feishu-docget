from urllib.parse import quote, urlencode

import requests

from src.core.config_loader import ConfigLoader, config

AUTHORIZE_URL = 'https://accounts.feishu.cn/open-apis/authen/v1/authorize'
TOKEN_URL = 'https://open.feishu.cn/open-apis/authen/v2/oauth/token'
USER_INFO_URL = 'https://open.feishu.cn/open-apis/authen/v1/user_info'

# 用户 token 需携带的权限点位（旧 passport 端点不支持 scope 参数，已切换到 accounts 端点）
OAUTH_SCOPES = [
    'offline_access',                    # 获取 refresh_token 必需
    'docx:document',                     # 创建及编辑新版文档（官方声明包含 docx:document:readonly）
    'docx:document:create',              # 创建新版文档
    'docx:document.block:convert',       # 转换文本为云文档块
    'docs:document:copy',                # 复制云文档
    'docs:document.comment:create',      # 添加、回复云文档中的评论
    'docs:document.comment:read',        # 获取云文档中的评论
    'docs:document.media:download',      # 下载云文档中的图片和附件
    'docs:permission.member:create',     # 添加云文档协作者
    'drive:drive',                       # 查看、评论、编辑和管理云空间中所有文件（官方声明包含 drive:drive:readonly）
    'drive:file:download',               # 下载云空间下的文件
    'space:document:retrieve',           # 获取云空间文件夹下的云文档清单
    'wiki:node:read',                    # 查看知识空间节点信息
    'wiki:node:retrieve',                # 查看知识空间节点列表
    'wiki:space:read',                   # 查看知识空间信息
    'contact:contact.base:readonly',   # 获取通讯录基本信息（调用 contact/v3/users 接口必需）
    'contact:user.base:readonly',      # 获取用户基本信息（返回 name 字段必需，用于@提及人名解析）
    'sheets:spreadsheet:readonly',       # 电子表格只读（保留）
    'board:whiteboard:node:read',        # 画板节点读取（保留）
]

logger = ConfigLoader.get_logger('feishu_docget')

_session = requests.Session()
_adapter = requests.adapters.HTTPAdapter(max_retries=3, pool_connections=10, pool_maxsize=10)
_session.mount('http://', _adapter)
_session.mount('https://', _adapter)


def build_authorize_url(app_id, redirect_uri, state):
    """构造飞书 OAuth v2 授权页跳转 URL（accounts 端点，支持 scope 参数）"""
    # quote_via=quote 使 scope 中空格编码为 %20（符合 RFC 3986），而非默认的 +
    query = urlencode({
        'client_id': app_id,
        'response_type': 'code',
        'redirect_uri': redirect_uri,
        'state': state,
        'scope': ' '.join(OAUTH_SCOPES),
    }, quote_via=quote)
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
    """用用户 access_token 获取当前登录用户信息（authen/v1，data 包裹），失败返回 {}"""
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
    info = {
        'name': data.get('name', ''),
        'open_id': data.get('open_id', ''),
        'union_id': data.get('union_id', ''),
        'user_id': data.get('user_id', ''),
        'avatar_url': data.get('avatar_url', ''),
    }
    return info
