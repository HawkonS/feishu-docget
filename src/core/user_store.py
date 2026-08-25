import os
import threading
from datetime import datetime

from src.core.config_loader import ConfigLoader, config
from src.core import sqlite_store

logger = ConfigLoader.get_logger('feishu_docget')

# 全局可重入锁保护用户读写；每个 open_id 另有独立锁用于串行刷新 token
_lock = threading.RLock()
_user_locks = {}
_token_refresh_margin = 300  # token 距过期不足 5 分钟时触发刷新
_refresh_token_default_ttl = 30 * 24 * 3600  # 飞书未返回 refresh_token 有效期时按 30 天兜底

# 密码登录对应的固定系统账号。它不代表任何飞书用户，因此不会覆盖飞书用户的真实姓名。
SYSTEM_ADMIN_OPEN_ID = '__system_admin__'
SYSTEM_ADMIN_NAME = '管理员'


def get_users_file():
    """用户库文件路径：<workspace>/<log.dir>/users.json（拼法参照 stats.get_stats_file）"""
    workspace = config.get('workspace.dir', '.')
    log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, 'users.json')


def _base_dir():
    return os.path.abspath(config.get('workspace.dir', '.'))


def _now_iso():
    return datetime.now().isoformat()


def is_system_admin(open_id):
    """返回 open_id 是否为固定的系统管理员账号。"""
    return str(open_id or '').strip() == SYSTEM_ADMIN_OPEN_ID


def ensure_system_admin():
    """确保固定系统管理员存在，且始终保持启用和管理员权限。"""
    now = _now_iso()
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, SYSTEM_ADMIN_OPEN_ID)
        if record is None:
            record = {
                'open_id': SYSTEM_ADMIN_OPEN_ID,
                'name': SYSTEM_ADMIN_NAME,
                'disabled': False,
                'is_admin': True,
                'created_at': now,
                'last_login_at': now,
            }
        else:
            record['name'] = SYSTEM_ADMIN_NAME
            record['disabled'] = False
            record['is_admin'] = True
            record.setdefault('created_at', now)
        return sqlite_store.upsert_user(_base_dir(), config, record)


def _get_user_lock(open_id):
    with _lock:
        if open_id not in _user_locks:
            _user_locks[open_id] = threading.Lock()
        return _user_locks[open_id]


def upsert_user(profile):
    """新增或更新用户：存在则更新资料/token 并刷新 last_login_at，否则创建并记 created_at；返回是否落盘成功"""
    profile = profile or {}
    open_id = profile.get('open_id') or ''
    if not open_id:
        return False
    if is_system_admin(open_id):
        # OAuth 不能占用固定系统账号，也不能覆盖其保护属性。
        return False
    now = _now_iso()
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
        if record is None:
            record = {
                'open_id': open_id,
                'union_id': profile.get('union_id', ''),
                'user_id': profile.get('user_id', ''),
                'name': profile.get('name', ''),
                'avatar': profile.get('avatar', ''),
                'disabled': False,
                'is_admin': False,
                'access_token': profile.get('access_token', ''),
                'refresh_token': profile.get('refresh_token', ''),
                'token_expire_at': profile.get('token_expire_at', 0),
                'refresh_token_expire_at': profile.get('refresh_token_expire_at', 0),
                'scope': profile.get('scope', ''),
                'token_invalid': False,
                'created_at': now,
                'last_login_at': now,
            }
        else:
            # 清理旧版本残留的 department 字段（部门代码已移除，顺带迁移存量数据）
            record.pop('department', None)
            for key in ('union_id', 'user_id', 'name', 'avatar',
                        'access_token', 'refresh_token', 'token_expire_at',
                        'refresh_token_expire_at', 'scope'):
                if key in profile:
                    record[key] = profile[key]
            # 重新登录/刷新成功会带来新凭证，清除凭证失效标记
            if profile.get('access_token'):
                record['token_invalid'] = False
            record['last_login_at'] = now
        return sqlite_store.upsert_user(_base_dir(), config, record)


def get_user(open_id):
    """获取单个用户记录副本，不存在返回 None"""
    if not open_id:
        return None
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
    if record is not None:
        record['is_system'] = is_system_admin(record.get('open_id'))
    return record


def list_users(page=None, page_size=None, query=''):
    """按 last_login_at 倒序返回用户；传分页参数时由 SQLite 直接分页。"""
    with _lock:
        result = sqlite_store.list_users(_base_dir(), config, page=page, page_size=page_size, query=query)
    for record in result['items']:
        record['is_system'] = is_system_admin(record.get('open_id'))
    if page is None and page_size is None:
        return result['items']
    return result


def set_disabled(open_id, disabled):
    """设置用户禁用状态，用户不存在或落盘失败时返回 False"""
    if is_system_admin(open_id):
        return False
    with _lock:
        return sqlite_store.update_user_fields(_base_dir(), config, open_id, disabled=disabled)


def set_admin(open_id, is_admin):
    """设置用户的后台管理员权限；用户不存在或落盘失败时返回 False"""
    if is_system_admin(open_id):
        return bool(is_admin)
    with _lock:
        return sqlite_store.update_user_fields(_base_dir(), config, open_id, is_admin=is_admin)


def is_admin(open_id):
    """返回用户是否拥有后台管理员权限。"""
    if not open_id:
        return False
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
    return bool(record and record.get('is_admin', False) and not record.get('disabled', False))


def is_disabled(open_id):
    """用户是否被禁用；用户不存在时按禁用处理返回 True"""
    if not open_id:
        return True
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
    if record is None:
        return True
    return bool(record.get('disabled', False))


def _mark_token_invalid(open_id):
    """刷新失败时标记该用户凭证失效，供管理后台后续展示；落盘失败仅记日志不阻断"""
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
        if record is None or record.get('token_invalid'):
            return
        if not sqlite_store.update_user_fields(_base_dir(), config, open_id, token_invalid=True):
            logger.warning(f'标记用户 {open_id} 凭证失效状态落库失败')


def get_refresh_token_expiry(open_id):
    """返回该用户 refresh_token 到期时间戳；未知（旧数据无字段）返回 0，不存在返回 None"""
    if not open_id:
        return None
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
    if record is None:
        return None
    try:
        return int(record.get('refresh_token_expire_at') or 0)
    except (TypeError, ValueError):
        return 0


def get_valid_access_token(open_id):
    """获取该用户有效的 access_token。

    - 用户不存在/被禁用返回 None
    - token 距过期超过 5 分钟直接返回缓存值
    - 否则在按 open_id 的专属锁内串行刷新，刷新成功落库并返回新 token
    - 刷新失败或异常返回 None
    """
    import time as _time
    if not open_id:
        return None
    with _lock:
        record = sqlite_store.get_user(_base_dir(), config, open_id)
    if record is None:
        return None
    if record.get('disabled'):
        return None
    # 升级前登录的旧记录无 scope 字段，其 token 不含新权限且刷新无法增量获取，
    # 视为凭证失效（与 token_invalid 相同的处理路径），引导用户重新登录
    if not (record.get('scope') or '').strip():
        logger.warning(f'用户 {open_id} 凭证缺少 scope（升级前登录），视为失效，需重新登录')
        _mark_token_invalid(open_id)
        return None
    now = int(_time.time())
    access_token = record.get('access_token') or ''
    try:
        expire_at = int(record.get('token_expire_at') or 0)
    except (TypeError, ValueError):
        expire_at = 0
    if access_token and expire_at - now > _token_refresh_margin:
        return access_token

    user_lock = _get_user_lock(open_id)
    with user_lock:
        # 双检：可能已被其他线程刷新
        with _lock:
            record = sqlite_store.get_user(_base_dir(), config, open_id)
        if record is None or record.get('disabled'):
            return None
        access_token = record.get('access_token') or ''
        try:
            expire_at = int(record.get('token_expire_at') or 0)
        except (TypeError, ValueError):
            expire_at = 0
        if access_token and expire_at - now > _token_refresh_margin:
            return access_token
        refresh_token = record.get('refresh_token') or ''
        if not refresh_token:
            logger.warning(f'用户 {open_id} 无 refresh_token，无法刷新 access_token')
            _mark_token_invalid(open_id)
            return None
        try:
            # 延迟导入避免与 feishu_oauth 的潜在循环依赖
            from src.core.feishu_oauth import refresh_user_token
            token_data = refresh_user_token(refresh_token)
        except Exception as e:
            logger.warning(f'刷新用户 {open_id} token 异常: {e}')
            return None
        if not token_data or not token_data.get('access_token'):
            logger.warning(f'刷新用户 {open_id} token 失败')
            _mark_token_invalid(open_id)
            return None
        with _lock:
            current = sqlite_store.get_user(_base_dir(), config, open_id)
            if current is None:
                return None
            current['access_token'] = token_data.get('access_token') or ''
            current['refresh_token'] = token_data.get('refresh_token') or refresh_token
            # 刷新响应携带 scope 时更新；未携带则保留记录中原有值
            if token_data.get('scope'):
                current['scope'] = token_data['scope']
            try:
                expires_in = int(token_data.get('expires_in') or 0)
            except (TypeError, ValueError):
                expires_in = 0
            current['token_expire_at'] = int(_time.time()) + expires_in
            try:
                refresh_expires_in = int(token_data.get('refresh_token_expires_in') or 0)
            except (TypeError, ValueError):
                refresh_expires_in = 0
            if refresh_expires_in <= 0:
                refresh_expires_in = _refresh_token_default_ttl
            current['refresh_token_expire_at'] = int(_time.time()) + refresh_expires_in
            # 刷新成功即凭证恢复有效
            current['token_invalid'] = False
            if not sqlite_store.upsert_user(_base_dir(), config, current):
                logger.error(f'用户 {open_id} refresh_token 已轮换但落库失败，重启后该用户需重新登录')
            return current['access_token']
