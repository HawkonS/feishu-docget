import os
import json
import threading
from datetime import datetime

from src.core.config_loader import ConfigLoader, config

logger = ConfigLoader.get_logger('feishu_docget')

# 全局可重入锁保护内存缓存与文件写入；每个 open_id 另有独立锁用于串行刷新 token
_lock = threading.RLock()
_user_locks = {}
_cache = None
_token_refresh_margin = 300  # token 距过期不足 5 分钟时触发刷新
_refresh_token_default_ttl = 30 * 24 * 3600  # 飞书未返回 refresh_token 有效期时按 30 天兜底


def get_users_file():
    """用户库文件路径：<workspace>/<log.dir>/users.json（拼法参照 stats.get_stats_file）"""
    workspace = config.get('workspace.dir', '.')
    log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, 'users.json')


def _now_iso():
    return datetime.now().isoformat()


def _load():
    """懒加载：首次访问时全量读入内存缓存"""
    global _cache
    if _cache is not None:
        return _cache
    path = get_users_file()
    data = {}
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    data = loaded
        except Exception as e:
            logger.error(f'读取用户库失败 {path}: {e}')
    _cache = data
    return _cache


def _save(users):
    """写临时文件后 os.replace() 原子替换，替换前先 chmod 0600 收紧权限；返回是否落盘成功"""
    path = get_users_file()
    tmp_path = path + '.tmp'
    try:
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
        os.chmod(tmp_path, 0o600)
        os.replace(tmp_path, path)
        try:
            os.chmod(path, 0o600)
        except Exception:
            pass
        return True
    except Exception as e:
        logger.error(f'写入用户库失败 {path}: {e}')
        return False


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
    now = _now_iso()
    with _lock:
        users = _load()
        record = users.get(open_id)
        if record is None:
            record = {
                'open_id': open_id,
                'union_id': profile.get('union_id', ''),
                'user_id': profile.get('user_id', ''),
                'name': profile.get('name', ''),
                'department': profile.get('department', ''),
                'avatar': profile.get('avatar', ''),
                'disabled': False,
                'access_token': profile.get('access_token', ''),
                'refresh_token': profile.get('refresh_token', ''),
                'token_expire_at': profile.get('token_expire_at', 0),
                'refresh_token_expire_at': profile.get('refresh_token_expire_at', 0),
                'token_invalid': False,
                'created_at': now,
                'last_login_at': now,
            }
        else:
            for key in ('union_id', 'user_id', 'name', 'department', 'avatar',
                        'access_token', 'refresh_token', 'token_expire_at',
                        'refresh_token_expire_at'):
                if key in profile:
                    record[key] = profile[key]
            # 重新登录/刷新成功会带来新凭证，清除凭证失效标记
            if profile.get('access_token'):
                record['token_invalid'] = False
            record['last_login_at'] = now
        users[open_id] = record
        return _save(users)


def get_user(open_id):
    """获取单个用户记录副本，不存在返回 None"""
    if not open_id:
        return None
    with _lock:
        record = _load().get(open_id)
        return dict(record) if record else None


def list_users():
    """按 last_login_at 倒序返回用户列表，剥离 access_token/refresh_token 敏感字段"""
    with _lock:
        records = [dict(r) for r in _load().values()]
    for record in records:
        record.pop('access_token', None)
        record.pop('refresh_token', None)
    records.sort(key=lambda r: r.get('last_login_at') or '', reverse=True)
    return records


def set_disabled(open_id, disabled):
    """设置用户禁用状态，用户不存在或落盘失败时返回 False"""
    with _lock:
        users = _load()
        record = users.get(open_id)
        if record is None:
            return False
        record['disabled'] = bool(disabled)
        return _save(users)


def is_disabled(open_id):
    """用户是否被禁用；用户不存在时按禁用处理返回 True"""
    if not open_id:
        return True
    with _lock:
        record = _load().get(open_id)
    if record is None:
        return True
    return bool(record.get('disabled', False))


def _mark_token_invalid(open_id):
    """刷新失败时标记该用户凭证失效，供管理后台后续展示；落盘失败仅记日志不阻断"""
    with _lock:
        users = _load()
        record = users.get(open_id)
        if record is None or record.get('token_invalid'):
            return
        record['token_invalid'] = True
        if not _save(users):
            logger.warning(f'标记用户 {open_id} 凭证失效状态落库失败')


def get_refresh_token_expiry(open_id):
    """返回该用户 refresh_token 到期时间戳；未知（旧数据无字段）返回 0，不存在返回 None"""
    if not open_id:
        return None
    with _lock:
        record = _load().get(open_id)
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
        record = _load().get(open_id)
    if record is None:
        return None
    if record.get('disabled'):
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
            record = _load().get(open_id)
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
            users = _load()
            current = users.get(open_id)
            if current is None:
                return None
            current['access_token'] = token_data.get('access_token') or ''
            current['refresh_token'] = token_data.get('refresh_token') or refresh_token
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
            if not _save(users):
                logger.error(f'用户 {open_id} refresh_token 已轮换但落库失败，重启后该用户需重新登录')
            return current['access_token']
