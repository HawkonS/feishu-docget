import sys
import os

# 确保项目根目录在 Python 路径中
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

import json
import threading
import mimetypes
import uuid
import shutil
import queue
import time
from collections import deque
from functools import wraps
from datetime import datetime, timedelta
import subprocess
import secrets
import tempfile
import zipfile
import hmac
from html import escape
from urllib.parse import urlsplit
from flask import Flask, jsonify, request, send_file, send_from_directory, session, redirect, url_for
from src.services.doc_service import process_document
from src.core.bot_store import normalize_bot_config
from src.core import user_store, feishu_oauth
from src.core.user_store import SYSTEM_ADMIN_OPEN_ID, SYSTEM_ADMIN_NAME
from src.core.config_loader import config, ConfigLoader, parse_size, SENSITIVE_KEYS, ALLOWED_CONFIG_KEYS, validate_config
from src.converters.docx.style_manager import TableStyleManager
from src.core.stats import update_download_stat, get_download_stats
from src.core.sqlite_store import initialize_database, migrate_legacy_data, delete_download_stats, DATABASE_FILENAME
from src.core.utils import sanitize_name
base_dir = os.path.abspath(config.get('workspace.dir', '.'))
# 飞书 OAuth 回调路径（redirect_uri 推导与回调路由共用单源）
OAUTH_CALLBACK_PATH = '/auth/feishu/callback'
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_DIR = os.path.join(CURRENT_DIR, 'web', 'templates')
logger = ConfigLoader.get_logger('feishu_docget')

# 首次升级时自动把旧 JSON/JSONL 导入数据库，保留原文件作为回滚备份；
# tools/migrate_json_to_sqlite.py 可在停服窗口中显式执行并生成额外备份。
try:
    initialize_database(base_dir, config)
    migration_result = migrate_legacy_data(base_dir, config, backup=False)
    if migration_result.get('users_imported') or migration_result.get('stats_imported'):
        logger.info(
            '旧用户/统计数据已导入 SQLite：users=%s, stats=%s',
            migration_result.get('users_imported', 0),
            migration_result.get('stats_imported', 0),
        )
    # 密码登录使用独立的固定系统账号，不占用任何飞书用户身份。
    user_store.ensure_system_admin()
except Exception as e:
    # 数据库故障不应在导入模块阶段吞掉堆栈；记录后让后续请求返回明确错误。
    logger.error(f'初始化 SQLite 数据库失败: {e}', exc_info=True)
app = Flask(__name__)
secret_key = config.get('server.secret_key', '').strip()
if not secret_key or secret_key == 'feishu_docget_secret_key_2025':
    secret_key = secrets.token_hex(32)
    try:
        ConfigLoader.save_config_from_admin({'server.secret_key': secret_key})
        logger.warning('已自动生成新的 SECRET_KEY 并写入配置文件，请妥善保管')
    except Exception as e:
        logger.error(f'自动保存 SECRET_KEY 失败: {e}')
app.secret_key = secret_key
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
# Cookie 的 Secure 标志必须在应用初始化时设置，避免使用 waitress/WSGI 导入方式
# 启动时退回不安全的 HTTP Cookie。生产环境应启用 server.https.enabled。
https_enabled = ConfigLoader.get_bool('server.https.enabled', False)
app.config['SESSION_COOKIE_SECURE'] = https_enabled
if https_enabled:
    # server.https.enabled 表示由一层可信反向代理终结 TLS；初始化阶段包装，
    # 确保通过 WSGI 导入 app 时同样正确识别客户端 IP 与 https scheme。
    from werkzeug.middleware.proxy_fix import ProxyFix
    # 不信任 X-Forwarded-For，防止直接访问服务时伪造来源 IP 绕过限流。
    # scheme/host/prefix 仅接受最外层一跳代理提供的值。
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=0, x_proto=1, x_host=1, x_prefix=1)
    app.config['PREFERRED_URL_SCHEME'] = 'https'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB


@app.errorhandler(500)
def handle_500(e):
    """记录未捕获异常的完整堆栈；下载等 URL 由浏览器直接导航访问，返回默认错误页"""
    logger.error(f'服务器内部错误: {request.method} {request.path}', exc_info=True)
    return e.get_response()

jobs = {}
jobs_lock = threading.Lock()
download_queue = queue.Queue()
active_downloads_lock = threading.Lock()
# active_downloads 仍表示整个进程中正在执行的任务总数，便于日志和管理端查看。
# 个人登录任务另按 open_id 计数；机器人任务共享 active_bot_downloads 计数。
active_downloads = 0
active_downloads_by_account = {}
active_bot_downloads = 0
download_dispatch_condition = threading.Condition(active_downloads_lock)

_login_attempts = {}  # {ip: {'count': N, 'locked_until': timestamp, 'ban_level': N}}
_admin_account_attempts = {}  # 单一后台账号的全局退避，防止轮换 IP 绕过限流
_upload_attempts = {}  # {ip: {'count': N, 'locked_until': timestamp, 'ban_level': N}} 模板上传密码尝试记录
_start_attempts = {}  # {ip: [timestamps]} 任务提交接口的滑动窗口限流
_attempts_lock = threading.RLock()
START_RATE_LIMIT = 10  # 单 IP 每分钟最多提交任务数


def _check_login_lock(attempts, client_ip):
    """检查 IP 是否处于锁定状态。返回 (是否锁定, 剩余秒数)，锁定期过期时自动清理记录"""
    with _attempts_lock:
        attempt = attempts.get(client_ip)
        if attempt and attempt.get('locked_until', 0) > 0:
            now = time.time()
            if now > attempt['locked_until']:
                attempts.pop(client_ip, None)
                return False, 0
            return True, int(attempt['locked_until'] - now)
        return False, 0


def _record_login_failure(attempts, client_ip, failure_limit=5):
    """记录一次密码校验失败，失败 5 次后按指数退避锁定"""
    with _attempts_lock:
        now = time.time()
        attempt = attempts.get(client_ip)
        if not attempt:
            attempts[client_ip] = {'count': 1, 'locked_until': 0, 'ban_level': 0}
        else:
            attempt['count'] = attempt.get('count', 0) + 1
            if attempt['count'] >= failure_limit:
                ban_level = attempt.get('ban_level', 0)
                lock_seconds = min(60 * (3 ** ban_level), 3600)
                attempt['locked_until'] = now + lock_seconds
                attempt['count'] = 0
                attempt['ban_level'] = ban_level + 1


def _admin_login_lock(client_ip):
    """同时检查 IP 和后台账号维度的登录退避状态。"""
    ip_locked, ip_remaining = _check_login_lock(_login_attempts, client_ip)
    account_locked, account_remaining = _check_login_lock(_admin_account_attempts, 'admin')
    if ip_locked or account_locked:
        return True, max(ip_remaining, account_remaining)
    return False, 0


def _record_admin_login_failure(client_ip):
    _record_login_failure(_login_attempts, client_ip)
    # 账号维度使用更高阈值，既阻止轮换 IP 暴力破解，也降低恶意锁死账号的风险。
    _record_login_failure(_admin_account_attempts, 'admin', failure_limit=20)


def _clear_admin_login_failures(client_ip):
    with _attempts_lock:
        _login_attempts.pop(client_ip, None)
        _admin_account_attempts.pop('admin', None)


def _get_csrf_token():
    token = session.get('_csrf_token')
    if not token:
        token = secrets.token_urlsafe(32)
        session['_csrf_token'] = token
        session.modified = True
    return token


@app.before_request
def ensure_csrf_token_and_protect_unsafe_requests():
    """为每个会话生成 CSRF token，并保护所有状态修改请求。"""
    _get_csrf_token()
    if request.method not in {'POST', 'PUT', 'PATCH', 'DELETE'}:
        return None
    expected = session.get('_csrf_token')
    supplied = request.headers.get('X-CSRF-Token') or request.form.get('_csrf_token')
    if not expected or not supplied or not hmac.compare_digest(str(expected), str(supplied)):
        return jsonify({'status': 'error', 'message': 'CSRF 校验失败，请刷新页面后重试'}), 403
    return None


def _check_start_rate_limit(client_ip):
    """任务提交接口的滑动窗口限流，超限返回 False"""
    with _attempts_lock:
        now = time.time()
        window = _start_attempts.setdefault(client_ip, [])
        window[:] = [ts for ts in window if now - ts < 60]
        if len(window) >= START_RATE_LIMIT:
            return False
        window.append(now)
        return True


def _verify_job_access(job_id):
    """校验当前会话是否有权访问指定任务：管理员或任务创建者可访问"""
    if _is_admin_session():
        return True
    tokens = session.get('job_tokens') or {}
    return bool(tokens.get(job_id))


def _cleanup_temp(path):
    """延迟清理临时文件"""
    threading.Timer(60, lambda: os.path.exists(path) and os.remove(path)).start()


def _login_enabled():
    """飞书登录开关，由配置 login.enabled 控制"""
    return ConfigLoader.get_bool('login.enabled', False)


def _is_admin_session():
    """当前会话是否具备后台管理员权限。

    ``is_admin`` 是管理员密码登录标志；飞书用户管理员权限持久化在用户库中，
    因此每次请求都重新读取，管理员被取消或禁用后立即失效。
    """
    if session.get('is_admin'):
        return True
    user = session.get('user') or {}
    return user_store.is_admin(user.get('open_id', ''))


def _is_system_admin_session():
    """密码登录的系统管理员会话；飞书管理员不会命中该标志。"""
    if not session.get('is_admin'):
        return False
    user = session.get('user') or {}
    # 兼容早期只写入 is_admin 的密码会话；若会话同时带有飞书用户身份，保留其真实姓名。
    return not user or user_store.is_system_admin(user.get('open_id', ''))


def _get_redirect_uri():
    """OAuth 回调地址：优先使用配置项，未配置时根据当前请求地址推导（HTTPS 模式下强制 https scheme）"""
    configured = str(config.get('login.oauth.redirect_uri', '') or '').strip()
    if configured:
        return configured
    base = request.url_root.rstrip('/')
    # nginx SSL 终结且未传 X-Forwarded-Proto 时 url_root 会是 http，需按配置纠正 scheme
    if ConfigLoader.get_bool('server.https.enabled', False) and base.startswith('http://'):
        base = 'https://' + base[len('http://'):]
    return base + OAUTH_CALLBACK_PATH


def _inject_csrf(html):
    """将会话 CSRF token 注入页面，并让同源写请求自动携带请求头。"""
    token = escape(_get_csrf_token(), quote=True)
    html = html.replace('[/* csrf_token */]', token)
    if 'name="csrf-token"' not in html:
        csrf_bootstrap = f'''<meta name="csrf-token" content="{token}">
<script>
(function () {{
  const csrfToken = document.querySelector('meta[name="csrf-token"]').content;
  const originalFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {{
    const options = Object.assign({{}}, init || {{}});
    const method = String(options.method || (input && input.method) || 'GET').toUpperCase();
    if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) {{
      const target = new URL(typeof input === 'string' ? input : input.url, window.location.href);
      if (target.origin === window.location.origin) {{
        const headers = new Headers(options.headers || (input && input.headers) || {{}});
        headers.set('X-CSRF-Token', csrfToken);
        options.headers = headers;
      }}
    }}
    return originalFetch(input, options);
  }};
}})();
</script>'''
        html = html.replace('</head>', csrf_bootstrap + '</head>', 1)
    return html


def _script_json(value):
    """将数据安全嵌入 script：JSON 中的 HTML 特殊字符使用 Unicode 转义。"""
    return json.dumps(value, ensure_ascii=False).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026').replace('\u2028', '\\u2028').replace('\u2029', '\\u2029')


def _safe_http_url(value, fallback='#'):
    raw = str(value or '').strip()
    try:
        parsed = urlsplit(raw)
        if parsed.scheme.lower() in {'http', 'https'} and parsed.netloc:
            return parsed.geturl()
    except ValueError:
        pass
    return fallback


def _safe_admin_path(value):
    raw = str(value or '').strip()
    return raw if raw.startswith('/') and not raw.startswith('//') else '/admin'


def login_required(f):
    # 未启用登录或当前会话具备管理员权限时直通；页面请求重定向 /login，API 请求返回 403 JSON
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not _login_enabled() or _is_admin_session():
            return f(*args, **kwargs)
        user = session.get('user')
        open_id = (user or {}).get('open_id', '')
        blocked = False
        if not open_id:
            blocked = True
        elif user_store.is_disabled(open_id):
            session.pop('user', None)
            blocked = True
        if blocked:
            if request.path == '/' or request.accept_mimetypes.best == 'text/html':
                return redirect('/login')
            return (jsonify({'status': 'error', 'message': '请先登录'}), 403)
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not _is_admin_session():
            return (jsonify({'status': 'error', 'message': '未登录'}), 403)
        return f(*args, **kwargs)
    return decorated_function

def _download_scope(user_open_id=''):
    """返回任务的并发作用域。

    user_open_id 只会在个人登录任务中传入；空值表示机器人/未登录任务，
    这些任务继续共享一个并发池，以保留原有的机器人请求限流行为。
    """
    user_open_id = str(user_open_id or '').strip()
    if user_open_id:
        return 'account', user_open_id
    return 'bot', '__bot__'


def _scope_limit(scope_type):
    """读取并发上限。配置项保持兼容，但登录任务按账号分别计算。"""
    return max(1, ConfigLoader.get_int('max.concurrent.downloads', 1))


def _scope_active_count_locked(scope_type, scope_id):
    if scope_type == 'account':
        return active_downloads_by_account.get(scope_id, 0)
    return active_bot_downloads


def _scope_has_slot_locked(scope_type, scope_id):
    return _scope_active_count_locked(scope_type, scope_id) < _scope_limit(scope_type)


def _scope_job_counts(user_open_id=''):
    """统计当前作用域中已运行/等待的任务，用于提交接口展示排队信息。"""
    scope_type, scope_id = _download_scope(user_open_id)
    active = 0
    pending = 0
    with jobs_lock:
        for job in jobs.values():
            job_scope_type, job_scope_id = _download_scope(job.get('user_open_id', ''))
            if (job_scope_type, job_scope_id) != (scope_type, scope_id):
                continue
            if job.get('status') == 'running':
                active += 1
            elif job.get('status') == 'pending':
                pending += 1
    return scope_type, scope_id, active, pending


def _mark_download_started(user_open_id=''):
    """在调度锁内占用一个并发槽位。"""
    global active_downloads, active_bot_downloads
    scope_type, scope_id = _download_scope(user_open_id)
    active_downloads += 1
    if scope_type == 'account':
        active_downloads_by_account[scope_id] = active_downloads_by_account.get(scope_id, 0) + 1
    else:
        active_bot_downloads += 1
    return scope_type, scope_id


def _mark_download_finished(scope_type, scope_id):
    """释放并发槽位并唤醒调度器。调用方必须持有 active_downloads_lock。"""
    global active_downloads, active_bot_downloads
    active_downloads = max(0, active_downloads - 1)
    if scope_type == 'account':
        count = active_downloads_by_account.get(scope_id, 1) - 1
        if count > 0:
            active_downloads_by_account[scope_id] = count
        else:
            active_downloads_by_account.pop(scope_id, None)
    else:
        active_bot_downloads = max(0, active_bot_downloads - 1)


def _run_download_job(job_args, scope_type, scope_id):
    """执行单个任务，并在结束后释放对应账号/机器人并发槽位。"""
    global active_downloads
    try:
        run_job(*job_args)
    except Exception as e:
        logger.error(f'工作线程任务错误: {e}')
    finally:
        with download_dispatch_condition:
            _mark_download_finished(scope_type, scope_id)
            download_dispatch_condition.notify_all()
            current_active = active_downloads
        download_queue.task_done()
        logger.info(f"任务完成，当前活动任务数: {current_active}")


def worker_thread():
    """调度下载队列。

    调度器只在对应作用域有空闲槽位时才取出任务执行，避免某个账号的排队任务
    占满工作线程后把其他账号也饿死。每个已获准的任务使用独立守护线程执行，
    因此不同个人账号可以同时下载，而同一账号仍受 max.concurrent.downloads 限制。
    """
    pending = deque()
    logger.info("下载任务调度线程已启动，等待任务...")
    while True:
        try:
            # 至少接收一个任务，随后尽可能批量收取，减少队列高峰时的调度延迟。
            if not pending:
                try:
                    pending.append(download_queue.get(timeout=0.5))
                except queue.Empty:
                    pass
            while True:
                try:
                    pending.append(download_queue.get_nowait())
                except queue.Empty:
                    break

            with download_dispatch_condition:
                selected_index = None
                selected_scope = None
                for index, job_args in enumerate(pending):
                    user_open_id = job_args[20] if len(job_args) > 20 else ''
                    scope_type, scope_id = _download_scope(user_open_id)
                    if _scope_has_slot_locked(scope_type, scope_id):
                        selected_index = index
                        selected_scope = (scope_type, scope_id)
                        break
                if selected_index is None:
                    if pending:
                        # 当前所有任务都在等待各自账号的槽位释放。
                        download_dispatch_condition.wait(timeout=0.5)
                    continue
                job_args = pending[selected_index]
                del pending[selected_index]
                _mark_download_started(job_args[20] if len(job_args) > 20 else '')
                scope_type, scope_id = selected_scope
                logger.info(f"调度任务，作用域={scope_type}:{scope_id}，当前活动任务数={active_downloads}")

            try:
                threading.Thread(
                    target=_run_download_job,
                    args=(job_args, scope_type, scope_id),
                    daemon=True,
                    name=f'download-{job_args[0]}',
                ).start()
            except Exception:
                # 极少见的线程创建失败也必须归还槽位并结束队列任务，避免永久卡死。
                with download_dispatch_condition:
                    _mark_download_finished(scope_type, scope_id)
                    download_dispatch_condition.notify_all()
                download_queue.task_done()
                raise
        except Exception as e:
            logger.error(f'下载任务调度循环错误: {e}', exc_info=True)
            time.sleep(1)


threading.Thread(target=worker_thread, daemon=True, name='download-dispatcher').start()

def token_pre_refresh_thread():
    """后台预刷新线程：定期为 access_token 即将过期的用户提前刷新，
    通过 refresh_token 轮换实现事实性续期，避免长期不下载导致 refresh_token 30 天过期"""
    logger.info('token 预刷新线程已启动')
    while True:
        try:
            time.sleep(30 * 60)
            if not _login_enabled():
                continue
            now = int(time.time())
            for record in user_store.list_users():
                try:
                    if record.get('disabled'):
                        continue
                    open_id = record.get('open_id') or ''
                    if not open_id:
                        continue
                    # 升级前的旧记录无 scope，预刷新无法获得新权限，跳过并引导重新登录
                    if not (record.get('scope') or '').strip():
                        logger.warning(f'用户 {open_id} 凭证缺少 scope（升级前登录），预刷新跳过，需重新登录')
                        continue
                    try:
                        expire_at = int(record.get('token_expire_at') or 0)
                    except (TypeError, ValueError):
                        expire_at = 0
                    # 仅对 access_token 距过期不足 1 小时（或已过期）的用户触发刷新
                    if expire_at - now > 3600:
                        continue
                    refresh_expire_at = user_store.get_refresh_token_expiry(open_id) or 0
                    if refresh_expire_at and refresh_expire_at <= now:
                        logger.warning(f'用户 {open_id} 的 refresh_token 已过期，预刷新跳过，需重新登录')
                        continue
                    # 复用 get_valid_access_token 的按用户锁 + 双重检查，天然防并发
                    token = user_store.get_valid_access_token(open_id)
                    if token:
                        logger.info(f'预刷新用户 {open_id} token 成功')
                    else:
                        logger.warning(f'预刷新用户 {open_id} token 失败，该用户下次任务前需重新登录')
                except Exception as e:
                    logger.warning(f'预刷新用户 {record.get("open_id", "")} 处理异常: {e}')
        except Exception as e:
            logger.error(f'token 预刷新线程循环错误: {e}')
            time.sleep(60)
threading.Thread(target=token_pre_refresh_thread, daemon=True).start()

def check_cleanup_output():
    try:
        output_dir = os.path.join(base_dir, config['output.dir'])
        if not os.path.exists(output_dir):
            return
        items = []
        total_size = 0
        for name in os.listdir(output_dir):
            path = os.path.join(output_dir, name)
            if os.path.isdir(path):
                size = 0
                for dirpath, dirnames, filenames in os.walk(path):
                    for f in filenames:
                        fp = os.path.join(dirpath, f)
                        if not os.path.islink(fp):
                            size += os.path.getsize(fp)
                items.append({'path': path, 'size': size, 'ctime': os.path.getctime(path)})
                total_size += size
        limit = ConfigLoader.get_size('output.max_size', parse_size('10G'))
        if total_size > limit:
            items.sort(key=lambda x: x['ctime'])
            for item in items:
                if total_size <= limit:
                    break
                try:
                    shutil.rmtree(item['path'])
                    total_size -= item['size']
                    logger.info(f"已清理: {item['path']}")
                except Exception as e:
                    logger.error(f'清理失败: {str(e)}')
    except Exception as e:
        logger.error(f'清理错误: {str(e)}')

def list_templates():
    # 重新加载配置以确保获取最新默认值
    current_config = ConfigLoader.load_config()
    template_dir = os.path.join(base_dir, current_config['template.dir'])
    if not os.path.isdir(template_dir):
        return []
    items = []
    default_template = current_config.get('template.default', 'template.docx')
    for name in os.listdir(template_dir):
        # 过滤掉临时文件，不显示在列表中
        if name.startswith('temp_'):
            continue
            
        if name.lower().endswith('.docx'):
            path = os.path.join(template_dir, name)
            size = os.path.getsize(path) if os.path.exists(path) else 0
            png_name = os.path.splitext(name)[0] + '.png'
            png_path = os.path.join(template_dir, png_name)
            has_png = os.path.exists(png_path)
            pdf_name = os.path.splitext(name)[0] + '.pdf'
            pdf_path = os.path.join(template_dir, pdf_name)
            has_pdf = os.path.exists(pdf_path)
            is_default = name == default_template
            
            items.append({'name': name, 'display_name': name, 'size': size, 'has_png': has_png, 'png_name': png_name if has_png else None, 'has_pdf': has_pdf, 'pdf_name': pdf_name if has_pdf else None, 'is_default': is_default, 'is_temp': False})
    items.sort(key=lambda x: (not x['is_default'], x['name']))
    return items


def paginate_items(items, page=1, page_size=20):
    """返回管理后台统一使用的服务端分页结果。"""
    try:
        page = max(int(page or 1), 1)
    except (TypeError, ValueError):
        page = 1
    try:
        page_size = min(max(int(page_size or 20), 1), 100)
    except (TypeError, ValueError):
        page_size = 20
    total = len(items)
    start = (page - 1) * page_size
    total_pages = (total + page_size - 1) // page_size if total else 0
    if total_pages:
        page = min(page, total_pages)
    else:
        page = 1
    start = (page - 1) * page_size
    page_items = items[start:start + page_size]
    return {
        'items': page_items,
        'total': total,
        'page': page,
        'page_size': page_size,
        'total_pages': total_pages,
        'has_more': start + len(page_items) < total,
    }


def _resolve_template_path(name, allow_empty=False):
    """解析模板文件名，并保证真实文件严格位于 template.dir 内。"""
    raw_name = str(name or '').strip()
    if not raw_name and allow_empty:
        return ''
    if not raw_name or raw_name != os.path.basename(raw_name) or raw_name in {'.', '..'}:
        return None
    if not raw_name.lower().endswith('.docx'):
        return None
    template_root = os.path.realpath(os.path.join(base_dir, config['template.dir']))
    candidate = os.path.realpath(os.path.join(template_root, raw_name))
    if not candidate.startswith(template_root + os.sep):
        return None
    if not os.path.isfile(candidate):
        return None
    return candidate

def list_projects(include_files=True):
    output_dir = os.path.join(base_dir, config['output.dir'])
    items = []
    if os.path.isdir(output_dir):
        for name in os.listdir(output_dir):
            path = os.path.join(output_dir, name)
            if os.path.isdir(path):
                try:
                    ctime = os.path.getctime(path)
                    size = 0
                    for root, _, filenames in os.walk(path):
                        for f in filenames:
                            fp = os.path.join(root, f)
                            if not os.path.islink(fp):
                                size += os.path.getsize(fp)
                except Exception:
                    ctime = 0
                    size = 0
                files = list_project_files(path) if include_files else []
                file_count = len(files) if include_files else None
                items.append({'name': name, 'path': path, 'ctime': datetime.fromtimestamp(ctime).isoformat(timespec='minutes'), 'ctime_ts': ctime, 'size': size, 'file_count': file_count, 'files': files})
    return sorted(items, key=lambda x: x.get('ctime_ts', 0), reverse=True)


def list_project_files(path):
    """按需读取单个项目的文件树，避免项目列表接口返回全部文件明细。"""
    files = []
    if not path or not os.path.isdir(path):
        return files
    for root, _, filenames in os.walk(path):
        for fname in filenames:
            abs_path = os.path.join(root, fname)
            rel_path = os.path.relpath(abs_path, path).replace('\\', '/')
            try:
                fctime = os.path.getctime(abs_path)
            except Exception:
                fctime = 0
            files.append({'name': fname, 'rel_path': rel_path, 'path': abs_path,
                          'ctime': datetime.fromtimestamp(fctime).isoformat(timespec='minutes'),
                          'is_md': fname.endswith('.md')})
    files.sort(key=lambda x: x['ctime'], reverse=True)
    return files


def list_project_summaries(page=1, page_size=20, query=''):
    """分页读取项目摘要。

    先只扫描项目目录本身完成排序/分页，再为当前页计算大小和文件数，
    这样不会因为历史项目文件树越来越大而拖慢每次 Tab 切换。
    """
    output_dir = os.path.join(base_dir, config['output.dir'])
    candidates = []
    query = str(query or '').strip().lower()
    if os.path.isdir(output_dir):
        for name in os.listdir(output_dir):
            path = os.path.join(output_dir, name)
            if not os.path.isdir(path) or (query and query not in name.lower()):
                continue
            try:
                ctime = os.path.getctime(path)
            except OSError:
                ctime = 0
            candidates.append((ctime, name, path))
    candidates.sort(key=lambda item: item[0], reverse=True)
    total = len(candidates)
    try:
        page = max(int(page or 1), 1)
    except (TypeError, ValueError):
        page = 1
    try:
        page_size = min(max(int(page_size or 20), 1), 100)
    except (TypeError, ValueError):
        page_size = 20
    start = (page - 1) * page_size
    total_pages = (total + page_size - 1) // page_size if total else 0
    if total_pages:
        page = min(page, total_pages)
    else:
        page = 1
    start = (page - 1) * page_size
    selected = candidates[start:start + page_size]
    items = []
    for ctime, name, path in selected:
        size = 0
        file_count = 0
        for root, _, filenames in os.walk(path):
            file_count += len(filenames)
            for fname in filenames:
                fp = os.path.join(root, fname)
                if not os.path.islink(fp):
                    try:
                        size += os.path.getsize(fp)
                    except OSError:
                        pass
        items.append({'name': name, 'path': path,
                      'ctime': datetime.fromtimestamp(ctime).isoformat(timespec='minutes'),
                      'ctime_ts': ctime, 'size': size, 'file_count': file_count, 'files': []})
    return {'items': items, 'total': total, 'page': page,
            'page_size': page_size, 'total_pages': total_pages,
            'has_more': start + len(items) < total}

def update_job(job_id, **fields):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return
        log_type = fields.pop('log_type', 'info')
        if 'message' in fields:
            msg = fields.get('message') or ''
            if msg:
                logs = job.get('logs') or []
                if log_type == 'dynamic' and logs and (logs[-1].get('type') == 'dynamic'):
                    logs[-1]['message'] = msg
                    logs[-1]['ts'] = datetime.now().isoformat(timespec='seconds')
                elif log_type == 'success' and logs and (logs[-1].get('type') == 'dynamic'):
                    logs[-1]['message'] = msg
                    logs[-1]['ts'] = datetime.now().isoformat(timespec='seconds')
                    logs[-1]['type'] = 'success'
                else:
                    logs.append({'ts': datetime.now().isoformat(timespec='seconds'), 'message': msg, 'type': log_type})
                job['logs'] = logs[-200:]
        job.update(fields)

def run_job(job_id, doc_url, template_name, table_style, delete_template=False, add_cover=False, client_ip='', check_stop_func=None, unordered_list_style='default', body_style=None, was_queued=False, image_style=None, ignore_mention=False, ignore_template_heading_num=False, table_config=None, margin_config=None, code_block_config=None, document_info=None, add_title=False, bot_config=None, user_open_id='', user_name=''):
    try:
        logger.info(f"开始执行任务 {job_id}: {doc_url}")
        user_token = ''
        if user_open_id:
            # get_valid_access_token 内部已判断禁用状态，失败返回 None
            user_token = user_store.get_valid_access_token(user_open_id) or ''
            if not user_token:
                # 用户身份任务不回退机器人身份，避免误导性权限报错；直接按现有失败分支标记任务失败
                error_msg = '您的飞书登录凭证已过期，请重新登录后再试'
                logger.error(f'任务 {job_id} 失败: 用户 {user_open_id} 凭证已过期或刷新失败')
                update_job(job_id, status='error', message=error_msg)
                update_download_stat(base_dir, config, job_id, '错误', doc_url, ip_address=client_ip, user_name=user_name)
                return
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        if was_queued:
            update_job(job_id, message='已完成下载任务排队，成功创建下载任务', log_type='info')
        update_job(job_id, status='running', progress=5, message='正在准备任务...', log_type='dynamic')
        update_download_stat(base_dir, config, job_id, '下载中', doc_url=doc_url, ip_address=client_ip, user_name=user_name)
        template_path = _resolve_template_path(template_name, allow_empty=True)
        if template_path is None:
            raise ValueError('模板文件不存在或路径无效')
        output_root = os.path.join(base_dir, config['output.dir'])
        result = process_document(doc_url=doc_url, template_path=template_path, table_style=table_style, base_dir=base_dir, output_root=output_root, progress_cb=lambda p, m, t='info': update_job(job_id, progress=p, message=m, log_type=t), add_cover=add_cover, check_stop_func=check_stop_func, unordered_list_style=unordered_list_style, body_style=body_style, image_style=image_style, ignore_mention=ignore_mention, ignore_template_heading_num=ignore_template_heading_num, table_config=table_config, margin_config=margin_config, code_block_config=code_block_config, document_info=document_info, add_title=add_title, bot_config=bot_config, user_access_token=user_token or None)
        if delete_template and template_path:
            if os.path.exists(template_path):
                try:
                    os.remove(template_path)
                except Exception:
                    pass
            # 同时删除对应的预览图片
            png_path = os.path.splitext(template_path)[0] + '.png'
            if os.path.exists(png_path):
                try:
                    os.remove(png_path)
                except Exception:
                    pass
        update_job(job_id, status='done', progress=100, message='已完成', docx_path=result['docx_path'], folder=result['folder'])
        update_download_stat(base_dir, config, job_id, '已完成', doc_url, result['docx_path'], title=result.get('title', os.path.basename(result['docx_path'])), ip_address=client_ip, user_name=user_name)
        threading.Thread(target=check_cleanup_output).start()
    except Exception as e:
        is_stopped = isinstance(e, InterruptedError) or (check_stop_func and check_stop_func())
        if is_stopped:
            logger.info(f'任务 {job_id} 已被用户停止')
            update_job(job_id, status='stopped', message='任务已停止', log_type='error')
            update_download_stat(base_dir, config, job_id, '已停止', doc_url, ip_address=client_ip, user_name=user_name)
        else:
            logger.error('任务失败: ' + str(e))
            update_job(job_id, status='error', message=str(e))
            update_download_stat(base_dir, config, job_id, '错误', doc_url, ip_address=client_ip, user_name=user_name)

@app.errorhandler(404)
def page_not_found(e):
    target = config.get('url.404', 'https://space.hawkon.tech/')
    if not target.startswith('http'):
        target = 'http://' + target
    return redirect(target)

@app.errorhandler(413)
def request_entity_too_large(e):
    # 返回 JSON，避免前端解析 HTML 错误页报 Unexpected token '<'
    logger.warning(f'请求体超过大小限制: {request.path}')
    return jsonify({'status': 'error', 'message': '上传文件超过大小限制（50MB），请压缩后重试；若经由反向代理访问，请检查代理层的请求体大小限制'}), 413

def _validate_document_info(document_info):
    if not isinstance(document_info, dict):
        return None
    field_labels = {'created': '创建时间', 'modified': '上次修改时间', 'lastPrinted': '上次打印时间'}
    formats = ['%Y-%m-%dT%H:%M', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M', '%Y/%m/%d %H:%M:%S']
    for key, label in field_labels.items():
        raw_value = str(document_info.get(key) or '').strip()
        if not raw_value:
            continue
        is_valid = False
        for fmt in formats:
            try:
                datetime.strptime(raw_value, fmt)
                is_valid = True
                break
            except ValueError:
                continue
        if not is_valid:
            return f'{label}格式无效，请重新选择有效时间'
    return None


@app.route('/', methods=['GET'])
@login_required
def index():
    templates = list_templates()
    template_json = _script_json(templates)
    table_styles = TableStyleManager.list_styles()
    style_json = _script_json(table_styles)
    style_css = TableStyleManager.get_frontend_css()
    with open(os.path.join(HTML_DIR, 'index.html'), 'r', encoding='utf-8') as f:
        html = f.read()
    html = html.replace('[/* template_json */]', template_json)
    html = html.replace('[/* style_json */]', style_json)
    html = html.replace('/* [style_css] */', style_css)
    html = html.replace('[/* usage_url */]', escape(_safe_http_url(config.get('usage.url', 'https://github.com/HawkonS/feishu-docget'))))
    copyright_text = escape(config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('Hawkon 2025 -2026', copyright_text)
    html = html.replace('Hawkon 2025', copyright_text)
    html = html.replace('[/* page_title */]', escape(config.get('page.title', '飞书文档下载工具')))
    html = html.replace('[/* page_description */]', escape(config.get('page.description', '支持将飞书文档链接下载为指定模板的 Word 文件')))
    html = html.replace('[/* page_placeholder */]', escape(config.get('page.placeholder', '输入飞书文档链接，如 https://hawkon.feishu.cn/wiki/...'), quote=True))
    html = html.replace('[/* usage_link_text */]', escape(config.get('page.usage_link_text', '使用说明')))
    html = html.replace('[/* usage_doc_url */]', escape(_safe_http_url(config.get('url.usage_doc', 'https://github.com/HawkonS/feishu-docget'))))
    html = html.replace('[/* default_template */]', escape(config.get('template.default', 'template.docx'), quote=True))
    html = html.replace('[/* default_template_json */]', _script_json(config.get('template.default', 'template.docx')))
    html = html.replace('[/* image_max_width */]', str(config.get('image.max_width', '16')))
    html = html.replace('[/* image_max_height */]', str(config.get('image.max_height', '23')))
    # 登录状态标志：管理员和普通飞书用户都需要显示首页右上角用户卡片。
    # 密码管理员 session 不带 user/open_id，不能只依赖普通用户的登录标志判断。
    is_admin = _is_admin_session()
    is_system_admin = _is_system_admin_session()
    user_logged_in = 'true' if is_admin else 'false'
    if _login_enabled() and (session.get('user') or {}).get('open_id'):
        user_logged_in = 'true'
    html = html.replace('[/* user_logged_in */]', user_logged_in)
    html = html.replace('[/* user_is_admin */]', 'true' if is_admin else 'false')
    # 登录用户 open_id：用户名缺失时前端作为兜底展示文案；同样转义防注入
    user_open_id = ''
    if _login_enabled() and not is_system_admin:
        user_open_id = (session.get('user') or {}).get('open_id', '') or ''
    html = html.replace('[/* user_open_id */]', escape(user_open_id))
    # 登录用户头像 URL：从 user_store 读取（会话中不含头像）；未登录/无头像时注入空串，前端回退首字母样式
    user_avatar = ''
    if _login_enabled() and user_open_id:
        record = user_store.get_user(user_open_id)
        if record:
            user_avatar = record.get('avatar', '') or ''
    html = html.replace('[/* user_avatar */]', escape(user_avatar))
    # 登录用户名：仅作展示，HTML 转义防止注入；必须保持为最后一个替换，避免注入内容被二次替换
    user_name = SYSTEM_ADMIN_NAME if is_system_admin else ''
    if _login_enabled() and not is_system_admin:
        user_name = (session.get('user') or {}).get('name', '') or ''
    html = html.replace('[/* user_name */]', escape(user_name))
    html = html.replace('[/* admin_url */]', escape(_safe_admin_path(admin_path), quote=True))
    html = html.replace('[/* user_avatar_js */]', _script_json(user_avatar))
    html = _inject_csrf(html)
    return html

@app.route('/login', methods=['GET'])
def user_login_page():
    if not _login_enabled():
        return redirect('/')
    user = session.get('user')
    if user and user.get('open_id') and not user_store.is_disabled(user.get('open_id')):
        return redirect('/')
    with open(os.path.join(HTML_DIR, 'user_login.html'), 'r', encoding='utf-8') as f:
        html = f.read()
    html = html.replace('[/* page_title */]', escape(config.get('page.title', '飞书文档下载工具')))
    copyright_text = escape(config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('Hawkon 2025 -2026', copyright_text)
    html = html.replace('Hawkon 2025', copyright_text)
    return html

@app.route('/auth/feishu/authorize', methods=['GET'])
def auth_feishu_authorize():
    if not _login_enabled():
        return redirect('/')
    state = secrets.token_urlsafe(16)
    session['oauth_state'] = state
    return redirect(feishu_oauth.build_authorize_url(config.get('feishu.app_id'), _get_redirect_uri(), state))

@app.route(f'{OAUTH_CALLBACK_PATH}', methods=['GET'])
def auth_feishu_callback():
    expected_state = session.pop('oauth_state', None)
    if not expected_state or not secrets.compare_digest(request.args.get('state') or '', expected_state):
        return redirect('/login?error=state_invalid')
    if request.args.get('error') or not request.args.get('code'):
        return redirect('/login?error=auth_denied')
    code = request.args.get('code')
    try:
        tokens = feishu_oauth.exchange_code(code, _get_redirect_uri())
        info = feishu_oauth.get_oauth_user_info(tokens['access_token'])
        open_id = info.get('open_id')
        if not open_id:
            return redirect('/login?error=user_info')
        if user_store.is_system_admin(open_id):
            logger.warning('OAuth 返回了保留的系统管理员 open_id，拒绝登录')
            return redirect('/login?error=user_info')
        profile = {
            'open_id': open_id,
            'union_id': info.get('union_id', ''),
            'user_id': info.get('user_id', ''),
            'name': info.get('name', ''),
            'avatar': info.get('avatar_url', ''),
            'access_token': tokens['access_token'],
            'refresh_token': tokens.get('refresh_token', ''),
            'token_expire_at': int(time.time()) + int(tokens.get('expires_in', 7200)),
            'refresh_token_expire_at': int(time.time()) + int(tokens.get('refresh_token_expires_in') or 30 * 24 * 3600),
            'scope': tokens.get('scope', ''),
        }
        # 仅对“已存在且被禁用”的记录拦截；is_disabled 对不存在用户返回 True，首次登录不能在此拦截
        existing = user_store.get_user(open_id)
        if existing and existing.get('disabled'):
            return redirect('/login?error=disabled')
        if not user_store.upsert_user(profile):
            logger.warning(f'用户信息落库失败 open_id={open_id}，不阻断本次登录')
        # OAuth 登录切换回飞书用户身份时，清除密码管理员标志。
        session.pop('is_admin', None)
        session['user'] = {'open_id': open_id, 'name': profile['name']}
        session.permanent = True
        return redirect('/')
    except Exception as e:
        logger.error(f'飞书 OAuth 登录失败: {e}', exc_info=True)
        return redirect('/login?error=oauth_failed')

admin_path = config.get('admin.path', '/admin')

@app.route('/favicon.ico', methods=['GET'])
def favicon():
    # 图标路径可通过 page.favicon 配置，默认使用内置图标
    rel = (config.get('page.favicon', 'src/static/favicon.ico') or '').strip()
    candidates = []
    if rel:
        candidates.append(os.path.join(base_dir, rel))
    candidates.append(os.path.join(CURRENT_DIR, 'static', 'favicon.ico'))
    for path in candidates:
        if os.path.isfile(path):
            mimetype = mimetypes.guess_type(path)[0] or 'image/x-icon'
            return send_file(path, mimetype=mimetype)
    return ('', 404)

@app.route(admin_path, methods=['GET'])
def admin_page():
    if _is_admin_session():
        with open(os.path.join(HTML_DIR, 'dashboard.html'), 'r', encoding='utf-8') as f:
            html = f.read()
        copyright_text = escape(config.get('copyright.text', 'Hawkon 2025 -2026'))
        html = html.replace('Hawkon 2025 -2026', copyright_text)
        html = html.replace('Hawkon 2025', copyright_text)
        html = html.replace('[/* page_title */]', escape(config.get('page.title', '飞书文档下载工具')))
        html = html.replace('[/* default_template */]', escape(config.get('template.default', 'template.docx'), quote=True))
        html = html.replace('[/* default_template_json */]', _script_json(config.get('template.default', 'template.docx')))
        html = html.replace('[/* usage_url */]', escape(_safe_http_url(config.get('usage.url', 'https://github.com/HawkonS/feishu-docget'))))
        html = html.replace('[/* image_max_width */]', str(config.get('image.max_width', '16')))
        html = html.replace('[/* image_max_height */]', str(config.get('image.max_height', '23')))
        html = html.replace('/* [style_css] */', TableStyleManager.get_frontend_css())
        html = _inject_csrf(html)
        return html
    with open(os.path.join(HTML_DIR, 'login.html'), 'r', encoding='utf-8') as f:
        html = f.read()
    copyright_text = escape(config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('Hawkon 2025 -2026', copyright_text)
    html = html.replace('Hawkon 2025', copyright_text)
    return _inject_csrf(html)

@app.route('/api/admin/login', methods=['POST'])
def api_admin_login():
    client_ip = request.remote_addr or 'unknown'
    locked, remaining = _admin_login_lock(client_ip)
    if locked:
        return jsonify({'status': 'error', 'message': f'账户已锁定，请 {remaining} 秒后重试'})

    data = request.get_json(silent=True) or {}
    password = (data.get('password') or '').strip()
    admin_password = str(config.get('admin.password') or '').strip()
    if admin_password and hmac.compare_digest(password, admin_password):
        _clear_admin_login_failures(client_ip)
        session['is_admin'] = True
        session['user'] = {'open_id': SYSTEM_ADMIN_OPEN_ID, 'name': SYSTEM_ADMIN_NAME}
        session.permanent = True
        return jsonify({'status': 'ok'})

    # 登录失败，增加计数
    _record_admin_login_failure(client_ip)
    return jsonify({'status': 'error', 'message': '密码错误'})

@app.route('/api/admin/logout', methods=['POST'])
def api_admin_logout():
    # 密码管理员与飞书管理员共用后台入口；飞书管理员退出时需清除用户会话。
    is_password_admin = bool(session.get('is_admin'))
    user = session.get('user') or {}
    user_is_admin = user_store.is_admin(user.get('open_id', ''))
    session.pop('is_admin', None)
    # 任务令牌属于登录会话，退出后台后不应继续用于查询/下载任务。
    session.pop('job_tokens', None)
    # 非密码管理员调用该入口时一并退出飞书会话，避免管理员权限刚被取消后按钮失效。
    if user_is_admin or not is_password_admin:
        session.pop('user', None)
    return jsonify({'status': 'ok'})

@app.route('/api/user/logout', methods=['POST'])
def api_user_logout():
    session.pop('user', None)
    # 退出后不再持有任务访问令牌；正在运行的任务由工作线程继续执行，不受影响
    session.pop('job_tokens', None)
    return jsonify({'status': 'ok'})

@app.route('/api/admin/users', methods=['GET'])
@admin_required
def api_admin_users():
    paged = user_store.list_users(
        page=request.args.get('page', 1),
        page_size=request.args.get('page_size', 20),
        query=request.args.get('q', ''),
    )
    return jsonify({'status': 'ok', **paged, 'users': paged['items']})

@app.route('/api/admin/users/toggle', methods=['POST'])
@admin_required
def api_admin_users_toggle():
    data = request.get_json(silent=True) or {}
    open_id = str(data.get('open_id') or '').strip()
    if not open_id:
        return jsonify({'status': 'error', 'message': '缺少 open_id'})
    if not user_store.get_user(open_id):
        return jsonify({'status': 'error', 'message': '用户不存在'})
    if user_store.is_system_admin(open_id):
        return jsonify({'status': 'error', 'message': '系统管理员不能被禁用'})
    if not user_store.set_disabled(open_id, bool(data.get('disabled'))):
        return jsonify({'status': 'error', 'message': '保存失败，请检查服务器磁盘与权限'})
    return jsonify({'status': 'ok'})


@app.route('/api/admin/users/set-admin', methods=['POST'])
@admin_required
def api_admin_users_set_admin():
    data = request.get_json(silent=True) or {}
    open_id = str(data.get('open_id') or '').strip()
    if not open_id:
        return jsonify({'status': 'error', 'message': '缺少 open_id'})
    if not user_store.get_user(open_id):
        return jsonify({'status': 'error', 'message': '用户不存在'})
    if user_store.is_system_admin(open_id) and not bool(data.get('is_admin')):
        return jsonify({'status': 'error', 'message': '系统管理员不能取消管理员权限'})
    if not user_store.set_admin(open_id, bool(data.get('is_admin'))):
        return jsonify({'status': 'error', 'message': '保存失败，请检查服务器磁盘与权限'})
    return jsonify({'status': 'ok'})

@app.route('/api/admin/projects', methods=['GET'])
@admin_required
def api_admin_projects():
    return jsonify(list_project_summaries(
        page=request.args.get('page', 1),
        page_size=request.args.get('page_size', 20),
        query=request.args.get('q', ''),
    ))


@app.route('/api/admin/project-files', methods=['GET'])
@admin_required
def api_admin_project_files():
    path = request.args.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'}), 400
    output_dir = os.path.realpath(os.path.join(base_dir, config['output.dir']))
    real_path = os.path.realpath(path)
    if not real_path.startswith(output_dir + os.sep) or not os.path.isdir(real_path):
        return jsonify({'status': 'error', 'message': '无效路径'}), 400
    return jsonify({'status': 'ok', 'items': list_project_files(real_path)})

@app.route('/api/admin/download_project', methods=['GET'])
@admin_required
def api_admin_download_project():
    path = request.args.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    real_output = os.path.realpath(output_dir)
    real_path = os.path.realpath(path)
    if not real_path.startswith(real_output + os.sep) and real_path != real_output:
        return jsonify({'status': 'error', 'message': '无效路径'})
    if not os.path.exists(real_path):
        return jsonify({'status': 'error', 'message': '无效路径'})
    try:
        fd, tmp_zip = tempfile.mkstemp(suffix='.zip')
        os.close(fd)
        with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(path):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, path)
                    zipf.write(abs_path, rel_path)
        _cleanup_temp(tmp_zip)
        return send_file(tmp_zip, as_attachment=True, download_name=f'{os.path.basename(path)}.zip')
    except Exception as e:
        logger.error(f'下载项目打包失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '打包下载失败，请稍后重试'})

@app.route('/api/admin/delete_project', methods=['POST'])
@admin_required
def api_admin_delete_project():
    data = request.get_json(silent=True) or {}
    path = data.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    real_output = os.path.realpath(output_dir)
    real_path = os.path.realpath(path)
    if not real_path.startswith(real_output + os.sep) and real_path != real_output:
        return jsonify({'status': 'error', 'message': '无效路径'})
    if not os.path.exists(real_path):
        return jsonify({'status': 'error', 'message': '无效路径'})
    try:
        shutil.rmtree(real_path)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'删除项目失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '删除项目失败，请稍后重试'})

@app.route('/api/admin/download_folder', methods=['GET'])
@admin_required
def api_admin_download_folder():
    path = request.args.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    real_output = os.path.realpath(output_dir)
    real_path = os.path.realpath(path)
    if not real_path.startswith(real_output + os.sep) and real_path != real_output:
        return jsonify({'status': 'error', 'message': '无效路径'})
    if not os.path.exists(real_path):
        return jsonify({'status': 'error', 'message': '无效路径'})
    try:
        fd, tmp_zip = tempfile.mkstemp(suffix='.zip')
        os.close(fd)
        with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(path):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, path)
                    zipf.write(abs_path, rel_path)
        _cleanup_temp(tmp_zip)
        return send_file(tmp_zip, as_attachment=True, download_name=f'{os.path.basename(path)}.zip')
    except Exception as e:
        logger.error(f'下载文件夹打包失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '打包下载失败，请稍后重试'})

@app.route('/api/admin/delete_file', methods=['POST'])
@admin_required
def api_admin_delete_file():
    data = request.get_json(silent=True) or {}
    path = data.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    real_output = os.path.realpath(output_dir)
    real_path = os.path.realpath(path)
    if not real_path.startswith(real_output + os.sep) or not os.path.isfile(real_path):
        return jsonify({'status': 'error', 'message': '无效文件'})
    try:
        os.remove(real_path)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'删除文件失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '删除文件失败，请稍后重试'})

@app.route('/api/upload_template', methods=['POST'])
@login_required
def api_upload_template():
    # 验证请求数据
    password = request.form.get('password')
    mode = request.form.get('mode')
    file = request.files.get('file')
    image_file = request.files.get('image')
    name = request.form.get('name')
    
    # 如果是管理员登录，跳过密码验证，强制为 long_term 模式
    if _is_admin_session():
        mode = 'long_term'
    else:
        # 非管理员的密码校验同样受速率限制保护，防止暴力破解
        client_ip = request.remote_addr or 'unknown'
        locked, remaining = _check_login_lock(_upload_attempts, client_ip)
        if locked:
            return jsonify({'status': 'error', 'message': f'尝试次数过多，请 {remaining} 秒后重试'})

        # 验证模式
        if mode not in ['long_term', 'one_time']:
             return jsonify({'status': 'error', 'message': '无效的存储模式'})
        
        # 验证密码
        if mode == 'long_term':
            correct_password = config.get('template.password.long_term')
        elif mode == 'one_time':
            correct_password = config.get('template.password.one_time')
        else:
            correct_password = None
        
        if not correct_password or not hmac.compare_digest(str(password or ''), str(correct_password)):
            _record_login_failure(_upload_attempts, client_ip)
            return jsonify({'status': 'error', 'message': '密码错误'})
        with _attempts_lock:
            _upload_attempts.pop(client_ip, None)

    # 验证文件和名称
    if not name:
         return jsonify({'status': 'error', 'message': '模板名称不能为空'})

    template_dir = os.path.join(base_dir, config['template.dir'])
    safe_name = name.strip()
    safe_name = os.path.basename(safe_name)
    
    # 移除 .docx 后缀（如果用户输入了），避免重复
    if safe_name.lower().endswith('.docx'):
        safe_name = safe_name[:-5]
    
    if not safe_name:
        safe_name = 'template'
    # 从落盘源头清洗控制字符等（扩展名 .docx 在 final_filename 中重新拼接）
    safe_name = sanitize_name(safe_name)
    
    # 长期存储模式下，文件名就是用户输入的名称
    if mode == 'long_term':
        final_filename = f'{safe_name}.docx'
    else:
        # 仅本次使用模式下，加上 temp_ 前缀和 uuid，避免冲突和方便清理
        final_filename = f'temp_{uuid.uuid4().hex[:8]}_{safe_name}.docx'

    path = os.path.join(template_dir, final_filename)
    
    # 检查是否是更新操作
    is_update = os.path.exists(path)
    
    # 如果是新文件，必须上传 docx
    if not is_update and (not file or not file.filename.endswith('.docx')):
        return jsonify({'status': 'error', 'message': '新模板必须上传 Word 文件'})
        
    # 如果提供了文件，则保存（覆盖）
    if file:
        if not file.filename.endswith('.docx'):
             return jsonify({'status': 'error', 'message': '无效的 Word 文件'})
        try:
            file.save(path)
        except Exception as e:
            logger.error(f'保存模板文件失败: {e}', exc_info=True)
            return jsonify({'status': 'error', 'message': '保存模板文件失败，请稍后重试'})
            
    # 处理预览图
    if image_file:
        try:
            # 图片文件名与模板同名，后缀改为 .png
            img_filename = os.path.splitext(final_filename)[0] + '.png'
            img_path = os.path.join(template_dir, img_filename)
            image_file.save(img_path)
        except Exception as e:
            logger.error(f'保存预览图失败: {e}', exc_info=True)
            return jsonify({'status': 'error', 'message': '保存预览图失败，请稍后重试'})
            
    return jsonify({'status': 'ok', 'filename': final_filename})

@app.route('/api/admin/rename_template', methods=['POST'])
@admin_required
def api_admin_rename_template():
    data = request.get_json(silent=True) or {}
    old_name = data.get('old_name')
    new_name = data.get('new_name')
    
    if not old_name or not new_name:
        return jsonify({'status': 'error', 'message': '参数不完整'})
        
    if old_name == new_name:
        return jsonify({'status': 'ok'})
        
    template_dir = os.path.realpath(os.path.join(base_dir, config['template.dir']))

    safe_old = os.path.basename(str(old_name).strip())
    if not safe_old.lower().endswith('.docx'):
        safe_old += '.docx'
    old_path = _resolve_template_path(safe_old)
    if not old_path:
        return jsonify({'status': 'error', 'message': '原模板不存在'})
        
    # 处理 new_name
    safe_new = os.path.basename(str(new_name).strip())
    if safe_new != str(new_name).strip() or safe_new in {'.', '..'} or not safe_new:
        return jsonify({'status': 'error', 'message': '无效文件名'})
    # 如果用户没输后缀，后端逻辑通常是加上，但这里 old_name 已经是带后缀的文件名吗？
    # 前端传过来的 name 通常是不带后缀的显示名，还是带后缀的？
    # list_templates 返回的 name 是带 .docx 的 (e.g. "template.docx")
    # 所以 old_name 是 "abc.docx", new_name 可能是 "xyz"
    
    if safe_new.lower().endswith('.docx'):
        safe_new_filename = safe_new
    else:
        safe_new_filename = safe_new + '.docx'
        
    new_path = os.path.join(template_dir, safe_new_filename)
    if os.path.realpath(new_path).startswith(template_dir + os.sep) is False:
        return jsonify({'status': 'error', 'message': '无效文件名'})
    
    if os.path.exists(new_path):
        return jsonify({'status': 'error', 'message': '新名称已存在'})
        
    try:
        # 重命名 docx
        os.rename(old_path, new_path)
        
        # 重命名 png (如果有)
        old_png = os.path.splitext(old_path)[0] + '.png'
        new_png = os.path.splitext(new_path)[0] + '.png'
        if os.path.exists(old_png):
            os.rename(old_png, new_png)
            
        # 如果是默认模板，更新配置
        default_template = config.get('template.default', 'template.docx')
        # default_template 是带后缀的
        if safe_old == default_template:
            ConfigLoader.save_config_from_admin({'template.default': safe_new_filename})
            
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'重命名模板失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '重命名模板失败，请稍后重试'})

@app.route('/api/admin/delete_template', methods=['POST'])
@admin_required
def api_admin_delete_template():
    data = request.get_json(silent=True) or {}
    name = data.get('name')
    if not name:
        return jsonify({'status': 'error', 'message': '模板名称不能为空'})
        
    safe_name = os.path.basename(name)
    if safe_name != str(name).strip() or not safe_name.lower().endswith('.docx'):
        return jsonify({'status': 'error', 'message': '无效文件名'})
    
    # 禁止删除默认模板
    default_template = config.get('template.default', 'template.docx')
    if safe_name == default_template:
        return jsonify({'status': 'error', 'message': '默认模板不能删除'})
        
    path = _resolve_template_path(safe_name)
    if not path:
        return jsonify({'status': 'error', 'message': '模板不存在'})
        
    try:
        os.remove(path)
        # 尝试删除对应的图片
        png_path = os.path.splitext(path)[0] + '.png'
        if os.path.exists(png_path):
            os.remove(png_path)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'删除模板失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '删除模板失败，请稍后重试'})

@app.route('/api/admin/set_default_template', methods=['POST'])
@admin_required
def api_admin_set_default_template():
    data = request.get_json(silent=True) or {}
    name = data.get('name')
    if not name:
        return jsonify({'status': 'error', 'message': '模板名称不能为空'})
        
    safe_name = os.path.basename(name)
    if '..' in name or safe_name != name:
        return jsonify({'status': 'error', 'message': '无效文件名'})

    if not _resolve_template_path(safe_name):
            return jsonify({'status': 'error', 'message': '模板文件不存在'})

    try:
        ConfigLoader.save_config_from_admin({'template.default': safe_name})
        # 更新内存中的 config 对象，确保立即生效。
        # 但为了保险，我们可以不操作，直接依赖 ConfigLoader 的单例特性。
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'设置默认模板失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '设置默认模板失败，请稍后重试'})

@app.route('/api/start', methods=['POST'])
@login_required
def api_start():
    client_ip = request.remote_addr or 'unknown'
    user = session.get('user') or {}
    is_system_admin = _is_system_admin_session()
    # 系统账号没有飞书 access_token，任务应继续使用系统机器人，但统计显示固定名称。
    user_open_id = '' if is_system_admin else user.get('open_id', '')
    user_name = SYSTEM_ADMIN_NAME if is_system_admin else user.get('name', '')  # 姓名快照，随任务链路传递
    if not _check_start_rate_limit(client_ip):
        return jsonify({'status': 'error', 'message': '提交过于频繁，请稍后再试'})
    data = request.get_json(silent=True) or {}
    doc_url = str(data.get('url') or '').strip()
    template = str(data.get('template') or '').strip()
    table_style = str(data.get('tableStyle') or '').strip()
    add_cover = bool(data.get('addCover'))
    ignore_mention = bool(data.get('ignoreMention'))
    ignore_template_heading_num = bool(data.get('ignoreTemplateHeadingNum'))
    unordered_list_style = str(data.get('unorderedListStyle') or 'default').strip()
    body_style = data.get('bodyStyle') # dict or None
    image_style = data.get('imageStyle') # dict or None
    table_config = data.get('tableConfig') # dict or None
    margin_config = data.get('marginConfig') # dict or None
    code_block_config = data.get('codeBlockConfig') # dict or None
    document_info = data.get('documentInfo') # dict or None
    add_title = bool(data.get('addTitle'))
    if not doc_url:
        return jsonify({'status': 'error', 'message': '缺少文档链接'})
    template_path = _resolve_template_path(template, allow_empty=True)
    if template_path is None:
        return jsonify({'status': 'error', 'message': '模板文件不存在或路径无效'})
    # 队列中仅保留经过校验的文件名；实际执行时再次 realpath 校验，防止竞态替换。
    template = os.path.basename(template_path) if template_path else ''
    try:
        bot_config = normalize_bot_config(data.get('botConfig'))
    except ValueError as e:
        return jsonify({'status': 'error', 'message': str(e)})
    document_info_error = _validate_document_info(document_info)
    if document_info_error:
        return jsonify({'status': 'error', 'message': document_info_error})
    job_id = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:8]
    # 任务访问令牌：仅创建者会话与管理员可查询/下载/停止该任务，防止他人枚举 job_id
    job_token = secrets.token_urlsafe(16)
    session_tokens = session.setdefault('job_tokens', {})
    session_tokens[job_id] = job_token
    session.modified = True
    is_temp_template = template.startswith('temp_')
    with jobs_lock:
        jobs[job_id] = {'status': 'pending', 'progress': 0, 'message': '等待中', 'job_id': job_id, 'created_at': datetime.now().isoformat(timespec='seconds'), 'doc_url': doc_url, 'template': template, 'table_style': table_style, 'unordered_list_style': unordered_list_style, 'body_style': body_style, 'image_style': image_style, 'table_config': table_config, 'margin_config': margin_config, 'code_block_config': code_block_config, 'document_info': document_info, 'custom_bot_enabled': bool(bot_config), 'client_ip': client_ip, 'user_open_id': user_open_id, 'user_name': user_name, 'logs': [{'ts': datetime.now().isoformat(timespec='seconds'), 'message': '任务已创建'}]}

    def check_stop():
        with jobs_lock:
            job = jobs.get(job_id)
            if job and job.get('status') == 'stopped':
                return True
        return False
    scope_type, scope_id, current_active, current_pending = _scope_job_counts(user_open_id)
    # 个人登录任务按 open_id 各自限流；机器人/未登录任务仍共用一个全局槽位。
    with active_downloads_lock:
        active_scope_count = _scope_active_count_locked(scope_type, scope_id)
        is_queued = active_scope_count >= _scope_limit(scope_type)
    if is_queued:
        # 调度线程占槽位早于 run_job 更新状态，取两者较大值避免显示“还需等待 0 份”。
        wait_count = max(current_active, active_scope_count) + max(0, current_pending - 1)
        scope_label = '账号' if scope_type == 'account' else '机器人'
        msg = f'因{scope_label}并发限制，创建下载任务排队中，您还需等待 {wait_count} 份文档下载完成'
        with jobs_lock:
            jobs[job_id]['message'] = msg
            jobs[job_id]['logs'].append({'ts': datetime.now().isoformat(timespec='seconds'), 'message': msg})
        update_download_stat(base_dir, config, job_id, '排队中', doc_url=doc_url, ip_address=client_ip, user_name=user_name)
    else:
        pass
    download_queue.put((job_id, doc_url, template, table_style, is_temp_template, add_cover, client_ip, check_stop, unordered_list_style, body_style, is_queued, image_style, ignore_mention, ignore_template_heading_num, table_config, margin_config, code_block_config, document_info, add_title, bot_config, user_open_id, user_name))
    return jsonify({'status': 'ok', 'job_id': job_id})

@app.route('/api/status/<job_id>', methods=['GET'])
@login_required
def api_status(job_id):
    if not _verify_job_access(job_id):
        return jsonify({'status': 'error', 'message': '任务未找到'})
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({'status': 'error', 'message': '任务未找到'})
        return jsonify(job)

@app.route('/api/download/<job_id>', methods=['GET'])
@login_required
def api_download(job_id):
    if not _verify_job_access(job_id):
        return jsonify({'status': 'error', 'message': '任务未找到'})
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.get('status') != 'done':
            return jsonify({'status': 'error', 'message': '任务未完成'})
        docx_path = job.get('docx_path')
        if not docx_path or not os.path.isfile(docx_path):
            return jsonify({'status': 'error', 'message': '文件未找到'})
    # 下载名与磁盘文件名解耦：清洗控制字符等，避免构造响应头时抛 ValueError；磁盘路径保持原样读取，保留扩展名避免被截断丢失
    base, ext = os.path.splitext(os.path.basename(docx_path))
    download_name = sanitize_name(base, ext)
    try:
        return send_file(docx_path, as_attachment=True, download_name=download_name)
    except Exception as e:
        logger.error(f'文件下载失败: {docx_path}, {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '文件下载失败'}), 500

@app.route('/api/stop/<job_id>', methods=['POST'])
@login_required
def api_stop(job_id):
    if not _verify_job_access(job_id):
        return jsonify({'status': 'error', 'message': '任务未找到'})
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({'status': 'error', 'message': '任务未找到'})
        if job.get('status') in ['done', 'error', 'stopped']:
            return jsonify({'status': 'error', 'message': '任务已结束，无法停止'})
        job['status'] = 'stopped'
        logs = job.get('logs') or []
        logs.append({'ts': datetime.now().isoformat(timespec='seconds'), 'message': '用户手动停止了任务', 'type': 'error'})
        job['logs'] = logs[-200:]
        folder = job.get('folder')
        if folder and os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                logs.append({'ts': datetime.now().isoformat(timespec='seconds'), 'message': '已清理未完成的任务文件', 'type': 'info'})
            except Exception as e:
                logger.error(f'停止任务时清理文件失败: {e}')
    return jsonify({'status': 'ok'})

@app.route('/api/jobs', methods=['GET'])
@admin_required
def api_jobs():
    with jobs_lock:
        items = list(jobs.values())
    items.sort(key=lambda x: x.get('created_at', ''), reverse=True)
    return jsonify({'items': items[:20]})

@app.route('/api/template/<name>', methods=['GET'])
def api_template(name):
    path = _resolve_template_path(name)
    if not path:
        return jsonify({'status': 'error', 'message': '模板未找到'})
    inline = request.args.get('inline', 'false').lower() == 'true'
    base, ext = os.path.splitext(os.path.basename(path))
    download_name = sanitize_name(base, ext)
    try:
        return send_file(path, as_attachment=not inline, download_name=download_name)
    except Exception as e:
        logger.error(f'模板下载失败: {path}, {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '模板下载失败'}), 500

@app.route('/api/template_preview/<name>', methods=['GET'])
def api_template_preview(name):
    safe_name = os.path.basename(str(name))
    if safe_name != str(name) or not safe_name.lower().endswith('.docx'):
        return (jsonify({'status': 'error', 'message': '预览图未找到'}), 404)
    png_name = os.path.splitext(safe_name)[0] + '.png'
    template_dir = os.path.realpath(os.path.join(base_dir, config['template.dir']))
    path = os.path.realpath(os.path.join(template_dir, png_name))
    if not path.startswith(template_dir + os.sep) or not os.path.isfile(path):
        return (jsonify({'status': 'error', 'message': '预览图未找到'}), 404)
    return send_file(path, mimetype='image/png')

@app.route('/api/admin/info', methods=['GET'])
@admin_required
def api_admin_info():
    return jsonify({'templates': list_templates(), 'table_styles': TableStyleManager.list_styles()})


@app.route('/api/admin/templates', methods=['GET'])
@admin_required
def api_admin_templates():
    templates = list_templates()
    query = str(request.args.get('q') or '').strip().lower()
    if query:
        templates = [t for t in templates if query in str(t.get('name') or '').lower()]
    paged = paginate_items(templates, request.args.get('page', 1), request.args.get('page_size', 12))
    return jsonify({'status': 'ok', **paged, 'templates': paged['items']})


@app.route('/api/admin/table-styles', methods=['GET'])
@admin_required
def api_admin_table_styles():
    return jsonify({'status': 'ok', 'items': TableStyleManager.list_styles()})

@app.route('/api/admin/download_file', methods=['GET'])
@admin_required
def api_admin_download_file():
    path = request.args.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效文件'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    real_output = os.path.realpath(output_dir)
    real_path = os.path.realpath(path)
    if not real_path.startswith(real_output + os.sep) and real_path != real_output:
        return jsonify({'status': 'error', 'message': '无效文件'})
    if not os.path.exists(real_path) or not os.path.isfile(real_path):
        return jsonify({'status': 'error', 'message': '无效文件'})
    # 清洗下载名中的控制字符等，保留扩展名避免被截断丢失
    base, ext = os.path.splitext(os.path.basename(real_path))
    download_name = sanitize_name(base, ext)
    try:
        return send_file(real_path, as_attachment=True, download_name=download_name)
    except Exception as e:
        logger.error(f'文件下载失败: {real_path}, {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '文件下载失败'}), 500

@app.route('/api/config', methods=['GET'])
@admin_required
def get_config_api():
    items = ConfigLoader.get_all_config_items()
    # 回调地址留空时下发按当前访问域名推导的预览值，供前端展示
    for item in items:
        if item['key'] == 'login.oauth.redirect_uri' and str(item.get('value') or '') == '':
            item['preview'] = _get_redirect_uri()
            break
    return jsonify(items)

@app.route('/api/config', methods=['POST'])
@admin_required
def save_config_api():
    data = request.get_json(silent=True) or {}
    new_config = {}
    for item in data:
        key = item.get('key', '')
        value = item.get('value', '')
        # 仅对敏感字段跳过脱敏值与空值，避免将 ****** 写入配置或清空密钥/密码
        if key in SENSITIVE_KEYS and (value == '******' or str(value).strip() == ''):
            continue
        # 未知键静默丢弃，对齐旧的白名单过滤行为
        if key not in ALLOWED_CONFIG_KEYS:
            continue
        new_config[key] = value
    # 保存前按 Schema 校验本次变更
    errors = validate_config(new_config)
    if errors:
        return jsonify({'status': 'error', 'message': '配置校验失败', 'errors': errors}), 422
    meta_map = ConfigLoader.get_meta_map()
    restart_required = [k for k in new_config if meta_map.get(k, {}).get('restart_required')]
    try:
        if not ConfigLoader.save_config_from_admin(new_config):
            return jsonify({'status': 'error', 'message': '写入配置文件失败，请检查文件权限或磁盘空间'})
        return jsonify({'status': 'ok', 'restart_required': restart_required})
    except Exception as e:
        logger.error(f'保存配置失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '保存配置失败，请稍后重试'})

@app.route('/api/stats', methods=['GET'])
@admin_required
def get_stats_api():
    stats = get_download_stats(
        base_dir, config,
        page=request.args.get('page', 1),
        page_size=request.args.get('page_size', 10),
        title=request.args.get('title', ''),
        url=request.args.get('url', ''),
        ip=request.args.get('ip', ''),
        date=request.args.get('date', ''),
    )
    return jsonify(stats)

@app.route('/api/admin/stats/delete', methods=['POST'])
@admin_required
def api_admin_stats_delete():
    data = request.get_json(silent=True) or {}
    ts_list = data.get('ts_list') or []
    id_list = data.get('id_list') or []
    if not ts_list and (not id_list):
        return jsonify({'status': 'error', 'message': '未选择记录'})
    try:
        deleted = delete_download_stats(base_dir, config, ts_list=ts_list, id_list=id_list)
        return jsonify({'status': 'ok', 'deleted': deleted})
    except Exception as e:
        logger.error(f'删除统计记录失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '删除统计记录失败，请稍后重试'})

@app.route('/api/download_all', methods=['GET'])
@admin_required
def download_all_api():
    output_dir = os.path.join(base_dir, config['output.dir'])
    zip_path = os.path.join(base_dir, 'all_downloads.zip')
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, output_dir)
                zipf.write(abs_path, rel_path)
    return send_file(zip_path, as_attachment=True, download_name='all_downloads.zip')

@app.route('/api/admin/logs', methods=['GET'])
@admin_required
def api_admin_logs():
    log_dir = os.path.join(base_dir, config.get('log.dir', 'logs'))
    if not os.path.exists(log_dir):
        return jsonify([])
    files = []
    for f in os.listdir(log_dir):
        if f in {'download_stats.jsonl', DATABASE_FILENAME,
                 DATABASE_FILENAME + '-wal', DATABASE_FILENAME + '-shm'}:
            continue
        path = os.path.join(log_dir, f)
        if os.path.isfile(path):
            files.append({'name': f, 'size': os.path.getsize(path), 'mtime': os.path.getmtime(path)})
    files.sort(key=lambda x: x['mtime'], reverse=True)
    return jsonify(files)

@app.route('/api/admin/logs/<filename>', methods=['GET'])
@admin_required
def api_admin_get_log(filename):
    if filename in {'download_stats.jsonl', DATABASE_FILENAME,
                    DATABASE_FILENAME + '-wal', DATABASE_FILENAME + '-shm'}:
         return jsonify({'status': 'error', 'message': 'Cannot read stats file'})
    safe_name = os.path.basename(filename)
    if safe_name != filename or '..' in filename:
        return jsonify({'status': 'error', 'message': '无效文件名'})
    log_dir = os.path.join(base_dir, config.get('log.dir', 'logs'))
    log_path = os.path.join(log_dir, safe_name)
    real_path = os.path.realpath(log_path)
    if not real_path.startswith(os.path.realpath(log_dir) + os.sep):
        return jsonify({'status': 'error', 'message': '无效文件'})
    if not os.path.exists(real_path) or not os.path.isfile(real_path):
        return jsonify({'status': 'error', 'message': 'File not found'})
    path = real_path
    try:
        # 分页参数: lines=每次返回行数(默认2000), skip_lines=从末尾跳过行数(默认0)
        max_lines = min(int(request.args.get('lines', 2000)), 10000)
        skip_lines = max(int(request.args.get('skip_lines', 0)), 0)
        
        file_size = os.path.getsize(path)
        
        if file_size == 0:
            return jsonify({'status': 'ok', 'content': '', 'total_lines': 0, 'returned_lines': 0, 'has_more': False, 'file_size': 0})
        
        # 高效读取文件末尾N行：从文件尾部反向读取块，拼接后分割
        with open(path, 'rb') as f:
            # 第一遍：从尾部反向读取块，收集原始字节直到获得足够行
            raw_tail = b''
            offset = file_size
            found_enough = False
            needed_lines = skip_lines + max_lines + 1  # 多读一行用于准确判断 has_more
            
            while offset > 0:
                chunk_size = min(65536, offset)
                offset -= chunk_size
                f.seek(offset)
                raw_tail = f.read(chunk_size) + raw_tail
                
                # 在安全边界（非文件起始时的完整行边界）统计行数
                check_data = raw_tail
                if offset > 0 and check_data and check_data[0:1] != b'\n':
                    # 首行可能不完整，跳过它再计数
                    first_nl = check_data.find(b'\n')
                    if first_nl >= 0:
                        check_data = check_data[first_nl + 1:]
                    else:
                        check_data = b''
                
                est_lines = check_data.count(b'\n')
                if est_lines >= needed_lines:
                    found_enough = True
                    break
            
            # 在行边界处分割
            raw_tail_began_incomplete = False
            if offset > 0 and raw_tail and raw_tail[0:1] != b'\n':
                first_nl = raw_tail.find(b'\n')
                if first_nl >= 0:
                    raw_tail = raw_tail[first_nl + 1:]
                    raw_tail_began_incomplete = True
                else:
                    raw_tail = b''
            
            # 分割得到所有行（尾部已经去掉了不完整的首行）
            tail_lines = raw_tail.split(b'\n')
            # 去除文件末尾换行产生的空元素
            if tail_lines and tail_lines[-1] == b'':
                tail_lines.pop()
            
            tail_count = len(tail_lines)
            
            # 计算总行数
            if offset > 0:
                # 文件还有未读取的前缀部分，统计其行数
                f.seek(0)
                prefix = f.read(offset)
                prefix_lines = prefix.count(b'\n')
                # 如果 raw_tail 的首行是不完整的（跨越了 chunk 边界），
                # 它会属于前缀的最后一行，需要额外 +1
                incomplete_first_line = (raw_tail_began_incomplete and tail_count > 0)
                total_lines = prefix_lines + tail_count + (1 if incomplete_first_line else 0)
            else:
                total_lines = tail_count
            
            # 应用分页：从 tail_lines 尾部跳过 skip_lines，取 max_lines
            if skip_lines >= tail_count:
                collected = []
            else:
                end_idx = tail_count - skip_lines
                start_idx = max(0, end_idx - max_lines)
                collected = tail_lines[start_idx:end_idx]
            
            has_more = total_lines > (skip_lines + len(collected))
        
        content = '\n'.join(line.decode('utf-8', errors='ignore') for line in collected)
        returned_lines = len(collected)
        
        return jsonify({
            'status': 'ok',
            'content': content,
            'total_lines': total_lines,
            'returned_lines': returned_lines,
            'skip_lines': skip_lines,
            'has_more': has_more,
            'file_size': file_size
        })
    except Exception as e:
        logger.error(f'读取日志文件失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '读取日志文件失败，请稍后重试'})

@app.route('/api/admin/logs/<filename>', methods=['DELETE'])
@admin_required
def api_admin_delete_log(filename):
    if filename in {'download_stats.jsonl', DATABASE_FILENAME,
                    DATABASE_FILENAME + '-wal', DATABASE_FILENAME + '-shm'}:
         return jsonify({'status': 'error', 'message': 'Cannot delete stats file'})
    safe_name = os.path.basename(filename)
    if safe_name != filename or '..' in filename:
        return jsonify({'status': 'error', 'message': '无效文件名'})
    log_dir = os.path.join(base_dir, config.get('log.dir', 'logs'))
    log_path = os.path.join(log_dir, safe_name)
    real_path = os.path.realpath(log_path)
    if not real_path.startswith(os.path.realpath(log_dir) + os.sep):
        return jsonify({'status': 'error', 'message': '无效文件'})
    if not os.path.exists(real_path):
        return jsonify({'status': 'error', 'message': 'File not found'})
    path = real_path
    try:
        os.remove(path)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'删除日志文件失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '删除日志文件失败，请稍后重试'})

@app.route('/api/admin/system', methods=['POST'])
@admin_required
def api_admin_system():
    data = request.get_json(silent=True) or {}
    action = data.get('action')

    if action == 'status':
        # 检查 systemctl 是否可用
        if shutil.which('systemctl') is None:
            return jsonify({'status': 'error', 'message': '未找到 systemctl，系统管理功能不可用'})
        try:
            # 使用 list-units 检查服务是否存在
            subprocess.check_call(['systemctl', 'status', 'feishu-docget'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # 获取状态详情
            output = subprocess.check_output(['systemctl', 'status', 'feishu-docget', '--no-pager'], stderr=subprocess.STDOUT)
            return jsonify({'status': 'ok', 'output': output.decode('utf-8')})
        except subprocess.CalledProcessError:
            return jsonify({'status': 'error', 'message': '服务未运行或不存在'})
        except Exception as e:
            logger.error(f'获取服务状态失败: {e}', exc_info=True)
            return jsonify({'status': 'error', 'message': '获取服务状态失败，请稍后重试'})

    elif action == 'update':
        script_path = os.path.join(base_dir, 'tools', 'update.sh')
        if not os.path.exists(script_path):
             return jsonify({'status': 'error', 'message': '更新脚本未找到'})
        
        def run_update_bg():
            try:
                # 使用 nohup 运行更新脚本，避免因服务重启导致脚本中断
                # 脚本内部会处理重启逻辑，输出追加到 logs/update.log
                log_dir = os.path.join(base_dir, config.get('log.dir', 'logs'))
                os.makedirs(log_dir, exist_ok=True)
                update_log_path = os.path.join(log_dir, 'update.log')
                with open(update_log_path, 'a') as update_log:
                    # 子进程会 dup 该文件描述符，父进程关闭句柄后后台脚本仍可继续写入
                    # stdin 设为 DEVNULL，切断脚本的交互式输入（如 sudo 密码询问），避免后台流程卡死
                    subprocess.Popen(['nohup', 'bash', script_path, '--yes'], cwd=base_dir, stdin=subprocess.DEVNULL, stdout=update_log, stderr=subprocess.STDOUT)
            except Exception as e:
                logger.error(f'启动更新脚本失败: {e}')
        
        threading.Thread(target=run_update_bg).start()
        return jsonify({'status': 'ok', 'message': '更新任务已在后台启动，服务稍后将自动重启'})

    elif action in ['restart', 'stop']:
        if shutil.which('systemctl') is None:
            return jsonify({'status': 'error', 'message': '未找到 systemctl，系统管理功能不可用'})
        cmd = ['sudo', 'systemctl', action, 'feishu-docget']
        try:
            sudo_pass = config.get('system.sudo_password')
            
            def run_cmd_bg():
                if sudo_pass:
                    proc = subprocess.Popen(
                        ['sudo', '-S'] + cmd[1:],
                        stdin=subprocess.PIPE,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    proc.communicate(input=(sudo_pass + '\n').encode())
                else:
                    subprocess.Popen(cmd)
            
            # 异步执行，防止阻塞 HTTP 响应
            threading.Thread(target=run_cmd_bg).start()
            
            msg = '正在重启服务...' if action == 'restart' else '正在停止服务...'
            return jsonify({'status': 'ok', 'message': msg})
        except Exception as e:
            logger.error(f'执行系统操作失败: {e}', exc_info=True)
            return jsonify({'status': 'error', 'message': '系统操作失败，请稍后重试'})
            
    else:
        return jsonify({'status': 'error', 'message': '无效的操作'})

@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    if request.path in {'/', '/login', admin_path} or request.path.startswith('/api/admin/'):
        response.headers['Cache-Control'] = 'no-store'
    # HSTS 仅在 HTTPS 模式下启用
    if app.config.get('SESSION_COOKIE_SECURE'):
        response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response

if __name__ == '__main__':
    port = ConfigLoader.get_int('server.port', 7800)
    if https_enabled:
        logger.info(f'服务启动于端口 {port} (HTTPS 代理模式)...')
    else:
        logger.info(f'服务启动于端口 {port}...')

    # 优先使用 waitress 生产级 WSGI 服务器，未安装时自动降级到 Flask 开发服务器
    try:
        from waitress import serve
        threads = ConfigLoader.get_int('server.threads', 8)
        logger.info(f'使用 waitress 生产服务器启动 (threads={threads})')
        serve(app, host='0.0.0.0', port=port, threads=threads)
    except ImportError:
        logger.warning('未安装 waitress，回退到 Flask 开发服务器（不建议用于生产环境，可执行 pip install waitress 升级）')
        app.run(host='0.0.0.0', port=port)
