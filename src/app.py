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
from functools import wraps
from datetime import datetime, timedelta
import subprocess
import secrets
import tempfile
import zipfile
from flask import Flask, jsonify, request, send_file, send_from_directory, session, redirect, url_for
from src.services.doc_service import process_document
from src.core.bot_store import normalize_bot_config
from src.core import user_store, feishu_oauth
from src.core.config_loader import config, ConfigLoader, parse_size, parse_bool, SENSITIVE_KEYS
from src.converters.docx.style_manager import TableStyleManager
from src.core.stats import update_download_stat, get_download_stats
base_dir = os.path.abspath(config.get('workspace.dir', '.'))
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_DIR = os.path.join(CURRENT_DIR, 'web', 'templates')
logger = ConfigLoader.get_logger('feishu_docget')
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
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
jobs = {}
jobs_lock = threading.Lock()
download_queue = queue.Queue()
active_downloads_lock = threading.Lock()
active_downloads = 0

_login_attempts = {}  # {ip: {'count': N, 'locked_until': timestamp, 'ban_level': N}}
_upload_attempts = {}  # {ip: {'count': N, 'locked_until': timestamp, 'ban_level': N}} 模板上传密码尝试记录
_start_attempts = {}  # {ip: [timestamps]} 任务提交接口的滑动窗口限流
START_RATE_LIMIT = 10  # 单 IP 每分钟最多提交任务数


def _check_login_lock(attempts, client_ip):
    """检查 IP 是否处于锁定状态。返回 (是否锁定, 剩余秒数)，锁定期过期时自动清理记录"""
    attempt = attempts.get(client_ip)
    if attempt and attempt.get('locked_until', 0) > 0:
        now = time.time()
        if now > attempt['locked_until']:
            attempts.pop(client_ip, None)
            return False, 0
        return True, int(attempt['locked_until'] - now)
    return False, 0


def _record_login_failure(attempts, client_ip):
    """记录一次密码校验失败，失败 5 次后按指数退避锁定"""
    now = time.time()
    attempt = attempts.get(client_ip)
    if not attempt:
        attempts[client_ip] = {'count': 1, 'locked_until': 0, 'ban_level': 0}
    else:
        attempt['count'] = attempt.get('count', 0) + 1
        if attempt['count'] >= 5:
            ban_level = attempt.get('ban_level', 0)
            lock_seconds = min(60 * (3 ** ban_level), 3600)
            attempt['locked_until'] = now + lock_seconds
            attempt['count'] = 0
            attempt['ban_level'] = ban_level + 1


def _check_start_rate_limit(client_ip):
    """任务提交接口的滑动窗口限流，超限返回 False"""
    now = time.time()
    window = _start_attempts.setdefault(client_ip, [])
    window[:] = [ts for ts in window if now - ts < 60]
    if len(window) >= START_RATE_LIMIT:
        return False
    window.append(now)
    return True


def _verify_job_access(job_id):
    """校验当前会话是否有权访问指定任务：管理员或任务创建者可访问"""
    if session.get('is_admin'):
        return True
    tokens = session.get('job_tokens') or {}
    return bool(tokens.get(job_id))


def _cleanup_temp(path):
    """延迟清理临时文件"""
    threading.Timer(60, lambda: os.path.exists(path) and os.remove(path)).start()


def _login_enabled():
    """飞书登录开关，由配置 login.enabled 控制"""
    return parse_bool(config.get('login.enabled', 'false'))


def _get_redirect_uri():
    """OAuth 回调地址：优先使用配置项，未配置时根据当前请求地址推导（HTTPS 模式下强制 https scheme）"""
    configured = str(config.get('login.oauth.redirect_uri', '') or '').strip()
    if configured:
        return configured
    base = request.url_root.rstrip('/')
    # nginx SSL 终结且未传 X-Forwarded-Proto 时 url_root 会是 http，需按配置纠正 scheme
    if parse_bool(config.get('server.https.enabled', 'false')) and base.startswith('http://'):
        base = 'https://' + base[len('http://'):]
    return base + '/auth/feishu/callback'


def login_required(f):
    # 未启用登录或已是管理员时直通；页面请求重定向 /login，API 请求返回 403 JSON
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not _login_enabled() or session.get('is_admin'):
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
        if not session.get('is_admin'):
            return (jsonify({'status': 'error', 'message': '未登录'}), 403)
        return f(*args, **kwargs)
    return decorated_function

def worker_thread():
    global active_downloads
    logger.info("工作线程已启动，等待任务...")
    while True:
        try:
            max_concurrent = int(config.get('max.concurrent.downloads', 1))
            with active_downloads_lock:
                if active_downloads >= max_concurrent:
                    time.sleep(1)
                    continue
            try:
                job_args = download_queue.get(timeout=1)
            except queue.Empty:
                continue
            
            logger.info(f"工作线程获取到任务，当前活动任务数: {active_downloads}")
            with active_downloads_lock:
                active_downloads += 1
            try:
                run_job(*job_args)
            except Exception as e:
                logger.error(f'工作线程任务错误: {e}')
            finally:
                with active_downloads_lock:
                    active_downloads -= 1
                download_queue.task_done()
                logger.info(f"任务完成，当前活动任务数: {active_downloads}")
        except Exception as e:
            logger.error(f'工作线程循环错误: {e}')
            time.sleep(1)
threading.Thread(target=worker_thread, daemon=True).start()

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
        limit = parse_size(config.get('output.max_size', '10G'))
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

def list_projects():
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
                files = []
                for root, _, filenames in os.walk(path):
                    for fname in filenames:
                        abs_path = os.path.join(root, fname)
                        rel_path = os.path.relpath(abs_path, path).replace('\\', '/')
                        try:
                            fctime = os.path.getctime(abs_path)
                        except Exception:
                            fctime = 0
                        files.append({'name': fname, 'rel_path': rel_path, 'path': abs_path, 'ctime': datetime.fromtimestamp(fctime).isoformat(timespec='minutes'), 'is_md': fname.endswith('.md')})
                files.sort(key=lambda x: x['ctime'], reverse=True)
                items.append({'name': name, 'path': path, 'ctime': datetime.fromtimestamp(ctime).isoformat(timespec='minutes'), 'ctime_ts': ctime, 'size': size, 'files': files})
    return sorted(items, key=lambda x: x.get('ctime_ts', 0), reverse=True)

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
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        if was_queued:
            update_job(job_id, message='已完成下载任务排队，成功创建下载任务', log_type='info')
        update_job(job_id, status='running', progress=5, message='正在准备任务...', log_type='dynamic')
        update_download_stat(base_dir, config, job_id, '下载中', doc_url=doc_url, ip_address=client_ip, user_name=user_name)
        template_path = ''
        if template_name:
            template_path = os.path.join(base_dir, config['template.dir'], template_name)
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
    template_json = json.dumps(templates, ensure_ascii=False)
    table_styles = TableStyleManager.list_styles()
    style_json = json.dumps(table_styles, ensure_ascii=False)
    style_css = TableStyleManager.get_frontend_css()
    with open(os.path.join(HTML_DIR, 'index.html'), 'r', encoding='utf-8') as f:
        html = f.read()
    html = html.replace('[/* template_json */]', template_json)
    html = html.replace('[/* style_json */]', style_json)
    html = html.replace('/* [style_css] */', style_css)
    html = html.replace('[/* usage_url */]', config.get('usage.url', 'https://github.com/HawkonS/feishu-docget'))
    html = html.replace('Hawkon 2025 -2026', config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('Hawkon 2025', config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('[/* page_title */]', config.get('page.title', '飞书文档下载工具'))
    html = html.replace('[/* page_description */]', config.get('page.description', '支持将飞书文档链接下载为指定模板的 Word 文件'))
    html = html.replace('[/* page_placeholder */]', config.get('page.placeholder', '输入飞书文档链接，如 https://hawkon.feishu.cn/wiki/...'))
    html = html.replace('[/* usage_link_text */]', config.get('page.usage_link_text', '使用说明'))
    html = html.replace('[/* usage_doc_url */]', config.get('url.usage_doc', 'https://github.com/HawkonS/feishu-docget'))
    html = html.replace('[/* default_template */]', config.get('template.default', 'template.docx'))
    html = html.replace('[/* image_max_width */]', str(config.get('image.max_width', '16')))
    html = html.replace('[/* image_max_height */]', str(config.get('image.max_height', '23')))
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
    html = html.replace('[/* page_title */]', config.get('page.title', '飞书文档下载工具'))
    html = html.replace('Hawkon 2025 -2026', config.get('copyright.text', 'Hawkon 2025 -2026'))
    html = html.replace('Hawkon 2025', config.get('copyright.text', 'Hawkon 2025 -2026'))
    return html

@app.route('/auth/feishu/authorize', methods=['GET'])
def auth_feishu_authorize():
    if not _login_enabled():
        return redirect('/')
    state = secrets.token_urlsafe(16)
    session['oauth_state'] = state
    return redirect(feishu_oauth.build_authorize_url(config.get('feishu.app_id'), _get_redirect_uri(), state))

@app.route('/auth/feishu/callback', methods=['GET'])
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
        department = feishu_oauth.resolve_department_names(info.get('department_ids') or [])
        profile = {
            'open_id': open_id,
            'union_id': info.get('union_id', ''),
            'user_id': info.get('user_id', ''),
            'name': info.get('name', ''),
            'department': department,
            'avatar': info.get('avatar_url', ''),
            'access_token': tokens['access_token'],
            'refresh_token': tokens.get('refresh_token', ''),
            'token_expire_at': int(time.time()) + int(tokens.get('expires_in', 7200)),
        }
        # 仅对“已存在且被禁用”的记录拦截；is_disabled 对不存在用户返回 True，首次登录不能在此拦截
        existing = user_store.get_user(open_id)
        if existing and existing.get('disabled'):
            return redirect('/login?error=disabled')
        if not user_store.upsert_user(profile):
            logger.warning(f'用户信息落库失败 open_id={open_id}，不阻断本次登录')
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
    if session.get('is_admin'):
        with open(os.path.join(HTML_DIR, 'dashboard.html'), 'r', encoding='utf-8') as f:
            html = f.read()
        html = html.replace('Hawkon 2025 -2026', config.get('copyright.text', 'Hawkon 2025 -2026'))
        html = html.replace('Hawkon 2025', config.get('copyright.text', 'Hawkon 2025 -2026'))
        html = html.replace('[/* page_title */]', config.get('page.title', '飞书文档下载工具'))
        html = html.replace('[/* default_template */]', config.get('template.default', 'template.docx'))
        html = html.replace('[/* image_max_width */]', str(config.get('image.max_width', '16')))
        html = html.replace('[/* image_max_height */]', str(config.get('image.max_height', '23')))
        html = html.replace('/* [style_css] */', TableStyleManager.get_frontend_css())
        return html
    return send_file(os.path.join(HTML_DIR, 'login.html'))

@app.route('/api/admin/login', methods=['POST'])
def api_admin_login():
    client_ip = request.remote_addr or 'unknown'
    locked, remaining = _check_login_lock(_login_attempts, client_ip)
    if locked:
        return jsonify({'status': 'error', 'message': f'账户已锁定，请 {remaining} 秒后重试'})

    data = request.get_json(silent=True) or {}
    password = (data.get('password') or '').strip()
    admin_password = str(config.get('admin.password') or '').strip()
    if password == admin_password:
        _login_attempts.pop(client_ip, None)
        session['is_admin'] = True
        session.permanent = True
        return jsonify({'status': 'ok'})

    # 登录失败，增加计数
    _record_login_failure(_login_attempts, client_ip)
    return jsonify({'status': 'error', 'message': '密码错误'})

@app.route('/api/admin/logout', methods=['POST', 'GET'])
def api_admin_logout():
    session.pop('is_admin', None)
    return jsonify({'status': 'ok'})

@app.route('/api/user/logout', methods=['POST'])
def api_user_logout():
    session.pop('user', None)
    return jsonify({'status': 'ok'})

@app.route('/api/admin/users', methods=['GET'])
@admin_required
def api_admin_users():
    return jsonify({'status': 'ok', 'users': user_store.list_users()})

@app.route('/api/admin/users/toggle', methods=['POST'])
@admin_required
def api_admin_users_toggle():
    data = request.get_json(silent=True) or {}
    open_id = str(data.get('open_id') or '').strip()
    if not open_id:
        return jsonify({'status': 'error', 'message': '缺少 open_id'})
    if not user_store.get_user(open_id):
        return jsonify({'status': 'error', 'message': '用户不存在'})
    if not user_store.set_disabled(open_id, bool(data.get('disabled'))):
        return jsonify({'status': 'error', 'message': '保存失败，请检查服务器磁盘与权限'})
    return jsonify({'status': 'ok'})

@app.route('/api/admin/projects', methods=['GET'])
@admin_required
def api_admin_projects():
    return jsonify(list_projects())

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
    if session.get('is_admin'):
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
        
        if not correct_password or password != correct_password:
            _record_login_failure(_upload_attempts, client_ip)
            return jsonify({'status': 'error', 'message': '密码错误'})
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
        
    template_dir = os.path.join(base_dir, config['template.dir'])
    
    # 处理 old_name
    safe_old = os.path.basename(old_name)
    if not safe_old.lower().endswith('.docx'):
        safe_old += '.docx'
    old_path = os.path.join(template_dir, safe_old)
    
    if not os.path.exists(old_path):
        return jsonify({'status': 'error', 'message': '原模板不存在'})
        
    # 处理 new_name
    safe_new = os.path.basename(new_name)
    # 如果用户没输后缀，后端逻辑通常是加上，但这里 old_name 已经是带后缀的文件名吗？
    # 前端传过来的 name 通常是不带后缀的显示名，还是带后缀的？
    # list_templates 返回的 name 是带 .docx 的 (e.g. "template.docx")
    # 所以 old_name 是 "abc.docx", new_name 可能是 "xyz"
    
    if safe_new.lower().endswith('.docx'):
        safe_new_filename = safe_new
    else:
        safe_new_filename = safe_new + '.docx'
        
    new_path = os.path.join(template_dir, safe_new_filename)
    
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
        
    template_dir = os.path.join(base_dir, config['template.dir'])
    safe_name = os.path.basename(name)
    
    # 禁止删除默认模板
    default_template = config.get('template.default', 'template.docx')
    if safe_name == default_template:
        return jsonify({'status': 'error', 'message': '默认模板不能删除'})
        
    path = os.path.join(template_dir, safe_name)
    
    if not os.path.exists(path):
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

    template_dir = os.path.join(base_dir, config['template.dir'])
    if not os.path.exists(os.path.join(template_dir, safe_name)):
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
    user_open_id = user.get('open_id', '')
    user_name = user.get('name', '')  # 姓名快照，随任务链路传递
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
        jobs[job_id] = {'status': 'pending', 'progress': 0, 'message': '等待中', 'job_id': job_id, 'created_at': datetime.now().isoformat(timespec='seconds'), 'doc_url': doc_url, 'template': template, 'table_style': table_style, 'unordered_list_style': unordered_list_style, 'body_style': body_style, 'image_style': image_style, 'table_config': table_config, 'margin_config': margin_config, 'code_block_config': code_block_config, 'document_info': document_info, 'custom_bot_enabled': bool(bot_config), 'client_ip': client_ip, 'user_name': user_name, 'logs': [{'ts': datetime.now().isoformat(timespec='seconds'), 'message': '任务已创建'}]}

    def check_stop():
        with jobs_lock:
            job = jobs.get(job_id)
            if job and job.get('status') == 'stopped':
                return True
        return False
    max_concurrent = int(config.get('max.concurrent.downloads', 1))
    current_active = 0
    current_pending = 0
    with jobs_lock:
        for j in jobs.values():
            if j['status'] == 'running':
                current_active += 1
            elif j['status'] == 'pending':
                current_pending += 1
    is_queued = False
    with active_downloads_lock:
        if active_downloads >= max_concurrent:
            is_queued = True
    if is_queued:
        wait_count = current_active + max(0, current_pending - 1)
        msg = f'因并发限制，创建下载任务排队中，您还需等待 {wait_count} 份文档下载完成'
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
        if not docx_path or not os.path.exists(docx_path):
            return jsonify({'status': 'error', 'message': '文件未找到'})
    return send_file(docx_path, as_attachment=True, download_name=os.path.basename(docx_path))

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
    template_dir = os.path.join(base_dir, config['template.dir'])
    safe_name = os.path.basename(name)
    path = os.path.join(template_dir, safe_name)
    if not os.path.exists(path):
        return jsonify({'status': 'error', 'message': '模板未找到'})
    inline = request.args.get('inline', 'false').lower() == 'true'
    return send_file(path, as_attachment=not inline, download_name=os.path.basename(path))

@app.route('/api/template_preview/<name>', methods=['GET'])
def api_template_preview(name):
    template_dir = os.path.join(base_dir, config['template.dir'])
    safe_name = os.path.basename(name)
    png_name = os.path.splitext(safe_name)[0] + '.png'
    path = os.path.join(template_dir, png_name)
    if not os.path.exists(path):
        return (jsonify({'status': 'error', 'message': '预览图未找到'}), 404)
    return send_file(path, mimetype='image/png')

@app.route('/api/admin/info', methods=['GET'])
@admin_required
def api_admin_info():
    return jsonify({'templates': list_templates(), 'table_styles': TableStyleManager.list_styles()})

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
    return send_file(real_path, as_attachment=True, download_name=os.path.basename(real_path))

@app.route('/api/config', methods=['GET'])
@admin_required
def get_config_api():
    return jsonify(ConfigLoader.get_all_config_items())

@app.route('/api/config', methods=['POST'])
@admin_required
def save_config_api():
    data = request.get_json(silent=True) or {}
    new_config = {}
    for item in data:
        key = item.get('key', '')
        value = item.get('value', '')
        # 仅对敏感字段跳过脱敏值，避免将 ****** 写入配置
        if key in SENSITIVE_KEYS and value == '******':
            continue
        new_config[key] = value
    try:
        ConfigLoader.save_config_from_admin(new_config)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f'保存配置失败: {e}', exc_info=True)
        return jsonify({'status': 'error', 'message': '保存配置失败，请稍后重试'})

@app.route('/api/stats', methods=['GET'])
@admin_required
def get_stats_api():
    stats = get_download_stats(base_dir, config)
    return jsonify(stats)

@app.route('/api/admin/stats/delete', methods=['POST'])
@admin_required
def api_admin_stats_delete():
    data = request.get_json(silent=True) or {}
    ts_list = data.get('ts_list') or []
    id_list = data.get('id_list') or []
    if not ts_list and (not id_list):
        return jsonify({'status': 'error', 'message': '未选择记录'})
    stats_file = os.path.join(base_dir, config['log.dir'], 'download_stats.jsonl')
    if not os.path.exists(stats_file):
        return jsonify({'status': 'ok'})
    try:
        lines = []
        with open(stats_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        new_lines = []
        ts_set = set((str(ts) for ts in ts_list))
        id_set = set((str(x) for x in id_list))
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
                if str(record.get('ts')) in ts_set:
                    continue
                if record.get('id') and str(record.get('id')) in id_set:
                    continue
                new_lines.append(line + '\n')
            except Exception:
                new_lines.append(line + '\n')
        with open(stats_file, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        return jsonify({'status': 'ok'})
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
        if f == 'download_stats.jsonl':
            continue
        path = os.path.join(log_dir, f)
        if os.path.isfile(path):
            files.append({'name': f, 'size': os.path.getsize(path), 'mtime': os.path.getmtime(path)})
    files.sort(key=lambda x: x['mtime'], reverse=True)
    return jsonify(files)

@app.route('/api/admin/logs/<filename>', methods=['GET'])
@admin_required
def api_admin_get_log(filename):
    if filename == 'download_stats.jsonl':
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
    if filename == 'download_stats.jsonl':
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
    # HSTS 仅在 HTTPS 模式下启用
    if app.config.get('SESSION_COOKIE_SECURE'):
        response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response

if __name__ == '__main__':
    port = int(config.get('server.port', '7800'))
    https_enabled = config.get('server.https.enabled', 'false').lower() == 'true'
    if https_enabled:
        from werkzeug.middleware.proxy_fix import ProxyFix
        app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)
        app.config['SESSION_COOKIE_SECURE'] = True
        app.config['PREFERRED_URL_SCHEME'] = 'https'
        logger.info(f'服务启动于端口 {port} (HTTPS 代理模式)...')
    else:
        logger.info(f'服务启动于端口 {port}...')

    # 优先使用 waitress 生产级 WSGI 服务器，未安装时自动降级到 Flask 开发服务器
    try:
        from waitress import serve
        threads = int(config.get('server.threads', '8'))
        logger.info(f'使用 waitress 生产服务器启动 (threads={threads})')
        serve(app, host='0.0.0.0', port=port, threads=threads)
    except ImportError:
        logger.warning('未安装 waitress，回退到 Flask 开发服务器（不建议用于生产环境，可执行 pip install waitress 升级）')
        app.run(host='0.0.0.0', port=port)
