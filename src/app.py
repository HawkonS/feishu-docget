import json
import os
import threading
import uuid
import shutil
import queue
import time
from functools import wraps
from datetime import datetime
from flask import Flask, jsonify, request, send_file, send_from_directory, session, redirect, url_for
from src.services.doc_service import process_document
from src.core.config_loader import config, ConfigLoader, parse_size
from src.converters.docx.style_manager import TableStyleManager
from src.core.stats import update_download_stat, get_download_stats
base_dir = os.path.abspath(config.get('workspace.dir', '.'))
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_DIR = os.path.join(CURRENT_DIR, 'web', 'templates')
logger = ConfigLoader.get_logger('feishu_docget')
app = Flask(__name__)
app.secret_key = config.get('server.secret_key', 'feishu_docget_secret_key_2025')
jobs = {}
jobs_lock = threading.Lock()
download_queue = queue.Queue()
active_downloads_lock = threading.Lock()
active_downloads = 0

def worker_thread():
    global active_downloads
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
    template_dir = os.path.join(base_dir, config['template.dir'])
    if not os.path.isdir(template_dir):
        return []
    items = []
    default_template = config.get('template.default', 'template.docx')
    for name in os.listdir(template_dir):
        if name.lower().endswith('.docx') and (not name.startswith('temp_')):
            path = os.path.join(template_dir, name)
            size = os.path.getsize(path) if os.path.exists(path) else 0
            png_name = os.path.splitext(name)[0] + '.png'
            png_path = os.path.join(template_dir, png_name)
            has_png = os.path.exists(png_path)
            pdf_name = os.path.splitext(name)[0] + '.pdf'
            pdf_path = os.path.join(template_dir, pdf_name)
            has_pdf = os.path.exists(pdf_path)
            is_default = name == default_template
            items.append({'name': name, 'size': size, 'has_png': has_png, 'png_name': png_name if has_png else None, 'has_pdf': has_pdf, 'pdf_name': pdf_name if has_pdf else None, 'is_default': is_default})
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

def run_job(job_id, doc_url, template_name, table_style, delete_template=False, add_cover=False, client_ip='', check_stop_func=None, unordered_list_style='default', was_queued=False):
    try:
        if check_stop_func and check_stop_func():
            raise InterruptedError('任务已停止')
        if was_queued:
            update_job(job_id, message='已完成下载任务排队，成功创建下载任务', log_type='info')
        update_job(job_id, status='running', progress=5, message='正在准备任务...', log_type='dynamic')
        update_download_stat(base_dir, config, job_id, '下载中', doc_url=doc_url, ip_address=client_ip)
        template_path = ''
        if template_name:
            template_path = os.path.join(base_dir, config['template.dir'], template_name)
        output_root = os.path.join(base_dir, config['output.dir'])
        result = process_document(doc_url=doc_url, template_path=template_path, table_style=table_style, base_dir=base_dir, output_root=output_root, progress_cb=lambda p, m, t='info': update_job(job_id, progress=p, message=m, log_type=t), add_cover=add_cover, check_stop_func=check_stop_func, unordered_list_style=unordered_list_style)
        if delete_template and template_path and os.path.exists(template_path):
            try:
                os.remove(template_path)
            except Exception:
                pass
        update_job(job_id, status='done', progress=100, message='已完成', docx_path=result['docx_path'], folder=result['folder'])
        update_download_stat(base_dir, config, job_id, '已完成', doc_url, result['docx_path'], title=result.get('title', os.path.basename(result['docx_path'])), ip_address=client_ip)
        threading.Thread(target=check_cleanup_output).start()
    except Exception as e:
        is_stopped = isinstance(e, InterruptedError) or (check_stop_func and check_stop_func())
        if is_stopped:
            logger.info(f'任务 {job_id} 已被用户停止')
            update_job(job_id, status='stopped', message='任务已停止', log_type='error')
            update_download_stat(base_dir, config, job_id, '已停止', doc_url, ip_address=client_ip)
        else:
            logger.error('任务失败: ' + str(e))
            update_job(job_id, status='error', message=str(e))
            update_download_stat(base_dir, config, job_id, '错误', doc_url, ip_address=client_ip)

@app.errorhandler(404)
def page_not_found(e):
    target = config.get('url.404', 'https://space.hawkon.tech/')
    if not target.startswith('http'):
        target = 'http://' + target
    return redirect(target)

@app.route('/', methods=['GET'])
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
    return html
admin_path = config.get('admin.path', '/admin')

@app.route(admin_path, methods=['GET'])
def admin_page():
    if session.get('is_admin'):
        with open(os.path.join(HTML_DIR, 'dashboard.html'), 'r', encoding='utf-8') as f:
            html = f.read()
        html = html.replace('Hawkon 2025 -2026', config.get('copyright.text', 'Hawkon 2025 -2026'))
        html = html.replace('Hawkon 2025', config.get('copyright.text', 'Hawkon 2025 -2026'))
        html = html.replace('[/* page_title */]', config.get('page.title', '飞书文档下载工具'))
        html = html.replace('[/* default_template */]', config.get('template.default', 'template.docx'))
        return html
    return send_file(os.path.join(HTML_DIR, 'login.html'))

@app.route('/api/admin/login', methods=['POST'])
def api_admin_login():
    data = request.get_json(silent=True) or {}
    password = (data.get('password') or '').strip()
    admin_password = str(config.get('admin.password') or '').strip()
    if password == admin_password:
        session['is_admin'] = True
        return jsonify({'status': 'ok'})
    return jsonify({'status': 'error', 'message': '密码错误'})

@app.route('/api/admin/logout', methods=['POST', 'GET'])
def api_admin_logout():
    session.pop('is_admin', None)
    return jsonify({'status': 'ok'})

def admin_required(f):

    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('is_admin'):
            return (jsonify({'status': 'error', 'message': '未登录'}), 403)
        return f(*args, **kwargs)
    return decorated_function

@app.route('/api/admin/projects', methods=['GET'])
@admin_required
def api_admin_projects():
    return jsonify(list_projects())

@app.route('/api/admin/download_project', methods=['GET'])
@admin_required
def api_admin_download_project():
    path = request.args.get('path')
    if not path or not os.path.exists(path) or (not path.startswith(os.path.join(base_dir, config['output.dir']))):
        return jsonify({'status': 'error', 'message': '无效路径'})
    import zipfile
    import tempfile
    try:
        tmp_zip = tempfile.mktemp(suffix='.zip')
        with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(path):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, path)
                    zipf.write(abs_path, rel_path)
        return send_file(tmp_zip, as_attachment=True, download_name=f'{os.path.basename(path)}.zip')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/admin/delete_project', methods=['POST'])
@admin_required
def api_admin_delete_project():
    data = request.get_json(silent=True) or {}
    path = data.get('path')
    if path and os.path.exists(path) and path.startswith(os.path.join(base_dir, config['output.dir'])):
        try:
            shutil.rmtree(path)
            return jsonify({'status': 'ok'})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)})
    return jsonify({'status': 'error', 'message': '无效路径'})

@app.route('/api/admin/download_folder', methods=['GET'])
@admin_required
def api_admin_download_folder():
    path = request.args.get('path')
    if not path or not os.path.exists(path) or (not path.startswith(os.path.join(base_dir, config['output.dir']))):
        return jsonify({'status': 'error', 'message': '无效路径'})
    import zipfile
    import tempfile
    try:
        tmp_zip = tempfile.mktemp(suffix='.zip')
        with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(path):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, path)
                    zipf.write(abs_path, rel_path)
        return send_file(tmp_zip, as_attachment=True, download_name=f'{os.path.basename(path)}.zip')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/admin/delete_file', methods=['POST'])
def api_admin_delete_file():
    data = request.get_json(silent=True) or {}
    path = data.get('path')
    if not path:
        return jsonify({'status': 'error', 'message': '无效路径'})
    output_dir = os.path.join(base_dir, config['output.dir'])
    abs_path = os.path.abspath(path)
    if not abs_path.startswith(output_dir) or not os.path.isfile(abs_path):
        return jsonify({'status': 'error', 'message': '无效文件'})
    try:
        os.remove(abs_path)
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/upload_template', methods=['POST'])
@admin_required
def api_upload_template():
    password = request.form.get('password')
    mode = request.form.get('mode')
    file = request.files.get('file')
    if password != config.get('template.upload_password'):
        return jsonify({'status': 'error', 'message': '密码错误'})
    if not file or not file.filename.endswith('.docx'):
        return jsonify({'status': 'error', 'message': '无效文件'})
    template_dir = os.path.join(base_dir, config['template.dir'])
    if mode == 'current':
        filename = f'temp_{uuid.uuid4().hex[:8]}_{file.filename}'
        path = os.path.join(template_dir, filename)
        file.save(path)
        return jsonify({'status': 'ok', 'temp_name': filename})
    else:
        filename = file.filename
        path = os.path.join(template_dir, filename)
        file.save(path)
        return jsonify({'status': 'ok'})

@app.route('/api/start', methods=['POST'])
def api_start():
    data = request.get_json(silent=True) or {}
    doc_url = str(data.get('url') or '').strip()
    template = str(data.get('template') or '').strip()
    table_style = str(data.get('tableStyle') or '').strip()
    add_cover = bool(data.get('addCover'))
    unordered_list_style = str(data.get('unorderedListStyle') or 'default').strip()
    if not doc_url:
        return jsonify({'status': 'error', 'message': '缺少文档链接'})
    job_id = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:8]
    client_ip = request.remote_addr
    is_temp_template = template.startswith('temp_')
    with jobs_lock:
        jobs[job_id] = {'status': 'pending', 'progress': 0, 'message': '等待中', 'job_id': job_id, 'created_at': datetime.now().isoformat(timespec='seconds'), 'doc_url': doc_url, 'template': template, 'table_style': table_style, 'unordered_list_style': unordered_list_style, 'client_ip': client_ip, 'logs': [{'ts': datetime.now().isoformat(timespec='seconds'), 'message': '任务已创建'}]}

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
        update_download_stat(base_dir, config, job_id, '排队中', doc_url=doc_url, ip_address=client_ip)
    else:
        pass
    download_queue.put((job_id, doc_url, template, table_style, is_temp_template, add_cover, client_ip, check_stop, unordered_list_style, is_queued))
    return jsonify({'status': 'ok', 'job_id': job_id})

@app.route('/api/status/<job_id>', methods=['GET'])
def api_status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({'status': 'error', 'message': '任务未找到'})
        return jsonify(job)

@app.route('/api/download/<job_id>', methods=['GET'])
def api_download(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.get('status') != 'done':
            return jsonify({'status': 'error', 'message': '任务未完成'})
        docx_path = job.get('docx_path')
        if not docx_path or not os.path.exists(docx_path):
            return jsonify({'status': 'error', 'message': '文件未找到'})
    return send_file(docx_path, as_attachment=True, download_name=os.path.basename(docx_path))

@app.route('/api/stop/<job_id>', methods=['POST'])
def api_stop(job_id):
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
    if not path or not os.path.exists(path) or (not path.startswith(os.path.join(base_dir, config['output.dir']))):
        return jsonify({'status': 'error', 'message': '无效文件'})
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

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
        new_config[item['key']] = item['value']
    try:
        ConfigLoader.save_config_from_admin(new_config)
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

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
        logger.error(f'删除统计记录失败: {e}')
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/download_all', methods=['GET'])
@admin_required
def download_all_api():
    output_dir = os.path.join(base_dir, config['output.dir'])
    zip_path = os.path.join(base_dir, 'all_downloads.zip')
    if os.path.exists(zip_path):
        os.remove(zip_path)
    import zipfile
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, output_dir)
                zipf.write(abs_path, rel_path)
    return send_file(zip_path, as_attachment=True, download_name='all_downloads.zip')
if __name__ == '__main__':
    port = int(config.get('server.port', '7800'))
    logger.info(f'服务启动于端口 {port}...')
    app.run(host='0.0.0.0', port=port)