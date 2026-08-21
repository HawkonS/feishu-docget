import os
import sys
import time
import gzip
import shutil
import logging
import secrets
from logging.handlers import RotatingFileHandler
CONFIG_META = [{'name': '飞书配置 (必填)', 'items': [{'key': 'feishu.app_id', 'default': '', 'desc': '# 飞书 App ID'}, {'key': 'feishu.app_secret', 'default': '', 'desc': '# 飞书 App Secret'}, {'key': 'login.oauth.redirect_uri', 'default': '', 'desc': '# 飞书 OAuth 回调地址（留空则按访问域名自动生成，需在飞书开放平台配置同值重定向 URL）'}]}, {'name': '服务器配置', 'items': [{'key': 'server.port', 'default': '7800', 'desc': '# 服务器端口'}, {'key': 'server.https.enabled', 'default': 'false', 'desc': '# 是否启用 HTTPS，适用于反向代理 SSL 终结场景 (true/false)'}, {'key': 'admin.path', 'default': '/admin', 'desc': '# 管理后台路径'}, {'key': 'admin.password', 'default': '', 'desc': '# 管理后台密码 (留空则首次启动时自动生成随机密码)'}, {'key': 'template.default', 'default': 'Hawkon.docx', 'desc': '# 默认 Word 模板名称'}, {'key': 'system.sudo_password', 'default': '', 'desc': '# 系统 sudo 密码 (用于自动更新脚本，选填)'}, {'key': 'server.secret_key', 'default': '', 'desc': '# 用于 Flask session 加密的密钥，留空则自动生成'}, {'key': 'login.enabled', 'default': 'false', 'desc': '# 是否启用登录机制（飞书 OAuth），开启后未登录不能访问主页 (true/false)'}]}, {'name': '页面显示配置', 'items': [{'key': 'page.title', 'default': '飞书文档下载工具', 'desc': '# 页面标题'}, {'key': 'page.favicon', 'default': 'src/static/favicon.ico', 'desc': '# 网站图标文件路径 (相对工作区目录，支持 ico/png/svg)'}, {'key': 'page.description', 'default': '支持将飞书文档链接下载为指定模板的 Word 文件', 'desc': '# 页面描述'}, {'key': 'page.placeholder', 'default': '输入飞书文档链接，如 https://hawkon.feishu.cn/wiki/...', 'desc': '# 输入框占位符'}, {'key': 'page.usage_link_text', 'default': '使用说明', 'desc': '# 使用说明链接文本'}, {'key': 'url.usage_doc', 'default': 'https://github.com/HawkonS/feishu-docget', 'desc': '# 使用文档 URL'}, {'key': 'usage.url', 'default': 'mailto:contact@hawkon.tech', 'desc': '# 使用文档 URL / 联系方式'}, {'key': 'copyright.text', 'default': 'Hawkon 2025 -2026', 'desc': '# 版权文本'}, {'key': 'url.404', 'default': 'https://space.hawkon.tech/', 'desc': '# 404 重定向 URL'}, {'key': 'contact.name', 'default': 'Hakwon', 'desc': '# 联系人名称'}, {'key': 'bot.name', 'default': 'Hawkon-Tool', 'desc': '# 机器人名称'}]}, {'name': '导出设置', 'items': [{'key': 'image.max_height', 'default': '23', 'desc': '# 图片最大高度 (cm)'}, {'key': 'image.max_width', 'default': '16', 'desc': '# 图片最大宽度 (cm)'}, {'key': 'download.threads', 'default': '4', 'desc': '# 图片下载并发线程数'}, {'key': 'max.concurrent.downloads', 'default': '1', 'desc': '# 最大并发下载数'}, {'key': 'download_images', 'default': True, 'desc': '# 是否下载图片'}]}, {'name': '路径与日志配置', 'items': [{'key': 'workspace.dir', 'default': '.', 'desc': '# 工作区目录'}, {'key': 'template.dir', 'default': 'template', 'desc': '# 模板目录'}, {'key': 'template.password.long_term', 'default': '', 'desc': '# 模板上传密码 - 长期存储模式'}, {'key': 'template.password.one_time', 'default': '', 'desc': '# 模板上传密码 - 仅本次使用模式'}, {'key': 'output.dir', 'default': 'output', 'desc': '# 输出资源目录'}, {'key': 'output.max_size', 'default': '10G', 'desc': '# 输出目录最大大小'}, {'key': 'log.dir', 'default': 'logs', 'desc': '# 日志目录'}, {'key': 'log.level', 'default': 'INFO', 'desc': '# 日志级别'}, {'key': 'log.max_size', 'default': '20M', 'desc': '# 最大日志大小'}, {'key': 'log.backup_count', 'default': '5', 'desc': '# 日志轮转保留的历史文件数量'}, {'key': 'log.compress', 'default': 'true', 'desc': '# 轮转出的历史日志是否压缩为 .gz (true/false)'}, {'key': 'log.archive.enabled', 'default': 'true', 'desc': '# 超出保留数量的历史日志是否归档保存，false 则直接删除 (true/false)'}, {'key': 'log.archive.dir', 'default': '', 'desc': '# 日志归档目录 (留空则使用 <日志目录>/archive)'}, {'key': 'log.archive.max_days', 'default': '30', 'desc': '# 归档日志保留天数 (0 表示永久保留)'}]}]
DEFAULT_CONFIG = {}
for group in CONFIG_META:
    for item in group['items']:
        DEFAULT_CONFIG[item['key']] = item['default']
# 配置键白名单：仅允许写入 CONFIG_META 中定义的键，防止配置注入
ALLOWED_CONFIG_KEYS = set(DEFAULT_CONFIG.keys())
CONFIG_FILE = 'feishu-docget.properties'

SENSITIVE_KEYS = {
    'feishu.app_secret',
    'admin.password',
    'system.sudo_password',
    'template.password.long_term',
    'template.password.one_time',
    'server.secret_key'
}

def parse_size(size_str):
    size_str = size_str.upper()
    if size_str.endswith('K'):
        return int(size_str[:-1]) * 1024
    elif size_str.endswith('M'):
        return int(size_str[:-1]) * 1024 * 1024
    elif size_str.endswith('G'):
        return int(size_str[:-1]) * 1024 * 1024 * 1024
    else:
        return int(size_str)

def parse_bool(value, default=False):
    if value is None or str(value).strip() == '':
        return default
    return str(value).strip().lower() in ('true', '1', 'yes', 'on')

class ArchivingRotatingFileHandler(RotatingFileHandler):
    """支持压缩与归档的轮转日志 Handler。

    - compress: 轮转出的历史日志压缩为 .gz
    - archive_dir: 超出 backupCount 的最旧历史日志移动到归档目录，而不是直接删除
    - archive_max_days: 归档文件按修改时间清理，超过天数的自动删除 (0 表示不清理)
    """

    def __init__(self, filename, maxBytes=0, backupCount=0, encoding=None,
                 compress=True, archive_dir=None, archive_max_days=0):
        self.compress = compress
        self.archive_dir = archive_dir
        self.archive_max_days = archive_max_days
        super().__init__(filename, maxBytes=maxBytes, backupCount=backupCount, encoding=encoding)

    def _log_stem(self):
        base = os.path.basename(self.baseFilename)
        return base[:-4] if base.endswith('.log') else base

    def rotation_filename(self, default_name):
        if not self.compress:
            return default_name
        return default_name + '.gz'

    def rotate(self, source, dest):
        if not self.compress:
            super().rotate(source, dest)
            return
        if os.path.exists(source):
            with open(source, 'rb') as sf, gzip.open(dest, 'wb') as df:
                shutil.copyfileobj(sf, df)
            os.remove(source)

    def doRollover(self):
        if self.archive_dir:
            oldest = self.rotation_filename("%s.%d" % (self.baseFilename, self.backupCount))
            if os.path.exists(oldest):
                self._archive_file(oldest)
        super().doRollover()
        if self.archive_dir:
            self._cleanup_old_archives()

    def _archive_file(self, path):
        try:
            os.makedirs(self.archive_dir, exist_ok=True)
            stem = self._log_stem()
            suffix = '.log.gz' if path.endswith('.gz') else '.log'
            dest = os.path.join(self.archive_dir, f'{stem}-{time.strftime("%Y%m%d-%H%M%S")}{suffix}')
            seq = 1
            while os.path.exists(dest):
                dest = os.path.join(self.archive_dir, f'{stem}-{time.strftime("%Y%m%d-%H%M%S")}-{seq}{suffix}')
                seq += 1
            shutil.move(path, dest)
        except Exception as e:
            self.handleError(logging.LogRecord('', 0, '', 0, f'日志归档失败 {path}: {e}', None, None))

    def _cleanup_old_archives(self):
        if not self.archive_max_days:
            return
        try:
            cutoff = time.time() - self.archive_max_days * 86400
            prefix = self._log_stem() + '-'
            for name in os.listdir(self.archive_dir):
                if not name.startswith(prefix):
                    continue
                fp = os.path.join(self.archive_dir, name)
                if os.path.isfile(fp) and os.path.getmtime(fp) < cutoff:
                    os.remove(fp)
        except Exception:
            pass

class ConfigLoader:
    _config = {}
    _initialized = False

    @classmethod
    def load_config(cls):
        if cls._initialized:
            return cls._config
        config_path = os.path.join(os.getcwd(), CONFIG_FILE)
        if os.path.exists(config_path):
            cls._config = cls._read_config(config_path)
        else:
            print(f'配置文件 {CONFIG_FILE} 未找到。正在创建默认配置...')
            cls._write_config(config_path, DEFAULT_CONFIG)
            cls._config = DEFAULT_CONFIG.copy()
        for k, v in DEFAULT_CONFIG.items():
            if k not in cls._config:
                cls._config[k] = v
        # 管理后台密码为空时自动生成高熵随机密码，避免使用可预测的固定密码
        if not str(cls._config.get('admin.password') or '').strip():
            generated_password = secrets.token_urlsafe(16)
            cls._config['admin.password'] = generated_password
            print(f'已自动生成管理后台密码并写入 {CONFIG_FILE}: {generated_password}')
            print('请及时登录管理后台修改为自定义密码，并妥善保管。')
        cls._write_config(config_path, cls._config)
        if not cls._config.get('feishu.app_id') or not cls._config.get('feishu.app_secret'):
            if sys.stdin.isatty():
                print('\n缺少飞书 App ID 或 App Secret。')
                if not cls._config.get('feishu.app_id'):
                    cls._config['feishu.app_id'] = input('请输入飞书 App ID: ').strip()
                if not cls._config.get('feishu.app_secret'):
                    cls._config['feishu.app_secret'] = input('请输入飞书 App Secret: ').strip()
                cls._write_config(config_path, cls._config)
            else:
                print('警告: 配置中缺少飞书 App ID 或 App Secret。')
        workspace = cls._config.get('workspace.dir', '.')
        log_dir = os.path.join(workspace, cls._config.get('log.dir', 'logs'))
        output_dir = os.path.join(workspace, cls._config.get('output.dir', 'output'))
        template_dir = os.path.join(workspace, cls._config.get('template.dir', 'template'))
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        os.makedirs(template_dir, exist_ok=True)
        cls._initialized = True
        return cls._config

    @classmethod
    def _read_config(cls, path):
        config = {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        k, v = line.split('=', 1)
                        config[k.strip()] = v.strip()
        except Exception as e:
            print(f'读取配置错误 {path}: {e}')
        return config

    @classmethod
    def get_comment_map(cls):
        comments = {}
        for group in CONFIG_META:
            for item in group['items']:
                comments[item['key']] = item['desc']
        return comments

    @classmethod
    def _write_config(cls, path, config):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                for group in CONFIG_META:
                    f.write('# ==========================================\n')
                    f.write(f"# {group['name']}\n")
                    f.write('# ==========================================\n')
                    for item in group['items']:
                        key = item['key']
                        desc = item['desc']
                        if desc:
                            f.write(f'{desc}\n')
                        val = config.get(key, item['default'])
                        if val is None: val = ''
                        f.write(f"{key}={val}\n")
                    f.write('\n')
        except PermissionError as e:
            print(f"警告: 无法写入配置文件 {path}: {e}")
            print("请检查文件是否被设置为只读、隐藏，或正在被其他程序使用。")
        except Exception as e:
            print(f"写入配置文件 {path} 失败: {e}")

    @classmethod
    def get_comment(cls, key):
        comments = cls.get_comment_map()
        return comments.get(key, key)

    @classmethod
    def get_all_config_items(cls):
        cls.load_config()
        items = []
        for group in CONFIG_META:
            for item in group['items']:
                items.append({'key': item['key'], 'value': cls._config.get(item['key'], ''), 'desc': item['desc'], 'group': group['name']})
        for item in items:
            if item.get('key') in SENSITIVE_KEYS and item.get('value'):
                item['value'] = '******'
        return items

    @classmethod
    def save_config_from_admin(cls, new_config):
        config_path = os.path.join(os.getcwd(), CONFIG_FILE)
        # 仅接受白名单内的配置键，拒绝任意键注入
        filtered = {k: v for k, v in new_config.items() if k in ALLOWED_CONFIG_KEYS}
        cls._config.update(filtered)
        cls._write_config(config_path, cls._config)
        return True

    @classmethod
    def get_logger(cls, name, filename=None):
        config = cls.load_config()
        workspace = config.get('workspace.dir', '.')
        log_dir = os.path.join(workspace, config.get('log.dir', 'logs'))
        os.makedirs(log_dir, exist_ok=True)
        if not filename:
            filename = f'{name}.log'
        log_file = os.path.join(log_dir, filename)
        max_bytes = parse_size(config.get('log.max_size', '20M'))
        try:
            backup_count = int(config.get('log.backup_count', '5'))
        except (TypeError, ValueError):
            backup_count = 5
        compress = parse_bool(config.get('log.compress', 'true'), True)
        archive_enabled = parse_bool(config.get('log.archive.enabled', 'true'), True)
        archive_dir = config.get('log.archive.dir', '').strip() or os.path.join(log_dir, 'archive')
        try:
            archive_max_days = int(config.get('log.archive.max_days', '30'))
        except (TypeError, ValueError):
            archive_max_days = 30
        level_str = config.get('log.level', 'INFO').upper()
        level = getattr(logging, level_str, logging.INFO)
        logger = logging.getLogger(name)
        logger.setLevel(level)
        if not logger.handlers:
            handler = ArchivingRotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8',
                compress=compress,
                archive_dir=archive_dir if archive_enabled else None,
                archive_max_days=archive_max_days)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            console = logging.StreamHandler()
            console.setFormatter(formatter)
            logger.addHandler(console)
        return logger
config = ConfigLoader.load_config()